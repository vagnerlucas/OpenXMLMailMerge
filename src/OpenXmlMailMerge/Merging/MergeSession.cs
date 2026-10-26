using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Xml;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Wordprocessing;
using OpenXmlMailMerge.Elements;
using OpenXmlMailMerge.Fields;

namespace OpenXmlMailMerge.Merging
{
    /// <summary>
    /// One merge: a private, in-memory copy of the template, the data merged into it, and what matched.
    /// It lives for a single <see cref="DocumentMerger"/> call and is disposed before the call returns.
    /// </summary>
    internal sealed class MergeSession : IDisposable
    {
        private readonly MemoryStream _buffer;
        private readonly WordprocessingDocument _document;
        private readonly MainDocumentPart _main;
        private readonly MergeOptions _options;
        private readonly HashSet<string> _usedBindingNames = new(StringComparer.OrdinalIgnoreCase);
        private List<(string Name, string PartUri)> _unboundFields = [];
        private MergeData _data;
        private IdAllocator _ids;
        private bool _completed;

        private MergeSession(MemoryStream buffer, MergeOptions options)
        {
            _buffer = buffer;
            _options = options;
            buffer.Position = 0;

            // AutoSave is off: changes reach the buffer only when the document is saved explicitly.
            // MaxCharactersInPart stops a part from inflating past the limit while it's read (0 means no limit).
            _document = WordprocessingDocument.Open(buffer, true, new OpenSettings
            {
                AutoSave = false,
                MaxCharactersInPart = options.MaxPartCharacters
            });

            try
            {
                // An empty writable stream opens as a new, empty package: a valid zip that isn't a Word document.
                _main = _document.MainDocumentPart;
                if (_main == null || LoadRoot(_main, () => _main.Document) == null)
                    throw new InvalidDataException("The package has no main document part.");

                // Parts are only parsed when first read. Reading the ones every merge works on now makes an oversized or
                // broken part fail here, before anything is merged.
                foreach (var header in _main.HeaderParts)
                    LoadRoot(header, () => header.Header);
                foreach (var footer in _main.FooterParts)
                    LoadRoot(footer, () => footer.Footer);
            }
            catch
            {
                _document.Dispose();
                throw;
            }
        }

        /// <summary>
        /// Opens a template held in memory, copying it once into the session's own buffer.
        /// </summary>
        /// <param name="template">The .docx template.</param>
        /// <param name="options">The merge options.</param>
        /// <returns>The session.</returns>
        internal static MergeSession Open(ReadOnlySpan<byte> template, MergeOptions options)
        {
            // Not new MemoryStream(byte[]): that stream is fixed-size and fails as soon as the package grows.
            var buffer = new MemoryStream(template.Length);
            buffer.Write(template);

            return OpenBuffer(buffer, options);
        }

        /// <summary>
        /// Opens a template read from a stream, to its end. The stream is not closed.
        /// </summary>
        /// <param name="template">The .docx template.</param>
        /// <param name="options">The merge options.</param>
        /// <param name="cancellationToken">Cancellation token.</param>
        /// <returns>The session.</returns>
        internal static async Task<MergeSession> OpenAsync(Stream template, MergeOptions options, CancellationToken cancellationToken)
        {
            var buffer = new MemoryStream();
            await template.CopyToAsync(buffer, cancellationToken).ConfigureAwait(false);

            return OpenBuffer(buffer, options);
        }

        private static MergeSession OpenBuffer(MemoryStream buffer, MergeOptions options)
        {
            try
            {
                return new MergeSession(buffer, options);
            }
            catch
            {
                buffer.Dispose();
                throw;
            }
        }

        /// <summary>
        /// Merges data into the document, then applies <see cref="MergeOptions.UnboundFields"/> to the fields that got no
        /// data. Without bindings, nothing is merged.
        /// </summary>
        /// <param name="data">The data to merge.</param>
        internal void Merge(MergeData data)
        {
            _data = data;
            if (data.Names.Count > 0)
            {
                foreach (var root in GetRoots())
                    ComplexFields.ConvertToSimpleFields(root, HasBinding);

                MergeText();
                MergeImages();
                MergeTables();
            }

            // Merged fields are gone, so every MERGEFIELD still here got no data. Recorded before Remove takes them out,
            // so the report lists them either way.
            _unboundFields = FindMergeFields();

            if (_options.UnboundFields == UnboundFieldBehavior.Remove)
                RemoveMergeFields();
        }

        /// <summary>
        /// Finishes the merge: saves the document, reports what matched, and closes the package so the buffer holds the
        /// final document. Every zip entry then gets the template's main-document timestamp: the package stamps the entries
        /// it adds with the current time, and without this, merging the same input twice would give different bytes.
        /// </summary>
        /// <returns>The report.</returns>
        /// <exception cref="MergeException">
        /// Fields got no data and <see cref="MergeOptions.UnboundFields"/> is Throw, or the result isn't valid and
        /// <see cref="MergeOptions.ValidateOutput"/> is set.
        /// </exception>
        internal MergeReport Complete()
        {
            if (_options.UnboundFields == UnboundFieldBehavior.Throw && _unboundFields.Count > 0)
            {
                var (name, partUri) = _unboundFields[0];
                var names = string.Join(", ", _unboundFields.Select(field => field.Name));
                throw new MergeException($"The template has fields without data: {names}.", name, partUri);
            }

            if (_options.ValidateOutput)
                Validate();

            _document.Save();
            var report = CreateReport();

            var mainEntryName = _main.Uri.OriginalString.TrimStart('/');
            _document.Dispose();
            _completed = true;

            using (var zip = new ZipArchive(_buffer, ZipArchiveMode.Update, leaveOpen: true))
            {
                var timestamp = zip.GetEntry(mainEntryName)?.LastWriteTime ?? new DateTimeOffset(1980, 1, 1, 0, 0, 0, TimeSpan.Zero);
                foreach (var entry in zip.Entries)
                    entry.LastWriteTime = timestamp;
            }

            return report;
        }

        /// <summary>
        /// The size of the merged document in bytes, after <see cref="Complete"/>.
        /// </summary>
        internal long Length
        {
            get { return _buffer.Length; }
        }

        /// <summary>
        /// Gets the merged document, after <see cref="Complete"/>.
        /// </summary>
        /// <returns>The .docx bytes.</returns>
        internal byte[] GetBytes()
        {
            return _buffer.ToArray();
        }

        /// <summary>
        /// Writes the merged document to a stream, from its current position, without closing it. Call after <see cref="Complete"/>.
        /// </summary>
        /// <param name="output">The destination.</param>
        /// <param name="cancellationToken">Cancellation token.</param>
        internal async Task WriteToAsync(Stream output, CancellationToken cancellationToken)
        {
            var content = _buffer.GetBuffer().AsMemory(0, (int)_buffer.Length);
            await output.WriteAsync(content, cancellationToken).ConfigureAwait(false);
        }

        /// <summary>
        /// Lists the template's merge fields without merging anything. Complex fields are converted on the session's
        /// private copy, exactly as a merge would, to tell which ones can be merged; the copy is never saved.
        /// </summary>
        /// <returns>The fields.</returns>
        internal TemplateInfo Inspect()
        {
            var converted = new HashSet<SimpleField>();
            foreach (var root in GetRoots())
                converted.UnionWith(ComplexFields.ConvertToSimpleFields(root, instruction => MergeFieldName.Get(instruction) != null));

            var fields = new List<TemplateField>();
            foreach (var (part, root) in GetParts())
            {
                foreach (var (instruction, field) in FieldInstructions.Find(root))
                {
                    var name = MergeFieldName.Get(instruction);
                    if (name == null)
                        continue;

                    // What's still a complex field after the conversion is one a merge can't handle.
                    fields.Add(new TemplateField(
                        FieldParameters.GetBindingName(name),
                        FieldParameters.Parse(name),
                        part.Uri.OriginalString,
                        isComplex: field == null || converted.Contains(field),
                        isMergeable: field != null));
                }
            }

            return new TemplateInfo(fields);
        }

        /// <summary>
        /// Validates the merged document against the Open XML schema, using the newest Office version the SDK knows, so
        /// templates saved by current Word aren't flagged for markup newer than Office 2007.
        /// </summary>
        /// <exception cref="MergeException">The document isn't valid; the first errors are in the message.</exception>
        private void Validate()
        {
            var errors = new OpenXmlValidator(FileFormatVersions.Microsoft365).Validate(_document).ToList();
            if (errors.Count == 0)
                return;

            var first = errors[0];
            var details = string.Join(" ", errors.Take(3).Select(error => $"{error.Part?.Uri} {error.Path?.XPath}: {error.Description}"));
            throw new MergeException($"The merged document isn't valid ({errors.Count} errors). {details}", null, first.Part?.Uri.OriginalString);
        }

        /// <summary>
        /// Reports the fields that got no data and the bindings no field used.
        /// </summary>
        private MergeReport CreateReport()
        {
            var unusedBindings = _data == null ? [] : _data.Names.Where(name => !_usedBindingNames.Contains(name)).ToList();

            return new MergeReport([.. _unboundFields.Select(field => field.Name)], unusedBindings);
        }

        /// <summary>
        /// Lists the MERGEFIELDs in the document by binding name, without duplicates (ignoring case), with the part each
        /// is first found in: headers, then footers, then the body, each in document order. Nested and multi-paragraph
        /// fields are included.
        /// </summary>
        private List<(string Name, string PartUri)> FindMergeFields()
        {
            var fields = new List<(string Name, string PartUri)>();
            var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (var (part, root) in GetParts())
            {
                foreach (var (instruction, _) in FieldInstructions.Find(root))
                {
                    var name = MergeFieldName.GetBindingName(instruction);
                    if (name != null && seen.Add(name))
                        fields.Add((name, part.Uri.OriginalString));
                }
            }

            return fields;
        }

        /// <summary>
        /// Removes every MERGEFIELD left in the document, with its display text. Complex fields are converted to simple
        /// ones first; those that can't be (nested in another field, or spread over paragraphs) stay.
        /// </summary>
        private void RemoveMergeFields()
        {
            foreach (var root in GetRoots())
                ComplexFields.ConvertToSimpleFields(root, instruction => MergeFieldName.Get(instruction) != null);

            foreach (var field in GetSimpleFields())
            {
                if (MergeFieldName.Get(field) != null)
                    field.Remove();
            }
        }

        /// <summary>
        /// Reads a part's XML. A part that isn't valid XML, or is larger than <see cref="MergeOptions.MaxPartCharacters"/>,
        /// is reported as a <see cref="MergeException"/> naming the part.
        /// </summary>
        private T LoadRoot<T>(OpenXmlPart part, Func<T> read) where T : OpenXmlPartRootElement
        {
            try
            {
                return read();
            }
            catch (XmlException exception)
            {
                var reason = _options.MaxPartCharacters > 0
                    ? string.Create(CultureInfo.InvariantCulture, $"isn't valid XML, or is larger than MaxPartCharacters ({_options.MaxPartCharacters} characters)")
                    : "isn't valid XML";
                throw new MergeException($"The template part {part.Uri} {reason}.", null, part.Uri.OriginalString, exception);
            }
        }

        /// <summary>
        /// The root of every part that can hold drawings, for the id scan: the body, headers, footers, footnotes,
        /// endnotes and comments.
        /// </summary>
        private IEnumerable<OpenXmlPartRootElement> GetDrawingRoots()
        {
            foreach (var root in GetRoots())
                yield return (OpenXmlPartRootElement)root;

            if (_main.FootnotesPart is { } footnotes)
                yield return LoadRoot(footnotes, () => footnotes.Footnotes);
            if (_main.EndnotesPart is { } endnotes)
                yield return LoadRoot(endnotes, () => endnotes.Endnotes);
            if (_main.WordprocessingCommentsPart is { } comments)
                yield return LoadRoot(comments, () => comments.Comments);
        }

        /// <summary>
        /// The culture for values merged into a field: <see cref="MergeOptions.Culture"/> when set, else the template's
        /// language.
        /// </summary>
        private CultureInfo GetCulture(SimpleField field)
        {
            return _options.Culture ?? TemplateFormatting.GetCulture(field, _main);
        }

        /// <summary>
        /// Disposes the package and its buffer.
        /// </summary>
        public void Dispose()
        {
            if (!_completed)
                _document.Dispose();

            _buffer.Dispose();
        }

        /// <summary>
        /// The parts fields are merged in, each with its root element: every header, every footer, then the body.
        /// An image becomes a part of the part that holds its field.
        /// </summary>
        private IEnumerable<(OpenXmlPart Part, OpenXmlElement Root)> GetParts()
        {
            foreach (var header in _main.HeaderParts)
                yield return (header, header.Header);

            foreach (var footer in _main.FooterParts)
                yield return (footer, footer.Footer);

            yield return (_main, _main.Document);
        }

        private IEnumerable<OpenXmlElement> GetRoots()
        {
            return GetParts().Select(part => part.Root);
        }

        /// <summary>
        /// Every simple field, snapshotted first: replacing a field changes the tree being searched.
        /// </summary>
        private List<SimpleField> GetSimpleFields()
        {
            return [.. GetRoots().SelectMany(root => root.Descendants<SimpleField>())];
        }

        /// <summary>
        /// Gets the value bound to a field, when a binding of the given kind matches the field name exactly (ignoring case,
        /// and ignoring the field's parameters): key "IMG" matches IMG and IMG@WIDTH=100, not IMG_LOGO.
        /// </summary>
        /// <param name="kind">The binding kind.</param>
        /// <param name="field">The field.</param>
        /// <param name="value">The bound value, which can itself be null.</param>
        /// <returns>True when a binding matches the field.</returns>
        private bool TryFindBinding(BindingKind kind, SimpleField field, out object value)
        {
            value = null;

            var name = MergeFieldName.GetBindingName(field.Instruction?.Value);
            if (name == null || !_data.TryGetBinding(name, out var bindingKind, out var bindingValue) || bindingKind != kind)
                return false;

            _usedBindingNames.Add(name);
            value = bindingValue;
            return true;
        }

        /// <summary>
        /// Tells whether a MERGEFIELD instruction has a binding of any kind.
        /// </summary>
        private bool HasBinding(string instruction)
        {
            var name = MergeFieldName.GetBindingName(instruction);
            return name != null && _data.TryGetBinding(name, out _, out _);
        }

        /// <summary>
        /// Replaces each text field, at its position, with a run holding the bound value.
        /// Only fields are touched; text elsewhere in the document is never changed, even when it contains a key.
        /// </summary>
        private void MergeText()
        {
            foreach (var field in GetSimpleFields())
            {
                // A bound null value still replaces the field (with nothing).
                if (!TryFindBinding(BindingKind.Text, field, out var value))
                    continue;

                var text = Convert.ToString(value, GetCulture(field)) ?? string.Empty;

                // Without xml:space="preserve", leading and trailing spaces would be dropped.
                var element = new Text(text);
                if (text.Length > 0 && (char.IsWhiteSpace(text[0]) || char.IsWhiteSpace(text[^1])))
                    element.Space = SpaceProcessingModeValues.Preserve;

                FieldReplacement.ReplaceWithInline(field, element);
            }
        }

        /// <summary>
        /// Replaces each image field with the image, sized the same way in every part: the size given in the field name,
        /// or the image's own size scaled down to the text width.
        /// </summary>
        private void MergeImages()
        {
            var textWidth = TemplateFormatting.GetTextWidthInPixels(_main);

            foreach (var (part, root) in GetParts())
            {
                foreach (var field in root.Descendants<SimpleField>().ToList())
                {
                    if (!TryFindBinding(BindingKind.Image, field, out var value) || value is not byte[] data)
                        continue;

                    var config = new ImageElementConfig { MaxWidth = textWidth };
                    config.ParseConfig(MergeFieldName.Get(field));

                    // Created on the first image only: scanning for existing ids loads parts that may otherwise stay untouched.
                    _ids ??= new IdAllocator(GetDrawingRoots());
                    var image = ImageElement.CreateImage(data, part, _ids.NextDrawingId(), IdAllocator.NextRelationshipId(part), config);
                    if (image != null)
                        FieldReplacement.ReplaceWithInline(field, image);
                }
            }
        }

        /// <summary>
        /// Replaces each table field with a table built from the bound data.
        /// </summary>
        private void MergeTables()
        {
            var textWidth = TemplateFormatting.GetTextWidthInPixels(_main);

            foreach (var field in GetSimpleFields())
            {
                if (!TryFindBinding(BindingKind.Table, field, out var value) || value is not TableValue table)
                    continue;

                FieldReplacement.ReplaceWithBlock(field, TableBuilder.Create(table, field, _main, textWidth, GetCulture(field)));
            }
        }
    }
}
