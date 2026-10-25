using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Linq;
using System.Data;
using System.Globalization;
using System.IO;
using OpenXmlMailMerge.Fields;
using OpenXmlMailMerge.Elements;

namespace OpenXmlMailMerge
{
    /// <summary>
    /// A OpenXmlDocument as word processor.
    /// </summary>
    internal sealed class WordprocessingMl : OpenXmlDocument
    {
        /// <summary>
        /// Opens the WordprocessingML package on the given stream.
        /// AutoSave is off: changes reach the stream only when the document is saved explicitly.
        /// </summary>
        /// <param name="stream">Expandable, seekable stream holding the package.</param>
        /// <returns>The content handler of the opened package.</returns>
        private protected override DocumentHandler OpenContent(Stream stream)
        {
            var document = WordprocessingDocument.Open(stream, true, new OpenSettings { AutoSave = false });

            return new DocumentHandler
            {
                Document = document,
                Content = document.MainDocumentPart
            };
        }

        /// <summary>
        /// Gets the field name from a MERGEFIELD instruction, including its parameters.
        /// For example, " MERGEFIELD  IMG@WIDTH=100 \* MERGEFORMAT " gives "IMG@WIDTH=100".
        /// </summary>
        /// <param name="field">The field.</param>
        /// <returns>The field name, or null when the field is not a MERGEFIELD.</returns>
        private static string GetFieldName(SimpleField field)
        {
            return GetFieldName(field.Instruction?.Value);
        }

        /// <summary>
        /// Gets the field name from a MERGEFIELD instruction, including its parameters.
        /// </summary>
        /// <param name="instruction">The field instruction, for example " MERGEFIELD  IMG \* MERGEFORMAT ".</param>
        /// <returns>The field name, or null when the instruction is not a MERGEFIELD.</returns>
        private static string GetFieldName(string instruction)
        {
            const string keyword = "MERGEFIELD";

            instruction = instruction?.Trim();
            if (instruction == null
                || instruction.Length <= keyword.Length
                || !instruction.StartsWith(keyword, StringComparison.OrdinalIgnoreCase)
                || !char.IsWhiteSpace(instruction[keyword.Length]))
                return null;

            var rest = instruction[keyword.Length..].TrimStart();
            string name;
            if (rest.StartsWith('"'))
            {
                var end = rest.IndexOf('"', 1);
                name = end < 0 ? rest[1..] : rest[1..end];
            }
            else
            {
                var end = rest.IndexOfAny([' ', '\t']);
                name = end < 0 ? rest : rest[..end];
            }

            return name.Length == 0 ? null : name;
        }

        /// <summary>
        /// Gets the value bound to a field. The binding key must equal the field name
        /// without its parameters (case-insensitive): key "IMG" matches IMG and IMG@WIDTH=100, not IMG_LOGO.
        /// </summary>
        /// <param name="dataType">The mailmerge data type.</param>
        /// <param name="field">The field.</param>
        /// <param name="value">The bound value, which can itself be null.</param>
        /// <returns>True when a binding matches the field.</returns>
        private bool TryFindBinding(MailMergeDataTypeEnum dataType, SimpleField field, out object value)
        {
            value = null;

            var name = GetFieldName(field);
            if (name == null || !Dictionary.TryGetValue(dataType, out var bindings))
                return false;

            name = FieldParameters.GetBindingName(name);

            foreach (var binding in bindings)
            {
                if (!string.Equals(binding.Key, name, StringComparison.OrdinalIgnoreCase))
                    continue;

                value = binding.Value;
                return true;
            }

            return false;
        }

        /// <summary>
        /// Tells whether a MERGEFIELD instruction has a binding of any data type.
        /// </summary>
        /// <param name="instruction">The field instruction.</param>
        /// <returns>True when a binding key matches the field name.</returns>
        private bool HasBinding(string instruction)
        {
            var name = GetFieldName(instruction);
            if (name == null)
                return false;

            name = FieldParameters.GetBindingName(name);
            return Dictionary.Values.Any(bindings => bindings.Keys.Any(key => string.Equals(key, name, StringComparison.OrdinalIgnoreCase)));
        }

        /// <summary>
        /// Rewrites every complex MERGEFIELD that has a binding as an equivalent w:fldSimple, so all fields are merged the same way.
        /// </summary>
        private void ConvertComplexFields()
        {
            var roots = ContentManager.Content.HeaderParts.Select(p => (OpenXmlElement)p.Header)
                .Concat(ContentManager.Content.FooterParts.Select(p => p.Footer))
                .Append(ContentManager.Content.Document);

            foreach (var root in roots)
                ComplexFields.ConvertToSimpleFields(root, HasBinding);
        }

        /// <summary>
        /// Gets the value bound to a field.
        /// </summary>
        /// <param name="dataType">The mailmerge data type.</param>
        /// <param name="field">The field.</param>
        /// <returns>The bound value, or null when no binding matches.</returns>
        private object FindBinding(MailMergeDataTypeEnum dataType, SimpleField field)
        {
            TryFindBinding(dataType, field, out var value);
            return value;
        }

        /// <summary>
        /// Gets the image bound to a field.
        /// </summary>
        /// <param name="field">The field.</param>
        /// <returns>Array of byte with the value from data dictionary.</returns>
        private byte[] TryReplaceImage(SimpleField field)
        {
            return FindBinding(MailMergeDataTypeEnum.Image, field) as byte[];
        }

        /// <summary>
        /// Gets the text bound to a field.
        /// </summary>
        /// <param name="field">The field.</param>
        /// <param name="text">The text value; a null value gives an empty string.</param>
        /// <returns>True when a text binding matches the field.</returns>
        private bool TryReplaceText(SimpleField field, out string text)
        {
            // A bound null value still replaces the field (with nothing), so "no binding" and "null" must stay distinct.
            if (!TryFindBinding(MailMergeDataTypeEnum.Text, field, out var value))
            {
                text = null;
                return false;
            }

            text = Convert.ToString(value, GetCulture(field)) ?? string.Empty;
            return true;
        }

        /// <summary>
        /// Gets the table bound to a field.
        /// </summary>
        /// <param name="field">The field.</param>
        /// <returns>DataTable with the value from data dictionary.</returns>
        private DataTable TryReplaceTable(SimpleField field)
        {
            return FindBinding(MailMergeDataTypeEnum.Table, field) as DataTable;
        }

        /// <summary>
        /// Executes text replacement: each text field is replaced, at its position, by a run with the bound text.
        /// Only fields are touched; text elsewhere in the document is never changed, even when it contains a key.
        /// </summary>
        private void ProcessText()
        {
            var roots = ContentManager.Content.HeaderParts.Select(p => (OpenXmlElement)p.Header)
                .Concat(ContentManager.Content.FooterParts.Select(p => p.Footer))
                .Append(ContentManager.Content.Document);

            // Snapshot the fields first: replacing a field changes the tree being searched.
            foreach (var field in roots.SelectMany(r => r.Descendants<SimpleField>()).ToList())
            {
                if (!TryReplaceText(field, out var text))
                    continue;

                // Without xml:space="preserve", leading and trailing spaces would be dropped.
                var element = new Text(text);
                if (text.Length > 0 && (char.IsWhiteSpace(text[0]) || char.IsWhiteSpace(text[^1])))
                    element.Space = SpaceProcessingModeValues.Preserve;

                ReplaceFieldWithInline(field, element);
            }
        }

        /// <summary>
        /// Gets the width between the margins of the document's last section, in pixels at 96 DPI (15 twips per pixel).
        /// Headers and footers span the same margins as the body.
        /// The last section is the one in the body's own w:sectPr, which covers the whole document unless it has section breaks.
        /// </summary>
        /// <returns>The text width in pixels, or 0 when the page size is not defined.</returns>
        private int GetTextWidthInPixels()
        {
            var section = ContentManager.Content.Document.Body?.Elements<SectionProperties>().LastOrDefault();
            var pageWidth = section?.GetFirstChild<PageSize>()?.Width?.Value;
            if (pageWidth == null)
                return 0;

            var margin = section.GetFirstChild<PageMargin>();
            var textWidth = (long)pageWidth.Value - (margin?.Left?.Value ?? 0) - (margin?.Right?.Value ?? 0);

            return textWidth > 0 ? (int)(textWidth / 15) : 0;
        }

        /// <summary>
        /// Executes image replacement.
        /// </summary>
        private void ProcessImage()
        {
            var textWidth = GetTextWidthInPixels();

            // Each image becomes a part of the part that holds its field, so the part travels with its root element.
            var parts = ContentManager.Content.HeaderParts.Select(p => ((OpenXmlPart)p, (OpenXmlElement)p.Header))
                .Concat(ContentManager.Content.FooterParts.Select(p => ((OpenXmlPart)p, (OpenXmlElement)p.Footer)))
                .Append(((OpenXmlPart)ContentManager.Content, (OpenXmlElement)ContentManager.Content.Document));

            foreach (var (part, root) in parts)
            {
                // Snapshot the fields first: replacing a field changes the tree being searched.
                foreach (var field in root.Descendants<SimpleField>().ToList())
                {
                    var data = TryReplaceImage(field);
                    if (data == null)
                        continue;

                    var image = CreateImage(data, part, field, textWidth);
                    if (image != null)
                        ReplaceFieldWithInline(field, image);
                }
            }
        }

        /// <summary>
        /// Replaces a field with a run holding inline content (text or an image), at the field's position.
        /// The run keeps the formatting of the field's first result run.
        /// </summary>
        /// <param name="field">The field to replace.</param>
        /// <param name="content">The inline content.</param>
        private static void ReplaceFieldWithInline(SimpleField field, OpenXmlElement content)
        {
            var run = new Run();

            var properties = field.Descendants<RunProperties>().FirstOrDefault();
            if (properties != null)
                run.Append(properties.CloneNode(true));

            run.Append(content);

            field.InsertBeforeSelf(run);
            field.Remove();
        }

        /// <summary>
        /// Creates the drawing for an image field, sized the same way in every part:
        /// the size given in the field name, or the image's own size scaled down to the text width.
        /// </summary>
        /// <param name="data">Byte array of the image data.</param>
        /// <param name="part">Part that will hold the image (main document, header or footer).</param>
        /// <param name="field">The image field.</param>
        /// <param name="maxWidth">Maximum width in pixels for an image shown at its own size. 0 for no limit.</param>
        /// <returns>The drawing element.</returns>
        private static Drawing CreateImage(byte[] data, OpenXmlPart part, SimpleField field, int maxWidth)
        {
            var imageElementConfig = new ImageElementConfig { MaxWidth = maxWidth };
            imageElementConfig.ParseConfig(GetFieldName(field));
            return ElementBuilder.CreateImage(data, part, imageElementConfig);
        }

        /// <summary>
        /// Builds the table for a table field: a header row with the column captions, repeated on each page, then one row
        /// per data row. The template decides how it looks:
        /// <list type="bullet">
        /// <item>cell text takes the field's run formatting (font, size, bold, ...), and the document's defaults otherwise;</item>
        /// <item>a STYLE parameter in the field name (TABLE@STYLE=GridTable4) applies that table style, when the document
        /// defines it; otherwise the table gets single borders;</item>
        /// <item>values are formatted in the template's language (see <see cref="GetCulture"/>).</item>
        /// </list>
        /// </summary>
        /// <param name="data">The table data.</param>
        /// <param name="field">The table field.</param>
        /// <param name="textWidthInPixels">Width between the page margins, shared by the columns; 0 when unknown.</param>
        /// <returns>The table.</returns>
        private Table CreateTable(DataTable data, SimpleField field, int textWidthInPixels)
        {
            var columns = data.Columns.Count;
            var config = new TableElementConfig
            {
                StyleId = FindTableStyleId(field),
                // 15 twips per pixel at 96 DPI.
                ColumnWidth = textWidthInPixels > 0 ? textWidthInPixels * 15 / columns : 0
            };

            var table = ElementBuilder.CreateTable(columns, config);
            var runProperties = field.Descendants<RunProperties>().FirstOrDefault();
            var culture = GetCulture(field);

            var header = ElementBuilder.CreateRow(isHeader: true);
            foreach (DataColumn column in data.Columns)
                header.Append(ElementBuilder.CreateCell(column.Caption, runProperties));
            table.Append(header);

            foreach (DataRow row in data.Rows)
            {
                var tableRow = ElementBuilder.CreateRow();
                foreach (var value in row.ItemArray)
                    tableRow.Append(ElementBuilder.CreateCell(Convert.ToString(value, culture), runProperties));
                table.Append(tableRow);
            }

            return table;
        }

        /// <summary>
        /// Finds the table style named by the STYLE parameter of a field, by style id or by style name (ignoring case).
        /// A name with spaces needs a quoted field name: MERGEFIELD "TABLE@STYLE=Grid Table 4".
        /// </summary>
        /// <param name="field">The table field.</param>
        /// <returns>The style id, or null when no STYLE is given or the document doesn't define that table style.</returns>
        private string FindTableStyleId(SimpleField field)
        {
            var parameters = FieldParameters.Parse(GetFieldName(field));
            if (!parameters.TryGetValue("STYLE", out var name) || name.Length == 0)
                return null;

            var style = ContentManager.Content.StyleDefinitionsPart?.Styles?.Elements<Style>()
                .FirstOrDefault(s => s.Type?.Value == StyleValues.Table
                                     && (string.Equals(s.StyleId?.Value, name, StringComparison.OrdinalIgnoreCase)
                                         || string.Equals(s.StyleName?.Val?.Value, name, StringComparison.OrdinalIgnoreCase)));

            return style?.StyleId?.Value;
        }

        /// <summary>
        /// Gets the culture used to format the values merged into a field, from the template's language:
        /// the language set on the field's text, else the document's default language, else the invariant culture.
        /// The machine running the merge never changes the result.
        /// </summary>
        /// <param name="field">The field.</param>
        /// <returns>The culture.</returns>
        private CultureInfo GetCulture(SimpleField field)
        {
            var language = field.Descendants<RunProperties>().FirstOrDefault()?.Languages?.Val?.Value
                           ?? ContentManager.Content.StyleDefinitionsPart?.Styles?.DocDefaults?.RunPropertiesDefault?.RunPropertiesBaseStyle?.Languages?.Val?.Value;

            if (string.IsNullOrEmpty(language))
                return CultureInfo.InvariantCulture;

            try
            {
                return CultureInfo.GetCultureInfo(language);
            }
            catch (CultureNotFoundException)
            {
                return CultureInfo.InvariantCulture;
            }
        }

        /// <summary>
        /// Executes table replacement.
        /// </summary>
        private void ProcessTable()
        {
            var textWidth = GetTextWidthInPixels();
            var roots = ContentManager.Content.HeaderParts.Select(p => (OpenXmlElement)p.Header)
            .Concat(ContentManager.Content.FooterParts.Select(p => p.Footer))
            .Append(ContentManager.Content.Document);

            // Snapshot the fields first: placing a table changes the tree being searched.
            foreach (var field in roots.SelectMany(r => r.Descendants<SimpleField>()).ToList())
            {
                var data = TryReplaceTable(field);
                // A table needs at least one column: a row without cells is invalid.
                if (data == null || data.Columns.Count == 0)
                    continue;

                ReplaceFieldWithBlock(field, CreateTable(data, field, textWidth));
            }
        }

        /// <summary>
        /// Puts a block element (such as a table) where a field is, and removes the field.
        /// A block can't be inside a paragraph, so:
        /// <list type="bullet">
        /// <item>when the field is the only content of its paragraph, the block replaces the paragraph;</item>
        /// <item>otherwise the block goes right after the paragraph, which keeps its other content.</item>
        /// </list>
        /// The paragraph is kept when it carries a section break, or when it is the last element of its container
        /// (table cells, headers and footers must end with a paragraph). For the same reason, an empty paragraph is
        /// added when the block would otherwise end its container.
        /// </summary>
        /// <param name="field">The field to replace.</param>
        /// <param name="block">The block element.</param>
        private static void ReplaceFieldWithBlock(SimpleField field, OpenXmlElement block)
        {
            var paragraph = field.Ancestors<Paragraph>().FirstOrDefault();
            if (paragraph == null)
                return;

            field.Remove();

            var isEmpty = paragraph.Descendants<Text>().All(t => string.IsNullOrWhiteSpace(t.Text)) && !paragraph.Descendants<Drawing>().Any()
                                                                                                    && !paragraph.Descendants<SimpleField>().Any();
            if (!isEmpty)
            {
                paragraph.InsertAfterSelf(block);
            }
            else
            {
                paragraph.InsertBeforeSelf(block);
                if (paragraph.NextSibling() != null && !paragraph.Descendants<SectionProperties>().Any())
                    paragraph.Remove();
            }

            if (block.NextSibling() == null)
                block.InsertAfterSelf(new Paragraph());
        }

        /// <summary>
        /// Process the document applying the mail merge.
        /// </summary>
        public override void Process()
        {
            base.Process();

            // Processing types
            
            ConvertComplexFields();
            
            ProcessText();

            ProcessImage();

            ProcessTable();
        }
    }
}
