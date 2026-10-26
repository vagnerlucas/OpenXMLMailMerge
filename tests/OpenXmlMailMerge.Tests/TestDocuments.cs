using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using System.Diagnostics;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace OpenXmlMailMerge.Tests
{
    /// <summary>
    /// Fixture paths, template builders and assertions shared by the test classes.
    /// </summary>
    internal static class TestDocuments
    {
        internal static readonly string BindFile = Path.Combine("Fixtures", "BIND.docx");
        internal static readonly string ImageFile = Path.Combine("Fixtures", "image.png");
        internal static readonly string ImagesDirectory = Path.Combine("Fixtures", "Images");

        internal static void AssertValid(byte[] data)
        {
            using var document = WordprocessingDocument.Open(new MemoryStream(data), false);
            var errors = new OpenXmlValidator().Validate(document)
            .Select(e => $"{e.Part?.Uri} {e.Path?.XPath}: {e.Description}")
            .ToList();

            Assert.AreEqual(0, errors.Count, "Invalid document:" + Environment.NewLine + string.Join(Environment.NewLine, errors));
        }

        internal static string ReadPart(byte[] package, string partName)
        {
            using var zip = new ZipArchive(new MemoryStream(package), ZipArchiveMode.Read);
            using var reader = new StreamReader(zip.GetEntry(partName)!.Open());
            return reader.ReadToEnd();
        }

        /// <summary>
        /// A w:fldSimple with the instruction "MERGEFIELD name" and the display text «name».
        /// </summary>
        internal static W.SimpleField Field(string name)
        {
            // Word quotes a field name that contains spaces.
            var quoted = name.Contains(' ') ? $"\"{name}\"" : name;
            return new W.SimpleField(new W.Run(new W.Text($"«{name}»"))) { Instruction = $" MERGEFIELD {quoted} \\* MERGEFORMAT " };
        }

        /// <summary>
        /// A field whose result run has the given formatting.
        /// </summary>
        internal static W.SimpleField Field(string name, W.RunProperties runProperties)
        {
            var field = Field(name);
            field.GetFirstChild<W.Run>()?.PrependChild(runProperties);
            return field;
        }

        /// <summary>
        /// A4 page with Word's default margins, the same as BIND.docx.
        /// </summary>
        internal static W.SectionProperties PageSetup(params OpenXmlElement[] headerReferences)
        {
            return new W.SectionProperties(headerReferences.Concat([
                new W.PageSize { Width = 11906, Height = 16838 },
                new W.PageMargin { Top = 1417, Right = 1701, Bottom = 1417, Left = 1701, Header = 708, Footer = 708, Gutter = 0 }
            ]));
        }

        /// <summary>
        /// Creates a template from body blocks. Like documents saved by Word, the body ends with a w:sectPr.
        /// </summary>
        /// <param name="blocks">Body content.</param>
        /// <param name="headers">Content of each header part, referenced in order as the default and first-page header.</param>
        /// <exception cref="ArgumentNullException"></exception>
        internal static byte[] CreateDocument(OpenXmlElement[] blocks, params W.Paragraph[][] headers)
        {
            return CreateDocument(blocks, headers, []);
        }

        /// <summary>
        /// Creates a template from body blocks, headers and footers.
        /// </summary>
        /// <param name="blocks">Body content.</param>
        /// <param name="headers">Content of each header part, referenced in order as the default and first-page header.</param>
        /// <param name="footers">Content of each footer part, referenced in order as the default and first-page footer.</param>
        /// <param name="styles">Style definitions, or null for a document without a styles part.</param>
        internal static byte[] CreateDocument(OpenXmlElement[] blocks, W.Paragraph[][] headers, W.Paragraph[][] footers, W.Styles styles = null)
        {
            var stream = new MemoryStream();
            using (var document = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document))
            {
                var main = document.AddMainDocumentPart();
                if (styles != null)
                    main.AddNewPart<StyleDefinitionsPart>().Styles = styles;
                var types = new[] { W.HeaderFooterValues.Default, W.HeaderFooterValues.First };
                var headerReferences = headers.Select((content, i) =>
                {
                    var part = main.AddNewPart<HeaderPart>();
                    // ReSharper disable once CoVariantArrayConversion
                    part.Header = new W.Header(content);
                    return (OpenXmlElement)new W.HeaderReference { Type = types[i], Id = main.GetIdOfPart(part) };
                });
                var footerReferences = footers.Select((content, i) =>
                {
                    var part = main.AddNewPart<FooterPart>();
                    // ReSharper disable once CoVariantArrayConversion
                    part.Footer = new W.Footer(content);
                    return (OpenXmlElement)new W.FooterReference { Type = types[i], Id = main.GetIdOfPart(part) };
                });

                main.Document = new W.Document(new W.Body(blocks.Append(PageSetup([.. headerReferences, .. footerReferences]))));
            }

            return stream.ToArray();
        }

        /// <summary>
        /// Creates a template whose body has one paragraph per field.
        /// </summary>
        internal static byte[] CreateTemplate(params string[] fieldNames)
        {
            return CreateDocument([.. fieldNames.Select(name => (OpenXmlElement)new W.Paragraph(Field(name)))]);
        }

        internal static W.Run FieldMarker(W.FieldCharValues type)
        {
            return new W.Run(new W.FieldChar { FieldCharType = type });
        }

        /// <summary>
        /// The runs of a complex field, as Word writes them: begin, instruction runs, separate, a bold result «name», end.
        /// </summary>
        /// <param name="name">Field name.</param>
        /// <param name="withResult">False for a field without a separate marker and result.</param>
        /// <param name="instructionParts">The instruction split over several runs; by default a single " MERGEFIELD name " run.</param>
        internal static OpenXmlElement[] ComplexField(string name, bool withResult = true, params string[] instructionParts)
        {
            string[] parts = instructionParts.Length > 0 ? instructionParts : [$" MERGEFIELD {name} \\* MERGEFORMAT "];

            var runs = new List<OpenXmlElement> { FieldMarker(W.FieldCharValues.Begin) };
            runs.AddRange(parts.Select(part => new W.Run(new W.FieldCode(part) { Space = SpaceProcessingModeValues.Preserve })));
            if (withResult)
            {
                runs.Add(FieldMarker(W.FieldCharValues.Separate));
                runs.Add(new W.Run(new W.RunProperties(new W.Bold()), new W.Text($"«{name}»")));
            }

            runs.Add(FieldMarker(W.FieldCharValues.End));
            return [.. runs];
        }

        internal static W.Body Body(WordprocessingDocument document)
        {
            var body = document.MainDocumentPart?.Document.Body;
            Debug.Assert(body != null, nameof(body) + " != null");
            return body;
        }
    }
}
