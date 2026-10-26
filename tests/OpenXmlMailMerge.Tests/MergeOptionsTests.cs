using System;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Xml;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using W = DocumentFormat.OpenXml.Wordprocessing;
using static OpenXmlMailMerge.Tests.TestDocuments;

namespace OpenXmlMailMerge.Tests
{
    [TestClass]
    public class MergeOptionsTests
    {
        // ReSharper disable once UnusedAutoPropertyAccessor.Global
        public TestContext TestContext { get; set; }

        private byte[] SaveAndValidate(byte[] document, string suffix = null)
        {
            File.WriteAllBytes(TestOutput.PathFor(TestContext, suffix), document);
            AssertValid(document);
            return document;
        }

        private static string AllParts(byte[] document)
        {
            return string.Concat(ReadPart(document, "word/header1.xml"), ReadPart(document, "word/footer1.xml"), ReadPart(document, "word/document.xml"));
        }

        private static bool HasMergeField(string xml, string name)
        {
            return Regex.IsMatch(xml, $"MERGEFIELD\\s+{name}\\b");
        }

        [TestMethod]
        public void Defaults()
        {
            var options = new MergeOptions();

            Assert.AreEqual(UnboundFieldBehavior.Keep, options.UnboundFields);
            Assert.IsNull(options.Culture);
            Assert.AreEqual(50_000_000, options.MaxPartCharacters);
        }

        [TestMethod]
        public void InvalidOptions_Throw()
        {
            Assert.ThrowsException<ArgumentOutOfRangeException>(() => new DocumentMerger(new MergeOptions { MaxPartCharacters = -1 }));
            Assert.ThrowsException<ArgumentOutOfRangeException>(() => new DocumentMerger(new MergeOptions { UnboundFields = (UnboundFieldBehavior)42 }));
        }

        // ---- UnboundFields ----

        [TestMethod]
        public void UnboundFields_Keep_LeavesFieldsAndPlaceholders()
        {
            var result = new DocumentMerger().Merge(File.ReadAllBytes(BindFile), new MergeData().Text("BIND", "Reis"));

            var xml = AllParts(SaveAndValidate(result.Document));
            Assert.IsTrue(HasMergeField(xml, "HEADER"));
            StringAssert.Contains(xml, "«HEADER»");
        }

        [TestMethod]
        public void UnboundFields_Remove_RemovesFieldsAndPlaceholders_AndStillReportsThem()
        {
            var merger = new DocumentMerger(new MergeOptions { UnboundFields = UnboundFieldBehavior.Remove });

            var result = merger.Merge(File.ReadAllBytes(BindFile), new MergeData().Text("BIND", "Reis"));

            var xml = AllParts(SaveAndValidate(result.Document));
            foreach (var name in new[] { "HEADER", "IMG", "FOOTER", "TABLE" })
            {
                Assert.IsFalse(HasMergeField(xml, name), $"{name} is still a field.");
                Assert.IsFalse(xml.Contains($"«{name}»"), $"{name}'s placeholder is still shown.");
            }

            StringAssert.Contains(xml, "<w:t>Reis</w:t>");
            CollectionAssert.AreEqual((string[])["HEADER", "IMG", "FOOTER", "TABLE"], result.Report.UnboundFields.ToList());
        }

        [TestMethod]
        public void UnboundFields_Remove_RemovesComplexFields()
        {
            var template = CreateDocument([new W.Paragraph([new W.Run(new W.Text("Before ") { Space = SpaceProcessingModeValues.Preserve }), .. ComplexField("OTHER"), new W.Run(new W.Text("After"))])]);
            var merger = new DocumentMerger(new MergeOptions { UnboundFields = UnboundFieldBehavior.Remove });

            var result = merger.Merge(template, new MergeData().Text("NAME", "Jane"));

            using var document = WordprocessingDocument.Open(new MemoryStream(SaveAndValidate(result.Document)), false);
            Assert.AreEqual(0, Body(document).Descendants<W.FieldChar>().Count());
            Assert.AreEqual("Before After", Body(document).GetFirstChild<W.Paragraph>()?.InnerText);
        }

        [TestMethod]
        public void UnboundFields_Remove_LeavesNestedFields()
        {
            // { IF { MERGEFIELD NAME } = "x" "yes" "no" }: removing the inner field would change what the IF evaluates.
            OpenXmlElement[] field =
            [
                FieldMarker(W.FieldCharValues.Begin),
                new W.Run(new W.FieldCode(" IF ") { Space = SpaceProcessingModeValues.Preserve }),
                .. ComplexField("NAME"),
                new W.Run(new W.FieldCode(" = \"x\" \"yes\" \"no\" ") { Space = SpaceProcessingModeValues.Preserve }),
                FieldMarker(W.FieldCharValues.Separate),
                new W.Run(new W.Text("no")),
                FieldMarker(W.FieldCharValues.End)
            ];
            var merger = new DocumentMerger(new MergeOptions { UnboundFields = UnboundFieldBehavior.Remove });

            var result = merger.Merge(CreateDocument([new W.Paragraph(field)]), new MergeData().Text("OTHER", "x"));

            using var document = WordprocessingDocument.Open(new MemoryStream(SaveAndValidate(result.Document)), false);
            Assert.AreEqual(6, Body(document).Descendants<W.FieldChar>().Count());
            CollectionAssert.AreEqual((string[])["NAME"], result.Report.UnboundFields.ToList());
        }

        [TestMethod]
        public void UnboundFields_Throw_NamesTheFirstFieldAndItsPart()
        {
            var merger = new DocumentMerger(new MergeOptions { UnboundFields = UnboundFieldBehavior.Throw });

            var exception = Assert.ThrowsException<MergeException>(() => merger.Merge(File.ReadAllBytes(BindFile), new MergeData().Text("BIND", "Reis")));

            Assert.AreEqual("HEADER", exception.FieldName);
            Assert.AreEqual("/word/header1.xml", exception.PartUri);
            StringAssert.Contains(exception.Message, "HEADER, IMG, FOOTER, TABLE");
        }

        [TestMethod]
        public async Task UnboundFields_Throw_WritesNothingToTheOutput()
        {
            var merger = new DocumentMerger(new MergeOptions { UnboundFields = UnboundFieldBehavior.Throw });
            using var template = File.OpenRead(BindFile);
            using var output = new MemoryStream();

            await Assert.ThrowsExceptionAsync<MergeException>(() => merger.MergeAsync(template, output, new MergeData().Text("BIND", "Reis")));

            Assert.AreEqual(0, output.Length);
        }

        [TestMethod]
        public void UnboundFields_Throw_WhenEveryFieldIsBound_Merges()
        {
            var merger = new DocumentMerger(new MergeOptions { UnboundFields = UnboundFieldBehavior.Throw });
            var data = new MergeData()
                .Text("HEADER", "Vagner")
                .Text("FOOTER", "Lucas")
                .Text("BIND", "Reis")
                .Image("IMG", File.ReadAllBytes(ImageFile))
                .Table("TABLE", new TableValue(["COL1"], [["ROW 0"]]));

            var result = merger.Merge(File.ReadAllBytes(BindFile), data);

            SaveAndValidate(result.Document);
            Assert.AreEqual(0, result.Report.UnboundFields.Count);
        }

        // ---- Culture ----

        [TestMethod]
        public void Culture_OverridesTheTemplatesLanguage()
        {
            // BIND.docx's default language is pt-BR, which would give "1234,5" and "25/10/2026".
            var merger = new DocumentMerger(new MergeOptions { Culture = CultureInfo.GetCultureInfo("en-US") });
            var data = new MergeData()
                .Text("BIND", 1234.5m)
                .Table("TABLE", new TableValue(["Date"], [[new DateTime(2026, 10, 25)]]));

            var body = ReadPart(SaveAndValidate(merger.Merge(File.ReadAllBytes(BindFile), data).Document), "word/document.xml");

            StringAssert.Contains(body, "<w:t>1234.5</w:t>");
            StringAssert.Contains(body, "10/25/2026");
        }

        // ---- MaxPartCharacters ----

        private static byte[] TemplateWithLargeText(int characters, bool inHeader = false)
        {
            var paragraph = new W.Paragraph(new W.Run(new W.Text(new string('x', characters))), Field("NAME"));
            return inHeader
                ? CreateDocument([new W.Paragraph(Field("NAME"))], [paragraph])
                : CreateDocument([paragraph]);
        }

        [TestMethod]
        [DataRow(false, "/word/document.xml")]
        [DataRow(true, "/word/header")]
        public void MaxPartCharacters_PartTooLarge_ThrowsNamingThePart(bool inHeader, string partUri)
        {
            var merger = new DocumentMerger(new MergeOptions { MaxPartCharacters = 10_000 });

            var exception = Assert.ThrowsException<MergeException>(() => merger.Merge(TemplateWithLargeText(200_000, inHeader), new MergeData().Text("NAME", "Jane")));

            StringAssert.StartsWith(exception.PartUri, partUri);
            Assert.IsInstanceOfType(exception.InnerException, typeof(XmlException));
            StringAssert.Contains(exception.Message, "MaxPartCharacters (10000 characters)");
        }

        [TestMethod]
        public void MaxPartCharacters_Zero_DisablesTheCheck()
        {
            var merger = new DocumentMerger(new MergeOptions { MaxPartCharacters = 0 });

            var result = merger.Merge(TemplateWithLargeText(200_000), new MergeData().Text("NAME", "Jane"));

            StringAssert.Contains(ReadPart(result.Document, "word/document.xml"), "Jane");
        }

        /// <summary>
        /// Writes a minimal .docx whose document part is a single paragraph of the given number of characters.
        /// The part is streamed into the zip, so the test never holds the whole text in memory.
        /// </summary>
        private static byte[] DecompressionBomb(long characters)
        {
            var stream = new MemoryStream();
            using (var zip = new ZipArchive(stream, ZipArchiveMode.Create, leaveOpen: true))
            {
                void Write(string name, string content)
                {
                    using var writer = new StreamWriter(zip.CreateEntry(name, CompressionLevel.SmallestSize).Open(), new UTF8Encoding(false));
                    writer.Write(content);
                }

                Write("[Content_Types].xml",
                    "<?xml version=\"1.0\" encoding=\"UTF-8\"?><Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">"
                    + "<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>"
                    + "<Default Extension=\"xml\" ContentType=\"application/xml\"/>"
                    + "<Override PartName=\"/word/document.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml\"/></Types>");
                Write("_rels/.rels",
                    "<?xml version=\"1.0\" encoding=\"UTF-8\"?><Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
                    + "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"word/document.xml\"/></Relationships>");

                using var document = new StreamWriter(zip.CreateEntry("word/document.xml", CompressionLevel.SmallestSize).Open(), new UTF8Encoding(false));
                document.Write("<?xml version=\"1.0\" encoding=\"UTF-8\"?><w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:body><w:p><w:r><w:t>");
                var chunk = new string('x', 1 << 20);
                for (var written = 0L; written < characters; written += chunk.Length)
                    document.Write(chunk, 0, (int)Math.Min(chunk.Length, characters - written));
                document.Write("</w:t></w:r></w:p></w:body></w:document>");
            }

            return stream.ToArray();
        }

        [TestMethod]
        public void MaxPartCharacters_Default_StopsADecompressionBomb()
        {
            var bomb = DecompressionBomb(MergeOptions.DefaultMaxPartCharacters + 1);
            Assert.IsTrue(bomb.Length < 1_000_000, $"The template is {bomb.Length} bytes; the point is that it's small.");

            var exception = Assert.ThrowsException<MergeException>(() => new DocumentMerger().Merge(bomb, new MergeData()));

            Assert.AreEqual("/word/document.xml", exception.PartUri);
        }
    }
}
