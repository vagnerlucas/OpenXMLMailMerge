using System;
using System.Data;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using W = DocumentFormat.OpenXml.Wordprocessing;

using static OpenXmlMailMerge.Tests.TestDocuments;

namespace OpenXmlMailMerge.Tests
{
    [TestClass]
    public class DocumentMergerTests
    {
        // ReSharper disable once UnusedAutoPropertyAccessor.Global
        public TestContext TestContext { get; set; }

        private static byte[] Bind()
        {
            return File.ReadAllBytes(BindFile);
        }

        private static MergeData BindData()
        {
            return new MergeData()
                .Text("HEADER", "Vagner")
                .Text("FOOTER", "Lucas")
                .Text("BIND", "Reis")
                .Image("IMG", File.ReadAllBytes(ImageFile))
                .Table("TABLE", new TableValue(["COL1", "COL2"], [["ROW 0", 1234.5m]]));
        }

        private byte[] SaveAndValidate(byte[] document, string suffix = null)
        {
            File.WriteAllBytes(TestOutput.PathFor(TestContext, suffix), document);
            AssertValid(document);
            return document;
        }

        private static void AssertBindMerged(byte[] document)
        {
            var body = ReadPart(document, "word/document.xml");
            StringAssert.Contains(body, "<w:t>Reis</w:t>");
            StringAssert.Contains(body, "COL1");
            // BIND.docx's default language is pt-BR.
            StringAssert.Contains(body, "1234,5");
            StringAssert.Contains(body, "<w:drawing>");
            StringAssert.Contains(ReadPart(document, "word/header1.xml"), "Vagner");
            StringAssert.Contains(ReadPart(document, "word/footer1.xml"), "Lucas");
        }

        [TestMethod]
        public void Merge_BindTemplate_MergesTextImageAndTable()
        {
            var result = new DocumentMerger().Merge(Bind(), BindData());

            AssertBindMerged(SaveAndValidate(result.Document));
            Assert.AreEqual(0, result.Report.UnboundFields.Count);
            Assert.AreEqual(0, result.Report.UnusedBindings.Count);
        }

        [TestMethod]
        public async Task MergeAsync_Streams_WritesMergedDocument_AndLeavesStreamsOpen()
        {
            using var template = new MemoryStream(Bind());
            using var output = new MemoryStream();

            var report = await new DocumentMerger().MergeAsync(template, output, BindData());

            Assert.IsTrue(template.CanRead && output.CanWrite, "Neither stream may be closed.");
            AssertBindMerged(SaveAndValidate(output.ToArray()));
            Assert.AreEqual(0, report.UnboundFields.Count);
        }

        [TestMethod]
        public void Report_ListsUnboundFields_AndUnusedBindings()
        {
            var data = new MergeData().Text("bind", "Reis").Text("Adress", "Main Street");

            var result = new DocumentMerger().Merge(Bind(), data);

            SaveAndValidate(result.Document);
            // Headers, then footers, then the body; IMG appears in all three but is listed once.
            CollectionAssert.AreEqual((string[])["HEADER", "IMG", "FOOTER", "TABLE"], result.Report.UnboundFields.ToList());
            CollectionAssert.AreEqual((string[])["Adress"], result.Report.UnusedBindings.ToList());
        }

        [TestMethod]
        public void Report_BoundFieldThatCantBeMerged_IsUnboundAndUnused()
        {
            // { IF { MERGEFIELD NAME } = "x" "yes" "no" }: a field nested in another one is left as it is.
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

            var result = new DocumentMerger().Merge(CreateDocument([new W.Paragraph(field)]), new MergeData().Text("NAME", "Jane"));

            SaveAndValidate(result.Document);
            CollectionAssert.AreEqual((string[])["NAME"], result.Report.UnboundFields.ToList());
            CollectionAssert.AreEqual((string[])["NAME"], result.Report.UnusedBindings.ToList());
        }

        [TestMethod]
        public void Merge_WithoutBindings_LeavesTemplateAsItIs()
        {
            var result = new DocumentMerger().Merge(Bind(), new MergeData());

            var body = ReadPart(SaveAndValidate(result.Document), "word/document.xml");
            StringAssert.Contains(body, "«BIND»");
            CollectionAssert.AreEqual((string[])["HEADER", "IMG", "FOOTER", "BIND", "TABLE"], result.Report.UnboundFields.ToList());
        }

        [TestMethod]
        public void Merge_TableValue_WithRepeatedCaptionsAndNulls()
        {
            var table = new TableValue(["Total", "Total"], [[1, null], [null, 2]]);

            var result = new DocumentMerger().Merge(CreateTemplate("TABLE"), new MergeData().Table("TABLE", table));

            using var document = WordprocessingDocument.Open(new MemoryStream(SaveAndValidate(result.Document)), false);
            var rows = Body(document).Descendants<W.TableRow>()
                .Select(row => string.Join("|", row.Elements<W.TableCell>().Select(cell => cell.InnerText)))
                .ToList();
            CollectionAssert.AreEqual((string[])["Total|Total", "1|", "|2"], rows);
        }

        [TestMethod]
        [DataRow(new byte[0])]
        [DataRow(new byte[] { 1, 2, 3, 4 })]
        [DataRow(new byte[] { 0x50, 0x4B, 0x03, 0x04, 0, 0 })]
        public void Merge_InvalidTemplate_ThrowsMergeException(byte[] template)
        {
            var exception = Assert.ThrowsException<MergeException>(() => new DocumentMerger().Merge(template, new MergeData()));

            Assert.IsNotNull(exception.InnerException);
        }

        [TestMethod]
        public async Task MergeAsync_InvalidTemplate_ThrowsMergeException()
        {
            using var template = new MemoryStream([1, 2, 3, 4]);

            await Assert.ThrowsExceptionAsync<MergeException>(() => new DocumentMerger().MergeAsync(template, new MemoryStream(), new MergeData()));
        }

        [TestMethod]
        public void Merger_IsReusable_AcrossThreads()
        {
            var merger = new DocumentMerger();
            var data = BindData();
            var template = Bind();

            var results = new MergeResult[8];
            Parallel.For(0, results.Length, i => results[i] = merger.Merge(template, data));

            foreach (var result in results)
            {
                AssertValid(result.Document);
                AssertBindMerged(result.Document);
            }
        }

        [TestMethod]
        public async Task Merge_SameInputTwice_GivesIdenticalBytes()
        {
            var merger = new DocumentMerger();

            var first = merger.Merge(Bind(), BindData()).Document;
            var second = merger.Merge(Bind(), BindData()).Document;
            using var output = new MemoryStream();
            await merger.MergeAsync(new MemoryStream(Bind()), output, BindData());

            CollectionAssert.AreEqual(first, second);
            CollectionAssert.AreEqual(first, output.ToArray());
        }

        [TestMethod]
        public void Merge_EveryZipEntry_HasTheTemplatesTimestamp()
        {
            // The package stamps the entries a merge adds with the current time; the merge replaces that with the
            // template's own timestamp, so the output doesn't depend on when it ran.
            DateTimeOffset expected;
            using (var template = new ZipArchive(new MemoryStream(Bind()), ZipArchiveMode.Read))
                expected = template.GetEntry("word/document.xml")!.LastWriteTime;

            var document = SaveAndValidate(new DocumentMerger().Merge(Bind(), BindData()).Document);

            using var zip = new ZipArchive(new MemoryStream(document), ZipArchiveMode.Read);
            Assert.IsTrue(zip.Entries.Any(entry => entry.FullName.StartsWith("media/")), "The merge should have added images.");
            foreach (var entry in zip.Entries)
                Assert.AreEqual(expected, entry.LastWriteTime, entry.FullName);
        }

        private static uint[] DrawingIds(byte[] document)
        {
            using var package = WordprocessingDocument.Open(new MemoryStream(document), false);
            var main = package.MainDocumentPart!;
            OpenXmlElement[] roots = [main.Document, .. main.HeaderParts.Select(p => p.Header), .. main.FooterParts.Select(p => p.Footer)];
            return [.. roots.SelectMany(root => root.Descendants<DW.DocProperties>()).Select(properties => properties.Id!.Value).Order()];
        }

        [TestMethod]
        public void Images_GetUniqueDrawingIds_AcrossHeaderFooterAndBody()
        {
            var document = SaveAndValidate(new DocumentMerger().Merge(Bind(), new MergeData().Image("IMG", File.ReadAllBytes(ImageFile))).Document);

            // BIND.docx has no drawing yet, and IMG appears in the header, the footer and the body.
            CollectionAssert.AreEqual(new uint[] { 1, 2, 3 }, DrawingIds(document));
            StringAssert.Contains(ReadPart(document, "word/document.xml"), "name=\"Picture 3\"");
        }

        [TestMethod]
        public void Images_DrawingIds_ContinueAfterTheTemplatesHighest()
        {
            var merger = new DocumentMerger();
            var image = File.ReadAllBytes(ImageFile);

            // A template that already holds a drawing, whose id is then raised to 41.
            var withDrawing = merger.Merge(CreateTemplate("FIRST", "SECOND"), new MergeData().Image("FIRST", image)).Document;
            var stream = new MemoryStream();
            stream.Write(withDrawing);
            using (var package = WordprocessingDocument.Open(stream, true))
                package.MainDocumentPart!.Document.Descendants<DW.DocProperties>().Single().Id = 41U;

            var document = SaveAndValidate(merger.Merge(stream.ToArray(), new MergeData().Image("SECOND", image)).Document);

            CollectionAssert.AreEqual(new uint[] { 41, 42 }, DrawingIds(document));
        }

        [TestMethod]
        public void Images_RelationshipId_IsTheFirstOneTheTemplateDoesntUse()
        {
            string expected;
            using (var template = WordprocessingDocument.Open(new MemoryStream(Bind()), false))
            {
                var used = template.MainDocumentPart!.Parts.Select(pair => pair.RelationshipId).ToHashSet();
                expected = Enumerable.Range(1, used.Count + 1).Select(i => $"rId{i}").First(id => !used.Contains(id));
            }

            var document = SaveAndValidate(new DocumentMerger().Merge(Bind(), new MergeData().Image("IMG", File.ReadAllBytes(ImageFile))).Document);

            using var package = WordprocessingDocument.Open(new MemoryStream(document), false);
            var main = package.MainDocumentPart!;
            var embed = main.Document.Descendants<DocumentFormat.OpenXml.Drawing.Blip>().Single().Embed!.Value;
            Assert.AreEqual(expected, embed);
            Assert.IsInstanceOfType(main.GetPartById(embed), typeof(ImagePart));
        }

        [TestMethod]
        public void MergeData_SameNameTwice_Throws_IgnoringCase()
        {
            var data = new MergeData().Text("Name", "Jane");

            Assert.ThrowsException<ArgumentException>(() => data.Image("NAME", new byte[] { 1 }));
            CollectionAssert.AreEqual((string[])["Name"], data.Names.ToList());
        }

        [TestMethod]
        public void MergeData_InvalidArguments_Throw()
        {
            var data = new MergeData();

            Assert.ThrowsException<ArgumentNullException>(() => data.Text(null, "x"));
            Assert.ThrowsException<ArgumentException>(() => data.Text(" ", "x"));
            Assert.ThrowsException<ArgumentException>(() => data.Image("IMG", ReadOnlyMemory<byte>.Empty));
            Assert.ThrowsException<ArgumentNullException>(() => data.Table("TABLE", null));
            Assert.AreEqual(0, data.Names.Count);
        }

        [TestMethod]
        public void MergeData_CopiesImageBytes()
        {
            var image = File.ReadAllBytes(ImageFile);
            var data = new MergeData().Image("IMG", image);

            Array.Clear(image);
            var result = new DocumentMerger().Merge(CreateTemplate("IMG"), data);

            using var document = WordprocessingDocument.Open(new MemoryStream(SaveAndValidate(result.Document)), false);
            using var stream = document.MainDocumentPart?.ImageParts.Single().GetStream();
            using var copy = new MemoryStream();
            stream?.CopyTo(copy);
            CollectionAssert.AreEqual(File.ReadAllBytes(ImageFile), copy.ToArray());
        }

        [TestMethod]
        public void TableValue_InvalidShape_Throws()
        {
            Assert.ThrowsException<ArgumentException>(() => new TableValue([], []));
            Assert.ThrowsException<ArgumentException>(() => new TableValue(["A", "B"], [["only one"]]));
            Assert.ThrowsException<ArgumentNullException>(() => new TableValue(["A"], [null]));
        }

        [TestMethod]
        public void TableValue_FromDataTable_KeepsCaptionsAndValues()
        {
            var dataTable = new DataTable();
            dataTable.Columns.Add(new DataColumn("a", typeof(string)) { Caption = "Name" });
            dataTable.Columns.Add(new DataColumn("b", typeof(int)) { Caption = "Age" });
            dataTable.Rows.Add("Jane", 42);
            dataTable.Rows.Add("John", DBNull.Value);

            var table = TableValue.From(dataTable);

            CollectionAssert.AreEqual((string[])["Name", "Age"], table.Columns.ToList());
            CollectionAssert.AreEqual(new object[] { "Jane", 42 }, table.Rows[0].ToList());
            CollectionAssert.AreEqual(new object[] { "John", null }, table.Rows[1].ToList());
        }
    }
}
