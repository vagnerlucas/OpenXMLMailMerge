using System.IO;
using System.Linq;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using W = DocumentFormat.OpenXml.Wordprocessing;
using static OpenXmlMailMerge.Tests.TestDocuments;

namespace OpenXmlMailMerge.Tests
{
    [TestClass]
    public class InspectTests
    {
        private static string Describe(TemplateField field)
        {
            return $"{field.Name}|{field.PartUri}|complex={field.IsComplex}|mergeable={field.IsMergeable}";
        }

        // ---- Inspect ----

        [TestMethod]
        public void Inspect_BindTemplate_ListsEveryOccurrenceInOrder()
        {
            var info = new DocumentMerger().Inspect(File.ReadAllBytes(BindFile));

            CollectionAssert.AreEqual(
                (string[])
                [
                    "HEADER|/word/header1.xml|complex=False|mergeable=True",
                    "IMG|/word/header1.xml|complex=False|mergeable=True",
                    "FOOTER|/word/footer1.xml|complex=False|mergeable=True",
                    "IMG|/word/footer1.xml|complex=False|mergeable=True",
                    "IMG|/word/document.xml|complex=False|mergeable=True",
                    "BIND|/word/document.xml|complex=False|mergeable=True",
                    "TABLE|/word/document.xml|complex=False|mergeable=True"
                ],
                info.Fields.Select(Describe).ToList());
        }

        [TestMethod]
        public void Inspect_ReadsParameters_IgnoringCase()
        {
            var field = new DocumentMerger().Inspect(CreateTemplate("IMG@WIDTH=120@HEIGHT=80")).Fields.Single();

            Assert.AreEqual("IMG", field.Name);
            Assert.AreEqual("120", field.Parameters["width"]);
            Assert.AreEqual("80", field.Parameters["HEIGHT"]);
            Assert.AreEqual(0, new DocumentMerger().Inspect(CreateTemplate("NAME")).Fields.Single().Parameters.Count);
        }

        [TestMethod]
        public void Inspect_TellsComplexAndUnmergeableFields()
        {
            // { IF { MERGEFIELD INNER } = "x" "yes" "no" }: the inner field is nested, so it can't be merged.
            OpenXmlElement[] nested =
            [
                FieldMarker(W.FieldCharValues.Begin),
                new W.Run(new W.FieldCode(" IF ") { Space = SpaceProcessingModeValues.Preserve }),
                .. ComplexField("INNER"),
                new W.Run(new W.FieldCode(" = \"x\" \"yes\" \"no\" ") { Space = SpaceProcessingModeValues.Preserve }),
                FieldMarker(W.FieldCharValues.Separate),
                new W.Run(new W.Text("no")),
                FieldMarker(W.FieldCharValues.End)
            ];
            var split = ComplexField("SPLIT");
            var template = CreateDocument(
            [
                new W.Paragraph(ComplexField("NAME")),
                new W.Paragraph(nested),
                new W.Paragraph(split[..3]),
                new W.Paragraph(split[3..]),
                new W.Paragraph(Field("PLAIN"))
            ]);

            var info = new DocumentMerger().Inspect(template);

            CollectionAssert.AreEqual(
                (string[])
                [
                    "NAME|/word/document.xml|complex=True|mergeable=True",
                    "INNER|/word/document.xml|complex=True|mergeable=False",
                    "SPLIT|/word/document.xml|complex=True|mergeable=False",
                    "PLAIN|/word/document.xml|complex=False|mergeable=True"
                ],
                info.Fields.Select(Describe).ToList());
        }

        [TestMethod]
        public void Inspect_SkipsFieldsThatArentMergeFields()
        {
            var template = CreateDocument(
            [
                new W.Paragraph(new W.SimpleField(new W.Run(new W.Text("1"))) { Instruction = " PAGE " }),
                new W.Paragraph([FieldMarker(W.FieldCharValues.Begin), new W.Run(new W.FieldCode(" DATE ")), FieldMarker(W.FieldCharValues.End)]),
                new W.Paragraph(Field("NAME"))
            ]);

            var info = new DocumentMerger().Inspect(template);

            CollectionAssert.AreEqual((string[])["NAME"], info.Fields.Select(field => field.Name).ToList());
        }

        [TestMethod]
        public void Inspect_ListsTheFieldsAMergeWithoutDataReportsAsUnbound()
        {
            var merger = new DocumentMerger();
            var template = File.ReadAllBytes(BindFile);

            var inspected = merger.Inspect(template).Fields.Select(field => field.Name).Distinct().ToList();
            var unbound = merger.Merge(template, new MergeData()).Report.UnboundFields.ToList();

            CollectionAssert.AreEqual(inspected, unbound);
        }

        [TestMethod]
        public void Inspect_InvalidTemplate_ThrowsMergeException()
        {
            Assert.ThrowsException<MergeException>(() => new DocumentMerger().Inspect([1, 2, 3, 4]));
        }

        // ---- ValidateOutput ----

        /// <summary>
        /// A template that is already schema-invalid: a table directly inside a paragraph. The merge keeps it as it is.
        /// </summary>
        private static byte[] InvalidTemplate()
        {
            return CreateDocument([new W.Paragraph(new W.Table(), Field("NAME"))]);
        }

        [TestMethod]
        public void ValidateOutput_IsOffByDefault()
        {
            var result = new DocumentMerger().Merge(InvalidTemplate(), new MergeData().Text("NAME", "Jane"));

            StringAssert.Contains(ReadPart(result.Document, "word/document.xml"), "Jane");
        }

        [TestMethod]
        public void ValidateOutput_ValidResult_Merges()
        {
            var merger = new DocumentMerger(new MergeOptions { ValidateOutput = true });

            var result = merger.Merge(File.ReadAllBytes(BindFile), new MergeData().Text("BIND", "Reis").Image("IMG", File.ReadAllBytes(ImageFile)));

            StringAssert.Contains(ReadPart(result.Document, "word/document.xml"), "Reis");
        }

        [TestMethod]
        public void ValidateOutput_InvalidResult_ThrowsNamingThePart()
        {
            var merger = new DocumentMerger(new MergeOptions { ValidateOutput = true });

            var exception = Assert.ThrowsException<MergeException>(() => merger.Merge(InvalidTemplate(), new MergeData().Text("NAME", "Jane")));

            Assert.AreEqual("/word/document.xml", exception.PartUri);
            StringAssert.Contains(exception.Message, "isn't valid");
        }

        [TestMethod]
        public async Task ValidateOutput_InvalidResult_WritesNothingToTheOutput()
        {
            var merger = new DocumentMerger(new MergeOptions { ValidateOutput = true });
            using var output = new MemoryStream();

            await Assert.ThrowsExceptionAsync<MergeException>(() => merger.MergeAsync(new MemoryStream(InvalidTemplate()), output, new MergeData().Text("NAME", "Jane")));

            Assert.AreEqual(0, output.Length);
        }
    }
}
