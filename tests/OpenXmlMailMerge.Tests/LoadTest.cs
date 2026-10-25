using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Data;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using System.Diagnostics;
using System.Globalization;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace OpenXmlMailMerge.Tests
{
    [TestClass]
    public class LoadTest
    {
        private static readonly string BindFile = Path.Combine("Fixtures", "BIND.docx");
        private static readonly string ImageFile = Path.Combine("Fixtures", "image.png");

        /// <summary>
        /// Every merged document is kept here (bin/&lt;configuration&gt;/&lt;tfm&gt;/Output) to be opened in Word.
        /// The folder is emptied at the start of each run.
        /// </summary>
        private static readonly string OutputDirectory = Path.Combine(AppContext.BaseDirectory, "Output");

        // ReSharper disable once UnusedAutoPropertyAccessor.Global
        public TestContext TestContext { get; set; }

        [ClassInitialize]
        public static void ClassInitialize(TestContext context)
        {
            if (Directory.Exists(OutputDirectory))
                Directory.Delete(OutputDirectory, true);

            Directory.CreateDirectory(OutputDirectory);
        }

        private string OutputPath(string suffix = null)
        {
            var name = suffix == null ? TestContext.TestName : $"{TestContext.TestName}_{suffix}";
            foreach (var c in Path.GetInvalidFileNameChars().Concat(['$', '*', '?', '"', '\\', '/', ':', '<', '>', '|', ' ']))
            {
                name = name?.Replace(c, '_');
            }

            var path = Path.Combine(OutputDirectory, name + ".docx");
            TestContext.WriteLine($"Output: {path}");
            return path;
        }

        /// <summary>
        /// Saves a merged document to the output folder, then asserts it is schema-valid.
        /// It is saved first so an invalid document can still be opened and inspected.
        /// </summary>
        private void SaveAndValidate(byte[] data, string suffix = null)
        {
            File.WriteAllBytes(OutputPath(suffix), data);
            AssertValid(data);
        }

        private static void AssertValid(byte[] data)
        {
            using var document = WordprocessingDocument.Open(new MemoryStream(data), false);
            var errors = new OpenXmlValidator().Validate(document)
            .Select(e => $"{e.Part?.Uri} {e.Path?.XPath}: {e.Description}")
            .ToList();

            Assert.AreEqual(0, errors.Count, "Invalid document:" + Environment.NewLine + string.Join(Environment.NewLine, errors));
        }

        private static OxMailMerge GetOxObject()
        {
            var imageData = File.ReadAllBytes(ImageFile);

            var dt = new DataTable();
            dt.Columns.Add(new DataColumn("COL1", typeof(string)));
            dt.Columns.Add(new DataColumn("COL2", typeof(int)));
            dt.Columns.Add(new DataColumn("COL3", typeof(DateTime)));
            dt.Columns.Add(new DataColumn("COL4", typeof(decimal)));

            for (var i = 0; i < 1; i++)
            {
                var row = dt.NewRow();
                row["COL1"] = $"ROW {i.ToString()}";
                row["COL2"] = i;
                row["COL3"] = DateTime.Now;
                row["COL4"] = Convert.ToDecimal(i);
                dt.Rows.Add(row);
            }

            var openXmlMailMerge = new OxMailMerge(DocumentType.Doc);
            openXmlMailMerge.Document.AddToDictionary(MailMergeDataTypeEnum.Text, "HEADER", "Vagner");
            openXmlMailMerge.Document.AddToDictionary(MailMergeDataTypeEnum.Text, "FOOTER", "Lucas");
            openXmlMailMerge.Document.AddToDictionary(MailMergeDataTypeEnum.Text, "BIND", "Reis");
            openXmlMailMerge.Document.AddToDictionary(MailMergeDataTypeEnum.Image, "IMG", imageData);
            openXmlMailMerge.Document.AddToDictionary(MailMergeDataTypeEnum.Table, "TABLE", dt);

            return openXmlMailMerge;
        }

        private static string ReadPart(byte[] package, string partName)
        {
            using var zip = new ZipArchive(new MemoryStream(package), ZipArchiveMode.Read);
            using var reader = new StreamReader(zip.GetEntry(partName)!.Open());
            return reader.ReadToEnd();
        }

        private static void AssertMerged(byte[] data)
        {
            var body = ReadPart(data, "word/document.xml");
            StringAssert.Contains(body, "Reis");
            StringAssert.Contains(body, "COL1");
            Assert.IsFalse(body.Contains("«BIND»"), "Body field was not merged.");
            StringAssert.Contains(ReadPart(data, "word/header1.xml"), "Vagner");
            StringAssert.Contains(ReadPart(data, "word/footer1.xml"), "Lucas");
        }

        private string MergeBindText(string key, string value, string outputSuffix = null)
        {
            var openXmlMailMerge = new OxMailMerge(DocumentType.Doc);
            openXmlMailMerge.Document.AddToDictionary(MailMergeDataTypeEnum.Text, key, value);

            openXmlMailMerge.Document.LoadFromBytes(File.ReadAllBytes(BindFile));
            openXmlMailMerge.Process();
            var data = openXmlMailMerge.Document.GetBytes();
            openXmlMailMerge.Terminate();

            SaveAndValidate(data, outputSuffix);
            return ReadPart(data, "word/document.xml");
        }

        [TestMethod]
        [DataRow("Price: $0.99")]
        [DataRow("$1 ${name} $$ \\1")]
        [DataRow("(a+)+ [x] .* ^$")]
        public void Text_ValueWithRegexSyntax_IsInsertedAsIs(string value)
        {
            var body = MergeBindText("BIND", value, value);

            StringAssert.Contains(body, System.Security.SecurityElement.Escape(value));
        }

        [TestMethod]
        public void Text_KeyCaseDiffersFromField_ReplacesField()
        {
            var body = MergeBindText("bind", "Reis");

            StringAssert.Contains(body, "<w:t>Reis</w:t>");
        }

        [TestMethod]
        public void Text_KeyIsPrefixOfAnotherField_MatchesOnlyExactField()
        {
            var body = MergeBody(CreateTemplate("NAME", "NAME_FULL"), MailMergeDataTypeEnum.Text, "NAME", "Jane");

            Assert.AreEqual(1, Regex.Matches(body, "Jane").Count);
            StringAssert.Contains(body, "«NAME_FULL»");
        }

        [TestMethod]
        public void Text_KeyIsPrefixOfField_DoesNotMatch()
        {
            var body = MergeBindText("BIN", "Reis");

            Assert.IsFalse(body.Contains("Reis"), "Key BIN must not match the field BIND.");
            StringAssert.Contains(body, "«BIND»");
        }

        [TestMethod]
        public void Text_LiteralTextContainingKey_IsNotChanged()
        {
            var header = new W.Paragraph(new W.Run(new W.Text("HEADER notes: ") { Space = SpaceProcessingModeValues.Preserve }), Field("HEADER"));
            var template = CreateDocument([new W.Paragraph(new W.Run(new W.Text("Body")))], [header]);

            var data = Merge(template, MailMergeDataTypeEnum.Text, "HEADER", "Vagner");

            Assert.AreEqual("HEADER notes: Vagner", FirstHeaderParagraphText(data));
        }

        [TestMethod]
        public void Text_IsInsertedAtFieldPosition_KeepingFieldFormatting()
        {
            var template = CreateDocument([ParagraphWithBoldField("NAME")]);

            var data = Merge(template, MailMergeDataTypeEnum.Text, "NAME", "Jane");

            using var document = WordprocessingDocument.Open(new MemoryStream(data), false);
            var paragraph = document.MainDocumentPart?.Document.Body?.GetFirstChild<W.Paragraph>();
            Debug.Assert(paragraph != null, nameof(paragraph) + " != null");
            CollectionAssert.AreEqual((string[])["r", "r", "r"], ChildNames(paragraph));
            var runs = paragraph.Elements<W.Run>().ToList();
            Assert.AreEqual("Before Jane After", paragraph.InnerText);
            Assert.IsNotNull(runs[1].RunProperties?.Bold, "The text run lost the field's formatting.");
        }

        [TestMethod]
        public void Text_BoundFieldsAreRemoved_UnboundFieldsStay()
        {
            var body = MergeBody(CreateTemplate("NAME", "OTHER"), MailMergeDataTypeEnum.Text, "NAME", "Jane");

            Assert.IsFalse(Regex.IsMatch(body, "w:instr=\"[^\"]*NAME"), "The bound field is still in the document.");
            Assert.IsTrue(Regex.IsMatch(body, "w:instr=\"[^\"]*OTHER"), "The unbound field was removed.");
        }

        [TestMethod]
        public void Text_NullValue_ClearsField()
        {
            var body = MergeBindText("BIND", null);

            Assert.IsFalse(body.Contains("«BIND»"));
            Assert.IsFalse(Regex.IsMatch(body, "w:instr=\"[^\"]*BIND"));
        }

        [TestMethod]
        public void Text_ValueWithSurroundingSpaces_IsPreserved()
        {
            var body = MergeBindText("BIND", "  Reis  ");

            StringAssert.Contains(body, "<w:t xml:space=\"preserve\">  Reis  </w:t>");
        }

        [TestMethod]
        public void Text_KeyWithRegexSyntax_IsMatchedLiterally()
        {
            var body = MergeBindText("B.ND", "Reis");

            Assert.IsFalse(body.Contains("Reis"), "A '.' in the key must not match any character.");
        }

        /// <summary>
        /// A w:fldSimple with the instruction "MERGEFIELD name" and the display text «name».
        /// </summary>
        private static W.SimpleField Field(string name)
        {
            // Word quotes a field name that contains spaces.
            var quoted = name.Contains(' ') ? $"\"{name}\"" : name;
            return new W.SimpleField(new W.Run(new W.Text($"«{name}»"))) { Instruction = $" MERGEFIELD {quoted} \\* MERGEFORMAT " };
        }

        /// <summary>
        /// A field whose result run has the given formatting.
        /// </summary>
        private static W.SimpleField Field(string name, W.RunProperties runProperties)
        {
            var field = Field(name);
            field.GetFirstChild<W.Run>()?.PrependChild(runProperties);
            return field;
        }

        /// <summary>
        /// A4 page with Word's default margins, the same as BIND.docx.
        /// </summary>
        private static W.SectionProperties PageSetup(params OpenXmlElement[] headerReferences)
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
        private static byte[] CreateDocument(OpenXmlElement[] blocks, params W.Paragraph[][] headers)
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
        private static byte[] CreateDocument(OpenXmlElement[] blocks, W.Paragraph[][] headers, W.Paragraph[][] footers, W.Styles styles = null)
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
        private static byte[] CreateTemplate(params string[] fieldNames)
        {
            return CreateDocument([.. fieldNames.Select(name => (OpenXmlElement)new W.Paragraph(Field(name)))]);
        }

        private byte[] Merge(byte[] template, MailMergeDataTypeEnum dataType, string key, object value, string outputSuffix = null)
        {
            var openXmlMailMerge = new OxMailMerge(DocumentType.Doc);
            openXmlMailMerge.Document.AddToDictionary(dataType, key, value);

            openXmlMailMerge.Document.LoadFromBytes(template);
            openXmlMailMerge.Process();
            var data = openXmlMailMerge.Document.GetBytes();
            openXmlMailMerge.Terminate();

            SaveAndValidate(data, outputSuffix);
            return data;
        }

        private string MergeBody(byte[] template, MailMergeDataTypeEnum dataType, string key, object value, string outputSuffix = null)
        {
            return ReadPart(Merge(template, dataType, key, value, outputSuffix), "word/document.xml");
        }

        private static string[] ChildNames(OpenXmlElement element)
        {
            return [.. element.ChildElements.Select(e => e.LocalName)];
        }

        private static int Count(string xml, string element)
        {
            return Regex.Matches(xml, $"<{element}[ >]").Count;
        }

        private static DataTable OneColumnTable()
        {
            var dt = new DataTable();
            dt.Columns.Add("COL1");
            dt.Rows.Add("ROW 0");
            return dt;
        }

        [TestMethod]
        public void Image_KeyIsPrefixOfAnotherField_MatchesOnlyExactField()
        {
            var body = MergeBody(CreateTemplate("IMG", "IMG_LOGO"), MailMergeDataTypeEnum.Image, "IMG", File.ReadAllBytes(ImageFile));

            Assert.AreEqual(1, Count(body, "w:drawing"));
            StringAssert.Contains(body, "«IMG_LOGO»");
        }

        [TestMethod]
        public void Image_KeyIsPrefixOfField_DoesNotMatch()
        {
            var body = MergeBody(File.ReadAllBytes(BindFile), MailMergeDataTypeEnum.Image, "IM", File.ReadAllBytes(ImageFile));

            Assert.AreEqual(0, Count(body, "w:drawing"));
        }

        [TestMethod]
        public void Image_FieldWithParameters_MatchesNameBeforeParameters()
        {
            var body = MergeBody(CreateTemplate("IMG@WIDTH=120@HEIGHT=80"), MailMergeDataTypeEnum.Image, "img", File.ReadAllBytes(ImageFile));

            Assert.AreEqual(1, Count(body, "w:drawing"));
            StringAssert.Contains(body, $"cx=\"{120 * 9525}\"");
        }

        private const long EmuPerPixel = 9525;

        /// <summary>
        /// Gets the size of the first inline image, in pixels.
        /// </summary>
        private static (long Width, long Height) ImageSize(string body)
        {
            var match = Regex.Match(body, "<wp:extent cx=\"(\\d+)\" cy=\"(\\d+)\"");
            Assert.IsTrue(match.Success, "No inline image found.");
            return (long.Parse(match.Groups[1].Value) / EmuPerPixel, long.Parse(match.Groups[2].Value) / EmuPerPixel);
        }

        private static byte[] BigEndian(int value, int bytes)
        {
            return [.. Enumerable.Range(0, bytes).Select(i => (byte)(value >> 8 * (bytes - 1 - i)))];
        }

        private static byte[] LittleEndian(int value, int bytes)
        {
            return [.. Enumerable.Range(0, bytes).Select(i => (byte)(value >> 8 * i))];
        }

        /// <summary>
        /// Builds just enough of an image file for its size to be read from the header.
        /// </summary>
        private static byte[] ImageHeader(string format, int width, int height)
        {
            return format switch
            {
                "png" =>
                [
                    0x89, (byte)'P', (byte)'N', (byte)'G', 0x0D, 0x0A, 0x1A, 0x0A, 0, 0, 0, 13, (byte)'I', (byte)'H', (byte)'D', (byte)'R',
                    .. BigEndian(width, 4), .. BigEndian(height, 4), .. new byte[9]
                ],
                "gif" => [.. "GIF89a"u8, .. LittleEndian(width, 2), .. LittleEndian(height, 2), .. new byte[3]],
                "bmp" => [.. "BM"u8, .. new byte[16], .. LittleEndian(width, 4), .. LittleEndian(-height, 4)],
                // SOI, an APP0 segment to skip, then SOF0 with the size.
                "jpeg" =>
                [
                    0xFF, 0xD8, 0xFF, 0xE0, 0x00, 0x10, .. new byte[14],
                    0xFF, 0xC0, 0x00, 0x11, 0x08, .. BigEndian(height, 2), .. BigEndian(width, 2), .. new byte[10]
                ],
                // A big-endian TIFF whose directory stores width and height as LONG (type 4) values.
                "tiff-long" =>
                [
                    .. "MM"u8, 0x00, 0x2A, 0, 0, 0, 8, 0, 2,
                    0x01, 0x00, 0, 4, 0, 0, 0, 1, .. BigEndian(width, 4),
                    0x01, 0x01, 0, 4, 0, 0, 0, 1, .. BigEndian(height, 4),
                    0, 0, 0, 0
                ],
                // Signatures only: these have no readable size.
                "tiff" => [.. "II"u8, 0x2A, 0x00, .. new byte[4]],
                "emf" => [0x01, 0x00, 0x00, 0x00, .. new byte[36], .. " EMF"u8, .. new byte[4]],
                "wmf" => [0xD7, 0xCD, 0xC6, 0x9A, .. new byte[18]],
                _ => [1, 2, 3, 4]
            };
        }

        [TestMethod]
        public void Image_WithoutSize_UsesImagePixelSize()
        {
            var body = MergeBody(File.ReadAllBytes(BindFile), MailMergeDataTypeEnum.Image, "IMG", File.ReadAllBytes(ImageFile));

            Assert.AreEqual((368L, 368L), ImageSize(body));
        }

        [TestMethod]
        [DataRow("png")]
        [DataRow("gif")]
        [DataRow("bmp")]
        [DataRow("jpeg")]
        [DataRow("tiff-long")]
        public void Image_WithoutSize_ReadsPixelSizeOfFormat(string format)
        {
            var body = MergeBody(CreateTemplate("IMG"), MailMergeDataTypeEnum.Image, "IMG", ImageHeader(format, 30, 20), format);

            Assert.AreEqual((30L, 20L), ImageSize(body));
        }

        [TestMethod]
        public void Image_UnreadableFormat_UsesDefaultSize()
        {
            var body = MergeBody(CreateTemplate("IMG"), MailMergeDataTypeEnum.Image, "IMG", ImageHeader("unknown", 0, 0));

            Assert.AreEqual((42L, 42L), ImageSize(body));
        }

        [TestMethod]
        [DataRow("IMG@WIDTH=120", 120, 60)]
        [DataRow("IMG@HEIGHT=50", 100, 50)]
        public void Image_OneDimension_KeepsAspectRatio(string field, int expectedWidth, int expectedHeight)
        {
            var body = MergeBody(CreateTemplate(field), MailMergeDataTypeEnum.Image, "IMG", ImageHeader("png", 200, 100), field);

            Assert.AreEqual(((long)expectedWidth, (long)expectedHeight), ImageSize(body));
        }

        [TestMethod]
        [DataRow("IMG@width=120")]
        [DataRow("IMG@Width=120")]
        public void Image_ParameterKeys_IgnoreCase(string field)
        {
            var body = MergeBody(CreateTemplate(field), MailMergeDataTypeEnum.Image, "IMG", ImageHeader("png", 200, 100), outputSuffix: field);

            Assert.AreEqual((120L, 60L), ImageSize(body));
        }

        [TestMethod]
        public void Image_ParameterKeys_AreReadUnderTurkishCulture()
        {
            // In tr-TR, "width".ToUpper() is "WİDTH" (dotted capital I), which never equals "WIDTH".
            var culture = CultureInfo.CurrentCulture;
            CultureInfo.CurrentCulture = new CultureInfo("tr-TR");
            try
            {
                var body = MergeBody(CreateTemplate("IMG@width=120@height=30"), MailMergeDataTypeEnum.Image, "IMG", ImageHeader("png", 200, 100));

                Assert.AreEqual((120L, 30L), ImageSize(body));
            }
            finally
            {
                CultureInfo.CurrentCulture = culture;
            }
        }

        [TestMethod]
        [DataRow("IMG@WIDTH=abc@HEIGHT=50", 100, 50)]
        [DataRow("IMG@WIDTH=-5", 200, 100)]
        [DataRow("IMG@=5@HEIGHT@WIDTH=", 200, 100)]
        public void Image_InvalidParameters_AreIgnored(string field, int expectedWidth, int expectedHeight)
        {
            var body = MergeBody(CreateTemplate(field), MailMergeDataTypeEnum.Image, "IMG", ImageHeader("png", 200, 100), outputSuffix: field);

            Assert.AreEqual(((long)expectedWidth, (long)expectedHeight), ImageSize(body));
        }

        [TestMethod]
        public void Image_WiderThanPage_IsScaledToTextWidth()
        {
            // BIND.docx: page 11906 twips, margins 1701 + 1701, so 8504 twips / 15 = 566 px between the margins.
            var body = MergeBody(File.ReadAllBytes(BindFile), MailMergeDataTypeEnum.Image, "IMG", ImageHeader("png", 2000, 1000));

            Assert.AreEqual((566L, 283L), ImageSize(body));
        }

        [TestMethod]
        public void Image_ExplicitSize_IsNotScaledToTextWidth()
        {
            var template = CreateTemplate("IMG@WIDTH=2000@HEIGHT=10");

            var body = MergeBody(template, MailMergeDataTypeEnum.Image, "IMG", ImageHeader("png", 200, 100));

            Assert.AreEqual((2000L, 10L), ImageSize(body));
        }

        [TestMethod]
        public void Image_InHeaderAndFooter_UsesImagePixelSize()
        {
            var data = Merge(File.ReadAllBytes(BindFile), MailMergeDataTypeEnum.Image, "IMG", File.ReadAllBytes(ImageFile));

            Assert.AreEqual((368L, 368L), ImageSize(ReadPart(data, "word/header1.xml")));
            Assert.AreEqual((368L, 368L), ImageSize(ReadPart(data, "word/footer1.xml")));
        }

        private static string FirstHeaderParagraphText(byte[] data)
        {
            using var document = WordprocessingDocument.Open(new MemoryStream(data), false);
            return document.MainDocumentPart?.HeaderParts.Single().Header.GetFirstChild<W.Paragraph>()?.InnerText;
        }

        private static string FirstHeaderXml(byte[] data)
        {
            using var document = WordprocessingDocument.Open(new MemoryStream(data), false);
            return document.MainDocumentPart?.HeaderParts.Single().Header.OuterXml;
        }

        private static byte[] HeaderTemplate(string fieldName)
        {
            return CreateDocument(
                [new W.Paragraph(new W.Run(new W.Text("Body")))],
                [new W.Paragraph(Field(fieldName))]);
        }

        [TestMethod]
        public void Image_InHeader_OneDimension_KeepsAspectRatio()
        {
            var data = Merge(HeaderTemplate("IMG@WIDTH=120"), MailMergeDataTypeEnum.Image, "IMG", ImageHeader("png", 200, 100));

            Assert.AreEqual((120L, 60L), ImageSize(FirstHeaderXml(data)));
        }

        [TestMethod]
        public void Image_InHeader_WiderThanPage_IsScaledToTextWidth()
        {
            var data = Merge(HeaderTemplate("IMG"), MailMergeDataTypeEnum.Image, "IMG", ImageHeader("png", 2000, 1000));

            Assert.AreEqual((566L, 283L), ImageSize(FirstHeaderXml(data)));
        }

        /// <summary>
        /// A paragraph with text around a bold field: "Before ", the field, " After".
        /// </summary>
        private static W.Paragraph ParagraphWithBoldField(string name)
        {
            var field = Field(name);
            field.GetFirstChild<W.Run>()?.PrependChild(new W.RunProperties(new W.Bold()));

            return new W.Paragraph(
                new W.Run(new W.Text("Before ") { Space = SpaceProcessingModeValues.Preserve }),
                field,
                new W.Run(new W.Text(" After") { Space = SpaceProcessingModeValues.Preserve }));
        }

        /// <summary>
        /// Asserts that the image replaced the field where it was, between the two text runs, keeping the field's formatting.
        /// </summary>
        private static void AssertImageAtFieldPosition(W.Paragraph paragraph)
        {
            CollectionAssert.AreEqual((string[])["r", "r", "r"], ChildNames(paragraph));

            var runs = paragraph.Elements<W.Run>().ToList();
            Assert.AreEqual("Before ", runs[0].InnerText);
            Assert.IsNotNull(runs[1].GetFirstChild<W.Drawing>(), "The image is not where the field was.");
            Assert.IsNotNull(runs[1].RunProperties?.Bold, "The image run lost the field's formatting.");
            Assert.AreEqual(" After", runs[2].InnerText);
        }

        [TestMethod]
        public void Image_InBodyHeaderAndFooter_IsInsertedAtFieldPosition()
        {
            var template = CreateDocument(
                [ParagraphWithBoldField("IMG")],
                [[ParagraphWithBoldField("IMG")]],
                [[ParagraphWithBoldField("IMG")]]);

            var data = Merge(template, MailMergeDataTypeEnum.Image, "IMG", File.ReadAllBytes(ImageFile));

            using var document = WordprocessingDocument.Open(new MemoryStream(data), false);
            var main = document.MainDocumentPart;
            Debug.Assert(main != null, nameof(main) + " != null");
            AssertImageAtFieldPosition(main.Document.Body?.GetFirstChild<W.Paragraph>());
            AssertImageAtFieldPosition(main.HeaderParts.Single().Header.GetFirstChild<W.Paragraph>());
            AssertImageAtFieldPosition(main.FooterParts.Single().Footer.GetFirstChild<W.Paragraph>());
        }

        [TestMethod]
        public void Image_InBindTemplate_HeaderAndFooterFieldsAreReplacedTheSameWay()
        {
            var data = Merge(File.ReadAllBytes(BindFile), MailMergeDataTypeEnum.Image, "IMG", File.ReadAllBytes(ImageFile));

            foreach (var part in new[] { "word/header1.xml", "word/footer1.xml" })
            {
                var xml = ReadPart(data, part);
                Assert.IsFalse(Regex.IsMatch(xml, "w:instr=\"[^\"]*IMG"), $"The image field is still in {part}.");
                Assert.AreEqual(1, Count(xml, "w:drawing"), part);
            }
        }

        [TestMethod]
        [DataRow("png", "image/png", ".png")]
        [DataRow("jpeg", "image/jpeg", ".jpeg")]
        [DataRow("gif", "image/gif", ".gif")]
        [DataRow("bmp", "image/bmp", ".bmp")]
        [DataRow("tiff", "image/tiff", ".tiff")]
        [DataRow("emf", "image/x-emf", ".emf")]
        [DataRow("wmf", "image/x-wmf", ".wmf")]
        [DataRow("unknown", "image/jpeg", ".jpeg")]
        public void Image_PartContentType_MatchesImageFormat(string format, string contentType, string extension)
        {
            var data = Merge(CreateTemplate("IMG"), MailMergeDataTypeEnum.Image, "IMG", ImageHeader(format, 30, 20), outputSuffix: format);

            using var document = WordprocessingDocument.Open(new MemoryStream(data), false);
            var imagePart = document.MainDocumentPart?.ImageParts.Single();
            Debug.Assert(imagePart != null, nameof(imagePart) + " != null");
            Assert.AreEqual(contentType, imagePart.ContentType);
            StringAssert.EndsWith(imagePart.Uri.OriginalString, extension);
        }

        private static W.Run FieldMarker(W.FieldCharValues type)
        {
            return new W.Run(new W.FieldChar { FieldCharType = type });
        }

        /// <summary>
        /// The runs of a complex field, as Word writes them: begin, instruction runs, separate, a bold result «name», end.
        /// </summary>
        /// <param name="name">Field name.</param>
        /// <param name="withResult">False for a field without a separate marker and result.</param>
        /// <param name="instructionParts">The instruction split over several runs; by default a single " MERGEFIELD name " run.</param>
        private static OpenXmlElement[] ComplexField(string name, bool withResult = true, params string[] instructionParts)
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

        /// <summary>
        /// A paragraph with text around a complex field: "Before ", the field, " After".
        /// </summary>
        private static W.Paragraph ParagraphWithComplexField(OpenXmlElement[] field)
        {
            return new W.Paragraph(
            [
                new W.Run(new W.Text("Before ") { Space = SpaceProcessingModeValues.Preserve }),
                .. field,
                new W.Run(new W.Text(" After") { Space = SpaceProcessingModeValues.Preserve })
            ]);
        }

        private static W.Body Body(WordprocessingDocument document)
        {
            var body = document.MainDocumentPart?.Document.Body;
            Debug.Assert(body != null, nameof(body) + " != null");
            return body;
        }

        [TestMethod]
        public void ComplexField_Text_IsReplacedAtFieldPosition_KeepingResultFormatting()
        {
            var template = CreateDocument([ParagraphWithComplexField(ComplexField("NAME"))]);

            var data = Merge(template, MailMergeDataTypeEnum.Text, "NAME", "Jane");

            using var document = WordprocessingDocument.Open(new MemoryStream(data), false);
            var paragraph = Body(document).GetFirstChild<W.Paragraph>();
            Debug.Assert(paragraph != null, nameof(paragraph) + " != null");
            CollectionAssert.AreEqual((string[])["r", "r", "r"], ChildNames(paragraph));
            Assert.AreEqual("Before Jane After", paragraph.InnerText);
            Assert.IsNotNull(paragraph.Elements<W.Run>().ElementAt(1).RunProperties?.Bold, "The text lost the field result's formatting.");
            Assert.AreEqual(0, Body(document).Descendants<W.FieldChar>().Count());
        }

        [TestMethod]
        public void ComplexField_InstructionSplitAcrossRuns_IsMatched()
        {
            var field = ComplexField("NAME", true, " MERGE", "FIELD  NA", "ME \\* MERGEFORMAT ");
            var template = CreateDocument([ParagraphWithComplexField(field)]);

            var data = Merge(template, MailMergeDataTypeEnum.Text, "NAME", "Jane");

            using var document = WordprocessingDocument.Open(new MemoryStream(data), false);
            Assert.AreEqual("Before Jane After", Body(document).GetFirstChild<W.Paragraph>()?.InnerText);
        }

        [TestMethod]
        public void ComplexField_WithoutResult_IsReplaced()
        {
            var template = CreateDocument([ParagraphWithComplexField(ComplexField("NAME", withResult: false))]);

            var data = Merge(template, MailMergeDataTypeEnum.Text, "NAME", "Jane");

            using var document = WordprocessingDocument.Open(new MemoryStream(data), false);
            Assert.AreEqual("Before Jane After", Body(document).GetFirstChild<W.Paragraph>()?.InnerText);
            Assert.AreEqual(0, Body(document).Descendants<W.FieldChar>().Count());
        }

        [TestMethod]
        public void ComplexField_Image_IsInsertedAtFieldPosition()
        {
            var template = CreateDocument([ParagraphWithComplexField(ComplexField("IMG"))]);

            var data = Merge(template, MailMergeDataTypeEnum.Image, "IMG", File.ReadAllBytes(ImageFile));

            using var document = WordprocessingDocument.Open(new MemoryStream(data), false);
            var runs = Body(document).GetFirstChild<W.Paragraph>()?.Elements<W.Run>().ToList();
            Debug.Assert(runs != null, nameof(runs) + " != null");
            Assert.AreEqual(3, runs.Count);
            Assert.IsNotNull(runs[1].GetFirstChild<W.Drawing>());
        }

        [TestMethod]
        public void ComplexField_Table_AloneInParagraph_ReplacesParagraph()
        {
            var template = CreateDocument([new W.Paragraph(ComplexField("TABLE"))]);

            var data = Merge(template, MailMergeDataTypeEnum.Table, "TABLE", OneColumnTable());

            using var document = WordprocessingDocument.Open(new MemoryStream(data), false);
            CollectionAssert.AreEqual((string[])["tbl", "sectPr"], ChildNames(Body(document)));
        }

        [TestMethod]
        public void ComplexField_InHeader_IsReplaced()
        {
            var template = CreateDocument([new W.Paragraph(new W.Run(new W.Text("Body")))], [ParagraphWithComplexField(ComplexField("HEADER"))]);

            var data = Merge(template, MailMergeDataTypeEnum.Text, "HEADER", "Vagner");

            Assert.AreEqual("Before Vagner After", FirstHeaderParagraphText(data));
        }

        [TestMethod]
        public void ComplexField_Unbound_IsLeftAsIs()
        {
            var template = CreateDocument([ParagraphWithComplexField(ComplexField("OTHER"))]);

            var data = Merge(template, MailMergeDataTypeEnum.Text, "NAME", "Jane");

            using var document = WordprocessingDocument.Open(new MemoryStream(data), false);
            Assert.AreEqual(3, Body(document).Descendants<W.FieldChar>().Count());
            Assert.AreEqual(0, Body(document).Descendants<W.SimpleField>().Count());
        }

        [TestMethod]
        public void ComplexField_NestedInAnotherField_IsLeftAsIs()
        {
            // { IF { MERGEFIELD NAME } = "x" "yes" "no" }: converting the inner field would change what the IF evaluates.
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
            var template = CreateDocument([new W.Paragraph(field)]);

            var data = Merge(template, MailMergeDataTypeEnum.Text, "NAME", "Jane");

            using var document = WordprocessingDocument.Open(new MemoryStream(data), false);
            Assert.AreEqual(6, Body(document).Descendants<W.FieldChar>().Count());
            Assert.IsFalse(Body(document).InnerText.Contains("Jane"));
        }

        [TestMethod]
        public void ComplexField_SpanningParagraphs_IsLeftAsIs()
        {
            var runs = ComplexField("NAME");
            var template = CreateDocument([new W.Paragraph(runs[..3]), new W.Paragraph(runs[3..])]);

            var data = Merge(template, MailMergeDataTypeEnum.Text, "NAME", "Jane");

            using var document = WordprocessingDocument.Open(new MemoryStream(data), false);
            Assert.AreEqual(3, Body(document).Descendants<W.FieldChar>().Count());
            Assert.IsFalse(Body(document).InnerText.Contains("Jane"));
        }

        private static readonly string ImagesDirectory = Path.Combine("Fixtures", "Images");

        [TestMethod]
        [DataRow("hippopotamus.jpeg", "image/jpeg", ".jpeg")]
        [DataRow("peacock.progressive.jpeg", "image/jpeg", ".jpeg")]
        [DataRow("hippopotamus.regular.png", "image/png", ".png")]
        [DataRow("hippopotamus.interlaced.png", "image/png", ".png")]
        [DataRow("hippopotamus.regular.gif", "image/gif", ".gif")]
        [DataRow("hippopotamus.interlaced.gif", "image/gif", ".gif")]
        [DataRow("hippopotamus.bmp", "image/bmp", ".bmp")]
        [DataRow("pal8.bmp", "image/bmp", ".bmp")]
        [DataRow("pal8topdown.bmp", "image/bmp", ".bmp")]
        [DataRow("pal8os2.bmp", "image/bmp", ".bmp")]
        [DataRow("hippopotamus.tiff", "image/tiff", ".tiff")]
        [DataRow("minisblack-1c-8b.tiff", "image/tiff", ".tiff")]
        [DataRow("63327.emf", "image/x-emf", ".emf")]
        [DataRow("61294.emf", "image/x-emf", ".emf")]
        [DataRow("64716_image1.wmf", "image/x-wmf", ".wmf")]
        [DataRow("61338.wmf", "image/x-wmf", ".wmf")]
        public void Image_RealFile_GetsItsContentType(string file, string contentType, string extension)
        {
            var data = Merge(CreateTemplate("IMG"), MailMergeDataTypeEnum.Image, "IMG", File.ReadAllBytes(Path.Combine(ImagesDirectory, file)), outputSuffix: file);

            using var document = WordprocessingDocument.Open(new MemoryStream(data), false);
            var imagePart = document.MainDocumentPart?.ImageParts.Single();
            Debug.Assert(imagePart != null, nameof(imagePart) + " != null");
            Assert.AreEqual(contentType, imagePart.ContentType);
            StringAssert.EndsWith(imagePart.Uri.OriginalString, extension);
        }

        /// <summary>
        /// Real files of the formats whose size is read. Raster sizes were measured with macOS sips; vector sizes are the
        /// header's physical size at 96 DPI (EMF frame 5917 x 5291 hundredths of a millimeter; WMF bounding box
        /// 1622 x 1093 at 2536 units per inch).
        /// </summary>
        [TestMethod]
        [DataRow("hippopotamus.jpeg", 36, 28)]
        [DataRow("peacock.progressive.jpeg", 100, 75)]
        [DataRow("hippopotamus.regular.png", 36, 28)]
        [DataRow("hippopotamus.interlaced.png", 36, 28)]
        [DataRow("hippopotamus.regular.gif", 36, 28)]
        [DataRow("hippopotamus.interlaced.gif", 36, 28)]
        [DataRow("hippopotamus.bmp", 36, 28)]
        [DataRow("pal8.bmp", 127, 64)]
        [DataRow("pal8topdown.bmp", 127, 64)]
        [DataRow("pal8os2.bmp", 127, 64)]
        [DataRow("hippopotamus.tiff", 36, 28)]
        [DataRow("minisblack-1c-8b.tiff", 157, 151)]
        [DataRow("61294.emf", 224, 200)]
        [DataRow("64716_image1.wmf", 61, 41)]
        public void Image_RealFile_UsesItsPixelSize(string file, int width, int height)
        {
            var body = MergeBody(CreateTemplate("IMG"), MailMergeDataTypeEnum.Image, "IMG", File.ReadAllBytes(Path.Combine(ImagesDirectory, file)), outputSuffix: file);

            Assert.AreEqual(((long)width, (long)height), ImageSize(body));
        }

        /// <summary>
        /// Styles with a table style (id "MyGrid", name "My Grid") and, optionally, the document's default language.
        /// </summary>
        private static W.Styles TemplateStyles(string language = null)
        {
            var styles = new W.Styles();
            if (language != null)
                styles.Append(new W.DocDefaults(new W.RunPropertiesDefault(new W.RunPropertiesBaseStyle(new W.Languages { Val = language }))));

            styles.Append(new W.Style(new W.StyleName { Val = "My Grid" }) { Type = W.StyleValues.Table, StyleId = "MyGrid" });
            return styles;
        }

        private static byte[] TableTemplate(W.SimpleField field, string language = null)
        {
            return CreateDocument([new W.Paragraph(field)], [], [], TemplateStyles(language));
        }

        private static DataTable ValuesTable()
        {
            var dt = new DataTable();
            dt.Columns.Add("Date", typeof(DateTime));
            dt.Columns.Add("Amount", typeof(decimal));
            dt.Rows.Add(new DateTime(2026, 10, 25, 13, 5, 0), 1234.5m);
            return dt;
        }

        private static W.Table FirstTable(WordprocessingDocument document)
        {
            var table = Body(document).Descendants<W.Table>().FirstOrDefault();
            Debug.Assert(table != null, nameof(table) + " != null");
            return table;
        }

        /// <summary>
        /// Runs an action with another current culture, to prove the result doesn't depend on the machine's culture.
        /// </summary>
        private static T WithCulture<T>(string name, Func<T> action)
        {
            var culture = CultureInfo.CurrentCulture;
            CultureInfo.CurrentCulture = new CultureInfo(name);
            try
            {
                return action();
            }
            finally
            {
                CultureInfo.CurrentCulture = culture;
            }
        }

        [TestMethod]
        public void Table_CellText_UsesFieldFormatting()
        {
            var formatting = new W.RunProperties(new W.RunFonts { Ascii = "Arial" }, new W.Bold(), new W.FontSize { Val = "18" });

            var data = Merge(TableTemplate(Field("TABLE", formatting)), MailMergeDataTypeEnum.Table, "TABLE", OneColumnTable());

            using var document = WordprocessingDocument.Open(new MemoryStream(data), false);
            var runs = FirstTable(document).Descendants<W.Run>().ToList();
            Assert.AreEqual(2, runs.Count);
            foreach (var run in runs)
            {
                Assert.AreEqual("Arial", run.RunProperties?.RunFonts?.Ascii?.Value);
                Assert.IsNotNull(run.RunProperties?.Bold);
                Assert.AreEqual("18", run.RunProperties?.FontSize?.Val?.Value);
            }
        }

        [TestMethod]
        public void Table_WithoutFieldFormatting_UsesDocumentDefaults()
        {
            var data = Merge(TableTemplate(Field("TABLE")), MailMergeDataTypeEnum.Table, "TABLE", OneColumnTable());

            using var document = WordprocessingDocument.Open(new MemoryStream(data), false);
            Assert.IsTrue(FirstTable(document).Descendants<W.Run>().All(r => r.RunProperties == null), "Cells must not get a fixed font or size.");
            Assert.IsFalse(Body(document).OuterXml.Contains("Calibri"));
        }

        [TestMethod]
        [DataRow("TABLE@STYLE=MyGrid")]
        [DataRow("TABLE@style=mygrid")]
        [DataRow("TABLE@STYLE=My Grid")]
        public void Table_StyleParameter_AppliesTemplateStyle(string field)
        {
            var data = Merge(TableTemplate(Field(field)), MailMergeDataTypeEnum.Table, "TABLE", OneColumnTable(), outputSuffix: field);

            using var document = WordprocessingDocument.Open(new MemoryStream(data), false);
            var properties = FirstTable(document).GetFirstChild<W.TableProperties>();
            Assert.AreEqual("MyGrid", properties?.TableStyle?.Val?.Value);
            Assert.IsNull(properties?.TableBorders, "A styled table must take its borders from the style.");
        }

        [TestMethod]
        [DataRow("TABLE")]
        [DataRow("TABLE@STYLE=Missing")]
        public void Table_WithoutKnownStyle_GetsSingleBorders(string field)
        {
            var data = Merge(TableTemplate(Field(field)), MailMergeDataTypeEnum.Table, "TABLE", OneColumnTable(), outputSuffix: field);

            using var document = WordprocessingDocument.Open(new MemoryStream(data), false);
            var properties = FirstTable(document).GetFirstChild<W.TableProperties>();
            Assert.IsNull(properties?.TableStyle);
            Assert.AreEqual(W.BorderValues.Single, properties?.TableBorders?.TopBorder?.Val?.Value);
        }

        [TestMethod]
        public void Table_HeaderRow_RepeatsOnEachPage_AndColumnsShareTextWidth()
        {
            var data = Merge(TableTemplate(Field("TABLE")), MailMergeDataTypeEnum.Table, "TABLE", ValuesTable());

            using var document = WordprocessingDocument.Open(new MemoryStream(data), false);
            var rows = FirstTable(document).Elements<W.TableRow>().ToList();
            Assert.IsNotNull(rows[0].TableRowProperties?.GetFirstChild<W.TableHeader>(), "The caption row must be a header row.");
            Assert.IsNull(rows[1].TableRowProperties);

            // 566 px between the margins, 15 twips per pixel, 2 columns.
            var widths = FirstTable(document).GetFirstChild<W.TableGrid>()?.Elements<W.GridColumn>().Select(c => c.Width?.Value).ToList();
            CollectionAssert.AreEqual(new[] { "4245", "4245" }, widths);
        }

        [TestMethod]
        [DataRow("pt-BR", null, "25/10/2026 13:05:00", "1234,5")]
        [DataRow("pt-BR", "de-DE", "25.10.2026 13:05:00", "1234,5")]
        [DataRow(null, null, "10/25/2026 13:05:00", "1234.5")]
        public void Table_Values_UseTemplateLanguage(string documentLanguage, string fieldLanguage, string date, string amount)
        {
            var field = fieldLanguage == null ? Field("TABLE") : Field("TABLE", new W.RunProperties(new W.Languages { Val = fieldLanguage }));

            // The merge runs under en-US, which must not leak into the output.
            var data = WithCulture("en-US", () => Merge(TableTemplate(field, documentLanguage), MailMergeDataTypeEnum.Table, "TABLE", ValuesTable(),
                outputSuffix: $"{documentLanguage ?? "none"}-{fieldLanguage ?? "none"}"));

            using var document = WordprocessingDocument.Open(new MemoryStream(data), false);
            var cells = FirstTable(document).Elements<W.TableRow>().ElementAt(1).Elements<W.TableCell>().Select(c => c.InnerText).ToList();
            CollectionAssert.AreEqual(new[] { date, amount }, cells);
        }

        [TestMethod]
        public void Text_Value_UsesTemplateLanguage()
        {
            // BIND.docx's default language is pt-BR.
            var body = WithCulture("en-US", () => MergeBody(File.ReadAllBytes(BindFile), MailMergeDataTypeEnum.Text, "BIND", 1234.5m));

            StringAssert.Contains(body, "<w:t>1234,5</w:t>");
        }

        [TestMethod]
        public void Image_RealNonPlaceableWmf_UsesDefaultSize()
        {
            // A WMF without the placeable header has no physical size.
            var body = MergeBody(CreateTemplate("IMG"), MailMergeDataTypeEnum.Image, "IMG", File.ReadAllBytes(Path.Combine(ImagesDirectory, "61338.wmf")));

            Assert.AreEqual((42L, 42L), ImageSize(body));
        }

        [TestMethod]
        public void Table_FieldAloneInParagraph_ReplacesParagraph()
        {
            var data = Merge(CreateTemplate("TABLE"), MailMergeDataTypeEnum.Table, "TABLE", OneColumnTable());

            using var document = WordprocessingDocument.Open(new MemoryStream(data), false);
            CollectionAssert.AreEqual(new[] { "tbl", "sectPr" }, ChildNames(document.MainDocumentPart?.Document.Body));
        }

        [TestMethod]
        public void Table_InBindTemplate_IsPlacedAtFieldBeforeSectionProperties()
        {
            var body = MergeBody(File.ReadAllBytes(BindFile), MailMergeDataTypeEnum.Table, "TABLE", OneColumnTable());

            Assert.IsFalse(Regex.IsMatch(body, "w:instr=\"[^\"]*TABLE"), "The table field was not removed.");
            Assert.IsTrue(body.IndexOf("<w:tbl", StringComparison.Ordinal) < body.LastIndexOf("<w:sectPr", StringComparison.Ordinal));
        }

        [TestMethod]
        public void Table_FieldWithOtherContent_GoesAfterParagraph()
        {
            var template = CreateDocument([new W.Paragraph(new W.Run(new W.Text("Items:")), Field("TABLE"))]);

            var data = Merge(template, MailMergeDataTypeEnum.Table, "TABLE", OneColumnTable());

            using var document = WordprocessingDocument.Open(new MemoryStream(data), false);
            var body = document.MainDocumentPart?.Document.Body;
            CollectionAssert.AreEqual(new[] { "p", "tbl", "sectPr" }, ChildNames(body));
            Assert.AreEqual("Items:", body?.GetFirstChild<W.Paragraph>()?.InnerText);
        }

        [TestMethod]
        public void Table_InTableCell_CellStillEndsWithParagraph()
        {
            var outer = new W.Table(
                new W.TableProperties(),
                new W.TableGrid(new W.GridColumn()),
                new W.TableRow(new W.TableCell(new W.Paragraph(Field("TABLE")))));

            var data = Merge(CreateDocument([outer]), MailMergeDataTypeEnum.Table, "TABLE", OneColumnTable());

            using var document = WordprocessingDocument.Open(new MemoryStream(data), false);
            var cell = document.MainDocumentPart?.Document.Body?.GetFirstChild<W.Table>()?.Descendants<W.TableCell>().First();
            CollectionAssert.AreEqual(new[] { "tbl", "p" }, ChildNames(cell));
        }

        [TestMethod]
        public void Table_InParagraphWithSectionBreak_KeepsSectionBreak()
        {
            var template = CreateDocument([
                new W.Paragraph(new W.ParagraphProperties(PageSetup()), Field("TABLE")),
                new W.Paragraph(new W.Run(new W.Text("Next section")))
            ]);

            var data = Merge(template, MailMergeDataTypeEnum.Table, "TABLE", OneColumnTable());

            using var document = WordprocessingDocument.Open(new MemoryStream(data), false);
            var body = document.MainDocumentPart?.Document.Body;
            CollectionAssert.AreEqual(new[] { "tbl", "p", "p", "sectPr" }, ChildNames(body));
            Debug.Assert(body != null, nameof(body) + " != null");
            Assert.AreEqual(2, body.Descendants<W.SectionProperties>().Count());
        }

        [TestMethod]
        public void Table_InSecondHeader_GoesIntoThatHeader()
        {
            var template = CreateDocument(
                [new W.Paragraph(new W.Run(new W.Text("Body")))],
                [new W.Paragraph(new W.Run(new W.Text("Default header")))],
                [new W.Paragraph(Field("TABLE"))]);

            var data = Merge(template, MailMergeDataTypeEnum.Table, "TABLE", OneColumnTable());

            using var document = WordprocessingDocument.Open(new MemoryStream(data), false);
            var headers = document.MainDocumentPart?.HeaderParts.Select(p => p.Header).ToList();
            Debug.Assert(headers != null, nameof(headers) + " != null");
            var defaultHeader = headers.Single(h => h.InnerText.Contains("Default header"));
            var fieldHeader = headers.Single(h => h != defaultHeader);
            Assert.AreEqual(0, defaultHeader.Descendants<W.Table>().Count());
            CollectionAssert.AreEqual((string[])["tbl", "p"], ChildNames(fieldHeader));
        }

        [TestMethod]
        public void Table_KeyIsPrefixOfAnotherField_MatchesOnlyExactField()
        {
            var body = MergeBody(CreateTemplate("TABLE", "TABLE_TOTALS"), MailMergeDataTypeEnum.Table, "TABLE", OneColumnTable());

            Assert.AreEqual(1, Count(body, "w:tbl"));
            StringAssert.Contains(body, "«TABLE_TOTALS»");
        }

        [TestMethod]
        public void Table_KeyIsPrefixOfField_DoesNotMatch()
        {
            var body = MergeBody(File.ReadAllBytes(BindFile), MailMergeDataTypeEnum.Table, "TAB", OneColumnTable());

            Assert.AreEqual(0, Count(body, "w:tbl"));
        }

        [TestMethod]
        public void LoadFromBytes_GetBytes_ReturnsMergedDocument()
        {
            var openXmlMailMerge = GetOxObject();

            openXmlMailMerge.Document.LoadFromBytes(File.ReadAllBytes(BindFile));
            openXmlMailMerge.Process();
            var data = openXmlMailMerge.Document.GetBytes();
            openXmlMailMerge.Terminate();

            SaveAndValidate(data);
            AssertMerged(data);
        }

        [TestMethod]
        public void LoadFromBytes_DoesNotModifyCallerArray()
        {
            var openXmlMailMerge = GetOxObject();
            var template = File.ReadAllBytes(BindFile);
            var original = template.ToArray();

            openXmlMailMerge.Document.LoadFromBytes(template);
            openXmlMailMerge.Process();
            openXmlMailMerge.Document.GetBytes();
            openXmlMailMerge.Terminate();

            CollectionAssert.AreEqual(original, template);
        }

        [TestMethod]
        public async Task LoadFromFileAsync_SaveToFileAsync_WritesMergedDocument()
        {
            var openXmlMailMerge = GetOxObject();
            var outputPath = OutputPath();

            await openXmlMailMerge.Document.LoadFromFileAsync(BindFile);
            openXmlMailMerge.Process();
            await openXmlMailMerge.Document.SaveToFileAsync(outputPath);
            openXmlMailMerge.Terminate();

            var data = await File.ReadAllBytesAsync(outputPath);
            AssertValid(data);
            AssertMerged(data);
        }

        [TestMethod]
        public async Task SaveToStreamAsync_MatchesGetBytes_AndDocumentStaysOpen()
        {
            var openXmlMailMerge = GetOxObject();

            await using (var input = File.OpenRead(BindFile))
            {
                await openXmlMailMerge.Document.LoadFromStreamAsync(input);
            }
            openXmlMailMerge.Process();

            var first = openXmlMailMerge.Document.GetBytes();
            var output = new MemoryStream();
            await openXmlMailMerge.Document.SaveToStreamAsync(output);
            var second = openXmlMailMerge.Document.GetBytes();
            openXmlMailMerge.Terminate();

            SaveAndValidate(output.ToArray());
            CollectionAssert.AreEqual(first, output.ToArray());
            CollectionAssert.AreEqual(first, second);
            AssertMerged(output.ToArray());
        }

        [TestMethod]
        public void GetBytes_WithoutLoad_Throws()
        {
            var openXmlMailMerge = GetOxObject();

            Assert.ThrowsException<InvalidOperationException>(() => openXmlMailMerge.Document.GetBytes());
        }
    }
}
