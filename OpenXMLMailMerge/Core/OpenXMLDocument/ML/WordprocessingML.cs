using DocumentFormat.OpenXml.Packaging;
using OpenXMLMailMerge.Core.OpenXMLDocument.Document;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Linq;
using OpenXMLMailMerge.Core.OpenXMLDocument.Dictionary;
using System.Text.RegularExpressions;
using System.Data;
using System.IO;
using OpenXMLMailMerge.Core.OpenXMLDocument.Configuration;
using System.Collections.Generic;
using DocumentFormat.OpenXml;
using OpenXMLMailMerge.Core.OpenXMLDocument.Document.Element;

namespace OpenXMLMailMerge.Core.OpenXMLDocument.ML
{
    /// <summary>
    /// A OpenXMLDocument as word processor.
    /// </summary>
    public class WordprocessingML : OpenXMLDocument
    {
        /// <summary>
        /// Loads a document from file.
        /// </summary>
        /// <param name="path">Path of the document.</param>
        /// <param name="isEditable">Whether is a editable copy.</param>
        /// <param name="removeCopy">Whether should remove a copy or not.</param>
        /// <param name="autoSave">Whether it should auto save the document.</param>
        /// <returns>OpenXMLDocument processor.</returns>
        public override OpenXMLDocument LoadFromFile(string path, bool isEditable = true, bool removeCopy = false, bool autoSave = false)
        {
            base.LoadFromFile(path, isEditable, removeCopy);

            DocumentSettings.AutoSave = autoSave;

            ContentManager = new DocumentHandler()
            {
                Document = WordprocessingDocument.Open(DocumentSettings.TempPath, isEditable, new OpenSettings() { AutoSave = DocumentSettings.AutoSave })
            };

            ContentManager.Content = (ContentManager.Document as WordprocessingDocument).MainDocumentPart;

            return this;
        }

        /// <summary>
        /// Gets the value from the data dictionary entry.
        /// </summary>
        /// <param name="arg">Text of the field.</param>
        /// <returns>Array of byte with the value from data dictionary.</returns>
        public virtual byte[] TryReplaceImage(string arg)
        {
            foreach (var item in Dictionary)
            {
                if (item.Key == MailMergeDataTypeEnum.Image)
                {
                    foreach (var value in item.Value)
                    {
                        if (arg.ToUpper().Contains(value.Key.ToString().ToUpper()))
                        {
                            return value.Value as byte[];
                        }
                    }
                }
            }

            return null;
        }

        /// <summary>
        /// Gets the value from the data dictionary entry.
        /// </summary>
        /// <param name="arg">Text of the field.</param>
        /// <returns>String with the value from data dictionary.</returns>
        protected virtual string TryReplaceString(string arg)
        {
            foreach (var item in Dictionary)
            {
                if (item.Key == MailMergeDataTypeEnum.Regex)
                {
                    foreach (var value in item.Value)
                    {
                        if (!arg.ToUpper().Contains(value.Key.ToString().ToUpper())) continue;

                        var regexText = new Regex(value.Key);
                        string toReplace = Convert.ToString(value.Value);
                        var result = regexText.Replace(arg, toReplace);
                        result = result
                            .Replace(StringConst.SPECIAL_STR_L, string.Empty)
                            .Replace(StringConst.SPECIAL_STR_R, string.Empty);
                        return result;
                    }
                }
            }
            return arg;
        }

        /// <summary>
        /// Gets the value from the data dictionary entry.
        /// </summary>
        /// <param name="arg">Text of the field.</param>
        /// <returns>DataTable with the value from data dictionary.</returns>
        public virtual DataTable TryReplaceTable(string arg)
        {
            foreach (var item in Dictionary)
            {
                if (item.Key != MailMergeDataTypeEnum.Table) continue;

                foreach (var value in item.Value)
                {
                    if (arg.ToUpper().Contains(value.Key.ToString().ToUpper()))
                    {
                        return value.Value as DataTable;
                    }
                }
            }

            return null;
        }

        /// <summary>
        /// Executes string regex replacement.
        /// </summary>
        protected virtual void ProcessRegex()
        {
            var regexText = new Regex(StringConst.MERGEFIELD_REGEX);

            ContentManager.Content.Document.MainDocumentPart.Document.Descendants<Paragraph>().Where(w => regexText.Match(w.InnerXml).Success).ToList().ForEach(w =>
                {
                    w.Descendants<Text>().ToList().ForEach(t => t.Text = TryReplaceString(t.Text));
                });

        }

        private void ProcessDataFromSimpleFieldList(OpenXmlPart part, List<SimpleField> simpleFieldList)
        {
            foreach (var simpleField in simpleFieldList)
            {
                var paragraph = simpleField.Parent as Paragraph;
                var key = simpleField.InnerText;
                var newText = TryReplaceString(key);

                var run = paragraph?.Descendants<Run>().FirstOrDefault();

                if (key != newText)
                {
                    paragraph?.Append(new Run(new ElementBuilder().CreateText(newText)) { RunProperties = run?.RunProperties?.Clone() as RunProperties });
                }

                var data = TryReplaceImage(key);

                if (data == null) continue;

                var imageElementConfig = ImageElementConfig.ParseConfig(key);
                var image = ContentManager.ElementBuilder.CreateImage(data, part, imageElementConfig);
                simpleField.Remove();
                paragraph?.Append(new Run(image) { RunProperties = run?.RunProperties?.Clone() as RunProperties });
            }
        }

        /// <summary>
        /// Process the header parts of the document
        /// </summary>
        protected virtual void ProcessHeaders()
        {
            ContentManager.Content.HeaderParts.Where(w => w.Header.Descendants<SimpleField>().Any()).ToList().ForEach(
                w =>
                {
                    var simpleFieldList = w.Header.Descendants<SimpleField>().ToList();
                    ProcessDataFromSimpleFieldList(w, simpleFieldList);
                });
        }

        /// <summary>
        /// Process the footer parts of the document
        /// </summary>
        protected virtual void ProcessFooters()
        {
            ContentManager.Content.FooterParts.Where(w => w.Footer.Descendants<SimpleField>().Any()).ToList().ForEach(
                w =>
                {
                    var simpleFieldList = w.Footer.Descendants<SimpleField>().ToList();
                    ProcessDataFromSimpleFieldList(w, simpleFieldList);
                });
        }

        /// <summary>
        /// Clear merge fields to prevent any change of view from ms-office when trying to find the data source of those mergefields
        /// </summary>
        protected virtual void ClearFooters()
        {
            ContentManager.Content.FooterParts.Where(w => w.Footer.Descendants<SimpleField>().Any()).ToList().ForEach(
                w =>
                {

                    foreach (var paragraph in w.Footer.Descendants<Paragraph>())
                    {
                        paragraph.RemoveAllChildren<SimpleField>();
                    }
                });
        }

        /// <summary>
        /// Clear merge fields to prevent any change of view from ms-office when trying to find the data source of those mergefields
        /// </summary>
        protected virtual void ClearHeaders()
        {
            ContentManager.Content.HeaderParts.Where(w => w.Header.Descendants<SimpleField>().Any()).ToList().ForEach(
                w =>
                {

                    foreach (var paragraph in w.Header.Descendants<Paragraph>())
                    {
                        paragraph.RemoveAllChildren<SimpleField>();
                    }
                });
        }

        /// <summary>
        /// Process the document regex binds, images and tables
        /// </summary>
        protected virtual void ProcessDocument()
        {
            ProcessRegex();
            ProcessImage();
            ProcessTable();
        }

        /// <summary>
        /// Executes image replacement.
        /// </summary>
        protected virtual void ProcessImage()
        {
            ContentManager.Content.Document.MainDocumentPart.Document.Descendants<Text>().ToList().ForEach(w =>
            {
                var data = TryReplaceImage(w.Text);
                if (data != null)
                {
                    var imageElementConfig = ImageElementConfig.ParseConfig(w.Text);
                    var image = ContentManager.ElementBuilder.CreateImage(data, ContentManager.Content, imageElementConfig);
                    w.Parent?.Parent?.Append(new Run(image));
                    w.Remove();
                }
            });
        }

        private void ProcessTable(ref DataTable data, ref ElementBuilder elementBuilder, ref Table table)
        {
            var captionTableRow = elementBuilder.CreateRow();
            foreach (DataColumn item in data.Columns)
            {
                var colCaption = elementBuilder.CreateParagraph(item.Caption, JustificationValues.Left, null);
                var cell = elementBuilder.CreateCell();
                cell.Append(colCaption);
                captionTableRow.Append(cell);
            }
            table.Append(captionTableRow);

            foreach (DataRow item in data.Rows)
            {
                var tableRow = elementBuilder.CreateRow();
                foreach (var i in item.ItemArray)
                {
                    var cell = elementBuilder.CreateCell();
                    var cellValue = elementBuilder.CreateParagraph(i.ToString(), JustificationValues.Left, null);
                    cell.Append(cellValue);
                    tableRow.Append(cell);
                }

                table.Append(tableRow);
            }
        }

        /// <summary>
        /// Executes table replacement.
        /// </summary>
        protected virtual void ProcessTable()
        {
            ContentManager.Content.Document.MainDocumentPart.Document.Descendants<Text>().ToList().ForEach(w =>
                {
                    var data = TryReplaceTable(w.Text);

                    if (data != null)
                    {
                        var elementBuilder = ContentManager.ElementBuilder;
                        var table = elementBuilder.CreateTable(data.Columns.Count);

                        ProcessTable(ref data, ref elementBuilder, ref table);

                        w.RemoveAllChildren();

                        //BUG: Try to solve the validation error: Description="The element has unexpected child element 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:tbl'."
                        ContentManager.Content.Document.Body.Append(table);
                    }
                }
            );
        }

        /// <summary>
        /// Process the document applying the mail merge.
        /// </summary>
        public override void Process()
        {
            base.Process();

            //Processing types, headers and footers

            ProcessDocument();
            ProcessHeaders();
            ProcessFooters();
            ClearFooters();
            ClearHeaders();

            //OpenXmlValidator validator = new OpenXmlValidator();
            //var errors = validator.Validate(ContentManager.Content);
            //Debug.Write(Environment.NewLine);
            //foreach (ValidationErrorInfo error in errors)
            //    Debug.Write(error.Description + Environment.NewLine);
        }

        /// <summary>
        /// Saves the temporary file.
        /// </summary>
        private void Save()
        {
            ContentManager?.Content?.Document?.Save();
            ContentManager.Content?.Document?.MainDocumentPart?.HeaderParts?.ToList().ForEach(h => h.Header?.Save());
            ContentManager?.Content?.Document?.MainDocumentPart?.FooterParts?.ToList().ForEach(f => f.Footer?.Save());
        }

        /// <summary>
        /// Saves the temporary file.
        /// </summary>
        public override void SaveToFile()
        {
            Save();
        }

        /// <summary>
        /// Loads a document from byte array.
        /// </summary>
        /// <param name="bytes">Array of bytes with file content.</param>
        /// <param name="tempPath">Temporary file path.</param>
        /// <param name="removeCopy">Whether should remove a copy or not.</param>
        /// <param name="autoSave">Whether it should auto save the document.</param>
        /// <returns>OpenXMLDocument processor.</returns>
        public override OpenXMLDocument LoadFromBytes(byte[] bytes, string tempPath, bool removeCopy = false, bool autoSave = false)
        {
            base.LoadFromBytes(bytes, tempPath, removeCopy, autoSave);

            DocumentSettings.AutoSave = autoSave;

            ContentManager = new DocumentHandler()
            {
                Document = WordprocessingDocument.Open(DocumentSettings.TempPath, true, new OpenSettings() { AutoSave = DocumentSettings.AutoSave })
            };

            ContentManager.Content = (ContentManager.Document as WordprocessingDocument).MainDocumentPart;

            return this;
        }

        /// <summary>
        /// Saves the temporary file.
        /// </summary>
        /// <param name="path">Path to save.</param>
        public override void SaveToFile(string path)
        {
            Save();

            var bytes = GetBytes();

            try
            {
                File.WriteAllBytes(path, bytes);
            }
            catch (Exception)
            {

                throw;
            }
        }

        /// <summary>
        /// Gets the document byte array.
        /// </summary>
        /// <returns>Array of byte with document's content.</returns>
        public override byte[] GetBytes()
        {
            var bytes = base.GetBytes();

            //This override was made because it was easier to copy the closed actual saved file rather than gets its memory copy

            ContentManager = new DocumentHandler()
            {
                Document = WordprocessingDocument.Open(DocumentSettings.TempPath, true, new OpenSettings() { AutoSave = DocumentSettings.AutoSave })
            };

            ContentManager.Content = (ContentManager.Document as WordprocessingDocument).MainDocumentPart;

            return bytes;
        }
    }
}
