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
                        if (arg.ToUpper().Contains(value.Key.ToString().ToUpper()))
                        {
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
                if (item.Key == MailMergeDataTypeEnum.Table)
                {
                    foreach (var value in item.Value)
                    {
                        if (arg.ToUpper().Contains(value.Key.ToString().ToUpper()))
                        {
                            return value.Value as DataTable;
                        }
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
            ContentManager.Content.HeaderParts.Where(w => w.Header.Descendants<SimpleField>().Count() > 0).ToList().ForEach(w =>
            {
                w.Header.Descendants<Text>().ToList().ForEach(t =>
                {
                    t.Text = TryReplaceString(t.Text).ToString();
                });
            });

            ContentManager.Content.FooterParts.Where(w => w.Footer.Descendants<SimpleField>().Count() > 0).ToList().ForEach(w =>
            {
                w.Footer.Descendants<Text>().ToList().ForEach(t =>
                {
                    t.Text = TryReplaceString(t.Text).ToString();
                });
            });

            ContentManager.Content.Document.MainDocumentPart.Document.Descendants<SimpleField>().ToList().ForEach(w =>
            {
                w.Descendants<Text>().ToList().ForEach(t =>
                {
                    t.Text = TryReplaceString(t.Text).ToString();
                });
            });

            //ContentManager.Content.Document.Save();
        }

        /// <summary>
        /// Executes image replacement.
        /// </summary>
        protected virtual void ProcessImage()
        {
            ContentManager.Content.HeaderParts.Where(w => w.Header.Descendants<SimpleField>().Count() > 0).ToList().ForEach(w =>
            {
                w.Header.Descendants<SimpleField>().ToList().ForEach(s =>
                {
                    s.Descendants<Text>().ToList().ForEach(t =>
                    {
                        var data = TryReplaceImage(t.Text);
                        if (data != null)
                        {
                            var image = ContentManager.ElementBuilder.CreateImage(data, w);
                            t.Parent.Parent.Parent.Append(new Run(image));
                            t.Text = string.Empty;
                        }
                    });
                });
            });

            ContentManager.Content.FooterParts.Where(w => w.Footer.Descendants<SimpleField>().Count() > 0).ToList().ForEach(w =>
            {
                w.Footer.Descendants<SimpleField>().ToList().ForEach(s =>
                {
                    s.Descendants<Text>().ToList().ForEach(t =>
                    {
                        var data = TryReplaceImage(t.Text);
                        if (data != null)
                        {
                            var image = ContentManager.ElementBuilder.CreateImage(data, w);
                            t.Parent.Parent.Parent.Append(new Run(image));
                            t.Remove();
                        }
                    });
                });
            });

            ContentManager.Content.Document.MainDocumentPart.Document.Descendants<SimpleField>().ToList().ForEach(w =>
            {
                w.Descendants<Text>().ToList().ForEach(t =>
                {
                    var data = TryReplaceImage(t.Text);
                    if (data != null)
                    {
                        var imageElementConfig = new ImageElementConfig();
                        imageElementConfig.ParseConfig(t.Text);
                        var image = ContentManager.ElementBuilder.CreateImage(data, ContentManager.Content, imageElementConfig);
                        w.Parent.Append(new Run(image));
                        t.Remove();
                    }
                });

            });

            //ContentManager.Content.Document.Save();
        }

        /// <summary>
        /// Executes table replacement.
        /// </summary>
        protected virtual void ProcessTable()
        {
            ContentManager.Content.HeaderParts.Where(w => w.Header.Descendants<SimpleField>().Count() > 0).ToList().ForEach(w =>
            {
                w.Header.Descendants<SimpleField>().ToList().ForEach(s =>
                {
                    s.Descendants<Text>().ToList().ForEach(t =>
                    {
                        var data = TryReplaceTable(t.Text);
                        if (data != null)
                        {
                            
                            var elementBuilder = ContentManager.ElementBuilder;
                            var table = elementBuilder.CreateTable(data.Columns.Count);
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

                            ContentManager.Content.HeaderParts.FirstOrDefault().Header.AppendChild(table);
                            t.Remove();
                        }
                    });
                });
            });

            ContentManager.Content.FooterParts.Where(w => w.Footer.Descendants<SimpleField>().Count() > 0).ToList().ForEach(w =>
            {
                w.Footer.Descendants<SimpleField>().ToList().ForEach(s =>
                {
                    s.Descendants<Text>().ToList().ForEach(t =>
                    {
                        var data = TryReplaceTable(t.Text);
                        if (data != null)
                        {
                            var elementBuilder = ContentManager.ElementBuilder;
                            var table = elementBuilder.CreateTable(data.Columns.Count);
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
                            
                            ContentManager.Content.FooterParts.FirstOrDefault().Footer.Append(table);
                            t.Remove();
                        }
                    });
                });
            });

            ContentManager.Content.Document.MainDocumentPart.Document.Descendants<SimpleField>().ToList().ForEach(w =>
            {
                w.Descendants<Text>().ToList().ForEach(t =>
                {
                    var data = TryReplaceTable(t.Text);
                    if (data != null)
                    {
                        var elementBuilder = ContentManager.ElementBuilder;
                        var table = elementBuilder.CreateTable(data.Columns.Count);
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
                       
                        w.RemoveAllChildren();

                        //BUG: Try to solve the validation error: Description="The element has unexpected child element 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:tbl'."
                        ContentManager.Content.Document.Body.Append(table);
                    }
                });
            });

            
            //ContentManager.Content.Document.Save();
        }

        /// <summary>
        /// Process the document applying the mail merge.
        /// </summary>
        public override void Process()
        {
            base.Process();

            //Processing types

            ProcessRegex();

            ProcessImage();

            ProcessTable();

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
            var header = ContentManager.Content.Document.MainDocumentPart.HeaderParts.FirstOrDefault();
            if (header != null)
                if (header.Header != null)
                    header.Header.Save();

            var footer = ContentManager.Content.Document.MainDocumentPart.FooterParts.FirstOrDefault();
            if (footer != null)
                if (footer.Footer != null)
                    footer.Footer.Save();

            ContentManager.Content.Document.Save();
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
