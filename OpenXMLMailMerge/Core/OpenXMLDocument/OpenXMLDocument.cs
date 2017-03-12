using DocumentFormat.OpenXml.Packaging;
using OpenXMLMailMerge.Core.OpenXMLDocument.Dictionary;
using OpenXMLMailMerge.Core.OpenXMLDocument.Document;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace OpenXMLMailMerge.Core.OpenXMLDocument
{
    /// <summary>
    /// The OpenXml document processor.
    /// </summary>
    public abstract class OpenXMLDocument
    {
        /// <summary>
        /// OpenXMLDocument definitions.
        /// </summary>
        protected class OpenXMLDocumentSettings
        {
            /// <summary>
            /// Temporary path.
            /// </summary>
            public string TempPath { get; set; }
            /// <summary>
            /// Sets whether it should remove the temporary copy.
            /// </summary>
            public bool RemoveFileCopy { get; set; }
            /// <summary>
            /// Indicates whether the file should be saved automatically.
            /// </summary>
            public bool AutoSave { get; set; }
        }

        /// <summary>
        /// Document settings.
        /// </summary>
        protected OpenXMLDocumentSettings DocumentSettings { get; set; }

        /// <summary>
        /// The data dictionary with the merge content and attributes to be processed.
        /// </summary>
        protected DataDictionary Dictionary { get; } = new DataDictionary();

        /// <summary>
        /// Manages the document content.
        /// </summary>
        protected DocumentHandler ContentManager { get; set; }

        /// <summary>
        /// Adds a data to the dictionary to be merged.
        /// </summary>
        /// <param name="dataType">The mailmerge data type.</param>
        /// <param name="id">The document's field id.</param>
        /// <param name="data">Data to replace the field.</param>
        public void AddToDictionary(MailMergeDataTypeEnum dataType, string id, object data)
        {
            if (!Dictionary.ContainsKey(dataType))
            {
                var tmpDictionary = new Dictionary<string, object>();
                Dictionary.Add(dataType, tmpDictionary);
            }

            Dictionary[dataType].Add(id, data);
        }

        /// <summary>
        /// Saves the temporary file.
        /// </summary>
        public virtual void SaveToFile()
        {
            throw new InvalidOperationException("Invalid call");
        }

        /// <summary>
        /// Saves the temporary file.
        /// </summary>
        /// <param name="path">Path to save.</param>
        public virtual void SaveToFile(string path)
        {
            throw new NotImplementedException();
        }

        /// <summary>
        /// Saves the temporary stream to file.
        /// </summary>
        /// <param name="path">Path to save.</param>
        /// <param name="stream">Document's data stream.</param>
        public virtual void SaveToFile(string path, Stream stream)
        {
            throw new NotImplementedException();
        }

        /// <summary>
        /// Loads a document from file.
        /// </summary>
        /// <param name="path">Path of the document.</param>
        /// <param name="isEditable">Whether is a editable copy.</param>
        /// <param name="removeCopy">Whether should remove a copy or not.</param>
        /// <param name="autoSave">Whether it should auto save the document.</param>
        /// <returns>OpenXMLDocument processor.</returns>
        public virtual OpenXMLDocument LoadFromFile(string path, bool isEditable = true, bool removeCopy = false, bool autoSave = false)
        {
            var fileName = Path.GetFileNameWithoutExtension(path);
            var filePath = Path.GetDirectoryName(path);

            DocumentSettings = new OpenXMLDocumentSettings()
            {
                RemoveFileCopy = removeCopy,
                TempPath = Path.Combine(filePath, fileName + Convert.ToString(new Random().Next()))
            };

            File.Copy(path, DocumentSettings.TempPath);

            return this;
        }

        /// <summary>
        /// Loads a document from byte array.
        /// </summary>
        /// <param name="bytes">Array of bytes with file content.</param>
        /// <param name="tempPath">Temporary file path.</param>
        /// <param name="removeCopy">Whether should remove a copy or not.</param>
        /// <param name="autoSave">Whether it should auto save the document.</param>
        /// <returns>OpenXMLDocument processor.</returns>
        public virtual OpenXMLDocument LoadFromBytes(byte[] bytes, string tempPath, bool removeCopy = false, bool autoSave = false)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(tempPath))
                    throw new ArgumentNullException("Invalid path");

                var fileName = Path.GetFileNameWithoutExtension(tempPath);
                var filePath = Path.GetDirectoryName(tempPath);

                DocumentSettings = new OpenXMLDocumentSettings()
                {
                    RemoveFileCopy = removeCopy,
                    TempPath = Path.Combine(filePath, fileName + Convert.ToString(new Random().Next()))
                };

                File.WriteAllBytes(DocumentSettings.TempPath, bytes);

                return this;
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
        public virtual byte[] GetBytes()
        {
            ContentManager.Document.Close();
            return File.ReadAllBytes(DocumentSettings.TempPath).Clone() as byte[];
        }

        /// <summary>
        /// Validates the document, content and dictionary.
        /// </summary>
        private void Validate()
        {
            if (ContentManager == null)
                throw new NullReferenceException("Content not found");

            if (ContentManager.Document == null)
                throw new NullReferenceException("Document not found");

            if (Dictionary == null || Dictionary.Count == 0)
                throw new InvalidOperationException("No data to merge");
        }

        /// <summary>
        /// Closes the document and streams.
        /// </summary>
        public virtual void Terminate()
        {
            ContentManager.Dispose();
            Dictionary.Clear();
            if (DocumentSettings.RemoveFileCopy)
            {
                try
                {
                    File.Delete(DocumentSettings.TempPath);
                }
                catch (Exception)
                {

                    throw;
                }
            }
            DocumentSettings = null;
        }

        /// <summary>
        /// Process the document applying the mail merge.
        /// </summary>
        public virtual void Process()
        {
            try
            {
                Validate();
            }
            catch (Exception)
            {
                throw;
            }            
        }
    }
}
