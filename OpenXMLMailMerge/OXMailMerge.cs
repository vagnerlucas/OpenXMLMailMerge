using OpenXMLMailMerge.Core.OpenXMLDocument;
using OpenXMLMailMerge.Core.OpenXMLDocument.Factory;
using System;

namespace OpenXMLMailMerge
{
    /// <summary>
    /// OpenXML MailMerge processor.
    /// This class manipulates the OpenXml document in order to process the mail merging.
    /// </summary>
    public class OXMailMerge
    {
        /// <summary>
        /// Document processor handler.
        /// </summary>
        public OpenXMLDocument Document { get; set; }

        /// <summary>
        /// Default constructor.
        /// </summary>
        public OXMailMerge() { }

        /// <summary>
        /// Creates a OpenXMLMailMerge instance with a specific document processor.
        /// </summary>
        /// <param name="documentType">Type of document processor.</param>
        public OXMailMerge(DocumentType documentType)
        {
            CreateDocument(documentType);
        }

        /// <summary>
        /// Creates a OpenXMLMailMerge instance with a external document processor.
        /// </summary>
        /// <param name="externalDocumentProcessor">External OpenXMLDocument processor.</param>
        public OXMailMerge(OpenXMLDocument externalDocumentProcessor)
        {
            Document = externalDocumentProcessor;
        }

        /// <summary>
        /// Creates a document by type.
        /// </summary>
        /// <param name="documentType">Type of document processor.</param>
        /// <returns>OpenXMLDocument processor.</returns>
        public OpenXMLDocument CreateDocument(DocumentType documentType)
        {
            Document = DocumentFactory.CreateDocument(documentType);
            return Document;
        }

        /// <summary>
        /// Sets the process to end and close all streams.
        /// </summary>
        public void Terminate()
        {
            try
            {
                Document.Terminate();
            }
            catch (Exception)
            {

                throw;
            }
        }

        /// <summary>
        /// Process a document and executes the mail merging activity.
        /// </summary>
        public void Process()
        {
            Document.Process();
        }
    }
}
