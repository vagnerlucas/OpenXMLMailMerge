using System;

namespace OpenXmlMailMerge
{
    /// <summary>
    /// Open XML mail merge processor.
    /// This class manipulates the OpenXml document in order to process the mail merging.
    /// </summary>
    public class OxMailMerge
    {
        /// <summary>
        /// Document processor handler.
        /// </summary>
        public OpenXmlDocument Document { get; }

        /// <summary>
        /// Creates a OpenXmlMailMerge instance with a specific document processor.
        /// </summary>
        /// <param name="documentType">Type of document processor.</param>
        public OxMailMerge(DocumentType documentType)
        {
            Document = CreateDocument(documentType);
        }

        /// <summary>
        /// Creates a document by type.
        /// </summary>
        /// <param name="documentType">Type of document processor.</param>
        /// <returns>OpenXmlDocument processor.</returns>
        private static OpenXmlDocument CreateDocument(DocumentType documentType)
        {
            return DocumentFactory.CreateDocument(documentType);
        }

        /// <summary>
        /// Sets the process to end and close all streams.
        /// </summary>
        public void Terminate()
        {
            Document.Terminate();
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
