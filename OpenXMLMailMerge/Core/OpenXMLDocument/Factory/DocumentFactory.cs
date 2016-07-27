using OpenXMLMailMerge.Core.OpenXMLDocument.ML;
using System;

namespace OpenXMLMailMerge.Core.OpenXMLDocument.Factory
{
    /// <summary>
    /// Document processor factory.
    /// </summary>
    internal static class DocumentFactory
    {
        /// <summary>
        /// Creates a document processor by the document type.
        /// </summary>
        /// <param name="documentType">Type of document processor.</param>
        /// <returns>OpenXmlDocument as document processor.</returns>
        internal static OpenXMLDocument CreateDocument(DocumentType documentType)
        {
            switch (documentType)
            {
                case DocumentType.DOC:
                    return new WordprocessingML();
                case DocumentType.XLS:
                    throw new NotImplementedException();//return new SpreadsheetML();
                case DocumentType.PPT:
                    throw new NotImplementedException();//return new PresentationML();
                default:
                    throw new ArgumentOutOfRangeException("invalid argument");
            }
        }
    }
}
