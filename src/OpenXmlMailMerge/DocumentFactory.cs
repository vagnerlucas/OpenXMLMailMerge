using System;

namespace OpenXmlMailMerge
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
        internal static OpenXmlDocument CreateDocument(DocumentType documentType)
        {
            return documentType switch
            {
                DocumentType.Doc => new WordprocessingMl(),
                _ => throw new ArgumentOutOfRangeException(nameof(documentType), documentType, "Unknown document type.")
            };
        }
    }
}
