using DocumentFormat.OpenXml.Packaging;
using OpenXMLMailMerge.Core.OpenXMLDocument.Document.Element;
using System;

namespace OpenXMLMailMerge.Core.OpenXMLDocument.Document
{
    /// <summary>
    /// OpenXmlDocument helper to handle the document's content.
    /// </summary>
    public sealed class DocumentHandler : IDisposable
    {
        /// <summary>
        /// The OpenXmlPackage itself.
        /// </summary>
        public OpenXmlPackage Document { get; set; }
        /// <summary>
        /// The main document part.
        /// </summary>
        public MainDocumentPart Content { get; set; }
        /// <summary>
        /// The element builder helper.
        /// </summary>
        public ElementBuilder ElementBuilder { get; set; }  = new ElementBuilder();

        #region IDisposable Support
        private bool disposedValue = false; // To detect redundant calls

        /// <summary>
        /// Disposable pattern
        /// </summary>
        /// <param name="disposing"></param>
        private void Dispose(bool disposing)
        {
            if (disposedValue) return;

            Document?.Close();

            if (disposing)
            {
                Document = null;
                Content = null;
                ElementBuilder = null;
            }

            disposedValue = true;
        }

        // This code added to correctly implement the disposable pattern.
        /// <summary>
        /// Disposable pattern
        /// </summary>
        public void Dispose()
        {
            // Do not change this code. Put cleanup code in Dispose(bool disposing) above.
            Dispose(true);
            // TODO: uncomment the following line if the finalizer is overridden above.
            // GC.SuppressFinalize(this);
        }
        #endregion
    }
}
