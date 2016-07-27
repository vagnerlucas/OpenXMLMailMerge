using DocumentFormat.OpenXml.Packaging;
using OpenXMLMailMerge.Core.OpenXMLDocument.Document.Element;

namespace OpenXMLMailMerge.Core.OpenXMLDocument.Document
{
    /// <summary>
    /// OpenXmlDocument helper to handle the document's content.
    /// </summary>
    public class DocumentHandler
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
    }
}
