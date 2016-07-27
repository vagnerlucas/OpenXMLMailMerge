using DocumentFormat.OpenXml.Wordprocessing;

namespace OpenXMLMailMerge.Core.OpenXMLDocument.Document.Element
{
    /// <summary>
    /// Table element definitions class.
    /// </summary>
    public class TableElementConfig
    {
        /// <summary>
        /// Option to draw border line.
        /// </summary>
        internal bool DrawLine { get; set; } = true;
        /// <summary>
        /// Table width.
        /// </summary>
        internal int Width { get; set; } = 5000;
        /// <summary>
        /// Wordprocessing width value.
        /// </summary>
        internal TableWidthUnitValues WidthUnitValues { get; set; } = TableWidthUnitValues.Pct;
    }
}
