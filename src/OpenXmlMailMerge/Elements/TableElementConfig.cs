using DocumentFormat.OpenXml.Wordprocessing;

namespace OpenXmlMailMerge.Elements
{
    /// <summary>
    /// Table element definitions class.
    /// </summary>
    internal class TableElementConfig
    {
        /// <summary>
        /// Id of a table style defined in the document. When set, the style decides borders, shading and the look of
        /// the header row; when null, the table gets single borders.
        /// </summary>
        internal string StyleId { get; set; }
        /// <summary>
        /// Table width.
        /// </summary>
        internal int Width { get; set; } = 5000;
        /// <summary>
        /// Wordprocessing width value.
        /// </summary>
        internal TableWidthUnitValues WidthUnitValues { get; set; } = TableWidthUnitValues.Pct;
        /// <summary>
        /// Width of each grid column, in twips. 0 leaves the widths to the application that opens the document.
        /// </summary>
        internal int ColumnWidth { get; set; }
    }
}
