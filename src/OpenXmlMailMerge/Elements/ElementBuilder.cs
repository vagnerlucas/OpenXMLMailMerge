using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OpenXmlMailMerge.Elements
{
    /// <summary>
    /// Builder to construct the OpenXml elements.
    /// </summary>
    internal class ElementBuilder
    {
        /// <summary>
        /// Creates a image (drawing) element.
        /// </summary>
        /// <param name="data">Byte array with image data.</param>
        /// <param name="part">Parent element.</param>
        /// <param name="imageElementConfig">Image element definitions.</param>
        /// <returns>OpenXml Drawing (image) element.</returns>
        public static Drawing CreateImage(byte[] data, OpenXmlPart part, ImageElementConfig imageElementConfig = null)
        {
            return ImageElement.CreateImage(data, part, imageElementConfig) as Drawing;
        }

        /// <summary>
        /// Creates a table element.
        /// </summary>
        /// <param name="cols">Number of columns.</param>
        /// <param name="tableElementConfig">Table properties definitions.</param>
        /// <returns>OpenXml Table element.</returns>
        public static Table CreateTable(int cols, TableElementConfig tableElementConfig = null)
        {
            return TableElement.CreateTable(cols, tableElementConfig) as Table;
        }

        /// <summary>
        /// Creates a row element.
        /// </summary>
        /// <param name="isHeader">True for a header row, repeated at the top of each page.</param>
        /// <returns>OpenXml TableRow element.</returns>
        public static TableRow CreateRow(bool isHeader = false)
        {
            return TableElement.CreateRow(isHeader) as TableRow;
        }

        /// <summary>
        /// Creates a cell holding one paragraph with the given text.
        /// </summary>
        /// <param name="text">Text value.</param>
        /// <param name="runProperties">Formatting of the text; null to use the document's defaults.</param>
        /// <returns>OpenXml TableCell element.</returns>
        public static TableCell CreateCell(string text, RunProperties runProperties = null)
        {
            return TableElement.CreateCell(text, runProperties) as TableCell;
        }
    }
}
