using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OpenXMLMailMerge.Core.OpenXMLDocument.Document.Element
{
    /// <summary>
    /// Builder to construct the OpenXml elements.
    /// </summary>
    public class ElementBuilder
    {
        /// <summary>
        /// Creates a Text element.
        /// </summary>
        /// <param name="text">Text value.</param>
        /// <returns>OpenXml Text element.</returns>
        public Text CreateText(string text)
        {
            return TextElement.CreateText(text) as Text;
        }

        /// <summary>
        /// Creates a paragraph element.
        /// </summary>
        /// <param name="text">Text value.</param>
        /// <param name="justificationValue">Justification option.</param>
        /// <param name="textElementConfig">Text properties definitions.</param>
        /// <returns>OpenXml Text within a Run and Parent element.</returns>
        public Paragraph CreateParagraph(string text, JustificationValues justificationValue, TextElementConfig textElementConfig)
        {
            return ParagraphElement.CreateParagraph(text, justificationValue,textElementConfig) as Paragraph;
        }

        /// <summary>
        /// Creates a text with a parent Run element.
        /// </summary>
        /// <param name="text">Text value.</param>
        /// <param name="textElementConfig">Text properties definitions.</param>
        /// <returns>OpenXml Text within a Run element.</returns>
        public Run CreateRun(string text, TextElementConfig textElementConfig)
        {
            return RunElement.CreateRunElement(text, textElementConfig) as Run;
        }

        /// <summary>
        /// Creates a image (drawing) element.
        /// </summary>
        /// <param name="data">Byte array with image data.</param>
        /// <param name="part">Parent element.</param>
        /// <param name="imageElementConfig">Image element definitions.</param>
        /// <returns>OpenXml Drawing (image) element.</returns>
        public Drawing CreateImage(byte[] data, OpenXmlPart part, ImageElementConfig imageElementConfig = null)
        {
            return ImageElement.CreateImage(data, part, imageElementConfig) as Drawing;
        }

        /// <summary>
        /// Creates a table element.
        /// </summary>
        /// <param name="cols">Number of columns.</param>
        /// <param name="tableElementConfig">Table properties definitions.</param>
        /// <returns>OpenXml Table element.</returns>
        public Table CreateTable(int cols, TableElementConfig tableElementConfig = null)
        {
            return TableElement.CreateTable(cols, tableElementConfig) as Table;
        }

        /// <summary>
        /// Creates a row element.
        /// </summary>
        /// <returns>OpenXml TableRow element.</returns>
        public TableRow CreateRow()
        {
            return TableElement.CreateRow() as TableRow;
        }

        /// <summary>
        /// Creates a cell element.
        /// </summary>
        /// <returns>OpenXml TableCell element.</returns>
        public TableCell CreateCell()
        {
            return TableElement.CreateCell() as TableCell;
        }

        /// <summary>
        /// Creates a grid element.
        /// </summary>
        /// <param name="cols">Number of columns.</param>
        /// <returns>OpenXml Grid element.</returns>
        public TableGrid CreateGrid(int cols = 1)
        {
            return TableElement.CreateGrid(cols) as TableGrid;
        }
    }
}
