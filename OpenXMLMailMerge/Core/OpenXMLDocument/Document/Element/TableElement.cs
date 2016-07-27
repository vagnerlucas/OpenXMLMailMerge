using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OpenXMLMailMerge.Core.OpenXMLDocument.Document.Element
{
    /// <summary>
    /// Table element class.
    /// </summary>
    internal static class TableElement
    {
        /// <summary>
        /// Creates a Table OpenXmlElement
        /// </summary>
        /// <param name="cols">Number of columns.</param>
        /// <param name="tableConfig">Table properties definitions.</param>
        /// <returns>Table OpenXmlElement.</returns>
        internal static OpenXmlElement CreateTable(int cols = 1, TableElementConfig tableConfig = null)
        {
            var result = new Table();
            tableConfig = tableConfig ?? new TableElementConfig();

            var tableProperties = new TableProperties()
            {
                TableWidth = new TableWidth()
                {
                    Width = tableConfig.Width.ToString(),
                    Type = tableConfig.WidthUnitValues
                },
                TableLook = new TableLook()
                {
                    Val = tableConfig.DrawLine ? "04A0" : null
                },
                TableBorders = tableConfig.DrawLine ? new TableBorders()
                {
                    TopBorder = new TopBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)0U },
                    LeftBorder = new LeftBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)0U },
                    RightBorder = new RightBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)0U },
                    BottomBorder = new BottomBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)0U },
                    InsideHorizontalBorder = new InsideHorizontalBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)0U },
                    InsideVerticalBorder = new InsideVerticalBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)0U }
                } : new TableBorders()
                {
                    TopBorder = new TopBorder() { Val = BorderValues.None, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)0U },
                    LeftBorder = new LeftBorder() { Val = BorderValues.None, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)0U },
                    RightBorder = new RightBorder() { Val = BorderValues.None, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)0U },
                    BottomBorder = new BottomBorder() { Val = BorderValues.None, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)0U },
                    InsideHorizontalBorder = new InsideHorizontalBorder() { Val = BorderValues.None, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)0U },
                    InsideVerticalBorder = new InsideVerticalBorder() { Val = BorderValues.None, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)0U }
                }
            };

            result.Append(tableProperties);

            var grid = CreateGrid(cols);

            result.Append(grid);

            return result;
        }

        /// <summary>
        /// Creates a row element.
        /// </summary>
        /// <returns>TableRow OpenXmlElement.</returns>
        internal static OpenXmlElement CreateRow()
        {
            return new TableRow();
        }

        /// <summary>
        /// Creates a grid element.
        /// </summary>
        /// <param name="cols">Number of columns.</param>
        /// <returns>TableGrid OpenXmlElement.</returns>
        internal static OpenXmlElement CreateGrid(int cols = 1)
        {
            var result = new TableGrid();
            for (int i = 0; i < cols; i++)
            {
                result.Append(new GridColumn());
            }
            return result;
        }

        /// <summary>
        /// Creates a cell element.
        /// </summary>
        /// <returns>TableCell OpenXmlElement.</returns>
        internal static OpenXmlElement CreateCell()
        {
            return new TableCell();
        }
    }
}
