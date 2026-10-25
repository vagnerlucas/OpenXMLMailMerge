using System.Globalization;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OpenXmlMailMerge.Elements
{
    /// <summary>
    /// Table element class.
    /// </summary>
    internal static class TableElement
    {
        /// <summary>
        /// Creates a Table OpenXmlElement, with its properties and grid but no rows.
        /// </summary>
        /// <param name="cols">Number of columns.</param>
        /// <param name="tableConfig">Table properties definitions.</param>
        /// <returns>Table OpenXmlElement.</returns>
        internal static OpenXmlElement CreateTable(int cols = 1, TableElementConfig tableConfig = null)
        {
            tableConfig ??= new TableElementConfig();

            var tableProperties = new TableProperties
            {
                TableWidth = new TableWidth
                {
                    Width = tableConfig.Width.ToString(CultureInfo.InvariantCulture),
                    Type = tableConfig.WidthUnitValues
                },
                // Word's default look (0x04A0): the style's header row and first column formatting, with banded rows.
                // The 2007 form (w:val only) is the one every Word version reads.
                TableLook = new TableLook { Val = "04A0" }
            };

            if (tableConfig.StyleId != null)
                tableProperties.TableStyle = new TableStyle { Val = tableConfig.StyleId };
            else
                tableProperties.TableBorders = CreateSingleBorders();

            var result = new Table();
            result.Append(tableProperties);
            result.Append(CreateGrid(cols, tableConfig.ColumnWidth));

            return result;
        }

        /// <summary>
        /// Creates a row element.
        /// </summary>
        /// <param name="isHeader">True for a header row, which Word repeats at the top of each page the table spans.</param>
        /// <returns>TableRow OpenXmlElement.</returns>
        internal static OpenXmlElement CreateRow(bool isHeader = false)
        {
            return isHeader ? new TableRow(new TableRowProperties(new TableHeader())) : new TableRow();
        }

        /// <summary>
        /// Creates a grid element.
        /// </summary>
        /// <param name="cols">Number of columns.</param>
        /// <param name="columnWidth">Width of each column in twips; 0 for no width.</param>
        /// <returns>TableGrid OpenXmlElement.</returns>
        internal static OpenXmlElement CreateGrid(int cols = 1, int columnWidth = 0)
        {
            var result = new TableGrid();
            for (var i = 0; i < cols; i++)
            {
                result.Append(columnWidth > 0
                    ? new GridColumn { Width = columnWidth.ToString(CultureInfo.InvariantCulture) }
                    : new GridColumn());
            }
            return result;
        }

        /// <summary>
        /// Creates a cell holding one paragraph with the given text.
        /// </summary>
        /// <param name="text">Text value.</param>
        /// <param name="runProperties">Formatting of the text, cloned into the cell; null to use the document's defaults.</param>
        /// <returns>TableCell OpenXmlElement.</returns>
        internal static OpenXmlElement CreateCell(string text, RunProperties runProperties = null)
        {
            var run = new Run();
            if (runProperties != null)
                run.Append(runProperties.CloneNode(true));

            run.Append(new Text(text ?? string.Empty) { Space = SpaceProcessingModeValues.Preserve });

            return new TableCell(new Paragraph(run));
        }

        private static TableBorders CreateSingleBorders()
        {
            return new TableBorders
            {
                TopBorder = new TopBorder { Val = BorderValues.Single, Color = "auto", Size = 4U, Space = 0U },
                LeftBorder = new LeftBorder { Val = BorderValues.Single, Color = "auto", Size = 4U, Space = 0U },
                RightBorder = new RightBorder { Val = BorderValues.Single, Color = "auto", Size = 4U, Space = 0U },
                BottomBorder = new BottomBorder { Val = BorderValues.Single, Color = "auto", Size = 4U, Space = 0U },
                InsideHorizontalBorder = new InsideHorizontalBorder { Val = BorderValues.Single, Color = "auto", Size = 4U, Space = 0U },
                InsideVerticalBorder = new InsideVerticalBorder { Val = BorderValues.Single, Color = "auto", Size = 4U, Space = 0U }
            };
        }
    }
}
