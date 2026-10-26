using System;
using System.Globalization;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OpenXmlMailMerge.Elements;

namespace OpenXmlMailMerge.Merging
{
    /// <summary>
    /// Builds the table that replaces a table field.
    /// </summary>
    internal static class TableBuilder
    {
        /// <summary>
        /// Builds the table for a table field: a header row with the column captions, repeated on each page, then one row
        /// per data row. The template decides how it looks:
        /// <list type="bullet">
        /// <item>cell text takes the field's run formatting (font, size, bold, ...), and the document's defaults otherwise;</item>
        /// <item>a STYLE parameter in the field name (TABLE@STYLE=GridTable4) applies that table style, when the document
        /// defines it; otherwise the table gets single borders;</item>
        /// <item>values are formatted with the given culture, normally the template's language.</item>
        /// </list>
        /// </summary>
        /// <param name="data">The table data.</param>
        /// <param name="field">The table field.</param>
        /// <param name="main">The template's main document part.</param>
        /// <param name="textWidthInPixels">Width between the page margins, shared by the columns; 0 when unknown.</param>
        /// <param name="culture">The culture values are formatted with.</param>
        /// <returns>The table.</returns>
        internal static Table Create(TableValue data, SimpleField field, MainDocumentPart main, int textWidthInPixels, CultureInfo culture)
        {
            var columns = data.Columns.Count;
            var config = new TableElementConfig
            {
                StyleId = TemplateFormatting.FindTableStyleId(field, main),
                // 15 twips per pixel at 96 DPI.
                ColumnWidth = textWidthInPixels > 0 ? textWidthInPixels * 15 / columns : 0
            };

            var table = TableElement.CreateTable(columns, config);
            var runProperties = field.Descendants<RunProperties>().FirstOrDefault();

            var header = TableElement.CreateRow(isHeader: true);
            foreach (var column in data.Columns)
                header.Append(TableElement.CreateCell(column, runProperties));
            table.Append(header);

            foreach (var row in data.Rows)
            {
                var tableRow = TableElement.CreateRow();
                foreach (var value in row)
                    tableRow.Append(TableElement.CreateCell(Convert.ToString(value, culture), runProperties));
                table.Append(tableRow);
            }

            return table;
        }
    }
}
