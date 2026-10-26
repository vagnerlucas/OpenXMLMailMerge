#nullable enable

using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;

namespace OpenXmlMailMerge
{
    /// <summary>
    /// Table data for a table field: column captions, and rows with one value per column.
    /// The values are copied when the table is created. Numbers and dates are formatted in the template's language.
    /// </summary>
    public sealed class TableValue
    {
        /// <summary>
        /// Creates table data.
        /// </summary>
        /// <param name="columns">The column captions, shown in the header row. At least one.</param>
        /// <param name="rows">The rows, each with exactly one value per column. A null value gives an empty cell.</param>
        /// <exception cref="ArgumentNullException">The columns, the rows, or one of the rows is null.</exception>
        /// <exception cref="ArgumentException">There are no columns, or a row doesn't have one value per column.</exception>
        public TableValue(IReadOnlyList<string?> columns, IEnumerable<IReadOnlyList<object?>> rows)
        {
            ArgumentNullException.ThrowIfNull(columns);
            ArgumentNullException.ThrowIfNull(rows);

            if (columns.Count == 0)
                throw new ArgumentException("A table needs at least one column.", nameof(columns));

            Columns = [.. columns.Select(column => column ?? string.Empty)];

            var copy = new List<IReadOnlyList<object?>>();
            foreach (var row in rows)
            {
                if (row == null)
                    throw new ArgumentNullException(nameof(rows), $"Row {copy.Count} is null.");

                if (row.Count != Columns.Count)
                    throw new ArgumentException($"Row {copy.Count} has {row.Count} values; the table has {Columns.Count} columns.", nameof(rows));

                copy.Add([.. row]);
            }

            Rows = copy;
        }

        /// <summary>
        /// The column captions.
        /// </summary>
        public IReadOnlyList<string> Columns { get; }

        /// <summary>
        /// The rows, each with one value per column.
        /// </summary>
        public IReadOnlyList<IReadOnlyList<object?>> Rows { get; }

        /// <summary>
        /// Creates table data from a <see cref="DataTable"/>: the column captions, then the rows. DBNull values give
        /// empty cells.
        /// </summary>
        /// <param name="table">The table.</param>
        /// <returns>The table data.</returns>
        /// <exception cref="ArgumentNullException">The table is null.</exception>
        /// <exception cref="ArgumentException">The table has no columns.</exception>
        public static TableValue From(DataTable table)
        {
            ArgumentNullException.ThrowIfNull(table);

            var columns = table.Columns.Cast<DataColumn>().Select(column => column.Caption).ToList();
            var rows = table.Rows.Cast<DataRow>()
                .Select(row => (IReadOnlyList<object?>)[.. row.ItemArray.Select(value => value == DBNull.Value ? null : value)]);

            return new TableValue(columns, rows);
        }
    }
}
