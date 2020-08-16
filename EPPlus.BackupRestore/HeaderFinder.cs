using OfficeOpenXml;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;

namespace EPPlus.BackupRestore
{
    /// <summary>
    /// A class used for discovering the positions of structured data columns in Excel Tables or Worksheets
    /// </summary>
    public class HeaderFinder : IDictionary<ColumnMapItem, int>
    {
        /// <summary>
        /// Collection of <see cref="ColumnMapItem"/>s to help delegate the locations of Source data.
        /// </summary>
        public IEnumerable<ColumnMapItem> SourceColumns { get; private set; }

        /// <summary>
        /// Reference to the Excel Worksheet row for which the column headers are defined.
        /// <para>Note: Column headers should be defined on the same row within the Worksheet.</para>
        /// </summary>
        public int HeaderRow { get; private set; }

        /// <summary>
        /// Reference to the Excel Worksheet containing the data table.
        /// </summary>
        public ExcelWorksheet Worksheet { get; private set; }

        /// <summary>
        /// Internal collection of the provided <see cref="ColumnMapItem"/>s and their respective column indices within the <see cref="Worksheet"/>.
        /// </summary>
        private Dictionary<ColumnMapItem, int> _map { get; set; } = new Dictionary<ColumnMapItem, int>();

        /// <summary>
        /// Reference to the internal mapping of <see cref="ColumnMapItem"/>s from the provided <see cref="SourceColumns"/>.
        /// </summary>
        public ICollection<ColumnMapItem> Keys => _map.Keys;

        /// <summary>
        /// Reference to the internal mapping of <see cref="ColumnMapItem"/> indices from the provided <see cref="SourceColumns"/>.
        /// </summary>
        public ICollection<int> Values => _map.Values;

        /// <summary>
        /// Number of mapped <see cref="ColumnMapItem"/>s.
        /// </summary>
        public int Count => _map.Count;

        public bool IsReadOnly => true;

        /// <summary>
        /// Gets the index of the mapped <see cref="ColumnMapItem"/> from the provided <see cref="SourceColumns"/>.
        /// </summary>
        /// <param name="key">Reference to the <see cref="ColumnMapItem"/>.</param>
        /// <returns>The column index of the header within the associated Excel Worksheet.</returns>
        public int this[ColumnMapItem key]
        {
            get => _map[key];
            set => _map[key] = value;
        }

        public HeaderFinder(int row, ExcelWorksheet worksheet, params ColumnMapItem[] columnAliases)
        {
            HeaderRow = row;
            Worksheet = worksheet;
            SourceColumns = columnAliases;
            _map = map();
        }

        private Dictionary<ColumnMapItem, int> map()
        {
            Dictionary<ColumnMapItem, int> map = new Dictionary<ColumnMapItem, int>();
            int col = 1;
            string headerName = string.Empty;
            do
            {
                ExcelRangeBase cell = Worksheet.Cells[HeaderRow, col];
                headerName = cell.GetValue<string>();
                if (!string.IsNullOrEmpty(headerName))
                {
                    ColumnMapItem alias = SourceColumns.FirstOrDefault(o => o.Aliases.Contains(headerName));
                    if (alias != null)
                    {
                        map.Add(alias, col);
                    }
                    else
                    {
                        System.Diagnostics.Debug.WriteLine($"Failed to find alias for {headerName}");
                    }
                }
                col++;
            } while (!string.IsNullOrEmpty(headerName));
            return map;
        }

        public bool ContainsKey(ColumnMapItem key) => _map.ContainsKey(key);

        public void Add(ColumnMapItem key, int value) => _map.Add(key, value);

        public bool Remove(ColumnMapItem key) => _map.Remove(key);

        public bool TryGetValue(ColumnMapItem key, out int value) => _map.TryGetValue(key, out value);

        public void Add(KeyValuePair<ColumnMapItem, int> item) => _map.Add(item.Key, item.Value);

        public void Clear() => _map.Clear();

        public bool Contains(KeyValuePair<ColumnMapItem, int> item) => _map.Contains(item);

        public void CopyTo(KeyValuePair<ColumnMapItem, int>[] array, int arrayIndex) => throw new NotImplementedException();

        public bool Remove(KeyValuePair<ColumnMapItem, int> item) => _map.Remove(item.Key);

        public IEnumerator<KeyValuePair<ColumnMapItem, int>> GetEnumerator() => _map.GetEnumerator();

        IEnumerator IEnumerable.GetEnumerator() => _map.GetEnumerator();
    }
}
