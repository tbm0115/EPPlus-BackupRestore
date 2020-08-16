using OfficeOpenXml;
using System;

namespace EPPlus.BackupRestore
{
    /// <summary>
    /// An interim object referring to an individual Excel table row that represents a single entity.
    /// </summary>
    /// <typeparam name="TEntity">Reference to the data entity type.</typeparam>
    public abstract class RecordInterim<TEntity>
    {
        /// <summary>
        /// Excel table row number containing data for the entity.
        /// </summary>
        public int Row { get; protected set; }

        /// <summary>
        /// Reference to the Excel Worksheet containing the table data.
        /// </summary>
        public ExcelWorksheet Sheet { get; protected set; }

        /// <summary>
        /// Reference to the parent <see cref="SheetInterim"/>.
        /// </summary>
        public SheetInterim SheetInterim { get; protected set; }

        public RecordInterim(int row, ExcelWorksheet worksheet, SheetInterim sheetInterim)
        {
            Row = row;
            Sheet = worksheet;
            SheetInterim = sheetInterim;
        }

        /// <summary>
        /// Construction method that converts this interim object into the <see cref="TEntity"/> type.
        /// </summary>
        /// <returns></returns>
        public abstract TEntity Construct();
    }
}
