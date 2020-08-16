using OfficeOpenXml;
using System.Collections.Generic;

namespace EPPlus.BackupRestore
{
    /// <summary>
    /// A generic container object used to find structured data in an Excel Worksheet.
    /// </summary>
    /// <typeparam name="TRecord">A reference to the strongly implemented <see cref="RecordInterim{TEntity}"/>.</typeparam>
    /// <typeparam name="TEntity"></typeparam>
    public abstract class ScopedSheetInterim<TRecord, TEntity> : SheetInterim where TRecord : RecordInterim<TEntity>
    {
        /// <summary>
        /// A collection for storing discovered <see cref="TRecord"/> objects within the provided Excel Worksheet.
        /// </summary>
        public List<TRecord> Records { get; set; } = new List<TRecord>();

        public ScopedSheetInterim(ExcelWorkbook workbook) : base(workbook)
        {

        }
    }
}
