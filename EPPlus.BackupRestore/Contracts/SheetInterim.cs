using OfficeOpenXml;
using System;

namespace EPPlus.BackupRestore
{
    /// <summary>
    /// A generic container object used to find structured data in an Excel Worksheet.
    /// </summary>
    public abstract class SheetInterim : ISheetInterim
    {
        /// <summary>
        /// Reference to the Excel Worksheet name.
        /// </summary>
        public string SheetName { get; set; }

        /// <summary>
        /// Internal reference to the Excel Workbook.
        /// </summary>
        private ExcelWorkbook _workbook { get; set; }

        /// <summary>
        /// Reference to the Excel Worksheet.
        /// </summary>
        public ExcelWorksheet Sheet { get; set; }

        /// <summary>
        /// Reference to the row for which structured table data, not including the header, can be found.
        /// </summary>
        public int StartRow { get; set; }

        /// <summary>
        /// Reference to the dynamic <see cref="HeaderFinder"/> to help discover the positions of structured table columns
        /// </summary>
        public HeaderFinder Header { get; set; }

        public SheetInterim(ExcelWorkbook workbook)
        {
            _workbook = workbook;
        }

        /// <summary>
        /// Initializes the interim by discovering the Excel Worksheet.
        /// </summary>
        public void Initialize()
        {
            Sheet = _workbook.Worksheets[SheetName];
        }

        /// <summary>
        /// A method for iterating through the Excel Worksheet's structured data table.
        /// </summary>
        public abstract void Process();
    }
}
