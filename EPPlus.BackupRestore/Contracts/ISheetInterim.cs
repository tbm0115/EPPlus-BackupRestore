using OfficeOpenXml;

namespace EPPlus.BackupRestore
{
    /// <summary>
    /// For implementing custom interim structures designed to translate collections of data in Excel Worksheets.
    /// </summary>
    public interface ISheetInterim
    {
        /// <summary>
        /// Reference to the Excel Worksheet name.
        /// </summary>
        string SheetName { get; set; }

        /// <summary>
        /// Reference to the found Excel Worksheet.
        /// </summary>
        ExcelWorksheet Sheet { get; set; }

        /// <summary>
        /// Reference to the data table row for which data records start, not including the header.
        /// </summary>
        int StartRow { get; set; }

        /// <summary>
        /// A method for iterating through the contents of the data table.
        /// </summary>
        void Process();
    }
}
