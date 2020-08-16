using OfficeOpenXml;
using System;
using System.Collections.Generic;

namespace EPPlus.BackupRestore.Contracts
{
    public interface IExportSheet
    {
        Dictionary<string, Delegate> Properties { get; }

        ExcelWorksheet BuildSheet<TDbEntity>(ExcelPackage package, IEnumerable<TDbEntity> source);
    }
}
