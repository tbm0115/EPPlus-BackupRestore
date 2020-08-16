using OfficeOpenXml;
using OfficeOpenXml.ConditionalFormatting;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;

namespace EPPlus.BackupRestore.Contracts
{
    public abstract class ExportSheet<TEntity> : IExportSheet
    {
        public Dictionary<string, Delegate> Properties { get; private set; } = new Dictionary<string, Delegate>();

        public string SheetName { get; set; } = typeof(TEntity).Name;

        public ExportSheet()
        {

        }

        public void AddProperty<TProp>(Expression<Func<TEntity, TProp>> key, string name = "")
        {
            Type propType = typeof(TProp);
            if (string.IsNullOrEmpty(name))
            {
                name = propType.Name;
            }
            if (Properties.ContainsKey(name))
            {
                int nameIteration = 0;
                bool nameUnique = false;
                do
                {
                    nameIteration++;
                    if (!Properties.ContainsKey($"{name}_{nameIteration}"))
                    {
                        nameUnique = true;
                    }
                } while (!nameUnique);
                name = $"{name}_{nameIteration}";
            }
            Properties.Add(name, key.Compile());
        }

        ExcelWorksheet IExportSheet.BuildSheet<TDbEntity>(ExcelPackage package, IEnumerable<TDbEntity> source) => BuildSheet(package, (IEnumerable<TEntity>)source);
            
        public virtual ExcelWorksheet BuildSheet(ExcelPackage package, IEnumerable<TEntity> source)
        {
            ExcelWorksheet worksheet = package.Workbook.Worksheets.Add(SheetName);

            List<string> headers = Properties.Keys.ToList();
            for (int i = 0; i < headers.Count; i++)
            {
                var headerCell = worksheet.Cells[1, i + 1];
                headerCell.Value = headers[i];
            }

            int row = 2;
            foreach (var item in source)
            {
                for (int c = 0; c < headers.Count; c++)
                {
                    var valueCell = worksheet.Cells[row, c + 1];
                    object value = string.Empty;
                    try
                    {
                        value = Properties[headers[c]].DynamicInvoke(new object[] { item })?.ToString();
                    }
                    catch (Exception ex)
                    {
                        valueCell.AddComment($"Error: {ex}", "RevolutionSystem");
                    }
                    valueCell.Value = value;
                }
                row++;
            }
            row--; // Go back one because of for loop

            ExcelAddressBase tableRange = new ExcelAddressBase(1, 1, row, headers.Count);
            var excelTable = worksheet.Tables.Add(tableRange, typeof(TEntity).Name);


            return worksheet;
        }
    }
}
