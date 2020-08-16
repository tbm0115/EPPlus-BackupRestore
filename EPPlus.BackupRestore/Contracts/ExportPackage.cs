using System;
using OfficeOpenXml;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EPPlus.BackupRestore.Contracts
{
    public class ExportPackage : IList<IExportSheet>
    {
        private List<IExportSheet> _exportSheets = new List<IExportSheet>();

        public ExportPackage()
        {

        }

        public ExcelPackage Build<TSourceCollection>(IEnumerable<TSourceCollection> sources) where TSourceCollection : IEnumerable
        {
            ExcelPackage package = new ExcelPackage();

            Dictionary<Type, IEnumerable> sourceTypes = sources.ToDictionary(s => s.GetType().GetGenericArguments().FirstOrDefault(), s => (IEnumerable)s);

            foreach (IExportSheet exportSheet in _exportSheets)
            {
                Type sheetType = exportSheet.GetType().GetGenericArguments().FirstOrDefault();

                IEnumerable source;
                if (sourceTypes.TryGetValue(sheetType, out source))
                {
                    sheetType.GetMethod("BuildSheet").Invoke(exportSheet, new object[] { package, source });
                }
            }

            return package;
        }

        public IExportSheet this[int index]
        {
            get => _exportSheets[index];
            set
            {
                _exportSheets[index] = value;
            }
        }

        public int Count => _exportSheets.Count;

        public bool IsReadOnly => false;

        public void Add(IExportSheet item) => _exportSheets.Add(item);

        public void Clear() => _exportSheets.Clear();

        public bool Contains(IExportSheet item) => _exportSheets.Contains(item);

        public void CopyTo(IExportSheet[] array, int arrayIndex) => _exportSheets.CopyTo(array, arrayIndex);

        public IEnumerator<IExportSheet> GetEnumerator() => _exportSheets.GetEnumerator();

        public int IndexOf(IExportSheet item) => _exportSheets.IndexOf(item);

        public void Insert(int index, IExportSheet item) => _exportSheets.Insert(index, item);

        public bool Remove(IExportSheet item) => _exportSheets.Remove(item);

        public void RemoveAt(int index) => _exportSheets.RemoveAt(index);

        IEnumerator IEnumerable.GetEnumerator() => _exportSheets.GetEnumerator();
    }
}
