using System.Collections.Generic;

namespace EPPlus.BackupRestore
{
    public class ColumnMapItem
    {
        public List<string> Aliases { get; set; }

        public ColumnMapItem(params string[] aliases)
        {
            Aliases = new List<string>(aliases);
        }
    }
}
