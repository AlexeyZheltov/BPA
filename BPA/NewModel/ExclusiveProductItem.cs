using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BPA.NewModel
{
    class ExclusiveProductItem
    {
        TableRow _row;
        public ExclusiveProductItem(TableRow row) => _row = row;

        #region Свойства таблицы
        public int Id
        {
            get => _row["№"];
            set => _row["№"] = value;
        }
        public string Name
        {
            get => _row["Эксклюзивность"];
            set => _row["Эксклюзивность"] = value;
        }
        #endregion
    }
}
