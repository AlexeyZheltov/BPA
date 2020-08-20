using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BPA.NewModel
{
    class STKItem
    {
        TableRow _row;
        public STKItem(TableRow row) => _row = row;

        #region Основные свойства
        public int Id
        {
            get => _row["№"];
            set => _row["№"] = value;
        }
        public string Article
        {
            get => _row["Артикул"];
            set => _row["Артикул"] = value;
        }
        public string STKEur
        {
            get => _row["STK 2.5, Eur"];
            set => _row["STK 2.5, Eur"] = value;
        }
        public string STKRub
        {
            get => _row["STK 2.5, руб."];
            set => _row["STK 2.5, руб."] = value;
        }
        public DateTime Date
        {
            get => _row["Дата принятия"];
            set => _row["Дата принятия"] = value;
        }

        #endregion
    }
}
