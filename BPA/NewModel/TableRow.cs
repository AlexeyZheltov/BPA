using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BPA.NewModel
{
    class TableRow
    {
        Dictionary<string, SheetColumn> _columns = new Dictionary<string, SheetColumn>();
        dynamic[] _data;

        public TableRow(dynamic[] data, Dictionary<string, SheetColumn> columns)
        {
            _data = data;
            _columns = columns;
        }

        /// <summary>
        /// Получить или установить значение ячейки по номеру строки и номеру столбца
        /// </summary>
        /// <param name="r">Номер строки</param>
        /// <param name="c">Номер столбца</param>
        /// <returns></returns>
        public dynamic this[int c]
        {
            get => _data[c];
            set => _data[c] = value;
        }

        /// <summary>
        /// Получить или установить значение ячейки по номеру строки и имени столбца
        /// </summary>
        /// <param name="r">Номер строки</param>
        /// <param name="c">Имя столбца</param>
        /// <returns></returns>
        public dynamic this[string c]
        {
            get => _data[_columns[c].Column];
            set => _data[_columns[c].Column] = value;
        }
    }
}
