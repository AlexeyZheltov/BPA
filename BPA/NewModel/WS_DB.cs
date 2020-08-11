using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace BPA.NewModel
{
    class WS_DB
    {
        Dictionary<string, NewModel.SheetColumn> _columns = new Dictionary<string, SheetColumn>();
        List<SheetColumn> _columns_by_number = new List<SheetColumn>();
        List<Dynamic[]> _data = new List<Dynamic[]>();
        Excel.ListObject _table;

        /// <summary>
        /// Получить или установить значение ячейки по номеру строки и номеру столбца
        /// </summary>
        /// <param name="r">Номер строки</param>
        /// <param name="c">Номер столбца</param>
        /// <returns></returns>
        public Dynamic this[int r, int c]
        {
            get => _data[r][c];
            set => _data[r][c] = value;
        }

        /// <summary>
        /// Получить или установить значение ячейки по номеру строки и имени столбца
        /// </summary>
        /// <param name="r">Номер строки</param>
        /// <param name="c">Имя столбца</param>
        /// <returns></returns>
        public Dynamic this[int r, string c]
        {
            get => _data[r][_columns[c].Column];
            set => _data[r][_columns[c].Column] = value;
        }

        public TableRow this[int r] => new TableRow(_data[r], _columns);

        /// <summary>
        /// Позволяет проходить по строкам в ForEach
        /// </summary>
        /// <returns>строка в виде Dynamic[]</returns>
        public IEnumerator<TableRow> GetEnumerator()
        {
            foreach (var item in _data) yield return new TableRow(item, _columns);
        }

        /// <summary>
        /// Загружает данные умной таблицы
        /// </summary>
        /// <param name="table">Объект умной таблицы</param>
        public void Load(Excel.ListObject table)
        {
            _table = table;
            _columns = new Dictionary<string, SheetColumn>();
            foreach (Excel.ListColumn column in table.ListColumns)
            {
                _columns.Add(column.Name,
                    new SheetColumn()
                    {
                        Column = column.Index - 1,
                        HasFormula = column.Range.Cells[2, 1].HasFormula
                    });
            }

            if (table.DataBodyRange != null)
            {
                object[,] buffer = table.DataBodyRange.Value;
                _data = Arr2List(buffer);
            }

            _columns_by_number = (from item in _columns
                                  orderby item.Value.Column ascending
                                  select item.Value).ToList();
        }

        private List<Dynamic[]> Arr2List(object[,] data)
        {
            List<Dynamic[]> ret_value = new List<Dynamic[]>();
            int arr_width = data.GetLength(1);
            for (int r = 0; r < data.GetLength(0); r++)
            {
                Dynamic[] buffer = new Dynamic[arr_width];
                ret_value.Add(buffer);
                for (int c = 0; c < arr_width; c++)
                {
                    buffer[c] = new Dynamic(data[r + 1, c + 1]);
                }
            }

            return ret_value;
        }

        /// <summary>
        /// Отчистка умной таблицы с последующим сохранением
        /// </summary>
        public void Save()
        {
            //_table.DataBodyRange?.Clear();
            Excel.Range _startCell = _table.HeaderRowRange.Cells[2, 1];

            int firstColumn = 0;
            int lastRow = _data.Count - 1;
            var buffer = GetSolidRangeFromData(firstColumn, _data);
            Excel.Worksheet ws = _startCell.Parent;
            Excel.Range targetRange;
            while (buffer.LastColumn < _columns.Count)// если первый столбец формула - выход
            {
                if (buffer.Data != null)
                {
                    //Проверить правильность создания диапазона
                    targetRange = ws.Range[_startCell.Offset[0, firstColumn], _startCell.Offset[lastRow, buffer.LastColumn]];
                    targetRange.Value = buffer.Data;
                }
                firstColumn = buffer.LastColumn + 1;
                buffer = GetSolidRangeFromData(firstColumn, _data);
            }
        }

        private (int LastColumn, object[,] Data) GetSolidRangeFromData(int firstColumn, List<Dynamic[]> data)
        {
            if (firstColumn >= _columns.Count) return (firstColumn, null);
            int lastColumn = GetLastColumn(firstColumn);
            if (lastColumn == -1) return (firstColumn, null);

            int rowAmount = _data.Count;
            //Нужен массив от 1
            Array rv = Array.CreateInstance(typeof(object), new int[] { rowAmount, lastColumn - firstColumn + 1 }, new int[] { 1, 1 });
            object[,] buffer = rv as object[,];

            for (int row = 0; row < rowAmount; row++)
            {
                int buf_col = 1;
                for (int col = firstColumn; col <= lastColumn; col++)
                    buffer[row + 1, buf_col++] = data[row][col].Value;
            }

            return (lastColumn, buffer);
        }

        private int GetLastColumn(int firstColumn)
        {
            if (_columns_by_number[firstColumn].HasFormula) return -1;
            if (firstColumn == _columns.Count - 1) return firstColumn;
            int lastColumn = firstColumn;

            for (int ptr = firstColumn + 1; ptr < _columns_by_number.Count; ptr++)
            {
                if (_columns_by_number[ptr].HasFormula) break;
                lastColumn++;
            }
            return lastColumn;
        }

        /// <summary>
        /// удалит строку
        /// </summary>
        public void Delete(int row) => _data.RemoveAt(row);

        /// <summary>
        /// Колличество строк в таблице
        /// </summary>
        /// <returns></returns>
        public int RowCount() => _data.Count;

        /// <summary>
        /// Колличество столбцов в таблице
        /// </summary>
        /// <returns></returns>
        public int ColumnCount() => _columns.Count;

        /// <summary>
        /// Проверить существует ли столбец с данным именем
        /// </summary>
        /// <param name="col_name">Имя столбца</param>
        /// <returns></returns>
        public bool ColumnExists(string col_name) => _columns.ContainsKey(col_name);

        /// <summary>
        /// Получить следующий ID
        /// </summary>
        /// <param name="id_name">Имя столбца с ID</param>
        /// <returns></returns>
        public int NextID(string id_name)
        {
            int id_col = _columns[id_name].Column;
            int max = 0;

            for (int p = 0; p < _data.Count; p++)
            {
                Dynamic obj = _data[p][id_col];
                if (int.TryParse(obj?.ToString() ?? "0", out int i_obj))
                {
                    if (i_obj > max) max = i_obj;
                }
            }

            return ++max;
        }

        /// <summary>
        /// Получить номер строки по имени столбйа и значению
        /// </summary>
        /// <param name="col_name">Имя столбца</param>
        /// <param name="id">Значение</param>
        /// <returns></returns>
        public int GetRow(string col_name, string id)
        {
            int id_col = _columns[col_name].Column;

            for (int row = 0; row < _data.Count; row++)
            {
                Dynamic obj = _data[row][id_col];
                if (obj == id) return row;
            }
            return 0;
        }

        /// <summary>
        /// Узнать номер столбца таблицы по названию столбца
        /// </summary>
        /// <param name="col_name">Название столбца</param>
        /// <returns></returns>
        public int GetColumnNumber(string col_name) => _columns.ContainsKey(col_name) ? _columns[col_name].Column : 0;

        /// <summary>
        /// Добавить строку
        /// </summary>
        /// <returns>Номер добавленной строки</returns>
        public int AddRow()
        {
            _data.Add(new Dynamic[_columns.Count]);
            return _data.Count - 1;
        }
    }
}
