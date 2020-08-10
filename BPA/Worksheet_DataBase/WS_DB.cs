using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BPA.Worksheet_DataBase
{
    class WS_DB
    {
        Dictionary<string, SheetColumn> _columns = new Dictionary<string, SheetColumn>();
        List<SheetColumn> _columns_by_number = new List<SheetColumn>();
        List<dynamic[]> _data = new List<dynamic[]>();
        TrueExcel.Range _startCell;
        /// <summary>
        /// Загрузка данных с умной таблицы
        /// </summary>

        public dynamic this[int r, int c]
        {
            get => _data[r][c];
            set => _data[r][c] = value;
        }

        public dynamic this[int r, string c]
        {
            get => _data[r][_columns[c].Column];
            set => _data[r][_columns[c].Column] = value;
        }

        public void Load(TrueExcel.ListObject table)
        {

            _columns = new Dictionary<string, SheetColumn>();
            foreach (TrueExcel.ListColumn column in table.ListColumns)
            {
                _columns.Add(column.Name,
                    new SheetColumn()
                    {
                        Column = column.Index - 1,
                        WSColumn = column.Range.Column,
                        HasFormula = column.Range.Cells[2, 1].HasFormula,
                        Name = column.Name
                    });
            }

            if (table.DataBodyRange == null)
            {
                _startCell = table.HeaderRowRange.Cells[2, 1];
            }
            else
            {
                dynamic[,] buffer = table.DataBodyRange.Value;
                _data = Arr2List(buffer);
                _startCell = table.DataBodyRange.Cells[1, 1];
            }

            _columns_by_number = (from item in _columns
                                  orderby item.Value.WSColumn ascending
                                  select item.Value).ToList();
        }

        private List<dynamic[]> Arr2List(dynamic[,] data)
        {
            List<dynamic[]> ret_value = new List<dynamic[]>();
            int arr_width = data.GetLength(1);
            for (int r = 0; r < data.GetLength(0); r++)
            {
                dynamic[] buffer = new dynamic[arr_width];
                ret_value.Add(buffer);
                for (int c = 0; c < arr_width; c++)
                {
                    buffer[c] = data[r + 1, c + 1];
                }
            }

            return ret_value;
        }
        /// <summary>
        /// Отчистка умной таблицы с последующим сохранением
        /// </summary>
        public void Save()
        {
            //dynamic[,] data_buffer = List2Arr(_data);
            //int shift = _columns_by_number[0].WSColumn;
            int firstColumn = 0;

            int lastRow = _data.Count - 1;
            var buffer = GetSolidRangeFromData(firstColumn, _data);
            TrueExcel.Worksheet ws = _startCell.Parent;
            TrueExcel.Range targetRange;
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
        private dynamic[,] List2Arr(List<dynamic[]> data)
        {
            int height = data.Count;
            int width = data[0].GetLength(0);
            Array rv = Array.CreateInstance(typeof(object), new int[] { height, width }, new int[] { 1, 1 });
            dynamic[,] ret_value = rv as dynamic[,]; //new dynamic[height, width];

            for (int r = 0; r < height; r++)
                for (int c = 0; c < width; c++)
                    ret_value[r + 1, c + 1] = data[r][c];

            return ret_value;
        }
        private (int LastColumn, dynamic[,] Data) GetSolidRangeFromData(int firstColumn, List<dynamic[]> data)
        {
            if (firstColumn >= _columns.Count) return (firstColumn, null);
            int lastColumn = GetLastColumn(firstColumn);
            if (lastColumn == -1) return (firstColumn, null);

            int rowAmount = _data.Count;
            //Нужен массив от 1
            Array rv = Array.CreateInstance(typeof(object), new int[] { rowAmount, lastColumn - firstColumn + 1 }, new int[] { 1, 1 });
            dynamic[,] buffer = rv as dynamic[,];

            for (int row = 0; row < rowAmount; row++)
            {
                int buf_col = 1;
                for (int col = firstColumn; col <= lastColumn; col++)
                    buffer[row + 1, buf_col++] = data[row][col];
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
        public int RowCount() => _data.Count;
        public int ColumnCount() => _data[0].GetLength(0);
        public bool ValidateTable(IDictionary<string, string> keys) => keys.All(x => ColumnExists(x.Value));
        public bool ColumnExists(string col_name) => _columns.ContainsKey(col_name);
        public int NextID(string id_name)
        {
            int id_col = _columns[id_name].Column;
            int max = 0;

            for (int p = 0; p < _data.Count; p++)
            {
                dynamic obj = _data[p][id_col];
                if (int.TryParse(obj?.ToString() ?? "0", out int i_obj))
                {
                    if (i_obj > max) max = i_obj;
                }
            }

            return ++max;
        }
        public int GetRow(string col_name, int id)
        {
            int id_col = _columns[col_name].Column;

            for (int row = 0; row < _data.Count; row++)
            {
                dynamic obj = _data[row][id_col];
                if (int.TryParse(obj.ToString(), out int i_obj))
                {
                    if (i_obj == id) return row;
                }
            }
            return 0;
        }
        public int GetColumnNumber(string col_name) => _columns.ContainsKey(col_name) ? _columns[col_name].Column : 0;
        public int AddRow()
        {
            _data.Add(new dynamic[_columns.Count]);
            return _data.Count - 1;
        }
    }

    struct SheetColumn
    {
        /// <summary>
        /// Номер столбца в таблице
        /// </summary>
        public int Column { get; set; }
        /// <summary>
        /// Номер столбца на листе
        /// </summary>
        public int WSColumn { get; set; }
        public bool HasFormula { get; set; }
        public string Name { get; set; }
    }
}
