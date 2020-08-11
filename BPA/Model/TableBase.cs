using BPA.Modules;
using Microsoft.Office.Interop.Excel;
using System;
using System.CodeDom;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows;

namespace BPA.Model
{
    /// <summary>
    /// Базовый класс таблицы
    /// </summary>
    class TableBase
    {
        /// <summary>
        /// Имя таблицы
        /// </summary>
        public virtual string TableName => "Table";
        public virtual string SheetName => "Table";

        /// <summary>
        /// Объект умной таблицы
        /// </summary>
        public ListObject Table => Globals.ThisWorkbook?.Sheets[SheetName].ListObjects[TableName];

        /// <summary>
        /// Список полей таблицы. Поле Id - обязательное во всех дочерних классах
        /// </summary>
        public virtual IDictionary<string, string> Filds { get { return _filds; } }
        private readonly Dictionary<string, string> _filds = new Dictionary<string, string>();

        /// <summary>
        /// номер первой строки
        /// </summary>
        public int FirstRow
        {
            get
            {
                if (_FirstRow == 0)
                    _FirstRow = Table.ListRows[1].Range.Row;
                return _FirstRow;
            }
            set
            {
                _FirstRow = value;
            }
        }
        private int _FirstRow;

        /// <summary>
        /// последняя строка
        /// </summary>
        public int LastRow
        {
            get
            {
                if (_LastRow == 0)
                    _LastRow = Table.ListRows[Table.ListRows.Count].Range.Row;
                return _LastRow;
            }
            set
            {
                _LastRow = value;
            }
        }
        private int _LastRow;


        /// <summary>
        /// Сохранение данных в таблице
        /// </summary>
        public void Save()
        {
            if (GetParametrValueId() == 0)
            {
                int id = Insert();

                GetType().GetProperty("Id").SetValue(this, id);

                //foreach (var prop in GetType().GetProperties())
                //{
                //    if (prop.Name == "Id")
                //    {
                //        prop.SetValue(this, id);
                //    }
                //}
            }
            else
            {
                Update();
            }
        }


        /// <summary>
        /// Вставка данных в таблицу
        /// </summary>
        /// <returns>Возвращает ID новой записи</returns>
        public int Insert()
        {
            int id = GetNextId();

            ListRow row;
            if (Table.ListRows.Count == 0)
            {
                Table.ListRows.Add();
                Table.ListRows[2].Delete();
                row = Table.ListRows[1];
            }
            else if (Table.ListRows[1].Range.Cells[1, 1].Text == "")
                row = Table.ListRows[1];
            else
                Table.Resize(Table.Range.Resize[Table.Range.Rows.Count + 1]);
            row = Table.ListRows[Table.ListRows.Count];
                //row = Table.ListRows.Add();
            FillRow(row);

            row.Range[1, Table.ListColumns[Filds["Id"]].Index].Value = id;
            return id;
        }

        /// <summary>
        /// Обновление данных в таблице
        /// </summary>
        public void Update()
        {
            ListRow row = GetRow(GetParametrValueId());
            FillRow(row);
        }

        /// <summary>
        /// Заполнение строки данными из класса
        /// </summary>
        /// <param name="row"></param>
        private void FillRow(ListRow row)
        {
            foreach (ListColumn column in Table.ListColumns)
            {
                Range range = row.Range[1, column.Index];
                if (!range.HasFormula)
                    if (GetParametrValue(GetKey(column.Name)) is object obj)
                        range.Value = obj;
            }
        }

        /// <summary>
        /// Удаление данных из таблицы
        /// </summary>
        public void Delete()
        {
            ListRow row = GetRow(GetParametrValueId());
            row?.Delete();
        }

        /// <summary>
        /// Установка столбцов. Необходимо вызвать единожды для всех экземляров
        /// </summary>
        public void ReadColNumbers()
        {
            string buffer = "";
            try
            {
                Dictionary<string, int> ColDict = new Dictionary<string, int>();
                foreach (KeyValuePair<string, string> item in Filds)
                {
                    buffer = item.Value;
                    try
                    {
                        ColDict.Add(item.Key, Table.ListColumns[buffer].Index);
                    } catch
                    {

                    }
                }
                PropertyInfo pi = GetType().GetProperty("ColDict");
                pi.SetValue(this, ColDict);
            }
            catch
            {
                throw new ApplicationException($"Ошибка в поиске столбцов { SheetName }");
            }
        }

        /// <summary>
        /// Запись свойств класса данными из строки ListRow
        /// </summary>
        /// <param name="row">Строка таблицы</param>
        protected void SetProperty(ListRow row)
        {
            //PropertyInfo[] pi = GetType().GetProperties();
            //PropertyInfo[] pir = GetType().GetRuntimeProperties().ToArray();
            Dictionary<string, int> ColDict = (Dictionary<string, int>)GetType().GetProperty("ColDict").GetValue(this);
            object[,] buffer = row.Range.Value;
            //Запомнить столбцы, читать строку целиком.

            foreach (var prop in GetType().GetProperties())
            {
                if (Filds.ContainsKey(prop.Name))
                {
                    try
                    {
                        //prop.SetValue(this, Convert.ChangeType(row.Range[1, Table.ListColumns[Filds[prop.Name]].Index].Value, prop.PropertyType));
                        if (!FunctionsForExcel.IsRangeValueError(buffer[1, ColDict[prop.Name]]))
                            prop.SetValue(this, Convert.ChangeType(buffer[1, ColDict[prop.Name]], prop.PropertyType));
                        else
                            prop.SetValue(this, prop.PropertyType.IsValueType ? Activator.CreateInstance(prop.PropertyType) : null);
                    }
                    catch
                    {

                    }
                }
            }
        }

        public ListRow GetRow(int id)
        {
            int index = FindIndexRow(id);
            if (index == 0) return null;
            return Table.ListRows[index];
        }


        public ListRow GetRow(string fildName, object value, Range afterCell = null)
        {
            ListRow listRow = null;
            Range range;
            if (afterCell != null)
            {
                range = Table.ListColumns[Filds[fildName]].Range.Find(value, After: afterCell, LookAt: XlLookAt.xlWhole);
            }
            else
            {
                range = Table.ListColumns[Filds[fildName]].Range.Find(value, LookAt: XlLookAt.xlWhole);
            }

            if (range != null)
            {
                listRow = Table.ListRows[range.Row - Table.HeaderRowRange.Row];
            }
            return listRow;
        }

        /// <summary>
        /// Заполнение свойств класса по Id
        /// </summary>
        /// <param name="id">идентификатор записи</param>
        public void SetProperty(int id)
        {
            int index = FindIndexRow(id);
            if (index == 0) return;
            SetProperty(Table.ListRows[index]);
        }

        /// <summary>
        /// Поиск индекса по id
        /// </summary>
        /// <param name="id">идентификатор записи</param>
        /// <returns></returns>
        public int FindIndexRow(int id)
        {
            Range range = Table.ListColumns[Filds["Id"]].Range.Find(id, LookAt: XlLookAt.xlWhole);
            if (range == null) return 0;

            return range.Row - Table.Range.Row;
        }

        /// <summary>
        /// Получение значения свойства по его наименованию
        /// </summary>
        /// <param name="name">Имя свойства</param>
        /// <returns></returns>
        protected object GetParametrValue(string name)
        {
            foreach (var prop in GetType().GetProperties())
            {
                if (prop.Name == name)
                {
                    return prop.GetValue(this);
                }
            }
            return null;
        }

        public int GetParametrValueId() => (int)GetType().GetProperty("Id").GetValue(this);


        /// <summary>
        /// Получение ключа поля по наименованию столбца Table
        /// </summary>
        /// <param name="keyValue"></param>
        /// <returns></returns>
        private string GetKey(string keyValue)
        {
            var quere = (from pair in Filds
                         where pair.Value == keyValue
                         select pair.Key).ToList();

            return quere.Count > 0 ? quere[0] : string.Empty;

            //foreach (var pair in Filds)
            //{
            //    if (pair.Value == keyValue) return pair.Key;
            //}
            //return string.Empty;
        }

        private int GetNextId()
        {
            return (int)Globals.ThisWorkbook.Application.WorksheetFunction.Max(Table.ListColumns[Filds["Id"]].Range) + 1;
        }

        /// <summary>
        /// Сортировка таблицы
        /// </summary>
        public void Sort(string sortFildName)
        {
            Table.Sort.SortFields.Clear();
            Table.Sort.SortFields.Add(Key: Table.ListColumns[Filds[sortFildName]].Range,
                                        XlSortOn.xlSortOnValues,
                                        XlSortOrder.xlAscending,
                                        XlSortDataOption.xlSortNormal);
            Table.Sort.Header = XlYesNoGuess.xlYes;
            Table.Sort.MatchCase = false;
            Table.Sort.Orientation = XlSortOrientation.xlSortColumns;
            Table.Sort.SortMethod = XlSortMethod.xlPinYin;
            Table.Sort.Apply();
        }

        public void Mark(string fildNameToMark)
        {
            Dictionary<string, int> ColDict = (Dictionary<string, int>)GetType().GetProperty("ColDict").GetValue(this);
            ListRow row = GetRow(GetParametrValueId());
            row.Range[1, ColDict[fildNameToMark]].Interior.Color = 65535;
        }

        public void ClearTable()
        {
            if (Table.ListRows.Count < 1) return; 

            Table.DataBodyRange.Rows.Delete();
            //if (Table.ListRows.Count < 1)
            //    return;

            //for (double rw = Table.ListRows.Count; rw > 0; rw--)
            //{
            //    Table.ListRows[rw].Delete();
            //}
        }
    }
}
