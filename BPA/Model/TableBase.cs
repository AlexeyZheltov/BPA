using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

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
        /// Сохранение данных в таблице
        /// </summary>
        public void Save()
        {
            if ((int)GetParametrValue("Id") == 0)
            {
                int id = Insert();

                foreach (var prop in GetType().GetProperties())
                {
                    if (prop.Name == "Id")
                    {
                        prop.SetValue(this, id);
                    }
                }
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
            ListRow row = Table.ListRows.Add();
            FillRow(row);
            int id = GetNextId();
            row.Range[1, Table.ListColumns[Filds["Id"]].Index].Value = id;
            return id;
        }

        /// <summary>
        /// Обновление данных в таблице
        /// </summary>
        public void Update()
        {
            ListRow row = GetRow((int)GetParametrValue("Id"));
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
                if (!range.HasFormula) range.Value = GetParametrValue(GetKey(column.Name));
            }
        }

        /// <summary>
        /// Удаление данных из таблицы
        /// </summary>
        public void Delete()
        {
            ListRow row = GetRow((int)GetParametrValue("Id"));
            row?.Delete();
        }
        
        /// <summary>
        /// Запись свойств класса данными из строки ListRow
        /// </summary>
        /// <param name="row">Строка таблицы</param>
        public void SetProperty(ListRow row)
        {
            foreach (var prop in GetType().GetProperties())
            {
                if (Filds.ContainsKey(prop.Name))
                {
                    prop.SetValue(this, row.Range[1, Filds[prop.Name]]);
                }
            }
        }

        private ListRow GetRow(int id)
        {
            int index = FindIndexRow(id);
            if (index == 0) return null;
            return Table.ListRows[index];
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
        private object GetParametrValue(string name)
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

        /// <summary>
        /// Получение ключа поля по наименованию столбца Table
        /// </summary>
        /// <param name="keyValue"></param>
        /// <returns></returns>
        private string GetKey(string keyValue)
        {
            foreach (var pair in Filds)
            {
                if (pair.Value == keyValue) return pair.Key;
            }
            return string.Empty;
        }

        private int GetNextId()
        {
            return (int)Globals.ThisWorkbook.Application.WorksheetFunction.Max(Table.ListColumns[Filds["Id"]].Range) + 1;
        }
    }
}
