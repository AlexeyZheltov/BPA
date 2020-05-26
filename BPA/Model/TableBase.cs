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
        public ListObject Table
        { 
            get 
            { 
                return Globals.ThisWorkbook?.Sheets[SheetName].ListObjects[TableName]; 
            } 
        }

        /// <summary>
        /// Список полей таблицы. Поле Id - обязательное во всех дочерних классах
        /// </summary>
        public virtual IDictionary<string, string> Filds { get { return _filds; } }
        private readonly Dictionary<string, string> _filds = new Dictionary<string, string>();


        /// <summary>
        /// Получение данных записи по id
        /// </summary>
        /// <param name="id">идентификатор</param>
        public void GetData(int id)
        {

        }

        /// <summary>
        /// Вставка данных в таблицу
        /// </summary>
        /// <returns></returns>
        public int Insert()
        {
            return 0;
        }

        /// <summary>
        /// Обновление данных в таблице
        /// </summary>
        public void Update()
        {
            
        }

        /// <summary>
        /// Удаление данных из таблицы
        /// </summary>
        public void Delete()
        {
            int index = FindIndexRow((int)GetParametrValue("Id"));
            if (index > 0) Table.ListRows[index].Delete();
        }

        private int FindIndexRow(int id)
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
    }
}
