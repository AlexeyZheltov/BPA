using BPA.Modules;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BPA.NewModel
{
    class ProductCalendarItem
    {
        TableRow _row;
        public ProductCalendarItem(TableRow row) => _row = row;

        #region Основные свойства

        public int Id
        {
            get => _row["№"];
            set => _row["№"] = value;
        }

        public string Name
        {
            get => _row["Название"];
            set => _row["Название"] = value;
        }

        public string Path
        {
            get => _row["Путь к файлу"];
            set => _row["Путь к файлу"] = value;
        }
        #endregion

        public void UpdateFromCalendar(FileCalendar fileCalendar) 
        {
            Name = fileCalendar.FileName;
            Path = fileCalendar.FileAddress;
        }

        public void UpdateProductsFromCalendar(List<FileCalendar.ProductFromCalendar> productsFrom)
        {

        }
    }
}
