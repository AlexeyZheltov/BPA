using Microsoft.Office.Interop.Excel;

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BPA.Model {
    /// <summary>
    /// Справочник продуктовых календарей
    /// </summary>
    class ProductCalendar : TableBase {
        public override string TableName => "Продуктовые_календари";
        public override string SheetName => "Продуктовые календари";

        public override IDictionary<string, string> Filds {
            get {
                return _filds;
            }
        }
        private readonly Dictionary<string, string> _filds = new Dictionary<string, string>
        {
            { "Id", "№" },
            { "Name", "Название" },
            { "Path", "Путь к файлу" },
        };

        /// <summary>
        /// Идентификатор
        /// </summary>
        public int Id {
            get; set;
        }
        /// <summary>
        /// Название
        /// </summary>
        public string Name {
            get; set;
        }
        /// <summary>
        /// Путь к файлу
        /// </summary>
        public string Path {
            get; set;
        }

        public ProductCalendar() { }

        public ProductCalendar(string name) 
        {
            Name = name;
            var listRow = GetRow("Name", name);
            if (listRow != null) SetProperty(listRow);
        }

        public List<ProductCalendar> GetProducts()
        {
            List<ProductCalendar> products = new List<ProductCalendar>();
            foreach (ListRow row in Table.ListRows)
            {
                ProductCalendar product = new ProductCalendar();
                product.SetProperty(row);
                products.Add(product);
            }
            return products;
        }

        public void UpdateProductFromCalendar()
        {
            //TODO: Цикл по календарям
            //TODO: открываем книги WB

            //TODO: Цикл по продуктам

            List<Product> products = new Product().GetProducts();

            foreach (Product product in products)
            {
                // product.SetFromCalendar(WB);
            }

            //TODO: закрываем WB
        }

    }
}
