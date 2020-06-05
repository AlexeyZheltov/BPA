using BPA.Forms;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace BPA.Model {
    /// <summary>
    /// Справочник продуктовых календарей
    /// </summary>
    class ProductCalendar : TableBase {
        public override string TableName => "Продуктовые_календари";
        public override string SheetName => "Продуктовые календари";
        private readonly Microsoft.Office.Interop.Excel.Application Application = Globals.ThisWorkbook.Application;

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

        private Workbook Workbook
        {
            get
            {
                if (_Workbook == null)
                    _Workbook = Application.Workbooks.Open(FileName);
                return _Workbook;
            }
            set
            {
                _Workbook = value;
            }
        }
        private Workbook _Workbook;
        private string FileName;

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
            ProcessBar progressCalendar;

            List<ProductCalendar> calendars = GetProducts();

            progressCalendar = new ProcessBar("Обработка календарей", calendars.Count);
            progressCalendar.Show();

            foreach (ProductCalendar productCalendar in calendars)
            {
                FileName = productCalendar.Path;
                if (!File.Exists(FileName)) continue;
                
                progressCalendar.TaskStart($"Обрабатывается календарь {productCalendar.Name}");
                if (progressCalendar.IsCancel) break;

                List<Product> products = new Product().GetProducts();
                UpdateProductFromCalendar(products);

                //Application.Workbooks(FileName).Close(false);
                Workbook.Close(false);
                
            }
            progressCalendar.Close();
        }

        private void UpdateProductFromCalendar(List<Product> products)
        {
            ProcessBar progressProduct;

            progressProduct = new ProcessBar("Обработка продуктов", products.Count);
            progressProduct.Show();

            foreach (Product product in products)
            {
                progressProduct.TaskStart($"Обрабатывается № {product.Id}");
                if (progressProduct.IsCancel) break;

                product.SetFromCalendar(Workbook);
            }
            progressProduct.Close();

        }

        public void UpdateProductFromCalendar(Product product)
        {
            if (product.Calendar == null) return;
            Model.ProductCalendar productCalendar = new ProductCalendar(product.Calendar);
            
            FileName = productCalendar.Path;
            
            if (!File.Exists(FileName)) return;
            product.SetFromCalendar(Workbook);

            Workbook.Close(false);
        }
    }
}
