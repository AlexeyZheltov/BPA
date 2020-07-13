using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;

namespace BPA.Model {
    /// <summary>
    /// Справочник продукт групп
    /// </summary>
    class ProductStatus : TableBase {
        public override string TableName => "Статусы_товаров";
        public override string SheetName => "Статусы товаров";

        public static Dictionary<string, int> ColDict { get; set; } = new Dictionary<string, int>();

        public override IDictionary<string, string> Filds {
            get {
                return _filds;
            }
        }
        private readonly Dictionary<string, string> _filds = new Dictionary<string, string>
        {
            { "Id", "№" },
            { "Status", "Статус" }
        };

        /// <summary>
        /// Идентификатор
        /// </summary>
        public int Id {
            get; set;
        }
        /// <summary>
        /// Статус
        /// </summary>
        public string Status {
            get; set;
        }

        public ProductStatus() { }

        public ProductStatus(string status)
        {
            ListRow listRow = GetRow(Status, status);
            if (listRow != null)
                SetProperty(listRow);
        }

        //public int GetStatusID(ProductForPlanningNewYear product)
        //{
        //    ListRow row = new ProductStatus().GetRow(Status, product.Status);
        //    ProductStatus status.SetProperty(row);
        //    return ProductStatus()
        //}

    }
}
