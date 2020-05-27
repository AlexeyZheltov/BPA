using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BPA.Model {
    /// <summary>
    /// Справочник продукт групп
    /// </summary>
    class ProductStatus : TableBase {
        public override string TableName => "Статусы_товаров";
        public override string SheetName => "Статусы товаров";

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

    }
}
