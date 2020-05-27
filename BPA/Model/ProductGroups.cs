using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BPA.Model {
    /// <summary>
    /// Справочник продукт групп
    /// </summary>
    class ProductGroups : TableBase {
        public override string TableName => "Продукт группы";
        public override string SheetName => "Продукт_Группы";

        public override IDictionary<string, string> Filds {
            get {
                return _filds;
            }
        }
        private readonly Dictionary<string, string> _filds = new Dictionary<string, string>
        {
            { "Id", "№" },
            { "ProductGroup", "Продукт группа" },
            { "ProductGroupEng", "Название продукт группы (ENG)" },
            { "ProductGroupRu", "Название продукт группы (RUS)" }
        };

        /// <summary>
        /// Идентификатор
        /// </summary>
        public int Id {
            get; set;
        }
        /// <summary>
        /// Продукт группа
        /// </summary>
        public string ProductGroup{
            get; set;
        }

        /// <summary>
        /// Название продукт группы (ENG)
        /// </summary>
        public string ProductGroupEng {
            get; set;
        }
        /// <summary>
        /// Название продукт группы (RUS)
        /// </summary>
        public string ProductGroupRu {
            get; set;
        }

    }
}
