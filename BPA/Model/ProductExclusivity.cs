using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BPA.Model {
    /// <summary>
    /// Справочник эксклюзивности товаров
    /// </summary>
    class ProductExclusivity : TableBase {
        public override string TableName => "Эксклюзивность";
        public override string SheetName => "Эксклюзивность";

        public static Dictionary<string, int> ColDict { get; set; } = new Dictionary<string, int>();

        public override IDictionary<string, string> Filds {
            get {
                return _filds;
            }
        }
        private readonly Dictionary<string, string> _filds = new Dictionary<string, string>
        {
            { "Id", "№" },
            { "Exclusivity", "Эксклюзивность" }
        };

        /// <summary>
        /// Идентификатор
        /// </summary>
        public int Id {
            get; set;
        }
        /// <summary>
        /// Эксклюзивность
        /// </summary>
        public string Exclusivity {
            get; set;
        }
        
    }
}
