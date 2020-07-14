using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BPA.Model {
    /// <summary>
    /// Справочник DIY
    /// </summary>
    class DIY : TableBase {
        public override string TableName => "DIY";
        public override string SheetName => "DIY";
        public static Dictionary<string, int> ColDict { get; set; } = new Dictionary<string, int>();

        public override IDictionary<string, string> Filds {
            get {
                return _filds;
            }
        }
        private readonly Dictionary<string, string> _filds = new Dictionary<string, string>
        {
            { "Id", "№" },
            { "SubGroup", "SubGroup" },
            { "RRCPercent", "Процент от РРЦ, без НДС" }
        };

        /// <summary>
        /// №
        /// </summary>
        public int Id {
            get; set;
        }
        /// <summary>
        /// SubGroup
        /// </summary>
        public string SubGroup {
            get; set;
        }

        /// <summary>
        /// Процент от РРЦ, без НДС
        /// </summary>
        public string RRCPercent {
            get; set;
        }

    }
}
