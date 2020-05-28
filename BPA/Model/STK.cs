using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BPA.Model {
    /// <summary>
    /// Справочник STK
    /// </summary>
    class STK : TableBase {
        public override string TableName => "STK";
        public override string SheetName => "STK";

        public override IDictionary<string, string> Filds {
            get {
                return _filds;
            }
        }
        private readonly Dictionary<string, string> _filds = new Dictionary<string, string>
        {
            { "Id","№" },
            { "Article", "Артикул" },
            { "STKEur", "STK 2.5, Eur" },
            { "STKRub", "STK 2.5, руб." },
            { "Date", "Дата принятия" }
        };

        /// <summary>
        /// №
        /// </summary>
        public int Id
        {
            get; set;
        }

        /// <summary>
        /// Артикул
        /// </summary>
        public int Article {
            get; set;
        }
        /// <summary>
        /// STK 2.5, Eur
        /// </summary>
        public string STKEur {
            get; set;
        }

        /// <summary>
        /// STK 2.5, руб.
        /// </summary>
        public string STKRub {
            get; set;
        }
        /// <summary>
        /// Дата принятия
        /// </summary>
        public string Date {
            get; set;
        }

    }
}
