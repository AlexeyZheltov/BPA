using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BPA.Model {
    /// <summary>
    /// Справочник РРЦ
    /// </summary>
    class RRC : TableBase {
        public override string TableName => "РРЦ";
        public override string SheetName => "РРЦ";

        public override IDictionary<string, string> Filds {
            get {
                return _filds;
            }
        }
        private readonly Dictionary<string, string> _filds = new Dictionary<string, string>
        {
            { "Id","№" },
            { "Article", "Артикул" },
            { "IRP", "IRP, Eur" },
            { "RRCNDS", "РРЦ, руб. с НДС" },
            { "DIY", "DIY price list, руб. без НДС" },
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
        /// IRP, Eur
        /// </summary>
        public string IRP {
            get; set;
        }

        /// <summary>
        /// РРЦ, руб. с НДС
        /// </summary>
        public string RRCNDS {
            get; set;
        }

        /// <summary>
        /// DIY price list, руб. без НДС
        /// </summary>
        public string DIY {
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
