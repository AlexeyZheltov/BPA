using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BPA.Model {
    /// <summary>
    /// Справочник бюджетных курсов
    /// </summary>
    class BugetCourse : TableBase {
        public override string TableName => "Бюджетные_курсы";
        public override string SheetName => "Бюджетные курсы";

        public override IDictionary<string, string> Filds {
            get {
                return _filds;
            }
        }
        private readonly Dictionary<string, string> _filds = new Dictionary<string, string>
        {
            { "Id", "№" },
            { "Date", "Дата принятия" },
            { "RRCPercent", "Процент от РРЦ, без НДС" }
        };

        /// <summary>
        /// №
        /// </summary>
        public int Id {
            get; set;
        }
   
        /// <summary>
        /// Дата принятия
        /// </summary>
        public string Date {
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
