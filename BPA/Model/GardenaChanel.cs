using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BPA.Model {
    /// <summary>
    /// Справочник клиентов
    /// </summary>
    class GardenaChannel : TableBase {
        public override string TableName => "Клиенты";
        public override string SheetName => "Клиенты";

        public override IDictionary<string, string> Filds {
            get {
                return _filds;
            }
        }
        private readonly Dictionary<string, string> _filds = new Dictionary<string, string>
        {
            { "Id", "Код" },
            { "GardenaChannel", "Gardena_Channel" },
            { "SalesManager", "Sales manager" }
        };

        /// <summary>
        /// Код
        /// </summary>
        public int Id {
            get; set;
        }

        /// <summary>
        /// GardenaChannel
        /// </summary>
        public int Gardena_Channel {
            get; set;
        }

        /// <summary>
        /// Sales manager
        /// </summary>
        public string SalesManager {
            get; set;
        }
    }
}
