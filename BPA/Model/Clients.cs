using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BPA.Model {
    /// <summary>
    /// Справочник клиентов
    /// </summary>
    class Clients : TableBase {
        public override string TableName => "Клиенты";
        public override string SheetName => "Клиенты";

        public override IDictionary<string, string> Filds {
            get {
                return _filds;
            }
        }
        private readonly Dictionary<string, string> _filds = new Dictionary<string, string>
        {
            { "Id", "№" },
            { "GardenaChannel", "GardenaChannel" },
            { "Customer", "Customer" },
            { "ChannelType", "Channel type" },
            { "CustomerStatus", "Customer status" },
            { "SalesManager", "Sales manager" },
            { "Mag", "Маг" },
            { "CustomerBudget", "CustomerBudget" }
        };


        /// <summary>
        /// №
        /// </summary>
        public int Id
        {
            get; set;
        }

        /// <summary>
        /// GardenaChannel
        /// </summary>
        public string GardenaChannel {
            get; set;
        }
        /// <summary>
        /// Customer
        /// </summary>
        public string Customer {
            get; set;
        }

        /// <summary>
        /// Channel type
        /// </summary>
        public string ChannelType {
            get; set;
        }
        /// <summary>
        /// Customer Status
        /// </summary>
        public string CustomerStatus {
            get; set;
        }

        /// <summary>
        /// Sales manager
        /// </summary>
        public string SalesManager {
            get; set;
        }

        /// <summary>
        /// Маг
        /// </summary>
        public string Mag {
            get; set;
        }

        /// <summary>
        /// CustomerBudget
        /// </summary>
        public string CustomerBudget {
            get; set;
        }

    }
}
