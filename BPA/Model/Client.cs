using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace BPA.Model {
    /// <summary>
    /// Справочник клиентов
    /// </summary>
    class Client : TableBase {
        //public static ComparerCustomer ComparerCustomer => new ComparerCustomer();
        private readonly Microsoft.Office.Interop.Excel.Application Application = Globals.ThisWorkbook.Application;

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

        public Client() { }

        public Client(Excel.ListRow row) => SetProperty(row);
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

        public class ComparerCustomer : IEqualityComparer<Client>
        {
            public bool Equals(Client x, Client y) => x.Customer == y.Customer;

            public int GetHashCode(Client obj) => obj?.Customer.GetHashCode() ?? 0;
        }

        //public Client GetCurrentClients()
        //{
        //    Range activeCell = Application.ActiveCell;
        //    Worksheet activeSheet = Application.ActiveSheet;
        //    if (activeSheet.Name != SheetName || activeCell.Row < FirstRow || activeCell.Row > LastRow) 
        //        return null;
            
        //    Client clients = new Client(Table.ListRows[activeCell.Row - FirstRow + 1]);
        //    return clients;
        //}

        public void FillFromRow(Excel.ListRow row) => SetProperty(row);

        public static Client GetCurrentClients()
        {
            Client client = new Client();
            Range activeCell = client.Application.ActiveCell;
            Worksheet activeSheet = client.Application.ActiveSheet;
            if (activeSheet.Name != client.SheetName || activeCell.Row < client.FirstRow || activeCell.Row > client.LastRow)
                return null;

            client.FillFromRow(client.Table.ListRows[activeCell.Row - client.FirstRow + 1]);
            return client;
        }

    }
}
