using BPA.Forms;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualStudio.Tools.Applications.Deployment;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using BPA.Modules;

namespace BPA.Model {
    /// <summary>
    /// Справочник клиентов
    /// </summary>
    class Client : TableBase {
        //public static ComparerCustomer ComparerCustomer => new ComparerCustomer();
        private readonly Microsoft.Office.Interop.Excel.Application Application = Globals.ThisWorkbook.Application;

        public override string TableName => "Клиенты";
        public override string SheetName => "Клиенты";

        public static Dictionary<string, int> ColDict { get; set; } = new Dictionary<string, int>();

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
        /// Нужно описать конструктор!!!!
        /// </summary>
        /// <param name="planning"></param>
        public Client(PlanningNewYear planning)
        {
            CustomerStatus = planning.CustomerStatus;
            ChannelType = planning.ChannelType;
        }

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

        public static Client GetCurrentClient()
        {
            Client client = new Client();
            Range activeCell = client.Application.ActiveCell;
            Worksheet activeSheet = client.Application.ActiveSheet;
            if (activeSheet.Name != client.SheetName || activeCell.Row < client.FirstRow || activeCell.Row > client.LastRow)
                return null;

            client.FillFromRow(client.Table.ListRows[activeCell.Row - client.FirstRow + 1]);
            return client;
        }

        public static List<Client> GetAllClients()
        {
            List<Client> clients = new List<Client>();
            Excel.ListObject table = new Client().Table;
            ProcessBar processBar = new ProcessBar("Обновление листа клиентов", table.ListRows.Count);
            bool isCancel = false;

            void Cancel() => isCancel = true;

            processBar.CancelClick += Cancel;
            //processBar.TaskStart("Чтение клиентов");
            processBar.Show(new ExcelWindows(Globals.ThisWorkbook));

            foreach (Excel.ListRow row in new Client().Table.ListRows)
            {
                if (isCancel) return null;
                processBar.TaskStart($"Обрабатывается строка {row.Index}");
                clients.Add(new Client(row));
                processBar.TaskDone(1);
            }
            processBar?.Close();

            return clients;
        }
        public List<Client> GetCustomers(string customerStatus, string channelType)
        {
            List<Client> clients = new List<Client>();

            ProcessBar processBar = new ProcessBar("Поиск клиентов", Table.ListRows.Count);
            bool isCancel = false;
            void Cancel() => isCancel = true;
            processBar.CancelClick += Cancel;
            processBar.Show(new ExcelWindows(Globals.ThisWorkbook));

            foreach (Excel.ListRow row in Table.ListRows)
            {
                if (isCancel) return null;
                processBar.TaskStart($"Обрабатывается строка {row.Index}");

                Client client = new Client(row);

                if (customerStatus == client.CustomerStatus && channelType == client.ChannelType)
                {
                    clients.Add(client);
                }

                processBar.TaskDone(1);
            }
            processBar?.Close();

            return clients;
        }
    }
}
