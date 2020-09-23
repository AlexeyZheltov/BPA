using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BPA.NewModel
{
    class ClientItem
    {
        TableRow _row;
        public ClientItem(TableRow row) => _row = row;

        #region Свойства таблицы
        public int Id
        {
            get => _row["№"];
            set => _row["№"] = value;
        }
        public string GardenaChannel
        {
            get => _row["GardenaChannel"];
            set => _row["GardenaChannel"] = value;
        }
        public string Customer
        {
            get => _row["Customer"];
            set => _row["Customer"] = value;
        }
        public string ChannelType
        {
            get => _row["Channel type"];
            set => _row["Channel type"] = value;
        }
        public string CustomerStatus
        {
            get => _row["Customer status"];
            set => _row["Customer status"] = value;
        }
        public string CustomerStatusForecast
        {
            get => _row["Customer status for forecast"];
            set => _row["Customer status for forecast"] = value;
        }
        public string SalesManager
        {
            get => _row["Sales manager"];
            set => _row["Sales manager"] = value;
        }
        public string Mag
        {
            get => _row["Маг"];
            set => _row["Маг"] = value;
        }
        public string CustomerBudget
        {
            get => _row["CustomerBudget"];
            set => _row["CustomerBudget"] = value;
        }
        #endregion

        public struct DataFromDescision
        {
            public string Customer { get; set; }
            public string GardenaChannel { get; set; }
        }
    }
}
