using BPA.Modules;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

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
        public string Article {
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

        public RRC(Excel.ListRow row) => SetProperty(row);
        public RRC() { }

        public RRC GetRRC(string article, string date)
        {
            RRC rrc;
            ListRow listRow = GetRow("Article", article);
            if (listRow != null)
            {
                Range firstCell = listRow.Range[1, Table.ListColumns[Filds["Date"]].Index];
                Range afterCell;
                do
                {
                    rrc = new RRC();
                    rrc.SetProperty(listRow);

                    if (rrc.Date == date)
                    { 
                        return rrc;
                    }
                    afterCell = listRow.Range[1, Table.ListColumns[Filds["Date"]].Index];
                    listRow = GetRow("Article", article, afterCell); 
                }
                while (afterCell != firstCell);
            }
            return null;
        }

        public static List<RRC> GetAllRRC()
        {
            List<RRC> rrcs = new List<RRC>();
            foreach(Excel.ListRow row in new RRC().Table.ListRows)
            {
                rrcs.Add(new RRC(row));
            }
            return rrcs;
        }

    }
}
