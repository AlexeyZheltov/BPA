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

        public DateTime? GetDateAsDateTime()
        {
            if (DateTime.TryParse(Date, out DateTime date)) return date.Date;
            else return null;
        }

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

        public static List<RRC> GetAllRRC(PBWrapper pB)
        {
            List<RRC> rrcs = new List<RRC>();
            RRC rrc = new RRC();
            pB.Start(rrc.Table.ListRows.Count);

            foreach(Excel.ListRow row in rrc.Table.ListRows)
            {
                if (pB.IsCancel)
                {
                    pB.Dispose();
                    return null;
                }
                pB.Action($"{row.Index}");
                rrcs.Add(new RRC(row));
                pB.Done(1);
            }
            pB.Dispose();
            return rrcs;
        }

        public static List<RRC> GetActualPriceList(DateTime currentDate)
        {
            PBWrapper pb = new PBWrapper($"Создание прайс-листа", $"Анализ артикулов с листа РРЦ [Index]");

            //подключится к ценам
            List<RRC> rrcs = RRC.GetAllRRC(new PBWrapper($"Создание прайс-листа", "Чтение РРЦ [Index]"));
            if (rrcs == null) return null;
            //список уникальных артикулов
            List<string> arts = (from rrc in rrcs
                                 select rrc.Article).Distinct().ToList();

            List<RRC> actualRRC = new List<RRC>();
            List<RRC> buffer = new List<RRC>();

            pb.Start(arts.Count);
            //взять пачку строк соответсвующих артикулу и вязть тот что с последней датой
            foreach (string art in arts)
            {
                if (pb.IsCancel)
                {
                    pb.Dispose();
                    return null;
                }
                pb.Action(art);
                buffer = rrcs.FindAll(x => x.Article == art)
                                .Where(x => x.GetDateAsDateTime() <= currentDate)
                                .ToList();

                buffer.Sort((x, y) =>
                {
                    if (x.GetDateAsDateTime() > y.GetDateAsDateTime()) return 1;
                    else if (x.GetDateAsDateTime() < y.GetDateAsDateTime()) return -1;
                    else return 0;
                });

                if (buffer.Count == 0) continue;
                actualRRC.Add(buffer[0]);
                pb.Done(1);
            }
            pb.Dispose();

            return actualRRC;
        }

    }
}
