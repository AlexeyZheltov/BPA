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
    class RRC : TableBase
    {
        public override string TableName => "РРЦ";
        public override string SheetName => "РРЦ";

        #region --- Словарь ---

        public override IDictionary<string, string> Filds
        {
            get
            {
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

        #endregion

        #region --- Свойства ---

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
        public string Article
        {
            get; set;
        }
        /// <summary>
        /// IRP, Eur
        /// </summary>
        public double IRP
        {
            get; set;
        }

        /// <summary>
        /// РРЦ, руб. с НДС
        /// </summary>
        public double RRCNDS
        {
            get; set;
        }

        /// <summary>
        /// DIY price list, руб. без НДС
        /// </summary>
        public double DIY
        {
            get; set;
        }

        /// <summary>
        /// Дата принятия
        /// </summary>
        DateTime _date;
        public DateTime Date
        {
            get => _date.Date;
            set => _date = value;
        }

        #endregion
        public RRC(Excel.ListRow row) => SetProperty(row);
        public RRC() { }

        //public DateTime? GetDateAsDateTime()
        //{
        //    if (DateTime.TryParse(Date, out DateTime date)) return date.Date;
        //    else return null;
        //}

        public RRC GetRRC(string article, DateTime date)
        {
            ListRow listRow = GetRow("Article", article);

            if (listRow != null)
            {
                RRC currentRRC = new RRC();
                currentRRC.SetProperty(listRow);
                Range firstCell = listRow.Range[1, Table.ListColumns[Filds["Article"]].Index];
                int firstCellRow = firstCell.Row;
                int afterCellRow;

                do
                {
                    RRC tmpRRC = new RRC();
                    tmpRRC.SetProperty(listRow);

                    if (tmpRRC.Date <= date)
                    {
                        currentRRC = tmpRRC.Date > currentRRC.Date ? tmpRRC : currentRRC;
                    }

                    listRow = GetRow("Article", article, firstCell);
                    Range afterCell = listRow.Range[1, Table.ListColumns[Filds["Article"]].Index];
                    firstCell = afterCell;
                    afterCellRow = afterCell.Row;
                }
                while (afterCellRow != firstCellRow);

                return currentRRC;
            }
            return null;
        }

        /// <summary>
        /// поиск в справочнике цен артикула article с указанной датой date
        /// </summary>
        /// <param name="article"></param>
        /// <param name="date"></param>
        /// <returns></returns>
        public RRC GetRRC(string article, DateTime date, bool accurate = true)
        {
            if (accurate != true)
                return GetRRC(article, date);

            ListRow listRow = GetRow("Article", article);

            if (listRow == null)
                return null;

            RRC currentRRC = new RRC();
            currentRRC.SetProperty(listRow);
            Range firstCell = listRow.Range[1, Table.ListColumns[Filds["Article"]].Index];
            int firstCellRow = firstCell.Row;
            int afterCellRow;
            
            do
            {
                RRC tmpRRC = new RRC();
                tmpRRC.SetProperty(listRow);

                if (tmpRRC.Date == date)
                {
                    currentRRC = tmpRRC;
                }

                listRow = GetRow("Article", article, firstCell);
                Range afterCell = listRow.Range[1, Table.ListColumns[Filds["Article"]].Index];
                firstCell = afterCell;
                afterCellRow = afterCell.Row;
            }
            while (afterCellRow != firstCellRow);

            return currentRRC.Date == date ? currentRRC : null;
        }

       /// <summary>
       /// Обновление Справочника из ProductForRRC
       /// </summary>
       /// <param name="product"></param>
        public void UpdatePriceFromProduct(ProductForRRC product)
        {
            if (product != null)
            {
                this.Date = product.DateOfPromotion;
                this.RRCNDS = product.RRCCalculated;
                this.DIY = product.DIY;
                this.Article = product.Article;
                this.IRP = product.IRP;
            }

            Update();
        }

        public static List<RRC> GetAllRRC(PBWrapper pB)
        {
            List<RRC> rrcs = new List<RRC>();
            RRC rrc = new RRC();
            pB.Start(rrc.Table.ListRows.Count);

            foreach (Excel.ListRow row in rrc.Table.ListRows)
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
                                .Where(x => x.Date <= currentDate)
                                .ToList();

                buffer.Sort((x, y) =>
                {
                    if (x.Date > y.Date) return 1;
                    else if (x.Date < y.Date) return -1;
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
