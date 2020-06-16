using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BPA.Model {
    /// <summary>
    /// Справочник РРЦ
    /// </summary>
    class RRC : TableBase
    {
        public override string TableName => "РРЦ";
        public override string SheetName => "РРЦ";

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
        public string IRP
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
        public DateTime Date
        {
            get; set;
        }

        public RRC()
        {
        }

        /// <summary>
        /// поиск в справочнике цен артикула article с датой date не познее указанной
        /// </summary>
        /// <param name="article"></param>
        /// <param name="date"></param>
        /// <returns></returns>
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

        public void UpdatePriceFromProduct(Product product)
        {
            if (product != null)
            {
                this.Date = product.DateOfPromotion;
                this.RRCNDS = product.RRCCalculated;
                this.DIY = product.DIY;
                this.Article = product.Article;
            }

            Update();
        }
    }
}
