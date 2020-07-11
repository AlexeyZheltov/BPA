using BPA.Forms;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BPA.Model
{
    /// <summary>
    /// Справочник товаров для РРЦ
    /// </summary>
    class ProductForRRC : TableBase
    {
        private readonly Microsoft.Office.Interop.Excel.Application Application = Globals.ThisWorkbook.Application;

        public override string TableName => "Товары";
        public override string SheetName => "Товары";

        #region --- Словарь ---

        public override IDictionary<string, string> Filds => _filds;
        private readonly Dictionary<string, string> _filds = new Dictionary<string, string>
        {
            { "Id","№" },
            { "Article","Артикул" },
            { "IRP","IRP, Eur" },
            { "RRCCurrent","РРЦ текущий" },
            { "DIYCurrent","DIY текущий" },
            { "RRCCalculated","РРЦ расчетная, руб." },
            { "DIY","DIY price list, руб. без НДС" }
        };

        #endregion

        #region -- Основные свойства столбцов для РРЦ ---

        /// <summary>
        /// №
        /// </summary>
        public int Id
        {
            get; set;
        }
        /// <summary>
        /// Категория для прайс-листа диллеров
        /// </summary>
        public string Article
        {
            get; set;
        }
        /// <summary>
        /// IRP, Eur
        /// </summary>
        public Double IRP
        {
            get; set;
        }

        /// <summary>
        /// РРЦ текущий
        /// </summary>
        public Double RRCCurrent
        {
            get; set;
        }

        /// <summary>
        /// DIY текущий
        /// </summary>
        public Double DIYCurrent
        {
            get; set;
        }

        /// <summary>
        /// РРЦ расчетная, руб.
        /// </summary>
        public Double RRCCalculated
        {
            get; set;
        }

        /// <summary>
        /// DIY price list, руб. без НДС
        /// </summary>
        public Double DIY
        {
            get; set;
        }
        #endregion

        public ProductForRRC() { }
        public ProductForRRC(ListRow listRow) => SetProperty(listRow);

        /// <summary>
        /// Дата повышения
        /// </summary>
        /// <returns></returns>
        public DateTime DateOfPromotion
        {
            get
            {
                string dateOfPromotionLabel = "Дата повышения";

                try
                {
                    Range dateCell = Table.DataBodyRange.Cells[1, 1].Parent.UsedRange.Find(dateOfPromotionLabel, LookAt: XlLookAt.xlWhole);
                    return DateTime.Parse(dateCell.Offset[0, 1].Text);
                }
                catch
                {
                    return new DateTime();
                }
            }

            set
            {
                _DateOfPromotion = value;
            }
        }
        private DateTime _DateOfPromotion;

        /// <summary>
        /// Получение списка продуктов для РРЦ
        /// </summary>
        public List<ProductForRRC> GetProducts()
        {
            bool isCancel = false;
            void CancelLocal() => isCancel = true;
            ProcessBar processBar = new ProcessBar("Получение списка продуктов", LastRow);
            processBar.CancelClick += CancelLocal;
            processBar.Show();

            List<ProductForRRC> products = new List<ProductForRRC>();
            foreach (ListRow row in Table.ListRows)
            {
                if (isCancel)
                {
                    processBar.Close();
                    return null;
                }
                processBar.TaskStart($"Обрабатывается товар {row.Index} из {LastRow - Table.HeaderRowRange.Row}");

                ProductForRRC product = new ProductForRRC(row);
                if ((int)product.Id !=0)
                    products.Add(product);

                processBar.TaskDone(1);

            }
            processBar.Close();
            return products;
        }

        public void UpdatePriceFromRRC(RRC rrc)
        {
            if (rrc != null)
            {
                this.RRCCurrent = rrc.RRCNDS;
                this.DIYCurrent = rrc.DIY;
                this.IRP = rrc.IRP;

                Update();
            }
        }

        //private bool IsCancel = false;
        /// <summary>
        /// Событие начала задачи
        /// </summary>
        //public event ActionsStart ActionStart;
        //public delegate void ActionsStart(string name);

        /// <summary>
        /// Событие завершения задачи
        /// </summary>
        //public event ActionsDone ActionDone;
        //public delegate void ActionsDone(int count);

        //public void Cancel()
        //{
        //    IsCancel = true;
        //}

    }
}
