using Microsoft.Office.Interop.Excel;
using System.Collections.Generic;
using BPA.Modules;

namespace BPA.Model
{
    /// <summary>
    /// Планирование нового года шаблон
    /// </summary>
    class PlanningNewYear : TableBase
    {
        public override string TableName => "Планирование_новый_год";
        public override string SheetName => "Планирование нового года шаблон";

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
            { "Article","Артикул"},
            { "RRCNDS","РРЦ, руб.с НДС"},
            { "PercentageOfChange","Процент изменения"},
            { "STKEur","STK 2.5,  Eur"},
            { "STKRub","STK 2.5, руб."},
            { "IRP","IRP, Eur"},
            { "RRCNDS2","РРЦ, руб.с НДС2"},
            { "IRPIndex","Индекс IRP"},
            { "DIYPriceList","DIY price list, руб. без НДС"}
        };

        #endregion

        #region --- Свойства ---
        /// <summary>
        /// Артикул
        /// </summary>
        public string Article
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
        /// Процент изменения
        /// </summary>
        public double PercentageOfChange
        {
            get; set;
        }

        /// <summary>
        /// STK 2.5, руб.
        /// </summary>
        public double STKRub
        {
            get; set;
        }
        
        /// <summary>
        /// STK 2.5, Eur
        /// </summary>
        public string STKEur
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
        public double RRCNDS2
        {
            get; set;
        }

        /// <summary>
        /// Индекс IRP
        /// </summary>
        public double IRPIndex
        {
            get; set;
        }

        /// <summary>
        /// DIY price list, руб. без НДС
        /// </summary>
        public double DIYPriceList
        {
            get; set;
        }
        /// <summary>
        /// Дата принятия
        /// </summary>

        #endregion
        public PlanningNewYear() { }
        public PlanningNewYear(ListRow row) => SetProperty(row);
        public PlanningNewYear(ProductForPlanningNewYear product)
        {
            this.Article = product.Article;
            this.RRCNDS = product.RRCFinal; //?
            this.PercentageOfChange = product.RRCPercent;  //?
            //            this.STKEur = product.st
            //            this.STKRub = 
            this.IRP = product.IRP;
            this.RRCNDS2 = product.RRCFinal; //?
            this.IRPIndex = product.IRPIndex;
            this.DIYPriceList = product.DIY;
        }

        public void GetSheetCopy()
        {
            ThisWorkbook workbook = Globals.ThisWorkbook;
            FunctionsForExcel.ShowSheet(SheetName);

            string newSheetName = SheetName.Replace("шаблон", "").Trim();
            Worksheet newSheet = FunctionsForExcel.CreateSheetCopy(workbook.Sheets[SheetName], newSheetName);
            newSheet.Activate();

            FunctionsForExcel.HideSheet(SheetName);
        }
    }
}
