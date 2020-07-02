using Microsoft.Office.Interop.Excel;
using System.Collections.Generic;
using BPA.Modules;
using System;
using BPA.Forms;

namespace BPA.Model
{
    /// <summary>
    /// Планирование нового года шаблон
    /// </summary>
    class PlanningNewYear : TableBase
    {
        //public override string TableName => "Планирование_новый_год";
        //public override string SheetName => "Планирование нового года шаблон";
        public override string TableName => GetTableName();
        public override string SheetName => _TableWorksheetName != "" ? _TableWorksheetName: templateSheetName;
        public string _TableWorksheetName;

        public string GetTableName()
        {
            try
            {
                ThisWorkbook workbook = Globals.ThisWorkbook;
                ListObject table = workbook.Sheets[SheetName].ListObjects[1];
                return table.Name;
            }
            catch
            {
                ThisWorkbook workbook = Globals.ThisWorkbook;
                ListObject table = workbook.Sheets[templateSheetName].ListObjects[1];
                return table.Name;
            }
        }

        public string templateSheetName = "Планирование нового года шаблон";
        private string CustomerStatusLabel = "Customer status";
        private string ChannelTypeLabel = "Channel type";
        private string YearLabel = "Период";

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
        
        public string ChanelType;
        public string CustomerStatus;
        public double Year;

        public PlanningNewYear() { }
        public PlanningNewYear(ListRow row) => SetProperty(row);

        public PlanningNewYear Clone(ProductForPlanningNewYear product)
        {
            PlanningNewYear planning = new PlanningNewYear();

            planning.Year = this.Year;

            planning.Article = product.Article;
            planning.RRCNDS = product.RRCFinal; //?
            planning.PercentageOfChange = product.RRCPercent;  //?
            //            planning.STKEur = product.st
            //            planning.STKRub = 
            planning.IRP = product.IRP;
            planning.RRCNDS2 = product.RRCFinal; //?
            planning.IRPIndex = product.IRPIndex;
            planning.DIYPriceList = product.DIY;

            return planning;
        }

        public void GetSheetCopy()
        {
            ThisWorkbook workbook = Globals.ThisWorkbook;
            FunctionsForExcel.ShowSheet(templateSheetName);

            string newSheetName = templateSheetName.Replace("шаблон", "").Trim();
            Worksheet newSheet = FunctionsForExcel.CreateSheetCopy(workbook.Sheets[templateSheetName], newSheetName);
            newSheet.Activate();

            FunctionsForExcel.HideSheet(templateSheetName);
        }

        public PlanningNewYear GetTmp(string worksheetName)
        {
            if (worksheetName == templateSheetName)
                return null;
            else
                _TableWorksheetName = worksheetName;

            try
            {
                PlanningNewYear planningNewYear = new PlanningNewYear();
                ThisWorkbook workbook = Globals.ThisWorkbook;
                Range rng = workbook.Sheets[SheetName].UsedRange;

                planningNewYear.CustomerStatus = val(CustomerStatusLabel);
                planningNewYear.ChanelType = val(ChannelTypeLabel);
                if (double.TryParse(val(YearLabel), out double year))
                    planningNewYear.Year = year;

                string val(string label)
                {
                    try
                    {
                        Range cell =rng.Find(label, LookAt: XlLookAt.xlWhole);
                        return cell.Offset[0, 1].Text;
                    }
                    catch
                    {
                        return "";
                    }
                }

                return planningNewYear;
            } catch
            {
                return null;
            }
        }

        public void Save(string worksheetName)
        {
            _TableWorksheetName = worksheetName;
            Save();
        }
    }
}
