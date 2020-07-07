﻿using Microsoft.Office.Interop.Excel;
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
            //try
            //{
            ThisWorkbook workbook = Globals.ThisWorkbook;
            ListObject table = workbook.Sheets[SheetName].ListObjects[1];
            return table.Name;
            //}
            //catch (Exception ex)
            //{
            //    MessageBox.Show(ex.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            //}
        }

        public string templateSheetName = Properties.Settings.Default.templateSheetName;
        private const string CustomerStatusLabel = "Customer status";
        private const string ChannelTypeLabel = "Channel type";
        private const string YearLabel = "Период";
        private const string MaximumBonusLabel = "максмальный годовой бонус, %";

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
            { "DIYPriceList","DIY price list, руб. без НДС"},

            { "PricePromo","Промо цена, руб."},
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
        public double STKEur
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
        /// Промо цена, руб.
        /// </summary>
        public string PricePromo
        {
            get; set;
        }


        #endregion

        public string ChanelType;
        public string CustomerStatus;
        public double Year;
        public double MaximumBonus;

        /// <summary>
        /// словарь соответствия key: Эксклюзивность, val: CustomerStatus, ChannalType
        /// </summary>
        public readonly Dictionary<string, string> ExclusivesDict = new Dictionary<string, string>
        {
            {"леруа мерлен", "леруа мерлен"},
            {"оби", "оби"},
            {"diy канал", "diy"},
            {"dealer", "dealers&regional distr"},
            {"regional", "dealers&regional distr"},
            {"online", "online"},
            {"all channels", ""}
        };

        public PlanningNewYear() { }
        public PlanningNewYear(ListRow row) => SetProperty(row);

        public PlanningNewYear Clone()
        {
            PlanningNewYear planning = new PlanningNewYear();

            planning.Year = this.Year;
            planning.CustomerStatus = this.CustomerStatus;
            planning.ChanelType = this.ChanelType;
            planning.MaximumBonus = this.MaximumBonus;

            return planning;
        }

        public void SetProduct(ProductForPlanningNewYear product)
        {
            this.Article = product.Article;
            this.RRCNDS = product.RRCFinal; //?
            //planning.PercentageOfChange = product.RRCPercent;  //?
            this.IRP = product.IRP;
            //planning.RRCNDS2 = product.RRCFinal; //?
            this.IRPIndex = product.IRPIndex;
            this.DIYPriceList = product.DIY;
        }
        public void GetSTK()
        {
            if (this.Article == "")
                return;

            STK sTK = new STK().GetSTK(this.Article, this.Year);
            if (sTK == null)
                return;

            this.STKEur = sTK.STKEur;
            this.STKRub = sTK.STKRub;
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

        public void ClearTable(string worksheetName)
        {
            _TableWorksheetName = worksheetName;

            ClearTable();

            CopyTemplateFirstRow();
        }

        private void CopyTemplateFirstRow()
        {
            ThisWorkbook workbook = Globals.ThisWorkbook;
            Worksheet worksheet = workbook.Sheets[templateSheetName];
            ListObject tableTemplate = worksheet.ListObjects[1];


            Table.ListRows.AddEx();
            Range firstCell = Table.ListRows[1].Range[1];

            tableTemplate.ListRows[1].Range.Copy();
            firstCell.PasteSpecial();

        }

        public void SetMaximumBonusValue()
        {
            ThisWorkbook workbook = Globals.ThisWorkbook;
            Worksheet worksheet = workbook.Sheets[SheetName];
            Range range = worksheet.UsedRange;
            
            Range cell;
            cell = range.Find(What: "%", LookIn: XlFindLookIn.xlValues, LookAt: XlLookAt.xlPart, MatchCase: false);
            cell = range.Find(What:MaximumBonusLabel,LookIn:XlFindLookIn.xlFormulas, LookAt:XlLookAt.xlPart, MatchCase:false);
            if (cell == null)
                return;

            cell.Offset[0, -2].Value = this.MaximumBonus;
        }

    }
}