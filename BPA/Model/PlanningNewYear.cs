using Microsoft.Office.Interop.Excel;
using System.Collections.Generic;
using BPA.Modules;
using System;
using BPA.Forms;
using SettingsBPA = BPA.Properties.Settings;

namespace BPA.Model
{
    /// <summary>
    /// Планирование нового года шаблон
    /// </summary>
    class PlanningNewYear : TableBase
    {
        public override string TableName => GetTableName();
        public override string SheetName => _TableWorksheetName != null ? _TableWorksheetName: templateSheetName;
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

        public readonly string templateSheetName = SettingsBPA.Default.SHEET_NAME_PLANNING_TEMPLATE;
        private const string CustomerStatusLabel = "Customer status";
        private const string ChannelTypeLabel = "Channel type";
        private const string YearLabel = "Период";
        private const string MaximumBonusLabel = "максмальный годовой бонус, %";
        public static Dictionary<string, int> ColDict { get; set; } = new Dictionary<string, int>();
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
            //{ "Article","Артикул"},
            //{ "RRCNDS","РРЦ, руб.с НДС"},
            //{ "PercentageOfChange","Процент изменения"},
            //{ "STKEur","STK 2.5,  Eur"},
            //{ "STKRub","STK 2.5, руб."},
            //{ "IRP","IRP, Eur"},
            //{ "RRCNDS2","РРЦ, руб.с НДС2"},
            //{ "IRPIndex","Индекс IRP"},
            //{ "DIYPriceList","DIY price list, руб. без НДС"},

            //{ "PricePromo","Промо цена, руб."},

            //добавил новые
            { "SupercategoryEng", "Суперкатегория" },
            { "Supercategory", "SuperCategory" },
            { "ProductGroup", "Product group" },
            { "ProductGroupEng", "Product group name" },
            { "SubGroup", "Subgroup" },
            { "GenericName", "Generic Name (long)" },
            { "PNS", "PNS" },
            { "Article", "Article" },
            { "ArticleOld", "Predessor - Local ID Gardena" },
            { "ArticleRu", "Description RUS" },
            { "CalendarSalesStartDate", "Sales Start Date" },
            { "CalendarPreliminaryEliminationDate", "Preliminary Elimination Date" },
            { "CalendarEliminationDate", "Elimination Date" },
            { "Status", "status" }
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

        ///// <summary>
        ///// Артикул
        ///// </summary>
        //public string Article
        //{
        //    get; set;
        //}

        ///// <summary>
        ///// РРЦ, руб. с НДС
        ///// </summary>
        //public double RRCNDS
        //{
        //    get; set;
        //}

        ///// <summary>
        ///// Процент изменения
        ///// </summary>
        //public double PercentageOfChange
        //{
        //    get; set;
        //}

        ///// <summary>
        ///// STK 2.5, руб.
        ///// </summary>
        //public double STKRub
        //{
        //    get; set;
        //}
        
        ///// <summary>
        ///// STK 2.5, Eur
        ///// </summary>
        //public double STKEur
        //{
        //    get; set;
        //}
        ///// <summary>
        ///// IRP, Eur
        ///// </summary>
        //public double IRP
        //{
        //    get; set;
        //}

        ///// <summary>
        ///// РРЦ, руб. с НДС
        ///// </summary>
        //public double RRCNDS2
        //{
        //    get; set;
        //}

        ///// <summary>
        ///// Индекс IRP
        ///// </summary>
        //public double IRPIndex
        //{
        //    get; set;
        //}

        ///// <summary>
        ///// DIY price list, руб. без НДС
        ///// </summary>
        //public double DIYPriceList
        //{
        //    get; set;
        //}
        ///// <summary>
        ///// Промо цена, руб.
        ///// </summary>
        //public string PricePromo
        //{
        //    get; set;
        //}

        /// <summary>
        /// Суперкатегория
        /// </summary>
        public string SupercategoryEng
        {
            get; set;
        }

        /// <summary>
        /// SuperCategory
        /// </summary>
        public string Supercategory
        {
            get; set;
        }

        ///<summary>
        /// Product group
        /// </summary>
        public string ProductGroup 
        { 
            get; set;
        }

        /// <summary>
        /// Product group name
        /// </summary>
        public string ProductGroupEng 
        {
            get; set;
        }
        /// <summary>
        /// Subgroup
        /// </summary>
        public string SubGroup
        {
            get; set;
        }

        /// <summary>
        /// Generic Name (long)
        /// </summary>
        public string GenericName
        {
            get; set;
        }

        /// <summary>
        /// PNS
        /// </summary>
        public string PNS
        {
            get; set;
        }

        /// <summary>
        /// Article
        /// </summary>
        public string Article
        {
            get; set;
        }

        /// <summary>
        /// Predessor - Local ID Gardena
        /// </summary>
        public string ArticleOld
        {
            get; set;
        }

        /// <summary>
        /// Description RUS
        /// </summary>
        public string ArticleRu
        {
            get; set;
        }

        /// <summary>
        /// Sales Start Date
        /// </summary>
        public Double CalendarSalesStartDate
        {
            get; set;
        }

        /// <summary>
        /// Preliminary Elimination Date
        /// </summary>
        public Double CalendarPreliminaryEliminationDate
        {
            get; set;
        }

        /// <summary>
        /// Elimination Date
        /// </summary>
        public Double CalendarEliminationDate
        {
            get; set;
        }
        /// <summary>
        /// status
        /// </summary>
        public string Status
        {
            get; set;
        }

        #endregion

        public string ChannelType;
        public string CustomerStatus;
        public int Year;
        public DateTime planningDate;
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
        public PlanningNewYear(string worksheetName)
        {
            if (worksheetName == templateSheetName)
                return;
            else
                _TableWorksheetName = worksheetName;
        }
        public PlanningNewYear(ListRow row)
        {
            _TableWorksheetName = row.Range.Cells[1, 1].Parent.Name;
            SetProperty(row);
        }

        public PlanningNewYear Clone()
        {
            PlanningNewYear planning = new PlanningNewYear();

            planning._TableWorksheetName = this.SheetName;
            planning.Year = this.Year;
            planning.CustomerStatus = this.CustomerStatus;
            planning.ChannelType = this.ChannelType;
            planning.MaximumBonus = this.MaximumBonus;

            return planning;
        }

        public DateTime CurrentDate = DateTime.Now;
        private int CurrentMonth => CurrentDate.Month;

        public bool HasData()
        {
            if (Table.ListRows.Count < 1)
                return false;
            
            Range cell= Table.ListRows[1].Range[1, Table.ListColumns[Filds["Id"]].Index];
            if (cell.Value == 0 || cell.Value == null)
                return false;

            return true;
        }
        public void SetProduct(ProductForPlanningNewYear product)
        {
            this.Article = product.Article;
            //this.RRCNDS = product.RRCFinal; //?
            ////planning.PercentageOfChange = product.RRCPercent;  //?
            //this.IRP = product.IRP;
            ////planning.RRCNDS2 = product.RRCFinal; //?
            //this.IRPIndex = product.IRPIndex;
            //this.DIYPriceList = product.DIY;

            this.SupercategoryEng = product.SupercategoryEng;
            this.Supercategory = product.Supercategory;
            this.ProductGroup = product.ProductGroup;
            this.ProductGroupEng = product.ProductGroupEng;
            this.SubGroup = product.SubGroup;
            this.GenericName = product.GenericName;
            this.PNS = product.PNS;
            this.Article = product.Article;
            this.ArticleOld = product.ArticleOld;
            this.ArticleRu = product.ArticleRu;
            this.CalendarSalesStartDate = product.CalendarSalesStartDate;
            this.CalendarPreliminaryEliminationDate = product.CalendarPreliminaryEliminationDate;
            this.CalendarEliminationDate = product.CalendarEliminationDate;
            this.Status = product.Status;
        }
        public void GetSTK()
        {
            //if (this.Article == "")
            //    return;

            //STK sTK = new STK().GetSTK(this.Article, this.Year);
            //if (sTK == null)
            //    return;

            //this.STKEur = sTK.STKEur;
            //this.STKRub = sTK.STKRub;
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
            try
            {
                PlanningNewYear planningNewYear = new PlanningNewYear(worksheetName);
                if (planningNewYear == null)
                    return null;
                ThisWorkbook workbook = Globals.ThisWorkbook;
                Range rng = workbook.Sheets[planningNewYear.SheetName].UsedRange;

                planningNewYear.CustomerStatus = val(CustomerStatusLabel);
                planningNewYear.ChannelType = val(ChannelTypeLabel);
                if (Int32.TryParse(val(YearLabel), out int year))
                {
                    planningNewYear.Year = year;
                    planningNewYear.planningDate = new DateTime(year, 1, 1);
                }
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
            string tableTemplateName = tableTemplate.Name;

            Table.ListRows.AddEx();
            Range firstCell = Table.ListRows[1].Range[1];

            tableTemplate.ListRows[1].Range.Copy();
            firstCell.PasteSpecial();

            foreach (Range cell in Table.ListRows[1].Range)
            {
                string formula = cell.FormulaLocal;

                if (formula != "" && formula.Substring(0,1) == "=")
                //formula = formula.Replace(tableTemplateName, "");
                //cell.FormulaLocal = formula;
                    cell.FormulaLocal = cell.FormulaLocal.Replace(tableTemplateName, "");
            }

        }

        public void SetMaximumBonusValue()
        {
            ThisWorkbook workbook = Globals.ThisWorkbook;
            Worksheet worksheet = workbook.Sheets[SheetName];
            Range range = worksheet.UsedRange;
            
            Range cell = range.Find(What:MaximumBonusLabel,LookIn:XlFindLookIn.xlFormulas, LookAt:XlLookAt.xlPart, MatchCase:false);
            if (cell == null)
                return;

            cell.Offset[0, -2].Value = this.MaximumBonus;
        }

        #region Получение списков
        public void SetLists(List<PlanningNewYearPrognosis> prognosises, List<PlanningNewYearPromo> promos)
        {
            List<PlanningNewYear> plannings = GetList();
            if (plannings == null)
                return;

            ProcessBar processBar = null;
            processBar = new ProcessBar($"Сбор данных для сохранения { SheetName } ", Table.ListRows.Count);
            bool isCancel = false;
            void CancelLocal() => isCancel = true;
            processBar.CancelClick += CancelLocal;
            processBar.Show();

            foreach (PlanningNewYear planning in plannings)
            {
                if (isCancel)
                    break;
                processBar.TaskStart($"Обрабатывается артикул { planning.Article }");

                prognosises.Add(new PlanningNewYearPrognosis(planning));
                promos.Add(new PlanningNewYearPromo(planning));
                processBar.TaskDone(1);
            }
            processBar.Close();
            processBar = null;
        }
        public void SetLists(List<PlanningNewYearSave> saves)
        {
            List<PlanningNewYear> plannings = GetList();
            if (plannings == null)
                return;
            
            ProcessBar processBar = null;
            processBar = new ProcessBar($"Сбор данных для сохранения { SheetName } ", Table.ListRows.Count);
            bool isCancel = false;
            void CancelLocal() => isCancel = true;
            processBar.CancelClick += CancelLocal;
            processBar.Show();

            foreach (PlanningNewYear planning in plannings)
            {
                if (isCancel)
                    break;
                processBar.TaskStart($"Обрабатывается артикул { planning.Article }");

                PlanningNewYearSave planningNewYearSave = new PlanningNewYearSave(planning);
                planningNewYearSave.SetValues();
                saves.Add(planningNewYearSave);
                
                processBar.TaskDone(1);
            }
            processBar.Close();
            processBar = null;
        }

        private List<PlanningNewYear> GetList()
        {
            ProcessBar processBar = null;
            if (Table.ListRows.Count < 1)
                return null;

            processBar = new ProcessBar($"Получение списка артикулов на листе { SheetName } ", Table.ListRows.Count);
            bool isCancel = false;
            void CancelLocal() => isCancel = true;
            FunctionsForExcel.SpeedOn();
            processBar.CancelClick += CancelLocal;
            processBar.Show();

            List<PlanningNewYear> plannings = new List<PlanningNewYear>();

            foreach (ListRow listRow in Table.ListRows)
            {
                if (isCancel)
                    break;

                PlanningNewYear planning = this.Clone();
                planning.SetProperty(listRow);
                processBar.TaskStart($"Обрабатывается артикул { planning.Article }");
                if ((int)planning.Id != 0)
                    plannings.Add(planning);

                processBar.TaskDone(1);
            }

            processBar.Close();
            processBar = null;

            return plannings;
        }
        #endregion

        #region проверка promo/prognosis
        /// <summary>
        /// проверка promo/prognosis
        /// </summary>
        /// <param name="articleQuantity"></param>
        /// <returns></returns>
        public bool isPromo(ArticleQuantity articleQuantity)
        {
            return articleQuantity.Campaign != "0" && articleQuantity.Campaign != null ? true : false;
        }

        #endregion

        public ArticleQuantity[] GetsArticleQuantities(List<ArticleQuantity> articleDescisionQuantities, List<ArticleQuantity> articleBugetQuantities)
        {
            ArticleQuantity[] articles = new ArticleQuantity[12];
            for (int m = 1; m <= 12; m++)
            {
                articles[m - 1] = SumMonth(m);
            }
            return articles;

            ArticleQuantity SumMonth(double month)
            {
                ArticleQuantity newArticleQuantity = new ArticleQuantity();

                List<ArticleQuantity> articleQuantities = month < CurrentMonth ?
                    articleDescisionQuantities : articleBugetQuantities;

                if (articleQuantities.Count < 1)
                    return newArticleQuantity;

                //на случай если будет несколько записей на один месяц по одному артикула
                List<ArticleQuantity> monthQuantities = articleQuantities.FindAll(x => x.Month == month);

                if (monthQuantities.Count < 1)
                    return newArticleQuantity;

                foreach (ArticleQuantity articleQuantity in monthQuantities)
                {
                    newArticleQuantity.Quantity += articleQuantity.Quantity;
                    newArticleQuantity.PriceList += articleQuantity.PriceList;
                    double bonus = month < CurrentMonth ? articleQuantity.Bonus : articleQuantity.PriceList * MaximumBonus;
                    newArticleQuantity.Bonus += bonus;
                }
                //

                newArticleQuantity.Article= monthQuantities[0].Article;
                newArticleQuantity.Campaign= monthQuantities[0].Campaign;

                return newArticleQuantity;
            }
            //
        }
    }
}
