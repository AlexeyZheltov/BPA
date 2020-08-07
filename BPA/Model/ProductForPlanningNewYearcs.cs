using BPA.Forms;
using BPA.Modules;
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
    class ProductForPlanningNewYear : TableBase
    {
        public override string TableName => "Товары";
        public override string SheetName => "Товары";

        public static Dictionary<string, int> ColDict { get; set; } = new Dictionary<string, int>();


        #region --- Словарь ---

        public override IDictionary<string, string> Filds => _filds;
        private readonly Dictionary<string, string> _filds = new Dictionary<string, string>
        {
            { "Id","№" },
            { "Article","Артикул" },

            { "Status","Актуальный статус" },
            { "Exclusive","Эксклюзив клиента или канала продажи" },

            //{ "IRP","IRP, Eur" },
            //{ "RRCCurrent","РРЦ текущий" },
            //{ "DIYCurrent","DIY текущий" },
            //{ "RRCPercent","Процент повышения РРЦ" },
            //{ "RRCCalculated","РРЦ расчетная, руб." },
            { "RRCFinal","РРЦ финальная, руб." },
            //{ "RRCEuro","РРЦ, евро" },
            //{ "IRPIndex","Индекс IRP" },
            //{ "DIYDiscount","Скидка DIY" },
            { "DIY","DIY price list, руб. без НДС" },

            { "SuperCategory","Суперкатегория" },
            { "SupercategoryEng","Суперкатегория (ENG)"  },
            { "ProductGroup","Продукт группа" },
            { "ProductGroupEng","Название продукт группы (ENG)" },
            { "SubGroup", "SubGroup" },
            { "GenericName", "Generic Name (long)" },
            { "PNS", "PNS" },

            { "ArticleOld","Артикул предшественника (если есть)" },
            { "ArticleRu","Название артикула (RUS)" },

            { "CalendarSalesStartDate","Sales Start Date" },
            { "CalendarPreliminaryEliminationDate","Preliminary Elimination Date" },
            { "CalendarEliminationDate","Elimination Date" },
        };

        #endregion

        #region -- Основные свойства столбцов ---

        /// <summary>
        /// №
        /// </summary>
        public int Id
        {
            get; set;
        }
        public string Article
        {
            get; set;
        }

        public string Status
        {
            get; set;
        }
        /// <summary>
        /// Эксклюзив клиента или канала продажи
        /// </summary>
        public string Exclusive
        {
            get; set;
        }
        #endregion

        //#region --- Свойства для РРЦ ---

        ///// <summary>
        ///// IRP, Eur
        ///// </summary>
        //public Double IRP
        //{
        //    get; set;
        //}

        ///// <summary>
        ///// РРЦ текущий
        ///// </summary>
        //public Double RRCCurrent
        //{
        //    get; set;
        //}

        ///// <summary>
        ///// DIY текущий
        ///// </summary>
        //public Double DIYCurrent
        //{
        //    get; set;
        //}

        ///// <summary>
        ///// Процент повышения РРЦ
        ///// </summary>
        //public Double RRCPercent
        //{
        //    get; set;
        //}

        ///// <summary>
        ///// РРЦ расчетная, руб.
        ///// </summary>
        //public Double RRCCalculated
        //{
        //    get; set;
        //}

        /// <summary>
        /// РРЦ финальная, руб.
        /// </summary>
        public Double RRCFinal
        {
            get; set;
        }

        ///// <summary>
        ///// РРЦ, евро
        ///// </summary>
        //public Double RRCEuro
        //{
        //    get; set;
        //}

        ///// <summary>
        ///// Индекс IRP
        ///// </summary>
        //public Double IRPIndex
        //{
        //    get; set;
        //}

        ///// <summary>
        ///// Скидка DIY
        ///// </summary>
        //public Double DIYDiscount
        //{
        //    get; set;
        //}
        /// <summary>
        /// DIY price list, руб. без НДС
        /// </summary>
        public Double DIY
        {
            get; set;
        }
        //#endregion

        #region --- Остальные свойства ---
        /// <summary>
        /// Суперкатегория(ENG)
        /// </summary>
        public string SuperCategory
        {
            get; set;
        }
        /// <summary>
        /// Суперкатегория
        /// </summary>
        public string SupercategoryEng
        {
            get; set;
        }
 
        public string ProductGroup
        {
            get; set;
        }

        /// <summary>
        /// Название продукт группы(ENG)
        /// </summary>
        public string ProductGroupEng
        {
            get; set;
        }

        /// <summary>
        /// SubGroup
        /// </summary>
        public string SubGroup
        {
            get; set;
        }
        /// <summary>
        /// Generic Name(long)
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
        /// Артикул предшественника(если есть)
        /// </summary>
        public string ArticleOld
        {
            get; set;
        }
        /// <summary>
        /// Название артикула(ENG)
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
        /// CalendarEliminationDate
        /// </summary>
        public Double CalendarEliminationDate
        {
            get; set;
        }

        #endregion

        public ProductForPlanningNewYear()
        {
        }
        public ProductForPlanningNewYear(ListRow listRow) => SetProperty(listRow);

        public int StatusId
        {
            get
            {
                ProductStatus status = new ProductStatus(this.Status);
                return status.Id;
            } set
            {
                StatusId = value;
            }
        }

        /// <summary>
        /// Возвращает список всех продуктов со статусом коме 2
        /// </summary>
        /// <returns></returns>
        //public List<ProductForPlanningNewYear> GetProducts()
        //{
        //    bool isCancel = false;
        //    void CancelLocal() => isCancel = true;
        //    ProcessBar processBar = new ProcessBar("Получение списка продуктов", LastRow);
        //    processBar.CancelClick += CancelLocal;
        //    processBar.Show();

        //    List<ProductForPlanningNewYear> products = new List<ProductForPlanningNewYear>();
        //    foreach (ListRow row in Table.ListRows)
        //    {
        //        if (isCancel)
        //        {
        //            processBar.Close();
        //            return null;
        //        }
        //        processBar.TaskStart($"Обрабатывается товар {row.Index} из {LastRow - Table.HeaderRowRange.Row}");

        //        ProductForPlanningNewYear product = new ProductForPlanningNewYear(row);

        //        if (StatusId != 2)
        //            products.Add(product);                

        //        processBar.TaskDone(1);
        //    }
        //    processBar.Close();
        //    return products;
        //}

        public List<ProductForPlanningNewYear> GetProducts(PlanningNewYear planningNewYearTmp)
        {
            bool isCancel = false;
            void CancelLocal() => isCancel = true;
            ProcessBar processBar = new ProcessBar("Получение списка продуктов", LastRow);
            processBar.CancelClick += CancelLocal;
            processBar.Show();

            List<ProductForPlanningNewYear> products = new List<ProductForPlanningNewYear>();
            foreach (ListRow row in Table.ListRows)
            {
                if (isCancel)
                {
                    processBar.Close();
                    return null;
                }
                processBar.TaskStart($"Обрабатывается товар {row.Index} из {LastRow - Table.HeaderRowRange.Row}");

                ProductForPlanningNewYear product = new ProductForPlanningNewYear(row);

                if (IsExclusive(planningNewYearTmp, product.Exclusive))
                    if (product.Status != "Выведено из ассортимента текущего года" && product.Status != "Выведено из глобального ассортимента")
                        if ((int)product.Id != 0)
                            products.Add(product);
                
                processBar.TaskDone(1);
            }
            processBar.Close();
            return products;
        }

        private bool IsExclusive(PlanningNewYear planningNewYearTmp, string exclusive)
        {
            if (exclusive == null)
                exclusive = "";
            try
            {
                exclusive = FunctionsForExcel.StringNormalize(exclusive, true);

                if (exclusive == "")
                    return true;

                if (planningNewYearTmp.ExclusivesDict.ContainsKey(exclusive))
                {
                    if (planningNewYearTmp.ExclusivesDict[exclusive] == FunctionsForExcel.StringNormalize(planningNewYearTmp.CustomerStatus,true) ||
                        planningNewYearTmp.ExclusivesDict[exclusive] == FunctionsForExcel.StringNormalize(planningNewYearTmp.ChannelType,true) || 
                        exclusive == "all channels")

                    {
                        return true;
                    }
                }
            }
            catch { return false; }
            return false;
        }
    }
}
