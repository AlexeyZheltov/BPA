using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using BPA.Modules;

namespace BPA.Model
{
    /// <summary>
    /// For PlanningNewYear.cs
    /// </summary>
    class PlanningNewYearPromo : TableBase
    {
        public PlanningNewYear planningNewYear;
        public PlanningNewYearPromo(PlanningNewYear planningNewYear)
        {
            this.planningNewYear = planningNewYear;
        }

        public override string TableName => this.planningNewYear.GetTableName();
        public override string SheetName => this.planningNewYear._TableWorksheetName != "" ?
            this.planningNewYear._TableWorksheetName :
            this.planningNewYear.templateSheetName;

        #region --- Словарь ---

        public override IDictionary<string, string> Filds => _filds;
        private readonly Dictionary<string, string> _filds = new Dictionary<string, string>
        {
            { "Id","№" },
            { "QuantityPromoYear","Промо прогноз за год, шт. " },
            { "QuantityPromo01","Промо прогноз январь, шт." },
            { "QuantityPromo02","Промо прогноз февраль, шт." },
            { "QuantityPromo03","Промо прогноз март, шт." },
            { "QuantityPromo04","Промо прогноз апрель, шт." },
            { "QuantityPromo05","Промо прогноз май, шт." },
            { "QuantityPromo06","Промо прогноз июнь, шт." },
            { "QuantityPromo07","Промо прогноз июль, шт." },
            { "QuantityPromo08","Промо прогноз август, шт." },
            { "QuantityPromo09","Промо прогноз сентрябрь, шт." },
            { "QuantityPromo10","Промо прогноз октябрь, шт." },
            { "QuantityPromo11","Промо прогноз ноябрь, шт." },
            { "QuantityPromo12","Промо прогноз декабрь, шт." },         
        };

        #endregion

        #region -- Основные свойства столбцов ---
        /// <summary>
        /// №
        /// </summary>
        public int Id
        {
            get
            {
                return this.planningNewYear.Id;
            }
            set
            {
            }
        }
        /// <summary>
        /// Промо прогноз за год, шт.  
        /// </summary>
        public double QuantityPromoYear
        {
            get; set;
        } 
        /// <summary>
        /// Промо прогноз январь, шт. 
        /// </summary>
        public double QuantityPromo01
        {
            get; set;
        } 
        /// <summary>
        /// Промо прогноз февраль, шт. 
        /// </summary>
        public double QuantityPromo02
        {
            get; set;
        } 
        /// <summary>
        /// Промо прогноз март, шт. 
        /// </summary>
        public double QuantityPromo03
        {
            get; set;
        } 
        /// <summary>
        /// Промо прогноз апрель, шт. 
        /// </summary>
        public double QuantityPromo04
        {
            get; set;
        } 
        /// <summary>
        /// Промо прогноз май, шт. 
        /// </summary>
        public double QuantityPromo05
        {
            get; set;
        } 
        /// <summary>
        /// Промо прогноз июнь, шт. 
        /// </summary>
        public double QuantityPromo06
        {
            get; set;
        } 
        /// <summary>
        /// Промо прогноз июль, шт. 
        /// </summary>
        public double QuantityPromo07
        {
            get; set;
        } 
        /// <summary>
        /// Промо прогноз август, шт. 
        /// </summary>
        public double QuantityPromo08
        {
            get; set;
        } 
        /// <summary>
        /// Промо прогноз сентрябрь, шт. 
        /// </summary>
        public double QuantityPromo09
        {
            get; set;
        } 
        /// <summary>
        /// Промо прогноз октябрь, шт. 
        /// </summary>
        public double QuantityPromo10
        {
            get; set;
        } 
        /// <summary>
        /// Промо прогноз ноябрь, шт. 
        /// </summary>
        public double QuantityPromo11
        {
            get; set;
        } 
        /// <summary>
        /// Промо прогноз декабрь, шт. 
        /// </summary>
        public double QuantityPromo12
        {
            get; set;
        } 


        #endregion

        private int CurrentMonth => DateTime.Now.Month;
        public void SetValues(List<ArticleQuantity> deicionQuantities, List<ArticleQuantity> bugetQuantities)
        {
            string article = this.planningNewYear.Article;

            List<ArticleQuantity> articleDescisionQuantities = deicionQuantities.FindAll(x => x.Article == article && isPromo(x)).ToList();
            List<ArticleQuantity> articleBugetQuantities = bugetQuantities.FindAll(x => x.Article == article && isPromo(x)).ToList();

            double[] quantities = new double[12];
            for (int m = 1; m <= 12; m++)
            {
                quantities[m - 1] = m < CurrentMonth ? SumMonthQuantity(m, articleDescisionQuantities) : SumMonthQuantity(m, articleBugetQuantities);
            }

            //как написать подобный перебор???
            ///
            QuantityPromo01 = quantities[0];
            QuantityPromo02 = quantities[1];
            QuantityPromo03 = quantities[2];
            QuantityPromo04 = quantities[3];
            QuantityPromo05 = quantities[4];
            QuantityPromo06 = quantities[5];
            QuantityPromo07 = quantities[6];
            QuantityPromo08 = quantities[7];
            QuantityPromo09 = quantities[8];
            QuantityPromo10 = quantities[9];
            QuantityPromo11 = quantities[10];
            QuantityPromo12 = quantities[11];
            ///

            bool isPromo(ArticleQuantity articleQuantity)
            {
                return articleQuantity.Campaign != "0" && articleQuantity.Campaign != null ? true : false;
            }

            double SumMonthQuantity(double month, List<ArticleQuantity> articleQuantities)
            {
                if (articleQuantities.Count <= 0)
                    return 0;

                List<ArticleQuantity> MohthQuantities = articleQuantities.FindAll(x => x.Month == month);
                double quantity = 0;

                foreach (ArticleQuantity articleQuantity in MohthQuantities)
                {
                    quantity += articleQuantity.Quantity;
                }
                return quantity;
            }
        }
    }    
}
