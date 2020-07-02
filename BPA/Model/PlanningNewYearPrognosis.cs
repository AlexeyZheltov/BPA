using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using BPA.Model;
using BPA.Modules;
using Microsoft.Office.Interop.Excel;

namespace BPA.Model
{
    /// <summary>
    /// For PlanningNewYear.cs
    /// </summary>
    class PlanningNewYearPrognosis : TableBase
    {
        public PlanningNewYear planningNewYear;
        public PlanningNewYearPrognosis(PlanningNewYear planningNewYear)
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
            { "QuantityPrognosisYear","Прогноз за год, шт." },
            { "QuantityPrognosis01","Прогноз январь, шт." },
            { "QuantityPrognosis02","Прогноз февраль, шт." },
            { "QuantityPrognosis03","Прогноз март, шт." },
            { "QuantityPrognosis04","Прогноз апрель, шт." },
            { "QuantityPrognosis05","Прогноз май, шт." },
            { "QuantityPrognosis06","Прогноз июнь, шт." },
            { "QuantityPrognosis07","Прогноз июль, шт." },
            { "QuantityPrognosis08","Прогноз август, шт." },
            { "QuantityPrognosis09","Прогноз сентрябрь, шт." },
            { "QuantityPrognosis10","Прогноз октябрь, шт." },
            { "QuantityPrognosis11","Прогноз ноябрь, шт." },
            { "QuantityPrognosis12","Прогноз декабрь, шт." },
        };

        #endregion

        #region -- Основные свойства столбцов ---
        /// <summary>
        /// Прогноз за год, шт. 
        /// </summary>
        public double QuantityPrognosisYear
        {
            get; set;
        }
        /// <summary>
        /// Прогноз январь, шт. 
        /// </summary>
        public double QuantityPrognosis01
        {
            get; set;
        }
        /// <summary>
        /// Прогноз февраль, шт. 
        /// </summary>
        public double QuantityPrognosis02
        {
            get; set;
        }
        /// <summary>
        /// Прогноз март, шт. 
        /// </summary>
        public double QuantityPrognosis03
        {
            get; set;
        }
        /// <summary>
        /// Прогноз апрель, шт. 
        /// </summary>
        public double QuantityPrognosis04
        {
            get; set;
        }
        /// <summary>
        /// Прогноз май, шт. 
        /// </summary>
        public double QuantityPrognosis05
        {
            get; set;
        }
        /// <summary>
        /// Прогноз июнь, шт. 
        /// </summary>
        public double QuantityPrognosis06
        {
            get; set;
        }
        /// <summary>
        /// Прогноз июль, шт. 
        /// </summary>
        public double QuantityPrognosis07
        {
            get; set;
        }
        /// <summary>
        /// Прогноз август, шт. 
        /// </summary>
        public double QuantityPrognosis08
        {
            get; set;
        }
        /// <summary>
        /// Прогноз сентрябрь, шт. 
        /// </summary>
        public double QuantityPrognosis09
        {
            get; set;
        }
        /// <summary>
        /// Прогноз октябрь, шт. 
        /// </summary>
        public double QuantityPrognosis10
        {
            get; set;
        }
        /// <summary>
        /// Прогноз ноябрь, шт. 
        /// </summary>
        public double QuantityPrognosis11
        {
            get; set;
        }
        /// <summary>
        /// Прогноз декабрь, шт. 
        /// </summary>
        public double QuantityPrognosis12
        {
            get; set;
        }
        #endregion

        private int CurrentMonth => DateTime.Now.Month;
        public void SetValues(List<ArticleQuantity> deicionQuantities, List<ArticleQuantity> bugetQuantities)
        {
            string article = this.planningNewYear.Article;

            List<ArticleQuantity> articleDescisionQuantities = deicionQuantities.FindAll(x => x.Article == article).ToList();
            List<ArticleQuantity> articleBugetQuantities = bugetQuantities.FindAll(x => x.Article == article).ToList();

            double[] quantities = new double[12];
            for (int m = 1; m <= 12; m++)
            {
                quantities[m - 1] = m < CurrentMonth ? SumMonthQuantity(m, articleDescisionQuantities) : SumMonthQuantity(m, articleBugetQuantities);
            }

            //как написать подобный перебор???
            ///
            QuantityPrognosis01 = quantities[1];
            QuantityPrognosis02 = quantities[2];
            QuantityPrognosis03 = quantities[3];
            QuantityPrognosis04 = quantities[4];
            QuantityPrognosis05 = quantities[5];
            QuantityPrognosis06 = quantities[6];
            QuantityPrognosis07 = quantities[7];
            QuantityPrognosis08 = quantities[8];
            QuantityPrognosis09 = quantities[9];
            QuantityPrognosis10 = quantities[10];
            QuantityPrognosis11 = quantities[11];
            QuantityPrognosis12 = quantities[12];
            ///

            double SumMonthQuantity(double month, List<ArticleQuantity> articleQuantities)
            {
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
