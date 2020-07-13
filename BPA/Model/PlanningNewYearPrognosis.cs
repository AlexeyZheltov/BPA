using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
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
            { "Id","№" },

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

            { "GSPrognosisYear","GS за год, руб." },
            { "GSPrognosis01","GS январь, руб." },
            { "GSPrognosis02","GS февраль, руб." },
            { "GSPrognosis03","GS март, руб." },
            { "GSPrognosis04","GS апрель, руб." },
            { "GSPrognosis05","GS май, руб." },
            { "GSPrognosis06","GS июнь, руб." },
            { "GSPrognosis07","GS июль, руб." },
            { "GSPrognosis08","GS август, руб." },
            { "GSPrognosis09","GS сентрябрь, руб." },
            { "GSPrognosis10","GS октябрь, руб." },
            { "GSPrognosis11","GS ноябрь, руб." },
            { "GSPrognosis12","GS декабрь, руб." },

            { "NSPrognosisYear","NS за год, руб." },
            { "NSPrognosis01","NS январь, руб." },
            { "NSPrognosis02","NS февраль, руб." },
            { "NSPrognosis03","NS март, руб." },
            { "NSPrognosis04","NS апрель, руб." },
            { "NSPrognosis05","NS май, руб." },
            { "NSPrognosis06","NS июнь, руб." },
            { "NSPrognosis07","NS июль, руб." },
            { "NSPrognosis08","NS август, руб." },
            { "NSPrognosis09","NS сентрябрь, руб." },
            { "NSPrognosis10","NS октябрь, руб." },
            { "NSPrognosis11","NS ноябрь, руб." },
            { "NSPrognosis12","NS декабрь, шт." }
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
        /// <summary>
        /// GS за год, руб. 
        /// </summary>
        public double GSPrognosisYear
        {
            get; set;
        }
        /// <summary>
        /// GS январь, руб. 
        /// </summary>
        public double GSPrognosis01
        {
            get; set;
        }
        /// <summary>
        /// GS февраль, руб. 
        /// </summary>
        public double GSPrognosis02
        {
            get; set;
        }
        /// <summary>
        /// GS март, руб. 
        /// </summary>
        public double GSPrognosis03
        {
            get; set;
        }
        /// <summary>
        /// GS апрель, руб. 
        /// </summary>
        public double GSPrognosis04
        {
            get; set;
        }
        /// <summary>
        /// GS май, руб. 
        /// </summary>
        public double GSPrognosis05
        {
            get; set;
        }
        /// <summary>
        /// GS июнь, руб. 
        /// </summary>
        public double GSPrognosis06
        {
            get; set;
        }
        /// <summary>
        /// GS июль, руб. 
        /// </summary>
        public double GSPrognosis07
        {
            get; set;
        }
        /// <summary>
        /// GS август, руб. 
        /// </summary>
        public double GSPrognosis08
        {
            get; set;
        }
        /// <summary>
        /// GS сентрябрь, руб. 
        /// </summary>
        public double GSPrognosis09
        {
            get; set;
        }
        /// <summary>
        /// GS октябрь, руб. 
        /// </summary>
        public double GSPrognosis10
        {
            get; set;
        }
        /// <summary>
        /// GS ноябрь, руб. 
        /// </summary>
        public double GSPrognosis11
        {
            get; set;
        }
        /// <summary>
        /// GS декабрь, руб. 
        /// </summary>
        public double GSPrognosis12
        {
            get; set;
        }
        /// <summary>
        /// NS за год, руб. 
        /// </summary>
        public double NSPrognosisYear
        {
            get; set;
        }
        /// <summary>
        /// NS январь, руб. 
        /// </summary>
        public double NSPrognosis01
        {
            get; set;
        }
        /// <summary>
        /// NS февраль, руб. 
        /// </summary>
        public double NSPrognosis02
        {
            get; set;
        }
        /// <summary>
        /// NS март, руб. 
        /// </summary>
        public double NSPrognosis03
        {
            get; set;
        }
        /// <summary>
        /// NS апрель, руб. 
        /// </summary>
        public double NSPrognosis04
        {
            get; set;
        }
        /// <summary>
        /// NS май, руб. 
        /// </summary>
        public double NSPrognosis05
        {
            get; set;
        }
        /// <summary>
        /// NS июнь, руб. 
        /// </summary>
        public double NSPrognosis06
        {
            get; set;
        }
        /// <summary>
        /// NS июль, руб. 
        /// </summary>
        public double NSPrognosis07
        {
            get; set;
        }
        /// <summary>
        /// NS август, руб. 
        /// </summary>
        public double NSPrognosis08
        {
            get; set;
        }
        /// <summary>
        /// NS сентрябрь, руб. 
        /// </summary>
        public double NSPrognosis09
        {
            get; set;
        }
        /// <summary>
        /// NS октябрь, руб. 
        /// </summary>
        public double NSPrognosis10
        {
            get; set;
        }
        /// <summary>
        /// NS ноябрь, руб. 
        /// </summary>
        public double NSPrognosis11
        {
            get; set;
        }
        /// <summary>
        /// NS декабрь, шт. 
        /// </summary>
        public double NSPrognosis12
        {
            get; set;
        }
        #endregion
        
        /// <summary>
        /// установка свойств Промо
        /// </summary>
        /// <param name="deicionQuantities"></param>
        /// <param name="bugetQuantities"></param>
        public void SetValues(List<ArticleQuantity> deicionQuantities, List<ArticleQuantity> bugetQuantities)
        {
            string article = this.planningNewYear.Article;

            //Извлечение из списков Descision и Buget элементы с соответствующим артикулом и НЕ Promo
            List<ArticleQuantity> articleDescisionQuantities = deicionQuantities.FindAll(x => x.Article == article && !this.planningNewYear.isPromo(x)).ToList();
            List<ArticleQuantity> articleBugetQuantities = bugetQuantities.FindAll(x => x.Article == article && !this.planningNewYear.isPromo(x)).ToList();

            ArticleQuantity[] articles = this.planningNewYear.GetsArticleQuantities(articleDescisionQuantities, articleBugetQuantities);
            #region setproperties
            //как написать подобный перебор???
            ///
            QuantityPrognosis01 = articles[0].Quantity;
            QuantityPrognosis02 = articles[1].Quantity;
            QuantityPrognosis03 = articles[2].Quantity;
            QuantityPrognosis04 = articles[3].Quantity;
            QuantityPrognosis05 = articles[4].Quantity;
            QuantityPrognosis06 = articles[5].Quantity;
            QuantityPrognosis07 = articles[6].Quantity;
            QuantityPrognosis08 = articles[7].Quantity;
            QuantityPrognosis09 = articles[8].Quantity;
            QuantityPrognosis10 = articles[9].Quantity;
            QuantityPrognosis11 = articles[10].Quantity;
            QuantityPrognosis12 = articles[11].Quantity;

            GSPrognosis01 = articles[0].PriceList;
            GSPrognosis02 = articles[1].PriceList;
            GSPrognosis03 = articles[2].PriceList;
            GSPrognosis04 = articles[3].PriceList;
            GSPrognosis05 = articles[4].PriceList;
            GSPrognosis06 = articles[5].PriceList;
            GSPrognosis07 = articles[6].PriceList;
            GSPrognosis08 = articles[7].PriceList;
            GSPrognosis09 = articles[8].PriceList;
            GSPrognosis10 = articles[9].PriceList;
            GSPrognosis11 = articles[10].PriceList;
            GSPrognosis12 = articles[11].PriceList;

            NSPrognosis01 = GSPrognosis01 - articles[0].Bonus;
            NSPrognosis02 = GSPrognosis02 - articles[1].Bonus;
            NSPrognosis03 = GSPrognosis03 - articles[2].Bonus;
            NSPrognosis04 = GSPrognosis04 - articles[3].Bonus;
            NSPrognosis05 = GSPrognosis05 - articles[4].Bonus;
            NSPrognosis06 = GSPrognosis06 - articles[5].Bonus;
            NSPrognosis07 = GSPrognosis07 - articles[6].Bonus;
            NSPrognosis08 = GSPrognosis08 - articles[7].Bonus;
            NSPrognosis09 = GSPrognosis09 - articles[8].Bonus;
            NSPrognosis10 = GSPrognosis10 - articles[9].Bonus;
            NSPrognosis11 = GSPrognosis11 - articles[10].Bonus;
            NSPrognosis12 = GSPrognosis12 - articles[11].Bonus;
            ///
            #endregion
        }
    }
}
