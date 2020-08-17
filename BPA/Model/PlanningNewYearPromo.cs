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
            planningNewYear.templateSheetName;

        public static Dictionary<string, int> ColDict { get; set; } = new Dictionary<string, int>();


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
            { "QuantityPromo09","Промо прогноз сентябрь, шт." },
            { "QuantityPromo10","Промо прогноз октябрь, шт." },
            { "QuantityPromo11","Промо прогноз ноябрь, шт." },
            { "QuantityPromo12","Промо прогноз декабрь, шт." },

            //{ "GSPromoYear","Промо GS за год, шт." },
            //{ "GSPromo01","Промо GS январь, шт." },
            //{ "GSPromo02","Промо GS февраль, шт." },
            //{ "GSPromo03","Промо GS март, шт." },
            //{ "GSPromo04","Промо GS апрель, шт." },
            //{ "GSPromo05","Промо GS май, шт." },
            //{ "GSPromo06","Промо GS июнь, шт." },
            //{ "GSPromo07","Промо GS июль, шт." },
            //{ "GSPromo08","Промо GS август, шт." },
            //{ "GSPromo09","Промо GS сентябрь, шт." },
            //{ "GSPromo10","Промо GS октябрь, шт." },
            //{ "GSPromo11","Промо GS ноябрь, шт." },
            //{ "GSPromo12","Промо GS декабрь, шт." },

            //{ "NSPromoYear","Промо NS за год, шт." },
            //{ "NSPromo01","Промо NS январь, шт." },
            //{ "NSPromo02","Промо NS февраль, шт." },
            //{ "NSPromo03","Промо NS март, шт." },
            //{ "NSPromo04","Промо NS апрель, шт." },
            //{ "NSPromo05","Промо NS май, шт." },
            //{ "NSPromo06","Промо NS июнь, шт." },
            //{ "NSPromo07","Промо NS июль, шт." },
            //{ "NSPromo08","Промо NS август, шт." },
            //{ "NSPromo09","Промо NS сентябрь, шт." },
            //{ "NSPromo10","Промо NS октябрь, шт." },
            //{ "NSPromo11","Промо NS ноябрь, шт." },
            //{ "NSPromo12","Промо NS декабрь, шт." }
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
        /// Промо прогноз сентябрь, шт. 
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
        /// <summary>
        /// Промо GS за год, руб. 
        /// </summary>
        public double GSPromoYear
        {
            get; set;
        } 
        /// <summary>
        /// Промо GS январь, руб. 
        /// </summary>
        public double GSPromo01
        {
            get; set;
        }
        /// <summary>
        /// Промо GS февраль, руб. 
        /// </summary>
        public double GSPromo02
        {
            get; set;
        }
        /// <summary>
        /// Промо GS март, руб. 
        /// </summary>
        public double GSPromo03
        {
            get; set;
        }
        /// <summary>
        /// Промо GS апрель, руб. 
        /// </summary>
        public double GSPromo04
        {
            get; set;
        }
        /// <summary>
        /// Промо GS май, руб. 
        /// </summary>
        public double GSPromo05
        {
            get; set;
        }
        /// <summary>
        /// Промо GS июнь, руб. 
        /// </summary>
        public double GSPromo06
        {
            get; set;
        }
        /// <summary>
        /// Промо GS июль, руб. 
        /// </summary>
        public double GSPromo07
        {
            get; set;
        }
        /// <summary>
        /// Промо GS август, руб. 
        /// </summary>
        public double GSPromo08
        {
            get; set;
        }
        /// <summary>
        /// Промо GS сентябрь, руб. 
        /// </summary>
        public double GSPromo09
        {
            get; set;
        }
        /// <summary>
        /// Промо GS октябрь, руб. 
        /// </summary>
        public double GSPromo10
        {
            get; set;
        }
        /// <summary>
        /// Промо GS ноябрь, руб. 
        /// </summary>
        public double GSPromo11
        {
            get; set;
        }
        /// <summary>
        /// Промо GS декабрь, руб. 
        /// </summary>
        public double GSPromo12
        {
            get; set;
        }
        /// <summary>
        /// Промо NS за год, руб. 
        /// </summary>
        public double NSPromoYear
        {
            get; set;
        }
        /// <summary>
        /// Промо NS январь, руб. 
        /// </summary>
        public double NSPromo01
        {
            get; set;
        }
        /// <summary>
        /// Промо NS февраль, руб. 
        /// </summary>
        public double NSPromo02
        {
            get; set;
        }
        /// <summary>
        /// Промо NS март, руб. 
        /// </summary>
        public double NSPromo03
        {
            get; set;
        }
        /// <summary>
        /// Промо NS апрель, руб. 
        /// </summary>
        public double NSPromo04
        {
            get; set;
        }
        /// <summary>
        /// Промо NS май, руб. 
        /// </summary>
        public double NSPromo05
        {
            get; set;
        }
        /// <summary>
        /// Промо NS июнь, руб. 
        /// </summary>
        public double NSPromo06
        {
            get; set;
        }
        /// <summary>
        /// Промо NS июль, руб. 
        /// </summary>
        public double NSPromo07
        {
            get; set;
        }
        /// <summary>
        /// Промо NS август, руб. 
        /// </summary>
        public double NSPromo08
        {
            get; set;
        }
        /// <summary>
        /// Промо NS сентябрь, руб. 
        /// </summary>
        public double NSPromo09
        {
            get; set;
        }
        /// <summary>
        /// Промо NS октябрь, руб. 
        /// </summary>
        public double NSPromo10
        {
            get; set;
        }
        /// <summary>
        /// Промо NS ноябрь, руб. 
        /// </summary>
        public double NSPromo11
        {
            get; set;
        }
        /// <summary>
        /// Промо NS декабрь, руб. 
        /// </summary>
        public double NSPromo12
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

            //Извлечение из списков Descision и Buget элементы с соответствующим артикулом и Promo
            List<ArticleQuantity> articleDescisionQuantities = deicionQuantities.FindAll(x => x.Article == article && this.planningNewYear.isPromo(x)).ToList();
            List<ArticleQuantity> articleBugetQuantities = bugetQuantities.FindAll(x => x.Article == article && this.planningNewYear.isPromo(x)).ToList();

            ArticleQuantity[] articles = this.planningNewYear.GetsArticleQuantities(articleDescisionQuantities, articleBugetQuantities);

            #region setproperties
            //как написать подобный перебор???
            ///
            QuantityPromo01 = articles[0].Quantity;
            QuantityPromo02 = articles[1].Quantity;
            QuantityPromo03 = articles[2].Quantity;
            QuantityPromo04 = articles[3].Quantity;
            QuantityPromo05 = articles[4].Quantity;
            QuantityPromo06 = articles[5].Quantity;
            QuantityPromo07 = articles[6].Quantity;
            QuantityPromo08 = articles[7].Quantity;
            QuantityPromo09 = articles[8].Quantity;
            QuantityPromo10 = articles[9].Quantity;
            QuantityPromo11 = articles[10].Quantity;
            QuantityPromo12 = articles[11].Quantity;

            //GSPromo01 = articles[0].PriceList;
            //GSPromo02 = articles[1].PriceList;
            //GSPromo03 = articles[2].PriceList;
            //GSPromo04 = articles[3].PriceList;
            //GSPromo05 = articles[4].PriceList;
            //GSPromo06 = articles[5].PriceList;
            //GSPromo07 = articles[6].PriceList;
            //GSPromo08 = articles[7].PriceList;
            //GSPromo09 = articles[8].PriceList;
            //GSPromo10 = articles[9].PriceList;
            //GSPromo11 = articles[10].PriceList;
            //GSPromo12 = articles[11].PriceList;

            //NSPromo01 = GSPromo01 - articles[0].Bonus;
            //NSPromo02 = GSPromo02 - articles[1].Bonus;
            //NSPromo03 = GSPromo03 - articles[2].Bonus;
            //NSPromo04 = GSPromo04 - articles[3].Bonus;
            //NSPromo05 = GSPromo05 - articles[4].Bonus;
            //NSPromo06 = GSPromo06 - articles[5].Bonus;
            //NSPromo07 = GSPromo07 - articles[6].Bonus;
            //NSPromo08 = GSPromo08 - articles[7].Bonus;
            //NSPromo09 = GSPromo09 - articles[8].Bonus;
            //NSPromo10 = GSPromo10 - articles[9].Bonus;
            //NSPromo11 = GSPromo11 - articles[10].Bonus;
            //NSPromo12 = GSPromo12 - articles[11].Bonus;
            ///
            #endregion
        }
    }    
}
