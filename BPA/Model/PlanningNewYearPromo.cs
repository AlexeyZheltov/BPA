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

            { "GSPromoYear","Промо GS за год, руб." },
            { "GSPromo01","Промо GS январь, руб." },
            { "GSPromo02","Промо GS февраль, руб." },
            { "GSPromo03","Промо GS март, руб." },
            { "GSPromo04","Промо GS апрель, руб." },
            { "GSPromo05","Промо GS май, руб." },
            { "GSPromo06","Промо GS июнь, руб." },
            { "GSPromo07","Промо GS июль, руб." },
            { "GSPromo08","Промо GS август, руб." },
            { "GSPromo09","Промо GS сентрябрь, руб." },
            { "GSPromo10","Промо GS октябрь, руб." },
            { "GSPromo11","Промо GS ноябрь, руб." },
            { "GSPromo12","Промо GS декабрь, руб." },

            { "NSPromoYear","Промо NS за год, руб." },
            { "NSPromo01","Промо NS январь, руб." },
            { "NSPromo02","Промо NS февраль, руб." },
            { "NSPromo03","Промо NS март, руб." },
            { "NSPromo04","Промо NS апрель, руб." },
            { "NSPromo05","Промо NS май, руб." },
            { "NSPromo06","Промо NS июнь, руб." },
            { "NSPromo07","Промо NS июль, руб." },
            { "NSPromo08","Промо NS август, руб." },
            { "NSPromo09","Промо NS сентрябрь, руб." },
            { "NSPromo10","Промо NS октябрь, руб." },
            { "NSPromo11","Промо NS ноябрь, руб." },
            { "NSPromo12","Промо NS декабрь, руб." }
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
        /// Промо GS сентрябрь, руб. 
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
        /// Промо NS сентрябрь, руб. 
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

            List<ArticleQuantity> articleDescisionQuantities = deicionQuantities.FindAll(x => x.Article == article && this.planningNewYear.isPromo(x)).ToList();
            List<ArticleQuantity> articleBugetQuantities = bugetQuantities.FindAll(x => x.Article == article && this.planningNewYear.isPromo(x)).ToList();

            double[] quantities = this.planningNewYear.GetQuantities(articleDescisionQuantities, articleBugetQuantities);

            #region setproperties
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
            #endregion
        }
    }    
}
