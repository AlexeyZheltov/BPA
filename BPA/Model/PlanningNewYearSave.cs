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
    class PlanningNewYearSave : TableBase
    {
        public PlanningNewYear planningNewYear;
        public PlanningNewYearSave(PlanningNewYear planningNewYear)
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
            { "Id", "№" },
            { "Article", "Артикул" },              //?
            { "STKRub", "STK 2.5, руб." },
            { "IRPEur", "IRP, Eur" },
            { "RRC", "РРЦ, руб.с НДС" },
            { "IRPIndex", "Индекс IRP" },
            { "DIY", "DIY price list, руб. без НДС" },
            { "Price2Net", "2 net цена, руб." },
            { "Price3Net", "3 net цена, руб." },
            { "PriceTransfer", "Transfer цена, руб." },

            //12 столбцов PriceList
            { "PriceList01", "январь" },
            { "PriceList02", "февраль" },
            { "PriceList03", "март" },
            { "PriceList04", "апрель" },
            { "PriceList05", "май" },
            { "PriceList06", "июнь" },
            { "PriceList07", "июль" },
            { "PriceList08", "август" },
            { "PriceList09", "сентябрь" },
            { "PriceList10", "октябрь" },
            { "PriceList11", "ноябрь" },
            { "PriceList12", "декабрь" },
            //

            //12 столбцов Объем продаж
            { "SalesVolume01", "январь" },
            { "SalesVolume02", "февраль" },
            { "SalesVolume03", "март" },
            { "SalesVolume04", "апрель" },
            { "SalesVolume05", "май" },
            { "SalesVolume06", "июнь" },
            { "SalesVolume07", "июль" },
            { "SalesVolume08", "август" },
            { "SalesVolume09", "сентябрь" },
            { "SalesVolume10", "октябрь" },
            { "SalesVolume11", "ноябрь" },
            { "SalesVolume12", "декабрь" },
            //

            { "GS01", "GS январь, руб." },
            { "GS02", "GS февраль, руб." },
            { "GS03", "GS март, руб." },
            { "GS04", "GS апрель, руб." },
            { "GS05", "GS май, руб." },
            { "GS06", "GS июнь, руб." },
            { "GS07", "GS июль, руб." },
            { "GS08", "GS август, руб." },
            { "GS09", "GS сентябрь, руб." },
            { "GS10", "GS октябрь, руб." },
            { "GS11", "GS ноябрь, руб." },
            { "GS12", "GS декабрь, руб." },

            { "NS01", "NS январь, руб." },
            { "NS02", "NS февраль, руб." },
            { "NS03", "NS март, руб." },
            { "NS04", "NS апрель, руб." },
            { "NS05", "NS май, руб." },
            { "NS06", "NS июнь, руб." },
            { "NS07", "NS июль, руб." },
            { "NS08", "NS август, руб." },
            { "NS09", "NS сентябрь, руб." },
            { "NS10", "NS октябрь, руб." },
            { "NS11", "NS ноябрь, руб." },
            { "NS12", "NS декабрь, руб." },

            //12 столбцов Promo PriceList
            { "PromoPriceList01", "январь" },
            { "PromoPriceList02", "февраль" },
            { "PromoPriceList03", "март" },
            { "PromoPriceList04", "апрель" },
            { "PromoPriceList05", "май" },
            { "PromoPriceList06", "июнь" },
            { "PromoPriceList07", "июль" },
            { "PromoPriceList08", "август" },
            { "PromoPriceList09", "сентябрь" },
            { "PromoPriceList10", "октябрь" },
            { "PromoPriceList11", "ноябрь" },
            { "PromoPriceList12", "декабрь" },
            //

            //12 столбцов Promo Объем продаж
            { "PromoSalesVolume01", "январь" },
            { "PromoSalesVolume02", "февраль" },
            { "PromoSalesVolume03", "март" },
            { "PromoSalesVolume04", "апрель" },
            { "PromoSalesVolume05", "май" },
            { "PromoSalesVolume06", "июнь" },
            { "PromoSalesVolume07", "июль" },
            { "PromoSalesVolume08", "август" },
            { "PromoSalesVolume09", "сентябрь" },
            { "PromoSalesVolume10", "октябрь" },
            { "PromoSalesVolume11", "ноябрь" },
            { "PromoSalesVolume12", "декабрь" },
            //

            { "PromoGS01", "Promo GS январь, руб." },
            { "PromoGS02", "Promo GS февраль, руб." },
            { "PromoGS03", "Promo GS март, руб." },
            { "PromoGS04", "Promo GS апрель, руб." },
            { "PromoGS05", "Promo GS май, руб." },
            { "PromoGS06", "Promo GS июнь, руб." },
            { "PromoGS07", "Promo GS июль, руб." },
            { "PromoGS08", "Promo GS август, руб." },
            { "PromoGS09", "Promo GS сентябрь, руб." },
            { "PromoGS10", "Promo GS октябрь, руб." },
            { "PromoGS11", "Promo GS ноябрь, руб." },
            { "PromoGS12", "Promo GS декабрь, руб." },

            { "PromoNS01", "Promo NS январь, руб." },
            { "PromoNS02", "Promo NS февраль, руб." },
            { "PromoNS03", "Promo NS март, руб." },
            { "PromoNS04", "Promo NS апрель, руб." },
            { "PromoNS05", "Promo NS май, руб." },
            { "PromoNS06", "Promo NS июнь, руб." },
            { "PromoNS07", "Promo NS июль, руб." },
            { "PromoNS08", "Promo NS август, руб." },
            { "PromoNS09", "Promo NS сентябрь, руб." },
            { "PromoNS10", "Promo NS октябрь, руб." },
            { "PromoNS11", "Promo NS ноябрь, руб." },
            { "PromoNS12", "Promo NS декабрь" },

            { "GPValue01","январь" },
            { "GPValue02","февраль" },
            { "GPValue03","март" },
            { "GPValue04","апрель" },
            { "GPValue05","май" },
            { "GPValue06","июнь" },
            { "GPValue07","июль" },
            { "GPValue08","август" },
            { "GPValue09","сентябрь" },
            { "GPValue10","октябрь" },
            { "GPValue11","ноябрь" },
            { "GPValue12","декабрь" }
        };
        #endregion

        #region -- Основные свойства столбцов ---
        /// <summary>
        /// Id
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
        /// Артикул
        /// </summary>
        public string Article
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
        ///  IRP
        /// </summary>
        public double IRPEur
        {
            get; set;
        }

        /// <summary>
        /// РРЦ, руб.с НДС
        /// </summary>
        public double RRC
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
        public double DIY
        {
            get; set;
        }

        /// <summary>
        /// 2 net цена, руб.
        /// </summary>
        public double Price2Net
        {
            get; set;
        }

        /// <summary>
        /// Зимние инструменты 2 net цена, руб.
        /// </summary>
        public double Price3Net
        {
            get; set;
        }

        /// <summary>
        /// Transfer цена, руб.
        /// </summary>
        public double PriceTransfer
        {
            get; set;
        }
        #endregion

        #region Price-list

        /// <summary>
        /// Price-list январь
        /// </summary>
        public double PriceList01
        {
            get; set;
        }

        /// <summary>
        /// Price-list февраль
        /// </summary>
        public double PriceList02
        {
            get; set;
        }

        /// <summary>
        /// Price-list март
        /// </summary>
        public double PriceList03
        {
            get; set;
        }

        /// <summary>
        /// Price-list апрель
        /// </summary>
        public double PriceList04
        {
            get; set;
        }

        /// <summary>
        /// Price-list май
        /// </summary>
        public double PriceList05
        {
            get; set;
        }

        /// <summary>
        /// Price-list июнь
        /// </summary>
        public double PriceList06
        {
            get; set;
        }

        /// <summary>
        /// Price-list июль
        /// </summary>
        public double PriceList07
        {
            get; set;
        }

        /// <summary>
        /// Price-list август
        /// </summary>
        public double PriceList08
        {
            get; set;
        }

        /// <summary>
        /// Price-list сентябрь
        /// </summary>
        public double PriceList09
        {
            get; set;
        }

        /// <summary>
        /// Price-list октябрь
        /// </summary>
        public double PriceList10
        {
            get; set;
        }

        /// <summary>
        /// Price-list ноябрь
        /// </summary>
        public double PriceList11
        {
            get; set;
        }

        /// <summary>
        /// Price-list декабрь
        /// </summary>
        public double PriceList12
        {
            get; set;
        }
        #endregion

        #region SalesVolume

        /// <summary>
        /// Объем продаж январь
        /// </summary>
        public double SalesVolume01
        {
            get; set;
        }

        /// <summary>
        /// Объем продаж февраль
        /// </summary>
        public double SalesVolume02
        {
            get; set;
        }

        /// <summary>
        /// Объем продаж март
        /// </summary>
        public double SalesVolume03
        {
            get; set;
        }

        /// <summary>
        /// Объем продаж апрель
        /// </summary>
        public double SalesVolume04
        {
            get; set;
        }

        /// <summary>
        /// Объем продаж май
        /// </summary>
        public double SalesVolume05
        {
            get; set;
        }

        /// <summary>
        /// Объем продаж июнь
        /// </summary>
        public double SalesVolume06
        {
            get; set;
        }

        /// <summary>
        /// Объем продаж июль
        /// </summary>
        public double SalesVolume07
        {
            get; set;
        }

        /// <summary>
        /// Объем продаж август
        /// </summary>
        public double SalesVolume08
        {
            get; set;
        }

        /// <summary>
        /// Объем продаж сентябрь
        /// </summary>
        public double SalesVolume09
        {
            get; set;
        }

        /// <summary>
        /// Объем продаж октябрь
        /// </summary>
        public double SalesVolume10
        {
            get; set;
        }

        /// <summary>
        /// Объем продаж ноябрь
        /// </summary>
        public double SalesVolume11
        {
            get; set;
        }

        /// <summary>
        /// Объем продаж декабрь
        /// </summary>
        public double SalesVolume12
        {
            get; set;
        }
        #endregion

        #region GS

        /// <summary>
        /// GS январь
        /// </summary>
        public double GS01
        {
            get; set;
        }

        /// <summary>
        /// GS февраль
        /// </summary>
        public double GS02
        {
            get; set;
        }

        /// <summary>
        /// GS март
        /// </summary>
        public double GS03
        {
            get; set;
        }

        /// <summary>
        /// GS апрель
        /// </summary>
        public double GS04
        {
            get; set;
        }

        /// <summary>
        /// GS май
        /// </summary>
        public double GS05
        {
            get; set;
        }

        /// <summary>
        /// GS июнь
        /// </summary>
        public double GS06
        {
            get; set;
        }

        /// <summary>
        /// GS июль
        /// </summary>
        public double GS07
        {
            get; set;
        }

        /// <summary>
        /// GS август
        /// </summary>
        public double GS08
        {
            get; set;
        }

        /// <summary>
        /// GS сентябрь
        /// </summary>
        public double GS09
        {
            get; set;
        }

        /// <summary>
        /// GS октябрь
        /// </summary>
        public double GS10
        {
            get; set;
        }

        /// <summary>
        /// GS ноябрь
        /// </summary>
        public double GS11
        {
            get; set;
        }

        /// <summary>
        /// GS декабрь
        /// </summary>
        public double GS12
        {
            get; set;
        }
        #endregion
        
        #region NS

        /// <summary>
        /// NS январь
        /// </summary>
        public double NS01
        {
            get; set;
        }

        /// <summary>
        /// NS февраль
        /// </summary>
        public double NS02
        {
            get; set;
        }

        /// <summary>
        /// NS март
        /// </summary>
        public double NS03
        {
            get; set;
        }

        /// <summary>
        /// NS апрель
        /// </summary>
        public double NS04
        {
            get; set;
        }

        /// <summary>
        /// NS май
        /// </summary>
        public double NS05
        {
            get; set;
        }

        /// <summary>
        /// NS июнь
        /// </summary>
        public double NS06
        {
            get; set;
        }

        /// <summary>
        /// NS июль
        /// </summary>
        public double NS07
        {
            get; set;
        }

        /// <summary>
        /// NS август
        /// </summary>
        public double NS08
        {
            get; set;
        }

        /// <summary>
        /// NS сентябрь
        /// </summary>
        public double NS09
        {
            get; set;
        }

        /// <summary>
        /// NS октябрь
        /// </summary>
        public double NS10
        {
            get; set;
        }

        /// <summary>
        /// NS ноябрь
        /// </summary>
        public double NS11
        {
            get; set;
        }

        /// <summary>
        /// NS декабрь
        /// </summary>
        public double NS12
        {
            get; set;
        }
        #endregion

        #region Promo Price-list

        /// <summary>
        /// Promo Price-list январь
        /// </summary>
        public double PromoPriceList01
        {
            get; set;
        }

        /// <summary>
        /// Promo Price-list февраль
        /// </summary>
        public double PromoPriceList02
        {
            get; set;
        }

        /// <summary>
        /// Promo Price-list март
        /// </summary>
        public double PromoPriceList03
        {
            get; set;
        }

        /// <summary>
        /// Promo Price-list апрель
        /// </summary>
        public double PromoPriceList04
        {
            get; set;
        }

        /// <summary>
        /// Promo Price-list май
        /// </summary>
        public double PromoPriceList05
        {
            get; set;
        }

        /// <summary>
        /// Promo Price-list июнь
        /// </summary>
        public double PromoPriceList06
        {
            get; set;
        }

        /// <summary>
        /// Promo Price-list июль
        /// </summary>
        public double PromoPriceList07
        {
            get; set;
        }

        /// <summary>
        /// Promo Price-list август
        /// </summary>
        public double PromoPriceList08
        {
            get; set;
        }

        /// <summary>
        /// Promo Price-list сентябрь
        /// </summary>
        public double PromoPriceList09
        {
            get; set;
        }

        /// <summary>
        /// Promo Price-list октябрь
        /// </summary>
        public double PromoPriceList10
        {
            get; set;
        }

        /// <summary>
        /// Promo Price-list ноябрь
        /// </summary>
        public double PromoPriceList11
        {
            get; set;
        }

        /// <summary>
        /// Promo Price-list декабрь
        /// </summary>
        public double PromoPriceList12
        {
            get; set;
        }
        #endregion

        #region Promo Объем продаж

        /// <summary>
        /// Promo Promo Объем продаж январь
        /// </summary>
        public double PromoSalesVolume01
        {
            get; set;
        }

        /// <summary>
        /// Promo Promo Объем продаж февраль
        /// </summary>
        public double PromoSalesVolume02
        {
            get; set;
        }

        /// <summary>
        /// Promo Promo Объем продаж март
        /// </summary>
        public double PromoSalesVolume03
        {
            get; set;
        }

        /// <summary>
        /// Promo Promo Объем продаж апрель
        /// </summary>
        public double PromoSalesVolume04
        {
            get; set;
        }

        /// <summary>
        /// Promo Promo Объем продаж май
        /// </summary>
        public double PromoSalesVolume05
        {
            get; set;
        }

        /// <summary>
        /// Promo Promo Объем продаж июнь
        /// </summary>
        public double PromoSalesVolume06
        {
            get; set;
        }

        /// <summary>
        /// Promo Promo Объем продаж июль
        /// </summary>
        public double PromoSalesVolume07
        {
            get; set;
        }

        /// <summary>
        /// Promo Promo Объем продаж август
        /// </summary>
        public double PromoSalesVolume08
        {
            get; set;
        }

        /// <summary>
        /// Promo Promo Объем продаж сентябрь
        /// </summary>
        public double PromoSalesVolume09
        {
            get; set;
        }

        /// <summary>
        /// Promo Promo Объем продаж октябрь
        /// </summary>
        public double PromoSalesVolume10
        {
            get; set;
        }

        /// <summary>
        /// Promo Promo Объем продаж ноябрь
        /// </summary>
        public double PromoSalesVolume11
        {
            get; set;
        }

        /// <summary>
        /// Promo Promo Объем продаж декабрь
        /// </summary>
        public double PromoSalesVolume12
        {
            get; set;
        }
        #endregion

        #region Promo GS

        /// <summary>
        /// Promo GS январь
        /// </summary>
        public double PromoGS01
        {
            get; set;
        }

        /// <summary>
        /// Promo GS февраль
        /// </summary>
        public double PromoGS02
        {
            get; set;
        }

        /// <summary>
        /// Promo GS март
        /// </summary>
        public double PromoGS03
        {
            get; set;
        }

        /// <summary>
        /// Promo GS апрель
        /// </summary>
        public double PromoGS04
        {
            get; set;
        }

        /// <summary>
        /// Promo GS май
        /// </summary>
        public double PromoGS05
        {
            get; set;
        }

        /// <summary>
        /// Promo GS июнь
        /// </summary>
        public double PromoGS06
        {
            get; set;
        }

        /// <summary>
        /// Promo GS июль
        /// </summary>
        public double PromoGS07
        {
            get; set;
        }

        /// <summary>
        /// Promo GS август
        /// </summary>
        public double PromoGS08
        {
            get; set;
        }

        /// <summary>
        /// Promo GS сентябрь
        /// </summary>
        public double PromoGS09
        {
            get; set;
        }

        /// <summary>
        /// Promo GS октябрь
        /// </summary>
        public double PromoGS10
        {
            get; set;
        }

        /// <summary>
        /// Promo GS ноябрь
        /// </summary>
        public double PromoGS11
        {
            get; set;
        }

        /// <summary>
        /// Promo GS декабрь
        /// </summary>
        public double PromoGS12
        {
            get; set;
        }
        #endregion
        
        #region Promo NS

        /// <summary>
        /// Promo NS январь
        /// </summary>
        public double PromoNS01
        {
            get; set;
        }

        /// <summary>
        /// Promo NS февраль
        /// </summary>
        public double PromoNS02
        {
            get; set;
        }

        /// <summary>
        /// Promo NS март
        /// </summary>
        public double PromoNS03
        {
            get; set;
        }

        /// <summary>
        /// Promo NS апрель
        /// </summary>
        public double PromoNS04
        {
            get; set;
        }

        /// <summary>
        /// Promo NS май
        /// </summary>
        public double PromoNS05
        {
            get; set;
        }

        /// <summary>
        /// Promo NS июнь
        /// </summary>
        public double PromoNS06
        {
            get; set;
        }

        /// <summary>
        /// Promo NS июль
        /// </summary>
        public double PromoNS07
        {
            get; set;
        }

        /// <summary>
        /// Promo NS август
        /// </summary>
        public double PromoNS08
        {
            get; set;
        }

        /// <summary>
        /// Promo NS сентябрь
        /// </summary>
        public double PromoNS09
        {
            get; set;
        }

        /// <summary>
        /// Promo NS октябрь
        /// </summary>
        public double PromoNS10
        {
            get; set;
        }

        /// <summary>
        /// Promo NS ноябрь
        /// </summary>
        public double PromoNS11
        {
            get; set;
        }

        /// <summary>
        /// Promo NS декабрь
        /// </summary>
        public double PromoNS12
        {
            get; set;
        }
        #endregion

        #region GP Value
        /// <summary>
        /// GP Value январь 
        /// </summary>
        public double GPValue01
        {
            get; set;
        }
        /// <summary>
        /// GP Value февраль 
        /// </summary>
        public double GPValue02
        {
            get; set;
        }
        /// <summary>
        /// GP Value март 
        /// </summary>
        public double GPValue03
        {
            get; set;
        }
        /// <summary>
        /// GP Value апрель 
        /// </summary>
        public double GPValue04
        {
            get; set;
        }
        /// <summary>
        /// GP Value май 
        /// </summary>
        public double GPValue05
        {
            get; set;
        }
        /// <summary>
        /// GP Value июнь 
        /// </summary>
        public double GPValue06
        {
            get; set;
        }
        /// <summary>
        /// GP Value июль 
        /// </summary>
        public double GPValue07
        {
            get; set;
        }
        /// <summary>
        /// GP Value август 
        /// </summary>
        public double GPValue08
        {
            get; set;
        }
        /// <summary>
        /// GP Value сентябрь 
        /// </summary>
        public double GPValue09
        {
            get; set;
        }
        /// <summary>
        /// GP Value октябрь 
        /// </summary>
        public double GPValue10
        {
            get; set;
        }
        /// <summary>
        /// GP Value ноябрь 
        /// </summary>
        public double GPValue11
        {
            get; set;
        }
        /// <summary>
        /// GP Value декабрь 
        /// </summary>
        public double GPValue12
        {
            get; set;
        }

        #endregion

        #region свойства
        /// <summary>
        /// Channel type
        /// </summary>
        public string ChannelType
        {
            get
            {
                return this.planningNewYear.ChannelType;
            }
            set
            {
            }
        }

        /// <summary>
        /// CustomerStatus
        /// </summary>
        public string CustomerStatus
        {
            get
            {
                return this.planningNewYear.CustomerStatus;
            }
            set
            {
            }
        }

        /// <summary>
        /// Дата прогноза
        /// </summary>
        public DateTime PrognosisDate
        {
            get 
            {
                return new DateTime(this.planningNewYear.Year,1,1);
            }
            set
            {
            }
        }
        #endregion

        public void SetValues()
        {
            SetProperty(Id);
        }
    }
}

