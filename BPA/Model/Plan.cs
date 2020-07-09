using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BPA.Model {
    /// <summary>
    /// Справочник планирования
    /// </summary>
    class Plan : TableBase {
        public override string TableName => "Планирование";
        public override string SheetName => "Планирование";

        public override IDictionary<string, string> Filds {
            get {
                return _filds;
            }
        }
        private readonly Dictionary<string, string> _filds = new Dictionary<string, string>
        {
            { "Id", "ID" },
            { "ChannelType", "Channel type" },
            { "CustomerStatus", "Customer status" },
            { "PrognosisDate", "Дата прогноза" },
            { "Data", "Данные справочника товаров" }, //?
            { "STKRub", "STK 2.5, руб." },
            { "IRPEur", "IRP, Eur" },
            { "RRC", "РРЦ, руб.с НДС" },
            { "IRPIndex", "Индекс IRP" },
            { "DIY", "DIY price list, руб. без НДС" },
            { "Price2Net", "2 net цена, руб." },
            { "Price3Net", "3 net цена, руб." },
            { "PriceTransfer", "Transfer цена, руб." },

            { "PriceList01", "Price-list январь" },
            { "PriceList02", "Price-list февраль" },
            { "PriceList03", "Price-list март" },
            { "PriceList04", "Price-list апрель" },
            { "PriceList05", "Price-list май" },
            { "PriceList06", "Price-list июнь" },
            { "PriceList07", "Price-list июль" },
            { "PriceList08", "Price-list август" },
            { "PriceList09", "Price-list сентябрь" },
            { "PriceList10", "Price-list октябрь" },
            { "PriceList11", "Price-list ноябрь" },
            { "PriceList12", "Price-list декабрь" },

            { "SalesVolume01", "Объем продаж январь" },
            { "SalesVolume02", "Объем продаж февраль" },
            { "SalesVolume03", "Объем продаж март" },
            { "SalesVolume04", "Объем продаж апрель" },
            { "SalesVolume05", "Объем продаж май" },
            { "SalesVolume06", "Объем продаж июнь" },
            { "SalesVolume07", "Объем продаж июль" },
            { "SalesVolume08", "Объем продаж август" },
            { "SalesVolume09", "Объем продаж сентябрь" },
            { "SalesVolume10", "Объем продаж октябрь" },
            { "SalesVolume11", "Объем продаж ноябрь" },
            { "SalesVolume12", "Объем продаж декабрь" },

            { "GS01", "GS январь" },
            { "GS02", "GS февраль" },
            { "GS03", "GS март" },
            { "GS04", "GS апрель" },
            { "GS05", "GS май" },
            { "GS06", "GS июнь" },
            { "GS07", "GS июль" },
            { "GS08", "GS август" },
            { "GS09", "GS сентябрь" },
            { "GS10", "GS октябрь" },
            { "GS11", "GS ноябрь" },
            { "GS12", "GS декабрь" },

            { "NS01", "NS январь" },
            { "NS02", "NS февраль" },
            { "NS03", "NS март" },
            { "NS04", "NS апрель" },
            { "NS05", "NS май" },
            { "NS06", "NS июнь" },
            { "NS07", "NS июль" },
            { "NS08", "NS август" },
            { "NS09", "NS сентябрь" },
            { "NS10", "NS октябрь" },
            { "NS11", "NS ноябрь" },
            { "NS12", "NS декабрь" },

            { "PromoPriceList01", "Promo Price-list январь" },
            { "PromoPriceList02", "Promo Price-list февраль" },
            { "PromoPriceList03", "Promo Price-list март" },
            { "PromoPriceList04", "Promo Price-list апрель" },
            { "PromoPriceList05", "Promo Price-list май" },
            { "PromoPriceList06", "Promo Price-list июнь" },
            { "PromoPriceList07", "Promo Price-list июль" },
            { "PromoPriceList08", "Promo Price-list август" },
            { "PromoPriceList09", "Promo Price-list сентябрь" },
            { "PromoPriceList10", "Promo Price-list октябрь" },
            { "PromoPriceList11", "Promo Price-list ноябрь" },
            { "PromoPriceList12", "Promo Price-list декабрь" },

            { "PromoSalesVolume01", "Promo Объем продаж январь" },
            { "PromoSalesVolume02", "Promo Объем февраль" },
            { "PromoSalesVolume03", "Promo Объем продаж март" },
            { "PromoSalesVolume04", "Promo Объем продаж апрель" },
            { "PromoSalesVolume05", "Promo Объем продаж май" },
            { "PromoSalesVolume06", "Promo Объем продаж июнь" },
            { "PromoSalesVolume07", "Promo Объем продаж июль" },
            { "PromoSalesVolume08", "Promo Объем продаж август" },
            { "PromoSalesVolume09", "Promo Объем продаж сентябрь" },
            { "PromoSalesVolume10", "Promo Объем продаж октябрь" },
            { "PromoSalesVolume11", "Promo Объем продаж ноябрь" },
            { "PromoSalesVolume12", "Promo Объем продаж декабрь" },

            { "PromoGS01", "Promo GS январь" },
            { "PromoGS02", "Promo GS февраль" },
            { "PromoGS03", "Promo GS март" },
            { "PromoGS04", "Promo GS апрель" },
            { "PromoGS05", "Promo GS май" },
            { "PromoGS06", "Promo GS июнь" },
            { "PromoGS07", "Promo GS июль" },
            { "PromoGS08", "Promo GS август" },
            { "PromoGS09", "Promo GS сентябрь" },
            { "PromoGS10", "Promo GS октябрь" },
            { "PromoGS11", "Promo GS ноябрь" },
            { "PromoGS12", "Promo GS декабрь" },

            { "PromoNS01", "Promo NS январь" },
            { "PromoNS02", "Promo NS февраль" },
            { "PromoNS03", "Promo NS март" },
            { "PromoNS04", "Promo NS апрель" },
            { "PromoNS05", "Promo NS май" },
            { "PromoNS06", "Promo NS июнь" },
            { "PromoNS07", "Promo NS июль" },
            { "PromoNS08", "Promo NS август" },
            { "PromoNS09", "Promo NS сентябрь" },
            { "PromoNS10", "Promo NS октябрь" },
            { "PromoNS11", "Promo NS ноябрь" },
            { "PromoNS12", "Promo NS декабрь" },

            { "GPValue01","GP Value январь" },
            { "GPValue02","GP Value февраль" },
            { "GPValue03","GP Value март" },
            { "GPValue04","GP Value апрель" },
            { "GPValue05","GP Value май" },
            { "GPValue06","GP Value июнь" },
            { "GPValue07","GP Value июль" },
            { "GPValue08","GP Value август" },
            { "GPValue09","GP Value сентябрь" },
            { "GPValue10","GP Value октябрь" },
            { "GPValue11","GP Value ноябрь" },
            { "GPValue12","GP Value декабрь" }
        };

        #region -- Основные свойства столбцов ---
        /// <summary>
        /// Id
        /// </summary>
        public int Id {
            get; set;
        }

        /// <summary>
        /// Channel type
        /// </summary>
        public string ChannelType {
            get; set;
        }

        /// <summary>
        /// CustomerStatus
        /// </summary>
        public string CustomerStatus {
            get; set;
        }

        /// <summary>
        /// Дата прогноза
        /// </summary>
        public DateTime PrognosisDate {
            get; set;
        }

        /// <summary>
        /// ??
        /// </summary>
        public string Data
        {
            get; set;
        }

        /// <summary>
        /// STK 2.5, руб.
        /// </summary>
        public double STKRub {
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
        public double RRC {
            get; set;
        }

        /// <summary>
        /// Индекс IRP
        /// </summary>
        public double IRPIndex {
            get; set;
        }

        /// <summary>
        /// DIY price list, руб. без НДС
        /// </summary>
        public double DIY {
            get; set;
        }

        /// <summary>
        /// 2 net цена, руб.
        /// </summary>
        public double Price2Net {
            get; set;
        }

        /// <summary>
        /// Зимние инструменты 2 net цена, руб.
        /// </summary>
        public double Price3Net {
            get; set;
        }

        /// <summary>
        /// Transfer цена, руб.
        /// </summary>
        public double PriceTransfer {
            get; set;
        }
        #endregion

        #region Price-list

        /// <summary>
        /// Price-list январь
        /// </summary>
        public double PriceList01 {
            get; set;
        }

        /// <summary>
        /// Price-list февраль
        /// </summary>
        public double PriceList02 {
            get; set;
        }

        /// <summary>
        /// Price-list март
        /// </summary>
        public double PriceList03 {
            get; set;
        }

        /// <summary>
        /// Price-list апрель
        /// </summary>
        public double PriceList04 {
            get; set;
        }

        /// <summary>
        /// Price-list май
        /// </summary>
        public double PriceList05 {
            get; set;
        }

        /// <summary>
        /// Price-list июнь
        /// </summary>
        public double PriceList06 {
            get; set;
        }

        /// <summary>
        /// Price-list июль
        /// </summary>
        public double PriceList07 {
            get; set;
        }

        /// <summary>
        /// Price-list август
        /// </summary>
        public double PriceList08 {
            get; set;
        }

        /// <summary>
        /// Price-list сентябрь
        /// </summary>
        public double PriceList09 {
            get; set;
        }

        /// <summary>
        /// Price-list октябрь
        /// </summary>
        public double PriceList10 {
            get; set;
        }

        /// <summary>
        /// Price-list ноябрь
        /// </summary>
        public double PriceList11 {
            get; set;
        }

        /// <summary>
        /// Price-list декабрь
        /// </summary>
        public double PriceList12 {
            get; set;
        }
        #endregion

        #region SalesVolume

        /// <summary>
        /// Объем продаж январь
        /// </summary>
        public double SalesVolume01 {
            get; set;
        }

        /// <summary>
        /// Объем продаж февраль
        /// </summary>
        public double SalesVolume02 {
            get; set;
        }

        /// <summary>
        /// Объем продаж март
        /// </summary>
        public double SalesVolume03 {
            get; set;
        }

        /// <summary>
        /// Объем продаж апрель
        /// </summary>
        public double SalesVolume04 {
            get; set;
        }

        /// <summary>
        /// Объем продаж май
        /// </summary>
        public double SalesVolume05 {
            get; set;
        }

        /// <summary>
        /// Объем продаж июнь
        /// </summary>
        public double SalesVolume06 {
            get; set;
        }

        /// <summary>
        /// Объем продаж июль
        /// </summary>
        public double SalesVolume07 {
            get; set;
        }

        /// <summary>
        /// Объем продаж август
        /// </summary>
        public double SalesVolume08 {
            get; set;
        }

        /// <summary>
        /// Объем продаж сентябрь
        /// </summary>
        public double SalesVolume09 {
            get; set;
        }

        /// <summary>
        /// Объем продаж октябрь
        /// </summary>
        public double SalesVolume10 {
            get; set;
        }

        /// <summary>
        /// Объем продаж ноябрь
        /// </summary>
        public double SalesVolume11 {
            get; set;
        }

        /// <summary>
        /// Объем продаж декабрь
        /// </summary>
        public double SalesVolume12 {
            get; set;
        }
        #endregion

        #region GS

        /// <summary>
        /// GS январь
        /// </summary>
        public double GS01 {
            get; set;
        }

        /// <summary>
        /// GS февраль
        /// </summary>
        public double GS02 {
            get; set;
        }

        /// <summary>
        /// GS март
        /// </summary>
        public double GS03 {
            get; set;
        }

        /// <summary>
        /// GS апрель
        /// </summary>
        public double GS04 {
            get; set;
        }

        /// <summary>
        /// GS май
        /// </summary>
        public double GS05 {
            get; set;
        }

        /// <summary>
        /// GS июнь
        /// </summary>
        public double GS06 {
            get; set;
        }

        /// <summary>
        /// GS июль
        /// </summary>
        public double GS07 {
            get; set;
        }

        /// <summary>
        /// GS август
        /// </summary>
        public double GS08 {
            get; set;
        }

        /// <summary>
        /// GS сентябрь
        /// </summary>
        public double GS09 {
            get; set;
        }

        /// <summary>
        /// GS октябрь
        /// </summary>
        public double GS10 {
            get; set;
        }

        /// <summary>
        /// GS ноябрь
        /// </summary>
        public double GS11 {
            get; set;
        }

        /// <summary>
        /// GS декабрь
        /// </summary>
        public double GS12 {
            get; set;
        }
        #endregion

        #region NS

        /// <summary>
        /// NS январь
        /// </summary>
        public double NS01 {
            get; set;
        }

        /// <summary>
        /// NS февраль
        /// </summary>
        public double NS02 {
            get; set;
        }

        /// <summary>
        /// NS март
        /// </summary>
        public double NS03 {
            get; set;
        }

        /// <summary>
        /// NS апрель
        /// </summary>
        public double NS04 {
            get; set;
        }

        /// <summary>
        /// NS май
        /// </summary>
        public double NS05 {
            get; set;
        }

        /// <summary>
        /// NS июнь
        /// </summary>
        public double NS06 {
            get; set;
        }

        /// <summary>
        /// NS июль
        /// </summary>
        public double NS07 {
            get; set;
        }

        /// <summary>
        /// NS август
        /// </summary>
        public double NS08 {
            get; set;
        }

        /// <summary>
        /// NS сентябрь
        /// </summary>
        public double NS09 {
            get; set;
        }

        /// <summary>
        /// NS октябрь
        /// </summary>
        public double NS10 {
            get; set;
        }

        /// <summary>
        /// NS ноябрь
        /// </summary>
        public double NS11 {
            get; set;
        }

        /// <summary>
        /// NS декабрь
        /// </summary>
        public double NS12 {
            get; set;
        }
        #endregion

        #region Promo Price-list

        /// <summary>
        /// Promo Price-list январь
        /// </summary>
        public double PromoPriceList01 {
            get; set;
        }

        /// <summary>
        /// Promo Price-list февраль
        /// </summary>
        public double PromoPriceList02 {
            get; set;
        }

        /// <summary>
        /// Promo Price-list март
        /// </summary>
        public double PromoPriceList03 {
            get; set;
        }

        /// <summary>
        /// Promo Price-list апрель
        /// </summary>
        public double PromoPriceList04 {
            get; set;
        }

        /// <summary>
        /// Promo Price-list май
        /// </summary>
        public double PromoPriceList05 {
            get; set;
        }

        /// <summary>
        /// Promo Price-list июнь
        /// </summary>
        public double PromoPriceList06 {
            get; set;
        }

        /// <summary>
        /// Promo Price-list июль
        /// </summary>
        public double PromoPriceList07 {
            get; set;
        }

        /// <summary>
        /// Promo Price-list август
        /// </summary>
        public double PromoPriceList08 {
            get; set;
        }

        /// <summary>
        /// Promo Price-list сентябрь
        /// </summary>
        public double PromoPriceList09 {
            get; set;
        }

        /// <summary>
        /// Promo Price-list октябрь
        /// </summary>
        public double PromoPriceList10 {
            get; set;
        }

        /// <summary>
        /// Promo Price-list ноябрь
        /// </summary>
        public double PromoPriceList11 {
            get; set;
        }

        /// <summary>
        /// Promo Price-list декабрь
        /// </summary>
        public double PromoPriceList12 {
            get; set;
        }
        #endregion

        #region Promo Объем продаж

        /// <summary>
        /// Promo Promo Объем продаж январь
        /// </summary>
        public double PromoSalesVolume01 {
            get; set;
        }

        /// <summary>
        /// Promo Promo Объем продаж февраль
        /// </summary>
        public double PromoSalesVolume02 {
            get; set;
        }

        /// <summary>
        /// Promo Promo Объем продаж март
        /// </summary>
        public double PromoSalesVolume03 {
            get; set;
        }

        /// <summary>
        /// Promo Promo Объем продаж апрель
        /// </summary>
        public double PromoSalesVolume04 {
            get; set;
        }

        /// <summary>
        /// Promo Promo Объем продаж май
        /// </summary>
        public double PromoSalesVolume05 {
            get; set;
        }

        /// <summary>
        /// Promo Promo Объем продаж июнь
        /// </summary>
        public double PromoSalesVolume06 {
            get; set;
        }

        /// <summary>
        /// Promo Promo Объем продаж июль
        /// </summary>
        public double PromoSalesVolume07 {
            get; set;
        }

        /// <summary>
        /// Promo Promo Объем продаж август
        /// </summary>
        public double PromoSalesVolume08 {
            get; set;
        }

        /// <summary>
        /// Promo Promo Объем продаж сентябрь
        /// </summary>
        public double PromoSalesVolume09 {
            get; set;
        }

        /// <summary>
        /// Promo Promo Объем продаж октябрь
        /// </summary>
        public double PromoSalesVolume10 {
            get; set;
        }

        /// <summary>
        /// Promo Promo Объем продаж ноябрь
        /// </summary>
        public double PromoSalesVolume11 {
            get; set;
        }

        /// <summary>
        /// Promo Promo Объем продаж декабрь
        /// </summary>
        public double PromoSalesVolume12 {
            get; set;
        }
        #endregion

        #region Promo GS

        /// <summary>
        /// Promo GS январь
        /// </summary>
        public double PromoGS01 {
            get; set;
        }

        /// <summary>
        /// Promo GS февраль
        /// </summary>
        public double PromoGS02 {
            get; set;
        }

        /// <summary>
        /// Promo GS март
        /// </summary>
        public double PromoGS03 {
            get; set;
        }

        /// <summary>
        /// Promo GS апрель
        /// </summary>
        public double PromoGS04 {
            get; set;
        }

        /// <summary>
        /// Promo GS май
        /// </summary>
        public double PromoGS05 {
            get; set;
        }

        /// <summary>
        /// Promo GS июнь
        /// </summary>
        public double PromoGS06 {
            get; set;
        }

        /// <summary>
        /// Promo GS июль
        /// </summary>
        public double PromoGS07 {
            get; set;
        }

        /// <summary>
        /// Promo GS август
        /// </summary>
        public double PromoGS08 {
            get; set;
        }

        /// <summary>
        /// Promo GS сентябрь
        /// </summary>
        public double PromoGS09 {
            get; set;
        }

        /// <summary>
        /// Promo GS октябрь
        /// </summary>
        public double PromoGS10 {
            get; set;
        }

        /// <summary>
        /// Promo GS ноябрь
        /// </summary>
        public double PromoGS11 {
            get; set;
        }

        /// <summary>
        /// Promo GS декабрь
        /// </summary>
        public double PromoGS12 {
            get; set;
        }
        #endregion
        
        #region Promo NS

        /// <summary>
        /// Promo NS январь
        /// </summary>
        public double PromoNS01 {
            get; set;
        }

        /// <summary>
        /// Promo NS февраль
        /// </summary>
        public double PromoNS02 {
            get; set;
        }

        /// <summary>
        /// Promo NS март
        /// </summary>
        public double PromoNS03 {
            get; set;
        }

        /// <summary>
        /// Promo NS апрель
        /// </summary>
        public double PromoNS04 {
            get; set;
        }

        /// <summary>
        /// Promo NS май
        /// </summary>
        public double PromoNS05 {
            get; set;
        }

        /// <summary>
        /// Promo NS июнь
        /// </summary>
        public double PromoNS06 {
            get; set;
        }

        /// <summary>
        /// Promo NS июль
        /// </summary>
        public double PromoNS07 {
            get; set;
        }

        /// <summary>
        /// Promo NS август
        /// </summary>
        public double PromoNS08 {
            get; set;
        }

        /// <summary>
        /// Promo NS сентябрь
        /// </summary>
        public double PromoNS09 {
            get; set;
        }

        /// <summary>
        /// Promo NS октябрь
        /// </summary>
        public double PromoNS10 {
            get; set;
        }

        /// <summary>
        /// Promo NS ноябрь
        /// </summary>
        public double PromoNS11 {
            get; set;
        }

        /// <summary>
        /// Promo NS декабрь
        /// </summary>
        public double PromoNS12 {
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
    }
}
