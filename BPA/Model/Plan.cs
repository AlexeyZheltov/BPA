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

            { "PromoSalesVolume01", "Promo Promo Объем продаж январь" },
            { "PromoSalesVolume02", "Promo Promo Объем продаж февраль" },
            { "PromoSalesVolume03", "Promo Promo Объем продаж март" },
            { "PromoSalesVolume04", "Promo Promo Объем продаж апрель" },
            { "PromoSalesVolume05", "Promo Promo Объем продаж май" },
            { "PromoSalesVolume06", "Promo Promo Объем продаж июнь" },
            { "PromoSalesVolume07", "Promo Promo Объем продаж июль" },
            { "PromoSalesVolume08", "Promo Promo Объем продаж август" },
            { "PromoSalesVolume09", "Promo Promo Объем продаж сентябрь" },
            { "PromoSalesVolume10", "Promo Promo Объем продаж октябрь" },
            { "PromoSalesVolume11", "Promo Promo Объем продаж ноябрь" },
            { "PromoSalesVolume12", "Promo Promo Объем продаж декабрь" },

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

        };


        /// <summary>
        /// Id
        /// </summary>
        public string Id {
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
        public int CustomerStatus {
            get; set;
        }

        /// <summary>
        /// Дата прогноза
        /// </summary>
        public string PrognosisDate {
            get; set;
        }

        /// <summary>
        /// STK 2.5, руб.
        /// </summary>
        public string STKRub {
            get; set;
        }

        /// <summary>
        /// РРЦ, руб.с НДС
        /// </summary>
        public string RRC {
            get; set;
        }

        /// <summary>
        /// Индекс IRP
        /// </summary>
        public string IRPIndex {
            get; set;
        }

        /// <summary>
        /// DIY price list, руб. без НДС
        /// </summary>
        public string DIY {
            get; set;
        }

        /// <summary>
        /// 2 net цена, руб.
        /// </summary>
        public string Price2Net {
            get; set;
        }

        /// <summary>
        /// Зимние инструменты 2 net цена, руб.
        /// </summary>
        public string Price3Net {
            get; set;
        }

        /// <summary>
        /// Transfer цена, руб.
        /// </summary>
        public string PriceTransfer {
            get; set;
        }

#region Price-list

        /// <summary>
        /// Price-list январь
        /// </summary>
        public string PriceList01 {
            get; set;
        }

        /// <summary>
        /// Price-list февраль
        /// </summary>
        public string PriceList02 {
            get; set;
        }

        /// <summary>
        /// Price-list март
        /// </summary>
        public string PriceList03 {
            get; set;
        }

        /// <summary>
        /// Price-list апрель
        /// </summary>
        public string PriceList04 {
            get; set;
        }

        /// <summary>
        /// Price-list май
        /// </summary>
        public string PriceList05 {
            get; set;
        }

        /// <summary>
        /// Price-list июнь
        /// </summary>
        public string PriceList06 {
            get; set;
        }

        /// <summary>
        /// Price-list июль
        /// </summary>
        public string PriceList07 {
            get; set;
        }

        /// <summary>
        /// Price-list август
        /// </summary>
        public string PriceList08 {
            get; set;
        }

        /// <summary>
        /// Price-list сентябрь
        /// </summary>
        public string PriceList09 {
            get; set;
        }

        /// <summary>
        /// Price-list октябрь
        /// </summary>
        public string PriceList10 {
            get; set;
        }

        /// <summary>
        /// Price-list ноябрь
        /// </summary>
        public string PriceList11 {
            get; set;
        }

        /// <summary>
        /// Price-list декабрь
        /// </summary>
        public string PriceList12 {
            get; set;
        }
        #endregion


#region SalesVolume

        /// <summary>
        /// Объем продаж январь
        /// </summary>
        public string SalesVolume01 {
            get; set;
        }

        /// <summary>
        /// Объем продаж февраль
        /// </summary>
        public string SalesVolume02 {
            get; set;
        }

        /// <summary>
        /// Объем продаж март
        /// </summary>
        public string SalesVolume03 {
            get; set;
        }

        /// <summary>
        /// Объем продаж апрель
        /// </summary>
        public string SalesVolume04 {
            get; set;
        }

        /// <summary>
        /// Объем продаж май
        /// </summary>
        public string SalesVolume05 {
            get; set;
        }

        /// <summary>
        /// Объем продаж июнь
        /// </summary>
        public string SalesVolume06 {
            get; set;
        }

        /// <summary>
        /// Объем продаж июль
        /// </summary>
        public string SalesVolume07 {
            get; set;
        }

        /// <summary>
        /// Объем продаж август
        /// </summary>
        public string SalesVolume08 {
            get; set;
        }

        /// <summary>
        /// Объем продаж сентябрь
        /// </summary>
        public string SalesVolume09 {
            get; set;
        }

        /// <summary>
        /// Объем продаж октябрь
        /// </summary>
        public string SalesVolume10 {
            get; set;
        }

        /// <summary>
        /// Объем продаж ноябрь
        /// </summary>
        public string SalesVolume11 {
            get; set;
        }

        /// <summary>
        /// Объем продаж декабрь
        /// </summary>
        public string SalesVolume12 {
            get; set;
        }
        #endregion


 #region GS

        /// <summary>
        /// GS январь
        /// </summary>
        public string GS01 {
            get; set;
        }

        /// <summary>
        /// GS февраль
        /// </summary>
        public string GS02 {
            get; set;
        }

        /// <summary>
        /// GS март
        /// </summary>
        public string GS03 {
            get; set;
        }

        /// <summary>
        /// GS апрель
        /// </summary>
        public string GS04 {
            get; set;
        }

        /// <summary>
        /// GS май
        /// </summary>
        public string GS05 {
            get; set;
        }

        /// <summary>
        /// GS июнь
        /// </summary>
        public string GS06 {
            get; set;
        }

        /// <summary>
        /// GS июль
        /// </summary>
        public string GS07 {
            get; set;
        }

        /// <summary>
        /// GS август
        /// </summary>
        public string GS08 {
            get; set;
        }

        /// <summary>
        /// GS сентябрь
        /// </summary>
        public string GS09 {
            get; set;
        }

        /// <summary>
        /// GS октябрь
        /// </summary>
        public string GS10 {
            get; set;
        }

        /// <summary>
        /// GS ноябрь
        /// </summary>
        public string GS11 {
            get; set;
        }

        /// <summary>
        /// GS декабрь
        /// </summary>
        public string GS12 {
            get; set;
        }
        #endregion


 #region NS

        /// <summary>
        /// NS январь
        /// </summary>
        public string NS01 {
            get; set;
        }

        /// <summary>
        /// NS февраль
        /// </summary>
        public string NS02 {
            get; set;
        }

        /// <summary>
        /// NS март
        /// </summary>
        public string NS03 {
            get; set;
        }

        /// <summary>
        /// NS апрель
        /// </summary>
        public string NS04 {
            get; set;
        }

        /// <summary>
        /// NS май
        /// </summary>
        public string NS05 {
            get; set;
        }

        /// <summary>
        /// NS июнь
        /// </summary>
        public string NS06 {
            get; set;
        }

        /// <summary>
        /// NS июль
        /// </summary>
        public string NS07 {
            get; set;
        }

        /// <summary>
        /// NS август
        /// </summary>
        public string NS08 {
            get; set;
        }

        /// <summary>
        /// NS сентябрь
        /// </summary>
        public string NS09 {
            get; set;
        }

        /// <summary>
        /// NS октябрь
        /// </summary>
        public string NS10 {
            get; set;
        }

        /// <summary>
        /// NS ноябрь
        /// </summary>
        public string NS11 {
            get; set;
        }

        /// <summary>
        /// NS декабрь
        /// </summary>
        public string NS12 {
            get; set;
        }
        #endregion


#region Promo Price-list

        /// <summary>
        /// Promo Price-list январь
        /// </summary>
        public string PromoPriceList01 {
            get; set;
        }

        /// <summary>
        /// Promo Price-list февраль
        /// </summary>
        public string PromoPriceList02 {
            get; set;
        }

        /// <summary>
        /// Promo Price-list март
        /// </summary>
        public string PromoPriceList03 {
            get; set;
        }

        /// <summary>
        /// Promo Price-list апрель
        /// </summary>
        public string PromoPriceList04 {
            get; set;
        }

        /// <summary>
        /// Promo Price-list май
        /// </summary>
        public string PromoPriceList05 {
            get; set;
        }

        /// <summary>
        /// Promo Price-list июнь
        /// </summary>
        public string PromoPriceList06 {
            get; set;
        }

        /// <summary>
        /// Promo Price-list июль
        /// </summary>
        public string PromoPriceList07 {
            get; set;
        }

        /// <summary>
        /// Promo Price-list август
        /// </summary>
        public string PromoPriceList08 {
            get; set;
        }

        /// <summary>
        /// Promo Price-list сентябрь
        /// </summary>
        public string PromoPriceList09 {
            get; set;
        }

        /// <summary>
        /// Promo Price-list октябрь
        /// </summary>
        public string PromoPriceList10 {
            get; set;
        }

        /// <summary>
        /// Promo Price-list ноябрь
        /// </summary>
        public string PromoPriceList11 {
            get; set;
        }

        /// <summary>
        /// Promo Price-list декабрь
        /// </summary>
        public string PromoPriceList12 {
            get; set;
        }
        #endregion


#region Promo Объем продаж

        /// <summary>
        /// Promo Promo Объем продаж январь
        /// </summary>
        public string PromoSalesVolume01 {
            get; set;
        }

        /// <summary>
        /// Promo Promo Объем продаж февраль
        /// </summary>
        public string PromoSalesVolume02 {
            get; set;
        }

        /// <summary>
        /// Promo Promo Объем продаж март
        /// </summary>
        public string PromoSalesVolume03 {
            get; set;
        }

        /// <summary>
        /// Promo Promo Объем продаж апрель
        /// </summary>
        public string PromoSalesVolume04 {
            get; set;
        }

        /// <summary>
        /// Promo Promo Объем продаж май
        /// </summary>
        public string PromoSalesVolume05 {
            get; set;
        }

        /// <summary>
        /// Promo Promo Объем продаж июнь
        /// </summary>
        public string PromoSalesVolume06 {
            get; set;
        }

        /// <summary>
        /// Promo Promo Объем продаж июль
        /// </summary>
        public string PromoSalesVolume07 {
            get; set;
        }

        /// <summary>
        /// Promo Promo Объем продаж август
        /// </summary>
        public string PromoSalesVolume08 {
            get; set;
        }

        /// <summary>
        /// Promo Promo Объем продаж сентябрь
        /// </summary>
        public string PromoSalesVolume09 {
            get; set;
        }

        /// <summary>
        /// Promo Promo Объем продаж октябрь
        /// </summary>
        public string PromoSalesVolume10 {
            get; set;
        }

        /// <summary>
        /// Promo Promo Объем продаж ноябрь
        /// </summary>
        public string PromoSalesVolume11 {
            get; set;
        }

        /// <summary>
        /// Promo Promo Объем продаж декабрь
        /// </summary>
        public string PromoSalesVolume12 {
            get; set;
        }
        #endregion

#region Promo GS

        /// <summary>
        /// Promo GS январь
        /// </summary>
        public string PromoGS01 {
            get; set;
        }

        /// <summary>
        /// Promo GS февраль
        /// </summary>
        public string PromoGS02 {
            get; set;
        }

        /// <summary>
        /// Promo GS март
        /// </summary>
        public string PromoGS03 {
            get; set;
        }

        /// <summary>
        /// Promo GS апрель
        /// </summary>
        public string PromoGS04 {
            get; set;
        }

        /// <summary>
        /// Promo GS май
        /// </summary>
        public string PromoGS05 {
            get; set;
        }

        /// <summary>
        /// Promo GS июнь
        /// </summary>
        public string PromoGS06 {
            get; set;
        }

        /// <summary>
        /// Promo GS июль
        /// </summary>
        public string PromoGS07 {
            get; set;
        }

        /// <summary>
        /// Promo GS август
        /// </summary>
        public string PromoGS08 {
            get; set;
        }

        /// <summary>
        /// Promo GS сентябрь
        /// </summary>
        public string PromoGS09 {
            get; set;
        }

        /// <summary>
        /// Promo GS октябрь
        /// </summary>
        public string PromoGS10 {
            get; set;
        }

        /// <summary>
        /// Promo GS ноябрь
        /// </summary>
        public string PromoGS11 {
            get; set;
        }

        /// <summary>
        /// Promo GS декабрь
        /// </summary>
        public string PromoGS12 {
            get; set;
        }
        #endregion


#region Promo NS

        /// <summary>
        /// Promo NS январь
        /// </summary>
        public string PromoNS01 {
            get; set;
        }

        /// <summary>
        /// Promo NS февраль
        /// </summary>
        public string PromoNS02 {
            get; set;
        }

        /// <summary>
        /// Promo NS март
        /// </summary>
        public string PromoNS03 {
            get; set;
        }

        /// <summary>
        /// Promo NS апрель
        /// </summary>
        public string PromoNS04 {
            get; set;
        }

        /// <summary>
        /// Promo NS май
        /// </summary>
        public string PromoNS05 {
            get; set;
        }

        /// <summary>
        /// Promo NS июнь
        /// </summary>
        public string PromoNS06 {
            get; set;
        }

        /// <summary>
        /// Promo NS июль
        /// </summary>
        public string PromoNS07 {
            get; set;
        }

        /// <summary>
        /// Promo NS август
        /// </summary>
        public string PromoNS08 {
            get; set;
        }

        /// <summary>
        /// Promo NS сентябрь
        /// </summary>
        public string PromoNS09 {
            get; set;
        }

        /// <summary>
        /// Promo NS октябрь
        /// </summary>
        public string PromoNS10 {
            get; set;
        }

        /// <summary>
        /// Promo NS ноябрь
        /// </summary>
        public string PromoNS11 {
            get; set;
        }

        /// <summary>
        /// Promo NS декабрь
        /// </summary>
        public string PromoNS12 {
            get; set;
        }
        #endregion

    }
}
