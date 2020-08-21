using BPA.Modules;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BPA.NewModel
{
    class PlanningNewYearItem
    {
        TableRow _row;
        public PlanningNewYearItem(TableRow row) => _row = row;

        public string CustomerStatus;
        public string ChannelType;
        public double MaximumBonus;
        public DateTime CurrentDate;
        public DateTime planningDate;

        public int Id
        {
            get => _row["№"];
            set => _row["№"] = value;
        }
        //public string RRCNDS
        //{
        //    get => _row["РРЦ 2021, руб. с НДС"];
        //    set => _row["РРЦ 2021, руб. с НДС"] = value;
        //}
        public string DIYPriceList
        {
            get => _row["DIY цена 2021, руб. с НДС"];
            set => _row["DIY цена 2021, руб. с НДС"] = value;
        }
        public double STKRub
        {
            get => _row["STK 2.5 2021, RUB"];
            set => _row["STK 2.5 2021, RUB"] = value;
        }
        public double RRCCUrent
        {
            get => _row["РРЦ 2020, руб. с НДС" ];
            set => _row["РРЦ 2020, руб. с НДС"] = value;
        }
        public double RRCPlan
        {
            get => _row["РРЦ 2021, руб. с НДС"];
            set => _row["РРЦ 2021, руб. с НДС"] = value;
        }
        public double DIYCurrent
        {
            get => _row["DIY цена 2020, руб. с НДС"];
            set => _row["DIY цена 2020, руб. с НДС"] = value;
        }
        public double DIYPlan
        {
            get => _row["DIY цена 2021, руб. с НДС"];
            set => _row["DIY цена 2021, руб. с НДС"] = value;
        }

        //public double IRPCurrent
        //{
        //    get => _row[""];
        //    set => _row[""] = value;
        //}
        //public double IRPPlan
        //{
        //    get => _row[""];
        //    set => _row[""] = value;
        //}
        //public double IRPIndexCurrent
        //{
        //    get => _row[""];
        //    set => _row[""] = value;
        //}
        //public double IRPIndexPlan
        //{
        //    get => _row[""];
        //    set => _row[""] = value;
        //}
        //public double RRPCurrent
        //{
        //    get => _row[""];
        //    set => _row[""] = value;
        //}
        //public double RRPPlan
        //{
        //    get => _row[""];
        //    set => _row[""] = value;
        //}
        public string SupercategoryEng
        {
            get => _row["Суперкатегория"];
            set => _row["Суперкатегория"] = value;
        }
        public string Supercategory
        {
            get => _row["SuperCategory"];
            set => _row["SuperCategory"] = value;
        }
        public string ProductGroup
        {
            get => _row["Product group"];
            set => _row["Product group"] = value;
        }
        public string ProductGroupEng
        {
            get => _row["Product group name"];
            set => _row["Product group name"] = value;
        }
        public string SubGroup
        {
            get => _row["Subgroup"];
            set => _row["Subgroup"] = value;
        }
        public string GenericName
        {
            get => _row["Generic Name (long)"];
            set => _row["Generic Name (long)"] = value;
        }
        public string PNS
        {
            get => _row["PNS"];
            set => _row["PNS"] = value;
        }
        public string Article
        {
            get => _row["Article"];
            set => _row["Article"] = value;
        }
        public string ArticleOld
        {
            get => _row["Predessor - Local ID Gardena"];
            set => _row["Predessor - Local ID Gardena"] = value;
        }
        public string ArticleRu
        {
            get => _row["Description RUS"];
            set => _row["Description RUS"] = value;
        }
        public double CalendarSalesStartDate
        {
            get => _row["Sales Start Date"];
            set => _row["Sales Start Date"] = value;
        }
        public double CalendarPreliminaryEliminationDate
        {
            get => _row["Preliminary Elimination Date"];
            set => _row["Preliminary Elimination Date"] = value;
        }
        public double CalendarEliminationDate
        {
            get => _row["Elimination Date"];
            set => _row["Elimination Date"] = value;
        }
        public string Status
        {
            get => _row["status"];
            set => _row["status"] = value;
        }
        public double PriceListPalan
        {
            get => _row["Price list цена 2021, руб."];
            set => _row["Price list цена 2021, руб."] = value;
        }
        public double PriceListCurrentn
        {
            get => _row["Price list цена 2020, руб."];
            set => _row["Price list цена 2020, руб."] = value;
        }

        public double STKRubCurrent {
            get => _row["STK 2.5 2020, RUB"];
            set => _row["STK 2.5 2020, RUB"] = value;
        }
        public double STKRubPlan
        {
            get => _row["STK 2.5 2021, RUB"];
            set => _row["STK 2.5 2021, RUB"] = value;
        }


        #region Prognosis
        //public double QuantityPrognosisYear {
        //    get => _row["ИТОГО Прогноз за год, шт."];
        //    set => _row["ИТОГО Прогноз за год, шт."] = value;
        //}
        public double QuantityPrognosis01
        {
            get => _row["ИТОГО Прогноз январь, шт."];
            set => _row["ИТОГО Прогноз январь, шт."] = value;
        }
        public double QuantityPrognosis02
        {
            get => _row["ИТОГО Прогноз февраль, шт."];
            set => _row["ИТОГО Прогноз февраль, шт."] = value;
        }
        public double QuantityPrognosis03
        {
            get => _row["ИТОГО Прогноз март, шт."];
            set => _row["ИТОГО Прогноз март, шт."] = value;
        }
        public double QuantityPrognosis04
        {
            get => _row["ИТОГО Прогноз апрель, шт."];
            set => _row["ИТОГО Прогноз апрель, шт."] = value;
        }
        public double QuantityPrognosis05
        {
            get => _row["ИТОГО Прогноз май, шт."];
            set => _row["ИТОГО Прогноз май, шт."] = value;
        }
        public double QuantityPrognosis06
        {
            get => _row["ИТОГО Прогноз июнь, шт."];
            set => _row["ИТОГО Прогноз июнь, шт."] = value;
        }
        public double QuantityPrognosis07
        {
            get => _row["ИТОГО Прогноз июль, шт."];
            set => _row["ИТОГО Прогноз июль, шт."] = value;
        }
        public double QuantityPrognosis08
        {
            get => _row["ИТОГО Прогноз август, шт."];
            set => _row["ИТОГО Прогноз август, шт."] = value;
        }
        public double QuantityPrognosis09
        {
            get => _row["ИТОГО Прогноз сентябрь, шт."];
            set => _row["ИТОГО Прогноз сентябрь, шт."] = value;
        }
        public double QuantityPrognosis10
        {
            get => _row["ИТОГО Прогноз октябрь, шт."];
            set => _row["ИТОГО Прогноз октябрь, шт."] = value;
        }
        public double QuantityPrognosis11
        {
            get => _row["ИТОГО Прогноз ноябрь, шт."];
            set => _row["ИТОГО Прогноз ноябрь, шт."] = value;
        }
        public double QuantityPrognosis12
        {
            get => _row["ИТОГО Прогноз декабрь, шт."];
            set => _row["ИТОГО Прогноз декабрь, шт."] = value;
        }

        //public double GSPrognosisYear {
        //    get => _row["ИТОГО GS за год, шт."];
        //    set => _row["ИТОГО GS за год, шт."] = value;
        //}
        public double GSPrognosis01
        {
            get => _row["ИТОГО GS январь, шт."];
            set => _row["ИТОГО GS январь, шт."] = value;
        }
        public double GSPrognosis02
        {
            get => _row["ИТОГО GS февраль, шт."];
            set => _row["ИТОГО GS февраль, шт."] = value;
        }
        public double GSPrognosis03
        {
            get => _row["ИТОГО GS март, шт."];
            set => _row["ИТОГО GS март, шт."] = value;
        }
        public double GSPrognosis04
        {
            get => _row["ИТОГО GS апрель, шт."];
            set => _row["ИТОГО GS апрель, шт."] = value;
        }
        public double GSPrognosis05
        {
            get => _row["ИТОГО GS май, шт."];
            set => _row["ИТОГО GS май, шт."] = value;
        }
        public double GSPrognosis06
        {
            get => _row["ИТОГО GS июнь, шт."];
            set => _row["ИТОГО GS июнь, шт."] = value;
        }
        public double GSPrognosis07
        {
            get => _row["ИТОГО GS июль, шт."];
            set => _row["ИТОГО GS июль, шт."] = value;
        }
        public double GSPrognosis08
        {
            get => _row["ИТОГО GS август, шт."];
            set => _row["ИТОГО GS август, шт."] = value;
        }
        public double GSPrognosis09
        {
            get => _row["ИТОГО GS сентябрь, шт."];
            set => _row["ИТОГО GS сентябрь, шт."] = value;
        }
        public double GSPrognosis10
        {
            get => _row["ИТОГО GS октябрь, шт."];
            set => _row["ИТОГО GS октябрь, шт."] = value;
        }
        public double GSPrognosis11
        {
            get => _row["ИТОГО GS ноябрь, шт."];
            set => _row["ИТОГО GS ноябрь, шт."] = value;
        }
        public double GSPrognosis12
        {
            get => _row["ИТОГО GS декабрь, шт."];
            set => _row["ИТОГО GS декабрь, шт."] = value;
        }

        //public double NSPrognosisYear
        //{
        //    get => _row["ИТОГО NS за год, шт."];
        //    set => _row["ИТОГО NS за год, шт."] = value;
        //}
        public double NSPrognosis01
        {
            get => _row["ИТОГО NS январь, шт."];
            set => _row["ИТОГО NS январь, шт."] = value;
        }
        public double NSPrognosis02
        {
            get => _row["ИТОГО NS февраль, шт."];
            set => _row["ИТОГО NS февраль, шт."] = value;
        }
        public double NSPrognosis03
        {
            get => _row["ИТОГО NS март, шт."];
            set => _row["ИТОГО NS март, шт."] = value;
        }
        public double NSPrognosis04
        {
            get => _row["ИТОГО NS апрель, шт."];
            set => _row["ИТОГО NS апрель, шт."] = value;
        }
        public double NSPrognosis05
        {
            get => _row["ИТОГО NS май, шт."];
            set => _row["ИТОГО NS май, шт."] = value;
        }
        public double NSPrognosis06
        {
            get => _row["ИТОГО NS июнь, шт."];
            set => _row["ИТОГО NS июнь, шт."] = value;
        }
        public double NSPrognosis07
        {
            get => _row["ИТОГО NS июль, шт."];
            set => _row["ИТОГО NS июль, шт."] = value;
        }
        public double NSPrognosis08
        {
            get => _row["ИТОГО NS август, шт."];
            set => _row["ИТОГО NS август, шт."] = value;
        }
        public double NSPrognosis09
        {
            get => _row["ИТОГО NS сентябрь, шт."];
            set => _row["ИТОГО NS сентябрь, шт."] = value;
        }
        public double NSPrognosis10
        {
            get => _row["ИТОГО NS октябрь, шт."];
            set => _row["ИТОГО NS октябрь, шт."] = value;
        }
        public double NSPrognosis11
        {
            get => _row["ИТОГО NS ноябрь, шт."];
            set => _row["ИТОГО NS ноябрь, шт."] = value;
        }
        public double NSPrognosis12
        {
            get => _row["ИТОГО NS декабрь, шт."];
            set => _row["ИТОГО NS декабрь, шт."] = value;
        }
        #endregion

        #region 12 столбцов PriceList
        //public double PriceList01
        //{ 
        //    get => _row["январь"];
        //    set => _row["январь"] = value;
        //}
        //public double PriceList02
        //{ 
        //    get => _row["февраль"];
        //    set => _row["февраль"] = value;
        //}
        //public double PriceList03
        //{ 
        //    get => _row["март"];
        //    set => _row["март"] = value;
        //}
        //public double PriceList04
        //{ 
        //    get => _row["апрель"];
        //    set => _row["апрель"] = value;
        //}
        //public double PriceList05
        //{ 
        //    get => _row["май"];
        //    set => _row["май"] = value;
        //}
        //public double PriceList06
        //{ 
        //    get => _row["июнь"];
        //    set => _row["июнь"] = value;
        //}
        //public double PriceList07
        //{ 
        //    get => _row["июль"];
        //    set => _row["июль"] = value;
        //}
        //public double PriceList08
        //{ 
        //    get => _row["август"];
        //    set => _row["август"] = value;
        //}
        //public double PriceList09
        //{ 
        //    get => _row["сентябрь"];
        //    set => _row["сентябрь"] = value;
        //}
        //public double PriceList10
        //{ 
        //    get => _row["октябрь"];
        //    set => _row["октябрь"] = value;
        //}
        //public double PriceList11
        //{ 
        //    get => _row["ноябрь"];
        //    set => _row["ноябрь"] = value;
        //}
        //public double PriceList12
        //{ 
        //    get => _row["декабрь"];
        //    set => _row["декабрь"] = value;
        //}
        #endregion

        #region 12 столбцов Объем продаж
        //public double SalesVolume01
        //{ 
        //    get => _row["январь"];
        //    set => _row["январь"] = value;
        //}
        //public double SalesVolume02
        //{ 
        //    get => _row["февраль"];
        //    set => _row["февраль"] = value;
        //}
        //public double SalesVolume03
        //{ 
        //    get => _row["март"];
        //    set => _row["март"] = value;
        //}
        //public double SalesVolume04
        //{ 
        //    get => _row["апрель"];
        //    set => _row["апрель"] = value;
        //}
        //public double SalesVolume05
        //{ 
        //    get => _row["май"];
        //    set => _row["май"] = value;
        //}
        //public double SalesVolume06
        //{ 
        //    get => _row["июнь"];
        //    set => _row["июнь"] = value;
        //}
        //public double SalesVolume07
        //{ 
        //    get => _row["июль"];
        //    set => _row["июль"] = value;
        //}
        //public double SalesVolume08
        //{ 
        //    get => _row["август"];
        //    set => _row["август"] = value;
        //}
        //public double SalesVolume09
        //{ 
        //    get => _row["сентябрь"];
        //    set => _row["сентябрь"] = value;
        //}
        //public double SalesVolume10
        //{ 
        //    get => _row["октябрь"];
        //    set => _row["октябрь"] = value;
        //}
        //public double SalesVolume11
        //{ 
        //    get => _row["ноябрь"];
        //    set => _row["ноябрь"] = value;
        //}
        //public double SalesVolume12
        //{ 
        //    get => _row["декабрь"];
        //    set => _row["декабрь"] = value;
        //}
        #endregion

        # region 12 столбцов Promo PriceList
        //public double PromoPriceList01
        //{ 
        //    get => _row["январь"];
        //    set => _row["январь"] = value;
        //}
        //public double PromoPriceList02
        //{ 
        //    get => _row["февраль"];
        //    set => _row["февраль"] = value;
        //}
        //public double PromoPriceList03
        //{ 
        //    get => _row["март"];
        //    set => _row["март"] = value;
        //}
        //public double PromoPriceList04
        //{ 
        //    get => _row["апрель"];
        //    set => _row["апрель"] = value;
        //}
        //public double PromoPriceList05
        //{ 
        //    get => _row["май"];
        //    set => _row["май"] = value;
        //}
        //public double PromoPriceList06
        //{ 
        //    get => _row["июнь"];
        //    set => _row["июнь"] = value;
        //}
        //public double PromoPriceList07
        //{ 
        //    get => _row["июль"];
        //    set => _row["июль"] = value;
        //}
        //public double PromoPriceList08
        //{ 
        //    get => _row["август"];
        //    set => _row["август"] = value;
        //}
        //public double PromoPriceList09
        //{ 
        //    get => _row["сентябрь"];
        //    set => _row["сентябрь"] = value;
        //}
        //public double PromoPriceList10
        //{ 
        //    get => _row["октябрь"];
        //    set => _row["октябрь"] = value;
        //}
        //public double PromoPriceList11
        //{ 
        //    get => _row["ноябрь"];
        //    set => _row["ноябрь"] = value;
        //}
        //public double PromoPriceList12
        //{ 
        //    get => _row["декабрь"];
        //    set => _row["декабрь"] = value;
        //}
        #endregion

        #region 12 столбцов Promo Объем продаж
        //public double PromoSalesVolume01
        //{ 
        //    get => _row["январь"];
        //    set => _row["январь"] = value;
        //}
        //public double PromoSalesVolume02
        //{ 
        //    get => _row["февраль"];
        //    set => _row["февраль"] = value;
        //}
        //public double PromoSalesVolume03
        //{ 
        //    get => _row["март"];
        //    set => _row["март"] = value;
        //}
        //public double PromoSalesVolume04
        //{ 
        //    get => _row["апрель"];
        //    set => _row["апрель"] = value;
        //}
        //public double PromoSalesVolume05
        //{ 
        //    get => _row["май"];
        //    set => _row["май"] = value;
        //}
        //public double PromoSalesVolume06
        //{ 
        //    get => _row["июнь"];
        //    set => _row["июнь"] = value;
        //}
        //public double PromoSalesVolume07
        //{ 
        //    get => _row["июль"];
        //    set => _row["июль"] = value;
        //}
        //public double PromoSalesVolume08
        //{ 
        //    get => _row["август"];
        //    set => _row["август"] = value;
        //}
        //public double PromoSalesVolume09
        //{ 
        //    get => _row["сентябрь"];
        //    set => _row["сентябрь"] = value;
        //}
        //public double PromoSalesVolume10
        //{ 
        //    get => _row["октябрь"];
        //    set => _row["октябрь"] = value;
        //}
        //public double PromoSalesVolume11
        //{ 
        //    get => _row["ноябрь"];
        //    set => _row["ноябрь"] = value;
        //}
        //public double PromoSalesVolume12
        //{ 
        //    get => _row["декабрь"];
        //    set => _row["декабрь"] = value;
        //}
        #endregion

        #region промо GS
        public double PromoGS01
        { 
            get => _row["ПРОМО GS январь, руб."];
            set => _row["ПРОМО GS январь, руб."] = value;
        }
        public double PromoGS02
        { 
            get => _row["ПРОМО GS февраль, руб."];
            set => _row["ПРОМО GS февраль, руб."] = value;
        }
        public double PromoGS03
        { 
            get => _row["ПРОМО GS март, руб."];
            set => _row["ПРОМО GS март, руб."] = value;
        }
        public double PromoGS04
        { 
            get => _row["ПРОМО GS апрель, руб."];
            set => _row["ПРОМО GS апрель, руб."] = value;
        }
        public double PromoGS05
        { 
            get => _row["ПРОМО GS май, руб."];
            set => _row["ПРОМО GS май, руб."] = value;
        }
        public double PromoGS06
        { 
            get => _row["ПРОМО GS июнь, руб."];
            set => _row["ПРОМО GS июнь, руб."] = value;
        }
        public double PromoGS07
        { 
            get => _row["ПРОМО GS июль, руб."];
            set => _row["ПРОМО GS июль, руб."] = value;
        }
        public double PromoGS08
        { 
            get => _row["ПРОМО GS август, руб."];
            set => _row["ПРОМО GS август, руб."] = value;
        }
        public double PromoGS09
        { 
            get => _row["ПРОМО GS сентябрь, руб."];
            set => _row["ПРОМО GS сентябрь, руб."] = value;
        }
        public double PromoGS10
        { 
            get => _row["ПРОМО GS октябрь, руб."];
            set => _row["ПРОМО GS октябрь, руб."] = value;
        }
        public double PromoGS11
        { 
            get => _row["ПРОМО GS ноябрь, руб."];
            set => _row["ПРОМО GS ноябрь, руб."] = value;
        }
        public double PromoGS12
        { 
            get => _row["ПРОМО GS декабрь, руб."];
            set => _row["ПРОМО GS декабрь, руб."] = value;
        }
        #endregion

        #region промо NS
        public double PromoNS01
        { 
            get => _row["ПРОМО NS январь, руб."];
            set => _row["ПРОМО NS январь, руб."] = value;
        }
        public double PromoNS02
        { 
            get => _row["ПРОМО NS февраль, руб."];
            set => _row["ПРОМО NS февраль, руб."] = value;
        }
        public double PromoNS03
        { 
            get => _row["ПРОМО NS март, руб."];
            set => _row["ПРОМО NS март, руб."] = value;
        }
        public double PromoNS04
        { 
            get => _row["ПРОМО NS апрель, руб."];
            set => _row["ПРОМО NS апрель, руб."] = value;
        }
        public double PromoNS05
        { 
            get => _row["ПРОМО NS май, руб."];
            set => _row["ПРОМО NS май, руб."] = value;
        }
        public double PromoNS06
        { 
            get => _row["ПРОМО NS июнь, руб."];
            set => _row["ПРОМО NS июнь, руб."] = value;
        }
        public double PromoNS07
        { 
            get => _row["ПРОМО NS июль, руб."];
            set => _row["ПРОМО NS июль, руб."] = value;
        }
        public double PromoNS08
        { 
            get => _row["ПРОМО NS август, руб."];
            set => _row["ПРОМО NS август, руб."] = value;
        }
        public double PromoNS09
        { 
            get => _row["ПРОМО NS сентябрь, руб."];
            set => _row["ПРОМО NS сентябрь, руб."] = value;
        }
        public double PromoNS10
        { 
            get => _row["ПРОМО NS октябрь, руб."];
            set => _row["ПРОМО NS октябрь, руб."] = value;
        }
        public double PromoNS11
        { 
            get => _row["ПРОМО NS ноябрь, руб."];
            set => _row["ПРОМО NS ноябрь, руб."] = value;
        }
        public double PromoNS12
        { 
            get => _row["ПРОМО NS декабрь, руб."];
            set => _row["ПРОМО NS декабрь, руб."] = value;
        }
        #endregion

        #region GpValue
        public double GPValue01
        { 
            get => _row["ИТОГО GP value январь, руб."];
            set => _row["ИТОГО GP value январь, руб."] = value;
        }
        public double GPValue02
        { 
            get => _row["ИТОГО GP value февраль, руб."];
            set => _row["ИТОГО GP value февраль, руб."] = value;
        }
        public double GPValue03
        { 
            get => _row["ИТОГО GP value март, руб."];
            set => _row["ИТОГО GP value март, руб."] = value;
        }
        public double GPValue04
        { 
            get => _row["ИТОГО GP value апрель, руб."];
            set => _row["ИТОГО GP value апрель, руб."] = value;
        }
        public double GPValue05
        { 
            get => _row["ИТОГО GP value май, руб."];
            set => _row["ИТОГО GP value май, руб."] = value;
        }
        public double GPValue06
        { 
            get => _row["ИТОГО GP value июнь, руб."];
            set => _row["ИТОГО GP value июнь, руб."] = value;
        }
        public double GPValue07
        { 
            get => _row["ИТОГО GP value июль, руб."];
            set => _row["ИТОГО GP value июль, руб."] = value;
        }
        public double GPValue08
        {
            get => _row["ИТОГО GP value август, руб."];
            set => _row["ИТОГО GP value август, руб."] = value;
        }
        public double GPValue09
        { 
            get => _row["ИТОГО GP value сентябрь, руб."];
            set => _row["ИТОГО GP value сентябрь, руб."] = value;
        }
        public double GPValue10
        { 
            get => _row["ИТОГО GP value октябрь, руб."];
            set => _row["ИТОГО GP value октябрь, руб."] = value;
        }
        public double GPValue11
        { 
            get => _row["ИТОГО GP value ноябрь, руб."];
            set => _row["ИТОГО GP value ноябрь, руб."] = value;
        }
        public double GPValue12
        { 
            get => _row["ИТОГО GP value декабрь, руб."];
            set => _row["ИТОГО GP value декабрь, руб."] = value;
        }
        #endregion

        /// <summary>
        /// Установка значений таблицы элементу
        /// </summary>
        /// <param name=""></param>
        public void SetParamsToItem(PlanningNewYearTable planningNewYears)
        {
            //planning._TableWorksheetName = this.SheetName;
            //planning.Year = this.Year;
            CustomerStatus = planningNewYears.CustomerStatus;
            ChannelType = planningNewYears.ChannelType;
            MaximumBonus = planningNewYears.MaximumBonus;
            CurrentDate = planningNewYears.CurrentDate;
            planningDate = planningNewYears.planningDate;
        }

        public void SetProduct(ProductItem product)
        {
            this.Article = product.Article;

            this.SupercategoryEng = product.SupercategoryEng;
            this.Supercategory = product.SuperCategory;
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
        }

        public void SetRRC(RRCItem rrcPlan, RRCItem rrcCurrent)
        {
            if (rrcPlan != null)
            {
                this.RRCCUrent = rrcCurrent.RRCNDS;
                this.DIYCurrent = rrcCurrent.DIY;
                //this.IRPCurrent = rrcCurrent.IRP;
                //this.IRPIndexCurrent = rrcCurrent.IRPIndex;
                //this.RRPCurrent = rrcCurrent.RRP;
            }

            if (rrcCurrent != null)
            {
                this.RRCPlan = rrcPlan.RRCNDS;
                this.DIYPlan = rrcPlan.DIY;
                //this.IRPPlan = rrcPlan.IRP;
                //this.IRPIndexPlan = rrcPlan.IRPIndex;
                //this.RRPPlan = rrcPlan.RRP;
            }
        }

        public void SetSTK(STKItem stkPlan, STKItem stkCurrent)
        {
            if (stkPlan != null)
            {
                this.STKRubPlan = stkPlan.STKRub;
            }

            if (stkCurrent != null)
            {
                this.STKRubCurrent = stkCurrent.STKRub;
            }
        }

        public void SetValuesPrognosis(List<ArticleQuantity> deicionQuantities, List<ArticleQuantity> bugetQuantities)
        {
            //Извлечение из списков Descision и Buget элементы с соответствующим артикулом и НЕ Promo
            List<ArticleQuantity> articleDescisionQuantities = deicionQuantities.FindAll(x => x.Article == Article && !IsPromo(x)).ToList();
            List<ArticleQuantity> articleBugetQuantities = bugetQuantities.FindAll(x => x.Article == Article && !IsPromo(x)).ToList();

            ArticleQuantity[] articles = GetArticlesQuantities(articleDescisionQuantities, articleBugetQuantities);
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
            //
            #endregion
        }

        #region проверка promo/prognosis
        /// <summary>
        /// проверка promo/prognosis
        /// </summary>
        /// <param name="articleQuantity"></param>
        /// <returns></returns>
        private bool IsPromo(ArticleQuantity articleQuantity)
        {
            return articleQuantity.Campaign != "0" && articleQuantity.Campaign != null ? true : false;
        }

        #endregion

        private ArticleQuantity[] GetArticlesQuantities(List<ArticleQuantity> articleDescisionQuantities, List<ArticleQuantity> articleBugetQuantities)
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

                List<ArticleQuantity> articleQuantities = month < CurrentDate.Month ?
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
                    double bonus = month < CurrentDate.Month ? articleQuantity.Bonus : articleQuantity.PriceList * MaximumBonus;
                    newArticleQuantity.Bonus += bonus;
                }
                //

                newArticleQuantity.Article = monthQuantities[0].Article;
                newArticleQuantity.Campaign = monthQuantities[0].Campaign;

                return newArticleQuantity;
            }
            //
        }
    }
}
