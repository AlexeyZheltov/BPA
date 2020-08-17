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

        public int Id
        {
            get => _row["№"];
            set => _row["№"] = value;
        }
        public string RRCNDS
        {
            get => _row["РРЦ 2021, руб. с НДС"];
            set => _row["РРЦ 2021, руб. с НДС"] = value;
        }
        public string DIYPriceList
        {
            get => _row["DIY цена 2021, руб. с НДС"];
            set => _row["DIY цена 2021, руб. с НДС"] = value;
        }
        public string STKRub
        {
            get => _row["STK 2.5 2021, RUB"];
            set => _row["STK 2.5 2021, RUB"] = value;
        }
        public string RRCCUrent
        {
            get => _row["РРЦ 2020, руб. с НДС" ];
            set => _row["РРЦ 2020, руб. с НДС"] = value;
        }
        public string RRCPlan
        {
            get => _row["РРЦ 2021, руб. с НДС"];
            set => _row["РРЦ 2021, руб. с НДС"] = value;
        }
        public string DIYCurrent
        {
            get => _row["DIY цена 2020, руб. с НДС"];
            set => _row["DIY цена 2020, руб. с НДС"] = value;
        }
        public string DIYPlan
        {
            get => _row["DIY цена 2021, руб. с НДС"];
            set => _row["DIY цена 2021, руб. с НДС"] = value;
        }
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
        public string CalendarSalesStartDate
        {
            get => _row["Sales Start Date"];
            set => _row["Sales Start Date"] = value;
        }
        public string CalendarPreliminaryEliminationDate
        {
            get => _row["Preliminary Elimination Date"];
            set => _row["Preliminary Elimination Date"] = value;
        }
        public string CalendarEliminationDate
        {
            get => _row["Elimination Date"];
            set => _row["Elimination Date"] = value;
        }
        public string Status
        {
            get => _row["status"];
            set => _row["status"] = value;
        }

        #region Prognosis
        public string QuantityPrognosisYear {
            get => _row["ИТОГО Прогноз за год, шт."];
            set => _row["ИТОГО Прогноз за год, шт."] = value;
        }
        public string QuantityPrognosis01 {
            get => _row["ИТОГО Прогноз январь, шт."];
            set => _row ["ИТОГО Прогноз январь, шт."] = value;
        }
        public string QuantityPrognosis02 {
            get => _row["ИТОГО Прогноз февраль, шт."];
            set => _row ["ИТОГО Прогноз февраль, шт."] = value;
        }
        public string QuantityPrognosis03 {
            get => _row["ИТОГО Прогноз март, шт."];
            set => _row ["ИТОГО Прогноз март, шт."] = value;
        }
        public string QuantityPrognosis04 {
            get => _row["ИТОГО Прогноз апрель, шт."];
            set => _row ["ИТОГО Прогноз апрель, шт."] = value;
        }
        public string QuantityPrognosis05 {
            get => _row["ИТОГО Прогноз май, шт."];
            set => _row ["ИТОГО Прогноз май, шт."] = value;
        }
        public string QuantityPrognosis06 {
            get => _row["ИТОГО Прогноз июнь, шт."];
            set => _row ["ИТОГО Прогноз июнь, шт."] = value;
        }
        public string QuantityPrognosis07 {
            get => _row["ИТОГО Прогноз июль, шт."];
            set => _row ["ИТОГО Прогноз июль, шт."] = value;
        }
        public string QuantityPrognosis08 {
            get => _row["ИТОГО Прогноз август, шт."];
            set => _row ["ИТОГО Прогноз август, шт."] = value;
        }
        public string QuantityPrognosis09 {
            get => _row["ИТОГО Прогноз сентябрь, шт."];
            set => _row ["ИТОГО Прогноз сентябрь, шт."] = value;
        }
        public string QuantityPrognosis10 {
            get => _row["ИТОГО Прогноз октябрь, шт."];
            set => _row ["ИТОГО Прогноз октябрь, шт."] = value;
        }
        public string QuantityPrognosis11 {
            get => _row["ИТОГО Прогноз ноябрь, шт."];
            set => _row ["ИТОГО Прогноз ноябрь, шт."] = value;
        }
        public string QuantityPrognosis12 {
            get => _row["ИТОГО Прогноз декабрь, шт."];
            set => _row ["ИТОГО Прогноз декабрь, шт."] = value;
        }

        public string GSPrognosisYear {
            get => _row["ИТОГО GS за год, шт."];
            set => _row ["ИТОГО GS за год, шт."] = value;
        }
        public string GSPrognosis01 {
            get => _row["ИТОГО GS январь, шт."];
            set => _row ["ИТОГО GS январь, шт."] = value;
        }
        public string GSPrognosis02 {
            get => _row["ИТОГО GS февраль, шт."];
            set => _row ["ИТОГО GS февраль, шт."] = value;
        }
        public string GSPrognosis03 {
            get => _row["ИТОГО GS март, шт."];
            set => _row ["ИТОГО GS март, шт."] = value;
        }
        public string GSPrognosis04 {
            get => _row["ИТОГО GS апрель, шт."];
            set => _row ["ИТОГО GS апрель, шт."] = value;
        }
        public string GSPrognosis05 {
            get => _row["ИТОГО GS май, шт."];
            set => _row ["ИТОГО GS май, шт."] = value;
        }
        public string GSPrognosis06 {
            get => _row["ИТОГО GS июнь, шт."];
            set => _row ["ИТОГО GS июнь, шт."] = value;
        }
        public string GSPrognosis07 {
            get => _row["ИТОГО GS июль, шт."];
            set => _row ["ИТОГО GS июль, шт."] = value;
        }
        public string GSPrognosis08 {
            get => _row["ИТОГО GS август, шт."];
            set => _row ["ИТОГО GS август, шт."] = value;
        }
        public string GSPrognosis09 {
            get => _row["ИТОГО GS сентябрь, шт."];
            set => _row ["ИТОГО GS сентябрь, шт."] = value;
        }
        public string GSPrognosis10 {
            get => _row["ИТОГО GS октябрь, шт."];
            set => _row ["ИТОГО GS октябрь, шт."] = value;
        }
        public string GSPrognosis11 {
            get => _row["ИТОГО GS ноябрь, шт."];
            set => _row ["ИТОГО GS ноябрь, шт."] = value;
        }
        public string GSPrognosis12 {
            get => _row["ИТОГО GS декабрь, шт."];
            set => _row ["ИТОГО GS декабрь, шт."] = value;
        }

        public string NSPrognosisYear {
            get => _row["ИТОГО NS за год, шт."];
            set => _row ["ИТОГО NS за год, шт."] = value;
        }
        public string NSPrognosis01 {
            get => _row["ИТОГО NS январь, шт."];
            set => _row ["ИТОГО NS январь, шт."] = value;
        }
        public string NSPrognosis02 {
            get => _row["ИТОГО NS февраль, шт."];
            set => _row ["ИТОГО NS февраль, шт."] = value;
        }
        public string NSPrognosis03 {
            get => _row["ИТОГО NS март, шт."];
            set => _row ["ИТОГО NS март, шт."] = value;
        }
        public string NSPrognosis04 {
            get => _row["ИТОГО NS апрель, шт."];
            set => _row ["ИТОГО NS апрель, шт."] = value;
        }
        public string NSPrognosis05 {
            get => _row["ИТОГО NS май, шт."];
            set => _row ["ИТОГО NS май, шт."] = value;
        }
        public string NSPrognosis06 {
            get => _row["ИТОГО NS июнь, шт."];
            set => _row ["ИТОГО NS июнь, шт."] = value;
        }
        public string NSPrognosis07 {
            get => _row["ИТОГО NS июль, шт."];
            set => _row ["ИТОГО NS июль, шт."] = value;
        }
        public string NSPrognosis08 {
            get => _row["ИТОГО NS август, шт."];
            set => _row ["ИТОГО NS август, шт."] = value;
        }
        public string NSPrognosis09 {
            get => _row["ИТОГО NS сентябрь, шт."];
            set => _row ["ИТОГО NS сентябрь, шт."] = value;
        }
        public string NSPrognosis10 {
            get => _row["ИТОГО NS октябрь, шт."];
            set => _row ["ИТОГО NS октябрь, шт."] = value;
        }
        public string NSPrognosis11 {
            get => _row["ИТОГО NS ноябрь, шт."];
            set => _row ["ИТОГО NS ноябрь, шт."] = value;
        }
        public string NSPrognosis12
        {
            get => _row["ИТОГО NS декабрь, шт."];
            set => _row["ИТОГО NS декабрь, шт."] = value;
        }
        #endregion

    }
}
