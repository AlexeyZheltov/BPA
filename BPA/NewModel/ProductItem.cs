using BPA.Modules;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;

namespace BPA.NewModel
{
    class ProductItem
    {
        TableRow _row;
        public ProductItem(TableRow row) => _row = row;

        #region Основные свойства
        public int Id
        {
            get => _row["№"];
            set => _row["№"] = value;
        }
        public string Category
        {
            get => _row["Категория для прайс-листа дилеров"];
            set => _row["Категория для прайс-листа дилеров"] = value;
        }
        public string SuperCategory
        {
            get => _row["Суперкатегория"];
            set => _row["Суперкатегория"] = value;
        }
        public string SupercategoryEng
        {
            get => _row["Суперкатегория (ENG)"];
            set => _row["Суперкатегория (ENG)"] = value;
        }
        public string SupercategoryRu
        {
            get => _row["Суперкатегория (RUS)"];
            set => _row["Суперкатегория (RUS)"] = value;
        }
        public string ProductGroup
        {
            get => _row["Продукт группа"];
            set => _row["Продукт группа"] = value;
        }
        public string ProductGroupEng
        {
            get => _row["Название продукт группы (ENG)"];
            set => _row["Название продукт группы (ENG)"] = value;
        }
        public string ProductGroupRu
        {
            get => _row["Название продукт группы (RUS)"];
            set => _row["Название продукт группы (RUS)"] = value;
        }
        public string SubGroup
        {
            get => _row["SubGroup"];
            set => _row["SubGroup"] = value;
        }
        public string GenericName
        {
            get => _row["Generic Name (long)"];
            set => _row["Generic Name (long)"] = value;
        }
        public string Model
        {
            get => _row["Model"];
            set => _row["Model"] = value;
        }
        public string PNS
        {
            get => _row["PNS"];
            set => _row["PNS"] = value;
        }
        public string Article
        {
            get => _row["Артикул"];
            set => _row["Артикул"] = value;
        }
        public string ArticleOld
        {
            get => _row["Артикул предшественника (если есть)"];
            set => _row["Артикул предшественника (если есть)"] = value;
        }
        public string ArticleEng
        {
            get => _row["Название артикула (ENG)"];
            set => _row["Название артикула (ENG)"] = value;
        }
        public string ArticleRu
        {
            get => _row["Название артикула (RUS)"];
            set => _row["Название артикула (RUS)"] = value;
        }
        public string Calendar
        {
            get => _row["Используемый календарь"];
            set => _row["Используемый календарь"] = value;
        }
        public string CalendarToBeSoldIn
        {
            get => _row["to be sold in"];
            set => _row["to be sold in"] = value;
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
        public string CalendarGTIN
        {
            get => _row["GTIN-13/EAN"];
            set => _row["GTIN-13/EAN"] = value;
        }
        public string CalendarCurrentProducingFactoryEntityReference
        {
            get => _row["Current Producing Factory Entity Reference"];
            set => _row["Current Producing Factory Entity Reference"] = value;
        }
        public string CalendarCountryOfOrigin
        {
            get => _row["Country of Origin"];
            set => _row["Country of Origin"] = value;
        }
        public string CalendarUnitOfMeasure
        {
            get => _row["Unit of measure"];
            set => _row["Unit of measure"] = value;
        }
        public string CalendarQuantityInMasterPack
        {
            get => _row["Quantity in Master pack"];
            set => _row["Quantity in Master pack"] = value;
        }
        public double CalendarArticleGrossWeightPreliminary
        {
            get => _row["Article gross weight, preliminary"];
            set => _row["Article gross weight, preliminary"] = value;
        }
        public double CalendarArticleGrossWeight
        {
            get => _row["Article gross weight"];
            set => _row["Article gross weight"] = value;
        }
        public double CalendarArticleNetWeightPreliminary
        {
            get => _row["Article net weight, preliminary"];
            set => _row["Article net weight, preliminary"] = value;
        }
        public double CalendarArticleNetWeight
        {
            get => _row["Article net weight"];
            set => _row["Article net weight"] = value;
        }
        public double CalendarPackagingLength
        {
            get => _row["Packaging length"];
            set => _row["Packaging length"] = value;
        }
        public double CalendarPackagingWidth
        {
            get => _row["Packaging width"];
            set => _row["Packaging width"] = value;
        }
        public double CalendarPackagingHeight
        {
            get => _row["Packaging height"];
            set => _row["Packaging height"] = value;
        }
        public double CalendarPackagingVolume
        {
            get => _row["Packaging volume"];
            set => _row["Packaging volume"] = value;
        }
        public double CalendarProductSizeLength
        {
            get => _row["Product size height"];
            set => _row["Product size height"] = value;
        }
        public double CalendarProductSizeHeight
        {
            get => _row["Product size length"];
            set => _row["Product size length"] = value;
        }
        public double CalendarProductSizeWidth
        {
            get => _row["Product size width"];
            set => _row["Product size width"] = value;
        }
        public double CalendarUnitsPerPallet
        {
            get => _row["Units Per Pallet"];
            set => _row["Units Per Pallet"] = value;
        }
        public string Status
        {
            get => _row["Актуальный статус"];
            set => _row["Актуальный статус"] = value;
        }
        public string Exclusive
        {
            get => _row["Эксклюзив клиента или канала продажи"];
            set => _row["Эксклюзив клиента или канала продажи"] = value;
        }
        public string LocalCertificate
        {
            get => _row["Локальный сертификат"];
            set => _row["Локальный сертификат"] = value;
        }
        public double IRP
        {
            get => _row["IRP, Eur"];
            set => _row["IRP, Eur"] = value;
        }
        public double RRCCurrent
        {
            get => _row["РРЦ текущий"];
            set => _row["РРЦ текущий"] = value;
        }
        public double DIYCurrent
        {
            get => _row["DIY текущий"];
            set => _row["DIY текущий"] = value;
        }
        public double RRCPercent
        {
            get => _row["Процент повышения РРЦ"];
            set => _row["Процент повышения РРЦ"] = value;
        }
        public double RRCCalculated
        {
            get => _row["РРЦ расчетная, руб."];
            set => _row["РРЦ расчетная, руб."] = value;
        }
        public double RRCFinal
        {
            get => _row["РРЦ финальная, руб."];
            set => _row["РРЦ финальная, руб."] = value;
        }
        public double RRCEuro
        {
            get => _row["РРЦ, евро"];
            set => _row["РРЦ, евро"] = value;
        }
        public double IRPIndex
        {
            get => _row["Индекс IRP"];
            set => _row["Индекс IRP"] = value;
        }
        public double DIYDiscount
        {
            get => _row["Скидка DIY"];
            set => _row["Скидка DIY"] = value;
        }
        public double DIY
        {
            get => _row["DIY price list, руб. без НДС"];
            set => _row["DIY price list, руб. без НДС"] = value;
        }

        #endregion

        public void UpdatePriceFromRRC(RRCItem item)
        {
            if(item != null)
            {
                RRCCurrent = item.RRCNDS;
                DIYCurrent = item.DIY;
                IRP = item.IRP;
                IRPIndex = item.IRPIndex;
            }
        }


        public void MarkColumn(String col_name)
        {
            //Dictionary<string, int> ColDict = (Dictionary<string, int>)GetType().GetProperty("ColDict").GetValue(this);
            //ListRow row = GetRow(GetParametrValueId());
            //row.Range[1, ColDict[fildNameToMark]].Interior.Color = 65535;
        }
    }
}
