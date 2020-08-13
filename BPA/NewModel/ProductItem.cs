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

        public string Article
        {
            get => _row["Артикул"];
            set => _row["Артикул"] = value;
        }

        public double IRP
        {
            get => _row["IRP, Eur"];
            set => _row["IRP, Eur"] = value;
        }

        public double IRPIndex
        {
            get => _row["Индекс IRP"];
            set => _row["Индекс IRP"] = value;
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

        public double DIY
        {
            get => _row["DIY price list, руб. без НДС"];
            set => _row["DIY price list, руб. без НДС"] = value;
        }

        #endregion

        #region Свойства из календаря
        public string CalendarName
        {
            get => _row["Используемый календарь"];
            set => _row["Используемый календарь"] = value;
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

        public string CalendarToBeSoldIn
        {
            get => _row["to be sold in"];
            set => _row["to be sold in"] = value;
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
        public string CalendarArticleGrossWeightPreliminary
        {
            get => _row["Article gross weight, preliminary"];
            set => _row["Article gross weight, preliminary"] = value;
        }
        public string CalendarArticleGrossWeight
        {
            get => _row["Article gross weight"];
            set => _row["Article gross weight"] = value;
        }
        public string CalendarArticleNetWeightPreliminary
        {
            get => _row["Article net weight, preliminary"];
            set => _row["Article net weight, preliminary"] = value;
        }
        public string CalendarArticleNetWeight
        {
            get => _row["Article net weight"];
            set => _row["Article net weight"] = value;
        }
        public string CalendarPackagingLength
        {
            get => _row["Packaging length"];
            set => _row["Packaging length"] = value;
        }
        public string CalendarPackagingHeight
        {
            get => _row["Packaging height"];
            set => _row["Packaging height"] = value;
        }
        public string CalendarPackagingWidth
        {
            get => _row["Packaging width"];
            set => _row["Packaging width"] = value;
        }
        public string CalendarPackagingVolume
        {
            get => _row["Packaging volume"];
            set => _row["Packaging volume"] = value;
        }
        public string CalendarProductSizeHeight
        {
            get => _row["Product size length"];
            set => _row["Product size length"] = value;
        }
        public string CalendarProductSizeWidth
        {
            get => _row["Product size width"];
            set => _row["Product size width"] = value;
        }
        public string CalendarProductSizeLength
        {
            get => _row["Product size height"];
            set => _row["Product size height"] = value;
        }
        public string CalendarUnitsPerPallet
        {
            get => _row["Units Per Pallet"];
            set => _row["Units Per Pallet"] = value;
        }

        public string Model
        {
            get => _row["Model"];
            set => _row["Model"] = value;
        }
        public string SubGroup
        {
            get => _row["SubGroup"];
            set => _row["SubGroup"] = value;
        }
            
        public string PNS
        {
            get => _row["PNS"];
            set => _row["PNS"] = value;
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
