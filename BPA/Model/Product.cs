using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BPA.Model {
    /// <summary>
    /// Справочник Товаров
    /// </summary>
    class Product : TableBase {
        public override string TableName => "Товары";
        public override string SheetName => "Товары";

        public override IDictionary<string, string> Filds => _filds;
        private readonly Dictionary<string, string> _filds = new Dictionary<string, string>
        {
            { "Id","№" },
            { "Category","Категория для прайс-листа диллеров" },
            { "SupercategoryEng","Суперкатегория(ENG)"  },
            { "SupercategoryRu","Суперкатегория(RUS)" },
            { "ProductGroup","Продукт группа" },
            { "ProductGroupEng","Название продукт группы(ENG)" },
            { "ProductGroupRu","Название продукт группы(RUS)" },
            { "SubGroup", "SubGroup" },
            { "GenericName", "Generic Name(long)" },
            { "Model", "Model" },
            { "PNS", "PNS" },
            { "Article","Артикул" },
            { "ArticleOld","Артикул предшественника(если есть)" },
            { "ArticleEng","Название артикула(ENG)" },
            { "ArticleRu","Название артикула(RUS)" },
            { "Calendar", "Используемый календарь" },

            { "CalendarToBeSoldIn","to be sold in" },
            { "CalendarSalesStartDate","Sales Start Date" },
            { "CalendarPreliminaryEliminationDate","Preliminary Elimination Date" },
            { "CalendarEliminationDate","Elimination Date" },
            { "CalendarGTIN","GTIN-13/EAN" },
            { "CalendarCurrentProducingFactoryEntityReference","Current Producing Factory Entity Reference" },
            { "CalendarCountryOfOrigin","Country of Origin" },
            { "CalendarUnitOfMeasure","Unit of measure" },
            { "CalendarQuantityInMasterPack","Quantity in Master pack" },
            { "CalendarArticleGrossWeightPreliminary","Article gross weight, preliminary" },
            { "CalendarArticleGrossWeight","Article gross weight" },
            { "CalendarArticleNetWeightPreliminary","Article net weight, preliminary" },
            { "CalendarArticleNetWeight","Article net weight" },
            { "CalendarPackagingLength","Packaging length" },
            { "CalendarPackagingWidth","Packaging width" },
            { "CalendarPackagingHeight","Packaging height" },
            { "CalendarPackagingVolume","Packaging volume" },
            { "CalendarProductSizeLength","Product size length" },
            { "CalendarProductSizeHeight","Product size height" },
            { "CalendarProductSizeWidth","Product size width" },
            { "CalendarUnitsPerPallet","Units Per Pallet" },

            { "Status","Статус" },
            { "Exclusive","Эксклюзив клиента или канала продажи" },
            { "LocalCertificate","Локальный сертификат" }
        };


        #region main
        /// <summary>
        /// №
        /// </summary>
        public string Id { get; set; }
        /// <summary>
        /// Категория для прайс-листа диллеров
        /// </summary>
        public string Category { get; set; }
        /// <summary>
        /// Суперкатегория(ENG)
        /// </summary>
        public string SupercategoryEng { get; set; }
        /// <summary>
        /// Суперкатегория(RUS)
        /// </summary>
        public string SupercategoryRu { get; set; }
        /// <summary>
        /// Продукт группа
        /// </summary>
        public string ProductGroup { get; set; }
        /// <summary>
        /// Название продукт группы(ENG)
        /// </summary>
        public string ProductGroupEng { get; set; }
        /// <summary>
        /// Название продукт группы(RUS)
        /// </summary>
        public string ProductGroupRu { get; set; }
        /// <summary>
        /// SubGroup
        /// </summary>
        public string SubGroup { get; set; }
        /// <summary>
        /// Generic Name(long)
        /// </summary>
        public string GenericName { get; set; }
        /// <summary>
        /// Model
        /// </summary>
        public string Model { get; set; }
        /// <summary>
        /// PNS
        /// </summary>
        public string PNS { get; set; }
        /// <summary>
        /// Артикул
        /// </summary>
        public string Article { get; set; }
        /// <summary>
        /// Артикул предшественника(если есть)
        /// </summary>
        public string ArticleOld { get; set; }
        /// <summary>
        /// Название артикула(ENG)
        /// </summary>
        public string ArticleEng { get; set; }
        /// <summary>
        /// Название артикула(RUS)
        /// </summary>
        public string ArticleRu { get; set; }
        /// <summary>
        /// Используемый календарь
        /// </summary>
        public string Calendar { get; set; }
        /// <summary>
        /// Статус
        /// </summary>
        public string Status { get; set; }
        /// <summary>
        /// Эксклюзив клиента или канала продажи
        /// </summary>
        public string Exclusive { get; set; }
        /// <summary>
        /// Локальный сертификат
        /// </summary>
        public string LocalCertificate { get; set;
        }
        #endregion

        #region From Calendar

        /// <summary>
        /// to be sold in
        /// </summary>
        public string CalendarToBeSoldIn { get; set; }

        /// <summary>
        /// Sales Start Date
        /// </summary>
        public string CalendarSalesStartDate
        {
            get; set;
        }

        /// <summary>
        /// Preliminary Elimination Date
        /// </summary>
        public string CalendarPreliminaryEliminationDate
        {
            get; set;
        }

        /// <summary>
        /// CalendarEliminationDate
        /// </summary>
        public string CalendarEliminationDate
        {
            get; set;
        }
        /// <summary>
         /// CalendarGTIN
         /// </summary>
        public string CalendarGTIN
        {
            get; set;
        }

        /// <summary>
        /// CalendarCurrentProducingFactoryEntityReference
        /// </summary>
        public string CalendarCurrentProducingFactoryEntityReference
        {
            get; set;
        }
        /// <summary>
         /// CalendarCountryOfOrigin
         /// </summary>
        public string CalendarCountryOfOrigin
        {
            get; set;
        }

        /// <summary>
        /// CalendarUnitOfMeasure
        /// </summary>
        public string CalendarUnitOfMeasure
        {
            get; set;
        }
        /// <summary>
        /// CalendarQuantityInMasterPack
        /// </summary>
        public string CalendarQuantityInMasterPack
        {
            get; set;
        }

        /// <summary>
        /// CalendarArticleGrossWeightPreliminary
        /// </summary>
        public string CalendarArticleGrossWeightPreliminary
        {
            get; set;
        }

        /// <summary>
        /// CalendarArticleGrossWeight
        /// </summary>
        public string CalendarArticleGrossWeight
        {
            get; set;
        }

        /// <summary>
        /// CalendarArticleNetWeightPreliminary
        /// </summary>
        public string CalendarArticleNetWeightPreliminary
        {
            get; set;
        }

        /// <summary>
        ///CalendarArticleNetWeight
        /// </summary>
        public string CalendarArticleNetWeight
        {
            get; set;
        }

        /// <summary>
        /// CalendarPackagingLength
        /// </summary>
        public string CalendarPackagingLength
        {
            get; set;
        }

        /// <summary>
        /// CalendarPackagingWidth
        /// </summary>
        public string CalendarPackagingWidth
        {
            get; set;
        }

        /// <summary>
        /// CalendarPackagingHeight
        /// </summary>
        public string CalendarPackagingHeight
        {
            get; set;
        }

        /// <summary>
        /// CalendarPackagingVolume
        /// </summary>
        public string CalendarPackagingVolume
        {
            get; set;
        }

        /// <summary>
        /// CalendarProductSizeLength
        /// </summary>
        public string CalendarProductSizeLength
        {
            get; set;
        }

        /// <summary>
        /// CalendarProductSizeHeight
        /// </summary>
        public string CalendarProductSizeHeight
        {
            get; set;
        }

        /// <summary>
        /// CalendarProductSizeWidth
        /// </summary>
        public string CalendarProductSizeWidth
        {
            get; set;
        }

        /// <summary>
        /// CalendarUnitsPerPallet
        /// </summary>
        public string CalendarUnitsPerPallet
        {
            get; set;
        }

        #endregion
        
        public Product GetProduct(string articul)
        {

            ListRow listRow = GetRow("Article", articul);
            if (listRow != null)
            {
                Product product = new Product();
                product.SetProperty(listRow);
                return product;
            }

            return null;
        }

    }
}



