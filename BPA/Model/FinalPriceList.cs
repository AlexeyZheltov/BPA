using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace BPA.Model
{

    /// <summary>
    /// Итоговый прайс-лист
    /// </summary>
    internal class FinalPriceList : TableBase
    {
        private readonly Microsoft.Office.Interop.Excel.Application Application = Globals.ThisWorkbook.Application;

        public override string TableName => "Прайс_лист";
        public override string SheetName => "Прайс лист";

        #region --- Словарь ---

        public override IDictionary<string, string> Filds => _filds;
        private readonly Dictionary<string, string> _filds = new Dictionary<string, string>
        {
            {  "Id", "Id" },
            {  "Category","КАТЕГОРИЯ"  },
            {  "Photo","Фото продукта"  },
            {  "ProductGroup","Продукт группа - название"  },
            {  "ArticleGardena","Артикул GARDENA"  },
            {  "ArticleOld","Артикул предшествующего продукта (если есть)"  },
            {  "Name","Название"  },
            {  "Description","Описание"  },
            {  "RRC","РРЦ 2020, руб. с НДС"  },
            {  "EAN","EAN штуки"  },
            {  "CountryOfOrigin","страна производства"  },
            {  "UnitOfMeasure","единица измерения"  },
            {  "QuantityInMasterPack","кол-во в мастер-паке"  },
            {  "ArticleGrossWeight","Вес гросс штуки, финальный"  },
            {  "ArticleNetWeight","Вес нетто штуки, финальный"  },
            {  "PackagingLength","Длина упаковки"  },
            {  "PackagingWidth","Ширина упаковки"  },
            {  "PackagingHeight","Высота упаковки"  },
            {  "PackagingVolume","Объем упаковки"  },
            {  "UnitsPerPallet","Кол-во штук на паллете"  },
            {  "Certificate","Сертификация"  },
            {  "Warranty","Гарантия"  }
        };
        #endregion

        #region -- Основные свойства столбцов ---

        /// <summary>
        /// Id
        /// </summary>
        public int Id
        {
            get; set;
        }

        /// <summary>
        /// КАТЕГОРИЯ
        /// </summary>
        public string Category
        {
            get; set;
        }

        /// <summary>
        /// Фото продукта
        /// </summary>
        public string Photo
        {
            get; set;
        }

        /// <summary>
        /// Продукт группа - название
        /// </summary>
        public string ProductGroup
        {
            get; set;
        }

        /// <summary>
        /// Артикул GARDENA
        /// </summary>
        public string ArticleGardena
        {
            get; set;
        }

        /// <summary>
        /// Артикул предшествующего продукта (если есть)
        /// </summary>
        public string ArticleOld
        {
            get; set;
        }

        /// <summary>
        /// Название
        /// </summary>
        public string Name
        {
            get; set;
        }

        /// <summary>
        /// Описание
        /// </summary>
        public string Description
        {
            get; set;
        }

        /// <summary>
        /// РРЦ 2020, руб. с НДС
        /// </summary>
        public double RRC
        {
            get; set;
        }

        /// <summary>
        /// EAN штуки
        /// </summary>
        public string EAN
        {
            get; set;
        }

        /// <summary>
        /// страна производства
        /// </summary>
        public string CountryOfOrigin
        {
            get; set;
        }

        /// <summary>
        /// единица измерения
        /// </summary>
        public string UnitOfMeasure
        {
            get; set;
        }

        /// <summary>
        /// кол-во в мастер-паке
        /// </summary>
        public string QuantityInMasterPack
        {
            get; set;
        }

        /// <summary>
        /// Вес гросс штуки, финальный
        /// </summary>
        public string ArticleGrossWeight
        {
            get; set;
        }

        /// <summary>
        /// Вес нетто штуки, финальный
        /// </summary>
        public string ArticleNetWeight
        {
            get; set;
        }

        /// <summary>
        /// Длина упаковки
        /// </summary>
        public string PackagingLength
        {
            get; set;
        }

        /// <summary>
        /// Ширина упаковки
        /// </summary>
        public string PackagingWidth
        {
            get; set;
        }

        /// <summary>
        /// Высота упаковки
        /// </summary>
        public string PackagingHeight
        {
            get; set;
        }

        /// <summary>
        /// Объем упаковки
        /// </summary>
        public string PackagingVolume
        {
            get; set;
        }

        /// <summary>
        /// Кол-во штук на паллете
        /// </summary>
        public string UnitsPerPallet
        {
            get; set;
        }

        /// <summary>
        /// Сертификация
        /// </summary>
        public string Certificate
        {
            get; set;
        }

        /// <summary>
        /// Гарантия
        /// </summary>
        public string Warranty
        {
            get; set;
        }

        #endregion

        public FinalPriceList() { }

        public FinalPriceList(Product product)
        {
            this.Category = product.Category;
            this.ProductGroup = product.ProductGroupRu;
            this.ArticleGardena = product.Article;
            this.ArticleOld = product.ArticleOld;
            this.Name = product.ArticleRu;
            //this.Description= product. //откуда описание
            //this.RRC = product.RRCFinal; //цена из
            this.EAN = product.CalendarGTIN;
            this.CountryOfOrigin = product.CalendarCountryOfOrigin;
            this.UnitOfMeasure = product.CalendarUnitOfMeasure;
            this.QuantityInMasterPack = product.CalendarQuantityInMasterPack;
            this.ArticleGrossWeight = product.CalendarArticleGrossWeight;
            this.ArticleNetWeight = product.CalendarArticleNetWeight;
            this.PackagingLength = product.CalendarPackagingLength;
            this.PackagingWidth = product.CalendarPackagingWidth;
            this.PackagingHeight = product.CalendarPackagingHeight;
            this.PackagingVolume = product.CalendarPackagingVolume;
            this.UnitsPerPallet = product.CalendarUnitsPerPallet;
            this.Certificate = product.LocalCertificate;

            //this.Warranty= product.CalendarPackagingWidth //?
        }

        public FinalPriceList(Excel.ListRow row) => SetProperty(row);

        public static List<FinalPriceList> GetAllFinalPriceList()
        {
            List<FinalPriceList> finalPriceLists = new List<FinalPriceList>();
            new FinalPriceList().ReadColNumbers();
            foreach (Excel.ListRow row in new FinalPriceList().Table.ListRows)
            {
                finalPriceLists.Add(new FinalPriceList(row));
            }
            return finalPriceLists;
        }
    }
}
