using BPA.Model;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BPA.NewModel
{
    class FinalPriceItem
    {
        TableRow _row;
        public FinalPriceItem(TableRow row) => _row = row;

        #region Свойства таблицы
        public int Id
        {
            get => _row["Id"];
            set => _row["Id"] = value;
        }
        public string Category
        {
            get => _row["КАТЕГОРИЯ"];
            set => _row["КАТЕГОРИЯ"] = value;
        }
        public string Photo
        {
            get => _row["Фото продукта"];
            set => _row["Фото продукта"] = value;
        }
        public string ProductGroup
        {
            get => _row["Продукт группа - название"];
            set => _row["Продукт группа - название"] = value;
        }
        public string ArticleGardena
        {
            get => _row["Артикул GARDENA"];
            set => _row["Артикул GARDENA"] = value;
        }
        public string ArticleOld
        {
            get => _row["Артикул предшествующего продукта (если есть)"];
            set => _row["Артикул предшествующего продукта (если есть)"] = value;
        }
        public string Name
        {
            get => _row["Название"];
            set => _row["Название"] = value;
        }
        public string Description
        {
            get => _row["Описание"];
            set => _row["Описание"] = value;
        }
        public double RRC
        {
            get => _row["РРЦ 2020, руб. с НДС"];
            set => _row["РРЦ 2020, руб. с НДС"] = value;
        }
        public string EAN
        {
            get => _row["EAN штуки"];
            set => _row["EAN штуки"] = value;
        }
        public string CountryOfOrigin
        {
            get => _row["страна производства"];
            set => _row["страна производства"] = value;
        }
        public string UnitOfMeasure
        {
            get => _row["единица измерения"];
            set => _row["единица измерения"] = value;
        }
        public string QuantityInMasterPack
        {
            get => _row["кол-во в мастер-паке"];
            set => _row["кол-во в мастер-паке"] = value;
        }
        public string ArticleGrossWeight
        {
            get => _row["Вес гросс штуки, финальный"];
            set => _row["Вес гросс штуки, финальный"] = value;
        }
        public string ArticleNetWeight
        {
            get => _row["Вес нетто штуки, финальный"];
            set => _row["Вес нетто штуки, финальный"] = value;
        }
        public string PackagingLength
        {
            get => _row["Длина упаковки"];
            set => _row["Длина упаковки"] = value;
        }
        public string PackagingWidth
        {
            get => _row["Ширина упаковки"];
            set => _row["Ширина упаковки"] = value;
        }
        public string PackagingHeight
        {
            get => _row["Высота упаковки"];
            set => _row["Высота упаковки"] = value;
        }
        public string PackagingVolume
        {
            get => _row["Объем упаковки"];
            set => _row["Объем упаковки"] = value;
        }
        public string UnitsPerPallet
        {
            get => _row["Кол-во штук на паллете"];
            set => _row["Кол-во штук на паллете"] = value;
        }
        public string Certificate
        {
            get => _row["Сертификация"];
            set => _row["Сертификация"] = value;
        }
        public string Warranty
        {
            get => _row["Гарантия"];
            set => _row["Гарантия"] = value;
        }
        #endregion

        public void Fill(ProductItem product)
        {
            Category = product.Category;
            ProductGroup = product.ProductGroupRu;
            ArticleGardena = product.Article;
            Name = product.ArticleRu;
            EAN = product.CalendarGTIN;
            CountryOfOrigin = product.CalendarCountryOfOrigin;
            UnitOfMeasure = product.CalendarUnitOfMeasure;
            QuantityInMasterPack = product.CalendarQuantityInMasterPack;
            ArticleGrossWeight = product.CalendarArticleGrossWeight;
            ArticleNetWeight = product.CalendarArticleNetWeight;
            PackagingLength = product.CalendarPackagingLength;
            PackagingWidth = product.CalendarPackagingWidth;
            PackagingHeight = product.CalendarPackagingHeight;
            PackagingVolume = product.CalendarPackagingVolume;
            UnitsPerPallet = product.CalendarUnitsPerPallet;
            Certificate = product.LocalCertificate;
        }
    }
}
