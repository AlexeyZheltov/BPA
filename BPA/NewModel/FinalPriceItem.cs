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

        public int Id
        {
            get => _row["№"];
            set => _row["№"] = value;
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

        public double QuantityInMasterPack
        {
            get => _row["кол-во в мастер-паке"];
            set => _row["кол-во в мастер-паке"] = value;
        }

        public double ArticleGrossWeight
        {
            get => _row["Вес гросс штуки, финальный"];
            set => _row["Вес гросс штуки, финальный"] = value;
        }


        public double ArticleNetWeight
        {
            get => _row["Вес нетто штуки, финальный"];
            set => _row["Вес нетто штуки, финальный"] = value;
        }

        public double ArticleNetWeight
        {
            get => _row["Вес нетто штуки, финальный"];
            set => _row["Вес нетто штуки, финальный"] = value;
        }

        public double PackagingLength
        {
            get => _row["Длина упаковки"];
            set => _row["Длина упаковки"] = value;
        }
        public double PackagingWidth
        {
            get => _row["Ширина упаковки"];
            set => _row["Ширина упаковки"] = value;
        }
        public double PackagingHeight
        {
            get => _row["Высота упаковки"];
            set => _row["Высота упаковки"] = value;
        }
        public double PackagingVolume
        {
            get => _row["Объем упаковки"];
            set => _row["Объем упаковки"] = value;
        }
        public double UnitsPerPallet
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
    }
}
