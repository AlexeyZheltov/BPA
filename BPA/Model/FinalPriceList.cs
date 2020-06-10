using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

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
        /// КАТЕГОРИЯ
        /// </summary>
        public int Category
        {
            get; set;
        }

        /// <summary>
        /// Фото продукта
        /// </summary>
        public int Photo
        {
            get; set;
        }

        /// <summary>
        /// Продукт группа - название
        /// </summary>
        public int ProductGroup
        {
            get; set;
        }

        /// <summary>
        /// Артикул GARDENA
        /// </summary>
        public int ArticleGardena
        {
            get; set;
        }

        /// <summary>
        /// Артикул предшествующего продукта (если есть)
        /// </summary>
        public int ArticleOld
        {
            get; set;
        }

        /// <summary>
        /// Название
        /// </summary>
        public int Name
        {
            get; set;
        }

        /// <summary>
        /// Описание
        /// </summary>
        public int Description
        {
            get; set;
        }

        /// <summary>
        /// РРЦ 2020, руб. с НДС
        /// </summary>
        public int RRC
        {
            get; set;
        }

        /// <summary>
        /// EAN штуки
        /// </summary>
        public int EAN
        {
            get; set;
        }

        /// <summary>
        /// страна производства
        /// </summary>
        public int CountryOfOrigin
        {
            get; set;
        }

        /// <summary>
        /// единица измерения
        /// </summary>
        public int UnitOfMeasure
        {
            get; set;
        }

        /// <summary>
        /// кол-во в мастер-паке
        /// </summary>
        public int QuantityInMasterPack
        {
            get; set;
        }

        /// <summary>
        /// Вес гросс штуки, финальный
        /// </summary>
        public int ArticleGrossWeight
        {
            get; set;
        }

        /// <summary>
        /// Вес нетто штуки, финальный
        /// </summary>
        public int ArticleNetWeight
        {
            get; set;
        }

        /// <summary>
        /// Длина упаковки
        /// </summary>
        public int PackagingLength
        {
            get; set;
        }

        /// <summary>
        /// Ширина упаковки
        /// </summary>
        public int PackagingWidth
        {
            get; set;
        }

        /// <summary>
        /// Высота упаковки
        /// </summary>
        public int PackagingHeight
        {
            get; set;
        }

        /// <summary>
        /// Объем упаковки
        /// </summary>
        public int PackagingVolume
        {
            get; set;
        }

        /// <summary>
        /// Кол-во штук на паллете
        /// </summary>
        public int UnitsPerPallet
        {
            get; set;
        }

        /// <summary>
        /// Сертификация
        /// </summary>
        public int Certificate
        {
            get; set;
        }

        /// <summary>
        /// Гарантия
        /// </summary>
        public int Warranty
        {
            get; set;
        }
        
        #endregion
    }
}
