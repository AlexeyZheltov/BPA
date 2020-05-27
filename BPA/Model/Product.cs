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

        public override IDictionary<string, string> Filds {
            get {
                return _filds;
            }
        }
        private readonly Dictionary<string, string> _filds = new Dictionary<string, string>
        {
            { "Категория для прайс-листа диллеров", "Category" },
            { "Суперкатегория(ENG)", "SupercategoryEng" },
            { "Суперкатегория(RUS)", "SupercategoryRu" },
            { "Продукт группа", "ProductGroup" },
            { "Название продукт группы(ENG)", "ProductGroupEng" },
            { "Название продукт группы(RUS)", "ProductGroupRu" },
            { "SubGroup", "SubGroup" },
            { "Generic Name(long)", "GenericName" },
            { "Model", "Model" },
            { "PNS", "PNS" },
            { "Артикул", "Article" },
            { "Артикул предшественника(если есть)", "ArticleOld" },
            { "Название артикула(ENG)", "ArticleEng" },
            { "Название артикула(RUS)", "ArticleRu" },
            { "Используемый календарь", "Calendar" },
            { "Статус", "Status" },
            { "Эксклюзив клиента или канала продажи", "Exclusive" },
            { "Локальный сертификат", "LocalCertificate" }
        };

        /// <summary>
        /// Категория для прайс-листа диллеров
        /// </summary>
        public string Category {
            get; set;
        }
        /// <summary>
        /// Суперкатегория(ENG)
        /// </summary>
        public string SupercategoryEng {
            get; set;
        }
        /// <summary>
        /// Суперкатегория(RUS)
        /// </summary>
        public string SupercategoryRu {
            get; set;
        }
        /// <summary>
        /// Продукт группа
        /// </summary>
        public string ProductGroup {
            get; set;
        }
        /// <summary>
        /// Название продукт группы(ENG)
        /// </summary>
        public string ProductGroupEng {
            get; set;
        }
        /// <summary>
        /// Название продукт группы(RUS)
        /// </summary>
        public string ProductGroupRu {
            get; set;
        }
        /// <summary>
        /// SubGroup
        /// </summary>
        public string SubGroup {
            get; set;
        }
        /// <summary>
        /// Generic Name(long)
        /// </summary>
        public string GenericName {
            get; set;
        }
        /// <summary>
        /// Model
        /// </summary>
        public string Model {
            get; set;
        }
        /// <summary>
        /// PNS
        /// </summary>
        public string PNS {
            get; set;
        }
        /// <summary>
        /// Артикул
        /// </summary>
        public string Article {
            get; set;
        }
        /// <summary>
        /// Артикул предшественника(если есть)
        /// </summary>
        public string ArticleOld {
            get; set;
        }
        /// <summary>
        /// Название артикула(ENG)
        /// </summary>
        public string ArticleEng {
            get; set;
        }
        /// <summary>
        /// Название артикула(RUS)
        /// </summary>
        public string ArticleRu {
            get; set;
        }
        /// <summary>
        /// Используемый календарь
        /// </summary>
        public string Calendar {
            get; set;
        }
        /// <summary>
        /// Статус
        /// </summary>
        public string Status {
            get; set;
        }
        /// <summary>
        /// Эксклюзив клиента или канала продажи
        /// </summary>
        public string Exclusive {
            get; set;
        }
        /// <summary>
        /// Локальный сертификат
        /// </summary>
        public string LocalCertificate {
            get; set;
        }

    }
}



