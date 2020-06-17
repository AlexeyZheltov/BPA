using BPA.Forms;
using BPA.Modules;

using Microsoft.Office.Interop.Excel;

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Windows.Forms;

namespace BPA.Model
{
    /// <summary>
    /// Справочник Товаров
    /// </summary>
    internal class Product : TableBase
    {
        private readonly Microsoft.Office.Interop.Excel.Application Application = Globals.ThisWorkbook.Application;

        public override string TableName => "Товары";
        public override string SheetName => "Товары";

        #region --- Словарь ---

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
            { "GenericName", "Generic Name (long)" },
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
            { "LocalCertificate","Локальный сертификат" },

            { "IRP","IRP, Eur" },
            { "RRCCurrent","РРЦ текущий" },
            { "DIYCurrent","DIY текущий" },
            { "RRCPercent","Процент повышения РРЦ" },
            { "RRCCalculated","РРЦ расчетная, руб." },
            { "RRCFinal","РРЦ финальная, руб." },
            { "RRCEuro","РРЦ, евро" },
            { "IRPIndex","Индекс IRP" },
            { "DIYDiscount","Скидка DIY" },
            { "DIY","DIY price list, руб. без НДС" }
        };

        #endregion

        #region -- Основные свойства столбцов ---

        /// <summary>
        /// №
        /// </summary>
        public int Id
        {
            get; set;
        }
        /// <summary>
        /// Категория для прайс-листа диллеров
        /// </summary>
        public string Category
        {
            get; set;
        }
        /// <summary>
        /// Суперкатегория(ENG)
        /// </summary>
        public string SupercategoryEng
        {
            get; set;
        }
        /// <summary>
        /// Суперкатегория(RUS)
        /// </summary>
        public string SupercategoryRu
        {
            get; set;
        }
        /// <summary>
        /// Продукт группа
        /// </summary>
        public string ProductGroup
        {
            get; set;
        }
        /// <summary>
        /// Название продукт группы(ENG)
        /// </summary>
        public string ProductGroupEng
        {
            get; set;
        }
        /// <summary>
        /// Название продукт группы(RUS)
        /// </summary>
        public string ProductGroupRu
        {
            get; set;
        }
        /// <summary>
        /// SubGroup
        /// </summary>
        public string SubGroup
        {
            get; set;
        }
        /// <summary>
        /// Generic Name(long)
        /// </summary>
        public string GenericName
        {
            get; set;
        }
        /// <summary>
        /// Model
        /// </summary>
        public string Model
        {
            get; set;
        }
        /// <summary>
        /// PNS
        /// </summary>
        public string PNS
        {
            get; set;
        }
        /// <summary>
        /// Артикул
        /// </summary>
        public string Article
        {
            get; set;
        }
        /// <summary>
        /// Артикул предшественника(если есть)
        /// </summary>
        public string ArticleOld
        {
            get; set;
        }
        /// <summary>
        /// Название артикула(ENG)
        /// </summary>
        public string ArticleEng
        {
            get; set;
        }
        /// <summary>
        /// Название артикула(RUS)
        /// </summary>
        public string ArticleRu
        {
            get; set;
        }
        /// <summary>
        /// Используемый календарь
        /// </summary>
        public string Calendar
        {
            get; set;
        }
        /// <summary>
        /// Статус
        /// </summary>
        public string Status
        {
            get; set;
        }
        /// <summary>
        /// Эксклюзив клиента или канала продажи
        /// </summary>
        public string Exclusive
        {
            get; set;
        }
        /// <summary>
        /// Локальный сертификат
        /// </summary>
        public string LocalCertificate
        {
            get; set;
        }

        #endregion

        #region --- Свойства из Prod Calendar ---

        /// <summary>
        /// to be sold in
        /// </summary>
        public string CalendarToBeSoldIn
        {
            get; set;
        }

        /// <summary>
        /// Sales Start Date
        /// </summary>
        public Double CalendarSalesStartDate
        {
            get; set;
        }

        /// <summary>
        /// Preliminary Elimination Date
        /// </summary>
        public Double CalendarPreliminaryEliminationDate
        {
            get; set;
        }

        /// <summary>
        /// CalendarEliminationDate
        /// </summary>
        public Double CalendarEliminationDate
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

        #region --- Свойства для РРЦ ---

        /// <summary>
        /// IRP, Eur
        /// </summary>
        public Double IRP
        {
            get; set;
        }

        /// <summary>
        /// РРЦ текущий
        /// </summary>
        public Double RRCCurrent
        {
            get; set;
        }

        /// <summary>
        /// DIY текущий
        /// </summary>
        public Double DIYCurrent
        {
            get; set;
        }

        /// <summary>
        /// Процент повышения РРЦ
        /// </summary>
        public Double RRCPercent
        {
            get; set;
        }

        /// <summary>
        /// РРЦ расчетная, руб.
        /// </summary>
        public Double RRCCalculated
        {
            get; set;
        }

        /// <summary>
        /// РРЦ финальная, руб.
        /// </summary>
        public Double RRCFinal
        {
            get; set;
        }

        /// <summary>
        /// РРЦ, евро
        /// </summary>
        public Double RRCEuro
        {
            get; set;
        }

        /// <summary>
        /// Индекс IRP
        /// </summary>
        public Double IRPIndex
        {
            get; set;
        }

        /// <summary>
        /// Скидка DIY
        /// </summary>
        public Double DIYDiscount
        {
            get; set;
        }
        /// <summary>
        /// DIY price list, руб. без НДС
        /// </summary>
        public Double DIY
        {
            get; set;
        }
        #endregion

        public Product() { }
        public Product(ListRow listRow) => SetProperty(listRow);

        /// <summary>
        /// Дата повышения
        /// </summary>
        /// <returns></returns>
        public DateTime DateOfPromotion
        {
            get
            {
                string dateOfPromotionLabel = "Дата повышения";

                try
                {
                    Range dateCell = Table.DataBodyRange.Cells[1, 1].Parent.UsedRange.Find(dateOfPromotionLabel, LookAt: XlLookAt.xlWhole);
                    return DateTime.Parse(dateCell.Offset[0, 1].Text);
                }
                catch
                {
                    return new DateTime();
                }
            }

            set
            {
                _DateOfPromotion = value;
            }
        }
        private DateTime _DateOfPromotion;

        public Product GetProduct(string article)
        {

            ListRow listRow = GetRow("Article", article);
            if (listRow != null)
            {
                Product product = new Product(listRow);
                return product;
            }

            return null;
        }

        /// <summary>
        /// Получение данных продукта выбранной ячейки
        /// </summary>
        /// <returns></returns>
        public Product GetPoductActive()
        {
            if (Application.ActiveCell.Row <= FirstRow || Application.ActiveCell.Row >= LastRow)
                return null;

            ListRow listRow = Table.ListRows[Application.Selection[1].Row - Table.Range.Row];
            if (listRow != null)
            {
                Product product = new Product(listRow);
                return product;
            }
            return null;
        }

        /// <summary>
        /// Получение списка продуктов со всеми данными ДОЛГО!!!!
        /// </summary>
        public List<Product> GetProducts()
        {
            bool isCancel = false;
            void CancelLocal() => isCancel = true;
            ProcessBar processBar = new ProcessBar("Получение списка продуктов", LastRow);
            processBar.CancelClick += CancelLocal;
            processBar.Show();

            List<Product> products = new List<Product>();
            foreach (ListRow row in Table.ListRows)
            {
                if (isCancel)
                {
                    processBar.Close();
                    return null;
                }
                processBar.TaskStart($"Обрабатывается товар {row.Index} из {LastRow - Table.HeaderRowRange.Row}");

                products.Add(new Product(row));

                processBar.TaskDone(1);

            }
            processBar.Close();
            return products;
        }

        /// <summary>
        /// Получает список продуктов с укороченными данными
        /// </summary>
        /// <returns></returns>
        public List<Product> GetProductsLight()
        {
            bool isCancel = false;
            void CancelLocal() => isCancel = true;
            ProcessBar processBar = new ProcessBar("Получение списка продуктов", LastRow);
            processBar.CancelClick += CancelLocal;
            processBar.Show();

            List<Product> products = new List<Product>();
            foreach (ListRow row in Table.ListRows)
            {
                if (isCancel)
                {
                    processBar.Close();
                    return null;
                }
                processBar.TaskStart($"Обрабатывается товар {row.Index} из {LastRow - Table.HeaderRowRange.Row}");

                Product product = new Product()
                {
                    Id = (int)row.Range[1, Table.ListColumns[Filds["Id"]].Index].Value,
                    Category = row.Range[1, Table.ListColumns[Filds["Category"]].Index].Text,
                    Article = row.Range[1, Table.ListColumns[Filds["Article"]].Index].Text,
                    Calendar = row.Range[1, Table.ListColumns[Filds["Calendar"]].Index].Text
                };

                products.Add(product);
                processBar.TaskDone(1);
            }
            processBar.Close();
            return products;
        }
        /// <summary>
        /// Получает список продуктов с укороченными данными для обновления справочника РРЦ
        /// </summary>
        /// <returns></returns>
        public List<Product> GetProductsLightForRRC()
        {
            bool isCancel = false;
            void CancelLocal() => isCancel = true;
            ProcessBar processBar = new ProcessBar("Получение списка продуктов", LastRow);
            processBar.CancelClick += CancelLocal;
            processBar.Show();

            List<Product> products = new List<Product>();
            foreach (ListRow row in Table.ListRows)
            {
                if (isCancel)
                {
                    processBar.Close();
                    return null;
                }
                processBar.TaskStart($"Обрабатывается товар {row.Index} из {LastRow - Table.HeaderRowRange.Row}");

                Product product = new Product()
                {
                    Id = (int)row.Range[1, Table.ListColumns[Filds["Id"]].Index].Value,
                    Article = row.Range[1, Table.ListColumns[Filds["Article"]].Index].Text,
                    RRCCurrent = row.Range[1, Table.ListColumns[Filds["RRCCurrent"]].Index].Value ?? 0,
                    DIYCurrent = row.Range[1, Table.ListColumns[Filds["DIYCurrent"]].Index].Value ?? 0,
                    RRCCalculated = row.Range[1, Table.ListColumns[Filds["RRCCalculated"]].Index].Value ?? 0,
                    DIY = row.Range[1, Table.ListColumns[Filds["DIY"]].Index].Value ?? 0,
                    IRP = row.Range[1, Table.ListColumns[Filds["IRP"]].Index].Value ?? 0
                };
                products.Add(product);
                processBar.TaskDone(1);
            }
            processBar.Close();
            return products;
        }

        /// <summary>
        /// Устанавливает свойстка из продукт календаря
        /// </summary>
        public void SetFromCalendar(Workbook workbook)
        {
            FileCalendar fileCalendar = new FileCalendar(workbook);

            Product product = fileCalendar.GetProduct(Article);
            if (product == null)
                return;
            this.CalendarSalesStartDate = product.CalendarSalesStartDate;
            this.CalendarPreliminaryEliminationDate = product.CalendarPreliminaryEliminationDate;
            this.CalendarEliminationDate = product.CalendarEliminationDate;
            this.CalendarToBeSoldIn = product.CalendarToBeSoldIn;
            this.CalendarGTIN = product.CalendarGTIN;
            this.CalendarCurrentProducingFactoryEntityReference = product.CalendarCurrentProducingFactoryEntityReference;
            this.CalendarCountryOfOrigin = product.CalendarCountryOfOrigin;
            this.CalendarUnitOfMeasure = product.CalendarUnitOfMeasure;
            this.CalendarQuantityInMasterPack = product.CalendarQuantityInMasterPack;
            this.CalendarArticleGrossWeightPreliminary = product.CalendarArticleGrossWeightPreliminary;
            this.CalendarArticleGrossWeight = product.CalendarArticleGrossWeight;
            this.CalendarArticleNetWeightPreliminary = product.CalendarArticleNetWeightPreliminary;
            this.CalendarArticleNetWeight = product.CalendarArticleNetWeight;
            this.CalendarPackagingLength = product.CalendarPackagingLength;
            this.CalendarPackagingHeight = product.CalendarPackagingHeight;
            this.CalendarPackagingWidth = product.CalendarPackagingWidth;
            this.CalendarPackagingVolume = product.CalendarPackagingVolume;
            this.CalendarProductSizeHeight = product.CalendarProductSizeHeight;
            this.CalendarProductSizeWidth = product.CalendarProductSizeWidth;
            this.CalendarProductSizeLength = product.CalendarProductSizeLength;
            this.CalendarUnitsPerPallet = product.CalendarUnitsPerPallet;

            this.GenericName = product.GenericName;
            this.Model = product.Model;
            this.SubGroup = product.SubGroup;
            this.ProductGroup = product.ProductGroup;
            this.PNS = product.PNS;

            this.Calendar = this.Calendar;

            Update();
        }

        public void UpdatePriceFromRRC(RRC rrc)
        {
            if (rrc != null)
            {
                this.RRCCurrent = rrc.RRCNDS;
                this.DIYCurrent = rrc.DIY;
            }
            else
            {
                this.RRCCurrent = 0;
                this.DIYCurrent = 0;
            }

            Update();
        }

        private bool IsCancel = false;
        /// <summary>
        /// Событие начала задачи
        /// </summary>
        public event ActionsStart ActionStart;
        public delegate void ActionsStart(string name);

        /// <summary>
        /// Событие завершения задачи
        /// </summary>
        public event ActionsDone ActionDone;
        public delegate void ActionsDone(int count);

        public void Cancel()
        {
            IsCancel = true;
        }

    }
}



