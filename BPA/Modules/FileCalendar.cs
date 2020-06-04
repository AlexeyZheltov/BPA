using BPA.Forms;
using BPA.Model;

using Microsoft.Office.Interop.Excel;

using System;
using System.IO;
using System.Windows.Forms;

namespace BPA.Modules
{
    internal class FileCalendar
    {
        private readonly string FileName;
        private readonly Microsoft.Office.Interop.Excel.Application Application = Globals.ThisWorkbook.Application;
        private ProcessBar progress;
        private readonly string ToBeSoldInNeed = "RUSSIA";
        private readonly int CalendarHeaderRow = 6;

        private Workbook Workbook
        {
            get
            {
                if (_Workbook == null)
                    _Workbook = Application.Workbooks.Open(FileName);
                return _Workbook;
            }
            set
            {
                _Workbook = value;
            }
        }
        private Workbook _Workbook;

        private Worksheet Worksheet => Workbook?.Sheets[1];

        public int LastRow
        {
            get
            {
                return Worksheet.UsedRange.Row + Worksheet.UsedRange.Rows.Count - 1;
            }
        }

        #region --- Columns ---

        public int IdColumn => FindColumn("<ID>");
        public int LocalIDGardenaColumn => FindColumn("Local ID Gardena");
        public int GenericNameColumn => FindColumn("Generic Name (long)");
        public int ModelColumn => FindColumn("Model (only integration)");
        public int SubgroupColumn => FindColumn("Subgroup ClassRef ID (only integration)");
        public int ProductGroupColumn => FindColumn("Product group Alpha (only integration)");
        public int IRPRRPColumn => FindColumn("IRP-RRP");
        public int IRPNetColumn => FindColumn("IRP-Net");
        public int ShortDescriptionColumn => FindColumn("Short Description");
        public int TechnicalPlatformColumn => FindColumn("Technical Platform");
        public int VariantDescriptionColumn => FindColumn("Variant Description");
        public int ToBeSoldInColumn => FindColumn("To be sold in");
        public int KeyAccountExclusiveForColumn => FindColumn("Key Account exclusive for");
        public int SalesStartDateColumn => FindColumn("Sales Start Date");
        public int PreliminaryEliminationDateColumn => FindColumn("Preliminary Elimination Date");
        public int EliminationDateColumn => FindColumn("Elimination Date");
        public int PredecessorProductReferenceColumn => FindColumn("Predecessor Product Reference");
        public int GTIN13Column => FindColumn("GTIN-13/EAN");
        public int GTIN12Column => FindColumn("GTIN-12/UPC-A");
        public int CurrentProducingFactoryColumn => FindColumn("Current Producing Factory Entity Reference");
        public int CountryOfOriginColumn => FindColumn("Country of Origin");
        public int ArticleManagerColumn => FindColumn("Article manager");
        public int UnitOfMeasureColumn => FindColumn("Unit of measure");
        public int QuantityInMasterPackColumn => FindColumn("Quantity in Master pack");
        public int ArticleGrossWeightPreliminaryColumn => FindColumn("Article gross weight, preliminary");
        public int ArticleGrossWeightColumn => FindColumn("Article gross weight");
        public int ArticleNetWeightPreliminaryColumn => FindColumn("Article net weight, preliminary");
        public int ArticleNetWeightColumn => FindColumn("Article net weight");
        public int PackagingLengthColumn => FindColumn("Packaging length");
        public int PackagingWidthColumn => FindColumn("Packaging width");
        public int PackagingHeightColumn => FindColumn("Packaging height");
        public int PackagingVolumeColumn => FindColumn("Packaging volume");
        public int ProductSizeLengthColumn => FindColumn("Product size length");
        public int ProductSizeHeightColumn => FindColumn("Product size height");
        public int ProductSizeWidthColumn => FindColumn("Product size width");
        public int UnitsPerPalletColumn => FindColumn("Units Per Pallet");

        #endregion

        public FileCalendar()
        {

            using (OpenFileDialog fileDialog = new OpenFileDialog()
            {
                Title = "Выберите расположение продуктового календаря",
                DefaultExt = "*.xls*",
                CheckFileExists = true,
                InitialDirectory = Globals.ThisWorkbook.Path,
                ValidateNames = true,
                Multiselect = false,
                Filter = "Excel|*.xls*"
            })
            {
                if (fileDialog.ShowDialog() == DialogResult.OK)
                {
                    FileName = fileDialog.FileName;
                }
            }

        }

        public FileCalendar(string filename)
        {
            if (!File.Exists(filename))
            {
                throw new FileNotFoundException("Файл не найден");
            }
            FileName = filename;
        }

        public FileCalendar(Workbook workbook)
        {
            Workbook = workbook;
        }

        public void LoadCalendar()
        {
            if (Workbook == null) return;

            progress = new ProcessBar("Заполнение документов", LastRow - CalendarHeaderRow + 1);
            progress.Show();

            ReadCalendarLoad();

            progress.Close();
            
            //TODO: Добавить в таблицу календарей
        }

        /// <summary>
        /// загрузка календаря
        /// </summary>
        private void ReadCalendarLoad()
        {
            Product product = null;

            for (int rw = CalendarHeaderRow + 1; rw < LastRow; rw++)
            {
                progress.TaskStart($"Обрабатывается строка {rw}");
                if (progress.IsCancel) break;

                if (Worksheet.Cells[rw, 1].value == "") continue;

                string tobesold = Worksheet.Cells[rw, ToBeSoldInColumn].Text;
                tobesold = tobesold.ToUpper();

                if (!tobesold.Contains(ToBeSoldInNeed)) continue;

                product = new Product().GetProduct(GetValueFromColumn(rw, LocalIDGardenaColumn));

                if (product != null)
                {
                    product = CreateProduct(rw, product);
                    product.Mark("Article");
                    product.Mark("PNS");
                    product.Mark("Calendar");
                    product.Update();

                    Model.ProductCalendar productCalendar = new Model.ProductCalendar();
                    productCalendar.Name = Workbook.Name;
                    productCalendar.Path = Workbook.Path;
                    productCalendar.Save();
                }
                else
                {
                    product = CreateProduct(rw, new Product());
                    product.Mark("Calendar");
                    product.Save();
                }
            }
            product.Sort("Id");
            product.Sort("ProductGroup");
        }


        public Product GetProduct(string articul)
        {
            int rowNumber = FindRow(articul);
            if (rowNumber == 0) return null;
            return GetProduct(rowNumber);
        }

        public Product GetProduct(int row)
        {
            Product product = new Product();

            if (DateTime.TryParse(GetValueFromColumn(row, SalesStartDateColumn), out DateTime tmpDateTime))
            {
                product.CalendarSalesStartDate = tmpDateTime;
            }

            if (DateTime.TryParse(GetValueFromColumn(row, PreliminaryEliminationDateColumn), out tmpDateTime))
            {
                product.CalendarPreliminaryEliminationDate = tmpDateTime;
            }

            if (DateTime.TryParse(GetValueFromColumn(row, EliminationDateColumn), out tmpDateTime))
            {
                product.CalendarEliminationDate = tmpDateTime;
            }

            product.CalendarToBeSoldIn = GetValueFromColumn(row, ToBeSoldInColumn);
            product.CalendarGTIN = GetValueFromColumn(row, GTIN13Column);
            product.CalendarCurrentProducingFactoryEntityReference = GetValueFromColumn(row, CurrentProducingFactoryColumn);
            product.CalendarCountryOfOrigin = GetValueFromColumn(row, CountryOfOriginColumn);
            product.CalendarUnitOfMeasure = GetValueFromColumn(row, UnitOfMeasureColumn);
            product.CalendarQuantityInMasterPack = GetValueFromColumn(row, QuantityInMasterPackColumn);
            product.CalendarArticleGrossWeightPreliminary = GetValueFromColumn(row, ArticleGrossWeightPreliminaryColumn);
            product.CalendarArticleGrossWeight = GetValueFromColumn(row, ArticleGrossWeightColumn);
            product.CalendarArticleNetWeightPreliminary = GetValueFromColumn(row, ArticleNetWeightPreliminaryColumn);
            product.CalendarArticleNetWeight = GetValueFromColumn(row, ArticleNetWeightColumn);
            product.CalendarPackagingLength = GetValueFromColumn(row, PackagingLengthColumn);
            product.CalendarPackagingHeight = GetValueFromColumn(row, PackagingHeightColumn);
            product.CalendarPackagingWidth = GetValueFromColumn(row, PackagingWidthColumn);
            product.CalendarPackagingVolume = GetValueFromColumn(row, PackagingVolumeColumn);
            product.CalendarProductSizeHeight = GetValueFromColumn(row, ProductSizeHeightColumn);
            product.CalendarProductSizeWidth = GetValueFromColumn(row, ProductSizeWidthColumn);
            product.CalendarProductSizeLength = GetValueFromColumn(row, ProductSizeLengthColumn);
            product.CalendarUnitsPerPallet = GetValueFromColumn(row, UnitsPerPalletColumn);
            product.Article = GetValueFromColumn(row, LocalIDGardenaColumn);
            product.GenericName = GetValueFromColumn(row, GenericNameColumn);
            product.Model = GetValueFromColumn(row, ModelColumn);
            product.SubGroup = GetValueFromColumn(row, SubgroupColumn);
            product.ProductGroup = GetValueFromColumn(row, ProductGroupColumn);
            product.PNS = GetValueFromColumn(row, IdColumn);

            return product;
        }

        /// <summary>
        /// получение номена строки по имени заголовка
        /// </summary>
        /// <param name="fildName"></param>
        /// <returns></returns>
        private int FindColumn(string fildName)
        {
            return Worksheet.Cells.Find(fildName, LookAt: XlLookAt.xlWhole)?.Column ?? 0;
        }

        private int FindRow(string articul)
        {
            return Worksheet.Cells.Find(articul, LookAt: XlLookAt.xlWhole)?.Row ?? 0;
        }


        /// <summary>
        /// получение данных из календаря
        /// </summary>
        /// <param name="rw"></param>
        /// <returns></returns>
        private Product CreateProduct(int rw, Product product)
        {
            DateTime tmpDateTime;

            product.CalendarToBeSoldIn = GetValueFromColumn(rw, ToBeSoldInColumn);

            if (DateTime.TryParse(GetValueFromColumn(rw, SalesStartDateColumn), out tmpDateTime))
            {
                product.CalendarSalesStartDate = tmpDateTime;
            }

            if (DateTime.TryParse(GetValueFromColumn(rw, PreliminaryEliminationDateColumn), out tmpDateTime))
            {
                product.CalendarPreliminaryEliminationDate = tmpDateTime;
            }

            if (DateTime.TryParse(GetValueFromColumn(rw, EliminationDateColumn), out tmpDateTime))
            {
                product.CalendarEliminationDate = tmpDateTime;
            }

            product.CalendarGTIN = GetValueFromColumn(rw, GTIN13Column);
            product.CalendarCurrentProducingFactoryEntityReference = GetValueFromColumn(rw, CurrentProducingFactoryColumn);
            product.CalendarCountryOfOrigin = GetValueFromColumn(rw, CountryOfOriginColumn);
            product.CalendarUnitOfMeasure = GetValueFromColumn(rw, UnitOfMeasureColumn);
            product.CalendarQuantityInMasterPack = GetValueFromColumn(rw, QuantityInMasterPackColumn);
            product.CalendarArticleGrossWeightPreliminary = GetValueFromColumn(rw, ArticleGrossWeightPreliminaryColumn);
            product.CalendarArticleGrossWeight = GetValueFromColumn(rw, ArticleGrossWeightColumn);
            product.CalendarArticleNetWeightPreliminary = GetValueFromColumn(rw, ArticleNetWeightPreliminaryColumn);
            product.CalendarArticleNetWeight = GetValueFromColumn(rw, ArticleNetWeightColumn);
            product.CalendarPackagingLength = GetValueFromColumn(rw, PackagingLengthColumn);
            product.CalendarPackagingHeight = GetValueFromColumn(rw, PackagingHeightColumn);
            product.CalendarPackagingWidth = GetValueFromColumn(rw, PackagingWidthColumn);
            product.CalendarPackagingVolume = GetValueFromColumn(rw, PackagingVolumeColumn);
            product.CalendarProductSizeHeight = GetValueFromColumn(rw, ProductSizeHeightColumn);
            product.CalendarProductSizeWidth = GetValueFromColumn(rw, ProductSizeWidthColumn);
            product.CalendarProductSizeLength = GetValueFromColumn(rw, ProductSizeLengthColumn);
            product.CalendarUnitsPerPallet = GetValueFromColumn(rw, UnitsPerPalletColumn);

            //
            product.Article = GetValueFromColumn(rw, LocalIDGardenaColumn);

            product.GenericName = GetValueFromColumn(rw, GenericNameColumn);
            product.Model = GetValueFromColumn(rw, ModelColumn);
            product.SubGroup = GetValueFromColumn(rw, SubgroupColumn);
            product.ProductGroup = GetValueFromColumn(rw, ProductGroupColumn);
            product.PNS = GetValueFromColumn(rw, IdColumn);

            return product;
        }

        /// <summary>
        /// получение значения из строки по номеру столбца
        /// </summary>
        /// <param name="rw"></param>
        /// <param name="col"></param>
        /// <returns></returns>
        private string GetValueFromColumn(int rw, int col)
        {
            return col != 0 ? Worksheet.Cells[rw, col].value.ToString() : "";
        }


        public void Close()
        {
            Workbook.Close(false);
        }
    }
}
