using BPA.Model;
using Microsoft.Office.Interop.Excel;
using BPA.Forms;
using System;
using System.IO;
using System.Windows.Forms;
using System.Windows.Media.Animation;

namespace BPA.Modules
{
    internal class FileCalendar : FileBase
    {
        private int fileHeaderRow = 6;

        #region --- Columns ---
        public int IdColumn
        {
            get
            {
                if (_IdColumn == 0) _IdColumn = FindColumn("<ID>");
                return _IdColumn;
            }
}
        private int _IdColumn = 0;
        public int LocalIDGardenaColumn
        {
            get
            {
                if (_LocalIDGardenaColumn == 0) _LocalIDGardenaColumn = FindColumn("Local ID Gardena");
                return _LocalIDGardenaColumn;
            }
        }
        private int _LocalIDGardenaColumn = 0;
        public int GenericNameColumn
        {
            get
            {
                if (_GenericNameColumn == 0) _GenericNameColumn = FindColumn("Generic Name (long)");
                return _GenericNameColumn;
            }
        }
        private int _GenericNameColumn = 0;
        public int ModelColumn
        {
            get
            {
                if (_ModelColumn == 0) _ModelColumn = FindColumn("Model (only integration)");
                return _ModelColumn;
            }
        }
        private int _ModelColumn = 0;
        public int SubgroupColumn
        {
            get
            {
                if (_SubgroupColumn == 0) _SubgroupColumn = FindColumn("Subgroup ClassRef ID (only integration)");
                return _SubgroupColumn;
            }
        }
        private int _SubgroupColumn = 0;
        public int ProductGroupColumn
        {
            get
            {
                if (_ProductGroupColumn == 0) _ProductGroupColumn = FindColumn("Product group Alpha (only integration)");
                return _ProductGroupColumn;
            }
        }
        private int _ProductGroupColumn = 0;
        public int IRPRRPColumn
        {
            get
            {
                if (_IRPRRPColumn == 0) _IRPRRPColumn = FindColumn("IRP-RRP");
                return _IRPRRPColumn;
            }
        }
        private int _IRPRRPColumn = 0;
        public int IRPNetColumn
        {
            get
            {
                if (_IRPNetColumn == 0) _IRPNetColumn = FindColumn("IRP-Net");
                return _IRPNetColumn;
            }
        }
        private int _IRPNetColumn = 0;
        public int ShortDescriptionColumn
        {
            get
            {
                if (_ShortDescriptionColumn == 0) _ShortDescriptionColumn = FindColumn("Short Description");
                return _ShortDescriptionColumn;
            }
        }
        private int _ShortDescriptionColumn = 0;
        public int TechnicalPlatformColumn
        {
            get
            {
                if (_TechnicalPlatformColumn == 0) _TechnicalPlatformColumn = FindColumn("Technical Platform");
                return _TechnicalPlatformColumn;
            }
        }
        private int _TechnicalPlatformColumn = 0;
        public int VariantDescriptionColumn
        {
            get
            {
                if (_VariantDescriptionColumn == 0) _VariantDescriptionColumn = FindColumn("Variant Description");
                return _VariantDescriptionColumn;
            }
        }
        private int _VariantDescriptionColumn = 0;
        public int ToBeSoldInColumn
        {
            get
            {
                if (_ToBeSoldInColumn == 0) _ToBeSoldInColumn = FindColumn("To be sold in");
                return _ToBeSoldInColumn;
            }
        }
        private int _ToBeSoldInColumn = 0;
        public int KeyAccountExclusiveForColumn
        {
            get
            {
                if (_KeyAccountExclusiveForColumn == 0) _KeyAccountExclusiveForColumn = FindColumn("Key Account exclusive for");
                return _KeyAccountExclusiveForColumn;
            }
        }
        private int _KeyAccountExclusiveForColumn = 0;
        public int SalesStartDateColumn
        {
            get
            {
                if (_SalesStartDateColumn == 0) _SalesStartDateColumn = FindColumn("Sales Start Date");
                return _SalesStartDateColumn;
            }
        }
        private int _SalesStartDateColumn = 0;
        public int PreliminaryEliminationDateColumn
        {
            get
            {
                if (_PreliminaryEliminationDateColumn == 0) _PreliminaryEliminationDateColumn = FindColumn("Preliminary Elimination Date");
                return _PreliminaryEliminationDateColumn;
            }
        }
        private int _PreliminaryEliminationDateColumn = 0;
        public int EliminationDateColumn
        {
            get
            {
                if (_EliminationDateColumn == 0) _EliminationDateColumn = FindColumn("Elimination Date");
                return _EliminationDateColumn;
            }
        }
        private int _EliminationDateColumn = 0;
        public int PredecessorProductReferenceColumn
        {
            get
            {
                if (_PredecessorProductReferenceColumn == 0) _PredecessorProductReferenceColumn = FindColumn("Predecessor Product Reference");
                return _PredecessorProductReferenceColumn;
            }
        }
        private int _PredecessorProductReferenceColumn = 0;
        public int GTIN13Column
        {
            get
            {
                if (_GTIN13Column == 0) _GTIN13Column = FindColumn("GTIN-13/EAN");
                return _GTIN13Column;
            }
        }
        private int _GTIN13Column = 0;
        public int GTIN12Column
        {
            get
            {
                if (_GTIN12Column == 0) _GTIN12Column = FindColumn("GTIN-12/UPC-A");
                return _GTIN12Column;
            }
        }
        private int _GTIN12Column = 0;
        public int CurrentProducingFactoryColumn
        {
            get
            {
                if (_CurrentProducingFactoryColumn == 0) _CurrentProducingFactoryColumn = FindColumn("Current Producing Factory Entity Reference");
                return _CurrentProducingFactoryColumn;
            }
        }
        private int _CurrentProducingFactoryColumn = 0;
        public int CountryOfOriginColumn
        {
            get
            {
                if (_CountryOfOriginColumn == 0) _CountryOfOriginColumn = FindColumn("Country of Origin");
                return _CountryOfOriginColumn;
            }
        }
        private int _CountryOfOriginColumn = 0;
        public int ArticleManagerColumn
        {
            get
            {
                if (_ArticleManagerColumn == 0) _ArticleManagerColumn = FindColumn("Article manager");
                return _ArticleManagerColumn;
            }
        }
        private int _ArticleManagerColumn = 0;
        public int UnitOfMeasureColumn
        {
            get
            {
                if (_UnitOfMeasureColumn == 0) _UnitOfMeasureColumn = FindColumn("Unit of measure");
                return _UnitOfMeasureColumn;
            }
        }
        private int _UnitOfMeasureColumn = 0;
        public int QuantityInMasterPackColumn
        {
            get
            {
                if (_QuantityInMasterPackColumn == 0) _QuantityInMasterPackColumn = FindColumn("Quantity in Master pack");
                return _QuantityInMasterPackColumn;
            }
        }
        private int _QuantityInMasterPackColumn = 0;
        public int ArticleGrossWeightPreliminaryColumn
        {
            get
            {
                if (_ArticleGrossWeightPreliminaryColumn == 0) _ArticleGrossWeightPreliminaryColumn = FindColumn("Article gross weight, preliminary");
                return _ArticleGrossWeightPreliminaryColumn;
            }
        }
        private int _ArticleGrossWeightPreliminaryColumn = 0;
        public int ArticleGrossWeightColumn
        {
            get
            {
                if (_ArticleGrossWeightColumn == 0) _ArticleGrossWeightColumn = FindColumn("Article gross weight");
                return _ArticleGrossWeightColumn;
            }
        }
        private int _ArticleGrossWeightColumn = 0;
        public int ArticleNetWeightPreliminaryColumn
        {
            get
            {
                if (_ArticleNetWeightPreliminaryColumn == 0) _ArticleNetWeightPreliminaryColumn = FindColumn("Article net weight, preliminary");
                return _ArticleNetWeightPreliminaryColumn;
            }
        }
        private int _ArticleNetWeightPreliminaryColumn = 0;
        public int ArticleNetWeightColumn
        {
            get
            {
                if (_ArticleNetWeightColumn == 0) _ArticleNetWeightColumn = FindColumn("Article net weight");
                return _ArticleNetWeightColumn;
            }
        }
        private int _ArticleNetWeightColumn = 0;
        public int PackagingLengthColumn
        {
            get
            {
                if (_PackagingLengthColumn == 0) _PackagingLengthColumn = FindColumn("Packaging length");
                return _PackagingLengthColumn;
            }
        }
        private int _PackagingLengthColumn = 0;
        public int PackagingWidthColumn
        {
            get
            {
                if (_PackagingWidthColumn == 0) _PackagingWidthColumn = FindColumn("Packaging width");
                return _PackagingWidthColumn;
            }
        }
        private int _PackagingWidthColumn = 0;
        public int PackagingHeightColumn
        {
            get
            {
                if (_PackagingHeightColumn == 0) _PackagingHeightColumn = FindColumn("Packaging height");
                return _PackagingHeightColumn;
            }
        }
        private int _PackagingHeightColumn = 0;
        public int PackagingVolumeColumn
        {
            get
            {
                if (_PackagingVolumeColumn == 0) _PackagingVolumeColumn = FindColumn("Packaging volume");
                return _PackagingVolumeColumn;
            }
        }
        private int _PackagingVolumeColumn = 0;
        public int ProductSizeLengthColumn
        {
            get
            {
                if (_ProductSizeLengthColumn == 0) _ProductSizeLengthColumn = FindColumn("Product size length");
                return _ProductSizeLengthColumn;
            }
        }
        private int _ProductSizeLengthColumn = 0;
        public int ProductSizeHeightColumn
        {
            get
            {
                if (_ProductSizeHeightColumn == 0) _ProductSizeHeightColumn = FindColumn("Product size height");
                return _ProductSizeHeightColumn;
            }
        }
        private int _ProductSizeHeightColumn = 0;
        public int ProductSizeWidthColumn
        {
            get
            {
                if (_ProductSizeWidthColumn == 0) _ProductSizeWidthColumn = FindColumn("Product size width");
                return _ProductSizeWidthColumn;
            }
        }
        private int _ProductSizeWidthColumn = 0;
        public int UnitsPerPalletColumn
        {
            get
            {
                if (_UnitsPerPalletColumn == 0) _UnitsPerPalletColumn = FindColumn("Units Per Pallet");
                return _UnitsPerPalletColumn;
            }
        }
        private int _UnitsPerPalletColumn = 0;
        #endregion

        public FileCalendar()
        {
            BPASettings settings = new BPASettings();
            
            if (settings.GetProductCalendarPath(out string path))
            {
                FileAddress = path;
                FileSheetName = "";
                FileHeaderRow = fileHeaderRow;

                IsOpen = true;
                SetFileData();
            }
            else
            {
                throw new ApplicationException("Загрузка отменена");
            }
        }

        public FileCalendar(string filename)
        {
            if (!File.Exists(filename))
            {
                throw new FileNotFoundException($"Файл {filename} не найден");
            }
            FileAddress = filename;
            IsOpen = true;
            FileHeaderRow = fileHeaderRow;

            SetFileData();
        }

        public FileCalendar(Workbook workbook)
        {
            Workbook = workbook;
            IsOpen = true;
            FileHeaderRow = fileHeaderRow;

            SetFileData();
        }

        //public void LoadCalendar()
        //{
        //    if (Workbook == null) return;

        //    if (!ReadCalendarLoad())
        //    {
        //        MessageBox.Show("Значимых записей не найдено", "BPA", MessageBoxButtons.OK, MessageBoxIcon.Information);
        //        return;
        //    }

        //    ProductCalendar productCalendar = new ProductCalendar
        //    {
        //        Name = Workbook.Name,
        //        Path = FileName
        //    };
        //    productCalendar.Save();

        //    IsCancel = true;
        //}

        //перенести часть функционала в Product
        /// <summary>
        /// загрузка календаря
        /// </summary>
        private bool ReadCalendarLoad()
        {
            Product product = null;

            for (int rowIndex = 2; rowIndex < ArrRrows; rowIndex++)
            {

                if (IsCancel) return false;
                ActionStart?.Invoke($"Обрабатывается строка {rowIndex}");

                if (GetValueFromColumnStr(rowIndex, 1) == "")
                {
                    ActionDone?.Invoke(1);
                    continue;
                }

                int temp = ToBeSoldInColumn;
                if (temp == 0)
                {
                    Close();
                    throw new ApplicationException("Файл имеет неверный формат");
                }
                string tobesold = GetValueFromColumnStr(rowIndex, ToBeSoldInColumn);

                if (!CheckToBeSold())
                {
                    ActionDone?.Invoke(1);
                    continue;
                }

                bool CheckToBeSold()
                {
                    if (tobesold.ToLower().Contains("without russia"))
                        return false;

                    if (tobesold.ToLower().Contains("global"))
                        return true;
                    if (tobesold.Contains("R4") || tobesold.Contains("R5"))
                        return true;
                    if (tobesold.Contains("RU") || tobesold.Contains("RUS"))
                        return true;
                    if (tobesold.ToLower().Contains("russia"))
                        return true;

                    return false;
                }

                string localIDGardena = GetValueFromColumnStr(rowIndex, LocalIDGardenaColumn);

                product = new Product().GetProduct(localIDGardena);

                if (product != null)
                {
                    product = CreateProduct(rowIndex, product);
                    product.Calendar = Workbook.Name;
                    product.Update();
                    product.Mark("Article");
                    product.Mark("PNS");
                    product.Mark("Calendar");
                }
                else
                {
                    product = CreateProduct(rowIndex, new Product());
                    product.Calendar = Workbook.Name;
                    product.Save();
                    product.Mark("Calendar");
                }

                ActionDone?.Invoke(1);
            }
            if (product == null) return false;
            product.Sort("Id");
            product.Sort("ProductGroup");
            return true;
        }

        public Product GetProduct(string articul)
        {
            int rowIndex = FindRow(articul, LocalIDGardenaColumn);
            if (rowIndex == 0) return null;

            return CreateProduct(rowIndex, new Product());
        }

        /// <summary>
        /// получение данных из календаря
        /// </summary>
        /// <param name="row"></param>
        /// <returns></returns>
        private Product CreateProduct(int rowIndex, Product product)
        {
            product.CalendarSalesStartDate = GetDoubleDateFromCell(rowIndex, SalesStartDateColumn);
            product.CalendarPreliminaryEliminationDate = GetDoubleDateFromCell(rowIndex, PreliminaryEliminationDateColumn);
            product.CalendarEliminationDate = GetDoubleDateFromCell(rowIndex, EliminationDateColumn);

            product.CalendarToBeSoldIn = GetValueFromColumnStr(rowIndex, ToBeSoldInColumn);
            product.CalendarGTIN = GetValueFromColumnStr(rowIndex, GTIN13Column);
            product.CalendarCurrentProducingFactoryEntityReference = GetValueFromColumnStr(rowIndex, CurrentProducingFactoryColumn);
            product.CalendarCountryOfOrigin = GetValueFromColumnStr(rowIndex, CountryOfOriginColumn);
            product.CalendarUnitOfMeasure = GetValueFromColumnStr(rowIndex, UnitOfMeasureColumn);
            product.CalendarQuantityInMasterPack = GetValueFromColumnStr(rowIndex, QuantityInMasterPackColumn);
            product.CalendarArticleGrossWeightPreliminary = GetValueFromColumnStr(rowIndex, ArticleGrossWeightPreliminaryColumn);
            product.CalendarArticleGrossWeight = GetValueFromColumnStr(rowIndex, ArticleGrossWeightColumn);
            product.CalendarArticleNetWeightPreliminary = GetValueFromColumnStr(rowIndex, ArticleNetWeightPreliminaryColumn);
            product.CalendarArticleNetWeight = GetValueFromColumnStr(rowIndex, ArticleNetWeightColumn);
            product.CalendarPackagingLength = GetValueFromColumnStr(rowIndex, PackagingLengthColumn);
            product.CalendarPackagingHeight = GetValueFromColumnStr(rowIndex, PackagingHeightColumn);
            product.CalendarPackagingWidth = GetValueFromColumnStr(rowIndex, PackagingWidthColumn);
            product.CalendarPackagingVolume = GetValueFromColumnStr(rowIndex, PackagingVolumeColumn);
            product.CalendarProductSizeHeight = GetValueFromColumnStr(rowIndex, ProductSizeHeightColumn);
            product.CalendarProductSizeWidth = GetValueFromColumnStr(rowIndex, ProductSizeWidthColumn);
            product.CalendarProductSizeLength = GetValueFromColumnStr(rowIndex, ProductSizeLengthColumn);
            product.CalendarUnitsPerPallet = GetValueFromColumnStr(rowIndex, UnitsPerPalletColumn);

            product.Article = GetValueFromColumnStr(rowIndex, LocalIDGardenaColumn);

            //product.GenericName = GetValueFromColumn(row, GenericNameColumn);
            product.Model = GetValueFromColumnStr(rowIndex, ModelColumn);
            product.SubGroup = GetValueFromColumnStr(rowIndex, SubgroupColumn);
            //product.ProductGroup = GetValueFromColumn(row, ProductGroupColumn);
            product.PNS = GetValueFromColumnStr(rowIndex, IdColumn);

            product.IRP = GetValueFromColumnDbl(rowIndex, IRPRRPColumn);

            return product;
        }
    }
}
