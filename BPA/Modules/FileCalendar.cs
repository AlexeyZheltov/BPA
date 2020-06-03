using BPA.Model;

using Microsoft.Office.Interop.Excel;

using System.IO;
using System.Windows.Forms;

namespace BPA.Modules
{
    internal class FileCalendar
    {
        private readonly string FileName;
        private readonly Microsoft.Office.Interop.Excel.Application Application = Globals.ThisWorkbook.Application;

        private Workbook Workbook
        {
            get
            {
                if (_Workbook == null) _Workbook = Application.Workbooks.Open(FileName);
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

        public Product GetProduct(string articul)
        {
            int rowNumber = FindRow(articul);
            if (rowNumber == 0) return null;
            return GetProduct(rowNumber);
        }

        public Product GetProduct(int rowNumber)
        {
            Product product = new Product();
            //TODO Добавить заполнение свойств
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
    }
}
