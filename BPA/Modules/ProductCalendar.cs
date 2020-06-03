using BPA.Forms;
using BPA.Model;

using Microsoft.Office.Interop.Excel;
using System;
using System.Windows.Forms;

namespace BPA.Modules
{
    class ProductCalendar
    {

        readonly Microsoft.Office.Interop.Excel.Application Application = Globals.ThisWorkbook.Application;
        Workbook WB;
        Worksheet ws;
        ProcessBar progress;

        private readonly string ToBeSoldInNeed = "text";

        int CalendarHeaderRow;
        int LastRow;

        //columns
        int idColumn;
        int LocalIDGardenaColumn;
        int GenericNameColumn;
        int ModelColumn;
        int SubgroupColumn;
        int ProductGroupColumn;
        int IRPRRPColumn;
        int IRPNetColumn;
        int ShortDescriptionColumn;
        int TechnicalPlatformColumn;
        int VariantDescriptionColumn;
        int ToBeSoldInColumn;
        int KeyAccountExclusiveForColumn;
        int SalesStartDateColumn;
        int PreliminaryEliminationDateColumn;
        int EliminationDateColumn;
        int PredecessorProductReferenceColumn;
        int GTIN13Column ;
        int GTIN12Column;
        int CurrentProducingFactoryColumn;
        int CountryOfOriginColumn;
        int ArticleManagerColumn;
        int UnitOfMeasureColumn;
        int QuantityInMasterPackColumn;
        int ArticleGrossWeightPreliminaryColumn;
        int ArticleGrossWeightColumn;
        int ArticleNetWeightPreliminaryColumn;
        int ArticleNetWeightColumn;
        int PackagingLengthColumn;
        int PackagingWidthColumn;
        int PackagingHeightColumn;
        int PackagingVolumeColumn;
        int ProductSizeLengthColumn;
        int ProductSizeHeightColumn;
        int ProductSizeWidthColumn;
        int UnitsPerPalletColumn;
        //

        public Worksheet Worksheet 
        { 
            get
            {
                if (_Worksheet == null)
                {
                    _Worksheet = WB?.Worksheets[1];
                }
                return _Worksheet;
            }
            set
            {
                _Worksheet = value;
            }

        }
        private Worksheet _Worksheet;

        public void Open()
        {
            OpenFileDialog openFileDialog = new OpenFileDialog
            {
                InitialDirectory = Globals.ThisWorkbook.Application.ActiveWorkbook.Path,
                Filter = "Excel files (*.xls*)|*.xls*",
                Title = "Выберите файл календаря"
            };

            if (openFileDialog.ShowDialog() != DialogResult.OK)
            {
                WB = null;                
            }
            else
            {
                string filePath = openFileDialog.FileName;
                WB = Application.Workbooks.Open(filePath);
            }
        }

        private void Sets()
        {
            ws = Worksheet;
            CalendarHeaderRow = 6;
            SetColumns();
            
            LastRow = ws.Cells[ws.Rows.Count, 1].End(XlDirection.xlUp).Row;
        }

        /// <summary>
        /// получение номеров колонок в календаре
        /// </summary>
        private void SetColumns()
        {
            idColumn = FindColumn("<ID>");
            LocalIDGardenaColumn = FindColumn("Local ID Gardena");
            GenericNameColumn = FindColumn("Generic Name (long)");
            ModelColumn = FindColumn("Model (only integration)");
            SubgroupColumn = FindColumn("Subgroup ClassRef ID (only integration)");
            ProductGroupColumn = FindColumn("Product group Alpha (only integration)");
            IRPRRPColumn = FindColumn("IRP-RRP");
            IRPNetColumn = FindColumn("IRP-Net");
            ShortDescriptionColumn = FindColumn("Short Description");
            TechnicalPlatformColumn = FindColumn("Technical Platform");
            VariantDescriptionColumn = FindColumn("Variant Description");
            ToBeSoldInColumn = FindColumn("To be sold in");
            KeyAccountExclusiveForColumn = FindColumn("Key Account exclusive for");
            SalesStartDateColumn = FindColumn("Sales Start Date");
            PreliminaryEliminationDateColumn = FindColumn("Preliminary Elimination Date");
            EliminationDateColumn = FindColumn("Elimination Date");
            PredecessorProductReferenceColumn = FindColumn("Predecessor Product Reference");
            GTIN13Column = FindColumn("GTIN-13/EAN");
            GTIN12Column = FindColumn("GTIN-12/UPC-A");
            CurrentProducingFactoryColumn = FindColumn("Current Producing Factory Entity Reference");
            CountryOfOriginColumn = FindColumn("Country of Origin");
            ArticleManagerColumn = FindColumn("Article manager");
            UnitOfMeasureColumn = FindColumn("Unit of measure");
            QuantityInMasterPackColumn = FindColumn("Quantity in Master pack");
            ArticleGrossWeightPreliminaryColumn = FindColumn("Article gross weight, preliminary");
            ArticleGrossWeightColumn = FindColumn("Article gross weight");
            ArticleNetWeightPreliminaryColumn = FindColumn("Article net weight, preliminary");
            ArticleNetWeightColumn = FindColumn("Article net weight");
            PackagingLengthColumn = FindColumn("Packaging length");
            PackagingWidthColumn = FindColumn("Packaging width");
            PackagingHeightColumn = FindColumn("Packaging height");
            PackagingVolumeColumn = FindColumn("Packaging volume");
            ProductSizeLengthColumn = FindColumn("Product size length");
            ProductSizeHeightColumn = FindColumn("Product size height");
            ProductSizeWidthColumn = FindColumn("Product size width");
            UnitsPerPalletColumn = FindColumn("Units Per Pallet");
        }

        public void LoadCalendar()
        {
            Open();
            if (WB == null) return;
            
            Sets();

            progress = new ProcessBar("Заполнение документов", LastRow - CalendarHeaderRow + 1);
            progress.Show();

            ReadCalendarLoad();
            //ActiveWorkbook.save(true);
            //            ThisWorkbook.Save;
            //WB.Close(true);
            progress.Close();
        }

        public void UpdateCalendar()
        {
            Open();
            if (WB == null) return;

            Sets();

            progress = new ProcessBar("Заполнение документов", LastRow - CalendarHeaderRow + 1);
            progress.Show();

            ReadCalendarUpdate();

            progress.Close();
        }

        private void ReadCalendarUpdate()
        {
            for (int rw = CalendarHeaderRow + 1; rw < LastRow; rw++)
            {
                if (ws.Cells[rw, 1].value == "") continue;
                if (ws.Cells[rw, ToBeSoldInColumn].value != ToBeSoldInNeed) continue;
                
                Product product = new Product().GetProduct(GetValueFromColumn(rw, LocalIDGardenaColumn));
                if (product != null)
                {
                    product = CreateProduct(rw, product);
                    product.Save();
                }
                
            }
        }

        private void ReadCalendarLoad()
        {
            for (int rw = CalendarHeaderRow + 1; rw < LastRow; rw++)
            {
                progress.TaskStart($"Обрабатывается строка {rw}");
                if (progress.IsCancel) break;

                if (ws.Cells[rw, 1].value == "") continue;
                if (ws.Cells[rw, ToBeSoldInColumn].value != ToBeSoldInNeed) continue;

                Product product = CreateProduct(rw, new Product());
                
                product.Save();
                product.Mark("Article");

                if (rw == LastRow)
                {
                    product.Sort("ProductGroup");
                }
            }
        }

        /// <summary>
        /// получение данных из календаря
        /// </summary>
        /// <param name="rw"></param>
        /// <returns></returns>
        private Product CreateProduct(int rw, Product product)
        {
            //Product product = new Product().GetProduct(GetValueFromColumn(rw, LocalIDGardenaColumn));

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
            product.CalendarArticleGrossWeight = GetValueFromColumn( rw, ArticleGrossWeightColumn);
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
            product.PNS = GetValueFromColumn(rw, idColumn);

            return product;
        }

        /// <summary>
        /// получение номена строки по имени заголовка
        /// </summary>
        /// <param name="fildName"></param>
        /// <returns></returns>
        private int FindColumn(string fildName)
        {
            
            //Console.WriteLine(ws.Rows[CalendarHeaderRow].Find(fildName, LookAt: XlLookAt.xlWhole).Column);
            return ws.Rows[CalendarHeaderRow].Find(fildName, LookAt: XlLookAt.xlWhole)?.Column ?? 0;
            //if tne nothing?
        }

        /// <summary>
        /// получение значения из строки по номеру столбца
        /// </summary>
        /// <param name="rw"></param>
        /// <param name="col"></param>
        /// <returns></returns>
        private string GetValueFromColumn(int rw, int col)
        {
            return col != 0 ? ws.Cells[rw, col].value.ToString() : "";
        }

        /// <summary>
        /// обновление лиса РРЦ
        /// </summary>
        /// <param name="rw"></param>
        private void UpdatePrice(int rw)
        {
            string article = GetValueFromColumn(rw, LocalIDGardenaColumn);
            string dateStart = GetValueFromColumn(rw, SalesStartDateColumn);
            RRC rrc = new RRC().GetRRC(article, dateStart);

            if (rrc == null)
            {
                rrc = new RRC();
            }
            rrc.Article = GetValueFromColumn(rw, LocalIDGardenaColumn);
            rrc.IRP = GetValueFromColumn(rw, IRPRRPColumn);
            rrc.Date = GetValueFromColumn(rw, SalesStartDateColumn);

            rrc.Save();

        }
    }

}
