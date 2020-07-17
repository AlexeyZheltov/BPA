using BPA.Model;

using Microsoft.Office.Interop.Excel;

using System;
using System.IO;
using System.Windows.Forms;

namespace BPA.Modules
{
    internal class FileCalendar
    {
        private readonly string FileName = "";
        private readonly Microsoft.Office.Interop.Excel.Application Application = Globals.ThisWorkbook.Application;
        private readonly int CalendarHeaderRow = 6;

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

        public int CountActions => LastRow - CalendarHeaderRow;
        private bool IsCancel = false;

        public bool IsOpen { get; set; } = false;

        public Workbook Workbook
        {
            get
            {
                if (_Workbook == null)
                {
                    try
                    {
                        _Workbook = Application.Workbooks.Open(FileName);
                    }
                    catch
                    {
                        _Workbook = null;
                    }
                }
                return _Workbook;
            }
            set
            {
                _Workbook = value;
            }
        }
        private Workbook _Workbook;

        private Worksheet worksheet => Workbook?.Sheets[1];

        public int LastRow
        {
            get
            {
                if (_LastRow == 0) _LastRow = worksheet.UsedRange.Row + worksheet.UsedRange.Rows.Count - 1;
                return _LastRow;
            }
        }
        private int _LastRow = 0;

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
            BPASettings settings = new BPASettings();
            
            if (settings.GetProductCalendarPath(out string path))
            {
                FileName = path;
                IsOpen = true;
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
            FileName = filename;
        }

        public FileCalendar(Workbook workbook)
        {
            Workbook = workbook;
        }

        public void LoadCalendar()
        {
            if (Workbook == null) return;

            if (!ReadCalendarLoad())
            {
                MessageBox.Show("Значимых записей не найдено", "BPA", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            ProductCalendar productCalendar = new ProductCalendar
            {
                Name = Workbook.Name,
                Path = FileName
            };
            productCalendar.Save();

            IsCancel = true;
        }

        /// <summary>
        /// загрузка календаря
        /// </summary>
        private bool ReadCalendarLoad()
        {
            Product product = null;

            for (int rw = CalendarHeaderRow + 1; rw < LastRow; rw++)
            {
                if (IsCancel) return false;
                ActionStart?.Invoke($"Обрабатывается строка {rw}");

                if (worksheet.Cells[rw, 1].Text == "")
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
                string tobesold = worksheet.Cells[rw, ToBeSoldInColumn].Text;

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

                product = new Product().GetProduct(GetValueFromColumn(rw, LocalIDGardenaColumn));

                if (product != null)
                {
                    product = CreateProduct(rw, product);
                    product.Calendar = Workbook.Name;
                    product.Update();
                    product.Mark("Article");
                    product.Mark("PNS");
                    product.Mark("Calendar");
                }
                else
                {
                    product = CreateProduct(rw, new Product());
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
            int rowNumber = FindRow(articul);
            if (rowNumber == 0) return null;

            return CreateProduct(rowNumber, new Product());
        }


        /// <summary>
        /// получение данных из календаря
        /// </summary>
        /// <param name="row"></param>
        /// <returns></returns>
        private Product CreateProduct(int row, Product product)
        {
            if (DateTime.TryParse(GetValueFromColumn(row, SalesStartDateColumn), out DateTime tmpDateTime))
            {
                if (tmpDateTime.ToOADate() > 0)
                    product.CalendarSalesStartDate = tmpDateTime.ToOADate();
            }

            if (DateTime.TryParse(GetValueFromColumn(row, PreliminaryEliminationDateColumn), out tmpDateTime))
            {
                if (tmpDateTime.ToOADate() > 0)
                    product.CalendarPreliminaryEliminationDate = tmpDateTime.ToOADate();

            }

            if (DateTime.TryParse(GetValueFromColumn(row, EliminationDateColumn), out tmpDateTime))
            {
                if (tmpDateTime.ToOADate() > 0)
                    product.CalendarEliminationDate = tmpDateTime.ToOADate();
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

            //product.GenericName = GetValueFromColumn(row, GenericNameColumn);
            product.Model = GetValueFromColumn(row, ModelColumn);
            product.SubGroup = GetValueFromColumn(row, SubgroupColumn);
            //product.ProductGroup = GetValueFromColumn(row, ProductGroupColumn);
            product.PNS = GetValueFromColumn(row, IdColumn);


            if (Double.TryParse(GetValueFromColumn(row, IRPRRPColumn), out Double tmpDouble))
            {
                product.IRP = tmpDouble;
            }

            return product;
        }

        /// <summary>
        /// получение номена строки по имени заголовка
        /// </summary>
        /// <param name="fildName"></param>
        /// <returns></returns>
        private int FindColumn(string fildName)
        {
            return worksheet.Cells.Find(fildName, LookAt: XlLookAt.xlWhole)?.Column ?? 0;
        }

        private int FindRow(string articul)
        {
            return worksheet.Cells.Find(articul, LookAt: XlLookAt.xlWhole)?.Row ?? 0;
        }

        /// <summary>
        /// получение значения из строки по номеру столбца
        /// </summary>
        /// <param name="rw"></param>
        /// <param name="col"></param>
        /// <returns></returns>
        private string GetValueFromColumn(int rw, int col)
        {
            return col != 0 ? worksheet.Cells[rw, col].value?.ToString() : "";
        }

        public void Close()
        {
            IsOpen = false;
            Workbook.Close(false);
        }

        public void Cancel()
        {
            IsCancel = true;
        }
    }
}
