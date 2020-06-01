using BPA.Forms;
using BPA.Model;

using Microsoft.Office.Interop.Excel;
using System.Windows.Forms;

namespace BPA.Modules
{
    class ProductCalendar
    {
        readonly Microsoft.Office.Interop.Excel.Application ex = new Microsoft.Office.Interop.Excel.Application();
        Workbook WB;
        Worksheet ws;
        int CalendarHeaderRow;

        private int CalendarheaderRow;

        public Workbook Open
        {
            get
            {
                OpenFileDialog openFileDialog = new OpenFileDialog
                {
                    InitialDirectory = Globals.ThisWorkbook.Application.ActiveWorkbook.Path,
                    Filter = "Excel files (*.xls*)|*.xls*",
                    Title = "Выберите файл календаря"
                };

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    string filePath = openFileDialog.FileName;
                    return WB = ex.Workbooks.Open(filePath);
                }
                else
                {
                    return null;
                }
            }
        }

        private void ReadCalendar(Workbook WB)
        {
            if (WB == null) return;


            ws = WB.Worksheets[1];
            CalendarHeaderRow = 6;
            int LastRow = ws.Cells[ws.Rows.Count, 1].End(XlDirection.xlUp).Row;

            
            ProcessBar progress = new ProcessBar("Заполнение документов", LastRow - CalendarheaderRow +1);
            progress.Show();


            for (int rw = CalendarheaderRow + 1; rw < LastRow; rw++)
            {
                progress.TaskStart($"Обрабатывается строка {rw}");
                if (progress.IsCancel) break;

                if (ws.Cells[rw, 1].value == "") continue;
                if (ws.Cells[rw, TobesoldinColumn].value == "") continue;
                //if tobesoldinrussia No be sold in Russia
                Product product = GetProduct(rw);

                //product.Insert();
                product.Save();

                //rrc

                RRC rrc = new RRC();
                string article = GetValueFromColumn(rw, LocalIDGardenaColumn);
                RRC rrcThis = rrc.GetRRC(article);

                if (rrcThis != null)
                {
                    rrc = rrcThis;
                }

                rrc.Article = GetValueFromColumn(rw, LocalIDGardenaColumn);
                rrc.IRP = GetValueFromColumn(rw, IRPRRPColumn);

               //rrc.Insert();
                rrc.Save();

            }

        }


        private Product GetProduct(int rw)
        {
            int idColumn = FindColumn("<ID>");
            int LocalIDGardenaColumn = FindColumn("Local ID Gardena");
            int GenericNameColumn = FindColumn("Generic Name (long)");
            int ModelColumn = FindColumn("Model (only integration)");
            int SubgroupColumn = FindColumn("Subgroup ClassRef ID (only integration)");
            int ProductGroupColumn = FindColumn("Product group Alpha (only integration)");
            int IRPRRPColumn = FindColumn("IRP-RRP");
            int IRPNetColumn = FindColumn("IRP-Net");
            int ShortDescriptionColumn = FindColumn("Short Description");
            int TechnicalPlatformColumn = FindColumn("Technical Platform");
            int VariantDescriptionColumn = FindColumn("Variant Description");
            int TobesoldinColumn = FindColumn("To be sold in");
            int KeyAccountExclusiveForColumn = FindColumn("Key Account exclusive for");
            int SalesStartDateColumn = FindColumn("Sales Start Date");
            int PreliminaryEliminationDateColumn = FindColumn("Preliminary Elimination Date");
            int EliminationDateColumn = FindColumn("Elimination Date");
            int PredecessorProductReferenceColumn = FindColumn("Predecessor Product Reference");
            int GTIN13Column = FindColumn("GTIN-13/EAN");
            int GTIN12Column = FindColumn("GTIN-12/UPC-A");
            int CurrentProducingFactoryColumn = FindColumn("Current Producing Factory Entity Reference");
            int CountryOfOriginColumn = FindColumn("Country of Origin");
            int ArticleManagerColumn = FindColumn("Article manager");
            int UnitOfMeasureColumn = FindColumn("Unit of measure");
            int QuantityInMasterPackColumn = FindColumn("Quantity in Master pack");
            int ArticleGrossWeightPreliminaryColumn = FindColumn("Article gross weight, preliminary");
            int ArticleGrossWeightColumn = FindColumn("Article gross weight");
            int ArticleNetWeightPreliminaryColumn = FindColumn("Article net weight, preliminary");
            int ArticleNetWeightColumn = FindColumn("Article net weight");
            int PackagingLengthColumn = FindColumn("Packaging length");
            int PackagingWidthColumn = FindColumn("Packaging width");
            int PackagingHeightColumn = FindColumn("Packaging height");
            int PackagingVolumeColumn = FindColumn("Packaging volume");
            int ProductSizeLengthColumn = FindColumn("Product size length");
            int ProductSizeHeightColumn = FindColumn("Product size height");
            int ProductSizeWidthColumn = FindColumn("Product size width");
            int UnitsPerPalletColumn = FindColumn("Units Per Pallet");


            Product product = new Product().GetProduct(GetValueFromColumn(rw, LocalIDGardenaColumn));

            product.CalendarToBeSoldIn =
                                        GetValueFromColumn(rw, TobesoldinColumn);
            product.CalendarSalesStartDate =
                                        GetValueFromColumn(rw, SalesStartDateColumn);
            product.CalendarPreliminaryEliminationDate =
                                        GetValueFromColumn(rw, PreliminaryEliminationDateColumn);
            product.CalendarEliminationDate =
                                        GetValueFromColumn(rw, EliminationDateColumn);
            product.CalendarGTIN =
                                        GetValueFromColumn(rw, GTIN13Column);
            product.CalendarCurrentProducingFactoryEntityReference =
                                        GetValueFromColumn(rw, CurrentProducingFactoryColumn);
            product.CalendarCountryOfOrigin =
                                        GetValueFromColumn(rw, CountryOfOriginColumn);
            product.CalendarUnitOfMeasure =
                                        GetValueFromColumn(rw, UnitOfMeasureColumn);
            product.CalendarQuantityInMasterPack =
                                        GetValueFromColumn(rw, QuantityInMasterPackColumn);
            product.CalendarArticleGrossWeightPreliminary =
                                        GetValueFromColumn(rw, ArticleGrossWeightPreliminaryColumn);
            product.CalendarArticleGrossWeight =
                                        GetValueFromColumn(rw, ArticleGrossWeightColumn);
            product.CalendarArticleNetWeightPreliminary =
                                        GetValueFromColumn(rw, ArticleNetWeightPreliminaryColumn);
            product.CalendarArticleNetWeight =
                                        GetValueFromColumn(rw, ArticleNetWeightColumn);
            product.CalendarPackagingLength =
                                        GetValueFromColumn(rw, PackagingLengthColumn);
            product.CalendarPackagingHeight =
                                        GetValueFromColumn(rw, PackagingHeightColumn);
            product.CalendarPackagingWidth =
                                        GetValueFromColumn(rw, PackagingWidthColumn);
            product.CalendarPackagingVolume =
                                        GetValueFromColumn(rw, PackagingVolumeColumn);
            product.CalendarProductSizeHeight =
                                        GetValueFromColumn(rw, ProductSizeHeightColumn);
            product.CalendarProductSizeWidth =
                                        GetValueFromColumn(rw, ProductSizeWidthColumn);
            product.CalendarProductSizeLength =
                                        GetValueFromColumn(rw, ProductSizeLengthColumn);
            product.CalendarUnitsPerPallet =
                                        GetValueFromColumn(rw, UnitsPerPalletColumn);

            //
            product.GenericName =
                                        GetValueFromColumn(rw, GenericNameColumn);
            product.Model =
                                        GetValueFromColumn(rw, ModelColumn);
            product.SubGroup =
                                        GetValueFromColumn(rw, SubgroupColumn);
            product.ProductGroup =
                                        GetValueFromColumn(rw, ProductGroupColumn);
            product.PNS =
                                        GetValueFromColumn(rw, idColumn);

            return product;
        }


        private int FindColumn(string fildName)
        {
            return ws.Rows[CalendarHeaderRow].Find(fildName, LookAt: XlLookAt.xlWhole)?.Column ?? 0;
            //if tne nothing?
        }


        private string GetValueFromColumn(int rw, int col)
        {
            return col != 0 ? GetValueFromColumn(rw, col) : "";
        }
    }

}
