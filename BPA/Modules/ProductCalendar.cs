using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace BPA.Modules
{
    class ProductCalendar
    {
        readonly Microsoft.Office.Interop.Excel.Application ex = new Microsoft.Office.Interop.Excel.Application();
        Workbook WB;
        Worksheet ws;
        private int CalendarHeaderRow;
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
            if (WB != null)
            {
                ws = WB.Worksheets[1];
                CalendarHeaderRow = 6;
                int LastRow = ws.Cells[ws.Rows.Count, 1].End(XlDirection.xlUp).Row;


                int idColumn = FindColumn("");
                int LocalIDGardenaColumn = FindColumn("");
                int GenericNameColumn = FindColumn("");
                int ModelColumn = FindColumn("");
                int SubgroupColumn = FindColumn("");
                int ProductGroupColumn = FindColumn("");
                int IRPRRPColumn = FindColumn("");
                int IRPNetColumn = FindColumn("");
                int ShortDescriptionColumn = FindColumn("");
                int TechnicalPlatformColumn = FindColumn("");
                int VariantDescriptionColumn = FindColumn("");
                int TobesoldinColumn = FindColumn("");
                int KeyAccountExclusiveForColumn = FindColumn("");
                int SalesStartDateColumn = FindColumn("");
                int PreliminaryEliminationDateColumn = FindColumn("");
                int EliminationDateColumn = FindColumn("");
                int PredecessorProductReferenceColumn = FindColumn("");
                int GTIN13Column = FindColumn("");
                int GTIN12Column = FindColumn("");
                int CurrentProducingFactoryColumn = FindColumn("");
                int CountryOfOriginColumn = FindColumn("");
                int ArticleManagerColumn = FindColumn("");
                int UnitOfMeasureColumn = FindColumn("");
                int QuantityInMasterPackColumn = FindColumn("");
                int ArticleGrossWeightPreliminaryColumn = FindColumn("");
                int ArticleGrossWeightColumn = FindColumn("");
                int ArticleNetWeightPreliminaryColumn = FindColumn("");
                int ArticleNetWeightColumn = FindColumn("");
                int PackagingLengthColumn = FindColumn("");
                int PackagingWidthColumn = FindColumn("");
                int PackagingHeightColumn = FindColumn("");
                int PackagingVolumeColumn = FindColumn("");
                int ProductSizeLengthColumn = FindColumn("");
                int ProductSizeHeightColumn = FindColumn("");
                int ProductSizeWidthColumn = FindColumn("");
                int UnitsPerPalletColumn = FindColumn("");




                for (int rw = CalendarheaderRow; rw < LastRow; rw++)
                {
                    if (ws.Cells[1, rw].value != "")
                    {
                                                
                    }
                }
                

            }
        }

        private int FindColumn(string fildName)
        {
            FindColumn = 1
        }
    }
    
}
