using System;
using System.IO;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using BPA.Modules;
using System.Globalization;
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;
using SettingsBPA = BPA.Properties.Settings;

namespace BPA.NewModel
{
    class PlanningNewYearTable : IEnumerable<PlanningNewYearItem>
    {
        private string SHEET  => _TableWorksheetName != null ? _TableWorksheetName: templateSheetName;
        private string _TableWorksheetName;
        public readonly string templateSheetName = SettingsBPA.Default.SHEET_NAME_PLANNING_TEMPLATE;
        private string TABLE => GetTableName();
        private Excel.Workbook workbook;
        public string GetTableName()
        {
            Excel.ListObject table = workbook.Sheets[SHEET].ListObjects[1];
            return table.Name;
        }

        #region labels for tmpParams
        private const string CustomerStatusLabel = "Customer status";
        private const string ChannelTypeLabel = "Channel type";
        private const string YearLabel = "Период";
        private const string MaximumBonusLabel = "максмальный годовой бонус, %";
        #endregion

        #region tmpParams
        public string ChannelType;
        public string CustomerStatus;
        public int Year;
        public DateTime planningDate;
        public double MaximumBonus;

        public bool TmpSeted = false;
        #endregion

        WS_DB _db = new WS_DB();
        Excel.ListObject _table = null;

        public PlanningNewYearTable()
        {
            
        }

        public PlanningNewYearTable(string worksheetName)
        {
            if (worksheetName == templateSheetName)
                return;
            else
                _TableWorksheetName = worksheetName;

            this.workbook = Globals.ThisWorkbook.InnerObject;
            Excel.Worksheet ws = workbook.Sheets[SHEET];
            _table = ws.ListObjects[TABLE];
        }

        public PlanningNewYearTable(Excel.Workbook workbook, string worksheetName)
        {
            this.workbook = workbook;
           _TableWorksheetName = worksheetName;

            Excel.Worksheet ws = workbook.Sheets[SHEET];
            _table = ws.ListObjects[TABLE];
        }

        public IEnumerator<PlanningNewYearItem> GetEnumerator()
        {
            foreach (TableRow item in _db) yield return new PlanningNewYearItem(item);
        }
        IEnumerator IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }

        public int Load()
        {
            _db.Load(_table);
            return _db.RowCount();
        }

        public void Save() => _db.Save();

        public int Count => _db.RowCount();

        public void DelFirstRow()
        {
            _db.Delete(0);
        }

        public PlanningNewYearItem Find(Predicate<PlanningNewYearItem> predicate)
        {
            foreach (PlanningNewYearItem item in this)
                if (predicate(item)) return item;
            return null;
        }

        public PlanningNewYearItem Add()
        {
            int row = _db.AddRow();
            PlanningNewYearItem item = new PlanningNewYearItem(_db[row]);
            item.Id = _db.NextID("№");
            return item;
        }

        public DateTime CurrentDate = DateTime.Now;
        private int CurrentMonth => CurrentDate.Month;

        /// <summary>
        /// Возвращает шаблон с предварительно заполненными даннными
        /// </summary>
        /// <param name="worksheetName"></param>
        /// <returns></returns>
        public void SetTmpParams()
        {
            try
            {
                Excel.Range rng = workbook.Sheets[SHEET].UsedRange;
                
                this.CustomerStatus = getVal(CustomerStatusLabel);
                this.ChannelType = getVal(ChannelTypeLabel);
                if (int.TryParse(getVal(YearLabel), out int year))
                {
                    this.Year = year;
                    this.planningDate = new DateTime(year, 1, 1);
                }

                TmpSeted = true;

                string getVal(string label)
                {
                    try
                    {
                        Excel.Range cell = rng.Find(label, LookAt: Excel.XlLookAt.xlWhole);
                        return cell.Offset[0, 1].Text ?? "";
                    }
                    catch
                    {
                        return "";
                    }
                }
            }
            catch (Exception ex)
            {
                throw new ApplicationException(ex.Message);
            }
        }

        public bool HasData()
        {
            if (_table.ListRows.Count < 1)
                return false;

            return true;
        }

        /// <summary>
        /// Очистка таблицы
        /// </summary>
        public void ClearTable()
        {
            if (_table.ListRows.Count < 1) return;

            _table.DataBodyRange.Rows.Delete();
        }

        private readonly List<string> month_names = new List<string>()
        {
            "январь",
            "февраль",
            "март",
            "апрель",
            "май",
            "июнь",
            "июль",
            "август",
            "сентябрь",
            "октябрь",
            "ноябрь",
            "декабрь"
        };

        private static List<string> DelFormulColumnsList = new List<string>();

        public void DelFormulas()
        {
            StringBuilder stringBuilder = new StringBuilder();
            try
            {
                SetDelFormulaDict();

                stringBuilder.Append($"_table is null {_table == null}\n");
                if (_table.ListRows.Count < 1)
                {
                    _table.ListRows.Add();
                    _table.ListRows[2].Delete();
                }
                stringBuilder.Append($"_table rows count {_table.ListRows.Count}\n");

                foreach (string colName in DelFormulColumnsList)
                {
                    stringBuilder.Append($"colName {colName}\n");
                    int colNum = _table.ListColumns[colName].Range.Column;
                    stringBuilder.Append($"colNum {colNum}\n");
                    int rowNum = _table.DataBodyRange.Row;
                    stringBuilder.Append($"rowNum {colNum}\n");

                    if (colNum == 0 || rowNum == 0) continue;

                    Excel.Workbook wb = Globals.ThisWorkbook.InnerObject;
                    stringBuilder.Append($"wb is null {wb==null}\n");
                    Excel.Worksheet ws = wb.Sheets[SHEET];
                    stringBuilder.Append($"ws is null {ws == null}\n");
                    Excel.Range cell = ws.Cells[rowNum, colNum];
                    stringBuilder.Append($"cell is null {wb == null}\n");

                    cell.Value = "";
                }

                /// <summary>
                /// установка удаляемых столбцов
                /// </summary>
                /// <param name=""></param>
                void SetDelFormulaDict()
                {
                    DelFormulColumnsList.Clear();
                    int month = CurrentDate.Month;

                    for (int m = 0; m < 12; m++)
                        if (month > m)
                            DelFormulColumnsList.Add($"ИТОГО GS { month_names[m] }, шт.");

                    for (int m = 0; m < 12; m++)
                        if (month > m)
                            DelFormulColumnsList.Add($"ИТОГО NS { month_names[m] }, шт.");
                }
            }
            catch(Exception ex)
            {
                stringBuilder.Append($"\nException Message:\n{ex.Message}");
                using (StreamWriter stream = new StreamWriter("errorlog.txt"))
                {
                    stream.WriteLine(stringBuilder.ToString());
                }
                throw new ApplicationException($"Ошибка в поиске столбцов { SHEET }");
            }
        }

        public object[,] GetDataForPlanning()
        {
            Excel.Worksheet ws = workbook.Sheets[SHEET];
            Excel.Range rng = ws.Range[ws.Cells[_table.ListRows[1].Range.Row, _table.ListColumns[2].Range.Column], ws.Cells[_table.ListRows[_table.ListRows.Count].Range.Row, _table.ListColumns[_table.ListColumns.Count].Range.Column]];

            return rng.Value;
        }
    }
}
