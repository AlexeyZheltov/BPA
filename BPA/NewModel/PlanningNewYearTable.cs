using System;
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
        public string GetTableName()
        {
            ThisWorkbook workbook = Globals.ThisWorkbook;
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

            Excel.Workbook wb = Globals.ThisWorkbook.InnerObject;
            Excel.Worksheet ws = wb.Sheets[SHEET];
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
                ThisWorkbook workbook = Globals.ThisWorkbook;
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
                        return cell.Offset[0, 1].Text;
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

            Range cell = _table.ListRows[1].Range[1, _table.ListColumns["№"].Index];
            //if (cell.Value == 0 || cell.Value == null)
            //    return false;

            return true;
        }

        /// <summary>
        /// Очистка таблицы
        /// </summary>
        public void ClearTable()
        {
            //_TableWorksheetName = worksheetName;

            if (_table.ListRows.Count < 1) return;

            _table.DataBodyRange.Rows.Delete();
        }

        private static Dictionary<string, bool> DelFormulaDict { get; set; } = new Dictionary<string, bool>();

        /// <summary>
        /// установка удаляемых столбцов
        /// </summary>
        /// <param name=""></param>
        private void SetDelFormulaDict()
        {
            DelFormulaDict.Clear();
            int month = CurrentDate.Month;

            DateTime dt = new DateTime();
            string m2 = $"{dt:MMMM}";

            //string cn = $"ИТОГО GS {m}, шт."
            //DateTime.
            //for(int m = 1; m < 13; m++)
            //    if(month >= m)
            //        DelFormulaDict.Add($"ИТОГО GS {CultureInfo.CurrentCulture.DateTimeFormat.GetMonthName(m).ToLower()}, шт.");

            //уязвимое место
            setIsDel("ИТОГО GS январь, шт.", 1);
            setIsDel("ИТОГО GS февраль, шт.", 2);
            setIsDel("ИТОГО GS март, шт.", 3);
            setIsDel("ИТОГО GS апрель, шт.", 4);
            setIsDel("ИТОГО GS май, шт.", 5);
            setIsDel("ИТОГО GS июнь, шт.", 6);
            setIsDel("ИТОГО GS июль, шт.", 7);
            setIsDel("ИТОГО GS август, шт.", 8);
            setIsDel("ИТОГО GS сентябрь, шт.", 9);
            setIsDel("ИТОГО GS октябрь, шт.", 10);
            setIsDel("ИТОГО GS ноябрь, шт.", 11);
            setIsDel("ИТОГО GS декабрь, шт.", 12);

            setIsDel("ИТОГО NS январь, шт.", 1);
            setIsDel("ИТОГО NS февраль, шт.", 2);
            setIsDel("ИТОГО NS март, шт.", 3);
            setIsDel("ИТОГО NS апрель, шт.", 4);
            setIsDel("ИТОГО NS май, шт.", 5);
            setIsDel("ИТОГО NS июнь, шт.", 6);
            setIsDel("ИТОГО NS июль, шт.", 7);
            setIsDel("ИТОГО NS август, шт.", 8);
            setIsDel("ИТОГО NS сентябрь, шт.", 9);
            setIsDel("ИТОГО NS октябрь, шт.", 10);
            setIsDel("ИТОГО NS ноябрь, шт.", 11);
            setIsDel("ИТОГО NS декабрь, шт.", 12);

            void setIsDel(string colName, int m)
            {
                bool isDel = (month >= m) ? true : false; 

                DelFormulaDict.Add(colName, isDel);
            }
        }

        public void DelFormulas()
        {
            try
            {
                SetDelFormulaDict();

                foreach (Excel.Range cell in _table.HeaderRowRange)
                
                {
                    if (!DelFormulaDict.ContainsKey(cell.Value)) continue;
                    if (DelFormulaDict[cell.Value] == false) continue;

                    //удаляем формулу
                    if (_table.ListRows.Count < 1)
                    {
                        PlanningNewYearItem planning = Add();
                        Save();
                    }
                    int idx = _table.ListColumns[cell.Value].Index;
                    Excel.Range rng = _table.ListRows[1].Range;
                    Range cell_1 = rng.Cells[1, idx];
                    cell_1.Value = "";
                }
            }
            catch
            {
                throw new ApplicationException($"Ошибка в поиске столбцов { SHEET }");
            }
        }
    }
}
