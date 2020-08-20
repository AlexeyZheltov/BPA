using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
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

        WS_DB _db = new WS_DB();
        Excel.ListObject _table = null;

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
            foreach (PlanningNewYearItem planning in this)
                if (predicate(planning)) return planning;
            return null;
        }

        public PlanningNewYearItem Add()
        {
            int row = _db.AddRow();
            PlanningNewYearItem item = new PlanningNewYearItem(_db[row]);
            item.Id = _db.NextID("№");
            return item;
        }

    }
}
