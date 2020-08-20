using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace BPA.NewModel
{
    class PlanTable : IEnumerable<PlanItem>
    {
        const string SHEET = "Планирование";
        const string TABLE = "Планирование";

        WS_DB _db = new WS_DB();
        Excel.ListObject _table = null;

        public string SheetName => SHEET;

        public PlanTable()
        {
            Excel.Workbook wb = Globals.ThisWorkbook.InnerObject;
            Excel.Worksheet ws = wb.Sheets[SHEET];
            _table = ws.ListObjects[TABLE];
        }


        IEnumerator IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }

        public IEnumerator<PlanItem> GetEnumerator()
        {
            foreach (TableRow item in _db) yield return new PlanItem(item);
        }

        public int Load()
        {
            _db.Load(_table);
            return _db.RowCount();
        }

        public void Save() => _db.Save();

        public int Count => _db.RowCount();

        public PlanItem Find(Predicate<PlanItem> predicate)
        {
            foreach (PlanItem item in this)
                if (predicate(item)) return item;
            return null;
        }

        public PlanItem Add()
        {
            int row = _db.AddRow();
            PlanItem item = new PlanItem(_db[row]);
            item.Id = _db.NextID("№");
            return item;
        }
    }
}
