using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace BPA.NewModel
{
    class STKTable : IEnumerable<STKItem>
    {
        const string SHEET = "STK";
        const string TABLE = "STK";

        WS_DB _db = new WS_DB();
        Excel.ListObject _table = null;

        public STKTable()
        {
            Excel.Workbook wb = Globals.ThisWorkbook.InnerObject;
            Excel.Worksheet ws = wb.Sheets[SHEET];
            _table = ws.ListObjects[TABLE];
        }

        public IEnumerator<STKItem> GetEnumerator()
        {
            foreach (TableRow item in _db) yield return new STKItem(item);
        }

        public int Load()
        {
            _db.Load(_table);
            return _db.RowCount();
        }

        public void Save() => _db.Save();

        public int Count => _db.RowCount();

        IEnumerator IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }

        public STKItem Find(Predicate<STKItem> predicate)
        {
            foreach (STKItem item in this)
                if (predicate(item)) return item;
            return null;
        }

        public STKItem Add()
        {
            int row = _db.AddRow();
            STKItem item = new STKItem(_db[row]);
            item.Id = _db.NextID("№");
            return item;
        }
    }
}
