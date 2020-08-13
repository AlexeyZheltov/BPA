using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace BPA.NewModel
{
    class DiscountTable : IEnumerable<DiscountItem>
    {
        const string SHEET = "РРЦ";
        const string TABLE = "РРЦ";

        WS_DB db = new WS_DB();
        Excel.ListObject _table = null;

        public DiscountTable()
        {
            Excel.Workbook wb = Globals.ThisWorkbook.InnerObject;
            Excel.Worksheet ws = wb.Sheets[SHEET];
            _table = ws.ListObjects[TABLE];
        }

        public DiscountItem Find(Predicate<DiscountItem> predicate)
        {
            foreach (DiscountItem item in this)
                if (predicate(item)) return item;

            return null;
        }

        public DiscountItem Add()
        {
            int row = db.AddRow();
            DiscountItem item = new DiscountItem(db[row]);
            item.Id = db.NextID("№");
            return item;
        }

        public IEnumerator<DiscountItem> GetEnumerator()
        {
            foreach (TableRow item in db) yield return new DiscountItem(item);
        }

        public int Load()
        {
            db.Load(_table);
            return db.RowCount();
        }

        public void Save() => db.Save();

        IEnumerator IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }
    }
}
