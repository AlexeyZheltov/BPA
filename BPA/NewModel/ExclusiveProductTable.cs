using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace BPA.NewModel
{
    class ExclusiveProductTable : IEnumerable<ExclusiveProductItem>
    {
        const string SHEET = "Эксклюзивность";
        const string TABLE = "Эксклюзивность";

        WS_DB db = new WS_DB();
        Excel.ListObject _table = null;

        public ExclusiveProductTable()
        {
            Excel.Workbook wb = Globals.ThisWorkbook.InnerObject;
            Excel.Worksheet ws = wb.Sheets[SHEET];
            _table = ws.ListObjects[TABLE];
        }

        public ExclusiveProductItem Find(Predicate<ExclusiveProductItem> predicate)
        {
            foreach (ExclusiveProductItem item in this)
                if (predicate(item)) return item;

            return null;
        }

        public ExclusiveProductItem Add()
        {
            int row = db.AddRow();
            ExclusiveProductItem item = new ExclusiveProductItem(db[row]);
            item.Id = db.NextID("№");
            return item;
        }

        public IEnumerator<ExclusiveProductItem> GetEnumerator()
        {
            foreach (TableRow item in db) yield return new ExclusiveProductItem(item);
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
