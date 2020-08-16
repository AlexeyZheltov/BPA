using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace BPA.NewModel
{
    class ExclusiveMagTable : IEnumerable<ExclusiveMagItem>
    {
        const string SHEET = "Exclusives";
        const string TABLE = "Exclusives";

        WS_DB db = new WS_DB();
        Excel.ListObject _table = null;

        public ExclusiveMagTable()
        {
            Excel.Workbook wb = Globals.ThisWorkbook.InnerObject;
            Excel.Worksheet ws = wb.Sheets[SHEET];
            _table = ws.ListObjects[TABLE];
        }

        public ExclusiveMagItem Find(Predicate<ExclusiveMagItem> predicate)
        {
            foreach (ExclusiveMagItem item in this)
                if (predicate(item)) return item;

            return null;
        }

        public ExclusiveMagItem Add()
        {
            int row = db.AddRow();
            ExclusiveMagItem item = new ExclusiveMagItem(db[row]);
            item.Id = db.NextID("№");
            return item;
        }

        public IEnumerator<ExclusiveMagItem> GetEnumerator()
        {
            foreach (TableRow item in db) yield return new ExclusiveMagItem(item);
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
