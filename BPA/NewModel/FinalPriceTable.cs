using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace BPA.NewModel
{
    class FinalPriceTable : IEnumerable<FinalPriceItem>
    {
        const string SHEET = "Прайс лист";
        const string TABLE = "Прайс_лист";
        public string SheetName => SHEET;

        WS_DB db = new WS_DB();
        Excel.ListObject _table = null;

        public FinalPriceTable()
        {
            Excel.Workbook wb = Globals.ThisWorkbook.InnerObject;
            Excel.Worksheet ws = wb.Sheets[SHEET];
            _table = ws.ListObjects[TABLE];
        }

        public FinalPriceItem Find(Predicate<FinalPriceItem> predicate)
        {
            foreach (FinalPriceItem item in this)
                if (predicate(item)) return item;

            return null;
        }

        public FinalPriceItem Add()
        {
            int row = db.AddRow();
            FinalPriceItem item = new FinalPriceItem(db[row]);
            item.Id = db.NextID("Id");
            return item;
        }

        public IEnumerator<FinalPriceItem> GetEnumerator()
        {
            foreach (TableRow item in db) yield return new FinalPriceItem(item);
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
