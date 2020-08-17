using BPA.Modules;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace BPA.NewModel
{
    class ProductCalendarTable : IEnumerable<ProductItem>
    {
        const string SHEET = "Продуктовые календари";
        const string TABLE = "Продуктовые_календари";

        WS_DB db = new WS_DB();
        Excel.ListObject _table = null;

        public ProductCalendarTable()
        {
            Excel.Workbook wb = Globals.ThisWorkbook.InnerObject;
            Excel.Worksheet ws = wb.Sheets[SHEET];
            _table = ws.ListObjects[TABLE];
        }

        public IEnumerator<ProductCalendarItem> GetEnumerator()
        {
            foreach (TableRow item in db) yield return new ProductCalendarItem(item);
        }
        IEnumerator<ProductItem> IEnumerable<ProductItem>.GetEnumerator()
        {
            throw new NotImplementedException();
        }

        public int Load()
        {
            db.Load(_table);
            return db.RowCount();
        }

        public void Save() => db.Save();

        public int Count => db.RowCount();

        IEnumerator IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }

        public ProductCalendarItem Find(Predicate<ProductCalendarItem> predicate)
        {
            foreach (ProductCalendarItem productCalendar in this)
                if (predicate(productCalendar)) return productCalendar;
            return null;
        }

        public ProductCalendarItem Add()
        {
            int row = db.AddRow();
            ProductCalendarItem item = new ProductCalendarItem(db[row]);
            item.Id = db.NextID("№");
            return item;
        }
    }
}
