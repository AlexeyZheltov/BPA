using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace BPA.NewModel
{
    class ProductTable : IEnumerable<ProductItem>
    {
        const string SHEET = "Товары";
        const string TABLE = "Товары";

        WS_DB db = new WS_DB();
        Excel.ListObject _table = null;

        public ProductTable()
        {
            Excel.Workbook wb = Globals.ThisWorkbook.InnerObject;
            Excel.Worksheet ws = wb.Sheets[SHEET];
            _table = ws.ListObjects[TABLE];
        }

        public IEnumerator<ProductItem> GetEnumerator()
        {
            foreach (TableRow item in db) yield return new ProductItem(item);
        }

        public int Load()
        {
            db.Load(_table);
            return db.RowCount();
        }

        public void Save() => db.Save();

        public int Count => db.RowCount();

        public DateTime DateOfPromotion()
        {
            string Label = "Дата повышения";

            try
            {
                Excel.Worksheet ws = _table.Parent;
                int i_row = _table.HeaderRowRange.Row - 1;
                Excel.Range rng = ws.Rows[i_row];
                rng = rng.Find(Label, LookAt: Excel.XlLookAt.xlWhole);
                rng = rng.Offset[0, 1];
                return DateTime.Parse(rng.Text);
            }
            catch
            {
                return new DateTime();
            }
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }
    }
}
