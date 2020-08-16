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

        WS_DB _db = new WS_DB();
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
            int row = _db.AddRow();
            DiscountItem item = new DiscountItem(_db[row]);
            item.Id = _db.NextID("№");
            return item;
        }

        public IEnumerator<DiscountItem> GetEnumerator()
        {
            foreach (TableRow item in _db) yield return new DiscountItem(item);
        }

        public int Load()
        {
            _db.Load(_table);
            return _db.RowCount();
        }

        public void Save() => _db.Save();

        IEnumerator IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }

        public DiscountItem GetCurrentDiscount(ClientItem client, DateTime date)
        {
            if (_db.RowCount() == 0) return null;

            var quere = (from d in
                             from item in _db
                             select new DiscountItem(item)
                         where d.ChannelType == client.ChannelType
                                 && d.CustomerStatus == client.CustomerStatus
                                 && d.Period <= date
                         orderby d.Period descending
                         select d).ToList();

            if (quere.Count == 0)
            {
                //MessageBox.Show($"Клиенту {client.Customer} нет соответствий на листе \"Скидки\"", "BPA", MessageBoxButtons.OK, MessageBoxIcon.Information);
                //Debug.Print($"Клиенту {client.Customer} нет соответствий на листе \"Скидки\"", "BPA", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return null;
            }
            DiscountItem currentDiscount = quere[0];

            //проверить формулы
            //Убрать пробелы и лишние знаки
            currentDiscount.NormaliseAllFormulas();

            return currentDiscount;
        }
    }
}
