using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace BPA.NewModel
{
    class DiscountForPlanningTable : IEnumerable<DiscountItem>
    {
        const string SHEET = "Скидки для планирования";
        const string TABLE = "СкидкиДляПланирования";

        WS_DB _db = new WS_DB();
        Excel.ListObject _table = null;

        public DiscountForPlanningTable()
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

        public DiscountForPlanningItem GetDiscountForPlanning(string channelType, string customerStatusForecast, DateTime planningDate)
        {
            if (_db.RowCount() == 0) return null;

            var quere = (from d in
                            (from item in _db
                             select new DiscountForPlanningItem(item))
                         where d.ChannelType == channelType
                                && d.CustomerStatusForecast == customerStatusForecast
                                && d.Period <= planningDate
                         orderby d.Period descending
                         select d).ToList();
            
            if (quere.Count == 0)
            {
                //MessageBox.Show($"Клиенту {client.Customer} нет соответствий на листе \"Скидки\"", "BPA", MessageBoxButtons.OK, MessageBoxIcon.Information);
                //Debug.Print($"Клиенту {client.Customer} нет соответствий на листе \"Скидки\"", "BPA", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return null;
            }

            DiscountForPlanningItem currentDiscount = quere[0];

            //проверить формулы
            //Убрать пробелы и лишние знаки
            currentDiscount.NormaliseAllFormulas();

            return currentDiscount;
        }
    }
}
