using BPA.Modules;
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

        public List<STKItem> GetActualPriceList(DateTime data)
        {
            PBWrapper pb = new PBWrapper($"Создание прайс-листа", $"Анализ артикулов с листа STK [Index]");

            //подключится к ценам
            if ((_db?.Count() ?? 0) == 0) return null;
            //список уникальных артикулов
            List<STKItem> list = (from item in _db
                                      select new STKItem(item))
                                           .Distinct(new STKItem.STKItemComparerForPrice())
                                           .ToList();

            List<STKItem> all = (from row in _db
                                     select new STKItem(row)).ToList();
            List<STKItem> actualSTK = new List<STKItem>();
            pb.Start(list.Count);
            //взять пачку строк соответсвующих артикулу и вязть тот что с последней датой
            foreach (STKItem item in list)
            {
                if (pb.IsCancel)
                {
                    pb.Dispose();
                    return null;
                }
                pb.Action(item.Article);

                List<STKItem> buffer = (from item_1 in all
                                        where item_1.Article == item.Article && item_1.Date <= data
                                        orderby item_1.Date descending
                                        select item_1).ToList();


                if (buffer.Count == 0) continue;
                actualSTK.Add(buffer[0]);
                pb.Done(1);
            }
            pb.Dispose();

            return actualSTK;
        }
    }
}
