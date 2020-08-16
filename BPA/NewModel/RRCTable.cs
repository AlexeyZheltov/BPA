using Microsoft.VisualStudio.Tools.Applications.Runtime;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using BPA.Forms;
using BPA.Modules;
using System.Security.Cryptography;

namespace BPA.NewModel
{
    class RRCTable : IEnumerable<RRCItem>
    {
        public const string SHEET = "РРЦ";
        const string TABLE = "РРЦ";

        WS_DB _db = new WS_DB();
        Excel.ListObject _table = null;

        public RRCTable()
        {
            Excel.Workbook wb = Globals.ThisWorkbook.InnerObject;
            Excel.Worksheet ws = wb.Sheets[SHEET];
            _table = ws.ListObjects[TABLE];
        }
        #region Стандартный набор
        public RRCItem Find(Predicate<RRCItem> predicate)
        {
            foreach (RRCItem rrc in this)
                if (predicate(rrc)) return rrc;
            
            return null;
        }

        public RRCItem Add()
        {
            int row = _db.AddRow();
            RRCItem item = new RRCItem(_db[row]);
            item.Id = _db.NextID("№");
            return item;
        }

        public int Load()
        {
            _db.Load(_table);
            return _db.RowCount();
        }

        public void Save() => _db.Save();

        public IEnumerator<RRCItem> GetEnumerator()
        {
            foreach (TableRow item in _db) yield return new RRCItem(item);
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }
        #endregion
        public List<RRCItem> GetActualPriceList(DateTime currentDate)
        {
            PBWrapper pb = new PBWrapper($"Создание прайс-листа", $"Анализ артикулов с листа РРЦ [Index]");

            //подключится к ценам
            if ((_db?.Count() ?? 0) == 0) return null;
            //список уникальных артикулов
            List<RRCItem> rrc_list = (from item in _db
                                           select new RRCItem(item))
                                           .Distinct(new RRCItem.RRCItemComparerForPrice())
                                           .ToList();

            List<RRCItem> actualRRC = new List<RRCItem>();
            pb.Start(rrc_list.Count);
            //взять пачку строк соответсвующих артикулу и вязть тот что с последней датой
            foreach (RRCItem item in rrc_list)
            {
                if (pb.IsCancel)
                {
                    pb.Dispose();
                    return null;
                }
                pb.Action(item.Article);

                List<RRCItem> buffer = (from rrc in rrc_list
                                        where rrc.Article == item.Article && rrc.Date <= currentDate
                                        orderby rrc.Date descending
                                        select rrc).ToList();


                if (buffer.Count == 0) continue;
                actualRRC.Add(buffer[0]);
                pb.Done(1);
            }
            pb.Dispose();

            return actualRRC;
        }       
    }
}
