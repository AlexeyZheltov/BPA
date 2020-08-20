using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace BPA.NewModel
{
    class ClientTable : IEnumerable<ClientItem>
    {
        public const string SHEET = "Клиенты";
        const string TABLE = "Клиенты";

        WS_DB _db = new WS_DB();
        Excel.ListObject _table = null;

        public static void SortExcelTable(string sortFildName)
        {
            Excel.Workbook wb = Globals.ThisWorkbook.InnerObject;
            Excel.Worksheet ws = wb.Sheets[SHEET];
            Excel.ListObject table = ws.ListObjects[TABLE];

            table.Sort.SortFields.Clear();
            table.Sort.SortFields.Add(Key: table.ListColumns[sortFildName].Range,
                                        Excel.XlSortOn.xlSortOnValues,
                                        Excel.XlSortOrder.xlAscending,
                                        Excel.XlSortDataOption.xlSortNormal);
            table.Sort.Header = Excel.XlYesNoGuess.xlYes;
            table.Sort.MatchCase = false;
            table.Sort.Orientation = Excel.XlSortOrientation.xlSortColumns;
            table.Sort.SortMethod = Excel.XlSortMethod.xlPinYin;
            table.Sort.Apply();
        }

        public ClientTable()
        {
            Excel.Workbook wb = Globals.ThisWorkbook.InnerObject;
            Excel.Worksheet ws = wb.Sheets[SHEET];
            _table = ws.ListObjects[TABLE];
        }

        public bool Contains(ClientItem.DataFromDescision value)
        {
            foreach (var item in this)
                if (item.Customer == value.Customer) return true;

            return false;
        }

        public ClientItem this[int row] => new ClientItem(_db[row]);

        public ClientItem GetById(int id)
        {
            var quere = (from item in
                             from db_item in _db
                             select new ClientItem(db_item)
                         where item.Id == id
                         select item).ToList();
            if (quere.Count > 0) return quere[0];
            else return null;
        }

        public ClientItem Add()
        {
            int row = _db.AddRow();
            ClientItem item = new ClientItem(_db[row]);
            item.Id = _db.NextID("№");
            return item;
        }

        public int Count() => _db.RowCount();

        public IEnumerator<ClientItem> GetEnumerator()
        {
            foreach (TableRow item in _db) yield return new ClientItem(item);
        }

        public int Load()
        {
            _db.Load(_table);
            return _db.RowCount();
        }

        public int GetCurrentClientID()
        {
            Excel.Workbook wb = Globals.ThisWorkbook.InnerObject;
            Excel.Worksheet ws = wb.Sheets[SHEET];
            Excel.ListObject table = ws.ListObjects[TABLE];

            Excel.Range rng = Globals.ThisWorkbook.Application.ActiveCell;
            if (rng.Worksheet.Name == ws.Name)
            {
                int row = rng.Row;
                int column = table.ListColumns["№"].DataBodyRange.Column;

                if (int.TryParse(ws.Cells[row, column].Value, out int Id))
                    return Id;
            }

            return 0;
        }

        public void Save() => _db.Save();

        IEnumerator IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }
    }
}
