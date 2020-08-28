using BPA.Modules;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using NM = BPA.NewModel;
using SettingsBPA = BPA.Properties.Settings;

namespace BPA.NewModel
{
    class FinalPriceTable : IEnumerable<FinalPriceItem>
    {
        public string SheetName => SHEET;
        private string SHEET => _TableWorksheetName != null ? _TableWorksheetName : templateSheetName;
        private string _TableWorksheetName;
        public readonly string templateSheetName = SettingsBPA.Default.SHEET_NAME_PRICELIST_TEMPLATE;
        private Excel.Workbook workbook;

        private string TABLE => GetTableName();
        private string GetTableName()
        {
            Excel.ListObject table = workbook.Sheets[SHEET].ListObjects[1];
            return table.Name;
        }

        WS_DB _db = new WS_DB();
        Excel.ListObject _table = null;

        public FinalPriceTable(Excel.Worksheet worksheet)
        {
            if (worksheet.Name == templateSheetName)
                return;
            else
                _TableWorksheetName = worksheet.Name;

            workbook = worksheet.Parent;
            Excel.Worksheet ws = workbook.Sheets[SHEET];
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
            int row = _db.AddRow();
            FinalPriceItem item = new FinalPriceItem(_db[row]);
            //item.Id = db.NextID("Id");
            return item;
        }

        public IEnumerator<FinalPriceItem> GetEnumerator()
        {
            foreach (TableRow item in _db) yield return new FinalPriceItem(item);
        }

        public int Load()
        {
            _db.Load(_table);
            return _db.RowCount();
        }

        public void Save() {
            _db.Save();
        }

        public void DelFirstRow()
        {
            _db.Delete(0);
        }
        IEnumerator IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }

        private const string labelDate = "Дата обновления:";
        private const string labelCustomer = "Customer status";
        private const string labelChannel = "Channel type";
        private DateTime DateOfPrice;
        private string CustomerStatus;
        private string ChannelType;

        public void SetParams(string customerStatus , string channelType, DateTime date)
        {
            DateOfPrice = date;
            CustomerStatus = customerStatus;
            ChannelType = channelType;

            SetCategoryParams();
        }

        private void SetCategoryParams()
        {
            try
            {
                Excel.Worksheet ws = _table.Parent;

                setVal(labelDate, DateOfPrice);
                setVal(labelCustomer, CustomerStatus);
                setVal(labelChannel, ChannelType);

                void setVal(string label, object val)
                {
                    Excel.Range LabelCell = ws.UsedRange.Find(label, LookAt: Excel.XlLookAt.xlWhole);

                    if (LabelCell == null) return;

                    Excel.Range cell = LabelCell.Offset[0, 1];
                    cell.Value = val;
                }
            } catch
            {

            }
        }
    }
}
