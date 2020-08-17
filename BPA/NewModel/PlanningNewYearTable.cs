using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using SettingsBPA = BPA.Properties.Settings;

namespace BPA.NewModel
{
    class PlanningNewYearTable : IEnumerable<PlanningNewYearItem>
    {
        private string SHEET  => _TableWorksheetName != null ? _TableWorksheetName: templateSheetName;
        private string _TableWorksheetName;
        public readonly string templateSheetName = SettingsBPA.Default.SHEET_NAME_PLANNING_TEMPLATE;
        private string TABLE => GetTableName();
        public string GetTableName()
        {
            ThisWorkbook workbook = Globals.ThisWorkbook;
            Excel.ListObject table = workbook.Sheets[SHEET].ListObjects[1];
            return table.Name;
        }

        WS_DB _db = new WS_DB();
        Excel.ListObject _table = null;

        public PlanningNewYearTable(string worksheetName)
        {
            if (worksheetName == templateSheetName)
                return;
            else
                _TableWorksheetName = worksheetName;

            Excel.Workbook wb = Globals.ThisWorkbook.InnerObject;
            Excel.Worksheet ws = wb.Sheets[SHEET];
            _table = ws.ListObjects[TABLE];
        }

        public IEnumerator<PlanningNewYearItem> GetEnumerator()
        {
            foreach (TableRow item in _db) yield return new PlanningNewYearItem(item);
        }
        IEnumerator IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }

        public int Load()
        {
            _db.Load(_table);
            return _db.RowCount();
        }

        public void Save() => _db.Save();

        public int Count => _db.RowCount();

        public PlanningNewYearItem Find(Predicate<PlanningNewYearItem> predicate)
        {
            foreach (PlanningNewYearItem planning in this)
                if (predicate(planning)) return planning;
            return null;
        }

        public PlanningNewYearItem Add()
        {
            int row = _db.AddRow();
            PlanningNewYearItem item = new PlanningNewYearItem(_db[row]);
            item.Id = _db.NextID("№");
            return item;
        }

        //public int GetId(int excelrow)
        //{
        //    Excel.Range rng = _table.DataBodyRange;
        //    if (excelrow < rng[1].Row || excelrow > rng[rng.Cells.Count].Row)
        //        return 0;

        //    Excel.ListRow listRow = _table.ListRows[excelrow - rng.Row + 1];

        //    object val = listRow.Range[1, _table.ListColumns["№"].Index].Value;

        //    if (val.ToString().All(char.IsNumber))
        //        return Convert.ToInt32(val);
        //    return 0;
        //}

        //public List<ProductItem> GetProductForClient(ClientItem client, List<string> exclusives)
        //{
        //    if (_db.RowCount() == 0) return null;


        //    List<ProductItem> prod_list = (from pl in
        //                                       from item in _db
        //                                       select new ProductItem(item)
        //                                   where pl.Status.ToLower() != "выведено из ассортимента текущего года"
        //                                         && pl.Status.ToLower() != "выведено из глобального ассортимента"
        //                                   select pl).ToList();

        //    List<ProductItem> actualProducts = new List<ProductItem>();
        //    foreach (ProductItem product in prod_list)
        //    {
        //        if (exclusives.Contains(product.Exclusive.ToLower()))
        //        {
        //            if (product.Exclusive.ToLower() == client.CustomerStatus.ToLower())
        //                actualProducts.Add(product);
        //        }
        //        else
        //        {
        //            switch (product.Exclusive.ToLower())
        //            {
        //                case "diy канал":
        //                    if (client.ChannelType.ToLower() == "diy")
        //                        actualProducts.Add(product);
        //                    break;
        //                case "online":
        //                    if (client.ChannelType.ToLower() == "online")
        //                        actualProducts.Add(product);
        //                    break;
        //                case "dealer":
        //                case "regional": //DEALERS&REGIONAL DISTR
        //                    if (client.ChannelType.ToLower() == "dealer&regional distr")
        //                        actualProducts.Add(product);
        //                    break;
        //                default:
        //                    actualProducts.Add(product);
        //                    break;
        //            }
        //        }
        //    }

        //    if (actualProducts.Count == 0) MessageBox.Show("Данному клиенту не соотвествует ни один акртикул", "BPA", MessageBoxButtons.OK, MessageBoxIcon.Information);
        //    return actualProducts;
        //}
    }
}
