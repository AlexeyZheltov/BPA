using BPA.Modules;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace BPA.NewModel
{
    class ProductTable : IEnumerable<ProductItem>
    {
        const string SHEET = "Товары";
        const string TABLE = "Товары";

        WS_DB _db = new WS_DB();
        Excel.ListObject _table = null;

        public ProductTable()
        {
            Excel.Workbook wb = Globals.ThisWorkbook.InnerObject;
            Excel.Worksheet ws = wb.Sheets[SHEET];
            _table = ws.ListObjects[TABLE];
        }

        public IEnumerator<ProductItem> GetEnumerator()
        {
            foreach (TableRow item in _db) yield return new ProductItem(item);
        }

        public int Load()
        {
            _db.Load(_table);
            return _db.RowCount();
        }

        public void Save() => _db.Save();

        public int Count => _db.RowCount();

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

        public double BudgetCourse()
        {
            string Label = "Бюджетный курс";

            try
            {
                Excel.Worksheet ws = _table.Parent;
                int i_row = _table.HeaderRowRange.Row - 1;
                Excel.Range rng = ws.Rows[i_row];
                rng = rng.Find(Label, LookAt: Excel.XlLookAt.xlWhole);
                rng = rng.Offset[0, 1];
                return double.Parse(rng.Text);
            }
            catch
            {
                return default;
            }
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }

        public ProductItem Find(Predicate<ProductItem> predicate)
        {
            foreach (ProductItem product in this)
                if (predicate(product)) return product;
            return null;
        }

        public ProductItem Add() 
        {
            int row = _db.AddRow();
            ProductItem item = new ProductItem(_db[row]);
            item.Id = _db.NextID("№");
            return item;
        }
        
        /// <summary>
        /// Сортировка умной таблицы по имени столбца
        /// </summary>
        /// <param name="col_name"></param>
        public void Sort(string col_name)
        {
            _table.Sort.SortFields.Clear();
            _table.Sort.SortFields.Add(Key: _table.ListColumns[col_name].Range,
                                        Excel.XlSortOn.xlSortOnValues,
                                        Excel.XlSortOrder.xlAscending,
                                        Excel.XlSortDataOption.xlSortNormal);
            _table.Sort.Header = Excel.XlYesNoGuess.xlYes;
            _table.Sort.MatchCase = false;
            _table.Sort.Orientation = Excel.XlSortOrientation.xlSortColumns;
            _table.Sort.SortMethod = Excel.XlSortMethod.xlPinYin;
            _table.Sort.Apply();
        }

        public int GetId(int excelrow)
        {
            Excel.Range rng = _table.DataBodyRange;
            if (excelrow < rng[1].Row || excelrow > rng[rng.Cells.Count].Row)
                return 0;

            Excel.ListRow listRow = _table.ListRows[excelrow - rng.Row + 1];

            object val = listRow.Range[1, _table.ListColumns["№"].Index].Value;
            
            if (val.ToString().All(char.IsNumber))
                return Convert.ToInt32(val);
            return 0;
        }

        public List<ProductItem> GetProductForClient(ClientItem client, List<string> exclusives)
        {
            if (_db.RowCount() == 0) return null;


            List<ProductItem> prod_list = (from pl in
                                               from item in _db
                                               select new ProductItem(item)
                                           where pl.Status.ToLower() != "выведено из ассортимента текущего года"
                                                 && pl.Status.ToLower() != "выведено из глобального ассортимента"
                                           select pl).ToList();

            List<ProductItem> actualProducts = new List<ProductItem>();
            foreach (ProductItem product in prod_list)
            {
                if (exclusives.Contains(product.Exclusive.ToLower()))
                {
                    if (product.Exclusive.ToLower() == client.CustomerStatus.ToLower())
                        actualProducts.Add(product);
                }
                else
                {
                    switch (product.Exclusive.ToLower())
                    {
                        case "diy канал":
                            if (client.ChannelType.ToLower() == "diy")
                                actualProducts.Add(product);
                            break;
                        case "online":
                            if (client.ChannelType.ToLower() == "online")
                                actualProducts.Add(product);
                            break;
                        case "dealer":
                        case "regional": //DEALERS&REGIONAL DISTR
                            if (client.ChannelType.ToLower() == "dealer&regional distr")
                                actualProducts.Add(product);
                            break;
                        default:
                            actualProducts.Add(product);
                            break;
                    }
                }
            }

            if (actualProducts.Count == 0) MessageBox.Show("Данному клиенту не соотвествует ни один акртикул", "BPA", MessageBoxButtons.OK, MessageBoxIcon.Information);
            return actualProducts;
        }
    }
}
