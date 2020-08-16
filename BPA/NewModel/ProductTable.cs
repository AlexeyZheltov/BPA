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

        public ProductItem Find(Predicate<ProductItem> predicate)
        {
            foreach (ProductItem product in this)
                if (predicate(product)) return product;
            return null;
        }

        public ProductItem Add() 
        {
            int row = db.AddRow();
            ProductItem item = new ProductItem(db[row]);
            item.Id = db.NextID("№");
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
                                        XlSortOn.xlSortOnValues,
                                        XlSortOrder.xlAscending,
                                        XlSortDataOption.xlSortNormal);
            _table.Sort.Header = XlYesNoGuess.xlYes;
            _table.Sort.MatchCase = false;
            _table.Sort.Orientation = XlSortOrientation.xlSortColumns;
            _table.Sort.SortMethod = XlSortMethod.xlPinYin;
            _table.Sort.Apply();
        }

        public int GetId(int excelrow)
        {
            Excel.Range rng = _table.DataBodyRange;
            if (excelrow < rng[1].Row || excelrow > rng[rng.Cells.Count].Row)
                return 0;

            ListRow listRow = _table.ListRows[excelrow - rng.Row + 1];

            object val = listRow.Range[1, _table.ListColumns["№"].Index].Value;
            
            if (val.ToString().All(char.IsNumber))
                return Convert.ToInt32(val);
            return 0;
        }
    }
}
