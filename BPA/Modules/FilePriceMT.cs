using BPA.Model;

using Microsoft.Office.Interop.Excel;

using System;
using System.IO;
using System.Windows.Forms;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BPA.Modules
{
    class Test
    {
        void Go()
        {
            FilePriceMT file = new FilePriceMT();
            file.Load("Название нужного магазина", new DateTime()); //Загрузить в внутринности класса список артикулов соответсвтующий даннаму магазину, за данную дату

            //цикл
            file.GetPrice("fhn");


        }
    }

    internal class FilePriceMT
    {

        private readonly string FileName;
        private readonly Microsoft.Office.Interop.Excel.Application Application = Globals.ThisWorkbook.Application;
        private readonly int CalendarHeaderRow = 1;

        /// <summary>
        /// Событие начала задачи
        /// </summary>
        public event ActionsStart ActionStart;
        public delegate void ActionsStart(string name);

        /// <summary>
        /// Событие завершения задачи
        /// </summary>
        public event ActionsDone ActionDone;
        public delegate void ActionsDone(int count);

        public int CountActions => LastRow - CalendarHeaderRow;
        private bool IsCancel = false;

        public Workbook Workbook
        {
            get
            {
                if (_Workbook == null)
                {
                    try
                    {
                        _Workbook = Application.Workbooks.Open(FileName);
                    }
                    catch
                    {
                        _Workbook = null;
                    }
                }
                return _Workbook;
            }
            set
            {
                _Workbook = value;
            }
        }
        private Workbook _Workbook;

        private Worksheet Worksheet => Workbook?.Sheets[1];

        public int LastRow
        {
            get
            {
                if (_LastRow == 0)
                    _LastRow = Worksheet.UsedRange.Row + Worksheet.UsedRange.Rows.Count - 1;
                return _LastRow;
            }
        }
        private int _LastRow = 0;

        #region --- Columns ---

        public int MainColumn => FindColumn("Главный");
        public int ArticleColumn => FindColumn("Артикул");
        public int NameColumn => FindColumn("Название");
        public int PriceForClientColumn => FindColumn("Цена_для_клиента");
        public int ValidFromDatColumn => FindColumn("ValidFromDat");
        public int ValidToDatColumn => FindColumn("ValidToDat");
        public int CustCodeColumn => FindColumn("CustCode");
        public int PriceOfListingColumn => FindColumn("Цена_листинга");
        public int PriceNewColumn => FindColumn("Цена_новая");
        public int FromColumn => FindColumn("От");
        public int ToColumn => FindColumn("До");
        public int MagColumn => FindColumn("Маг");


        #endregion


        public FilePriceMT()
        {
            using (OpenFileDialog fileDialog = new OpenFileDialog()
            {
                Title = "Выберите расположение продуктового календаря",
                DefaultExt = "*.xls*",
                CheckFileExists = true,
                InitialDirectory = Globals.ThisWorkbook.Path,
                ValidateNames = true,
                Multiselect = false,
                Filter = "Excel|*.xls*"
            })
            {
                if (fileDialog.ShowDialog() == DialogResult.OK)
                {
                    FileName = fileDialog.FileName;
                }
            }

        }

        public FilePriceMT(string filename)
        {
            if (!File.Exists(filename))
            {
                throw new FileNotFoundException($"Файл {filename} не найден");
            }
            FileName = filename;
        }

        public FilePriceMT(Workbook workbook)
        {
            Workbook = workbook;
        }

        private List<Client> _clients = new List<Client>();
        public struct Client
        {
            public string Name {
                get; set;
            }
            public double Price
            {
                get; set;
            }
            public string Art
            {
                get; set;
            }
        }

        public void Load(string mag, DateTime date)
        {
            if (Workbook == null)
                return;

            _clients.Clear();
            //Тут загрузить все цены по магазину и дате.
            for (int rw = CalendarHeaderRow + 1; rw <= LastRow; rw++)
            {
                if (!double.TryParse(GetValueFromColumn(rw, PriceForClientColumn), out double price))
                {
                    price = 0;
                }
                
                clients.Add(new Client
                {
                    Name = GetValueFromColumn(rw, MainColumn),
                    Price =  price,
                    Art = GetValueFromColumn(rw, ArticleColumn)
                });
            }

            Close();
            IsCancel = true;
        }

        public double GetPrice(string Art)
        {
            return clients.Find(x => x.Art == Art).Price;
        }


        /// <summary>
        /// получение номена строки по имени заголовка
        /// </summary>
        /// <param name="fildName"></param>
        /// <returns></returns>
        private int FindColumn(string fildName)
        {
            return Worksheet.Cells.Find(fildName, LookAt: XlLookAt.xlWhole)?.Column ?? 0;
        }


        private int FindRow(int column, string articul)
        {
            return Worksheet.Columns[column].Find(articul, LookAt: XlLookAt.xlWhole)?.Row ?? 0;
        }

        private int FindRow(int column, string articul, Range afterCell)
        {
            return Worksheet.Columns[column].Find(articul, After:afterCell, LookAt: XlLookAt.xlWhole)?.Row ?? 0;
        }

        /// <summary>
        /// получение значения из строки по номеру столбца
        /// </summary>
        /// <param name="rw"></param>
        /// <param name="col"></param>
        /// <returns></returns>
        private string GetValueFromColumn(int rw, int col)
        {
            return col != 0 ? Worksheet.Cells[rw, col].value?.ToString() : "";
        }

        public void Close()
        {
            Workbook.Close(false);
        }

        public void Cancel()
        {
            IsCancel = true;
        }

    }
}
