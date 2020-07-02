using BPA.Model;
using Excel = Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Windows.Forms;
using System.IO;

namespace BPA.Modules
{
    class FileBuget
    {
        private readonly string FileName = "";
        private readonly Microsoft.Office.Interop.Excel.Application Application = Globals.ThisWorkbook.Application;
        private readonly int FileHeaderRow = 1;

        /// <summary>
        /// Событие начала задачи
        /// </summary>
        public event Action<string> ActionStart;
        //public delegate void ActionsStart(string name);

        /// <summary>
        /// Событие завершения задачи
        /// </summary>
        public event Action<int> ActionDone;
        //public delegate void ActionsDone(int count);

        public int CountActions => LastRow - FileHeaderRow;
        private bool IsCancel = false;

        public Excel.Workbook Workbook
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
        private Excel.Workbook _Workbook;

        private Excel.Worksheet Worksheet => Workbook?.Sheets[1];

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

        //public int CustomerColumn => FindColumn("Customer");
        //public int GardenaChannelColumn => FindColumn("GardenaChannel");
        public int DateColumn => FindColumn("Date");
        public int ArticleColumn => FindColumn("Code");
        public int CampaignColumn => FindColumn("CampaignDisc");
        public int QuantitynColumn => FindColumn("Qty");

        #endregion

        public FileBuget()
        {
            using (OpenFileDialog fileDialog = new OpenFileDialog()
            {
                Title = "Выберите расположение файла Descision",
                DefaultExt = "*.xls*",
                CheckFileExists = true,
                //InitialDirectory = Globals.ThisWorkbook.Path,
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

        public FileBuget(string filename)
        {
            if (!File.Exists(filename))
            {
                throw new FileNotFoundException($"Файл {filename} не найден");
            }
            FileName = filename;
        }

        public FileBuget(Excel.Workbook workbook)
        {
            Workbook = workbook;
        }


        public List<ArticleQuantity> ArticleQuantities = new List<ArticleQuantity>();

        public bool IsNotOpen() => FileName == "";
        
        //получение списка артикулов и месяцов
        public void LoadForPlanning(PlanningNewYear planning)
        {
            if (DateColumn == 0 || ArticleColumn == 0 || CampaignColumn == 0)
                throw new ApplicationException("Файл имеет не верный формат");

            for (int rowIndex = FileHeaderRow + 1; rowIndex <= LastRow; rowIndex++)
            {

                if (IsCancel)
                    return;
                ActionStart?.Invoke($"Обрабатывается строка {rowIndex}");

                var campaign = GetValueFromColumn(rowIndex, CampaignColumn);
                if (campaign != "" && (int.TryParse(campaign, out int res) && res == 0))
                    continue;

                DateTime date = GetDateFromCell(rowIndex, DateColumn);
                if (planning.Year != date.Year)
                    continue;

                ////уточнить >=
                //DateTime now = DateTime.Now;
                //if (now.Month > date.Month)
                //    continue;

                string article = GetValueFromColumn(rowIndex, ArticleColumn);
                double quantity;
                if (article != "")
                {
                    quantity = double.TryParse(GetValueFromColumn(rowIndex, QuantitynColumn), out quantity) ? quantity : 0;

                    ArticleQuantities.Add(new ArticleQuantity
                    {
                        Article = article,
                        Quantity = quantity,
                        Month = date.Month
                    });
                }

                ActionDone?.Invoke(1);
            }
            Close();

        }

        /////////////////////////////////
        /// <summary>
        /// получение номена строки по имени заголовка
        /// </summary>
        /// <param name="fildName"></param>
        /// <returns></returns>
        private int FindColumn(string fildName)
        {
            return Worksheet.Cells.Find(fildName, LookAt: Excel.XlLookAt.xlWhole)?.Column ?? 0;
        }

        private int FindRow(string articul)
        {
            return Worksheet.Cells.Find(articul, LookAt: Excel.XlLookAt.xlWhole)?.Row ?? 0;
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

        private DateTime GetDateFromCell(int rw, int col)
        {
            if (Double.TryParse(GetValueFromColumn(rw, col), out double dateDouble))
                return DateTime.FromOADate(dateDouble);
            else if (DateTime.TryParse(GetValueFromColumn(rw, col), out DateTime dateTmp))
                return dateTmp;
            else
                return new DateTime();
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
