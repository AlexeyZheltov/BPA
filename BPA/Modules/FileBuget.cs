using BPA.Model;
using Excel = Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Windows.Forms;
using System.IO;
using SettingsBPA = BPA.Properties.Settings;

namespace BPA.Modules
{
    class FileBuget
    {
        private readonly string FileName = "";
        private readonly string FileSheetName = SettingsBPA.Default.SHEET_NAME_FILE_BUGET;
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

        public bool IsOpen { get; set; } = false;
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

        //        private Excel.Worksheet worksheet => Workbook?.Sheets[FileSheetName];
        public Excel.Worksheet worksheet
        {
            get
            {
                if (_worksheet == null)
                {
                    try
                    {
                        _worksheet = Workbook?.Sheets[FileSheetName];
                    }
                    catch
                    {
                        throw new ApplicationException($"Лист { FileSheetName } в книге { FileName } не найден!");
                    }
                }
                return _worksheet;
            }
            set
            {
                _worksheet = value;
            }
        }
        private Excel.Worksheet _worksheet;

        public int LastRow
        {
            get
            {
                if (_LastRow == 0)
                    _LastRow = worksheet.UsedRange.Row + worksheet.UsedRange.Rows.Count - 1;
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
        public int PriceListColumn => FindColumn("PriceList");

        #endregion

        public FileBuget()
        {
            BPASettings settings = new BPASettings();

            if (settings.GetBudgetPath(out string path))
            {
                FileName = path;
                IsOpen = true;
            }
            else
            {
                throw new ApplicationException("Загрузка отменена");
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

        //public bool IsNotOpen() => FileName == "";
        
        //получение списка артикулов и месяцов
        public void LoadForPlanning(PlanningNewYear planning)
        {
            if (DateColumn == 0 || ArticleColumn == 0 || CampaignColumn == 0)
            {
                throw new ApplicationException("Файл имеет неверный формат");
            }

            for (int rowIndex = FileHeaderRow + 1; rowIndex <= LastRow; rowIndex++)
            {
                if (IsCancel)
                    return;
                ActionStart?.Invoke($"Обрабатывается строка {rowIndex}");

                DateTime date = GetDateFromCell(rowIndex, DateColumn);
                if (planning.Year != date.Year)
                    continue;

                string article = GetValueFromColumn(rowIndex, ArticleColumn);
                string campaign = GetValueFromColumn(rowIndex, CampaignColumn);
                double quantity;
                double priceList;
                if (article != "")
                {
                    quantity = double.TryParse(GetValueFromColumn(rowIndex, QuantitynColumn), out quantity) ? quantity : 0;
                    priceList = double.TryParse(GetValueFromColumn(rowIndex, PriceListColumn), out priceList) ? priceList : 0;

                    ArticleQuantities.Add(new ArticleQuantity
                    {
                        Article = article,
                        Quantity = quantity,
                        Month = date.Month,
                        Campaign = campaign == "" ? "0": campaign,
                        PriceList = priceList
                    });
                }

                ActionDone?.Invoke(1);
            }
        }

        /////////////////////////////////
        /// <summary>
        /// получение номена строки по имени заголовка
        /// </summary>
        /// <param name="fildName"></param>
        /// <returns></returns>
        private int FindColumn(string fildName)
        {
            return worksheet.Cells.Find(fildName, LookAt: Excel.XlLookAt.xlWhole)?.Column ?? 0;
        }

        private int FindRow(string articul)
        {
            return worksheet.Cells.Find(articul, LookAt: Excel.XlLookAt.xlWhole)?.Row ?? 0;
        }

        /// <summary>
        /// получение значения из строки по номеру столбца
        /// </summary>
        /// <param name="rw"></param>
        /// <param name="col"></param>
        /// <returns></returns>
        private string GetValueFromColumn(int rw, int col)
        {
            return col != 0 ? worksheet.Cells[rw, col].value?.ToString() : "";
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
            IsOpen = false;
            Workbook.Close(false);
        }

        public void Cancel()
        {
            IsCancel = true;
        }

    }
}
