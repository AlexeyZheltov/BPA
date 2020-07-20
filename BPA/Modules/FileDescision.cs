using BPA.Model;
using Excel = Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Windows.Forms;
using SettingsBPA = BPA.Properties.Settings;


namespace BPA.Modules
{
    class FileDescision
    {
        private readonly string FileName = "";
        private readonly string FileSheetName = SettingsBPA.Default.SHEET_NAME_FILE_DECISION;
        private readonly Microsoft.Office.Interop.Excel.Application Application = Globals.ThisWorkbook.Application;
        //private readonly string ToBeSoldInNeed = "RUSSIA";
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

        //private Excel.Worksheet worksheet => Workbook?.Sheets[FileSheetName];
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
                        throw new ApplicationException($"Лист \"{ FileSheetName }\" в книге { FileName } не найден!");
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
                if (_LastRow == 0) _LastRow = worksheet.UsedRange.Row + worksheet.UsedRange.Rows.Count - 1;
                return _LastRow;
            }
        }
        private int _LastRow = 0;

        #region --- Columns ---

        public int CustomerColumn => FindColumn("Customer");
        public int GardenaChannelColumn => FindColumn("GardenaChannel");
        public int DateColumn => FindColumn("Date");
        public int ArticleColumn => FindColumn("Code");
        public int CampaignColumn => FindColumn("Campaign");
        public int QuantitynColumn => FindColumn("Quantity");
        public int PriceListTotalColumn => FindColumn("PricelistPriceTotal");
        public int BonusColumn => FindColumn("Bonus");

        #endregion

        public FileDescision()
        {
            BPASettings settings = new BPASettings();

            if (settings.GetDecisionPath(out string path))
            {
                FileName = path;
                IsOpen = true;
            }
            else
            {
                throw new ApplicationException("Загрузка отменена");
            }
        }

        public FileDescision(string filename)
        {
            if (!File.Exists(filename))
            {
                throw new FileNotFoundException($"Файл {filename} не найден");
            }
            FileName = filename;
        }

        public FileDescision(Excel.Workbook workbook)
        {
            Workbook = workbook;
        }


        public List<ArticleQuantity> ArticleQuantities = new List<ArticleQuantity>();
        
        //public bool IsNotOpen() => FileName == "";
        
        public List<Client> LoadClients()
        {
            List<Client> buffer = new List<Client>();

            if (CustomerColumn == 0 || GardenaChannelColumn == 0)
            {
                Close();
                throw new ApplicationException("Файл имеет неверный формат");
            }

            for(int rowIndex = FileHeaderRow + 1; rowIndex <= LastRow; rowIndex++)
            {
                if (IsCancel) return null;
                ActionStart?.Invoke($"Обрабатывается строка {rowIndex}");
                Excel.Range range = worksheet.Cells[rowIndex, CustomerColumn];
                string customer = range.Text;
                if(customer.Trim().Length > 0)
                {
                    range = worksheet.Cells[rowIndex, GardenaChannelColumn];
                    string gardenaChannel = range.Text;

                    if(!buffer.Exists(x => x.Customer == customer)) buffer.Add(new Client()
                    {
                        Customer = customer,
                        GardenaChannel = gardenaChannel
                    });
                }

                ActionDone?.Invoke(1);
            }

            if (buffer.Count == 0) throw new ApplicationException("Файл не содержит значемых данных");
            return buffer;
        }


        //gjkextybt 
        //получение списка артикулов и месяцов
        public void LoadForPlanning(PlanningNewYear planning)
        {
            if (DateColumn == 0 || ArticleColumn == 0 || CampaignColumn ==0)
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
                double bonus;

                if (article != "")
                {
                    quantity = double.TryParse(GetValueFromColumn(rowIndex, QuantitynColumn), out quantity) ? quantity : 0;
                    priceList = double.TryParse(GetValueFromColumn(rowIndex, PriceListTotalColumn), out priceList) ? priceList : 0;
                    bonus = double.TryParse(GetValueFromColumn(rowIndex, BonusColumn), out bonus) ? bonus : 0;

                    ArticleQuantities.Add(new ArticleQuantity
                    {
                        Article = article,
                        Quantity = quantity,
                        Month = date.Month,
                        Campaign = campaign == "" ? "0" : campaign,
                        PriceList = priceList,
                        Bonus = bonus
                    });
                }

                ActionDone?.Invoke(1);
            }
        }

        //public PlanningNewYear LoadPrognosis(PlanningNewYearPrognosis planningNewYearPrognosis)
        //{
        //    if (DateColumn == 0 || ArticleColumn == 0)
            //{
            //    Close();
            //    throw new ApplicationException("Файл имеет неверный формат");
            //}
    //    //временный лист
    //    //List<PlanningNewYear> buffer = new List<PlanningNewYear>();

    //    for (int rowIndex = FileHeaderRow + 1; rowIndex <= LastRow; rowIndex++)
    //    {
    //        if (IsCancel)
    //            return null;
    //        ActionStart?.Invoke($"Обрабатывается строка {rowIndex}");

    //        if (planningNewYear.Article != GetValueFromColumn(rowIndex, ArticleColumn))
    //            continue;

    //        var campaign = GetValueFromColumn(rowIndex, CampaignColumn);
    //        if (campaign != "" && (int.TryParse(campaign, out int res) && res == 0))                
    //            continue;

    //        DateTime date = GetDateFromCell(rowIndex, DateColumn);
    //        if (planningNewYear.Year != date.Year)
    //            continue;

    //        //Здесь добавляем помесячно
    //        //date.Month;

    //        ActionDone?.Invoke(1);
    //    }

    //    return planningNewYear;

    //    //if (buffer.Count == 0)
    //    //    throw new ApplicationException("Файл не содержит значемых данных");
    //    //return buffer;
    //}

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
