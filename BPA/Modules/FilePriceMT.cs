﻿using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;
using System;
using System.IO;
using System.Windows.Forms;
using System.Collections.Generic;
using SettingsBPA = BPA.Properties.Settings;



namespace BPA.Modules
{
    internal class FilePriceMT
    {

        private readonly string FileName = "";
        private readonly string FileSheetName = SettingsBPA.Default.SHEET_NAME_FILE_PRICELISTMT;
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

        public bool IsOpen { get; set; } = false;

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
                if (_LastRow == 0)
                    _LastRow = worksheet.UsedRange.Row + worksheet.UsedRange.Rows.Count - 1;
                return _LastRow;
            }
        }
        private int _LastRow = 0;

        #region --- Columns ---
        public int CustomerColumn => FindColumn("Покупатель");
        public int SearchColumn => FindColumn("Поиск");
        public int MainColumn => FindColumn("Главный");
        public int ArticleColumn => FindColumn("Артикул");
        public int NameColumn => FindColumn("Название");
        public int PriceForClientColumn => FindColumn("Цена_для_клиента");
        public int ValidFromDatColumn => FindColumn("ValidFromDat");
        public int ValidToDatColumn => FindColumn("ValidToDat");
        public int CustCodeColumn => FindColumn("CustCode");
        public int PriceOfListingColumn => FindColumn("Цена_листинга");
        public int PriceNewColumn => FindColumn("Цена_новая");
        public int DateFromColumn => FindColumn("От");
        public int DateToColumn => FindColumn("До");
        public int MagColumn => FindColumn("Маг");


        #endregion


        public FilePriceMT()
        {
            BPASettings settings = new BPASettings();

            if (settings.GetPriceListMT(out string path))
            {
                FileName = path;
                IsOpen = true;
            }
            else
            {
                throw new ApplicationException("Загрузка отменена");
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

        public List<Client> clients = new List<Client>();
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
        
        /// <summary>
        /// получение магазина по дате
        /// </summary>
        /// <param name="mag"></param>
        /// <param name="date"></param>
        public void Load(string mag, DateTime date)
        {
            if (Workbook == null)
                return;

            clients.Clear();

            int magColumn = MagColumn;
            if (magColumn == 0)
            {
                Workbook.Close();
                throw new ApplicationException($"Файл {Path.GetFileName(FileName)} имеет ошибочный формат");
            }
            int rw = FindRow(MagColumn, mag);
            if (rw == 0)
                return;

            IsCancel = false;
            ActionStart?.Invoke("Загрузка файла PriceListMT");
            int firstFindedRw = rw;

            do
            {
                if (IsCancel) return;

                DateTime firstDate = GetDateFromCell(rw, DateFromColumn);
                DateTime lastDate = GetDateFromCell(rw, DateToColumn);

                if (lastDate.Year >= 9999)
                {
                    AddClient(rw, PriceOfListingColumn);
                }
                else if (date <= lastDate && date >= firstDate)
                {
                    AddClient(rw, PriceNewColumn);
                }
                rw = FindRow(MagColumn, mag, worksheet.Cells[rw, MagColumn]);
                ActionDone?.Invoke(1);
            } 
            while (firstFindedRw != rw);
            IsOpen = true;
        }

        /// <summary>
        /// получение магазина по дате
        /// </summary>
        /// <param name="date"></param>
        public void Load(DateTime date)
        {
            if (Workbook == null)
                return;

            clients.Clear();

            int dateFromCol = FindColumn("От");
            int dateToCol = FindColumn("До");
            IsCancel = false;
            ActionStart?.Invoke("Загрузка файла PriceListMT");

            for (int rw = 2; rw<LastRow; rw++)
            {
                if (IsCancel) return;

                DateTime firstDate = GetDateFromCell(rw, dateFromCol);
                DateTime lastDate = GetDateFromCell(rw, dateToCol);

                if (lastDate.Year >= 9999)
                {
                    AddClient(rw, PriceOfListingColumn);
                }
                else if (date <= lastDate && date >= firstDate)
                {
                    AddClient(rw, PriceNewColumn);
                }
                ActionDone?.Invoke(1);
            }
            IsOpen = true;
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

        private void AddClient(int rw, int priceColumn)
        {
            if (!double.TryParse(GetValueFromColumn(rw, priceColumn), out double price))
            {
                price = 0;
            }

            clients.Add(new Client
            {
                Name = GetValueFromColumn(rw, CustomerColumn),
                Price = price,
                Art = GetValueFromColumn(rw, ArticleColumn)
            });
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
            return worksheet.Cells.Find(fildName, LookAt: XlLookAt.xlWhole)?.Column ?? 0;
        }


        private int FindRow(int column, string articul)
        {
            return worksheet.Columns[column].Find(articul, LookAt: XlLookAt.xlWhole)?.Row ?? 0;
        }

        private int FindRow(int column, string articul, Range afterCell)
        {
            return worksheet.Columns[column].Find(articul, After:afterCell, LookAt: XlLookAt.xlWhole)?.Row ?? 0;
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
