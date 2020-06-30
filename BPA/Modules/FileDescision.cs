using BPA.Model;

using Excel = Microsoft.Office.Interop.Excel;

using System;
using System.Collections.Generic;
using System.IO;
using System.Windows.Forms;

namespace BPA.Modules
{
    class FileDescision
    {
        private readonly string FileName = "";
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
                if (_LastRow == 0) _LastRow = Worksheet.UsedRange.Row + Worksheet.UsedRange.Rows.Count - 1;
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

        #endregion

        public FileDescision()
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

        public bool IsNotOpen() => FileName == "";
        
        public List<Client> LoadClients()
        {
            List<Client> buffer = new List<Client>();

            if (CustomerColumn == 0 || GardenaChannelColumn == 0) throw new ApplicationException("Файл имеет не верный формат");

            for(int rowIndex = FileHeaderRow + 1; rowIndex <= LastRow; rowIndex++)
            {
                if (IsCancel) return null;
                ActionStart?.Invoke($"Обрабатывается строка {rowIndex}");
                Excel.Range range = Worksheet.Cells[rowIndex, CustomerColumn];
                string customer = range.Text;
                if(customer.Trim().Length > 0)
                {
                    range = Worksheet.Cells[rowIndex, GardenaChannelColumn];
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

        public PlanningNewYear LoadPlanningNewYear(PlanningNewYear planningNewYear)
        {
            if (DateColumn == 0 || ArticleColumn == 0)
                throw new ApplicationException("Файл имеет не верный формат");
            
            //временный лист
            //List<PlanningNewYear> buffer = new List<PlanningNewYear>();

            for (int rowIndex = FileHeaderRow + 1; rowIndex <= LastRow; rowIndex++)
            {
                if (IsCancel)
                    return null;
                ActionStart?.Invoke($"Обрабатывается строка {rowIndex}");

                if (planningNewYear.Article != GetValueFromColumn(rowIndex, ArticleColumn))
                    continue;

                var campaign = GetValueFromColumn(rowIndex, CampaignColumn);
                if (campaign == null ||
                    campaign == "")
                {
                    if (int.TryParse(campaign, out int res) && res == 0)
                    continue;
                }

                DateTime date = GetDateFromCell(rowIndex, DateColumn);
                if (planningNewYear.Year != date.Year)
                    continue;

                //Здесь добавляем помесячно
                //date.Month;

                ActionDone?.Invoke(1);
            }

            return planningNewYear;

            //if (buffer.Count == 0)
            //    throw new ApplicationException("Файл не содержит значемых данных");
            //return buffer;
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
