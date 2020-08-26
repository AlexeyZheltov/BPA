using Excel = Microsoft.Office.Interop.Excel;
//using Application = Microsoft.Office.Interop.Excel.Application;
using System;
using System.Windows.Forms;
using BPA.Forms;

namespace BPA.Modules
{
    class FileBase
    {
        public readonly Excel.Application Application = Globals.ThisWorkbook.Application;
        protected string FileSheetName = "";
        protected int FileHeaderRow = 1;
        private int ActionCounter = 1;

        public void SetProcessBarForLoad(ref ProcessBar pB)
        {
            pB = new ProcessBar($"Загрузка файла  { FileName } ", CountActions);
            pB.CancelClick += Cancel;
            ActionStart += pB.TaskStart;
            ActionDone += pB.TaskDone;
            pB.Show(new ExcelWindows(Globals.ThisWorkbook));
        }

        /// <summary>
        /// Событие начала задачи
        /// </summary>
        public Action<string> ActionStart;
        /// <summary>
        /// Событие завершения задачи
        /// </summary>
        private Action<int> ActionDone;
        public void ActionD()
        {
            if (ActionCounter++ >= coef)
            {
                ActionDone?.Invoke(1);
                ActionCounter = 1;
            }
            //ActionCounter++;
        }

        public int coef = 1;
        public int CountActions
        {
            get
            {
                double digit = Math.Log10(Convert.ToDouble(ArrRrows - 1));
                int pow = 1;

                if (digit >= 8)
                    pow = 7;
                else if (digit >= 5)
                    pow = 4;

                coef = Convert.ToInt32(Math.Pow(10, pow));

                return ArrRrows / coef;
            }
            set => CountActions = value;
        }
        //public int CountActions => ArrRrows;
        public bool IsCancel = false;

        public bool IsOpen { get; set; } = false;

        public string FileAddress = "";
        public string FileName
        {
            get
            {
                if (_FileName == null)
                {
                    try
                    {
                        _FileName = Workbook.Name;
                    }
                    catch
                    {
                        _FileName = null;
                    }
                }
                return _FileName;
            }
            set
            {
                _FileName = value;
            }
        }
        private string _FileName;

        public Excel.Workbook Workbook
        {
            get
            {
                if (_Workbook == null)
                {
                    try
                    {
                        _Workbook = Application.Workbooks.Open(FileAddress);
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

        public Excel.Worksheet worksheet
        {
            get
            {
                if (_worksheet == null)
                {
                    try
                    {
                        _worksheet = FileSheetName != "" ? Workbook?.Sheets[FileSheetName] : Workbook?.Sheets[1];
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

        private int LastColumn
        {
            get
            {
                if (_LastColumn == 0) _LastColumn = worksheet.Cells[FileHeaderRow, worksheet.Columns.Count].End[Microsoft.Office.Interop.Excel.XlDirection.xlToLeft].Column;
                    //_LastColumn = worksheet.Cells[FileHeaderRow, worksheet.UsedRange.Columns.Count].Column;
                return _LastColumn;
            }
        }
        private int _LastColumn = 0;

        private object[,] FileArray;
        public int ArrRrows;
        private int ArrColumns;
        /// <summary>
        /// Загрузка массива данных
        /// </summary>
        public void SetFileData()
        {
            BPA.Forms.WaitForm waitForm = new Forms.WaitForm();
            waitForm.Show();
            System.Windows.Forms.Application.DoEvents();
            try
            {
                FileArray = worksheet.Range[worksheet.Cells[FileHeaderRow, 1], worksheet.Cells[LastRow, LastColumn]].Value;
                ArrRrows = FileArray.GetUpperBound(0);
                ArrColumns = FileArray.GetLength(1);
            }
            catch (Exception ex)
            {
                throw new ApplicationException(ex.Message);
            }
            finally
            {
                waitForm.Close();
            }
        }

        /// <summary>
        /// Загрузка массива данных
        /// </summary>
        public void SetFileData(params string[] colNames)
        {
            BPA.Forms.WaitForm waitForm = new Forms.WaitForm();
            waitForm.Show();
            System.Windows.Forms.Application.DoEvents();
            try
            {
                object[,] headers = worksheet.Range[worksheet.Cells[FileHeaderRow, 1], worksheet.Cells[FileHeaderRow, LastColumn]].Value;
                int newCol = 1;

                foreach (string colName in colNames)
                    for (int c = 1; c <= headers.Length; c++)
                        if (colName == headers[1, c].ToString())
                        {
                            SetDataDescision(c);
                            newCol++;
                            break;
                        }
                ArrRrows = FileArray.GetUpperBound(0);
                ArrColumns = FileArray.GetLength(1);

                void SetDataDescision(int colNum)
                {
                    object[,] bufer = worksheet.Range[worksheet.Cells[FileHeaderRow, colNum], worksheet.Cells[LastRow, colNum]].Value;
                    if (FileArray == null)
                    {
                        //Нужен массив от 1
                        Array rv = Array.CreateInstance(typeof(object), new int[] { LastRow, colNames.Length }, new int[] { 1, 1 });
                        FileArray = rv as object[,];
                    }

                    for (int r = 1; r <= bufer.Length; r++)
                        FileArray[r, newCol] = bufer[r, 1];
                }

            } catch(Exception ex)
            {
                throw new ApplicationException(ex.Message);
            }
            finally
            {
                waitForm.Close();
            }
        }

        public void ClearData()
        {
            FileArray = null;
        }

        /// <summary>
        /// Загрузка массивы данных с одной строки
        /// </summary>
        /// <param name="row"></param>
        public void SetFileData(int row)
        {
            FileArray = worksheet.Range[worksheet.Cells[row, 1], worksheet.Cells[row, LastColumn]].Value;
            ArrRrows = FileArray.GetUpperBound(0);
            ArrColumns = FileArray.GetLength(1);
        }

        public FileBase() { }

        /// <summary>
        /// получение номена строки по имени заголовка
        /// </summary>
        /// <param name="fildName"></param>
        /// <returns></returns>
        public int FindColumn(string fildName)
        {
            for (int col = 1; col <= ArrColumns; col++)
            {
                object obj = FileArray[1, col];
                if ((obj is string) && Convert.ToString(obj) == fildName)
                    return col;
            }
            return 0;
        }

        /// <summary>
        /// Поиск значения String в столбце
        /// </summary>
        /// <param name="str"></param>
        /// <param name="col"></param>
        /// <returns></returns>
        public int FindRow(string str, int col)
        {
            for (int rowIndex = 2; rowIndex < ArrRrows; rowIndex++)
            {
                object obj = FileArray[rowIndex, col];
                if (obj is string && Convert.ToString(obj) == str) return rowIndex;
            }
            return 0;
        }

        /// <summary>
        /// Поиск значения Double в столбце
        /// </summary>
        /// <param name="dbl"></param>
        /// <param name="col"></param>
        /// <returns></returns>
        public int FindRow(double dbl, int col)
        {
            for (int rowIndex = 2; rowIndex < ArrRrows; rowIndex++)
            {
                object obj = FileArray[rowIndex, col];
                if (obj is double && Convert.ToDouble(obj) == dbl) return rowIndex;
            }
            return 0;
        }

        /// <summary>
        /// Поиск значения DataTime в столбце
        /// </summary>
        /// <param name="dt"></param>
        /// <param name="col"></param>
        /// <returns></returns>
        public int FindRow(DateTime dt, int col)
        {
            for (int rowIndex = 2; rowIndex < ArrRrows; rowIndex++)
            {
                object obj = FileArray[rowIndex, col];
                if (obj is DateTime && Convert.ToDateTime(obj) == dt) return rowIndex;
            }
            return 0;
        }
        /// <summary>
        /// Получение даты из ячейки
        /// </summary>
        /// <param name="rw"></param>
        /// <param name="col"></param>
        /// <returns></returns>
        public DateTime GetDateFromCell(int rw, int col)
        {
            object obj = GetValueFromColumn(rw, col);

            if (obj == null)
                return new DateTime();

            if (obj is double)
                return DateTime.FromOADate(Convert.ToDouble(obj));
            else if (obj is DateTime)
                return Convert.ToDateTime(obj);
            else
                return new DateTime();
        }

        /// <summary>
        /// получение значения из строки по номеру столбца
        /// </summary>
        /// <param name="rw"></param>
        /// <param name="col"></param>
        /// <returns></returns>
        public object GetValueFromColumn(int rw, int col) {
            return col != 0 ? FileArray[rw, col] : null;
        }

        public string GetValueFromColumnStr(int rw, int col)
        {
            object obj = FileArray[rw, col];
            //return obj.ToString();
            //return obj is null ? null : obj.ToString();
            return obj?.ToString() ?? "";
            //return obj is string ? Convert.ToString(obj) : "";
        }
        public double GetValueFromColumnDbl(int rw, int col)
        {
            object obj = FileArray[rw, col];
            if (obj is double) return Convert.ToDouble(obj);

            string objStr = GetValueFromColumnStr(rw, col).Trim();
            return Double.TryParse(objStr, out double dbl) ? dbl : 0;
        }
        public DateTime GetValueFromColumnDT(int rw, int col)
        {
            object obj = FileArray[rw, col];
            return obj is DateTime ? Convert.ToDateTime(obj) : new DateTime();
        }

        /// <summary>
        /// Возвращает значение DataTime в формате Double
        /// </summary>
        /// <param name="rw"></param>
        /// <param name="col"></param>
        /// <returns></returns>
        public double GetDoubleDateFromCell(int rw, int col)
        {
            object obj = FileArray[rw, col];

            if (obj is DateTime)
            {
                DateTime tmpDateTime = Convert.ToDateTime(obj);
                if (tmpDateTime.ToOADate() > 0)
                    return tmpDateTime.ToOADate();
            }
            return 0;
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
