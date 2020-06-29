using Excel = Microsoft.Office.Interop.Excel;
using System.Collections.Generic;
using Microsoft.Office.Interop.Excel;

namespace BPA.Modules
{
    /// <summary>
    /// Класс частоиспользуемых в Excel функций
    /// </summary>
    public static class FunctionsForExcel
    {
        public static Excel.Application Application = Globals.ThisWorkbook?.Application;

        /// <summary>
        /// Ускорение работы Excel
        /// </summary>
        public static void SpeedOn()
        {
            Application.Calculation = Excel.XlCalculation.xlCalculationManual;
            Application.ScreenUpdating = false;
            Application.DisplayAlerts = false;
        }

        /// <summary>
        /// Отключение ускорения работы
        /// </summary>
        public static void SpeedOff()
        {
            Application.Calculation = Excel.XlCalculation.xlCalculationAutomatic;
            Application.ScreenUpdating = true;
            Application.DisplayAlerts = true;
        }

        /// <summary>
        /// Максимальная строка на листе
        /// </summary>
        /// <param name="worksheet">Ссылка на лист</param>
        public static int MaxRow(Excel.Worksheet worksheet)
        {
            return worksheet.UsedRange.Row + worksheet.UsedRange.Rows.Count - 1;
        }

        /// <summary>
        ///  Максимальный столбец
        /// </summary>
        /// <param name="worksheet"></param>
        public static int MaxColumn(Excel.Worksheet worksheet)
        {
            return worksheet.UsedRange.Column + worksheet.UsedRange.Columns.Count - 1;
        }
        
        /// <summary>
        /// Убирает лишние пробельные симовлы и если надо приводит к нижнему регистру
        /// </summary>
        /// <param name="value">Строка в которой нужно удалить пробелы</param>
        /// <param name="toLower">Нужно ли приводить к нижнему регистру. Значение по умолчанию - false</param>
        /// <returns></returns>
        public static string StringNormalize(string value, bool toLower = false)
        {
            value = value.Trim();
            if (toLower) value = value.ToLower();
            return System.Text.RegularExpressions.Regex.Replace(value, @"\s+", " ");
        }

        //Function MS_SheetExist(ByVal NameSheet As String) As Boolean
        //    Dim sh As Object
        //    On Error Resume Next
        //    Set sh = ActiveWorkbook.Sheets(NameSheet)
        //    If Err.Number = 0 Then MS_SheetExist = True
        //End Function

        public static bool IsSheetExists(string sheetName)
        {
            Excel.Worksheet worksheet;
            try
            {
                worksheet = Globals.ThisWorkbook.Sheets[sheetName];
                return true;
            } 
            catch { return false; }
        }

        /// <summary>
        /// Создает и возвращает копию листа в текщей книге с указанным именем или именем по умолчанию
        /// </summary>
        /// <param name="worksheet"></param>
        /// <param name="worksheetName"></param>
        /// <returns></returns>
        public static Excel.Worksheet CreateSheetCopy(Excel.Worksheet worksheet, string copyWorksheetName = "", string afterSheetName = "")
        {
            Excel.Worksheet afterSheet = (afterSheetName != null && IsSheetExists(afterSheetName)) ? worksheet.Parent.Sheets[afterSheetName] : worksheet;

            worksheet.Copy(After: afterSheet);
            Excel.Worksheet newWorksheet = Application.ActiveSheet;

            newWorksheet.Name = nextNumSheet(copyWorksheetName);

            return newWorksheet;
        }

        /// <summary>
        /// проверка наличия листа с текущим именем
        /// </summary>
        /// <param name="sheetName"></param>
        /// <returns></returns>
        private static string nextNumSheet(string sheetName)
        {
            if (!IsSheetExists(sheetName))
                return sheetName;

            if (double.TryParse(sheetName.Substring(sheetName.Length - 1, 1), out double strNum))
            {
                strNum++;
                sheetName = sheetName.Substring(0, sheetName.Length - 2);
            }
            else
            {
                strNum = 1;
            }

            sheetName=sheetName + "_" + strNum.ToString();                    
            return nextNumSheet(sheetName);
        }


        public static void HideShowSettingsSheets()
        {
            List<string> AlwaysShowSheets = new List<string>
            {
                "Товары",
                "Клиенты",
                "Скидки",
                "РРЦ",
                "DIY",
                "Планирование",
                "Прайс лист"
            };

            Excel.XlSheetVisibility status = Excel.XlSheetVisibility.xlSheetVisible;
            foreach (Excel.Worksheet sheet in Globals.ThisWorkbook.Sheets)
                if (!AlwaysShowSheets.Contains(sheet.Name) && sheet.Visible == Excel.XlSheetVisibility.xlSheetVisible)
                {
                    status = Excel.XlSheetVisibility.xlSheetHidden;
                    break;
                }

            foreach (Excel.Worksheet sheet in Globals.ThisWorkbook.Sheets)
                if (!AlwaysShowSheets.Contains(sheet.Name))
                    sheet.Visible = status;
        }

        public static void ShowSheet(string sheetName)
        {
            if (!IsSheetExists(sheetName))
                return;

            Worksheet worksheet = Globals.ThisWorkbook.Sheets[sheetName];
            if (worksheet.Visible == XlSheetVisibility.xlSheetHidden)
            {
                worksheet.Visible = XlSheetVisibility.xlSheetVisible;
            }
        }

        public static void HideSheet(string sheetName)
        {
            if (!IsSheetExists(sheetName))
                return;

            Worksheet worksheet = Globals.ThisWorkbook.Sheets[sheetName];
            if (worksheet.Visible == XlSheetVisibility.xlSheetVisible)
            {
                worksheet.Visible = XlSheetVisibility.xlSheetHidden;
            }
        }
    }
}
