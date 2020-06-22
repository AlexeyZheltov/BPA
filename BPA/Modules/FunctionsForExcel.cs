using Excel = Microsoft.Office.Interop.Excel;
using System.Collections.Generic;

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
    }
}
