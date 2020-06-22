using Microsoft.Office.Interop.Excel;

namespace BPA.Modules
{
    /// <summary>
    /// Класс частоиспользуемых в Excel функций
    /// </summary>
    public static class FunctionsForExcel
    {
        public static Application Application = Globals.ThisWorkbook?.Application;

        /// <summary>
        /// Ускорение работы Excel
        /// </summary>
        public static void SpeedOn()
        {
            Application.Calculation = XlCalculation.xlCalculationManual;
            Application.ScreenUpdating = false;
            Application.DisplayAlerts = false;
        }

        /// <summary>
        /// Отключение ускорения работы
        /// </summary>
        public static void SpeedOff()
        {
            Application.Calculation = XlCalculation.xlCalculationAutomatic;
            Application.ScreenUpdating = true;
            Application.DisplayAlerts = true;
        }

        /// <summary>
        /// Максимальная строка на листе
        /// </summary>
        /// <param name="worksheet">Ссылка на лист</param>
        public static int MaxRow(Worksheet worksheet)
        {
            return worksheet.UsedRange.Row + worksheet.UsedRange.Rows.Count - 1;
        }

        /// <summary>
        ///  Максимальный столбец
        /// </summary>
        /// <param name="worksheet"></param>
        public static int MaxColumn(Worksheet worksheet)
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
            Worksheet worksheet;
            try
            {
                worksheet = Globals.ThisWorkbook.Sheets[sheetName];
                return true;
            } 
            catch { return false; }
        }
    }
}
