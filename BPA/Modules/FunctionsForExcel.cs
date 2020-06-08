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
    }
}
