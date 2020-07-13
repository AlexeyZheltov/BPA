using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BPA.Model {
    /// <summary>
    /// Справочник STK
    /// </summary>
    class STK : TableBase {
        public override string TableName => "STK";
        public override string SheetName => "STK";

        public static Dictionary<string, int> ColDict { get; set; } = new Dictionary<string, int>();

        #region --- Словарь ---
        public override IDictionary<string, string> Filds {
            get {
                return _filds;
            }
        }
        private readonly Dictionary<string, string> _filds = new Dictionary<string, string>
        {
            { "Id","№" },
            { "Article", "Артикул" },
            { "STKEur", "STK 2.5, Eur" },
            { "STKRub", "STK 2.5, руб." },
            { "Date", "Дата принятия" }
        };
        #endregion

        #region --- Свойства ---
        /// <summary>
        /// №
        /// </summary>
        public int Id
        {
            get; set;
        }

        /// <summary>
        /// Артикул
        /// </summary>
        public string Article {
            get; set;
        }
        /// <summary>
        /// STK 2.5, Eur
        /// </summary>
        public double STKEur {
            get; set;
        }
        
        /// <summary>
        /// STK 2.5, руб.
        /// </summary>
        public double STKRub {
            get; set;
        }
        /// <summary>
        /// Дата принятия
        /// </summary>
        public string Date {
            get; set;
        }
        #endregion

        Opti<DateTime?> PeriodAsDateTime;
        public DateTime? GetPeriodAsDateTime()
        {
            if (!PeriodAsDateTime.isCalculated)
            {
                if (DateTime.TryParse(Date, out DateTime dateTime))
                    PeriodAsDateTime.Value = dateTime;
                else
                    PeriodAsDateTime.Value = null;

                PeriodAsDateTime.isCalculated = true;
            }

            return PeriodAsDateTime.Value;
        }

        struct Opti<T>
        {
            public T Value
            {
                get; set;
            }
            public bool isCalculated
            {
                get; set;
            }
        }

        public STK() { }

        public STK(ListRow listRow) => SetProperty(listRow);
        
        public STK GetSTK(string article, double year)
        {
            ListRow listRow = GetRow("Article", article);

            if (listRow == null)
                return null;

            double firstIndex = listRow.Index;

            do
            {
                STK sTK = new STK(listRow);
                DateTime? date = sTK.GetPeriodAsDateTime();

                if (date != null && (date?.Year ?? -1) == year)
                {
                    return sTK;
                }

                listRow = GetRow("Article", article, listRow.Range[1, Table.ListColumns[Filds["Article"]].Index]);
            } while (listRow.Index != firstIndex);

            return null;
        }
    }
}
