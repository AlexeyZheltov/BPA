using BPA.Modules;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace BPA.Model {
    /// <summary>
    /// Справочник скидок
    /// </summary>
    class Discount : TableBase {
        public override string TableName => "Скидки";
        public override string SheetName => "Скидки";
        public static Dictionary<string, int> ColDict { get; set; } = new Dictionary<string, int>();


        public override IDictionary<string, string> Filds {
            get {
                return _filds;
            }
        }
        private readonly Dictionary<string, string> _filds = new Dictionary<string, string>
        {
            { "Id", "№" },
            { "ChannelType", "Channel type" },
            { "CustomerStatus", "Customer status" },
            { "Period", "Период" },
            { "IrrigationEquipments", "Оборудование для полива" },
            { "Electricians", "Электрика (Готовая продукция)" },
            { "Lawnmowers", "Газонокосилки-роботы" },
            { "Pumps", "Насосное оборудование" },
            { "CuttingTools", "Ручные и режущие инструменты" },
            { "WinterTools", "Зимние инструменты ClassicLine" },
            { "MaximumBonus", "Максимальный годовой бонус" }
        };

        public Discount() { }

        public Discount(Excel.ListRow row) => SetProperty(row);

        public string GetFormulaByName(string name)
        {
            foreach(KeyValuePair<string, string> item in _filds)
                if (item.Value == name)
                    return (string)GetParametrValue(item.Key);
            return "";
        }

        #region --- Свойства ---

        /// <summary>
        /// №
        /// </summary>
        public int Id
        {
            get; set;
        }

        /// <summary>
        /// Channel type
        /// </summary>
        public string ChannelType {
            get; set;
        }

        /// <summary>
        /// CustomerStatus
        /// </summary>
        public string CustomerStatus {
            get; set;
        }

        /// <summary>
        /// Period
        /// </summary>
        public string Period {
            get; set;
        }

        /// <summary>
        /// Оборудование для полива
        /// </summary>
        public string IrrigationEquipments {
            get; set;
        }

        /// <summary>
        /// Электрика (Готовая продукция)
        /// </summary>
        public string Electricians {
            get; set;
        }

        /// <summary>
        /// Газонокосилки-роботы
        /// </summary>
        public string Lawnmowers {
            get; set;
        }

        /// <summary>
        /// Насосное оборудование
        /// </summary>
        public string Pumps {
            get; set;
        }

        /// <summary>
        /// Ручные и режущие инструменты
        /// </summary>
        public string CuttingTools {
            get; set;
        }

        /// <summary>
        /// Зимние инструменты ClassicLine
        /// </summary>
        public string WinterTools {
            get; set;
        }

        /// <summary>
        /// Максимальный годовой бонус
        /// </summary>
        public double MaximumBonus {
            get; set;
        }
        #endregion

        Opti<DateTime?> PeriodAsDateTime;
        public DateTime? GetPeriodAsDateTime()
        {
            if (!PeriodAsDateTime.isCalculated)
            {
                if (DateTime.TryParse(Period, out DateTime dateTime))
                    PeriodAsDateTime.Value = dateTime;
                else
                    PeriodAsDateTime.Value = null;

                PeriodAsDateTime.isCalculated = true;
            }

            return PeriodAsDateTime.Value;
        }

        public static List<Discount> GetAllDiscounts(PBWrapper pB)
        {
            List<Discount> discounts = new List<Discount>();
            Discount discount = new Discount();
            pB.Start(discount.Table.ListRows.Count);

            foreach(Excel.ListRow row in discount.Table.ListRows)
            {
                if (pB.IsCancel)
                {
                    pB.Dispose();
                    return null;
                }
                pB.Action($"{row.Index}");
                discounts.Add(new Discount(row));
                pB.Done(1);
            }
            pB.Dispose();
            return discounts;
        }

        public static Discount GetCurrentDiscount(Client client, DateTime currentDate)
        {
            new Discount().ReadColNumbers();

            List<Discount> discounts = Discount.GetAllDiscounts(new PBWrapper($"Создание прайс-листа для {client.Customer}", "Чтение скидок [Index]"));
            if (discounts == null) return null;
            discounts = discounts.FindAll(x => x.ChannelType == client.ChannelType
                                                && x.CustomerStatus == client.CustomerStatus
                                                && x.GetPeriodAsDateTime() != null
                                                && x.GetPeriodAsDateTime() <= currentDate);

            discounts.Sort((x, y) =>
            {
                if (x.GetPeriodAsDateTime() < y.GetPeriodAsDateTime()) return 1;
                else if (x.GetPeriodAsDateTime() > y.GetPeriodAsDateTime()) return -1;
                else return 0;
            });

            if (discounts.Count == 0)
            {
                //MessageBox.Show($"Клиенту {client.Customer} нет соответствий на листе \"Скидки\"", "BPA", MessageBoxButtons.OK, MessageBoxIcon.Information);
                Debug.Print($"Клиенту {client.Customer} нет соответствий на листе \"Скидки\"", "BPA", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return null;
            }
            Discount currentDiscount = discounts[0];

            //проверить формулы
            //Убрать пробелы и лишние знаки
            currentDiscount.NormaliseAllFormulas();

            return currentDiscount;
        }

        public void NormaliseAllFormulas()
        {
            IrrigationEquipments = FormulaNormalize(IrrigationEquipments);
            Electricians = FormulaNormalize(Electricians);
            Lawnmowers = FormulaNormalize(Lawnmowers);
            Pumps = FormulaNormalize(Pumps);
            CuttingTools = FormulaNormalize(CuttingTools);
            WinterTools = FormulaNormalize(WinterTools);
        }

        string FormulaNormalize(string value, bool RemoveMarks = false)
        {
            //оставить только [метка], а вне ее только [1-9], +, - , *, /, (), %, =
            StringBuilder builder = new StringBuilder();
            bool isMark = false;

            value = value.ToLower();
            foreach (char ch in value.ToCharArray())
            {
                if (ch == '[' & !RemoveMarks) isMark = true;
                else if (ch == ']' & isMark)
                {
                    builder.Append(ch);
                    isMark = false;
                }

                if (!isMark)
                {
                    if (Char.IsDigit(ch)) builder.Append(ch);
                    else
                    {
                        switch (ch)
                        {
                            case '+':
                            case '-':
                            case '*':
                            case '/':
                            case '(':
                            case ')':
                            case '%':
                            case '=':
                                builder.Append(ch);
                                break;
                            case ',':
                            case '.':
                                builder.Append('.');
                                break;
                        }
                    }

                }
                else builder.Append(ch);
            }

            string temp = System.Text.RegularExpressions.Regex.Replace(builder.ToString(), @"\s+", " ");
            return temp;
        }

        public bool NeedFilePriceMT()
        {
            return IrrigationEquipments.Contains("[pricelist mt]") ||
                    Electricians.Contains("[pricelist mt]") ||
                    Lawnmowers.Contains("[pricelist mt]") ||
                    Pumps.Contains("[pricelist mt]") ||
                    CuttingTools.Contains("[pricelist mt]") ||
                    WinterTools.Contains("[pricelist mt]");
        }

        struct Opti<T>
        {
            public T Value { get; set; }
            public bool isCalculated { get; set; }
        }

        public double GetDiscountForPlanning(PlanningNewYear planning)
        {
            ListRow listRow = GetRow("ChannelType", planning.ChannelType);
            if (listRow == null) return 0;

            double firstIndex = listRow.Index;

            do
            {
                Discount discount = new Discount(listRow);
                if (discount.CustomerStatus == planning.CustomerStatus)
                {
                    DateTime? date = discount.GetPeriodAsDateTime();
                    
                    if (date != null && (date?.Year ?? -1) == planning.Year)
                    {
                        return discount.MaximumBonus;                    
                    }
                }
              
                listRow = GetRow("ChannelType", planning.ChannelType, listRow.Range[1, Table.ListColumns[Filds["ChannelType"]].Index]);
            } while (listRow.Index != firstIndex);

            return 0;
        }
    }
}
