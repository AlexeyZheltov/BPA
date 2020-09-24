using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BPA.NewModel
{
    class DiscountForPlanningItem
    {
        TableRow _row;
        public DiscountForPlanningItem(TableRow row) => _row = row;

        #region Свойства таблицы
        public int Id
        {
            get => _row["№"];
            set => _row["№"] = value;
        }
        public string ChannelType
        {
            get => _row["Channel type"];
            set => _row["Channel type"] = value;
        }
        public string CustomerStatusForecast
        {
            get => _row["Customer status for forecast"];
            set => _row["Customer status for forecast"] = value;
        }
        public DateTime Period
        {
            get => _row["Период"];
            set => _row["Период"] = value;
        }
        public string IrrigationEquipments
        {
            get => _row["Оборудование для полива"];
            set => _row["Оборудование для полива"] = value;
        }
        public string Electricians
        {
            get => _row["Электрика (Готовая продукция)"];
            set => _row["Электрика (Готовая продукция)"] = value;
        }
        public string Lawnmowers
        {
            get => _row["Газонокосилки-роботы"];
            set => _row["Газонокосилки-роботы"] = value;
        }
        public string Pumps
        {
            get => _row["Насосное оборудование"];
            set => _row["Насосное оборудование"] = value;
        }
        public string CuttingTools
        {
            get => _row["Ручные и режущие инструменты"];
            set => _row["Ручные и режущие инструменты"] = value;
        }
        public string WinterTools
        {
            get => _row["Зимние инструменты ClassicLine"];
            set => _row["Зимние инструменты ClassicLine"] = value;
        }
        public double MaximumBonus
        {
            get => _row["Максимальный годовой бонус"];
            set => _row["Максимальный годовой бонус"] = value;
        }
        #endregion

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

        public string GetFormulaByName(string name)
        {
            if (_row.ColumnExsists(name)) return _row[name];
            else return "";
        }
    }
}
