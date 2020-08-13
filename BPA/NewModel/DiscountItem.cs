using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BPA.NewModel
{
    class DiscountItem
    {
        TableRow _row;
        public DiscountItem(TableRow row) => _row = row;

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
        public string CustomerStatus
        {
            get => _row["Customer status"];
            set => _row["Customer status"] = value;
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

    }
}
