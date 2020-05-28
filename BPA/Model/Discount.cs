using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BPA.Model {
    /// <summary>
    /// Справочник скидок
    /// </summary>
    class Discount : TableBase {
        public override string TableName => "Скидки";
        public override string SheetName => "Скидки";

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
        public int CustomerStatus {
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
        public string MaximumBonus {
            get; set;
        }

    }
}
