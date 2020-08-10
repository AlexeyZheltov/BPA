using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BPA.NewModel
{
    class ProductItem
    {
        TableRow _row;
        public ProductItem(TableRow row) => _row = row;

        public int Id
        {
            get => _row["№"];
            set => _row["№"] = value;
        }

        public string Article
        {
            get => _row["Артикул"];
            set => _row["Артикул"] = value;
        }

        public double? IRP
        {
            get => _row["IRP, Eur"];
            set => _row["IRP, Eur"] = value;
        }

        public double? IRPIndex
        {
            get => _row["Индекс IRP"];
            set => _row["Индекс IRP"] = value;
        }

        public double? RRCCurrent
        {
            get => _row["РРЦ текущий"];
            set => _row["РРЦ текущий"] = value;
        }

        public double? DIYCurrent
        {
            get => _row["DIY текущий"];
            set => _row["DIY текущий"] = value;
        }

        public double? RRCCalculated
        {
            get => _row["РРЦ расчетная, руб."];
            set => _row["РРЦ расчетная, руб."] = value;
        }

        public double? RRCFinal
        {
            get => _row["РРЦ финальная, руб."];
            set => _row["РРЦ финальная, руб."] = value;
        }

        public double? DIY
        {
            get => _row["DIY price list, руб. без НДС"];
            set => _row["DIY price list, руб. без НДС"] = value;
        }

        public void UpdatePriceFromRRC(RRCItem item)
        {
            if(item != null)
            {
                RRCCurrent = item.RRCNDS;
                DIYCurrent = item.DIY;
                IRP = item.IRP;
                IRPIndex = item.IRPIndex;
            }
        }
    }
}
