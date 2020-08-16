using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using System.Threading.Tasks;

namespace BPA.NewModel
{
    class RRCItem
    {
        TableRow _row;
        public RRCItem(TableRow row) => _row = row;

        #region Свойства таблицы
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
        public double IRP
        {
            get => _row["IRP, Eur"];
            set => _row["IRP, Eur"] = value;
        }
        public double RRP
        {
            get => _row["RRP, Eur"];
            set => _row["RRP, Eur"] = value;
        }
        public double IRPIndex
        {
            get => _row["IRP index"];
            set => _row["IRP index"] = value;
        }
        public double RRCNDS
        {
            get => _row["РРЦ, руб. с НДС"];
            set => _row["РРЦ, руб. с НДС"] = value;
        }
        public double DIY
        {
            get => _row["DIY price list, руб. без НДС"];
            set => _row["DIY price list, руб. без НДС"] = value;
        }
        public DateTime Date
        {
            get => _row["Дата принятия"];
            set => _row["Дата принятия"] = value;
        }
        #endregion

        public void UpdateRRCFromProduct(ProductItem product, DateTime dateOfPromotion, double budgetCourse)
        {
            if (product != null)
            {
                Date = dateOfPromotion;
                RRCNDS = product.RRCFinal;
                DIY = product.DIY;
                Article = product.Article;
                IRP = product.IRP;
                IRPIndex = product.IRPIndex;
                RRP = RRCNDS / budgetCourse;
            }
        }

        public class RRCItemComparerForPrice : IEqualityComparer<RRCItem>
        {
            public bool Equals(RRCItem x, RRCItem y) => x.Article == y.Article;

            public int GetHashCode(RRCItem obj) => obj.Article.GetHashCode();
        }
    }
}
