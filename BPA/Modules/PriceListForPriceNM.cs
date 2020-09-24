using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NM = BPA.NewModel;

namespace BPA.Modules
{
    class PriceListForPriceNM
    {
        /// <summary>
        /// изменяется после SetProduct
        /// </summary>
        public bool FormulaChecked;

        private FilePriceMT filePriceMT;
        private NM.DiscountItem Discount;
        private NM.ProductItem Product;
        private string Formula;

        public PriceListForPriceNM(FilePriceMT filePriceMT, NM.DiscountItem discount) 
        {
            this.filePriceMT = filePriceMT;
            this.Discount = discount;
        }

        private PriceListForPriceNM() { }
        /// <summary>
        /// Устанвка формулы исходя из категории товара, предварительно необходимо установить Discount (в конструкторе)
        /// </summary>
        /// <param name="product"></param>
        public void SetProduct(NM.ProductItem product)
        {
            if (product.Category == "") throw new ApplicationException($"Для {product.Article} не указана категория");

            this.Product = product;
            this.Formula = Discount.GetFormulaByName(product.Category);
            this.FormulaChecked = true;
        }
        public void CheckFormula()
        {

        }
        /// <summary>
        /// Получение цены, предварительно необходимо установить фрмулу и Продукт (SetProduct), и filePriceMT (в конструкторе)
        /// </summary>
        /// <param name="rrcList"></param>
        /// <returns></returns>
        public double GetPrice(List<NM.RRCItem> rrcList)
        {
            try
            {
                string formula = this.Formula;
                //if (formula = "" ) 
                //Найти метку или метки. [Pricelist MT]  [DIY Pricelist] [РРЦ] и заменить
                while (formula.Contains("[pricelist mt]"))
                    formula = formula.Replace("[pricelist mt]", filePriceMT.GetPrice(Product.Article).ToString());

                while (formula.Contains("[diy price list]"))
                    formula = formula.Replace("[diy price list]", rrcList.Find(x => x.Article == Product.Article)?.DIY.ToString() ?? "0");

                while (formula.Contains("[ррц]"))
                    formula = formula.Replace("[ррц]", rrcList.Find(x => x.Article == Product.Article)?.RRCNDS.ToString() ?? "0");


                if (Parsing.Calculation(formula) is double result)
                {
                    return result;
                }
                else
                {
                    throw new ApplicationException($"В одной из формул содержится ошибка");
                }
            }
            catch
            {
                throw new ApplicationException($"В одной из формул содержится ошибка");
            }
        }
    }
}
