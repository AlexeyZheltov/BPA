using BPA.Forms;
using BPA.Model;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BPA.Modules
{
    class PriceListForPlaning : IDisposable//Выполнить чтение сталбцов RRC
    {
        readonly Client currentClient;
        readonly DateTime currentDate;
        List<FinalPriceList> priceList;
        FilePriceMT filePriceMT = null;
        bool isLoaded = false;

        public PriceListForPlaning(PlanningNewYear planningNewYear)
        {
            if (planningNewYear.Clients != null && planningNewYear.Clients.Count > 0)
            {
                if (planningNewYear.Clients[0] is Client cl && planningNewYear.CurrentDate is DateTime dt)
                {
                    currentClient = cl;
                    currentDate = dt;
                } else throw new ArgumentException();
            }
            else throw new ArgumentException();
        }

        /// <summary>
        /// Загружает необходимые данные
        /// </summary>
        public void Load()
        {
            List<RRC> actualRRC = RRC.GetActualPriceList(currentDate);
            if (actualRRC == null) throw new ApplicationException("Не удалось загрузить данные с листа RRC");

            Discount currentDiscount = Discount.GetCurrentDiscount(currentClient, currentDate);
            if (currentDiscount == null) throw new ApplicationException("Не удалось загрузить данные с с листа скидок");

            //подгрузить PriceMT если неужно, подключится к РРЦ                   
            if (currentDiscount.NeedFilePriceMT())
            {
                //Загурзить файл price list MT
                ProcessBar processBar = null;
                filePriceMT = new FilePriceMT();
                if (!filePriceMT.IsOpen)
                    return;
                filePriceMT.SetProcessBarForLoad(ref processBar);
                filePriceMT.SetFileData();
                filePriceMT.Load(currentDate);
                processBar.Close();
                if (!filePriceMT.IsOpen) throw new ApplicationException("Не удалось загрузить File PriceListMT");
            }

            //Загрузка списка артикулов, какие из них актуальные?
            List<Product> products = Product.GetProductForClient(currentClient);
            if (products == null) return;

            //в цикле менять метки на значения из цен, с заменой;
            priceList = new List<FinalPriceList>();

            foreach (Product product in products)
            {
                if (product.Category == "")
                {
                    if (filePriceMT?.IsOpen ?? false) filePriceMT.Close();
                    throw new ApplicationException($"Для {product.Article} не указана категория");
                }
                //получить формулу
                string formula = currentDiscount.GetFormulaByName(product.Category);

                try
                {
                    //Найти метку или метки. [Pricelist MT]  [DIY Pricelist] [РРЦ] и заменить
                    while (formula.Contains("[pricelist mt]"))
                        formula = formula.Replace("[pricelist mt]", filePriceMT.GetPrice(product.Article).ToString());

                    while (formula.Contains("[diy price list]"))
                        formula = formula.Replace("[diy price list]", actualRRC.Find(x => x.Article == product.Article).DIY.ToString());

                    while (formula.Contains("[ррц]"))
                        formula = formula.Replace("[ррц]", actualRRC.Find(x => x.Article == product.Article).RRCNDS.ToString());
                    if (Parsing.Calculation(formula) is double result)
                        priceList.Add(new FinalPriceList(product)
                        {
                            RRC = result
                        });
                } 
                catch
                {
                    if (filePriceMT?.IsOpen ?? false) filePriceMT.Close();
                    throw new ApplicationException($"В одной из формул для {currentClient.Customer} содержится ошибка");
                    //Debug.Print($"В одной из формул для {currentClient.Customer} содержится ошибка");
                    //continue;
                }
            }
            if (filePriceMT?.IsOpen ?? false) filePriceMT.Close();
            filePriceMT = null;
            isLoaded = true;
        }

        /// <summary>
        /// Возвращает цену артикула
        /// </summary>
        /// <param name="article">артикул</param>
        /// <returns></returns>
        public double GetPrice(string article)
        {
            if (!isLoaded) throw new ApplicationException("Не выполнена загрузка Price List Planing");

            if (priceList.Exists(x => x.ArticleGardena == article))
                return priceList.Find(x => x.ArticleGardena == article).RRC;
            else
                return 0.0;
        }

        /// <summary>
        /// Освобождает ресурсы
        /// </summary>
        public void Dispose()
        {
            filePriceMT?.Close();
        }
    }
}
