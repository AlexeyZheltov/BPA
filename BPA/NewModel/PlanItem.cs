using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BPA.NewModel
{
    class PlanItem
    {
        TableRow _row;
        public PlanItem(TableRow row) => _row = row;

        public int Id
        {
            get => _row["№"];
            set => _row["№"] = value;
        }

        public string Article 
        {
            set => _row["Артикул"] = value;
            get => _row["Артикул"];
        }
        public string ChannelType 
        {
            get => _row ["Channel type"];
            set => _row ["Channel type"] = value;
        }
        public string CustomerStatus 
        {
            get => _row ["Customer status"];
            set => _row ["Customer status"] = value;
        }
        public DateTime PrognosisDate 
        {
            get => _row ["Дата прогноза"];
            set => _row ["Дата прогноза"] = value;
        }
        public DateTime Data 
        {
            get => _row ["Данные справочника товаров"];
            set => _row ["Данные справочника товаров"] = value;
        }
        public double STKRub 
        {
            get => _row ["STK 2.5, руб."];
            set => _row ["STK 2.5, руб."] = value;
        }
        public double IRPEur 
        {
            get => _row ["IRP, Eur"];
            set => _row ["IRP, Eur"] = value;
        }
        public double RRC 
        {
            get => _row ["РРЦ, руб.с НДС"];
            set => _row ["РРЦ, руб.с НДС"] = value;
        }
        public double IRPIndex 
        {
            get => _row ["Индекс IRP"];
            set => _row ["Индекс IRP"] = value;
        }
        public double DIY 
        {
            get => _row ["DIY price list, руб. без НДС"];
            set => _row ["DIY price list, руб. без НДС"] = value;
        }
        public double Price2Net 
        {
            get => _row ["2 net цена, руб."];
            set => _row ["2 net цена, руб."] = value;
        }
        public double Price3Net 
        {
            get => _row ["3 net цена, руб."];
            set => _row ["3 net цена, руб."] = value;
        }
        public double PriceTransfer 
        {
            get => _row ["Transfer цена, руб."];
            set => _row ["Transfer цена, руб."] = value;
        }

        public double PriceList01 
        {
            get => _row ["Price-list январь"];
            set => _row ["Price-list январь"] = value;
        }
        public double PriceList02 
        {
            get => _row ["Price-list февраль"];
            set => _row ["Price-list февраль"] = value;
        }
        public double PriceList03 
        {
            get => _row ["Price-list март"];
            set => _row ["Price-list март"] = value;
        }
        public double PriceList04 
        {
            get => _row ["Price-list апрель"];
            set => _row ["Price-list апрель"] = value;
        }
        public double PriceList05 
        {
            get => _row ["Price-list май"];
            set => _row ["Price-list май"] = value;
        }
        public double PriceList06 
        {
            get => _row ["Price-list июнь"];
            set => _row ["Price-list июнь"] = value;
        }
        public double PriceList07 
        {
            get => _row ["Price-list июль"];
            set => _row ["Price-list июль"] = value;
        }
        public double PriceList08 
        {
            get => _row ["Price-list август"];
            set => _row ["Price-list август"] = value;
        }
        public double PriceList09 
        {
            get => _row ["Price-list сентябрь"];
            set => _row ["Price-list сентябрь"] = value;
        }
        public double PriceList10 
        {
            get => _row ["Price-list октябрь"];
            set => _row ["Price-list октябрь"] = value;
        }
        public double PriceList11 
        {
            get => _row ["Price-list ноябрь"];
            set => _row ["Price-list ноябрь"] = value;
        }
        public double PriceList12 
        {
            get => _row ["Price-list декабрь"];
            set => _row ["Price-list декабрь"] = value;
        }

        public double SalesVolume01 
        {
            get => _row ["Объем продаж январь"];
            set => _row ["Объем продаж январь"] = value;
        }
        public double SalesVolume02 
        {
            get => _row ["Объем продаж февраль"];
            set => _row ["Объем продаж февраль"] = value;
        }
        public double SalesVolume03 
        {
            get => _row ["Объем продаж март"];
            set => _row ["Объем продаж март"] = value;
        }
        public double SalesVolume04 
        {
            get => _row ["Объем продаж апрель"];
            set => _row ["Объем продаж апрель"] = value;
        }
        public double SalesVolume05 
        {
            get => _row ["Объем продаж май"];
            set => _row ["Объем продаж май"] = value;
        }
        public double SalesVolume06 
        {
            get => _row ["Объем продаж июнь"];
            set => _row ["Объем продаж июнь"] = value;
        }
        public double SalesVolume07 
        {
            get => _row ["Объем продаж июль"];
            set => _row ["Объем продаж июль"] = value;
        }
        public double SalesVolume08 
        {
            get => _row ["Объем продаж август"];
            set => _row ["Объем продаж август"] = value;
        }
        public double SalesVolume09 
        {
            get => _row ["Объем продаж сентябрь"];
            set => _row ["Объем продаж сентябрь"] = value;
        }
        public double SalesVolume10 
        {
            get => _row ["Объем продаж октябрь"];
            set => _row ["Объем продаж октябрь"] = value;
        }
        public double SalesVolume11 
        {
            get => _row ["Объем продаж ноябрь"];
            set => _row ["Объем продаж ноябрь"] = value;
        }
        public double SalesVolume12 
        {
            get => _row ["Объем продаж декабрь"];
            set => _row ["Объем продаж декабрь"] = value;
        }

        public double GS01 
        {
            get => _row ["GS январь"];
            set => _row ["GS январь"] = value;
        }
        public double GS02 
        {
            get => _row ["GS февраль"];
            set => _row ["GS февраль"] = value;
        }
        public double GS03 
        {
            get => _row ["GS март"];
            set => _row ["GS март"] = value;
        }
        public double GS04 
        {
            get => _row ["GS апрель"];
            set => _row ["GS апрель"] = value;
        }
        public double GS05 
        {
            get => _row ["GS май"];
            set => _row ["GS май"] = value;
        }
        public double GS06 
        {
            get => _row ["GS июнь"];
            set => _row ["GS июнь"] = value;
        }
        public double GS07 
        {
            get => _row ["GS июль"];
            set => _row ["GS июль"] = value;
        }
        public double GS08 
        {
            get => _row ["GS август"];
            set => _row ["GS август"] = value;
        }
        public double GS09 
        {
            get => _row ["GS сентябрь"];
            set => _row ["GS сентябрь"] = value;
        }
        public double GS10 
        {
            get => _row ["GS октябрь"];
            set => _row ["GS октябрь"] = value;
        }
        public double GS11 
        {
            get => _row ["GS ноябрь"];
            set => _row ["GS ноябрь"] = value;
        }
        public double GS12 
        {
            get => _row ["GS декабрь"];
            set => _row ["GS декабрь"] = value;
        }

        public double NS01 
        {
            get => _row ["NS январь"];
            set => _row ["NS январь"] = value;
        }
        public double NS02 
        {
            get => _row ["NS февраль"];
            set => _row ["NS февраль"] = value;
        }
        public double NS03 
        {
            get => _row ["NS март"];
            set => _row ["NS март"] = value;
        }
        public double NS04 
        {
            get => _row ["NS апрель"];
            set => _row ["NS апрель"] = value;
        }
        public double NS05 
        {
            get => _row ["NS май"];
            set => _row ["NS май"] = value;
        }
        public double NS06 
        {
            get => _row ["NS июнь"];
            set => _row ["NS июнь"] = value;
        }
        public double NS07 
        {
            get => _row ["NS июль"];
            set => _row ["NS июль"] = value;
        }
        public double NS08 
        {
            get => _row ["NS август"];
            set => _row ["NS август"] = value;
        }
        public double NS09 
        {
            get => _row ["NS сентябрь"];
            set => _row ["NS сентябрь"] = value;
        }
        public double NS10 
        {
            get => _row ["NS октябрь"];
            set => _row ["NS октябрь"] = value;
        }
        public double NS11 
        {
            get => _row ["NS ноябрь"];
            set => _row ["NS ноябрь"] = value;
        }
        public double NS12 
        {
            get => _row ["NS декабрь"];
            set => _row ["NS декабрь"] = value;
        }

        public double PromoPriceList01 
        {
            get => _row ["Promo Price-list январь"];
            set => _row ["Promo Price-list январь"] = value;
        }
        public double PromoPriceList02 
        {
            get => _row ["Promo Price-list февраль"];
            set => _row ["Promo Price-list февраль"] = value;
        }
        public double PromoPriceList03 
        {
            get => _row ["Promo Price-list март"];
            set => _row ["Promo Price-list март"] = value;
        }
        public double PromoPriceList04 
        {
            get => _row ["Promo Price-list апрель"];
            set => _row ["Promo Price-list апрель"] = value;
        }
        public double PromoPriceList05 
        {
            get => _row ["Promo Price-list май"];
            set => _row ["Promo Price-list май"] = value;
        }
        public double PromoPriceList06 
        {
            get => _row ["Promo Price-list июнь"];
            set => _row ["Promo Price-list июнь"] = value;
        }
        public double PromoPriceList07 
        {
            get => _row ["Promo Price-list июль"];
            set => _row ["Promo Price-list июль"] = value;
        }
        public double PromoPriceList08 
        {
            get => _row ["Promo Price-list август"];
            set => _row ["Promo Price-list август"] = value;
        }
        public double PromoPriceList09 
        {
            get => _row ["Promo Price-list сентябрь"];
            set => _row ["Promo Price-list сентябрь"] = value;
        }
        public double PromoPriceList10 
        {
            get => _row ["Promo Price-list октябрь"];
            set => _row ["Promo Price-list октябрь"] = value;
        }
        public double PromoPriceList11 
        {
            get => _row ["Promo Price-list ноябрь"];
            set => _row ["Promo Price-list ноябрь"] = value;
        }
        public double PromoPriceList12 
        {
            get => _row ["Promo Price-list декабрь"];
            set => _row ["Promo Price-list декабрь"] = value;
        }

        public double PromoSalesVolume01 
        {
            get => _row ["Promo Объем продаж январь"];
            set => _row ["Promo Объем продаж январь"] = value;
        }
        public double PromoSalesVolume02 
        {
            get => _row ["Promo Объем февраль"];
            set => _row ["Promo Объем февраль"] = value;
        }
        public double PromoSalesVolume03 
        {
            get => _row ["Promo Объем продаж март"];
            set => _row ["Promo Объем продаж март"] = value;
        }
        public double PromoSalesVolume04 
        {
            get => _row ["Promo Объем продаж апрель"];
            set => _row ["Promo Объем продаж апрель"] = value;
        }
        public double PromoSalesVolume05 
        {
            get => _row ["Promo Объем продаж май"];
            set => _row ["Promo Объем продаж май"] = value;
        }
        public double PromoSalesVolume06 
        {
            get => _row ["Promo Объем продаж июнь"];
            set => _row ["Promo Объем продаж июнь"] = value;
        }
        public double PromoSalesVolume07 
        {
            get => _row ["Promo Объем продаж июль"];
            set => _row ["Promo Объем продаж июль"] = value;
        }
        public double PromoSalesVolume08 
        {
            get => _row ["Promo Объем продаж август"];
            set => _row ["Promo Объем продаж август"] = value;
        }
        public double PromoSalesVolume09 
        {
            get => _row ["Promo Объем продаж сентябрь"];
            set => _row ["Promo Объем продаж сентябрь"] = value;
        }
        public double PromoSalesVolume10 
        {
            get => _row ["Promo Объем продаж октябрь"];
            set => _row ["Promo Объем продаж октябрь"] = value;
        }
        public double PromoSalesVolume11 
        {
            get => _row ["Promo Объем продаж ноябрь"];
            set => _row ["Promo Объем продаж ноябрь"] = value;
        }
        public double PromoSalesVolume12 
        {
            get => _row ["Promo Объем продаж декабрь"];
            set => _row ["Promo Объем продаж декабрь"] = value;
        }

        public double PromoGS01 
        {
            get => _row ["Promo GS январь"];
            set => _row ["Promo GS январь"] = value;
        }
        public double PromoGS02 
        {
            get => _row ["Promo GS февраль"];
            set => _row ["Promo GS февраль"] = value;
        }
        public double PromoGS03 
        {
            get => _row ["Promo GS март"];
            set => _row ["Promo GS март"] = value;
        }
        public double PromoGS04 
        {
            get => _row ["Promo GS апрель"];
            set => _row ["Promo GS апрель"] = value;
        }
        public double PromoGS05 
        {
            get => _row ["Promo GS май"];
            set => _row ["Promo GS май"] = value;
        }
        public double PromoGS06 
        {
            get => _row ["Promo GS июнь"];
            set => _row ["Promo GS июнь"] = value;
        }
        public double PromoGS07 
        {
            get => _row ["Promo GS июль"];
            set => _row ["Promo GS июль"] = value;
        }
        public double PromoGS08 
        {
            get => _row ["Promo GS август"];
            set => _row ["Promo GS август"] = value;
        }
        public double PromoGS09 
        {
            get => _row ["Promo GS сентябрь"];
            set => _row ["Promo GS сентябрь"] = value;
        }
        public double PromoGS10 
        {
            get => _row ["Promo GS октябрь"];
            set => _row ["Promo GS октябрь"] = value;
        }
        public double PromoGS11 
        {
            get => _row ["Promo GS ноябрь"];
            set => _row ["Promo GS ноябрь"] = value;
        }
        public double PromoGS12 
        {
            get => _row ["Promo GS декабрь"];
            set => _row ["Promo GS декабрь"] = value;
        }

        public double PromoNS01 
        {
            get => _row ["Promo NS январь"];
            set => _row ["Promo NS январь"] = value;
        }
        public double PromoNS02 
        {
            get => _row ["Promo NS февраль"];
            set => _row ["Promo NS февраль"] = value;
        }
        public double PromoNS03 
        {
            get => _row ["Promo NS март"];
            set => _row ["Promo NS март"] = value;
        }
        public double PromoNS04 
        {
            get => _row ["Promo NS апрель"];
            set => _row ["Promo NS апрель"] = value;
        }
        public double PromoNS05 
        {
            get => _row ["Promo NS май"];
            set => _row ["Promo NS май"] = value;
        }
        public double PromoNS06 
        {
            get => _row ["Promo NS июнь"];
            set => _row ["Promo NS июнь"] = value;
        }
        public double PromoNS07 
        {
            get => _row ["Promo NS июль"];
            set => _row ["Promo NS июль"] = value;
        }
        public double PromoNS08 
        {
            get => _row ["Promo NS август"];
            set => _row ["Promo NS август"] = value;
        }
        public double PromoNS09 
        {
            get => _row ["Promo NS сентябрь"];
            set => _row ["Promo NS сентябрь"] = value;
        }
        public double PromoNS10 
        {
            get => _row ["Promo NS октябрь"];
            set => _row ["Promo NS октябрь"] = value;
        }
        public double PromoNS11 
        {
            get => _row ["Promo NS ноябрь"];
            set => _row ["Promo NS ноябрь"] = value;
        }
        public double PromoNS12 
        {
            get => _row ["Promo NS декабрь"];
            set => _row ["Promo NS декабрь"] = value;
        }

        public double GPValue01 
        {
            get => _row ["GP Value январь"];
            set => _row ["GP Value январь"] = value;
        }
        public double GPValue02 
        {
            get => _row ["GP Value февраль"];
            set => _row ["GP Value февраль"] = value;
        }
        public double GPValue03 
        {
            get => _row ["GP Value март"];
            set => _row ["GP Value март"] = value;
        }
        public double GPValue04 
        {
            get => _row ["GP Value апрель"];
            set => _row ["GP Value апрель"] = value;
        }
        public double GPValue05 
        {
            get => _row ["GP Value май"];
            set => _row ["GP Value май"] = value;
        }
        public double GPValue06 
        {
            get => _row ["GP Value июнь"];
            set => _row ["GP Value июнь"] = value;
        }
        public double GPValue07 
        {
            get => _row ["GP Value июль"];
            set => _row ["GP Value июль"] = value;
        }
        public double GPValue08 
        {
            get => _row ["GP Value август"];
            set => _row ["GP Value август"] = value;
        }
        public double GPValue09 
        {
            get => _row ["GP Value сентябрь"];
            set => _row ["GP Value сентябрь"] = value;
        }
        public double GPValue10 
        {
            get => _row ["GP Value октябрь"];
            set => _row ["GP Value октябрь"] = value;
        }
        public double GPValue11 
        {
            get => _row ["GP Value ноябрь"];
            set => _row ["GP Value ноябрь"] = value;
        }
        public double GPValue12 
        {
            get => _row ["GP Value декабрь"];
            set => _row["GP Value декабрь"] = value;
        }

        public void SetPlan(PlanningNewYearItem planningNewYearSave)
        {
            ChannelType = planningNewYearSave.ChannelType;
            CustomerStatus = planningNewYearSave.CustomerStatus;
            PrognosisDate = planningNewYearSave.planningDate;
            Article = planningNewYearSave.Article;
            STKRub = planningNewYearSave.STKRub;
            //IRPEur = planningNewYearSave.IRPEur;
            //RRC = planningNewYearSave.RRC;
            //IRPIndex = planningNewYearSave.IRPIndex;
            //DIY = planningNewYearSave.DIY;
            //Price2Net = planningNewYearSave.Price2Net;
            //Price3Net = planningNewYearSave.Price3Net;
            //PriceTransfer = planningNewYearSave.PriceTransfer;

            //PriceList01 = planningNewYearSave.PriceList01;
            //PriceList02 = planningNewYearSave.PriceList02;
            //PriceList03 = planningNewYearSave.PriceList03;
            //PriceList04 = planningNewYearSave.PriceList04;
            //PriceList05 = planningNewYearSave.PriceList05;
            //PriceList06 = planningNewYearSave.PriceList06;
            //PriceList07 = planningNewYearSave.PriceList07;
            //PriceList08 = planningNewYearSave.PriceList08;
            //PriceList09 = planningNewYearSave.PriceList09;
            //PriceList10 = planningNewYearSave.PriceList10;
            //PriceList11 = planningNewYearSave.PriceList11;
            //PriceList12 = planningNewYearSave.PriceList12;

            //SalesVolume01 = planningNewYearSave.SalesVolume01;
            //SalesVolume02 = planningNewYearSave.SalesVolume02;
            //SalesVolume03 = planningNewYearSave.SalesVolume03;
            //SalesVolume04 = planningNewYearSave.SalesVolume04;
            //SalesVolume05 = planningNewYearSave.SalesVolume05;
            //SalesVolume06 = planningNewYearSave.SalesVolume06;
            //SalesVolume07 = planningNewYearSave.SalesVolume07;
            //SalesVolume08 = planningNewYearSave.SalesVolume08;
            //SalesVolume09 = planningNewYearSave.SalesVolume09;
            //SalesVolume10 = planningNewYearSave.SalesVolume10;
            //SalesVolume11 = planningNewYearSave.SalesVolume11;
            //SalesVolume12 = planningNewYearSave.SalesVolume12;

            GS01 = planningNewYearSave.GSPrognosis01;
            GS02 = planningNewYearSave.GSPrognosis02;
            GS03 = planningNewYearSave.GSPrognosis03;
            GS04 = planningNewYearSave.GSPrognosis04;
            GS05 = planningNewYearSave.GSPrognosis05;
            GS06 = planningNewYearSave.GSPrognosis06;
            GS07 = planningNewYearSave.GSPrognosis07;
            GS08 = planningNewYearSave.GSPrognosis08;
            GS09 = planningNewYearSave.GSPrognosis09;
            GS10 = planningNewYearSave.GSPrognosis10;
            GS11 = planningNewYearSave.GSPrognosis11;
            GS12 = planningNewYearSave.GSPrognosis12;

            NS01 = planningNewYearSave.NSPrognosis01;
            NS02 = planningNewYearSave.NSPrognosis02;
            NS03 = planningNewYearSave.NSPrognosis03;
            NS04 = planningNewYearSave.NSPrognosis04;
            NS05 = planningNewYearSave.NSPrognosis05;
            NS06 = planningNewYearSave.NSPrognosis06;
            NS07 = planningNewYearSave.NSPrognosis07;
            NS08 = planningNewYearSave.NSPrognosis08;
            NS09 = planningNewYearSave.NSPrognosis09;
            NS10 = planningNewYearSave.NSPrognosis10;
            NS11 = planningNewYearSave.NSPrognosis11;
            NS12 = planningNewYearSave.NSPrognosis12;

            //PromoPriceList01 = planningNewYearSave.PromoPriceList01;
            //PromoPriceList02 = planningNewYearSave.PromoPriceList02;
            //PromoPriceList03 = planningNewYearSave.PromoPriceList03;
            //PromoPriceList04 = planningNewYearSave.PromoPriceList04;
            //PromoPriceList05 = planningNewYearSave.PromoPriceList05;
            //PromoPriceList06 = planningNewYearSave.PromoPriceList06;
            //PromoPriceList07 = planningNewYearSave.PromoPriceList07;
            //PromoPriceList08 = planningNewYearSave.PromoPriceList08;
            //PromoPriceList09 = planningNewYearSave.PromoPriceList09;
            //PromoPriceList10 = planningNewYearSave.PromoPriceList10;
            //PromoPriceList11 = planningNewYearSave.PromoPriceList11;
            //PromoPriceList12 = planningNewYearSave.PromoPriceList12;

            //PromoSalesVolume01 = planningNewYearSave.PromoSalesVolume01;
            //PromoSalesVolume02 = planningNewYearSave.PromoSalesVolume02;
            //PromoSalesVolume03 = planningNewYearSave.PromoSalesVolume03;
            //PromoSalesVolume04 = planningNewYearSave.PromoSalesVolume04;
            //PromoSalesVolume05 = planningNewYearSave.PromoSalesVolume05;
            //PromoSalesVolume06 = planningNewYearSave.PromoSalesVolume06;
            //PromoSalesVolume07 = planningNewYearSave.PromoSalesVolume07;
            //PromoSalesVolume08 = planningNewYearSave.PromoSalesVolume08;
            //PromoSalesVolume09 = planningNewYearSave.PromoSalesVolume09;
            //PromoSalesVolume10 = planningNewYearSave.PromoSalesVolume10;
            //PromoSalesVolume11 = planningNewYearSave.PromoSalesVolume11;
            //PromoSalesVolume12 = planningNewYearSave.PromoSalesVolume12;

            PromoGS01 = planningNewYearSave.PromoGS01;
            PromoGS02 = planningNewYearSave.PromoGS02;
            PromoGS03 = planningNewYearSave.PromoGS03;
            PromoGS04 = planningNewYearSave.PromoGS04;
            PromoGS05 = planningNewYearSave.PromoGS05;
            PromoGS06 = planningNewYearSave.PromoGS06;
            PromoGS07 = planningNewYearSave.PromoGS07;
            PromoGS08 = planningNewYearSave.PromoGS08;
            PromoGS09 = planningNewYearSave.PromoGS09;
            PromoGS10 = planningNewYearSave.PromoGS10;
            PromoGS11 = planningNewYearSave.PromoGS11;
            PromoGS12 = planningNewYearSave.PromoGS12;

            PromoNS01 = planningNewYearSave.PromoNS01;
            PromoNS02 = planningNewYearSave.PromoNS02;
            PromoNS03 = planningNewYearSave.PromoNS03;
            PromoNS04 = planningNewYearSave.PromoNS04;
            PromoNS05 = planningNewYearSave.PromoNS05;
            PromoNS06 = planningNewYearSave.PromoNS06;
            PromoNS07 = planningNewYearSave.PromoNS07;
            PromoNS08 = planningNewYearSave.PromoNS08;
            PromoNS09 = planningNewYearSave.PromoNS09;
            PromoNS10 = planningNewYearSave.PromoNS10;
            PromoNS11 = planningNewYearSave.PromoNS11;
            PromoNS12 = planningNewYearSave.PromoNS12;

            GPValue01 = planningNewYearSave.GPValue01;
            GPValue02 = planningNewYearSave.GPValue02;
            GPValue03 = planningNewYearSave.GPValue03;
            GPValue04 = planningNewYearSave.GPValue04;
            GPValue05 = planningNewYearSave.GPValue05;
            GPValue06 = planningNewYearSave.GPValue06;
            GPValue07 = planningNewYearSave.GPValue07;
            GPValue08 = planningNewYearSave.GPValue08;
            GPValue09 = planningNewYearSave.GPValue09;
            GPValue10 = planningNewYearSave.GPValue10;
            GPValue11 = planningNewYearSave.GPValue11;
            GPValue12 = planningNewYearSave.GPValue12;
        }

    }
}
