using BPA.Model;
using Excel = Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using SettingsBPA = BPA.Properties.Settings;
using NM = BPA.NewModel;
using BPA.Forms;
using BPA.NewModel;

namespace BPA.Modules
{
    class FileDescision : FileBase
    {
        #region --- Columns ---
        public int CustomerColumn
        {
            get
            {
                if (_CustomerColumn == 0) _CustomerColumn = FindColumn("Customer");
                return _CustomerColumn;
            }
        }
        private int _CustomerColumn = 0;

        public int GardenaChannelColumn
        {
            get
            {
                if (_GardenaChannel == 0) _GardenaChannel = FindColumn("GardenaChannel");
                return _GardenaChannel;
            }
        }
        private int _GardenaChannel = 0;

        public int DateColumn
        {
            get
            {
                if (_DateColumn == 0) _DateColumn = FindColumn("Date");
                return _DateColumn;
            }
        }
        private int _DateColumn = 0;

        public int ArticleColumn
        {
            get
            {
                if (_ArticleColumn == 0) _ArticleColumn = FindColumn("Code");
                return _ArticleColumn;
            }
        }
        private int _ArticleColumn = 0;

        public int CampaignColumn
        {
            get
            {
                if (_CampaignColumn == 0) _CampaignColumn = FindColumn("Campaign");
                return _CampaignColumn;
            }
        }
        private int _CampaignColumn = 0;

        public int QuantitynColumn
        {
            get
            {
                if (_QuantitynColumn == 0) _QuantitynColumn = FindColumn("Quantity");
                return _QuantitynColumn;
            }
        }
        private int _QuantitynColumn = 0;

        public int PriceListColumn
        {
            get
            {
                if (_PriceListColumn == 0) _PriceListColumn = FindColumn("PricelistPriceTotal");
                return _PriceListColumn;
            }
        }
        private int _PriceListColumn = 0;
      
        public int BonusColumn
        {
            get
            {
                if (_BonusColumn == 0) _BonusColumn = FindColumn("Bonus");
                return _BonusColumn;
            }
        }
        private int _BonusColumn = 0;
        #endregion

        public FileDescision() 
        {
            BPASettings settings = new BPASettings();

            if (settings.GetDecisionPath(out string path))
            {
                FileAddress = path;
                FileHeaderRow = 1;
                FileSheetName = SettingsBPA.Default.SHEET_NAME_FILE_DECISION;
                
                IsOpen = true;
            }
            else
            {
                throw new ApplicationException("Загрузка отменена");
            }
        }

        public FileDescision(string filename)
        {
            if (!File.Exists(filename))
            {
                throw new FileNotFoundException($"Файл {filename} не найден");
            }
            FileAddress = filename;
        }

        public FileDescision(Excel.Workbook workbook)
        {
            //Workbook = workbook;
            FileAddress = Workbook.Path;
            IsOpen = true;
        }

        public List<ArticleQuantity> ArticleQuantities = new List<ArticleQuantity>();
        
        public List<NM.ClientItem.DataFromDescision> LoadClients()
        {
            List<NM.ClientItem.DataFromDescision> buffer = new List<NM.ClientItem.DataFromDescision>();

            if (CustomerColumn == 0 || GardenaChannelColumn == 0)
            {
                Close();
                throw new ApplicationException("Файл имеет неверный формат");
            }


            for (int rowIndex = 2; rowIndex < ArrRrows; rowIndex++)
            {
                if (IsCancel) return null;
                ActionStart?.Invoke($"Обрабатывается строка {rowIndex}");

                string customer = GetValueFromColumnStr(rowIndex, CustomerColumn);
                if(customer.Trim().Length > 0)
                {
                    string gardenaChannel = GetValueFromColumnStr(rowIndex, GardenaChannelColumn);

                    if (!buffer.Exists(x => x.Customer == customer)) buffer.Add(new NM.ClientItem.DataFromDescision() 
                    {
                        Customer = customer,
                        GardenaChannel = gardenaChannel
                    });
                    //if (!buffer.Exists(x => x.Customer == customer)) buffer.Add(new Client()
                    //{
                    //    Customer = customer,
                    //    GardenaChannel = gardenaChannel
                    //});
                }
                ActionDone?.Invoke(1);
            }

            if (buffer.Count == 0) throw new ApplicationException("Файл не содержит значемых данных");
            return buffer;
        }

        /// <summary>
        /// Получение списка 
        /// </summary>
        /// <param name="planning"></param>
        public void LoadForPlanning(PlanningNewYear planning) 
        {
            
            if (DateColumn == 0 || ArticleColumn == 0 || CampaignColumn == 0 || CustomerColumn== 0)
            {
                throw new ApplicationException("Файл имеет неверный формат");
            }            

            for (int rowIndex = 2; rowIndex < ArrRrows; rowIndex++)
            {
                if (IsCancel)
                    return;
                ActionStart?.Invoke($"Обрабатывается строка {rowIndex}");

                DateTime date = GetDateFromCell(rowIndex, DateColumn);
                string customer = GetValueFromColumnStr(rowIndex, CustomerColumn);

                //проверка на соответствие года и customer
                if (date.Year != planning.CurrentDate.Year || planning.Clients.Find(x => x.Customer == customer) == null)
                    continue;

                string article = GetValueFromColumnStr(rowIndex, ArticleColumn);
                string campaign = GetValueFromColumnStr(rowIndex, CampaignColumn);

                if (article != "")
                {
                    double quantity = GetValueFromColumnDbl(rowIndex, QuantitynColumn);
                    double priceList = GetValueFromColumnDbl(rowIndex, PriceListColumn);
                    double bonus = GetValueFromColumnDbl(rowIndex, BonusColumn);

                    ArticleQuantities.Add(new ArticleQuantity
                    {
                        Article = article,
                        Quantity = quantity,
                        Month = date.Month,
                        Campaign = campaign == "" ? "0" : campaign,
                        PriceList = priceList,
                        Bonus = bonus
                    });
                }
                ActionDone?.Invoke(1);
            }
        }

        /// <summary>
        /// Получение списка 
        /// </summary>
        /// <param name="planning"></param>
        public void LoadForPlanning(DateTime currentDate, List<ClientItem> client_list)
        {

            if (DateColumn == 0 || ArticleColumn == 0 || CampaignColumn == 0 || CustomerColumn == 0)
            {
                throw new ApplicationException("Файл имеет неверный формат");
            }

            for (int rowIndex = 2; rowIndex < ArrRrows; rowIndex++)
            {
                if (IsCancel)
                    return;
                ActionStart?.Invoke($"Обрабатывается строка {rowIndex}");

                DateTime date = GetDateFromCell(rowIndex, DateColumn);
                string customer = GetValueFromColumnStr(rowIndex, CustomerColumn);

                //проверка на соответствие года и customer
                if (date.Year != currentDate.Year || client_list.Find(x => x.Customer == customer) == null)
                    continue;

                string article = GetValueFromColumnStr(rowIndex, ArticleColumn);
                string campaign = GetValueFromColumnStr(rowIndex, CampaignColumn);

                if (article != "")
                {
                    double quantity = GetValueFromColumnDbl(rowIndex, QuantitynColumn);
                    double priceList = GetValueFromColumnDbl(rowIndex, PriceListColumn);
                    double bonus = GetValueFromColumnDbl(rowIndex, BonusColumn);

                    ArticleQuantities.Add(new ArticleQuantity
                    {
                        Article = article,
                        Quantity = quantity,
                        Month = date.Month,
                        Campaign = campaign == "" ? "0" : campaign,
                        PriceList = priceList,
                        Bonus = bonus
                    });
                }
                ActionDone?.Invoke(1);
            }
        }
    }
}
