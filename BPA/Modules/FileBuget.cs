using BPA.Model;
using Excel = Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using SettingsBPA = BPA.Properties.Settings;
using BPA.Forms;
using BPA.NewModel;

namespace BPA.Modules
{
    class FileBuget : FileBase
    {
        #region --- Columns ---
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
                if (_CampaignColumn == 0) _CampaignColumn = FindColumn("CampaignDisc");
                return _CampaignColumn;
            }
        }
        private int _CampaignColumn = 0;

        public int QuantitynColumn
        {
            get
            {
                if (_QuantitynColumn == 0) _QuantitynColumn = FindColumn("Qty");
                return _QuantitynColumn;
            }
        }
        private int _QuantitynColumn = 0;

        public int PriceListColumn
        {
            get
            {
                if (_PriceListColumn == 0) _PriceListColumn = FindColumn("PriceList");
                return _PriceListColumn;
            }
        }
        private int _PriceListColumn = 0;

        public int CustomerBugetColumn
        {
            get
            {
                if (_CustomerBugetColumn == 0) _CustomerBugetColumn = FindColumn("Customer");
                return _CustomerBugetColumn;
            }
        }
        private int _CustomerBugetColumn = 0;
        #endregion

        public FileBuget()
        {
            BPASettings settings = new BPASettings();

            if (settings.GetBudgetPath(out string path))
            {
                FileAddress = path;
                FileHeaderRow = 2;
                FileSheetName = SettingsBPA.Default.SHEET_NAME_FILE_BUGET;
                
                IsOpen = true;
            }
            else
            {
                throw new ApplicationException("Загрузка отменена");
            }
        }

        public FileBuget(string filename)
        {
            if (!File.Exists(filename))
            {
                throw new FileNotFoundException($"Файл {filename} не найден");
            }
            FileAddress = filename;
        }

        public FileBuget(Excel.Workbook workbook)
        {
            //Workbook = workbook;
            FileAddress = Workbook.Path;
            IsOpen = true;
        }

        public List<ArticleQuantity> ArticleQuantities = new List<ArticleQuantity>();
        
        //получение списка артикулов и месяцов
        public void LoadForPlanning(PlanningNewYear planning)
        {
            if (DateColumn == 0 || ArticleColumn == 0 || CampaignColumn == 0)
            {
                throw new ApplicationException("Файл имеет неверный формат");
            }

            for (int rowIndex = 2; rowIndex < ArrRrows; rowIndex++)
            {
                if (IsCancel)
                    return;

                ActionStart?.Invoke($"Обрабатывается строка {rowIndex}");

                DateTime date = GetDateFromCell(rowIndex, DateColumn);
                string customerBuget = GetValueFromColumnStr(rowIndex, CustomerBugetColumn); ;

                //проверка на соответствие года и customer
                if (date.Year != planning.CurrentDate.Year || planning.Clients.Find(x => x.CustomerBudget == customerBuget) == null)
                    continue;

                string article = GetValueFromColumnStr(rowIndex, ArticleColumn);
                string campaign = GetValueFromColumnStr(rowIndex, CampaignColumn);
                if (article != "")
                {
                    double quantity = GetValueFromColumnDbl(rowIndex, QuantitynColumn);
                    double priceList = GetValueFromColumnDbl(rowIndex, PriceListColumn);

                    ArticleQuantities.Add(new ArticleQuantity
                    {
                        Article = article,
                        Quantity = quantity,
                        Month = date.Month,
                        Campaign = campaign == "" ? "0": campaign,
                        PriceList = priceList
                    });
                }
                ActionD();
            }
        }

        public void LoadForPlanning(DateTime currentDate, List<ClientItem> client_list)
        {
            if (DateColumn == 0 || ArticleColumn == 0 || CampaignColumn == 0)
            {
                throw new ApplicationException("Файл имеет неверный формат");
            }

            for (int rowIndex = 2; rowIndex < ArrRrows; rowIndex++)
            {
                if (IsCancel)
                    return;

                ActionStart?.Invoke($"Обрабатывается строка {rowIndex}");

                DateTime date = GetDateFromCell(rowIndex, DateColumn);
                string customerBuget = GetValueFromColumnStr(rowIndex, CustomerBugetColumn); ;

                //проверка на соответствие года и customer
                if (date.Year != currentDate.Year || client_list.Find(x => x.CustomerBudget == customerBuget) == null)
                {
                    ActionD();
                    continue;
                }

                string article = GetValueFromColumnStr(rowIndex, ArticleColumn);
                string campaign = GetValueFromColumnStr(rowIndex, CampaignColumn);
                if (article != "")
                {
                    double quantity = GetValueFromColumnDbl(rowIndex, QuantitynColumn);
                    double priceList = GetValueFromColumnDbl(rowIndex, PriceListColumn);

                    ArticleQuantities.Add(new ArticleQuantity
                    {
                        Article = article,
                        Quantity = quantity,
                        Month = date.Month,
                        Campaign = campaign == "" ? "0" : campaign,
                        PriceList = priceList
                    });
                }
                ActionD();
            }
        }
    }
}
