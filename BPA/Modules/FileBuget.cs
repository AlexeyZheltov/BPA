using BPA.Model;
using Excel = Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using SettingsBPA = BPA.Properties.Settings;
using BPA.Forms;

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
        #endregion

        public FileBuget()
        {
            BPASettings settings = new BPASettings();

            if (settings.GetBudgetPath(out string path))
            {
                FileName = path;
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
            FileName = filename;
        }

        public FileBuget(Excel.Workbook workbook)
        {
            Workbook = workbook;
        }

        public List<ArticleQuantity> ArticleQuantities = new List<ArticleQuantity>();
        
        //получение списка артикулов и месяцов
        public void LoadForPlanning(PlanningNewYear planning)
        {
            ProcessBar processBar = null;

            SetFileData();

            if (DateColumn == 0 || ArticleColumn == 0 || CampaignColumn == 0)
            {
                throw new ApplicationException("Файл имеет неверный формат");
            }

            processBar = new ProcessBar($"Загрузка файла  { FileName } ", LastRow - FileHeaderRow);
            processBar.Show();
            ActionStart += processBar.TaskStart;
            ActionDone += processBar.TaskDone;
            processBar.CancelClick += Cancel;


            for (int rowIndex = 2; rowIndex < ArrRrows; rowIndex++)
            {
                if (IsCancel)
                    return;
                OnActionStart($"Обрабатывается строка {rowIndex}");

                DateTime date = GetDateFromCell(rowIndex, DateColumn);
                if (planning.Year != date.Year)
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
                OnActionDone(1);
            }
        }
    }
}
