using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;
using System;
using System.IO;
using System.Windows.Forms;
using System.Collections.Generic;
using SettingsBPA = BPA.Properties.Settings;
using BPA.Forms;

namespace BPA.Modules
{
    internal class FilePriceMT : FileBase
    {
        #region --- Columns ---
        public int CustomerColumn {
            get {
                if (_CustomerColumn == 0)
                    _CustomerColumn = FindColumn("Покупатель");
                return _CustomerColumn;
            }
        }
        private int _CustomerColumn = 0;
        public int SearchColumn
        {
            get
            {
                if (_SearchColumn == 0) _SearchColumn = FindColumn("Поиск");
                return _SearchColumn;
            }
        }
        private int _SearchColumn = 0;
        public int MainColumn {
            get
            {
                if (_MainColumn == 0) _MainColumn = FindColumn("Главный");
                return _MainColumn;
            }
        }
        private int _MainColumn = 0;
        public int ArticleColumn {
            get
            {
                if (_ArticleColumn == 0) _ArticleColumn = FindColumn("Артикул");
                return _ArticleColumn;
            }
        }
        private int _ArticleColumn = 0;
        public int NameColumn {
            get
            {
                if (_NameColumn == 0) _NameColumn = FindColumn("Название");
                return _NameColumn;
            }
        }
        private int _NameColumn = 0;
        public int PriceForClientColumn {
            get
            {
                if (_PriceForClientColumn == 0) _PriceForClientColumn = FindColumn("Цена_для_клиента");
                return _PriceForClientColumn;
            }
        }
        private int _PriceForClientColumn = 0;
        public int ValidFromDatColumn {
            get
            {
                if (_ValidFromDatColumn == 0) _ValidFromDatColumn = FindColumn("ValidFromDat");
                return _ValidFromDatColumn;
            }
        }
        private int _ValidFromDatColumn = 0;
        public int ValidToDatColumn {
            get
            {
                if (_ValidToDatColumn == 0) _ValidToDatColumn = FindColumn("ValidToDat");
                return _ValidToDatColumn;
            }
        }
        private int _ValidToDatColumn = 0;
        public int CustCodeColumn {
            get
            {
                if (_CustCodeColumn == 0) _CustCodeColumn = FindColumn("CustCode");
                return _CustCodeColumn;
            }
        }
        private int _CustCodeColumn = 0;
        public int PriceOfListingColumn {
            get
            {
                if (_PriceOfListingColumn == 0) _PriceOfListingColumn = FindColumn("Цена_листинга");
                return _PriceOfListingColumn;
            }
        }
        private int _PriceOfListingColumn = 0;
        public int PriceNewColumn {
            get
            {
                if (_PriceNewColumn == 0) _PriceNewColumn = FindColumn("Цена_новая");
                return _PriceNewColumn;
            }
        }
        private int _PriceNewColumn = 0;
        public int DateFromColumn {
            get
            {
                if (_DateFromColumn == 0) _DateFromColumn = FindColumn("От");
                return _DateFromColumn;
            }
        }
        private int _DateFromColumn = 0;
        public int DateToColumn
        {
            get
            {
                if (_DateToColumn == 0) _DateToColumn = FindColumn("До");
                return _DateToColumn;
            }
        }
        private int _DateToColumn = 0;
        public int MagColumn {
            get
            {
                if (_MagColumn == 0) _MagColumn = FindColumn("Маг");
                return _MagColumn;
            }
        }
        private int _MagColumn = 0;
        #endregion

        public FilePriceMT()
        {
            BPASettings settings = new BPASettings();

            if (settings.GetPriceListMT(out string path))
            {
                //FileName = path;
                FileAddress = path;
                FileSheetName = SettingsBPA.Default.SHEET_NAME_FILE_PRICELISTMT;
                FileHeaderRow = 1;

                IsOpen = true;
            }
            else
            {
                throw new ApplicationException("Загрузка отменена");
            }
        }

        public FilePriceMT(string filename)
        {
            if (!File.Exists(filename))
            {
                throw new FileNotFoundException($"Файл {filename} не найден");
            }
            FileName = filename;
        }

        public FilePriceMT(Workbook workbook)
        {
            //Workbook = workbook;
            FileAddress = Workbook.Path;
            IsOpen = true;
        }

        public List<Client> clients = new List<Client>();
        public struct Client
        {
            public string Name {
                get; set;
            }
            public double Price
            {
                get; set;
            }
            public string Art
            {
                get; set;
            }
        }
        
        /// <summary>
        /// получение магазина по дате
        /// </summary>
        /// <param name="mag"></param>
        /// <param name="date"></param>
        public void Load(DateTime date, string mag = "")
        {

            if (Workbook == null)
                return;

            clients.Clear();

            if (DateFromColumn == 0 || DateToColumn == 0 || MagColumn == 0)
            {
                Workbook.Close();
                throw new ApplicationException($"Файл {Path.GetFileName(FileName)} имеет ошибочный формат");
            }

            IsCancel = false;

            for (int rowIndex = 2; rowIndex < ArrRrows; rowIndex++) 
            {
                if (IsCancel)
                    return;

                ActionStart?.Invoke("Загрузка файла PriceListMT");

                string magVal;

                if (mag != "")
                {
                    magVal = GetValueFromColumnStr(rowIndex, MagColumn);
                    if (magVal != mag)
                    {
                        ActionD();
                        continue;
                    }
                }

                double priceNew = GetValueFromColumnDbl(rowIndex, PriceNewColumn);
                double priceOfListing= GetValueFromColumnDbl(rowIndex, PriceOfListingColumn);

                if (priceNew != 0)
                    AddClient(rowIndex, priceNew);
                else
                    AddClient(rowIndex, priceOfListing);

                ActionD();
            }

            void AddClient(int rowIndex, double price)
            {
                clients.Add(new Client
                {
                    Name = GetValueFromColumnStr(rowIndex, CustomerColumn),
                    Art = GetValueFromColumnStr(rowIndex, ArticleColumn),
                    Price = price
                });
            }
        }

        public double GetPrice(string Art)
        {
            return clients.Find(x => x.Art == Art).Price;
        }
    }
}
