using Microsoft.Office.Interop.Excel;
using System.Collections.Generic;

namespace BPA.Modules
{
    public class WorksheetsSettings
    {
        Workbook workbook;

        /// <summary>
        /// Cписок листов для сокрытия
        /// </summary>
        private readonly string[] SheetNames = new string[]
        {
            "Exclusives",
            "Channel type",
            "Customer status",
            "Продуктовые календари",
            "Структура ассортимента",
            "Статусы товаров",
            "STK",
            "Бюджетные курсы",
            "Эксклюзивность",
            "GardenChannel",
            "Сертификация",
            "Структура цен DIY",
            "Планирование НГ шаблон",
            "Прайс лист шаблон"
        };

        /// <summary>
        /// словарь со статусом скрытости листа
        /// </summary>
        Dictionary<string, bool> SheetsVisibleStatus
        {
            get
            {
                if (_sheetsVisibleStatus == null)
                {
                    try
                    {
                        _sheetsVisibleStatus = CheckStatus();
                    }
                    catch
                    {
                        _sheetsVisibleStatus = null;
                    }
                }
                return _sheetsVisibleStatus;
            }
            set
            {
                _sheetsVisibleStatus = value;
            }
        }
        Dictionary<string, bool> _sheetsVisibleStatus;

        /// <summary>
        /// Заполнение словаря статусов
        /// </summary>
        /// <returns></returns>
        private Dictionary<string, bool> CheckStatus()
        {
            Dictionary<string, bool> sheetStatusDict = new Dictionary<string, bool>();
            foreach (string sheetName in SheetNames)
            {
                if (!FunctionsForExcel.IsSheetExists(workbook, sheetName))
                    continue;

                Worksheet ws = workbook.Sheets[sheetName];
                sheetStatusDict.Add(sheetName, ws.Visible == XlSheetVisibility.xlSheetVisible);             
            }
            return sheetStatusDict;
        }

        public WorksheetsSettings(Workbook workbook)
        {
            this.workbook = workbook;
        }

        /// <summary>
        /// Проверят показан ли хоть один лист предназначенный для сокрытия. Если да, то скрывает все, иначе показывает все
        /// </summary>
        public void ShowUnshowSheets()
        {
            foreach (string sheetName in SheetsVisibleStatus.Keys)
            {
                if (SheetsVisibleStatus[sheetName])
                {
                    SetVisibleToSheets(XlSheetVisibility.xlSheetHidden);
                    return;
                }
            }
            SetVisibleToSheets(XlSheetVisibility.xlSheetVisible);
        }
        /// <summary>
        /// Скрывает или показывает лист в зависимости от переданного значения
        /// </summary>
        /// <param name="xlSheetVisibility"></param>
        private void SetVisibleToSheets(XlSheetVisibility sheetVisibilityStatus)
        {
            foreach (string sheetName in SheetsVisibleStatus.Keys)
            { 
                if (FunctionsForExcel.IsSheetExists(workbook, sheetName))
                {
                    if (sheetName.Contains("шаблон"))
                    {
                        workbook.Sheets[sheetName].Visible = XlSheetVisibility.xlSheetHidden;
                        continue;
                    }
                    workbook.Sheets[sheetName].Visible = sheetVisibilityStatus;
                }
            }
        }
    }
}

                
