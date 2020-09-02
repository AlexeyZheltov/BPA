#define ENABLE_TRY
//#undef ENABLE_TRY

using System;
using System.Collections.Generic;
using System.Windows.Forms;
using System.Linq;

using BPA.Forms;
using BPA.Modules;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools.Ribbon;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;
using System.Diagnostics;
using System.Text;
using Microsoft.Office.Core;
using SettingsBPA = BPA.Properties.Settings;
using NM = BPA.NewModel;
using System.Windows.Controls.Primitives;
using ClientFromDescision = BPA.NewModel.ClientItem.DataFromDescision;
using BPA.NewModel;
using BPA.Model;

namespace BPA
{
    public partial class RibbonBPA
    {
        private void RibbonBPA_Load(object sender, RibbonUIEventArgs e)
        {

        }

        /// <summary>
        /// кнопка загрузки
        /// </summary>
        private void AddNewCalendar_Click(object sender, RibbonControlEventArgs e)
        {
            FileCalendar fileCalendar = null;
            ProcessBar processBar = null;
            bool isCancel = false;
            void CancelLocal() => isCancel = true;
#if ENABLE_TRY
            try
            {
#endif
                FunctionsForExcel.SpeedOn();

                NM.ProductTable products = new NM.ProductTable();
                NM.ProductCalendarTable productCalendars = new NM.ProductCalendarTable();
                products.Load();
                productCalendars.Load();

                Globals.ThisWorkbook.Activate();
                
                //Загрузка календаря
                fileCalendar = new FileCalendar();
                if (!fileCalendar.IsOpen)
                    return;
                fileCalendar.SetFileData();
                fileCalendar.SetProcessBarForLoad(ref processBar);
                fileCalendar.LoadProductsFromCalendar();
                fileCalendar.Close();
                processBar?.Close();
                if (fileCalendar?.IsOpen ?? false) fileCalendar.Close();
                //

                if (fileCalendar.ProductsFromCalendar == null)
                {
                    MessageBox.Show("Значимых записей не найдено", "BPA", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                processBar = new ProcessBar("Обновление цен из справочника", fileCalendar.ProductsFromCalendar.Count);
                processBar.CancelClick += CancelLocal;
                processBar.Show();

                //Обновление списка продуктов
                foreach(FileCalendar.ProductFromCalendar productFromCalendar in fileCalendar.ProductsFromCalendar) 
                {
                    processBar.TaskStart($"Обрабатывается артикул {productFromCalendar.LocalIDGardena}");
                    if (isCancel) break;

                    ProductItem product = products.Find(x=>x.Article == productFromCalendar.LocalIDGardena);

                    if (product == null) 
                        product = products.Add();

                    product.UpdateFromCalendar(productFromCalendar);

                    processBar.TaskDone(1);
                }


                //Нужгно добавить маркировку выше
                //if (product != null)
                //{
                //    product.Mark("Article");
                //    product.Mark("PNS");
                //    product.Mark("Calendar");
                //}
                //else
                //{
                //    product.Mark("Calendar");
                //}

                //Обновление Справочника календарей
                ProductCalendarItem productCalendar = productCalendars.Find(x => x.Name == fileCalendar.FileName);

                if (productCalendar == null)
                    productCalendar = productCalendars.Add();

                productCalendar.UpdateFromCalendar(fileCalendar);


                products.Save();
                products.Sort("Продукт группа");
                productCalendars.Save();

                isCancel = true;
                MessageBox.Show("Загрузка календаря завершена", "BPA", MessageBoxButtons.OK, MessageBoxIcon.Information);
#if ENABLE_TRY
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
#endif
                FunctionsForExcel.SpeedOff();
                if (fileCalendar?.IsOpen ?? false) fileCalendar.Close();
                processBar?.Close();
#if ENABLE_TRY
            }
#endif
        }

        /// <summary>
        /// кнопка обновления
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void UpdateProducts_Click(object sender, RibbonControlEventArgs e)
        {
            if (MessageBox.Show
                (
                    $"Если вы выбирете действие \"Обновить календарь\", все ранее заполенные данные, будут обновлены в соответствии с продуктовыми календарями и заменены на содержащиеся там данные. \n\nОбновить календари?",
                    "Внимание",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Question,
                    MessageBoxDefaultButton.Button2
                )
                == DialogResult.No) return;

            FileCalendar fileCalendar = null;
            ProcessBar processBar = null;
            bool isCancel = false;
            void CancelLocal() => isCancel = true;

            //List<ProductCalendar> calendars = new ProductCalendar().GetProductCalendars();
            try
            {
                FunctionsForExcel.SpeedOn();

                NM.ProductTable products = new NM.ProductTable();
                NM.ProductCalendarTable productCalendars = new NM.ProductCalendarTable();
                products.Load();
                productCalendars.Load();

                processBar = new ProcessBar("Обновление продуктовых календарей", productCalendars.Count);
                processBar.CancelClick += CancelLocal;
                processBar.Show();

                Globals.ThisWorkbook.Activate();

                foreach (ProductCalendarItem productCalendar in productCalendars)
                {

                    if (isCancel) break;
                    //Если нет календаря, просто пропускаем?
                    processBar.TaskStart($"Обрабатывается календарь {productCalendar.Name}");
                    if (!File.Exists(productCalendar.Path))
                    {
                        processBar.TaskDone(1);
                        continue;
                    }
                    fileCalendar = new FileCalendar(productCalendar.Path);
                    if (!fileCalendar.IsOpen)
                        return;

                    ProcessBar pbForFileCalendar = null;
                    fileCalendar.SetFileData();
                    fileCalendar.SetProcessBarForLoad(ref pbForFileCalendar);
                    fileCalendar.LoadProductsFromCalendar();
                    fileCalendar.Close();
                    pbForFileCalendar?.Close();
                    if (fileCalendar?.IsOpen ?? false) fileCalendar.Close();

                    try
                    {
                        List<FileCalendar.ProductFromCalendar> productsFromCalendar = fileCalendar.ProductsFromCalendar;

                        foreach(ProductItem product in products)
                        {
                            //здесь добавить суббар
                            if (product.Calendar != productCalendar.Name) continue;

                            FileCalendar.ProductFromCalendar? productFromCalendar = productsFromCalendar.Find(x => x.LocalIDGardena == product.Article);

                            //проверить на нулл
                            if (productFromCalendar == null) continue;
                            product.UpdateFromCalendar((FileCalendar.ProductFromCalendar)productFromCalendar);
                        }
                    }
                    catch(FileNotFoundException)
                    {

                    }
                    processBar.TaskDone(1);
                }

                products.Save();
                MessageBox.Show("Обновление календарей завершено", "BPA", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                FunctionsForExcel.SpeedOff();
                if (processBar != null)
                {
                    processBar?.Close();
                }
            }
            
        }

        /// <summary>
        /// Кнопка информация
        /// </summary>
        private void About_Click(object sender, RibbonControlEventArgs e)
        {
            MessageBox.Show("Функционал в разработке", "BPA", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        /// <summary>
        /// Кнопка настроек
        /// </summary>
        private void Settings_Click(object sender, RibbonControlEventArgs e)
        {
            SettingsForm form = new SettingsForm();
            form.ShowDialog(new ExcelWindows(Globals.ThisWorkbook));
            //try
            //{
            //    FunctionsForExcel.SpeedOn();

            //    //FunctionsForExcel.HideShowSettingsSheets();
            //    WorksheetsSettings WS = new WorksheetsSettings();
            //    WS.ShowUnshowSheets();
            //}
            //catch (Exception ex)
            //{
            //    MessageBox.Show(ex.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            //}
            //finally
            //{
            //    FunctionsForExcel.SpeedOff();
            //}
        }

        /// <summary>
        /// Обновление продукта
        /// </summary>
        private void UpdateProduct_Click(object sender, RibbonControlEventArgs e)
        {
            FileCalendar fileCalendar = null;
            ProcessBar processBar = null;

            try
            {
                FunctionsForExcel.SpeedOn();

                NM.ProductTable products = new NM.ProductTable();
                NM.ProductCalendarTable productCalendars = new NM.ProductCalendarTable();
                products.Load();
                productCalendars.Load();

                Excel.Range activeCell = Globals.ThisWorkbook.Application.ActiveCell;
                int activeId = products.GetId(activeCell.Row);
                if (activeId == 0)
                {
                    throw new ApplicationException("Выберите товар");
                }

                ProductItem product = products.Find(x => x.Id == activeId);
                ProductCalendarItem calendar = productCalendars.Find(x=>x.Name == product.Calendar);
                
                if (calendar == null)
                {
                    throw new ApplicationException($"Файл { product.Calendar } не найден") ;
                }
                fileCalendar = new FileCalendar(calendar.Path);
                fileCalendar.SetFileData();
                fileCalendar.LoadProductsFromCalendar();
                //fileCalendar.GetArticle(product.Article);
                fileCalendar.SetProcessBarForLoad(ref processBar);
                fileCalendar.Close();
                processBar?.Close();

                if (fileCalendar != null)
                {
                    FileCalendar.ProductFromCalendar? productFromCalendar = fileCalendar.ProductsFromCalendar.Find(x=>x.LocalIDGardena == product.Article);
                        product.UpdateFromCalendar((FileCalendar.ProductFromCalendar)productFromCalendar);
                }
                products.Save();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                if (fileCalendar?.IsOpen ?? false) fileCalendar.Close();
                if (processBar != null)
                    processBar?.Close();

                FunctionsForExcel.SpeedOff();
            }
        }

        private void UploadPrice_Click(object sender, RibbonControlEventArgs e)
        {
            ProcessBar processBar = null;
            bool isCancel = false;
            void CancelLocal() => isCancel = true;
#if ENABLE_TRY
            try
            {
#endif
                FunctionsForExcel.SpeedOn();

                NM.ProductTable products = new NM.ProductTable();
                NM.RRCTable rrcs = new NM.RRCTable();

                products.Load();
                rrcs.Load();

                processBar = new ProcessBar("Обновление цен из справочника", products.Count);
                processBar.CancelClick += CancelLocal;
                processBar.Show();
                Globals.ThisWorkbook.Activate();

                DateTime date = products.DateOfPromotion;

                foreach(NM.ProductItem product in products)
                {
                    if (isCancel) break;
                    processBar.TaskStart($"Обрабатывается артикул {product.Article}");

                    var quere = (from rrc in rrcs
                                 where rrc.Article == product.Article && rrc.Date <= date
                                 orderby rrc.Date descending
                                 select rrc).ToList();

                    product.UpdatePriceFromRRC(quere.Count > 0 ? quere.First() : null);
                    processBar.TaskDone(1);
                }

                products.Save();
#if ENABLE_TRY
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
#endif
                FunctionsForExcel.SpeedOff();
                processBar?.Close();
#if ENABLE_TRY
            }
#endif
        }

        private void SavePrice_Click(object sender, RibbonControlEventArgs e)
        {
            ProcessBar processBar = null;
            bool isCancel = false;
            void CancelLocal() => isCancel = true;
#if ENABLE_TRY
            try
            {
#endif
                FunctionsForExcel.SpeedOn();

                NM.ProductTable products = new NM.ProductTable();
                NM.RRCTable rrcs = new NM.RRCTable();

                products.Load();
                rrcs.Load();

                processBar = new ProcessBar("Обновление цен из справочника", products.Count);
                processBar.CancelClick += CancelLocal;
                processBar.Show();
                Globals.ThisWorkbook.Activate();

                DateTime date = products.DateOfPromotion;
                double budget_cource = products.BudgetCourse();

                foreach(NM.ProductItem product in products)
                {
                    if (isCancel) break;
                    processBar.TaskStart($"Обрабатывается артикул {product.Article}");

                    NM.RRCItem rrc = rrcs.Find(x => x.Article == product.Article && x.Date == date);

                    if (rrc == null) rrc = rrcs.Add();

                    rrc.UpdateRRCFromProduct(product, date, budget_cource);

                    processBar.TaskDone(1);
                }

                rrcs.Save();
#if ENABLE_TRY
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
#endif
                FunctionsForExcel.SpeedOff();
                processBar?.Close();
#if ENABLE_TRY
            }
#endif
        }

        private void ClientsUpdate_Click(object sender, RibbonControlEventArgs e)
        {
            string[] ColumnsForLoadFromDescision = new string[] { "Code", "Date", "Campaign", "Customer", "Quantity", "PricelistPriceTotal", "Bonus" , "GardenaChannel" };

            FunctionsForExcel.SpeedOn();
            FileDescision fileDescision = null;
            ProcessBar processBar = null;

            try
            {
                NM.ClientTable clients = new NM.ClientTable();
                
                fileDescision = new FileDescision();
                if (!fileDescision.IsOpen) 
                    return;
                fileDescision.SetFileData(ColumnsForLoadFromDescision);
                fileDescision.SetProcessBarForLoad(ref processBar);
                if (fileDescision?.IsOpen ?? false) fileDescision.Close();
                List<ClientFromDescision> clientsFromDecision = fileDescision.LoadClients();
                fileDescision.ClearData();

                processBar.Close();

                //Загрузить данные из листа клиентов
                if (clients.Load() == 0) return;

                //Получить разницу

                //List<ClientFromDescision> newClients = new List<ClientFromDescision>();

                var newClients = (from c in clientsFromDecision
                                  where !clients.Contains(c)
                                  select c).ToList();

                if(newClients.Count == 0)
                {
                    MessageBox.Show("В файле Decision не обнаружено новых клиентов", "BPA", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                //Выгрузить разницу как новых клиентов
                //newClients.ForEach(x => x.Save());
                bool isCancel = false;

                void Cancel() => isCancel = true;

                processBar = new ProcessBar("Обновление клиентов", newClients.Count); ///???
                processBar.CancelClick += Cancel;
                processBar.Show(new ExcelWindows(Globals.ThisWorkbook));

                foreach (var item in newClients)
                {
                    if (isCancel) return;
                    processBar.TaskStart($"Сохраняется клиент {item.Customer}");

                    NM.ClientItem client = clients.Add();
                    client.Customer = item.Customer;
                    client.GardenaChannel = item.GardenaChannel;

                    processBar.TaskDone(1);
                }

                clients.Save();

                NM.ClientTable.SortExcelTable("№");
                Excel.Worksheet ws = Globals.ThisWorkbook.Sheets[NM.ClientTable.SHEET];
                ws.Activate();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                if(fileDescision?.IsOpen ?? false) fileDescision.Close();
                processBar?.Close();
                FunctionsForExcel.SpeedOff();
            }
        }

        /// <summary>
        /// прайс-лист клиента
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void GetClientPrice_Click(object sender, RibbonControlEventArgs e)
        {
            CreatePrice(false);
        }

        private void GetAllPrices_Click(object sender, RibbonControlEventArgs e)
        {
            CreatePrice(true);
        }

        private void CreatePrice(bool All = false)
        {
            string MessageDoesentLoadDiscount = "";
            string MessageDoesentLoadProduct = "";
            ProcessBar processBar = null;
            FilePriceMT filePriceMT = null;

            bool isCancel = false;
            void Cancel() => isCancel = true;
#if ENABLE_TRY
            try
            {
#endif
                NM.ExclusiveMagTable exclusives = new NM.ExclusiveMagTable();
                NM.ClientTable clients = new NM.ClientTable();
                NM.RRCTable rrcs = new NM.RRCTable();
                NM.ProductTable products = new NM.ProductTable();
                NM.DiscountTable discounts = new NM.DiscountTable();
                List<ClientCategory> priceClients = new List<ClientCategory>();

                FunctionsForExcel.SpeedOn();
                clients.Load();

                if (All)
                {
                    List <ClientCategory> clientCategories  = new ClientCategory().GetCategoryListFromClients(clients);
                    //загрузить всех клиентов                    
                    processBar = new ProcessBar($"Загрузка списка клиентов", clientCategories.Count());
                    processBar.CancelClick += Cancel;
                    processBar.Show();

                    foreach(ClientCategory clientCategory in clientCategories)
                    {
                        if (isCancel)
                        {
                            processBar.Close();
                            return;
                        }
                        string ClientsCategoryName = $"{clientCategory.CustomerStatus}_{clientCategory.ChannelType}";
                        processBar.TaskStart($"Загружаем { ClientsCategoryName }");

                        priceClients.Add(clientCategory);
                        processBar.TaskDone(1);
                    }

                    processBar?.Close();
                }
                else
                {
                    //выбрать того на кого указал перст божий
                    //получить активного клиента, если нет, то на нет и суда нет
                    int currents_id = clients.GetCurrentClientID();
                    if (currents_id == 0)
                    {
                        MessageBox.Show("Выберите клиента на листе \"Клиенты\"", "BPA", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        return;
                    }
                    NM.ClientItem client = clients.GetById(currents_id);
                    if (client == null) return;


                    ClientCategory clientCategory = new ClientCategory().GetCategoryFromClient(client);
                    priceClients.Add(clientCategory); //need getByID
                }

                clients = null;

                //Запросить дату
                MSCalendar calendar = new MSCalendar();
                DateTime currentDate;
                if (calendar.ShowDialog(new ExcelWindows(Globals.ThisWorkbook)) == DialogResult.OK) currentDate = calendar.SelectedDate;
                else return;
                calendar.Close();

                exclusives.Load();
                List<string> str_exclus = (from e in exclusives
                                           select e.Name.ToLower()).ToList();
                exclusives = null;

                //pb загрузки если будет лаг
                rrcs.Load();
                discounts.Load();
                products.Load();

                //сюда вынести загрузку общих данных
                List<NM.RRCItem> actualRRC = rrcs.GetActualPriceList(currentDate);
                if (actualRRC == null) 
                    return;
                    //continue;
                
                string priceTemplateSheetName = SettingsBPA.Default.SHEET_NAME_PRICELIST_TEMPLATE;
                ThisWorkbook workbook = Globals.ThisWorkbook;
                FunctionsForExcel.ShowSheet(priceTemplateSheetName);

                string newSheetName = priceTemplateSheetName.Replace("шаблон", "").Trim();

                //Загружаем массив данных из PriceListMT
                filePriceMT = new FilePriceMT();
                if (!filePriceMT.IsOpen)
                    return;
                filePriceMT.SetFileData();
                if (filePriceMT?.IsOpen ?? false) filePriceMT.Close();

                foreach (ClientCategory currentClient in priceClients)
                {
                    //Подготовка новой книги
                    string ClientsCategoryName = $"{currentClient.CustomerStatus}_{currentClient.ChannelType}";
                    Excel.Worksheet newSheet = FunctionsForExcel.CreateSheetСopyNewWB(workbook.Sheets[priceTemplateSheetName],
                                                                                        $"{ newSheetName }_{ ClientsCategoryName }");

                    NM.FinalPriceTable finalPrices = new NM.FinalPriceTable(newSheet);
                    finalPrices.Load();
                    finalPrices.DelFirstRow();
                    finalPrices.SetParams(currentClient.CustomerStatus,
                                            currentClient.ChannelType,
                                            currentDate);


                    NM.DiscountItem currentDiscount = discounts.GetCurrentDiscount(currentClient.ChannelType, currentClient.CustomerStatus, currentDate);
                    if (currentDiscount == null)
                    {
                        MessageDoesentLoadDiscount = $"{ MessageDoesentLoadDiscount }{ ClientsCategoryName },\n";
                        continue;
                    }

                    processBar = null;
                    filePriceMT.SetProcessBarForLoad(ref processBar); //зачем тут ref?
                    filePriceMT.Load(currentDate, currentClient.Mag);
                    processBar.Close();

                    PriceListForPlanningNM priceListModule = new PriceListForPlanningNM(filePriceMT, currentDiscount);

                    //Загрузка списка артикулов, какие из них актуальные?
                    List<NM.ProductItem> clients_products = products.GetProductForClient(currentClient.CustomerStatus, currentClient.ChannelType, str_exclus);
                    if (clients_products == null)
                    {
                        MessageDoesentLoadProduct = $"{ MessageDoesentLoadDiscount }{ ClientsCategoryName },\n";
                        continue;
                    }

                    processBar = new ProcessBar($"Создание прайс-листа для { ClientsCategoryName }", clients_products.Count);
                    processBar.CancelClick += Cancel;
                    processBar.Show();
                    foreach (NM.ProductItem product in clients_products)
                    {
                        if (isCancel) return;
                        //получить формулу
                        processBar.TaskStart($"Расчет цены для {product.Article}");

                        //получение прайслистцены 
                        priceListModule.SetProduct(product);
                        if (!priceListModule.FormulaChecked) return;
                        double priceListPrice = priceListModule.GetPrice(actualRRC);

                        NM.FinalPriceItem priceItem = finalPrices.Add();
                        priceItem.Fill(product);
                        priceItem.RRC = priceListPrice;

                       processBar.TaskDone(1);
                    }

                    finalPrices.Save();
                    processBar.Close();
                }

                FunctionsForExcel.HideSheet(priceTemplateSheetName);

                workbook.Activate();

                if (MessageDoesentLoadDiscount.Length > 0)
                {
                    MessageDoesentLoadDiscount = $"{ MessageDoesentLoadDiscount.Substring(0, MessageDoesentLoadDiscount.Length - 2) }\nна листе скидок не найдены";
                    MessageBox.Show(MessageDoesentLoadDiscount, "BPA", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }

                if (MessageDoesentLoadProduct.Length > 0)
                {
                    MessageDoesentLoadProduct = $"Для { MessageDoesentLoadProduct.Substring(0, MessageDoesentLoadProduct.Length - 2) }\nтовары не найдены";
                    MessageBox.Show(MessageDoesentLoadProduct, "BPA", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }

                MessageBox.Show("Создание прайс-листов завершено", "BPA", MessageBoxButtons.OK, MessageBoxIcon.Information);
#if ENABLE_TRY
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
#endif
                if (filePriceMT?.IsOpen ?? false) filePriceMT.Close();
                processBar?.Close();
                FunctionsForExcel.SpeedOff();
#if ENABLE_TRY
            }
#endif
        }

        /// <summary>
        /// Кнопка создать планирование
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void PlanningAdd_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                FunctionsForExcel.SpeedOn();

                string planningTemplateSheetName = SettingsBPA.Default.SHEET_NAME_PLANNING_TEMPLATE;

                ThisWorkbook workbook = Globals.ThisWorkbook;
                FunctionsForExcel.ShowSheet(planningTemplateSheetName);

                string newSheetName = planningTemplateSheetName.Replace("шаблон", "").Trim();
                Excel.Worksheet newSheet = FunctionsForExcel.CreateSheetCopy(workbook.Sheets[planningTemplateSheetName], newSheetName);
                newSheet.Activate();

                FunctionsForExcel.HideSheet(planningTemplateSheetName);

            } catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                FunctionsForExcel.SpeedOff();
            }
        }

        /// <summary>
        /// Формирование планирования
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void GetPlanningData_Click(object sender, RibbonControlEventArgs e)
        {
            string[] ColumnsForLoadFromDescision = new string[] { "Code", "Date", "Campaign", "Customer", "Quantity", "PricelistPriceTotal", "Bonus" };

            ProcessBar processBar = null;
            FileDescision fileDescision = null;
            FileBuget fileBuget = null;
            FilePriceMT filePriceMT = null;

            Worksheet worksheet = Globals.ThisWorkbook.Application.ActiveSheet;
            if (!FunctionsForExcel.HasRange(worksheet, SettingsBPA.Default.PlannningNYIndicatorCellName) ||
                worksheet.Name == SettingsBPA.Default.SHEET_NAME_PLANNING_TEMPLATE)
            {
                MessageBox.Show("Перейдите на страницу планирования (или создайте её) и повторите попытку", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            try
            {
                NM.DiscountTable discounts = new NM.DiscountTable();
                NM.ProductTable products = new NM.ProductTable();
                NM.ClientTable clients = new NM.ClientTable();
                NM.RRCTable rrcs = new NM.RRCTable();
                NM.STKTable stks = new NM.STKTable();
                NM.ExclusiveProductTable exclusives = new NM.ExclusiveProductTable();

                //получаем заполненые данне
                NM.PlanningNewYearTable planningNewYears = new NM.PlanningNewYearTable(worksheet.Name);

                //устанавливаем необходимые стобцы к удалению формул
                planningNewYears.ClearTable();
                planningNewYears.DelFormulas();
                planningNewYears.Load();
                planningNewYears.DelFirstRow();

                planningNewYears.SetTmpParams();
                if (!planningNewYears.TmpSeted)
                {
                    MessageBox.Show("Создайте копию листа планирования нового года", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                discounts.Load();
                NM.DiscountItem discount = discounts.GetDiscountForPlanning(planningNewYears);                
                if (discount != null) planningNewYears.MaximumBonus = discount.MaximumBonus;

                exclusives.Load();
                List<string> str_exclus = (from exlus in exclusives
                                           select exlus.Name.ToLower()).ToList();
                exclusives = null;

                products.Load();
                clients.Load();
                rrcs.Load();
                stks.Load();

                List<NM.ProductItem> planning_products = products.GetProductForPlanning(planningNewYears, str_exclus);
                List<NM.ClientItem> planning_clients = clients.GetClientsForPlanning(planningNewYears.ChannelType, planningNewYears.CustomerStatus);
                List<NM.RRCItem> actualRRC = rrcs.GetActualPriceList(planningNewYears.CurrentDate);
                List<NM.RRCItem> planRRC = rrcs.GetActualPriceList(planningNewYears.planningDate);
                List<NM.STKItem> actualSTK = stks.GetActualPriceList(planningNewYears.CurrentDate);
                List<NM.STKItem> planSTK = stks.GetActualPriceList(planningNewYears.planningDate);

                //получаем Desicion
                //processBar = null;
                fileDescision = new FileDescision();
                if (!fileDescision.IsOpen)
                    return;
                fileDescision.SetFileData(ColumnsForLoadFromDescision);
                if (fileDescision?.IsOpen ?? false) fileDescision.Close();
                fileDescision.SetProcessBarForLoad(ref processBar);
                fileDescision.LoadForPlanning(planningNewYears.CurrentDate, planning_clients);
                fileDescision.ClearData();
                processBar.Close();
                //

                //загружаем  Buget
                processBar = null;
                fileBuget = new FileBuget();
                if (!fileBuget.IsOpen)
                    return;
                fileBuget.SetFileData();
                if (fileBuget?.IsOpen ?? false) fileBuget.Close();
                fileBuget.SetProcessBarForLoad(ref processBar);
                fileBuget.LoadForPlanning(planningNewYears.CurrentDate, planning_clients);
                fileBuget.ClearData();
                processBar.Close();
                //

                //загружаем  FilePriceListMT
                processBar = null;
                filePriceMT = new FilePriceMT();
                if (!filePriceMT.IsOpen)
                    return;
                filePriceMT.SetFileData();
                if (filePriceMT?.IsOpen ?? false) filePriceMT.Close();
                filePriceMT.SetProcessBarForLoad(ref processBar);
                filePriceMT.Load(planningNewYears.CurrentDate); //почему не передаем Mag??
                filePriceMT.ClearData();
                processBar.Close();
                //

                //PriceListForPlaning priceListForPlaning = new PriceListForPlaning(planningNewYearTmp);
                //priceListForPlaning.Load();
                //
                PriceListForPlanningNM priceListModule = new PriceListForPlanningNM(filePriceMT, discount);

                processBar = new ProcessBar("Обновление планирования", products.Count);
                bool isCancel = false;
                void CancelLocal() => isCancel = true;
                FunctionsForExcel.SpeedOn();
                processBar.CancelClick += CancelLocal;
                processBar.Show();


                // заполняем планнингньюер 
                foreach (NM.ProductItem product in planning_products)
                {
                    if (isCancel)
                        break;
                    processBar.TaskStart($"Обрабатывается артикул {product.Article}");

                    NM.PlanningNewYearItem planning = planningNewYears.Add();
                    planning.SetParamsToItem(planningNewYears);
                    planning.SetProduct(product);

                    //уточнить отбор цены по дате
                    NM.RRCItem RRCPlan = planRRC.Find(x => x.Article == product.Article);
                    NM.RRCItem RRCCurrent = actualRRC.Find(x => x.Article == product.Article);
                    planning.SetRRC(RRCPlan, RRCCurrent);

                    //уточнить отбор цены по дате
                    NM.STKItem STKPlan = planSTK.Find(x => x.Article == product.Article);
                    NM.STKItem STKPCurrent = actualSTK.Find(x => x.Article == product.Article);
                    planning.SetSTK(STKPlan, STKPCurrent);

                    planning.SetValuesPrognosis(fileDescision.ArticleQuantities, fileBuget.ArticleQuantities);
                    //planning.DIYPriceList = priceListForPlaning.GetPrice(product.Article);

                    //priceList
                    priceListModule.SetProduct(product);
                    if (!priceListModule.FormulaChecked) return;

                    planning.PriceListCurrentn= priceListModule.GetPrice(actualRRC);
                    planning.PriceListPalan = priceListModule.GetPrice(planRRC);
                    //

                    processBar.TaskDone(1);
                }
                planningNewYears.Save();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                FunctionsForExcel.SpeedOff();
                processBar?.Close();
                if (fileBuget?.IsOpen ?? false) fileBuget.Close();
                if (fileDescision?.IsOpen ?? false) fileDescision.Close();
            }
            
            //MessageBox.Show("Функционал в разработке", "BPA", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        /// <summary> 
        ///Планирование/обновить факт
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void FactUpdate_Click(object sender, RibbonControlEventArgs e)
        {
            string[] ColumnsForLoadFromDescision = new string[] { "Code", "Date", "Campaign", "Customer", "Quantity", "PricelistPriceTotal", "Bonus" };

            ProcessBar processBar = null;
            FileDescision fileDescision = null;
            FileBuget fileBuget = null;
            //FilePriceMT filePriceMT = null;

            Worksheet worksheet = Globals.ThisWorkbook.Application.ActiveSheet;
            if (!FunctionsForExcel.HasRange(worksheet, SettingsBPA.Default.PlannningNYIndicatorCellName) ||
                worksheet.Name == SettingsBPA.Default.SHEET_NAME_PLANNING_TEMPLATE)
            {
                MessageBox.Show("Перейдите на страницу планирования (или создайте её) и повторите попытку", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            try
            {
                NM.ClientTable clients = new ClientTable();
                NM.DiscountTable discounts = new DiscountTable();

                NM.PlanningNewYearTable planningNewYears = new PlanningNewYearTable(worksheet.Name);
                planningNewYears.Load();
                planningNewYears.SetTmpParams();

                if (!planningNewYears.HasData())
                {
                    MessageBox.Show($"Заполните { worksheet.Name } и повторите попытку", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                if (planningNewYears.Count < 1)
                {
                    MessageBox.Show($"Заполните { worksheet.Name } и повторите попытку", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                discounts.Load();
                NM.DiscountItem discount = discounts.GetDiscountForPlanning(planningNewYears);
                if (discount != null) planningNewYears.MaximumBonus = discount.MaximumBonus;

                clients.Load();

                List<NM.ClientItem> planning_clients = clients.GetClientsForPlanning(planningNewYears.ChannelType, planningNewYears.CustomerStatus);

                //получаем Desicion
                processBar = null;
                fileDescision = new FileDescision();
                fileDescision.SetFileData(ColumnsForLoadFromDescision);
                fileDescision.SetProcessBarForLoad(ref processBar);
                if (!fileDescision.IsOpen)
                    return;
                if (fileDescision?.IsOpen ?? false) fileDescision.Close();
                fileDescision.LoadForPlanning(planningNewYears.CurrentDate, planning_clients);
                fileDescision.ClearData();
                processBar.Close();
                //

                //получаем Buget
                processBar = null;
                fileBuget= new FileBuget();
                fileBuget.SetFileData();
                fileBuget.SetProcessBarForLoad(ref processBar);
                if (!fileBuget.IsOpen)
                    return;
                fileBuget.LoadForPlanning(planningNewYears.CurrentDate, planning_clients);
                processBar.Close();
                if (fileBuget?.IsOpen ?? false) fileBuget.Close();
                //

                ////загружаем  FilePriceListMT
                //processBar = null;
                //filePriceMT = new FilePriceMT();
                //if (!filePriceMT.IsOpen)
                //    return;
                //filePriceMT.SetFileData();
                //filePriceMT.SetProcessBarForLoad(ref processBar);
                //filePriceMT.Load(planningNewYears.CurrentDate); //почему не передаем Mag??
                //processBar.Close();
                //if (filePriceMT?.IsOpen ?? false) filePriceMT.Close();

                processBar = new ProcessBar("Обновление клиентов", planningNewYears.Count);
                bool isCancel = false;
                void CancelLocal() => isCancel = true;
                FunctionsForExcel.SpeedOn();
                processBar.CancelClick += CancelLocal;
                processBar.Show();

                foreach(NM.PlanningNewYearItem planning in planningNewYears)
                {
                    if (isCancel)
                        break;

                    string article = planning.Article;
                    processBar.TaskStart($"Обрабатывается артикул {article}");

                    planning.SetValuesPrognosis(fileDescision.ArticleQuantities, fileBuget.ArticleQuantities);


                    processBar.TaskDone(1);
                }
                planningNewYears.Save();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                FunctionsForExcel.SpeedOff();
                processBar?.Close();
                if (fileBuget?.IsOpen ?? false) fileBuget.Close();
                if (fileDescision?.IsOpen ?? false) fileDescision.Close();
            }
        }

        private void AddNewIRP_Click(object sender, RibbonControlEventArgs e)
        {
            //MessageBox.Show("Функционал в разработке", "BPA", MessageBoxButtons.OK, MessageBoxIcon.Information);
            BPASettings settings = new BPASettings();
            if(settings.GetProductCalendarPath(out string path, true))
            {
                MessageBox.Show(path);
            }
            else
            {
                MessageBox.Show("Kernel Panic");
            }

        }

        /// <summary>
        /// Планирование / сохранить планирование
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void PlanningSave_Click(object sender, RibbonControlEventArgs e)
        {
            Excel.Workbook TWB = Globals.ThisWorkbook.InnerObject;
            Excel.Workbook workbook = null;
            ProcessBar processBar = null;
            //WaitForm waitForm = null;
            const string SHEET_NAME_PLAN = "Планирование";

            try
            {
                FunctionsForExcel.SpeedOn();

                if (!FunctionsForExcel.IsSheetExists(SHEET_NAME_PLAN))
                    throw new ApplicationException($"Лист \"{ SHEET_NAME_PLAN }\" отсутствует");
                Excel.Worksheet planWS = TWB.Sheets[SHEET_NAME_PLAN];

                if (planWS.ListObjects.Count < 1)
                    throw new ApplicationException($"Таблица на листе \"{ SHEET_NAME_PLAN }\" отсутствует");
                Excel.ListObject tableAllPlan = planWS.ListObjects[1];


                //узнаем последний номер и не добавляем одну строку
                double num = 0;
                if (tableAllPlan.ListRows.Count > 0)
                {
                    Excel.ListColumn firstColumn = tableAllPlan.ListColumns[1];

                    object[,] firstColumnArray = firstColumn.Range.Value;

                    int r = 0;
                    for (r = firstColumnArray.GetLength(0); r >= 1; r--)
                    {
                        object val = firstColumnArray[r, 1];
                        if (val != null)
                            if (Double.TryParse(val.ToString(), out num))
                                if (num != 0)
                                    break;
                    }

                    if (num == 0)
                        foreach (Excel.ListRow listRow in tableAllPlan.ListRows)
                            listRow.Delete(); 
                    else
                        tableAllPlan.Resize(tableAllPlan.Range.Resize[r, tableAllPlan.ListColumns.Count]);
                }
                //


                //перебор файлов
                OpenFileDialog openFileDialog = new OpenFileDialog();
                openFileDialog.Multiselect = true;
                openFileDialog.Filter= "Excel files (*.xls*)|*.xls*";
                if (openFileDialog.ShowDialog() == DialogResult.Cancel)
                    throw new ApplicationException($"Файлы не выбраны");
                string[] fileNames = openFileDialog.FileNames;

                processBar = new ProcessBar("Обновление сводного планирования", fileNames.Length);
                bool isCancel = false;
                void CancelLocal() => isCancel = true;
                processBar.CancelClick += CancelLocal;
                processBar.Show(new ExcelWindows(Globals.ThisWorkbook));

                foreach (string fileName in fileNames)
                {
                    if (isCancel)
                        return;

                    string fn = Path.GetFileName(fileName);
                    processBar.TaskStart($"Обрабатывается книга { fn }");

                    workbook = Globals.ThisWorkbook.Application.Workbooks.Open(fileName);

                    Excel.Worksheet worksheet = workbook.Sheets[1];
                    if (!FunctionsForExcel.HasRange(worksheet, SettingsBPA.Default.PlannningNYIndicatorCellName)
                    || worksheet.ListObjects.Count < 1
                    || worksheet.ListObjects[1].ListRows.Count < 1)
                    { 
                        workbook.Close(false);
                        workbook = null;
                        continue;
                    }

                    //выгрузка
                    PlanningNewYearTable tablePlanLoaded = new PlanningNewYearTable(workbook, worksheet.Name);
                    tablePlanLoaded.SetTmpParams();
                    object[,] planningData = tablePlanLoaded.GetDataForPlanning();
                    int dataRows = planningData.GetLength(0);
                    int dataColumns = planningData.GetLength(1);
                    workbook.Close(false);
                    workbook = null;
                    //

                    //вставка
                    if (tableAllPlan.ListRows.Count == 0) 
                    {
                        tableAllPlan.ListRows.Add();
                        tableAllPlan.ListRows[2].Delete(); //тупой ексель
                    } else
                        tableAllPlan.ListRows.Add();

                    Excel.ListRow listRow = tableAllPlan.ListRows[tableAllPlan.ListRows.Count];

                    int tmpArrColumns = 4;
                    Array rv = Array.CreateInstance(typeof(object), new int[] { dataRows, tmpArrColumns }, new int[] { 1, 1 });
                    object[,] buffer = rv as object[,];

                    for (int r = 1; r <= dataRows; r++)
                    {
                        buffer[r, 1] = ++num;
                        buffer[r, 2] = tablePlanLoaded.ChannelType;
                        buffer[r, 3] = tablePlanLoaded.planningDate.Year;
                        buffer[r, 4] = tablePlanLoaded.CustomerStatus;
                    }

                    //waitForm = new WaitForm();
                    //waitForm.Show();
                    //System.Windows.Forms.Application.DoEvents();

                    Excel.Range tmpCell;
                    tmpCell = listRow.Range[1];
                    tmpCell.Resize[dataRows, tmpArrColumns].Value = buffer;

                    tmpCell = listRow.Range[tmpArrColumns + 1];
                    tmpCell.Resize[dataRows, dataColumns].Value = planningData;
                    //waitForm.Close();

                    processBar.TaskDone(1);
                }
                planWS.Activate();
                MessageBox.Show("Обработка книг завершена", "BPA", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }

            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            } 
            finally
            {
                workbook?.Close(false);
                processBar?.Close();
                //waitForm?.Close();
                FunctionsForExcel.SpeedOff();
            }

            //ProcessBar processBar = null;

            //Worksheet worksheet = Globals.ThisWorkbook.Application.ActiveSheet;

            //NM.PlanningNewYearTable planningNewYears = new PlanningNewYearTable(worksheet.Name);
            //NM.tableAllPlan plans = new tableAllPlan();

            //if (!FunctionsForExcel.HasRange(worksheet, SettingsBPA.Default.PlannningNYIndicatorCellName) ||
            //    worksheet.Name == SettingsBPA.Default.SHEET_NAME_PLANNING_TEMPLATE)
            //{
            //    MessageBox.Show("Перейдите на страницу планирования (или создайте её) и повторите попытку", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            //    return;
            //}

            //try
            //{
            //    planningNewYears.Load();
            //    plans.Load();

            //    if (!planningNewYears.HasData())
            //    {
            //        MessageBox.Show($"Заполните { worksheet.Name } и повторите попытку", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            //        return;
            //    }

            //    planningNewYears.SetTmpParams();

            //    if (planningNewYears == null || planningNewYears.Count < 1)
            //    {
            //        MessageBox.Show($"Заполните { worksheet.Name } и повторите попытку", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            //        return;
            //    }

            //    processBar = new ProcessBar("Обновление клиентов", planningNewYears.Count);
            //    bool isCancel = false;
            //    void CancelLocal() => isCancel = true;
            //    FunctionsForExcel.SpeedOn();
            //    processBar.CancelClick += CancelLocal;
            //    processBar.Show();

            //    foreach (NM.PlanningNewYearItem planningNewYear in planningNewYears)
            //    {
            //        if (isCancel)
            //            break;

            //        processBar.TaskStart($"Обрабатывается артикул { planningNewYear.Article}");
            //        NM.PlanItem plan = plans.Find(x => x.Article == planningNewYear.Article && x.PrognosisDate == planningNewYear.planningDate);
            //        if (plan != null)
            //            continue;
                    
            //        plan = plans.Add();
            //        plan.SetPlan(planningNewYear);

            //        processBar.TaskDone(1);
            //    }
            //    Globals.ThisWorkbook.Sheets[plans.SheetName].Activate();

            //    plans.Save();
            //}
            //catch (Exception ex)
            //{
            //    MessageBox.Show(ex.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            //}
            //finally
            //{
            //    FunctionsForExcel.SpeedOff();
            //    processBar?.Close();
            //}
        }
    }
}
