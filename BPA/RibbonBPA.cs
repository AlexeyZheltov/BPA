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

                processBar = new ProcessBar("Обновление цен из справочника", products.Count);
                processBar.CancelClick += CancelLocal;
                processBar.Show();

                //Обновление списка продуктов
                foreach(FileCalendar.ProductFromCalendar productFromCalendar in fileCalendar.ProductsFromCalendar) 
                {
                    if (isCancel) break;

                    ProductItem product = products.Find(x=>x.Article == productFromCalendar.LocalIDGardena);

                    if (product == null) 
                        product = products.Add();

                    product.UpdateFromCalendar(productFromCalendar);
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

                DateTime date = products.DateOfPromotion();

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

                DateTime date = products.DateOfPromotion();
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
            FunctionsForExcel.SpeedOn();
            FileDescision fileDescision = null;
            ProcessBar processBar = null;

            try
            {
                NM.ClientTable clients = new NM.ClientTable();
                
                fileDescision = new FileDescision();
                if (!fileDescision.IsOpen) 
                    return;
                fileDescision.SetFileData();
                fileDescision.SetProcessBarForLoad(ref processBar);
                List<ClientFromDescision> clientsFromDecision = fileDescision.LoadClients();
                
                processBar.Close();
                if (fileDescision?.IsOpen ?? false) fileDescision.Close();

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
            ProcessBar processBar = null;
            FilePriceMT filePriceMT = null;

            bool isCancel = false;
            void Cancel() => isCancel = true;
            List<NM.ClientItem> priceClients = new List<NM.ClientItem>();
#if ENABLE_TRY
            try
            {
#endif
                NM.FinalPriceTable finalPrices = new NM.FinalPriceTable();
                NM.ExclusiveMagTable exclusives = new NM.ExclusiveMagTable();
                NM.ClientTable clients = new NM.ClientTable();
                NM.RRCTable rrcs = new NM.RRCTable();
                NM.ProductTable products = new NM.ProductTable();
                NM.DiscountTable discounts = new NM.DiscountTable();

                FunctionsForExcel.SpeedOn();
                clients.Load();
                if (All)
                {
                    //загрузить всех подопытных
                    
                    processBar = new ProcessBar($"Загрузка списка клиентов", clients.Count());
                    processBar.CancelClick += Cancel;
                    processBar.Show();

                    foreach(NM.ClientItem client in clients)
                    {
                        if (isCancel)
                        {
                            processBar.Close();
                            return;
                        }
                        processBar.TaskStart($"Загружаем {client.Id}");
                        priceClients.Add(client);
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
                    NM.ClientItem clientItem = clients.GetById(currents_id);
                    if (clientItem != null) priceClients.Add(clientItem); //need getByID
                    else return;
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
                finalPrices.Load();

                //сюда вынести загрузку общих данных
                List<NM.RRCItem> actualRRC = rrcs.GetActualPriceList(currentDate);
                if (actualRRC == null) return;
                
                foreach (NM.ClientItem currentClient in priceClients)
                {
                    NM.DiscountItem currentDiscount = discounts.GetCurrentDiscount(currentClient, currentDate);
                    if (currentDiscount == null) return;
                    //if (currentDiscount == null) continue;
                    
                    //подгрузить PriceMT если неужно, подключится к РРЦ                   
                    if (currentDiscount.NeedFilePriceMT() && (!filePriceMT?.IsOpen ?? true))
                    {
                        //Загурзить файл price list MT
                        processBar = null;
                        filePriceMT = new FilePriceMT();
                        if (!filePriceMT.IsOpen)
                            return;
                        filePriceMT.SetFileData();
                        filePriceMT.SetProcessBarForLoad(ref processBar); //зачем тут ref?
                        filePriceMT.Load(currentDate, currentClient.Mag);
                        processBar.Close();
                        
                        if (!All)
                        {
                            if (filePriceMT?.IsOpen ?? false) filePriceMT.Close();
                            //processBar.Close(); ///else not close???
                        }
                    }
                    
                    //Загрузка списка артикулов, какие из них актуальные?
                    List<NM.ProductItem> clients_products = products.GetProductForClient(currentClient, str_exclus);
                    if (clients_products.Count == 0) return;
                    /////Дописались до селе
                    ////в цикле менять метки на значения из цен, с заменой;
                    //List<FinalPriceList> priceList = new List<FinalPriceList>();
                    ////вместо него добовлять в finalPrices

                    processBar = new ProcessBar($"Создание прайс-листа для {currentClient.Customer}", products.Count);
                    processBar.CancelClick += Cancel;
                    processBar.Show();
                    foreach (NM.ProductItem product in clients_products)
                    {
                        if (isCancel) return;
                        //получить формулу
                        processBar.TaskStart($"Расчет цены для {product.Article}");
                        if (product.Category == "")
                        {
                            MessageBox.Show($"Для {product.Article} не указана категория", "BPA", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            return;
                        }
                        string formula = currentDiscount.GetFormulaByName(product.Category);
#if ENABLE_TRY
                        try
                        {
#endif
                            //Найти метку или метки. [Pricelist MT]  [DIY Pricelist] [РРЦ] и заменить
                            while (formula.Contains("[pricelist mt]"))
                                formula = formula.Replace("[pricelist mt]", filePriceMT.GetPrice(product.Article).ToString());

                            while (formula.Contains("[diy price list]"))
                                formula = formula.Replace("[diy price list]", actualRRC.Find(x => x.Article == product.Article)?.DIY.ToString() ?? "0");

                            while (formula.Contains("[ррц]"))
                                formula = formula.Replace("[ррц]", actualRRC.Find(x => x.Article == product.Article)?.RRCNDS.ToString() ?? "0");

                            if (Parsing.Calculation(formula) is double result)
                            {
                                NM.FinalPriceItem priceItem = finalPrices.Add();
                                priceItem.Fill(product);
                                priceItem.RRC = result;
                            }
                            else
                            {
                                MessageBox.Show($"В одной из формул для {currentClient.Customer} содержится ошибка", "BPA", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                return;
                            }
#if ENABLE_TRY
                        }
                        catch
                        {
                            MessageBox.Show($"{currentClient.Customer} не найден на листе { RRCTable.SHEET }", "BPA", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            return;
                        }
#endif
                        processBar.TaskDone(1);
                    }
                    processBar.Close();

                    ////Вывести
                    //processBar = new ProcessBar($"Создание прайс-листа для {currentClient.Customer}", products.Count);
                    //processBar.CancelClick += Cancel;
                    //foreach (FinalPriceList item in priceList)
                    //{
                    //    if (isCancel) return;
                    //    processBar.TaskStart($"Сохранение: {item.ArticleGardena}");
                    //    item.Save();
                    //    processBar.TaskDone(1);
                    //}
                    //processBar.TaskDone(1);
                }
                finalPrices.Save();
                Excel.Worksheet ws = Globals.ThisWorkbook.Sheets[finalPrices.SheetName];
                ws.Activate();
                MessageBox.Show("Создание прайс-листа завершено", "BPA", MessageBoxButtons.OK, MessageBoxIcon.Information);
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

                planningNewYears.Load();
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
                fileDescision.SetFileData();
                fileDescision.SetProcessBarForLoad(ref processBar);
                fileDescision.LoadForPlanning(planningNewYears.CurrentDate, planning_clients);
                processBar.Close();
                if (fileDescision?.IsOpen ?? false) fileDescision.Close();
                //

                //загружаем  Buget
                processBar = null;
                fileBuget = new FileBuget();
                if (!fileBuget.IsOpen)
                    return;
                fileBuget.SetFileData();
                fileBuget.SetProcessBarForLoad(ref processBar);
                fileBuget.LoadForPlanning(planningNewYears.CurrentDate, planning_clients);
                processBar.Close();
                if (fileBuget?.IsOpen ?? false) fileBuget.Close();
                //

                //загружаем  FilePriceListMT
                processBar = null;
                filePriceMT = new FilePriceMT();
                if (!filePriceMT.IsOpen)
                    return;
                filePriceMT.SetFileData();
                filePriceMT.SetProcessBarForLoad(ref processBar);
                filePriceMT.Load(planningNewYears.CurrentDate); //почему не передаем Mag??
                processBar.Close();
                if (filePriceMT?.IsOpen ?? false) filePriceMT.Close();
                //

                //PriceListForPlaning priceListForPlaning = new PriceListForPlaning(planningNewYearTmp);
                //priceListForPlaning.Load();
                //

                processBar = new ProcessBar("Обновление планирования", products.Count);
                bool isCancel = false;
                void CancelLocal() => isCancel = true;
                FunctionsForExcel.SpeedOn();
                processBar.CancelClick += CancelLocal;
                processBar.Show();


                //устанавливаем необходимые стобцы к удалению формул
                planningNewYears.DelFormulas();
                planningNewYears.ClearTable();

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
                NM.ClientTable clients = new ClientTable();
                //NM.ProductTable products = new ProductTable();
                NM.DiscountTable discounts = new DiscountTable();
                NM.RRCTable rrcs = new RRCTable();
                NM.ExclusiveMagTable exclusives = new NM.ExclusiveMagTable();

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

                exclusives.Load();
                List<string> str_exclus = (from exlus in exclusives
                                           select exlus.Name.ToLower()).ToList();
                exclusives = null;

                //products.Load();
                clients.Load();
                rrcs.Load();

                //List<NM.ProductItem> planning_products = products.GetProductForPlanning(planningNewYears, str_exclus);
                List<NM.ClientItem> planning_clients = clients.GetClientsForPlanning(planningNewYears.ChannelType, planningNewYears.CustomerStatus);
                List<NM.RRCItem> actualRRC = rrcs.GetActualPriceList(planningNewYears.CurrentDate);

                //получаем Desicion
                processBar = null;
                fileDescision = new FileDescision();
                fileDescision.SetFileData();
                fileDescision.SetProcessBarForLoad(ref processBar);
                if (!fileDescision.IsOpen)
                    return;
                fileDescision.LoadForPlanning(planningNewYears.CurrentDate, planning_clients);
                processBar.Close();
                if (fileDescision?.IsOpen ?? false) fileDescision.Close();
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

                //загружаем  FilePriceListMT
                processBar = null;
                filePriceMT = new FilePriceMT();
                if (!filePriceMT.IsOpen)
                    return;
                filePriceMT.SetFileData();
                filePriceMT.SetProcessBarForLoad(ref processBar);
                filePriceMT.Load(planningNewYears.CurrentDate); //почему не передаем Mag??
                processBar.Close();
                if (filePriceMT?.IsOpen ?? false) filePriceMT.Close();
                //

                //PriceListForPlaning priceListForPlaning = new PriceListForPlaning(planningNewYearTmp);
                //priceListForPlaning.Load();
                //

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
            ProcessBar processBar = null;

            Worksheet worksheet = Globals.ThisWorkbook.Application.ActiveSheet;

            NM.PlanningNewYearTable planningNewYears = new PlanningNewYearTable();
            NM.PlanTable plans = new PlanTable();

            if (!FunctionsForExcel.HasRange(worksheet, SettingsBPA.Default.PlannningNYIndicatorCellName) ||
                worksheet.Name == SettingsBPA.Default.SHEET_NAME_PLANNING_TEMPLATE)
            {
                MessageBox.Show("Перейдите на страницу планирования (или создайте её) и повторите попытку", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            try
            {
                if (!planningNewYears.HasData())
                {
                    MessageBox.Show($"Заполните { worksheet.Name } и повторите попытку", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                planningNewYears.SetTmpParams();
                planningNewYears.Load();
                plans.Load();

                if (planningNewYears == null || planningNewYears.Count < 1)
                {
                    MessageBox.Show($"Заполните { worksheet.Name } и повторите попытку", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                processBar = new ProcessBar("Обновление клиентов", planningNewYears.Count);
                bool isCancel = false;
                void CancelLocal() => isCancel = true;
                FunctionsForExcel.SpeedOn();
                processBar.CancelClick += CancelLocal;
                processBar.Show();

                foreach (NM.PlanningNewYearItem planningNewYear in planningNewYears)
                {
                    if (isCancel)
                        break;

                    processBar.TaskStart($"Обрабатывается артикул { planningNewYear.Article}");
                    NM.PlanItem plan = plans.Find(x => x.Article == planningNewYear.Article && x.PrognosisDate == planningNewYear.planningDate);
                    if (plan != null)
                        continue;
                    
                    plan = plans.Add();
                    plan.SetPlan(planningNewYear);

                    processBar.TaskDone(1);
                }
                Globals.ThisWorkbook.Sheets[plans.SheetName].Activate();

                plans.Save();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                FunctionsForExcel.SpeedOff();
                processBar?.Close();
            }
        }
    }
}
