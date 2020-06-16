using System;
using System.Collections.Generic;
using System.Windows.Forms;
using System.Linq;

using BPA.Forms;
using BPA.Model;
using BPA.Modules;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools.Ribbon;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;
using System.Diagnostics;
using System.Text;
using Microsoft.Office.Core;

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
            FileCalendar fileCalendar = new FileCalendar();
            ProcessBar processBar = new ProcessBar("Загрузка данных календаря", fileCalendar.CountActions);
            try
            {
                FunctionsForExcel.SpeedOn();
                Globals.ThisWorkbook.Activate();
                processBar.Show();
                fileCalendar.ActionStart += processBar.TaskStart;
                fileCalendar.ActionDone += processBar.TaskDone;
                processBar.CancelClick += fileCalendar.Cancel;
                fileCalendar.LoadCalendar();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                FunctionsForExcel.SpeedOff();
                processBar.Close();
            }
        }

        /// <summary>
        /// кнопка обновления
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void UpdateProducts_Click(object sender, RibbonControlEventArgs e)
        {
            List<ProductCalendar> calendars = new ProductCalendar().GetProductCalendars();
            ProcessBar processBar = new ProcessBar("Обновление продуктовых календарей", calendars.Count);
            try
            {
                FunctionsForExcel.SpeedOn();
                
                Product product = new Product();
                processBar.Show();
                Globals.ThisWorkbook.Activate();
                foreach (ProductCalendar calendar in calendars)
                {
                    if (processBar.IsCancel) break;
                    processBar.TaskStart($"Обрабатывается календарь {calendar.Name}");
                    processBar.AddSubBar("Обновление данных", product.Table.ListRows.Count);
                    calendar.ActionStart += processBar.SubBar.TaskStart;
                    calendar.ActionDone += processBar.SubBar.TaskDone;
                    processBar.SubBar.CancelClick += calendar.Cancel;
                    calendar.UpdateProducts();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                FunctionsForExcel.SpeedOff();
                processBar.SubBar.Close();
                processBar.Close();
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
            MessageBox.Show("Функционал в разработке", "BPA", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        /// <summary>
        /// Обновление продукта
        /// </summary>
        private void UpdateProduct_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                FunctionsForExcel.SpeedOn();
                Product product = new Product().GetPoductActive();
                ProductCalendar calendar = new ProductCalendar(product.Calendar);
                FileCalendar fileCalendar = new FileCalendar(calendar.Path);
                product.SetFromCalendar(fileCalendar.Workbook);
                fileCalendar.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                FunctionsForExcel.SpeedOff();
            }
        }

        private void UploadPrice_Click(object sender, RibbonControlEventArgs e)
        {
            MessageBox.Show("Функционал в разработке", "BPA", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void SavePrice_Click(object sender, RibbonControlEventArgs e)
        {
            MessageBox.Show("Функционал в разработке", "BPA", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void ClientsUpdate_Click(object sender, RibbonControlEventArgs e)
        {
            FunctionsForExcel.SpeedOn();
            ProcessBar processBar = null;
            FileDescision fileDescision = null;

            try
            {
                fileDescision = new FileDescision();
                if (fileDescision.IsNotOpen()) return;

                processBar = new ProcessBar("Обновление клиентов", fileDescision.CountActions);
                processBar.CancelClick += fileDescision.Cancel;
                fileDescision.ActionStart += processBar.TaskStart;
                fileDescision.ActionDone += processBar.TaskDone;
                processBar.Show(new ExcelWindows(Globals.ThisWorkbook));

                List<Client> clientsFromDecision = fileDescision.LoadClients();

                processBar.Close();

                //Загрузить данные из листа клиентов
                List<Client> clients = Client.GetAllClients();
                if (clients == null) return;

                //Получить разницу
                List<Client> newClients = clientsFromDecision.Except(clients, new Client.ComparerCustomer()).ToList();
                if(newClients.Count == 0)
                {
                    MessageBox.Show("В файле Decision не обнаружено новых клиентов", "BPA", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                //Выгрузить разницу как новых клиентов
                //newClients.ForEach(x => x.Save());
                bool isCancel = false;

                void Cancel() => isCancel = true;

                processBar = new ProcessBar("Обновление клиентов", fileDescision.CountActions);
                processBar.CancelClick += Cancel;
                processBar.Show(new ExcelWindows(Globals.ThisWorkbook));

                foreach (Client client in newClients)
                {
                    if (isCancel) return;
                    processBar.TaskStart($"Сохраняется клиент {client.Customer}");
                    client.Save();
                    processBar.TaskDone(1);
                }

                Client ClientForSort = newClients.First();
                ClientForSort.Sort("Id");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                if(!fileDescision?.IsNotOpen() ?? false) fileDescision.Close();
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
            ProcessBar processBar = null;
            try
            {
                //получить активного клиента, если нет, то на нет и суда нет
                Client currentClient = Client.GetCurrentClient();
                if (currentClient == null) return;

                //Запросить дату
                MSCalendar calendar = new MSCalendar();
                DateTime currentDate;
                if (calendar.ShowDialog(new ExcelWindows(Globals.ThisWorkbook)) == DialogResult.OK) currentDate = calendar.SelectedDate;
                else return;
                calendar.Close();

                //найти клиента в списке скидок
                List<Discount> discounts = Discount.GetAllDiscounts();
                discounts = discounts.FindAll(x => x.ChannelType == currentClient.ChannelType
                                                    && x.CustomerStatus == currentClient.CustomerStatus
                                                    && x.GetPeriodAsDateTime() != null
                                                    && x.GetPeriodAsDateTime() <= currentDate);

                discounts.Sort((x, y) =>
                {
                    if (x.GetPeriodAsDateTime() > y.GetPeriodAsDateTime()) return 1;
                    else if (x.GetPeriodAsDateTime() < y.GetPeriodAsDateTime()) return -1;
                    else return 0;
                });

                if(discounts.Count == 0)
                {
                    MessageBox.Show("Данному клиенту нет соответствий на листе \"Скидки\"", "BPA", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }
                Discount currentDiscount = discounts[0];
                discounts = null;

                //проверить формулы
                //Убрать пробелы и лишние знаки
                string FormulaNormalize(string value, bool RemoveMarks = false)
                {
                    //оставить только [метка], а вне ее только [1-9], +, - , *, /, (), %, =
                    StringBuilder builder = new StringBuilder();
                    bool isMark = false;

                    value = value.ToLower();
                    foreach (char ch in value.ToCharArray())
                    {
                        if (ch == '[' & !RemoveMarks) isMark = true;
                        else if (ch == ']' & isMark)
                        {
                            builder.Append(ch);
                            isMark = false;
                        }

                        if (!isMark)
                        {
                            if (Char.IsDigit(ch)) builder.Append(ch);
                            else
                            {
                                switch (ch)
                                {
                                    case '+':
                                    case '-':
                                    case '*':
                                    case '/':
                                    case '(':
                                    case ')':
                                    case '%':
                                    case '=':
                                        builder.Append(ch);
                                        break;
                                    case ',':
                                    case '.':
                                        builder.Append('.');
                                        break;
                                }
                            }

                        }
                        else builder.Append(ch);
                    }

                    string temp = System.Text.RegularExpressions.Regex.Replace(builder.ToString(), @"\s+", " ");
                    return builder.ToString();
                }

                currentDiscount.IrrigationEquipments = FormulaNormalize(currentDiscount.IrrigationEquipments);
                currentDiscount.Electricians = FormulaNormalize(currentDiscount.Electricians);
                currentDiscount.Lawnmowers = FormulaNormalize(currentDiscount.Lawnmowers);
                currentDiscount.Pumps = FormulaNormalize(currentDiscount.Pumps);
                currentDiscount.CuttingTools = FormulaNormalize(currentDiscount.CuttingTools);
                currentDiscount.WinterTools = FormulaNormalize(currentDiscount.WinterTools);

                //подгрузить PriceMT если неужно, подключится к РРЦ
                FilePriceMT filePriceMT = null;
                if (
                    currentDiscount.IrrigationEquipments.Contains("[pricelist mt]") ||
                    currentDiscount.Electricians.Contains("[pricelist mt]") ||
                    currentDiscount.Lawnmowers.Contains("[pricelist mt]") ||
                    currentDiscount.Pumps.Contains("[pricelist mt]") ||
                    currentDiscount.CuttingTools.Contains("[pricelist mt]") ||
                    currentDiscount.WinterTools.Contains("[pricelist mt]")
                    )
                {
                    //Загурзить файл price list MT
                    filePriceMT = new FilePriceMT();
                    filePriceMT.Load(currentClient.Mag, currentDate);
                    filePriceMT.Close();
                }

                //Загрузка списка артикулов, какие из них актуальные?
                List<Product> products = Product.GetProductsForDiscounts();
                products = products.FindAll(x => x.Status.ToLower() == "активный");

                //подключится к ценам
                List<RRC> rrcs = RRC.GetAllRRC();
                List<string> arts = (from rrc in rrcs
                                     select rrc.Article).Distinct().ToList();

                List<RRC> actualRRC = new List<RRC>();
                List<RRC> buffer = new List<RRC>();

                foreach (string art in arts)
                {
                    buffer = rrcs.FindAll(x => x.Article == art)
                                    .Where(x => x.GetDateAsDateTime() <= currentDate)
                                    .ToList();

                    buffer.Sort((x, y) =>
                    {
                        if (x.GetDateAsDateTime() > y.GetDateAsDateTime()) return 1;
                        else if (x.GetDateAsDateTime() < y.GetDateAsDateTime()) return -1;
                        else return 0;
                    });

                    if (buffer.Count == 0) continue;
                    actualRRC.Add(buffer[0]);
                }
                rrcs = null;
                arts = null;
                buffer = null;


                //в цикле менять метки на значения из цен, с заменой;
                List<FinalPriceList> priceList = new List<FinalPriceList>();

                foreach (Product product in products)
                {
                    //получить формулу
                    string formula = currentDiscount.GetFormulaByName(product.Category);

                    //Найти метку или метки. [Pricelist MT]  [DIY Pricelist] [РРЦ] и заменить
                    while (formula.Contains("[pricelist mt]"))
                        formula = formula.Replace("[pricelist mt]", filePriceMT.GetPrice(product.Article).ToString());

                    while (formula.Contains("[diy pricelist]"))
                        formula = formula.Replace("[diy pricelist]", actualRRC.Find(x => x.Article == product.Article).DIY);

                    while (formula.Contains("[ррц]"))
                        formula = formula.Replace("[ррц]", actualRRC.Find(x => x.Article == product.Article).RRCNDS);

                    if(Parsing.Calculation(formula) is double result)
                        priceList.Add(new FinalPriceList(product)
                        {
                        
                            RRC = result
                        });
                    else
                    {
                        MessageBox.Show($"В одной из формул для {currentClient.Customer} содержится ошибка", "BPA", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                }

                //Вывести
                priceList.ForEach(x => x.Save());
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                processBar?.Close();
                FunctionsForExcel.SpeedOff();
            }
        }

        private void GetAllPrices_Click(object sender, RibbonControlEventArgs e)
        {
            MessageBox.Show("Функционал в разработке", "BPA", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void PlanningAdd_Click(object sender, RibbonControlEventArgs e)
        {
            MessageBox.Show("Функционал в разработке", "BPA", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void GetPlanningData_Click(object sender, RibbonControlEventArgs e)
        {
            MessageBox.Show("Функционал в разработке", "BPA", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void FactUpdate_Click(object sender, RibbonControlEventArgs e)
        {
            MessageBox.Show("Функционал в разработке", "BPA", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void AddNewIRP_Click(object sender, RibbonControlEventArgs e)
        {
            MessageBox.Show("Функционал в разработке", "BPA", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void PlanningSave_Click(object sender, RibbonControlEventArgs e)
        {
            MessageBox.Show("Функционал в разработке", "BPA", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
    }
}
