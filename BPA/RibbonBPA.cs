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

                processBar = new ProcessBar("Загрузка", fileDescision.CountActions);
                processBar.CancelClick += fileDescision.Cancel;
                fileDescision.ActionStart += processBar.TaskStart;
                fileDescision.ActionDone += processBar.TaskDone;
                processBar.TaskStart("Загрузка из файла Decision");
                processBar.Show(new ExcelWindows(Globals.ThisWorkbook));


                List<Client> clientsFromDecision = fileDescision.LoadClients();

                processBar.Close();

                //Загрузить данные из листа клиентов
                List<Client> clients = new List<Client>();
                foreach (Excel.ListRow row in new Client().Table.ListRows)
                {
                    clients.Add(new Client(row));
                }

                //Получить разницу
                List<Client> newClients = clientsFromDecision.Except(clients, new Client.ComparerCustomer()).ToList();

                //Выгрузить разницу как новых клиентов
                newClients.ForEach(x => x.Save());
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
                    else if (ch == ']' & isMark) isMark = false;

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
            if(
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
            }

            //Загрузка списка артикулов, какие из них актуальные?
            List<Product> products = Product.GetProductsForDiscounts();
            products = products.FindAll(x => x.Status.ToLower() == "активный");

            //подключится к ценам
            List<RRC> rrcs = RRC.GetAllRRC();

            //в цикле менять метки на значения из цен, с заменой;
            foreach(Product product in products)
            {
                
            }

            //Еще раз прогнать по очистке формул

            //расчитать значения по функции руслана

            //подключится к листу вывода цен со скидками

            //Вывести

            //Client client = new Client();
            ////можно в клиенте написать метод возвращающий  
            ////клиетна по активной ячейке, или строке. что-то типа
            ////Clients client = new Clients().activeRow;
            ////string clientMag = client.Mag;
            //string clientMag = "ЛЕРУ";
            ////

            ////dataTime выбраная пользователем
            //DateTime date = new DateTime(2017, 08, 15);
            
            ////

            //FilePriceMT filePriceMT = new FilePriceMT();
            //filePriceMT.Load(clientMag, date);
            //List<FilePriceMT.Client> clientsPriceList = filePriceMT.clients;

            //ProcessBar processBar = new ProcessBar("Формирование прайс-листа", clientsPriceList.Count);
            //try
            //{
            //    FunctionsForExcel.SpeedOn();
            //    processBar.Show();
            //    Globals.ThisWorkbook.Activate();
                
            //    foreach (FilePriceMT.Client clientPrice in clientsPriceList)
            //    {
            //        if (processBar.IsCancel)
            //            break;
            //        processBar.TaskStart($"Обрабатывается клиент {clientPrice.Name}");


            //        //проверяем в discount
            //        //Discount discount = new Discount().GetDiscount();
            //        //discount.status

            //        //double price = clientPrice.Price
            //        double price = filePriceMT.GetPrice(clientPrice.Art);
            //        Debug.WriteLine(price);
            //        //здесь создаем новый лист
            //    }
            //}
            //catch (Exception ex)
            //{
            //    MessageBox.Show(ex.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            //}
            //finally
            //{
            //    FunctionsForExcel.SpeedOff();
            //    processBar.Close();
            //}
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
