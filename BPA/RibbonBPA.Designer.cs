namespace BPA
{
    partial class RibbonBPA : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Обязательная переменная конструктора.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public RibbonBPA()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

        /// <summary> 
        /// Освободить все используемые ресурсы.
        /// </summary>
        /// <param name="disposing">истинно, если управляемый ресурс должен быть удален; иначе ложно.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Код, автоматически созданный конструктором компонентов

        /// <summary>
        /// Требуемый метод для поддержки конструктора — не изменяйте 
        /// содержимое этого метода с помощью редактора кода.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(RibbonBPA));
            this.tabBPA = this.Factory.CreateRibbonTab();
            this.grpProducts = this.Factory.CreateRibbonGroup();
            this.grpPrice = this.Factory.CreateRibbonGroup();
            this.group2 = this.Factory.CreateRibbonGroup();
            this.group3 = this.Factory.CreateRibbonGroup();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.btnAddNewCalendar = this.Factory.CreateRibbonButton();
            this.btnUpdateProducts = this.Factory.CreateRibbonButton();
            this.btnUpdateProduct = this.Factory.CreateRibbonButton();
            this.btnUploadPrice = this.Factory.CreateRibbonButton();
            this.btnSavePrice = this.Factory.CreateRibbonButton();
            this.btnClientsUpdate = this.Factory.CreateRibbonButton();
            this.btnGetClientPrice = this.Factory.CreateRibbonButton();
            this.btnGetAllPrices = this.Factory.CreateRibbonButton();
            this.btnPlanningAdd = this.Factory.CreateRibbonButton();
            this.btnGetPlanningData = this.Factory.CreateRibbonButton();
            this.btnFactUpdate = this.Factory.CreateRibbonButton();
            this.btnPlanningSave = this.Factory.CreateRibbonButton();
            this.btnSettings = this.Factory.CreateRibbonButton();
            this.btnSheetsHide = this.Factory.CreateRibbonButton();
            this.btnInfo = this.Factory.CreateRibbonButton();
            this.tabBPA.SuspendLayout();
            this.grpProducts.SuspendLayout();
            this.grpPrice.SuspendLayout();
            this.group2.SuspendLayout();
            this.group3.SuspendLayout();
            this.group1.SuspendLayout();
            this.SuspendLayout();
            // 
            // tabBPA
            // 
            this.tabBPA.Groups.Add(this.grpProducts);
            this.tabBPA.Groups.Add(this.grpPrice);
            this.tabBPA.Groups.Add(this.group2);
            this.tabBPA.Groups.Add(this.group3);
            this.tabBPA.Groups.Add(this.group1);
            this.tabBPA.Label = "BPA";
            this.tabBPA.Name = "tabBPA";
            this.tabBPA.Position = this.Factory.RibbonPosition.BeforeOfficeId("TabHome");
            // 
            // grpProducts
            // 
            this.grpProducts.Items.Add(this.btnAddNewCalendar);
            this.grpProducts.Items.Add(this.btnUpdateProducts);
            this.grpProducts.Items.Add(this.btnUpdateProduct);
            this.grpProducts.Label = "Справочник товаров";
            this.grpProducts.Name = "grpProducts";
            // 
            // grpPrice
            // 
            this.grpPrice.Items.Add(this.btnUploadPrice);
            this.grpPrice.Items.Add(this.btnSavePrice);
            this.grpPrice.Label = "Работа с ценами";
            this.grpPrice.Name = "grpPrice";
            // 
            // group2
            // 
            this.group2.Items.Add(this.btnClientsUpdate);
            this.group2.Items.Add(this.btnGetClientPrice);
            this.group2.Items.Add(this.btnGetAllPrices);
            this.group2.Label = "Работа с клиентами";
            this.group2.Name = "group2";
            // 
            // group3
            // 
            this.group3.Items.Add(this.btnPlanningAdd);
            this.group3.Items.Add(this.btnGetPlanningData);
            this.group3.Items.Add(this.btnFactUpdate);
            this.group3.Items.Add(this.btnPlanningSave);
            this.group3.Label = "Планирование";
            this.group3.Name = "group3";
            // 
            // group1
            // 
            this.group1.Items.Add(this.btnSettings);
            this.group1.Items.Add(this.btnSheetsHide);
            this.group1.Items.Add(this.btnInfo);
            this.group1.Label = "Настройки";
            this.group1.Name = "group1";
            // 
            // btnAddNewCalendar
            // 
            this.btnAddNewCalendar.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnAddNewCalendar.Image = ((System.Drawing.Image)(resources.GetObject("btnAddNewCalendar.Image")));
            this.btnAddNewCalendar.Label = "Загрузить новый";
            this.btnAddNewCalendar.Name = "btnAddNewCalendar";
            this.btnAddNewCalendar.ScreenTip = "Добавление товаров в справочник товаров из нового календаря";
            this.btnAddNewCalendar.ShowImage = true;
            this.btnAddNewCalendar.SuperTip = "Новые позиции, отмеченные в календаре как to be sold in Russia, импортируются в с" +
    "правочник товаров в конец списка соответствующей группы товаров, а также выделяю" +
    "тся заливкой для информативности.";
            this.btnAddNewCalendar.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.AddNewCalendar_Click);
            // 
            // btnUpdateProducts
            // 
            this.btnUpdateProducts.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnUpdateProducts.Image = ((System.Drawing.Image)(resources.GetObject("btnUpdateProducts.Image")));
            this.btnUpdateProducts.Label = "Обновить календари";
            this.btnUpdateProducts.Name = "btnUpdateProducts";
            this.btnUpdateProducts.ScreenTip = "Обновление справочника товаров из календарей";
            this.btnUpdateProducts.ShowImage = true;
            this.btnUpdateProducts.SuperTip = "Программа сопоставляет данные справочника товаров и выбранного продуктового кален" +
    "даря.  Данные по существующим позициям обновляются в справочнике товаров.";
            this.btnUpdateProducts.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.UpdateProducts_Click);
            // 
            // btnUpdateProduct
            // 
            this.btnUpdateProduct.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnUpdateProduct.Image = ((System.Drawing.Image)(resources.GetObject("btnUpdateProduct.Image")));
            this.btnUpdateProduct.Label = "Обновить артикул";
            this.btnUpdateProduct.Name = "btnUpdateProduct";
            this.btnUpdateProduct.ScreenTip = "Обновление выбранного артикула из соответствующего календаря";
            this.btnUpdateProduct.ShowImage = true;
            this.btnUpdateProduct.SuperTip = "Обновляется товар выделенный в справочнике товаров. Данные подгружаются из соотве" +
    "тствующего товару продуктового календаря.";
            this.btnUpdateProduct.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.UpdateProduct_Click);
            // 
            // btnUploadPrice
            // 
            this.btnUploadPrice.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnUploadPrice.Image = ((System.Drawing.Image)(resources.GetObject("btnUploadPrice.Image")));
            this.btnUploadPrice.Label = "Загрузить текущие цены";
            this.btnUploadPrice.Name = "btnUploadPrice";
            this.btnUploadPrice.ScreenTip = "Загрузка текущих цен из справочника РРЦ в справочник товаров";
            this.btnUploadPrice.ShowImage = true;
            this.btnUploadPrice.SuperTip = "Цены в справочнике товаров обновляются в соответствии с данными из справочника РР" +
    "Ц. Данные берутся актуальные на дату повышения.";
            this.btnUploadPrice.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.UploadPrice_Click);
            // 
            // btnSavePrice
            // 
            this.btnSavePrice.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnSavePrice.Image = ((System.Drawing.Image)(resources.GetObject("btnSavePrice.Image")));
            this.btnSavePrice.Label = "Сохранить новые РРЦ";
            this.btnSavePrice.Name = "btnSavePrice";
            this.btnSavePrice.ScreenTip = "Сохранение новых цен из справочника товаров в справочник РРЦ";
            this.btnSavePrice.ShowImage = true;
            this.btnSavePrice.SuperTip = "Текущие цены из справочника товаров переносятся в справочник РРЦ с указанием даты" +
    " повышения. В случае если на выбранную дату уже имеются заполненные данные, пози" +
    "ция в справочнике РРЦ обновляется.";
            this.btnSavePrice.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.SavePrice_Click);
            // 
            // btnClientsUpdate
            // 
            this.btnClientsUpdate.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnClientsUpdate.Image = ((System.Drawing.Image)(resources.GetObject("btnClientsUpdate.Image")));
            this.btnClientsUpdate.Label = "Обновить клиентов";
            this.btnClientsUpdate.Name = "btnClientsUpdate";
            this.btnClientsUpdate.ScreenTip = "Обновление справочника клиентов";
            this.btnClientsUpdate.ShowImage = true;
            this.btnClientsUpdate.SuperTip = "Из загружаемого файла Decision в справочник клиентов вносятся данные новых клиент" +
    "ов, отсутствующих в справочнике. Путь к файлу Decision указывается в разделе нас" +
    "троек.";
            this.btnClientsUpdate.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ClientsUpdate_Click);
            // 
            // btnGetClientPrice
            // 
            this.btnGetClientPrice.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnGetClientPrice.Image = ((System.Drawing.Image)(resources.GetObject("btnGetClientPrice.Image")));
            this.btnGetClientPrice.Label = "Прайс-лист клиента";
            this.btnGetClientPrice.Name = "btnGetClientPrice";
            this.btnGetClientPrice.ScreenTip = "Формирование прайс-листа для выбранной категории клиентов";
            this.btnGetClientPrice.ShowImage = true;
            this.btnGetClientPrice.SuperTip = resources.GetString("btnGetClientPrice.SuperTip");
            this.btnGetClientPrice.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.GetClientPrice_Click);
            // 
            // btnGetAllPrices
            // 
            this.btnGetAllPrices.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnGetAllPrices.Image = ((System.Drawing.Image)(resources.GetObject("btnGetAllPrices.Image")));
            this.btnGetAllPrices.Label = "Все прайс-листы";
            this.btnGetAllPrices.Name = "btnGetAllPrices";
            this.btnGetAllPrices.ScreenTip = "Формирование прайс-листов для всех категорий клиентов";
            this.btnGetAllPrices.ShowImage = true;
            this.btnGetAllPrices.SuperTip = resources.GetString("btnGetAllPrices.SuperTip");
            this.btnGetAllPrices.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.GetAllPrices_Click);
            // 
            // btnPlanningAdd
            // 
            this.btnPlanningAdd.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnPlanningAdd.Image = ((System.Drawing.Image)(resources.GetObject("btnPlanningAdd.Image")));
            this.btnPlanningAdd.Label = "Создать планирование";
            this.btnPlanningAdd.Name = "btnPlanningAdd";
            this.btnPlanningAdd.ScreenTip = "Создание листа для формирования планирования";
            this.btnPlanningAdd.ShowImage = true;
            this.btnPlanningAdd.SuperTip = "Создается копия листа шаблона планирования, на котором необходимо указать год, Ch" +
    "annel type, Customer status для последующего формирования планирования. ";
            this.btnPlanningAdd.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.PlanningAdd_Click);
            // 
            // btnGetPlanningData
            // 
            this.btnGetPlanningData.Image = ((System.Drawing.Image)(resources.GetObject("btnGetPlanningData.Image")));
            this.btnGetPlanningData.Label = "Загрузить данные";
            this.btnGetPlanningData.Name = "btnGetPlanningData";
            this.btnGetPlanningData.ScreenTip = "Формирование листа планирования для категории клиентов";
            this.btnGetPlanningData.ShowImage = true;
            this.btnGetPlanningData.SuperTip = resources.GetString("btnGetPlanningData.SuperTip");
            this.btnGetPlanningData.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.GetPlanningData_Click);
            // 
            // btnFactUpdate
            // 
            this.btnFactUpdate.Image = ((System.Drawing.Image)(resources.GetObject("btnFactUpdate.Image")));
            this.btnFactUpdate.Label = "Обновить факт";
            this.btnFactUpdate.Name = "btnFactUpdate";
            this.btnFactUpdate.ScreenTip = "Обновление листа планирования для категории клиентов";
            this.btnFactUpdate.ShowImage = true;
            this.btnFactUpdate.SuperTip = resources.GetString("btnFactUpdate.SuperTip");
            this.btnFactUpdate.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.FactUpdate_Click);
            // 
            // btnPlanningSave
            // 
            this.btnPlanningSave.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnPlanningSave.Image = ((System.Drawing.Image)(resources.GetObject("btnPlanningSave.Image")));
            this.btnPlanningSave.Label = "Собрать данные";
            this.btnPlanningSave.Name = "btnPlanningSave";
            this.btnPlanningSave.ScreenTip = "Добавление товаров на лист сводного планирования";
            this.btnPlanningSave.ShowImage = true;
            this.btnPlanningSave.SuperTip = "Из указанных книг, которые содержат в качестве первых страниц раннее сформированн" +
    "ые листы планирования, данные построчно добавляются на лист сводного планировани" +
    "я.";
            this.btnPlanningSave.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.PlanningSave_Click);
            // 
            // btnSettings
            // 
            this.btnSettings.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnSettings.Image = ((System.Drawing.Image)(resources.GetObject("btnSettings.Image")));
            this.btnSettings.Label = "Настройки программы";
            this.btnSettings.Name = "btnSettings";
            this.btnSettings.ScreenTip = "Установка путей к файлам";
            this.btnSettings.ShowImage = true;
            this.btnSettings.SuperTip = "В данном разделе указываются пути к файлам Decision, Budget и PriceListMT.";
            this.btnSettings.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.Settings_Click);
            // 
            // btnSheetsHide
            // 
            this.btnSheetsHide.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnSheetsHide.Image = ((System.Drawing.Image)(resources.GetObject("btnSheetsHide.Image")));
            this.btnSheetsHide.Label = "Скрыть/ показать листы";
            this.btnSheetsHide.Name = "btnSheetsHide";
            this.btnSheetsHide.ScreenTip = "Скрытие/ показ листов с дополнительными справочниками";
            this.btnSheetsHide.ShowImage = true;
            this.btnSheetsHide.SuperTip = resources.GetString("btnSheetsHide.SuperTip");
            this.btnSheetsHide.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnSheetsHide_Click);
            // 
            // btnInfo
            // 
            this.btnInfo.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnInfo.Image = ((System.Drawing.Image)(resources.GetObject("btnInfo.Image")));
            this.btnInfo.Label = "Информация";
            this.btnInfo.Name = "btnInfo";
            this.btnInfo.ScreenTip = "Информация о приложении";
            this.btnInfo.ShowImage = true;
            this.btnInfo.SuperTip = "Информация о приложении и контактная информация.";
            this.btnInfo.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.About_Click);
            // 
            // RibbonBPA
            // 
            this.Name = "RibbonBPA";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.tabBPA);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.RibbonBPA_Load);
            this.tabBPA.ResumeLayout(false);
            this.tabBPA.PerformLayout();
            this.grpProducts.ResumeLayout(false);
            this.grpProducts.PerformLayout();
            this.grpPrice.ResumeLayout(false);
            this.grpPrice.PerformLayout();
            this.group2.ResumeLayout(false);
            this.group2.PerformLayout();
            this.group3.ResumeLayout(false);
            this.group3.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tabBPA;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnInfo;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup grpProducts;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnUpdateProducts;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnUpdateProduct;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup grpPrice;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group2;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnClientsUpdate;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnGetClientPrice;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnGetAllPrices;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnUploadPrice;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnSavePrice;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group3;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnPlanningAdd;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnGetPlanningData;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnFactUpdate;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnPlanningSave;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnSettings;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnSheetsHide;
        private Microsoft.Office.Tools.Ribbon.RibbonButton btnAddNewCalendar;
    }

    partial class ThisRibbonCollection
    {
        internal RibbonBPA RibbonBPA
        {
            get { return this.GetRibbon<RibbonBPA>(); }
        }
    }
}
