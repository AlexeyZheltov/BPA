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
            this.tabBPA = this.Factory.CreateRibbonTab();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.button1 = this.Factory.CreateRibbonButton();
            this.group2 = this.Factory.CreateRibbonGroup();
            this.button2 = this.Factory.CreateRibbonButton();
            this.button3 = this.Factory.CreateRibbonButton();
            this.tabBPA.SuspendLayout();
            this.group1.SuspendLayout();
            this.group2.SuspendLayout();
            this.SuspendLayout();
            // 
            // tabBPA
            // 
            this.tabBPA.Groups.Add(this.group1);
            this.tabBPA.Groups.Add(this.group2);
            this.tabBPA.Label = "BPA";
            this.tabBPA.Name = "tabBPA";
            this.tabBPA.Position = this.Factory.RibbonPosition.BeforeOfficeId("TabHome");
            // 
            // group1
            // 
            this.group1.Items.Add(this.button1);
            this.group1.Label = "group1";
            this.group1.Name = "group1";
            // 
            // button1
            // 
            this.button1.Label = "Добавить суперкатегорию";
            this.button1.Name = "button1";
            this.button1.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button1_Click);
            // 
            // group2
            // 
            this.group2.Items.Add(this.button2);
            this.group2.Items.Add(this.button3);
            this.group2.Label = "Продуктовый календарь";
            this.group2.Name = "group2";
            // 
            // button2
            // 
            this.button2.Label = "Загрузить";
            this.button2.Name = "button2";
            this.button2.ScreenTip = "Загрузить новый продуктовый календарь";
            this.button2.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button2_Click);
            // 
            // button3
            // 
            this.button3.Label = "Обновить";
            this.button3.Name = "button3";
            this.button3.ScreenTip = "Обновить из продуктового календаря";
            this.button3.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button3_Click);
            // 
            // RibbonBPA
            // 
            this.Name = "RibbonBPA";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.tabBPA);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.RibbonBPA_Load);
            this.tabBPA.ResumeLayout(false);
            this.tabBPA.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.group2.ResumeLayout(false);
            this.group2.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tabBPA;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group2;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button2;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button3;
    }

    partial class ThisRibbonCollection
    {
        internal RibbonBPA RibbonBPA
        {
            get { return this.GetRibbon<RibbonBPA>(); }
        }
    }
}
