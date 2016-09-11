namespace DataFetcherAddIn
{
    partial class MyRibbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Обязательная переменная конструктора.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public MyRibbon()
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(MyRibbon));
            this.tab1 = this.Factory.CreateRibbonTab();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.buttonRefreshData = this.Factory.CreateRibbonButton();
            this.toggleButtonShowFields = this.Factory.CreateRibbonToggleButton();
            this.buttonGroup1 = this.Factory.CreateRibbonButtonGroup();
            this.buttonPrev = this.Factory.CreateRibbonButton();
            this.buttonNext = this.Factory.CreateRibbonButton();
            this.group2 = this.Factory.CreateRibbonGroup();
            this.buttonAliases = this.Factory.CreateRibbonButton();
            this.buttonTextFromExcel = this.Factory.CreateRibbonButton();
            this.buttonTable = this.Factory.CreateRibbonButton();
            this.button3 = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.group1.SuspendLayout();
            this.buttonGroup1.SuspendLayout();
            this.group2.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Groups.Add(this.group1);
            this.tab1.Groups.Add(this.group2);
            this.tab1.Label = "Сбор данных";
            this.tab1.Name = "tab1";
            // 
            // group1
            // 
            this.group1.Items.Add(this.buttonRefreshData);
            this.group1.Items.Add(this.toggleButtonShowFields);
            this.group1.Items.Add(this.buttonGroup1);
            this.group1.Label = "Документ";
            this.group1.Name = "group1";
            // 
            // buttonRefreshData
            // 
            this.buttonRefreshData.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.buttonRefreshData.Image = ((System.Drawing.Image)(resources.GetObject("buttonRefreshData.Image")));
            this.buttonRefreshData.Label = "Обновить";
            this.buttonRefreshData.Name = "buttonRefreshData";
            this.buttonRefreshData.ShowImage = true;
            this.buttonRefreshData.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonRefreshData_Click);
            // 
            // toggleButtonShowFields
            // 
            this.toggleButtonShowFields.Label = "Показать код";
            this.toggleButtonShowFields.Name = "toggleButtonShowFields";
            this.toggleButtonShowFields.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.toggleButtonShowFields_Click);
            // 
            // buttonGroup1
            // 
            this.buttonGroup1.Items.Add(this.buttonPrev);
            this.buttonGroup1.Items.Add(this.buttonNext);
            this.buttonGroup1.Name = "buttonGroup1";
            this.buttonGroup1.Visible = false;
            // 
            // buttonPrev
            // 
            this.buttonPrev.Label = "button1";
            this.buttonPrev.Name = "buttonPrev";
            this.buttonPrev.Visible = false;
            // 
            // buttonNext
            // 
            this.buttonNext.Label = "button2";
            this.buttonNext.Name = "buttonNext";
            this.buttonNext.Visible = false;
            // 
            // group2
            // 
            this.group2.Items.Add(this.buttonAliases);
            this.group2.Items.Add(this.buttonTextFromExcel);
            this.group2.Items.Add(this.buttonTable);
            this.group2.Items.Add(this.button3);
            this.group2.Label = "Поля";
            this.group2.Name = "group2";
            // 
            // buttonAliases
            // 
            this.buttonAliases.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.buttonAliases.Image = ((System.Drawing.Image)(resources.GetObject("buttonAliases.Image")));
            this.buttonAliases.Label = "Подстановки";
            this.buttonAliases.Name = "buttonAliases";
            this.buttonAliases.ShowImage = true;
            this.buttonAliases.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.toggleButtonShowPanel_Click);
            // 
            // buttonTextFromExcel
            // 
            this.buttonTextFromExcel.Image = ((System.Drawing.Image)(resources.GetObject("buttonTextFromExcel.Image")));
            this.buttonTextFromExcel.Label = "Текст из ячейки";
            this.buttonTextFromExcel.Name = "buttonTextFromExcel";
            this.buttonTextFromExcel.ShowImage = true;
            this.buttonTextFromExcel.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonTextFromExcel_Click);
            // 
            // buttonTable
            // 
            this.buttonTable.Image = ((System.Drawing.Image)(resources.GetObject("buttonTable.Image")));
            this.buttonTable.Label = "Таблица";
            this.buttonTable.Name = "buttonTable";
            this.buttonTable.ShowImage = true;
            this.buttonTable.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonTable_Click);
            // 
            // button3
            // 
            this.button3.Image = ((System.Drawing.Image)(resources.GetObject("button3.Image")));
            this.button3.Label = "Диаграмма";
            this.button3.Name = "button3";
            this.button3.ShowImage = true;
            this.button3.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button3_Click);
            // 
            // MyRibbon
            // 
            this.Name = "MyRibbon";
            this.RibbonType = "Microsoft.Word.Document";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.MyRibbon_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.buttonGroup1.ResumeLayout(false);
            this.buttonGroup1.PerformLayout();
            this.group2.ResumeLayout(false);
            this.group2.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonRefreshData;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group2;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonTextFromExcel;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonTable;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button3;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonAliases;
        internal Microsoft.Office.Tools.Ribbon.RibbonToggleButton toggleButtonShowFields;
        internal Microsoft.Office.Tools.Ribbon.RibbonButtonGroup buttonGroup1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonPrev;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonNext;
    }

    partial class ThisRibbonCollection
    {
        internal MyRibbon MyRibbon
        {
            get { return this.GetRibbon<MyRibbon>(); }
        }
    }
}
