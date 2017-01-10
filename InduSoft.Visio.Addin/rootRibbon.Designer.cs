namespace InduSoft.Visio.Addin
{
    partial class rootRibbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public rootRibbon()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

        /// <summary> 
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Component Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.tab1 = this.Factory.CreateRibbonTab();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.editTime = this.Factory.CreateRibbonEditBox();
            this.btnWork = this.Factory.CreateRibbonToggleButton();
            this.editBox1 = this.Factory.CreateRibbonEditBox();
            this.editPeriodInSeconds = this.Factory.CreateRibbonEditBox();
            this.separator1 = this.Factory.CreateRibbonSeparator();
            this.tab1.SuspendLayout();
            this.group1.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Groups.Add(this.group1);
            this.tab1.Label = "INDUSOFT";
            this.tab1.Name = "tab1";
            // 
            // group1
            // 
            this.group1.Items.Add(this.editTime);
            this.group1.Items.Add(this.editPeriodInSeconds);
            this.group1.Items.Add(this.separator1);
            this.group1.Items.Add(this.btnWork);
            this.group1.Label = "Отладка схемы";
            this.group1.Name = "group1";
            // 
            // editTime
            // 
            this.editTime.Label = "Время значений";
            this.editTime.Name = "editTime";
            this.editTime.Text = "*";
            this.editTime.TextChanged += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.editTime_TextChanged);
            // 
            // btnWork
            // 
            this.btnWork.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnWork.Image = global::InduSoft.Visio.Addin.GlobalResource.iBug;
            this.btnWork.Label = "Режим";
            this.btnWork.Name = "btnWork";
            this.btnWork.ShowImage = true;
            this.btnWork.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.toggleButton1_Click);
            // 
            // editBox1
            // 
            this.editBox1.Label = "editBox1";
            this.editBox1.Name = "editBox1";
            // 
            // editPeriodInSeconds
            // 
            this.editPeriodInSeconds.Label = "Период обновления";
            this.editPeriodInSeconds.Name = "editPeriodInSeconds";
            this.editPeriodInSeconds.Text = "15";
            this.editPeriodInSeconds.TextChanged += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.editPeriodInSeconds_TextChanged);
            // 
            // separator1
            // 
            this.separator1.Name = "separator1";
            // 
            // rootRibbon
            // 
            this.Name = "rootRibbon";
            this.RibbonType = "Microsoft.Visio.Drawing";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.InduSoft_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonToggleButton btnWork;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox editTime;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox editPeriodInSeconds;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator1;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox editBox1;
    }

    partial class ThisRibbonCollection
    {
        internal rootRibbon InduSoft
        {
            get { return this.GetRibbon<rootRibbon>(); }
        }
    }
}
