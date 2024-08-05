namespace OutlookMacroPlugin
{
    partial class MyRibbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {

        private System.ComponentModel.IContainer components = null;

        public MyRibbon()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

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

        private void InitializeComponent()
        {
            this.tab1 = this.Factory.CreateRibbonTab();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.MoveToTreatedSDButton = this.Factory.CreateRibbonButton();
            this.SignButton = this.Factory.CreateRibbonButton();
            this.IBMButton = this.Factory.CreateRibbonButton();
            this.EGButton = this.Factory.CreateRibbonButton();
            this.NCRButton = this.Factory.CreateRibbonButton();
            this.WincorButton = this.Factory.CreateRibbonButton();
            this.NDButton = this.Factory.CreateRibbonButton();
            this.HPEButton = this.Factory.CreateRibbonButton();
            this.SIMAButton = this.Factory.CreateRibbonButton();
            this.MIMButton = this.Factory.CreateRibbonButton();
            this.LexmarkButton = this.Factory.CreateRibbonButton();
            this.AteaButton = this.Factory.CreateRibbonButton();
            this.RicohButton = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.group1.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Groups.Add(this.group1);
            this.tab1.Label = "Servicedesk";
            this.tab1.Name = "tab1";
            // 
            // group1
            // 
            this.group1.Items.Add(this.MoveToTreatedSDButton);
            this.group1.Items.Add(this.SignButton);
            this.group1.Items.Add(this.IBMButton);
            this.group1.Items.Add(this.EGButton);
            this.group1.Items.Add(this.NCRButton);
            this.group1.Items.Add(this.WincorButton);
            this.group1.Items.Add(this.NDButton);
            this.group1.Items.Add(this.HPEButton);
            this.group1.Items.Add(this.SIMAButton);
            this.group1.Items.Add(this.MIMButton);
            this.group1.Items.Add(this.LexmarkButton);
            this.group1.Items.Add(this.RicohButton);
            this.group1.Items.Add(this.AteaButton);
            this.group1.Label = "Custom Group";
            this.group1.Name = "group1";
            // 
            // MoveToTreatedSDButton
            // 
            this.MoveToTreatedSDButton.Label = "Move to Treated SD";
            this.MoveToTreatedSDButton.Name = "MoveToTreatedSDButton";
            this.MoveToTreatedSDButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.MoveToTreatedSDButton_Click);
            // 
            // SignButton
            // 
            this.SignButton.Label = "Sign";
            this.SignButton.Name = "SignButton";
            this.SignButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.SignButton_Click);
            // 
            // IBMButton
            // 
            this.IBMButton.Label = "Send to IBM";
            this.IBMButton.Name = "IBMButton";
            this.IBMButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.IBMButton_Click);
            // 
            // EGButton
            // 
            this.EGButton.Label = "Send to EG";
            this.EGButton.Name = "EGButton";
            this.EGButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.EGButton_Click);
            // 
            // NCRButton
            // 
            this.NCRButton.Label = "Send to NCR";
            this.NCRButton.Name = "NCRButton";
            this.NCRButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.NCRButton_Click);
            // 
            // WincorButton
            // 
            this.WincorButton.Label = "Send to Wincor";
            this.WincorButton.Name = "WincorButton";
            this.WincorButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.WincorButton_Click);
            // 
            // NDButton
            // 
            this.NDButton.Label = "Send to ND";
            this.NDButton.Name = "NDButton";
            this.NDButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.NDButton_Click);
            // 
            // HPEButton
            // 
            this.HPEButton.Label = "Send to HPE";
            this.HPEButton.Name = "HPEButton";
            this.HPEButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.HPEButton_Click);
            // 
            // SIMAButton
            // 
            this.SIMAButton.Label = "Send to SIMA";
            this.SIMAButton.Name = "SIMAButton";
            this.SIMAButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.SIMAButton_Click);
            // 
            // MIMButton
            // 
            this.MIMButton.Label = "Send to MIM";
            this.MIMButton.Name = "MIMButton";
            this.MIMButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.MIMButton_Click);
            // 
            // LexmarkButton
            // 
            this.LexmarkButton.Label = "Send to Lexmark";
            this.LexmarkButton.Name = "LexmarkButton";
            this.LexmarkButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.LexmarkButton_Click);
            // 
            // AteaButton
            // 
            this.AteaButton.Label = "Send to Atea";
            this.AteaButton.Name = "AteaButton";
            this.AteaButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.AteaButton_Click);
            // 
            // RicohButton
            // 
            this.RicohButton.Label = "Send to Ricoh";
            this.RicohButton.Name = "RicohButton";
            this.RicohButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.RicohButton_Click);
            // 
            // MyRibbon
            // 
            this.Name = "MyRibbon";
            this.RibbonType = "Microsoft.Outlook.Explorer, Microsoft.Outlook.Mail.Read";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.MyRibbon_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton MoveToTreatedSDButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton SignButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton IBMButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton EGButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton NCRButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton WincorButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton NDButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton HPEButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton SIMAButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton MIMButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton LexmarkButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton AteaButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton RicohButton;
    }

    partial class ThisRibbonCollection
    {
        internal MyRibbon MyRibbon
        {
            get { return this.GetRibbon<MyRibbon>(); }
        }
    }
}
