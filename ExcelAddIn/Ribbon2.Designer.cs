namespace ExcelAddIn
{
    partial class Ribbon2 : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public Ribbon2()
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
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl1 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl2 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl3 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl4 = this.Factory.CreateRibbonDropDownItem();
            this.tab1 = this.Factory.CreateRibbonTab();
            this.groupHelp = this.Factory.CreateRibbonGroup();
            this.btn_help = this.Factory.CreateRibbonButton();
            this.groupTextToSpeech = this.Factory.CreateRibbonGroup();
            this.btn_TextToSpeech = this.Factory.CreateRibbonButton();
            this.groupRGB = this.Factory.CreateRibbonGroup();
            this.btn_RGB = this.Factory.CreateRibbonButton();
            this.dropDownColorRGB = this.Factory.CreateRibbonDropDown();
            this.editSaturationPeak = this.Factory.CreateRibbonDropDown();
            this.tab1.SuspendLayout();
            this.groupHelp.SuspendLayout();
            this.groupTextToSpeech.SuspendLayout();
            this.groupRGB.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Groups.Add(this.groupHelp);
            this.tab1.Groups.Add(this.groupTextToSpeech);
            this.tab1.Groups.Add(this.groupRGB);
            this.tab1.Label = "Add-In";
            this.tab1.Name = "tab1";
            // 
            // groupHelp
            // 
            this.groupHelp.Items.Add(this.btn_help);
            this.groupHelp.Label = "Help";
            this.groupHelp.Name = "groupHelp";
            // 
            // btn_help
            // 
            this.btn_help.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btn_help.Label = "Help";
            this.btn_help.Name = "btn_help";
            this.btn_help.ScreenTip = "Trợ giúp";
            this.btn_help.ShowImage = true;
            this.btn_help.SuperTip = "Click để truy cập trang";
            this.btn_help.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btn_help_Click);
            // 
            // groupTextToSpeech
            // 
            this.groupTextToSpeech.Items.Add(this.btn_TextToSpeech);
            this.groupTextToSpeech.Label = "TextToSpeech";
            this.groupTextToSpeech.Name = "groupTextToSpeech";
            // 
            // btn_TextToSpeech
            // 
            this.btn_TextToSpeech.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btn_TextToSpeech.Label = "TextToSpeech";
            this.btn_TextToSpeech.Name = "btn_TextToSpeech";
            this.btn_TextToSpeech.ScreenTip = "Chuyển nội dung sang giọng nói";
            this.btn_TextToSpeech.ShowImage = true;
            this.btn_TextToSpeech.SuperTip = "Convert Text to Speech";
            this.btn_TextToSpeech.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btn_TextToSpeech_Click);
            // 
            // groupRGB
            // 
            this.groupRGB.Items.Add(this.btn_RGB);
            this.groupRGB.Items.Add(this.dropDownColorRGB);
            this.groupRGB.Items.Add(this.editSaturationPeak);
            this.groupRGB.Label = "RGB";
            this.groupRGB.Name = "groupRGB";
            // 
            // btn_RGB
            // 
            this.btn_RGB.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btn_RGB.Label = "RGB";
            this.btn_RGB.Name = "btn_RGB";
            this.btn_RGB.OfficeImageId = "BlackAndWhiteLightGrayscale";
            this.btn_RGB.ScreenTip = "Màu hóa ma trận giá trị";
            this.btn_RGB.ShowImage = true;
            this.btn_RGB.SuperTip = "Lựa chọn một bảng, sau đó bấm nút Màu hóa. Các cell trong bảng sẽ được tô màu với" +
    " mức xám thay đổi  từ màu đen (0) tới mức cực đại, của màu chỉ định trong dropbo" +
    "x";
            // 
            // dropDownColorRGB
            // 
            ribbonDropDownItemImpl1.Label = "Đỏ";
            ribbonDropDownItemImpl1.OfficeImageId = "AppointmentColor1";
            ribbonDropDownItemImpl2.Label = "Xanh lá";
            ribbonDropDownItemImpl2.OfficeImageId = "AppointmentColor3";
            ribbonDropDownItemImpl3.Label = "Xanh dương";
            ribbonDropDownItemImpl3.OfficeImageId = "AppointmentColor2";
            ribbonDropDownItemImpl4.Label = "Xám";
            ribbonDropDownItemImpl4.OfficeImageId = "AppointmentColor4";
            this.dropDownColorRGB.Items.Add(ribbonDropDownItemImpl1);
            this.dropDownColorRGB.Items.Add(ribbonDropDownItemImpl2);
            this.dropDownColorRGB.Items.Add(ribbonDropDownItemImpl3);
            this.dropDownColorRGB.Items.Add(ribbonDropDownItemImpl4);
            this.dropDownColorRGB.Label = "Màu";
            this.dropDownColorRGB.Name = "dropDownColorRGB";
            this.dropDownColorRGB.SelectionChanged += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.dropDown1_SelectionChanged);
            // 
            // editSaturationPeak
            // 
            this.editSaturationPeak.Label = "Cực đại";
            this.editSaturationPeak.Name = "editSaturationPeak";
            // 
            // Ribbon2
            // 
            this.Name = "Ribbon2";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon2_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.groupHelp.ResumeLayout(false);
            this.groupHelp.PerformLayout();
            this.groupTextToSpeech.ResumeLayout(false);
            this.groupTextToSpeech.PerformLayout();
            this.groupRGB.ResumeLayout(false);
            this.groupRGB.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup groupHelp;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btn_help;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup groupTextToSpeech;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btn_TextToSpeech;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup groupRGB;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btn_RGB;
        internal Microsoft.Office.Tools.Ribbon.RibbonDropDown dropDownColorRGB;
        internal Microsoft.Office.Tools.Ribbon.RibbonDropDown editSaturationPeak;
    }

    partial class ThisRibbonCollection
    {
        internal Ribbon2 Ribbon2
        {
            get { return this.GetRibbon<Ribbon2>(); }
        }
    }
}
