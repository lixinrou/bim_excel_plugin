namespace ExcelAddIn
{
    partial class Ribbon1 : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// 必需的设计器变量。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public Ribbon1()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

        /// <summary> 
        /// 清理所有正在使用的资源。
        /// </summary>
        /// <param name="disposing">如果应释放托管资源，为 true；否则为 false。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region 组件设计器生成的代码

        /// <summary>
        /// 设计器支持所需的方法 - 不要
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InitializeComponent()
        {
            Microsoft.Office.Tools.Ribbon.RibbonDialogLauncher ribbonDialogLauncherImpl1 = this.Factory.CreateRibbonDialogLauncher();
            Microsoft.Office.Tools.Ribbon.RibbonDialogLauncher ribbonDialogLauncherImpl2 = this.Factory.CreateRibbonDialogLauncher();
            Microsoft.Office.Tools.Ribbon.RibbonDialogLauncher ribbonDialogLauncherImpl3 = this.Factory.CreateRibbonDialogLauncher();
            this.tab1 = this.Factory.CreateRibbonTab();
            this.group = this.Factory.CreateRibbonGroup();
            this.SingleCell = this.Factory.CreateRibbonButton();
            this.separator1 = this.Factory.CreateRibbonSeparator();
            this.MultiCell = this.Factory.CreateRibbonButton();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.button2 = this.Factory.CreateRibbonButton();
            this.separator2 = this.Factory.CreateRibbonSeparator();
            this.button4 = this.Factory.CreateRibbonButton();
            this.button3 = this.Factory.CreateRibbonButton();
            this.group2 = this.Factory.CreateRibbonGroup();
            this.button5 = this.Factory.CreateRibbonButton();
            this.separator3 = this.Factory.CreateRibbonSeparator();
            this.button6 = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.group.SuspendLayout();
            this.group1.SuspendLayout();
            this.group2.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.Groups.Add(this.group);
            this.tab1.Groups.Add(this.group1);
            this.tab1.Groups.Add(this.group2);
            this.tab1.Label = "BIM应用插件";
            this.tab1.Name = "tab1";
            // 
            // group
            // 
            this.group.DialogLauncher = ribbonDialogLauncherImpl1;
            this.group.Items.Add(this.SingleCell);
            this.group.Items.Add(this.separator1);
            this.group.Items.Add(this.MultiCell);
            this.group.Label = "文本预处理";
            this.group.Name = "group";
            this.group.DialogLauncherClick += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.group_DialogLauncherClick);
            // 
            // SingleCell
            // 
            this.SingleCell.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.SingleCell.Image = global::ExcelAddIn.Properties.Resources._1145472;
            this.SingleCell.Label = "单元格格式化";
            this.SingleCell.Name = "SingleCell";
            this.SingleCell.ScreenTip = "单元格格式化";
            this.SingleCell.ShowImage = true;
            this.SingleCell.SuperTip = "格式化单个单元格的文本为SQL需要的格式，可同时操作多个单元格。";
            this.SingleCell.Tag = "";
            this.SingleCell.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.SingleCell_Click);
            // 
            // separator1
            // 
            this.separator1.Name = "separator1";
            // 
            // MultiCell
            // 
            this.MultiCell.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.MultiCell.Image = global::ExcelAddIn.Properties.Resources._1145507;
            this.MultiCell.Label = "跨单元格格式化";
            this.MultiCell.Name = "MultiCell";
            this.MultiCell.OfficeImageId = "Call";
            this.MultiCell.ShowImage = true;
            this.MultiCell.SuperTip = "格式化多个单元格的文本为SQL需要的格式，将多个单元格的文本组装在一起。";
            this.MultiCell.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.MultiCell_Click);
            // 
            // group1
            // 
            this.group1.DialogLauncher = ribbonDialogLauncherImpl2;
            this.group1.Items.Add(this.button2);
            this.group1.Items.Add(this.separator2);
            this.group1.Items.Add(this.button4);
            this.group1.Items.Add(this.button3);
            this.group1.Label = "数据导入";
            this.group1.Name = "group1";
            this.group1.DialogLauncherClick += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.group1_DialogLauncherClick);
            // 
            // button2
            // 
            this.button2.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button2.Image = global::ExcelAddIn.Properties.Resources._1145566;
            this.button2.Label = "运行参数";
            this.button2.Name = "button2";
            this.button2.OfficeImageId = "Call";
            this.button2.ShowImage = true;
            // 
            // separator2
            // 
            this.separator2.Name = "separator2";
            // 
            // button4
            // 
            this.button4.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button4.Image = global::ExcelAddIn.Properties.Resources._1145595;
            this.button4.Label = "二级系统";
            this.button4.Name = "button4";
            this.button4.OfficeImageId = "Call";
            this.button4.ShowImage = true;
            // 
            // button3
            // 
            this.button3.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button3.Image = global::ExcelAddIn.Properties.Resources._1145507;
            this.button3.Label = "三级系统";
            this.button3.Name = "button3";
            this.button3.OfficeImageId = "Call";
            this.button3.ShowImage = true;
            // 
            // group2
            // 
            this.group2.DialogLauncher = ribbonDialogLauncherImpl3;
            this.group2.Items.Add(this.button5);
            this.group2.Items.Add(this.separator3);
            this.group2.Items.Add(this.button6);
            this.group2.Label = "数据查询";
            this.group2.Name = "group2";
            this.group2.DialogLauncherClick += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.group2_DialogLauncherClick);
            // 
            // button5
            // 
            this.button5.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button5.Image = global::ExcelAddIn.Properties.Resources._1145472;
            this.button5.Label = "系统分组";
            this.button5.Name = "button5";
            this.button5.ShowImage = true;
            this.button5.SuperTip = "格式化文本为SQL需要的格式";
            // 
            // separator3
            // 
            this.separator3.Name = "separator3";
            // 
            // button6
            // 
            this.button6.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button6.Image = global::ExcelAddIn.Properties.Resources._1145507;
            this.button6.Label = "系统分类";
            this.button6.Name = "button6";
            this.button6.OfficeImageId = "Call";
            this.button6.ShowImage = true;
            this.button6.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button6_Click);
            // 
            // Ribbon1
            // 
            this.Name = "Ribbon1";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon1_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.group.ResumeLayout(false);
            this.group.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.group2.ResumeLayout(false);
            this.group2.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton SingleCell;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton MultiCell;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button2;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator2;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button4;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button3;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group2;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button5;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator3;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button6;
    }

    partial class ThisRibbonCollection
    {
        internal Ribbon1 Ribbon1
        {
            get { return this.GetRibbon<Ribbon1>(); }
        }
    }
}
