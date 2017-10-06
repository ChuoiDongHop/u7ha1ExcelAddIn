namespace u7ha1ExcelAddIn
{
   partial class Ribbon : Microsoft.Office.Tools.Ribbon.RibbonBase
   {
      /// <summary>
      /// Required designer variable.
      /// </summary>
      private System.ComponentModel.IContainer components = null;

      public Ribbon()
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
         this.u7ha1Tab = this.Factory.CreateRibbonTab();
         this.u7ha1Group = this.Factory.CreateRibbonGroup();
         this.u7ha1Label = this.Factory.CreateRibbonLabel();
         this.u7ha1Tab.SuspendLayout();
         this.u7ha1Group.SuspendLayout();
         this.SuspendLayout();
         // 
         // u7ha1Tab
         // 
         this.u7ha1Tab.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
         this.u7ha1Tab.Groups.Add(this.u7ha1Group);
         this.u7ha1Tab.Label = "u7ha1";
         this.u7ha1Tab.Name = "u7ha1Tab";
         // 
         // u7ha1Group
         // 
         this.u7ha1Group.Items.Add(this.u7ha1Label);
         this.u7ha1Group.Label = "u7ha1";
         this.u7ha1Group.Name = "u7ha1Group";
         // 
         // u7ha1Label
         // 
         this.u7ha1Label.Label = "u7ha1";
         this.u7ha1Label.Name = "u7ha1Label";
         // 
         // Ribbon
         // 
         this.Name = "Ribbon";
         this.RibbonType = "Microsoft.Excel.Workbook";
         this.Tabs.Add(this.u7ha1Tab);
         this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon_Load);
         this.u7ha1Tab.ResumeLayout(false);
         this.u7ha1Tab.PerformLayout();
         this.u7ha1Group.ResumeLayout(false);
         this.u7ha1Group.PerformLayout();
         this.ResumeLayout(false);

      }

      #endregion

      internal Microsoft.Office.Tools.Ribbon.RibbonTab u7ha1Tab;
      internal Microsoft.Office.Tools.Ribbon.RibbonGroup u7ha1Group;
      internal Microsoft.Office.Tools.Ribbon.RibbonLabel u7ha1Label;
   }

   partial class ThisRibbonCollection
   {
      internal Ribbon Ribbon
      {
         get { return this.GetRibbon<Ribbon>(); }
      }
   }
}
