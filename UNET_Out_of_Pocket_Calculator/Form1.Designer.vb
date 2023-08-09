<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Form1
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.BackgroundWorker1 = New System.ComponentModel.BackgroundWorker()
        Me.MenuStrip1 = New System.Windows.Forms.MenuStrip()
        Me.MHIHistoryToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.GetMMIToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.CalculateOOPToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.GatherProvInfoToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ResetViewToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ClearToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.MHIShortOptionToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.PaitentNameAndDOSToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.SortByICNToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ProviderTinAndSuffixToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.DateOfServiceToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ProviderTinToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.DeductibleIndicatorToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.SortByPercentToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.PatientAndProcessedDateToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ProcessedDateAndDraftToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ProcessedDateOnlyToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ELGSLetterToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.FormatMHIToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.CleanClaimToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.FormatMHISheetToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.InstructionsToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.AboutToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ClearAllFilterToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ApplyColorToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.Chk_select_memlist = New System.Windows.Forms.CheckBox()
        Me.yearList = New System.Windows.Forms.ComboBox()
        Me.btnClaimInfo = New System.Windows.Forms.Button()
        Me.btnMain = New System.Windows.Forms.Button()
        Me.endSelect = New System.Windows.Forms.DateTimePicker()
        Me.startSelect = New System.Windows.Forms.DateTimePicker()
        Me.txt_SSN = New System.Windows.Forms.TextBox()
        Me.txt_Policy = New System.Windows.Forms.TextBox()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.memberList = New System.Windows.Forms.CheckedListBox()
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.txtUserPass = New System.Windows.Forms.TextBox()
        Me.txtUserID = New System.Windows.Forms.TextBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.Panel2 = New System.Windows.Forms.Panel()
        Me.RichTextBox1 = New System.Windows.Forms.RichTextBox()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.TabPage13 = New System.Windows.Forms.TabPage()
        Me.DGridMHI = New Zuby.ADGV.AdvancedDataGridView()
        Me.From = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Thru = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Svc = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.PS = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Nbr = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.OV = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.P = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.N = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.RC = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Charge = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.NotCov = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.BM = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Covered = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Deduct = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.D = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Perc = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Paid = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.S = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.DC = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Sanc = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.CauseCode = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.P1 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Tin = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Suffix = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.ClaimNumber = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Draft = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.ProcDate = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.AdjNo = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.TotalBilled = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.TotalPaid = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.ICN = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Suf = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.FLN = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.PRS = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.SI = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.PT_Name = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Blank = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.PTRel = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.PTName = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.INN_Ded = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.INN_OOP = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.OON_Ded = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.OON_OOP = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.INNDed = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.INNOOP = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.OONDed = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.OONOOP = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.OI_OIM = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.OOPCalcRun = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.ICNandSuffix = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Inp_Facility = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.ProviderName = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.ProviderType = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.M1 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.M2 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.M3 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.M4 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.TabPage11 = New System.Windows.Forms.TabPage()
        Me.DGridADJ = New System.Windows.Forms.DataGridView()
        Me.TabPage12 = New System.Windows.Forms.TabPage()
        Me.btnOOPExport = New System.Windows.Forms.Button()
        Me.tblOOP = New System.Windows.Forms.DataGridView()
        Me.TabPage10 = New System.Windows.Forms.TabPage()
        Me.tblCopay = New System.Windows.Forms.DataGridView()
        Me.Policy = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Plan_CodeReportingCode_PlanVar = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Year = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.PatientName = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.CauseCodes = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.PlaceofService = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.MajorMedCalcCode = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.SpecialProcCode = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.CopaySet = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.DataGridViewTextBoxColumn110 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.TabPage8 = New System.Windows.Forms.TabPage()
        Me.DGrid_PG10 = New System.Windows.Forms.DataGridView()
        Me.DataGridViewTextBoxColumn29 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.DataGridViewTextBoxColumn30 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.DataGridViewTextBoxColumn31 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.DataGridViewTextBoxColumn32 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.DataGridViewTextBoxColumn33 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.DataGridViewTextBoxColumn34 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.DataGridViewTextBoxColumn35 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Column_H = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Column_I = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Column_J = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Column_K = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Column_L = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Column_M = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Column_N = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Column_O = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Column_P = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.TabPage5 = New System.Windows.Forms.TabPage()
        Me.DGrid_PG5 = New System.Windows.Forms.DataGridView()
        Me.DataGridViewTextBoxColumn15 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.DataGridViewTextBoxColumn16 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.DataGridViewTextBoxColumn17 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.DataGridViewTextBoxColumn18 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.DataGridViewTextBoxColumn19 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.DataGridViewTextBoxColumn20 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.DataGridViewTextBoxColumn21 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Colu_H = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Colu_I = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Colu_J = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Colu_K = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Colu_L = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Colu_M = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Colu_N = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Colu_O = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Colu_P = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.TabPage4 = New System.Windows.Forms.TabPage()
        Me.DGrid_PG4 = New System.Windows.Forms.DataGridView()
        Me.DataGridViewTextBoxColumn8 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.DataGridViewTextBoxColumn9 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.DataGridViewTextBoxColumn10 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.DataGridViewTextBoxColumn11 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.DataGridViewTextBoxColumn12 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.DataGridViewTextBoxColumn13 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.DataGridViewTextBoxColumn14 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.C_H = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.C_I = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.C_J = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.C_K = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.C_L = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.C_M = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.C_N = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.TabPage3 = New System.Windows.Forms.TabPage()
        Me.DGrid_PG1 = New System.Windows.Forms.DataGridView()
        Me.DataGridViewTextBoxColumn1 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.DataGridViewTextBoxColumn2 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.DataGridViewTextBoxColumn3 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.DataGridViewTextBoxColumn4 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.DataGridViewTextBoxColumn5 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.DataGridViewTextBoxColumn6 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.DataGridViewTextBoxColumn7 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.CL_H = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.CL_I = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.CLL_J = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.CLL_K = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.CLL_L = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.CLL_M = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.CLL_N = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.TabPage1 = New System.Windows.Forms.TabPage()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.MMI_Message = New System.Windows.Forms.RichTextBox()
        Me.DGridOverview = New System.Windows.Forms.DataGridView()
        Me.Col_A = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.COL_O = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.COL_P = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Col_B = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Col_C = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Col_D = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Col_E = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Col_F = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Col_G = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Col_H = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Col_I = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Col_J = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Cl_J = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Col_k = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Col_L = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Col_M = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Col_N = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.TabPage2 = New System.Windows.Forms.TabPage()
        Me.btnCEIExport = New System.Windows.Forms.Button()
        Me.DGridMInfo = New Zuby.ADGV.AdvancedDataGridView()
        Me.MLast_Name = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.MAddress = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.PayLoc_Eng = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.DGridCEI = New Zuby.ADGV.AdvancedDataGridView()
        Me.PFName = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.PRelation = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.DOB = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.End_DT = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.MHI = New System.Windows.Forms.TabPage()
        Me.Button2 = New System.Windows.Forms.Button()
        Me.tblMHI = New Zuby.ADGV.AdvancedDataGridView()
        Me.TabControl1 = New System.Windows.Forms.TabControl()
        Me.TabPage6 = New System.Windows.Forms.TabPage()
        Me.DGridMHI_II = New System.Windows.Forms.DataGridView()
        Me.MenuStrip1.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.TabPage13.SuspendLayout()
        CType(Me.DGridMHI, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TabPage11.SuspendLayout()
        CType(Me.DGridADJ, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TabPage12.SuspendLayout()
        CType(Me.tblOOP, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TabPage10.SuspendLayout()
        CType(Me.tblCopay, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TabPage8.SuspendLayout()
        CType(Me.DGrid_PG10, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TabPage5.SuspendLayout()
        CType(Me.DGrid_PG5, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TabPage4.SuspendLayout()
        CType(Me.DGrid_PG4, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TabPage3.SuspendLayout()
        CType(Me.DGrid_PG1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TabPage1.SuspendLayout()
        CType(Me.DGridOverview, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TabPage2.SuspendLayout()
        CType(Me.DGridMInfo, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.DGridCEI, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.MHI.SuspendLayout()
        CType(Me.tblMHI, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TabControl1.SuspendLayout()
        Me.TabPage6.SuspendLayout()
        CType(Me.DGridMHI_II, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'MenuStrip1
        '
        Me.MenuStrip1.ImageScalingSize = New System.Drawing.Size(20, 20)
        Me.MenuStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.MHIHistoryToolStripMenuItem, Me.MHIShortOptionToolStripMenuItem, Me.ELGSLetterToolStripMenuItem, Me.FormatMHIToolStripMenuItem, Me.InstructionsToolStripMenuItem, Me.AboutToolStripMenuItem, Me.ClearAllFilterToolStripMenuItem, Me.ApplyColorToolStripMenuItem})
        Me.MenuStrip1.Location = New System.Drawing.Point(0, 0)
        Me.MenuStrip1.Name = "MenuStrip1"
        Me.MenuStrip1.Size = New System.Drawing.Size(1924, 28)
        Me.MenuStrip1.TabIndex = 0
        Me.MenuStrip1.Text = "MenuStrip1"
        '
        'MHIHistoryToolStripMenuItem
        '
        Me.MHIHistoryToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.GetMMIToolStripMenuItem, Me.CalculateOOPToolStripMenuItem, Me.GatherProvInfoToolStripMenuItem, Me.ResetViewToolStripMenuItem, Me.ClearToolStripMenuItem})
        Me.MHIHistoryToolStripMenuItem.Name = "MHIHistoryToolStripMenuItem"
        Me.MHIHistoryToolStripMenuItem.Size = New System.Drawing.Size(102, 24)
        Me.MHIHistoryToolStripMenuItem.Text = "MHI History"
        '
        'GetMMIToolStripMenuItem
        '
        Me.GetMMIToolStripMenuItem.Name = "GetMMIToolStripMenuItem"
        Me.GetMMIToolStripMenuItem.Size = New System.Drawing.Size(199, 26)
        Me.GetMMIToolStripMenuItem.Text = "Get MMI"
        '
        'CalculateOOPToolStripMenuItem
        '
        Me.CalculateOOPToolStripMenuItem.Name = "CalculateOOPToolStripMenuItem"
        Me.CalculateOOPToolStripMenuItem.Size = New System.Drawing.Size(199, 26)
        Me.CalculateOOPToolStripMenuItem.Text = "Calculate OOP"
        '
        'GatherProvInfoToolStripMenuItem
        '
        Me.GatherProvInfoToolStripMenuItem.Name = "GatherProvInfoToolStripMenuItem"
        Me.GatherProvInfoToolStripMenuItem.Size = New System.Drawing.Size(199, 26)
        Me.GatherProvInfoToolStripMenuItem.Text = "Gather Prov Info"
        '
        'ResetViewToolStripMenuItem
        '
        Me.ResetViewToolStripMenuItem.Name = "ResetViewToolStripMenuItem"
        Me.ResetViewToolStripMenuItem.Size = New System.Drawing.Size(199, 26)
        Me.ResetViewToolStripMenuItem.Text = "Reset View"
        '
        'ClearToolStripMenuItem
        '
        Me.ClearToolStripMenuItem.Name = "ClearToolStripMenuItem"
        Me.ClearToolStripMenuItem.Size = New System.Drawing.Size(199, 26)
        Me.ClearToolStripMenuItem.Text = "Clear"
        '
        'MHIShortOptionToolStripMenuItem
        '
        Me.MHIShortOptionToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.PaitentNameAndDOSToolStripMenuItem, Me.SortByICNToolStripMenuItem, Me.ProviderTinAndSuffixToolStripMenuItem, Me.DateOfServiceToolStripMenuItem, Me.ProviderTinToolStripMenuItem, Me.DeductibleIndicatorToolStripMenuItem, Me.SortByPercentToolStripMenuItem, Me.PatientAndProcessedDateToolStripMenuItem, Me.ProcessedDateAndDraftToolStripMenuItem, Me.ProcessedDateOnlyToolStripMenuItem})
        Me.MHIShortOptionToolStripMenuItem.Name = "MHIShortOptionToolStripMenuItem"
        Me.MHIShortOptionToolStripMenuItem.Size = New System.Drawing.Size(136, 24)
        Me.MHIShortOptionToolStripMenuItem.Text = "MHI_Sort_Option"
        '
        'PaitentNameAndDOSToolStripMenuItem
        '
        Me.PaitentNameAndDOSToolStripMenuItem.Name = "PaitentNameAndDOSToolStripMenuItem"
        Me.PaitentNameAndDOSToolStripMenuItem.Size = New System.Drawing.Size(272, 26)
        Me.PaitentNameAndDOSToolStripMenuItem.Text = "Paitent Name and DOS"
        '
        'SortByICNToolStripMenuItem
        '
        Me.SortByICNToolStripMenuItem.Name = "SortByICNToolStripMenuItem"
        Me.SortByICNToolStripMenuItem.Size = New System.Drawing.Size(272, 26)
        Me.SortByICNToolStripMenuItem.Text = "Sort by ICN"
        '
        'ProviderTinAndSuffixToolStripMenuItem
        '
        Me.ProviderTinAndSuffixToolStripMenuItem.Name = "ProviderTinAndSuffixToolStripMenuItem"
        Me.ProviderTinAndSuffixToolStripMenuItem.Size = New System.Drawing.Size(272, 26)
        Me.ProviderTinAndSuffixToolStripMenuItem.Text = "Provider Tin and Suffix"
        '
        'DateOfServiceToolStripMenuItem
        '
        Me.DateOfServiceToolStripMenuItem.Name = "DateOfServiceToolStripMenuItem"
        Me.DateOfServiceToolStripMenuItem.Size = New System.Drawing.Size(272, 26)
        Me.DateOfServiceToolStripMenuItem.Text = "Date of Service"
        '
        'ProviderTinToolStripMenuItem
        '
        Me.ProviderTinToolStripMenuItem.Name = "ProviderTinToolStripMenuItem"
        Me.ProviderTinToolStripMenuItem.Size = New System.Drawing.Size(272, 26)
        Me.ProviderTinToolStripMenuItem.Text = "Provider Tin"
        '
        'DeductibleIndicatorToolStripMenuItem
        '
        Me.DeductibleIndicatorToolStripMenuItem.Name = "DeductibleIndicatorToolStripMenuItem"
        Me.DeductibleIndicatorToolStripMenuItem.Size = New System.Drawing.Size(272, 26)
        Me.DeductibleIndicatorToolStripMenuItem.Text = "Deductible Indicator"
        '
        'SortByPercentToolStripMenuItem
        '
        Me.SortByPercentToolStripMenuItem.Name = "SortByPercentToolStripMenuItem"
        Me.SortByPercentToolStripMenuItem.Size = New System.Drawing.Size(272, 26)
        Me.SortByPercentToolStripMenuItem.Text = "Sort by Percent"
        '
        'PatientAndProcessedDateToolStripMenuItem
        '
        Me.PatientAndProcessedDateToolStripMenuItem.Name = "PatientAndProcessedDateToolStripMenuItem"
        Me.PatientAndProcessedDateToolStripMenuItem.Size = New System.Drawing.Size(272, 26)
        Me.PatientAndProcessedDateToolStripMenuItem.Text = "Patient and Processed Date"
        '
        'ProcessedDateAndDraftToolStripMenuItem
        '
        Me.ProcessedDateAndDraftToolStripMenuItem.Name = "ProcessedDateAndDraftToolStripMenuItem"
        Me.ProcessedDateAndDraftToolStripMenuItem.Size = New System.Drawing.Size(272, 26)
        Me.ProcessedDateAndDraftToolStripMenuItem.Text = "Processed Date and Draft"
        '
        'ProcessedDateOnlyToolStripMenuItem
        '
        Me.ProcessedDateOnlyToolStripMenuItem.Name = "ProcessedDateOnlyToolStripMenuItem"
        Me.ProcessedDateOnlyToolStripMenuItem.Size = New System.Drawing.Size(272, 26)
        Me.ProcessedDateOnlyToolStripMenuItem.Text = "Processed Date Only"
        '
        'ELGSLetterToolStripMenuItem
        '
        Me.ELGSLetterToolStripMenuItem.Name = "ELGSLetterToolStripMenuItem"
        Me.ELGSLetterToolStripMenuItem.Size = New System.Drawing.Size(98, 24)
        Me.ELGSLetterToolStripMenuItem.Text = "ELGS Letter"
        '
        'FormatMHIToolStripMenuItem
        '
        Me.FormatMHIToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.CleanClaimToolStripMenuItem, Me.FormatMHISheetToolStripMenuItem})
        Me.FormatMHIToolStripMenuItem.Name = "FormatMHIToolStripMenuItem"
        Me.FormatMHIToolStripMenuItem.Size = New System.Drawing.Size(102, 24)
        Me.FormatMHIToolStripMenuItem.Text = "Format MHI"
        '
        'CleanClaimToolStripMenuItem
        '
        Me.CleanClaimToolStripMenuItem.Name = "CleanClaimToolStripMenuItem"
        Me.CleanClaimToolStripMenuItem.Size = New System.Drawing.Size(212, 26)
        Me.CleanClaimToolStripMenuItem.Text = "Clean Claim"
        '
        'FormatMHISheetToolStripMenuItem
        '
        Me.FormatMHISheetToolStripMenuItem.Name = "FormatMHISheetToolStripMenuItem"
        Me.FormatMHISheetToolStripMenuItem.Size = New System.Drawing.Size(212, 26)
        Me.FormatMHISheetToolStripMenuItem.Text = "Format MHI Sheet"
        '
        'InstructionsToolStripMenuItem
        '
        Me.InstructionsToolStripMenuItem.Name = "InstructionsToolStripMenuItem"
        Me.InstructionsToolStripMenuItem.Size = New System.Drawing.Size(98, 24)
        Me.InstructionsToolStripMenuItem.Text = "Instructions"
        '
        'AboutToolStripMenuItem
        '
        Me.AboutToolStripMenuItem.Name = "AboutToolStripMenuItem"
        Me.AboutToolStripMenuItem.Size = New System.Drawing.Size(64, 24)
        Me.AboutToolStripMenuItem.Text = "About"
        '
        'ClearAllFilterToolStripMenuItem
        '
        Me.ClearAllFilterToolStripMenuItem.Name = "ClearAllFilterToolStripMenuItem"
        Me.ClearAllFilterToolStripMenuItem.Size = New System.Drawing.Size(120, 24)
        Me.ClearAllFilterToolStripMenuItem.Text = "Clear_All_Filter"
        '
        'ApplyColorToolStripMenuItem
        '
        Me.ApplyColorToolStripMenuItem.Name = "ApplyColorToolStripMenuItem"
        Me.ApplyColorToolStripMenuItem.Size = New System.Drawing.Size(104, 24)
        Me.ApplyColorToolStripMenuItem.Text = "Apply_Color"
        '
        'GroupBox1
        '
        Me.GroupBox1.BackColor = System.Drawing.Color.FromArgb(CType(CType(165, Byte), Integer), CType(CType(210, Byte), Integer), CType(CType(224, Byte), Integer))
        Me.GroupBox1.Controls.Add(Me.Chk_select_memlist)
        Me.GroupBox1.Controls.Add(Me.yearList)
        Me.GroupBox1.Controls.Add(Me.btnClaimInfo)
        Me.GroupBox1.Controls.Add(Me.btnMain)
        Me.GroupBox1.Controls.Add(Me.endSelect)
        Me.GroupBox1.Controls.Add(Me.startSelect)
        Me.GroupBox1.Controls.Add(Me.txt_SSN)
        Me.GroupBox1.Controls.Add(Me.txt_Policy)
        Me.GroupBox1.Controls.Add(Me.Label5)
        Me.GroupBox1.Controls.Add(Me.Label4)
        Me.GroupBox1.Controls.Add(Me.Label3)
        Me.GroupBox1.Controls.Add(Me.memberList)
        Me.GroupBox1.Font = New System.Drawing.Font("Microsoft Sans Serif", 11.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupBox1.Location = New System.Drawing.Point(40, 202)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(304, 450)
        Me.GroupBox1.TabIndex = 1
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Member Info"
        '
        'Chk_select_memlist
        '
        Me.Chk_select_memlist.AutoSize = True
        Me.Chk_select_memlist.Location = New System.Drawing.Point(10, 253)
        Me.Chk_select_memlist.Name = "Chk_select_memlist"
        Me.Chk_select_memlist.Size = New System.Drawing.Size(110, 28)
        Me.Chk_select_memlist.TabIndex = 13
        Me.Chk_select_memlist.Text = "Select All"
        Me.Chk_select_memlist.UseVisualStyleBackColor = True
        '
        'yearList
        '
        Me.yearList.FormattingEnabled = True
        Me.yearList.Location = New System.Drawing.Point(137, 133)
        Me.yearList.Name = "yearList"
        Me.yearList.Size = New System.Drawing.Size(136, 30)
        Me.yearList.TabIndex = 12
        '
        'btnClaimInfo
        '
        Me.btnClaimInfo.Location = New System.Drawing.Point(53, 410)
        Me.btnClaimInfo.Name = "btnClaimInfo"
        Me.btnClaimInfo.Size = New System.Drawing.Size(205, 32)
        Me.btnClaimInfo.TabIndex = 11
        Me.btnClaimInfo.Text = "Get History"
        Me.btnClaimInfo.UseVisualStyleBackColor = True
        '
        'btnMain
        '
        Me.btnMain.Location = New System.Drawing.Point(45, 215)
        Me.btnMain.Name = "btnMain"
        Me.btnMain.Size = New System.Drawing.Size(213, 32)
        Me.btnMain.TabIndex = 10
        Me.btnMain.Text = "Gather Account Info"
        Me.btnMain.UseVisualStyleBackColor = True
        '
        'endSelect
        '
        Me.endSelect.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.endSelect.Location = New System.Drawing.Point(150, 180)
        Me.endSelect.Name = "endSelect"
        Me.endSelect.Size = New System.Drawing.Size(131, 28)
        Me.endSelect.TabIndex = 9
        '
        'startSelect
        '
        Me.startSelect.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.startSelect.Location = New System.Drawing.Point(6, 180)
        Me.startSelect.Name = "startSelect"
        Me.startSelect.Size = New System.Drawing.Size(131, 28)
        Me.startSelect.TabIndex = 8
        '
        'txt_SSN
        '
        Me.txt_SSN.Location = New System.Drawing.Point(142, 84)
        Me.txt_SSN.Name = "txt_SSN"
        Me.txt_SSN.Size = New System.Drawing.Size(131, 28)
        Me.txt_SSN.TabIndex = 6
        '
        'txt_Policy
        '
        Me.txt_Policy.Location = New System.Drawing.Point(142, 39)
        Me.txt_Policy.Name = "txt_Policy"
        Me.txt_Policy.Size = New System.Drawing.Size(131, 28)
        Me.txt_Policy.TabIndex = 5
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label5.Location = New System.Drawing.Point(17, 139)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(38, 18)
        Me.Label5.TabIndex = 4
        Me.Label5.Text = "Year"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.Location = New System.Drawing.Point(17, 90)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(81, 18)
        Me.Label4.TabIndex = 3
        Me.Label4.Text = "Member ID"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.Location = New System.Drawing.Point(17, 39)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(66, 18)
        Me.Label3.TabIndex = 2
        Me.Label3.Text = "Policy ID"
        '
        'memberList
        '
        Me.memberList.FormattingEnabled = True
        Me.memberList.Location = New System.Drawing.Point(6, 286)
        Me.memberList.Name = "memberList"
        Me.memberList.Size = New System.Drawing.Size(275, 119)
        Me.memberList.TabIndex = 0
        '
        'GroupBox2
        '
        Me.GroupBox2.BackColor = System.Drawing.Color.FromArgb(CType(CType(165, Byte), Integer), CType(CType(210, Byte), Integer), CType(CType(224, Byte), Integer))
        Me.GroupBox2.Controls.Add(Me.txtUserPass)
        Me.GroupBox2.Controls.Add(Me.txtUserID)
        Me.GroupBox2.Controls.Add(Me.Label2)
        Me.GroupBox2.Controls.Add(Me.Label1)
        Me.GroupBox2.Font = New System.Drawing.Font("Microsoft Sans Serif", 11.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupBox2.Location = New System.Drawing.Point(40, 75)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(304, 121)
        Me.GroupBox2.TabIndex = 8
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Unet Credentials"
        '
        'txtUserPass
        '
        Me.txtUserPass.Location = New System.Drawing.Point(142, 85)
        Me.txtUserPass.Name = "txtUserPass"
        Me.txtUserPass.PasswordChar = Global.Microsoft.VisualBasic.ChrW(42)
        Me.txtUserPass.Size = New System.Drawing.Size(131, 28)
        Me.txtUserPass.TabIndex = 4
        '
        'txtUserID
        '
        Me.txtUserID.Location = New System.Drawing.Point(142, 44)
        Me.txtUserID.Name = "txtUserID"
        Me.txtUserID.Size = New System.Drawing.Size(131, 28)
        Me.txtUserID.TabIndex = 3
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.Location = New System.Drawing.Point(17, 91)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(75, 18)
        Me.Label2.TabIndex = 2
        Me.Label2.Text = "Password"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(17, 50)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(58, 18)
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "User ID"
        '
        'Panel1
        '
        Me.Panel1.BackColor = System.Drawing.Color.FromArgb(CType(CType(50, Byte), Integer), CType(CType(60, Byte), Integer), CType(CType(100, Byte), Integer))
        Me.Panel1.Location = New System.Drawing.Point(12, 31)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(1701, 21)
        Me.Panel1.TabIndex = 3
        '
        'Panel2
        '
        Me.Panel2.BackColor = System.Drawing.Color.FromArgb(CType(CType(50, Byte), Integer), CType(CType(60, Byte), Integer), CType(CType(100, Byte), Integer))
        Me.Panel2.Location = New System.Drawing.Point(12, 53)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(22, 822)
        Me.Panel2.TabIndex = 4
        '
        'RichTextBox1
        '
        Me.RichTextBox1.BackColor = System.Drawing.Color.FromArgb(CType(CType(165, Byte), Integer), CType(CType(210, Byte), Integer), CType(CType(224, Byte), Integer))
        Me.RichTextBox1.Location = New System.Drawing.Point(46, 678)
        Me.RichTextBox1.Name = "RichTextBox1"
        Me.RichTextBox1.ScrollBars = System.Windows.Forms.RichTextBoxScrollBars.ForcedVertical
        Me.RichTextBox1.Size = New System.Drawing.Size(298, 176)
        Me.RichTextBox1.TabIndex = 13
        Me.RichTextBox1.Text = ""
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.BackColor = System.Drawing.Color.White
        Me.Label6.Location = New System.Drawing.Point(43, 655)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(64, 16)
        Me.Label6.TabIndex = 14
        Me.Label6.Text = "Summary"
        '
        'TabPage13
        '
        Me.TabPage13.Controls.Add(Me.DGridMHI)
        Me.TabPage13.Location = New System.Drawing.Point(4, 28)
        Me.TabPage13.Name = "TabPage13"
        Me.TabPage13.Size = New System.Drawing.Size(1278, 744)
        Me.TabPage13.TabIndex = 13
        Me.TabPage13.Text = "MHITab"
        Me.TabPage13.UseVisualStyleBackColor = True
        '
        'DGridMHI
        '
        Me.DGridMHI.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DGridMHI.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.From, Me.Thru, Me.Svc, Me.PS, Me.Nbr, Me.OV, Me.P, Me.N, Me.RC, Me.Charge, Me.NotCov, Me.BM, Me.Covered, Me.Deduct, Me.D, Me.Perc, Me.Paid, Me.S, Me.DC, Me.Sanc, Me.CauseCode, Me.P1, Me.Tin, Me.Suffix, Me.ClaimNumber, Me.Draft, Me.ProcDate, Me.AdjNo, Me.TotalBilled, Me.TotalPaid, Me.ICN, Me.Suf, Me.FLN, Me.PRS, Me.SI, Me.PT_Name, Me.Blank, Me.PTRel, Me.PTName, Me.INN_Ded, Me.INN_OOP, Me.OON_Ded, Me.OON_OOP, Me.INNDed, Me.INNOOP, Me.OONDed, Me.OONOOP, Me.OI_OIM, Me.OOPCalcRun, Me.ICNandSuffix, Me.Inp_Facility, Me.ProviderName, Me.ProviderType, Me.M1, Me.M2, Me.M3, Me.M4})
        Me.DGridMHI.FilterAndSortEnabled = True
        Me.DGridMHI.FilterStringChangedInvokeBeforeDatasourceUpdate = True
        Me.DGridMHI.Location = New System.Drawing.Point(3, 0)
        Me.DGridMHI.Name = "DGridMHI"
        Me.DGridMHI.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.DGridMHI.RowHeadersWidth = 51
        Me.DGridMHI.RowTemplate.Height = 24
        Me.DGridMHI.Size = New System.Drawing.Size(1216, 579)
        Me.DGridMHI.SortStringChangedInvokeBeforeDatasourceUpdate = True
        Me.DGridMHI.TabIndex = 0
        '
        'From
        '
        Me.From.HeaderText = "From"
        Me.From.MinimumWidth = 22
        Me.From.Name = "From"
        Me.From.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Programmatic
        Me.From.Width = 125
        '
        'Thru
        '
        Me.Thru.HeaderText = "Thru"
        Me.Thru.MinimumWidth = 22
        Me.Thru.Name = "Thru"
        Me.Thru.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Programmatic
        Me.Thru.Width = 125
        '
        'Svc
        '
        Me.Svc.HeaderText = "Svc"
        Me.Svc.MinimumWidth = 22
        Me.Svc.Name = "Svc"
        Me.Svc.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Programmatic
        Me.Svc.Width = 125
        '
        'PS
        '
        Me.PS.HeaderText = "PS"
        Me.PS.MinimumWidth = 22
        Me.PS.Name = "PS"
        Me.PS.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Programmatic
        Me.PS.Width = 125
        '
        'Nbr
        '
        Me.Nbr.HeaderText = "Nbr"
        Me.Nbr.MinimumWidth = 22
        Me.Nbr.Name = "Nbr"
        Me.Nbr.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Programmatic
        Me.Nbr.Width = 125
        '
        'OV
        '
        Me.OV.HeaderText = "OV"
        Me.OV.MinimumWidth = 22
        Me.OV.Name = "OV"
        Me.OV.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Programmatic
        Me.OV.Width = 125
        '
        'P
        '
        Me.P.HeaderText = "P"
        Me.P.MinimumWidth = 22
        Me.P.Name = "P"
        Me.P.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Programmatic
        Me.P.Width = 125
        '
        'N
        '
        Me.N.HeaderText = "N"
        Me.N.MinimumWidth = 22
        Me.N.Name = "N"
        Me.N.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Programmatic
        Me.N.Width = 125
        '
        'RC
        '
        Me.RC.HeaderText = "RC"
        Me.RC.MinimumWidth = 22
        Me.RC.Name = "RC"
        Me.RC.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Programmatic
        Me.RC.Width = 125
        '
        'Charge
        '
        Me.Charge.HeaderText = "Charge"
        Me.Charge.MinimumWidth = 22
        Me.Charge.Name = "Charge"
        Me.Charge.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Programmatic
        Me.Charge.Width = 125
        '
        'NotCov
        '
        Me.NotCov.HeaderText = "NotCov"
        Me.NotCov.MinimumWidth = 22
        Me.NotCov.Name = "NotCov"
        Me.NotCov.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Programmatic
        Me.NotCov.Width = 125
        '
        'BM
        '
        Me.BM.HeaderText = "BM"
        Me.BM.MinimumWidth = 22
        Me.BM.Name = "BM"
        Me.BM.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Programmatic
        Me.BM.Width = 125
        '
        'Covered
        '
        Me.Covered.HeaderText = "Covered"
        Me.Covered.MinimumWidth = 22
        Me.Covered.Name = "Covered"
        Me.Covered.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Programmatic
        Me.Covered.Width = 125
        '
        'Deduct
        '
        Me.Deduct.HeaderText = "Deduct"
        Me.Deduct.MinimumWidth = 22
        Me.Deduct.Name = "Deduct"
        Me.Deduct.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Programmatic
        Me.Deduct.Width = 125
        '
        'D
        '
        Me.D.HeaderText = "D"
        Me.D.MinimumWidth = 22
        Me.D.Name = "D"
        Me.D.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Programmatic
        Me.D.Width = 125
        '
        'Perc
        '
        Me.Perc.HeaderText = "Perc"
        Me.Perc.MinimumWidth = 22
        Me.Perc.Name = "Perc"
        Me.Perc.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Programmatic
        Me.Perc.Width = 125
        '
        'Paid
        '
        Me.Paid.HeaderText = "Paid"
        Me.Paid.MinimumWidth = 22
        Me.Paid.Name = "Paid"
        Me.Paid.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Programmatic
        Me.Paid.Width = 125
        '
        'S
        '
        Me.S.HeaderText = "S"
        Me.S.MinimumWidth = 22
        Me.S.Name = "S"
        Me.S.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Programmatic
        Me.S.Width = 125
        '
        'DC
        '
        Me.DC.HeaderText = "DC"
        Me.DC.MinimumWidth = 22
        Me.DC.Name = "DC"
        Me.DC.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Programmatic
        Me.DC.Width = 125
        '
        'Sanc
        '
        Me.Sanc.HeaderText = "Sanc"
        Me.Sanc.MinimumWidth = 22
        Me.Sanc.Name = "Sanc"
        Me.Sanc.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Programmatic
        Me.Sanc.Width = 125
        '
        'CauseCode
        '
        Me.CauseCode.HeaderText = "CauseCode"
        Me.CauseCode.MinimumWidth = 22
        Me.CauseCode.Name = "CauseCode"
        Me.CauseCode.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Programmatic
        Me.CauseCode.Width = 125
        '
        'P1
        '
        Me.P1.HeaderText = "P1"
        Me.P1.MinimumWidth = 22
        Me.P1.Name = "P1"
        Me.P1.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Programmatic
        Me.P1.Width = 125
        '
        'Tin
        '
        Me.Tin.HeaderText = "Tin"
        Me.Tin.MinimumWidth = 22
        Me.Tin.Name = "Tin"
        Me.Tin.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Programmatic
        Me.Tin.Width = 125
        '
        'Suffix
        '
        Me.Suffix.HeaderText = "Suffix"
        Me.Suffix.MinimumWidth = 22
        Me.Suffix.Name = "Suffix"
        Me.Suffix.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Programmatic
        Me.Suffix.Width = 125
        '
        'ClaimNumber
        '
        Me.ClaimNumber.HeaderText = "ClaimNumber"
        Me.ClaimNumber.MinimumWidth = 22
        Me.ClaimNumber.Name = "ClaimNumber"
        Me.ClaimNumber.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Programmatic
        Me.ClaimNumber.Width = 125
        '
        'Draft
        '
        Me.Draft.HeaderText = "Draft"
        Me.Draft.MinimumWidth = 22
        Me.Draft.Name = "Draft"
        Me.Draft.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Programmatic
        Me.Draft.Width = 125
        '
        'ProcDate
        '
        Me.ProcDate.HeaderText = "ProcDate"
        Me.ProcDate.MinimumWidth = 22
        Me.ProcDate.Name = "ProcDate"
        Me.ProcDate.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Programmatic
        Me.ProcDate.Width = 125
        '
        'AdjNo
        '
        Me.AdjNo.HeaderText = "AdjNo"
        Me.AdjNo.MinimumWidth = 22
        Me.AdjNo.Name = "AdjNo"
        Me.AdjNo.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Programmatic
        Me.AdjNo.Width = 125
        '
        'TotalBilled
        '
        Me.TotalBilled.HeaderText = "TotalBilled"
        Me.TotalBilled.MinimumWidth = 22
        Me.TotalBilled.Name = "TotalBilled"
        Me.TotalBilled.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Programmatic
        Me.TotalBilled.Width = 125
        '
        'TotalPaid
        '
        Me.TotalPaid.HeaderText = "TotalPaid"
        Me.TotalPaid.MinimumWidth = 22
        Me.TotalPaid.Name = "TotalPaid"
        Me.TotalPaid.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Programmatic
        Me.TotalPaid.Width = 125
        '
        'ICN
        '
        Me.ICN.HeaderText = "ICN"
        Me.ICN.MinimumWidth = 22
        Me.ICN.Name = "ICN"
        Me.ICN.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Programmatic
        Me.ICN.Width = 125
        '
        'Suf
        '
        Me.Suf.HeaderText = "Suf"
        Me.Suf.MinimumWidth = 22
        Me.Suf.Name = "Suf"
        Me.Suf.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Programmatic
        Me.Suf.Width = 125
        '
        'FLN
        '
        Me.FLN.HeaderText = "FLN"
        Me.FLN.MinimumWidth = 22
        Me.FLN.Name = "FLN"
        Me.FLN.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Programmatic
        Me.FLN.Width = 125
        '
        'PRS
        '
        Me.PRS.HeaderText = "PRS"
        Me.PRS.MinimumWidth = 22
        Me.PRS.Name = "PRS"
        Me.PRS.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Programmatic
        Me.PRS.Width = 125
        '
        'SI
        '
        Me.SI.HeaderText = "SI"
        Me.SI.MinimumWidth = 22
        Me.SI.Name = "SI"
        Me.SI.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Programmatic
        Me.SI.Width = 125
        '
        'PT_Name
        '
        Me.PT_Name.HeaderText = "PT_Name"
        Me.PT_Name.MinimumWidth = 22
        Me.PT_Name.Name = "PT_Name"
        Me.PT_Name.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Programmatic
        Me.PT_Name.Width = 125
        '
        'Blank
        '
        Me.Blank.HeaderText = "Blank"
        Me.Blank.MinimumWidth = 22
        Me.Blank.Name = "Blank"
        Me.Blank.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Programmatic
        Me.Blank.Width = 125
        '
        'PTRel
        '
        Me.PTRel.HeaderText = "PTRel"
        Me.PTRel.MinimumWidth = 22
        Me.PTRel.Name = "PTRel"
        Me.PTRel.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Programmatic
        Me.PTRel.Width = 125
        '
        'PTName
        '
        Me.PTName.HeaderText = "PTName"
        Me.PTName.MinimumWidth = 22
        Me.PTName.Name = "PTName"
        Me.PTName.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Programmatic
        Me.PTName.Width = 125
        '
        'INN_Ded
        '
        Me.INN_Ded.HeaderText = "INN_Ded"
        Me.INN_Ded.MinimumWidth = 22
        Me.INN_Ded.Name = "INN_Ded"
        Me.INN_Ded.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Programmatic
        Me.INN_Ded.Width = 125
        '
        'INN_OOP
        '
        Me.INN_OOP.HeaderText = "INN_OOP"
        Me.INN_OOP.MinimumWidth = 22
        Me.INN_OOP.Name = "INN_OOP"
        Me.INN_OOP.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Programmatic
        Me.INN_OOP.Width = 125
        '
        'OON_Ded
        '
        Me.OON_Ded.HeaderText = "OON_Ded"
        Me.OON_Ded.MinimumWidth = 22
        Me.OON_Ded.Name = "OON_Ded"
        Me.OON_Ded.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Programmatic
        Me.OON_Ded.Width = 125
        '
        'OON_OOP
        '
        Me.OON_OOP.HeaderText = "OON_OOP"
        Me.OON_OOP.MinimumWidth = 22
        Me.OON_OOP.Name = "OON_OOP"
        Me.OON_OOP.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Programmatic
        Me.OON_OOP.Width = 125
        '
        'INNDed
        '
        Me.INNDed.HeaderText = "INNDed"
        Me.INNDed.MinimumWidth = 22
        Me.INNDed.Name = "INNDed"
        Me.INNDed.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Programmatic
        Me.INNDed.Width = 125
        '
        'INNOOP
        '
        Me.INNOOP.HeaderText = "INNOOP"
        Me.INNOOP.MinimumWidth = 22
        Me.INNOOP.Name = "INNOOP"
        Me.INNOOP.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Programmatic
        Me.INNOOP.Width = 125
        '
        'OONDed
        '
        Me.OONDed.HeaderText = "OONDed"
        Me.OONDed.MinimumWidth = 22
        Me.OONDed.Name = "OONDed"
        Me.OONDed.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Programmatic
        Me.OONDed.Width = 125
        '
        'OONOOP
        '
        Me.OONOOP.HeaderText = "OONOOP"
        Me.OONOOP.MinimumWidth = 22
        Me.OONOOP.Name = "OONOOP"
        Me.OONOOP.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Programmatic
        Me.OONOOP.Width = 125
        '
        'OI_OIM
        '
        Me.OI_OIM.HeaderText = "OI_OIM"
        Me.OI_OIM.MinimumWidth = 22
        Me.OI_OIM.Name = "OI_OIM"
        Me.OI_OIM.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Programmatic
        Me.OI_OIM.Width = 125
        '
        'OOPCalcRun
        '
        Me.OOPCalcRun.HeaderText = "OOPCalcRun"
        Me.OOPCalcRun.MinimumWidth = 22
        Me.OOPCalcRun.Name = "OOPCalcRun"
        Me.OOPCalcRun.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Programmatic
        Me.OOPCalcRun.Width = 125
        '
        'ICNandSuffix
        '
        Me.ICNandSuffix.HeaderText = "ICNandSuffix"
        Me.ICNandSuffix.MinimumWidth = 22
        Me.ICNandSuffix.Name = "ICNandSuffix"
        Me.ICNandSuffix.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Programmatic
        Me.ICNandSuffix.Width = 125
        '
        'Inp_Facility
        '
        Me.Inp_Facility.HeaderText = "Inp_Facility"
        Me.Inp_Facility.MinimumWidth = 22
        Me.Inp_Facility.Name = "Inp_Facility"
        Me.Inp_Facility.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Programmatic
        Me.Inp_Facility.Width = 125
        '
        'ProviderName
        '
        Me.ProviderName.HeaderText = "ProviderName"
        Me.ProviderName.MinimumWidth = 22
        Me.ProviderName.Name = "ProviderName"
        Me.ProviderName.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Programmatic
        Me.ProviderName.Width = 125
        '
        'ProviderType
        '
        Me.ProviderType.HeaderText = "ProviderType"
        Me.ProviderType.MinimumWidth = 22
        Me.ProviderType.Name = "ProviderType"
        Me.ProviderType.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Programmatic
        Me.ProviderType.Width = 125
        '
        'M1
        '
        Me.M1.HeaderText = "M1"
        Me.M1.MinimumWidth = 22
        Me.M1.Name = "M1"
        Me.M1.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Programmatic
        Me.M1.Width = 125
        '
        'M2
        '
        Me.M2.HeaderText = "M2"
        Me.M2.MinimumWidth = 22
        Me.M2.Name = "M2"
        Me.M2.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Programmatic
        Me.M2.Width = 125
        '
        'M3
        '
        Me.M3.HeaderText = "M3"
        Me.M3.MinimumWidth = 22
        Me.M3.Name = "M3"
        Me.M3.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Programmatic
        Me.M3.Width = 125
        '
        'M4
        '
        Me.M4.HeaderText = "M4"
        Me.M4.MinimumWidth = 22
        Me.M4.Name = "M4"
        Me.M4.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Programmatic
        Me.M4.Width = 125
        '
        'TabPage11
        '
        Me.TabPage11.Controls.Add(Me.DGridADJ)
        Me.TabPage11.Location = New System.Drawing.Point(4, 28)
        Me.TabPage11.Name = "TabPage11"
        Me.TabPage11.Size = New System.Drawing.Size(1278, 744)
        Me.TabPage11.TabIndex = 11
        Me.TabPage11.Text = "Vend Adj"
        Me.TabPage11.UseVisualStyleBackColor = True
        '
        'DGridADJ
        '
        Me.DGridADJ.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DGridADJ.Location = New System.Drawing.Point(-4, 0)
        Me.DGridADJ.Name = "DGridADJ"
        Me.DGridADJ.RowHeadersWidth = 51
        Me.DGridADJ.RowTemplate.Height = 24
        Me.DGridADJ.Size = New System.Drawing.Size(1215, 492)
        Me.DGridADJ.TabIndex = 0
        '
        'TabPage12
        '
        Me.TabPage12.Controls.Add(Me.btnOOPExport)
        Me.TabPage12.Controls.Add(Me.tblOOP)
        Me.TabPage12.Location = New System.Drawing.Point(4, 28)
        Me.TabPage12.Name = "TabPage12"
        Me.TabPage12.Size = New System.Drawing.Size(1278, 744)
        Me.TabPage12.TabIndex = 12
        Me.TabPage12.Text = "OOP_Spreadsheet"
        Me.TabPage12.UseVisualStyleBackColor = True
        '
        'btnOOPExport
        '
        Me.btnOOPExport.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnOOPExport.Location = New System.Drawing.Point(242, 697)
        Me.btnOOPExport.Name = "btnOOPExport"
        Me.btnOOPExport.Size = New System.Drawing.Size(198, 35)
        Me.btnOOPExport.TabIndex = 1
        Me.btnOOPExport.Text = "Export"
        Me.btnOOPExport.UseVisualStyleBackColor = True
        '
        'tblOOP
        '
        Me.tblOOP.BackgroundColor = System.Drawing.SystemColors.ActiveCaption
        Me.tblOOP.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.tblOOP.Location = New System.Drawing.Point(3, 0)
        Me.tblOOP.Name = "tblOOP"
        Me.tblOOP.RowHeadersWidth = 51
        Me.tblOOP.RowTemplate.Height = 24
        Me.tblOOP.Size = New System.Drawing.Size(1272, 685)
        Me.tblOOP.TabIndex = 0
        '
        'TabPage10
        '
        Me.TabPage10.Controls.Add(Me.tblCopay)
        Me.TabPage10.Location = New System.Drawing.Point(4, 28)
        Me.TabPage10.Name = "TabPage10"
        Me.TabPage10.Size = New System.Drawing.Size(1278, 744)
        Me.TabPage10.TabIndex = 10
        Me.TabPage10.Text = "Copays"
        Me.TabPage10.UseVisualStyleBackColor = True
        '
        'tblCopay
        '
        Me.tblCopay.BackgroundColor = System.Drawing.SystemColors.ActiveCaption
        Me.tblCopay.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.tblCopay.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.Policy, Me.Plan_CodeReportingCode_PlanVar, Me.Year, Me.PatientName, Me.CauseCodes, Me.PlaceofService, Me.MajorMedCalcCode, Me.SpecialProcCode, Me.CopaySet, Me.DataGridViewTextBoxColumn110})
        Me.tblCopay.Location = New System.Drawing.Point(3, 3)
        Me.tblCopay.Name = "tblCopay"
        Me.tblCopay.RowHeadersWidth = 51
        Me.tblCopay.RowTemplate.Height = 24
        Me.tblCopay.Size = New System.Drawing.Size(1232, 629)
        Me.tblCopay.TabIndex = 0
        '
        'Policy
        '
        Me.Policy.HeaderText = "Policy"
        Me.Policy.MinimumWidth = 6
        Me.Policy.Name = "Policy"
        Me.Policy.Width = 125
        '
        'Plan_CodeReportingCode_PlanVar
        '
        Me.Plan_CodeReportingCode_PlanVar.HeaderText = "Plan_CodeReportingCode_PlanVar"
        Me.Plan_CodeReportingCode_PlanVar.MinimumWidth = 6
        Me.Plan_CodeReportingCode_PlanVar.Name = "Plan_CodeReportingCode_PlanVar"
        Me.Plan_CodeReportingCode_PlanVar.Width = 125
        '
        'Year
        '
        Me.Year.HeaderText = "Year"
        Me.Year.MinimumWidth = 6
        Me.Year.Name = "Year"
        Me.Year.Width = 125
        '
        'PatientName
        '
        Me.PatientName.HeaderText = "PatientName"
        Me.PatientName.MinimumWidth = 6
        Me.PatientName.Name = "PatientName"
        Me.PatientName.Width = 125
        '
        'CauseCodes
        '
        Me.CauseCodes.HeaderText = "CauseCodes"
        Me.CauseCodes.MinimumWidth = 6
        Me.CauseCodes.Name = "CauseCodes"
        Me.CauseCodes.Width = 125
        '
        'PlaceofService
        '
        Me.PlaceofService.HeaderText = "PlaceofService"
        Me.PlaceofService.MinimumWidth = 6
        Me.PlaceofService.Name = "PlaceofService"
        Me.PlaceofService.Width = 125
        '
        'MajorMedCalcCode
        '
        Me.MajorMedCalcCode.HeaderText = "MajorMedCalcCode"
        Me.MajorMedCalcCode.MinimumWidth = 6
        Me.MajorMedCalcCode.Name = "MajorMedCalcCode"
        Me.MajorMedCalcCode.Width = 125
        '
        'SpecialProcCode
        '
        Me.SpecialProcCode.HeaderText = "SpecialProcCode"
        Me.SpecialProcCode.MinimumWidth = 6
        Me.SpecialProcCode.Name = "SpecialProcCode"
        Me.SpecialProcCode.Width = 125
        '
        'CopaySet
        '
        Me.CopaySet.HeaderText = "CopaySet"
        Me.CopaySet.MinimumWidth = 6
        Me.CopaySet.Name = "CopaySet"
        Me.CopaySet.Width = 125
        '
        'DataGridViewTextBoxColumn110
        '
        Me.DataGridViewTextBoxColumn110.HeaderText = "DataGridViewTextBoxColumn110"
        Me.DataGridViewTextBoxColumn110.MinimumWidth = 6
        Me.DataGridViewTextBoxColumn110.Name = "DataGridViewTextBoxColumn110"
        Me.DataGridViewTextBoxColumn110.Width = 125
        '
        'TabPage8
        '
        Me.TabPage8.Controls.Add(Me.DGrid_PG10)
        Me.TabPage8.Location = New System.Drawing.Point(4, 28)
        Me.TabPage8.Name = "TabPage8"
        Me.TabPage8.Size = New System.Drawing.Size(1278, 744)
        Me.TabPage8.TabIndex = 8
        Me.TabPage8.Text = "MMI Page 10"
        Me.TabPage8.UseVisualStyleBackColor = True
        '
        'DGrid_PG10
        '
        Me.DGrid_PG10.BackgroundColor = System.Drawing.SystemColors.ActiveCaption
        Me.DGrid_PG10.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DGrid_PG10.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.DataGridViewTextBoxColumn29, Me.DataGridViewTextBoxColumn30, Me.DataGridViewTextBoxColumn31, Me.DataGridViewTextBoxColumn32, Me.DataGridViewTextBoxColumn33, Me.DataGridViewTextBoxColumn34, Me.DataGridViewTextBoxColumn35, Me.Column_H, Me.Column_I, Me.Column_J, Me.Column_K, Me.Column_L, Me.Column_M, Me.Column_N, Me.Column_O, Me.Column_P})
        Me.DGrid_PG10.Location = New System.Drawing.Point(7, 6)
        Me.DGrid_PG10.Name = "DGrid_PG10"
        Me.DGrid_PG10.RowHeadersWidth = 51
        Me.DGrid_PG10.RowTemplate.Height = 24
        Me.DGrid_PG10.Size = New System.Drawing.Size(1178, 615)
        Me.DGrid_PG10.TabIndex = 3
        '
        'DataGridViewTextBoxColumn29
        '
        Me.DataGridViewTextBoxColumn29.HeaderText = "Col_A"
        Me.DataGridViewTextBoxColumn29.MinimumWidth = 6
        Me.DataGridViewTextBoxColumn29.Name = "DataGridViewTextBoxColumn29"
        Me.DataGridViewTextBoxColumn29.Width = 125
        '
        'DataGridViewTextBoxColumn30
        '
        Me.DataGridViewTextBoxColumn30.HeaderText = "Col_B"
        Me.DataGridViewTextBoxColumn30.MinimumWidth = 6
        Me.DataGridViewTextBoxColumn30.Name = "DataGridViewTextBoxColumn30"
        Me.DataGridViewTextBoxColumn30.Width = 125
        '
        'DataGridViewTextBoxColumn31
        '
        Me.DataGridViewTextBoxColumn31.HeaderText = "Col_C"
        Me.DataGridViewTextBoxColumn31.MinimumWidth = 6
        Me.DataGridViewTextBoxColumn31.Name = "DataGridViewTextBoxColumn31"
        Me.DataGridViewTextBoxColumn31.Width = 125
        '
        'DataGridViewTextBoxColumn32
        '
        Me.DataGridViewTextBoxColumn32.HeaderText = "Col_D"
        Me.DataGridViewTextBoxColumn32.MinimumWidth = 6
        Me.DataGridViewTextBoxColumn32.Name = "DataGridViewTextBoxColumn32"
        Me.DataGridViewTextBoxColumn32.Width = 125
        '
        'DataGridViewTextBoxColumn33
        '
        Me.DataGridViewTextBoxColumn33.HeaderText = "Col_E"
        Me.DataGridViewTextBoxColumn33.MinimumWidth = 6
        Me.DataGridViewTextBoxColumn33.Name = "DataGridViewTextBoxColumn33"
        Me.DataGridViewTextBoxColumn33.Width = 125
        '
        'DataGridViewTextBoxColumn34
        '
        Me.DataGridViewTextBoxColumn34.HeaderText = "Col_F"
        Me.DataGridViewTextBoxColumn34.MinimumWidth = 6
        Me.DataGridViewTextBoxColumn34.Name = "DataGridViewTextBoxColumn34"
        Me.DataGridViewTextBoxColumn34.Width = 125
        '
        'DataGridViewTextBoxColumn35
        '
        Me.DataGridViewTextBoxColumn35.HeaderText = "Col_G"
        Me.DataGridViewTextBoxColumn35.MinimumWidth = 6
        Me.DataGridViewTextBoxColumn35.Name = "DataGridViewTextBoxColumn35"
        Me.DataGridViewTextBoxColumn35.Width = 125
        '
        'Column_H
        '
        Me.Column_H.HeaderText = "Column_H"
        Me.Column_H.MinimumWidth = 6
        Me.Column_H.Name = "Column_H"
        Me.Column_H.Width = 125
        '
        'Column_I
        '
        Me.Column_I.HeaderText = "Column_I"
        Me.Column_I.MinimumWidth = 6
        Me.Column_I.Name = "Column_I"
        Me.Column_I.Width = 125
        '
        'Column_J
        '
        Me.Column_J.HeaderText = "Column_J"
        Me.Column_J.MinimumWidth = 6
        Me.Column_J.Name = "Column_J"
        Me.Column_J.Width = 125
        '
        'Column_K
        '
        Me.Column_K.HeaderText = "Column_K"
        Me.Column_K.MinimumWidth = 6
        Me.Column_K.Name = "Column_K"
        Me.Column_K.Width = 125
        '
        'Column_L
        '
        Me.Column_L.HeaderText = "Column_L"
        Me.Column_L.MinimumWidth = 6
        Me.Column_L.Name = "Column_L"
        Me.Column_L.Width = 125
        '
        'Column_M
        '
        Me.Column_M.HeaderText = "Column_M"
        Me.Column_M.MinimumWidth = 6
        Me.Column_M.Name = "Column_M"
        Me.Column_M.Width = 125
        '
        'Column_N
        '
        Me.Column_N.HeaderText = "Column_N"
        Me.Column_N.MinimumWidth = 6
        Me.Column_N.Name = "Column_N"
        Me.Column_N.Width = 125
        '
        'Column_O
        '
        Me.Column_O.HeaderText = "Column_O"
        Me.Column_O.MinimumWidth = 6
        Me.Column_O.Name = "Column_O"
        Me.Column_O.Width = 125
        '
        'Column_P
        '
        Me.Column_P.HeaderText = "Column_P"
        Me.Column_P.MinimumWidth = 6
        Me.Column_P.Name = "Column_P"
        Me.Column_P.Width = 125
        '
        'TabPage5
        '
        Me.TabPage5.Controls.Add(Me.DGrid_PG5)
        Me.TabPage5.Location = New System.Drawing.Point(4, 28)
        Me.TabPage5.Name = "TabPage5"
        Me.TabPage5.Size = New System.Drawing.Size(1278, 744)
        Me.TabPage5.TabIndex = 5
        Me.TabPage5.Text = "MMI Page 5"
        Me.TabPage5.UseVisualStyleBackColor = True
        '
        'DGrid_PG5
        '
        Me.DGrid_PG5.BackgroundColor = System.Drawing.SystemColors.ActiveCaption
        Me.DGrid_PG5.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DGrid_PG5.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.DataGridViewTextBoxColumn15, Me.DataGridViewTextBoxColumn16, Me.DataGridViewTextBoxColumn17, Me.DataGridViewTextBoxColumn18, Me.DataGridViewTextBoxColumn19, Me.DataGridViewTextBoxColumn20, Me.DataGridViewTextBoxColumn21, Me.Colu_H, Me.Colu_I, Me.Colu_J, Me.Colu_K, Me.Colu_L, Me.Colu_M, Me.Colu_N, Me.Colu_O, Me.Colu_P})
        Me.DGrid_PG5.Location = New System.Drawing.Point(3, 3)
        Me.DGrid_PG5.Name = "DGrid_PG5"
        Me.DGrid_PG5.RowHeadersWidth = 51
        Me.DGrid_PG5.RowTemplate.Height = 24
        Me.DGrid_PG5.Size = New System.Drawing.Size(1179, 615)
        Me.DGrid_PG5.TabIndex = 3
        '
        'DataGridViewTextBoxColumn15
        '
        Me.DataGridViewTextBoxColumn15.HeaderText = "Col_A"
        Me.DataGridViewTextBoxColumn15.MinimumWidth = 6
        Me.DataGridViewTextBoxColumn15.Name = "DataGridViewTextBoxColumn15"
        Me.DataGridViewTextBoxColumn15.Width = 125
        '
        'DataGridViewTextBoxColumn16
        '
        Me.DataGridViewTextBoxColumn16.HeaderText = "Col_B"
        Me.DataGridViewTextBoxColumn16.MinimumWidth = 6
        Me.DataGridViewTextBoxColumn16.Name = "DataGridViewTextBoxColumn16"
        Me.DataGridViewTextBoxColumn16.Width = 125
        '
        'DataGridViewTextBoxColumn17
        '
        Me.DataGridViewTextBoxColumn17.HeaderText = "Col_C"
        Me.DataGridViewTextBoxColumn17.MinimumWidth = 6
        Me.DataGridViewTextBoxColumn17.Name = "DataGridViewTextBoxColumn17"
        Me.DataGridViewTextBoxColumn17.Width = 125
        '
        'DataGridViewTextBoxColumn18
        '
        Me.DataGridViewTextBoxColumn18.HeaderText = "Col_D"
        Me.DataGridViewTextBoxColumn18.MinimumWidth = 6
        Me.DataGridViewTextBoxColumn18.Name = "DataGridViewTextBoxColumn18"
        Me.DataGridViewTextBoxColumn18.Width = 125
        '
        'DataGridViewTextBoxColumn19
        '
        Me.DataGridViewTextBoxColumn19.HeaderText = "Col_E"
        Me.DataGridViewTextBoxColumn19.MinimumWidth = 6
        Me.DataGridViewTextBoxColumn19.Name = "DataGridViewTextBoxColumn19"
        Me.DataGridViewTextBoxColumn19.Width = 125
        '
        'DataGridViewTextBoxColumn20
        '
        Me.DataGridViewTextBoxColumn20.HeaderText = "Col_F"
        Me.DataGridViewTextBoxColumn20.MinimumWidth = 6
        Me.DataGridViewTextBoxColumn20.Name = "DataGridViewTextBoxColumn20"
        Me.DataGridViewTextBoxColumn20.Width = 125
        '
        'DataGridViewTextBoxColumn21
        '
        Me.DataGridViewTextBoxColumn21.HeaderText = "Col_G"
        Me.DataGridViewTextBoxColumn21.MinimumWidth = 6
        Me.DataGridViewTextBoxColumn21.Name = "DataGridViewTextBoxColumn21"
        Me.DataGridViewTextBoxColumn21.Width = 125
        '
        'Colu_H
        '
        Me.Colu_H.HeaderText = "Colu_H"
        Me.Colu_H.MinimumWidth = 6
        Me.Colu_H.Name = "Colu_H"
        Me.Colu_H.Width = 125
        '
        'Colu_I
        '
        Me.Colu_I.HeaderText = "Colu_I"
        Me.Colu_I.MinimumWidth = 6
        Me.Colu_I.Name = "Colu_I"
        Me.Colu_I.Width = 125
        '
        'Colu_J
        '
        Me.Colu_J.HeaderText = "Colu_J"
        Me.Colu_J.MinimumWidth = 6
        Me.Colu_J.Name = "Colu_J"
        Me.Colu_J.Width = 125
        '
        'Colu_K
        '
        Me.Colu_K.HeaderText = "Colu_J"
        Me.Colu_K.MinimumWidth = 6
        Me.Colu_K.Name = "Colu_K"
        Me.Colu_K.Width = 125
        '
        'Colu_L
        '
        Me.Colu_L.HeaderText = "Colu_L"
        Me.Colu_L.MinimumWidth = 6
        Me.Colu_L.Name = "Colu_L"
        Me.Colu_L.Width = 125
        '
        'Colu_M
        '
        Me.Colu_M.HeaderText = "Colu_M"
        Me.Colu_M.MinimumWidth = 6
        Me.Colu_M.Name = "Colu_M"
        Me.Colu_M.Width = 125
        '
        'Colu_N
        '
        Me.Colu_N.HeaderText = "Colu_N"
        Me.Colu_N.MinimumWidth = 6
        Me.Colu_N.Name = "Colu_N"
        Me.Colu_N.Width = 125
        '
        'Colu_O
        '
        Me.Colu_O.HeaderText = "Colu_O"
        Me.Colu_O.MinimumWidth = 6
        Me.Colu_O.Name = "Colu_O"
        Me.Colu_O.Width = 125
        '
        'Colu_P
        '
        Me.Colu_P.HeaderText = "Colu_P"
        Me.Colu_P.MinimumWidth = 6
        Me.Colu_P.Name = "Colu_P"
        Me.Colu_P.Width = 125
        '
        'TabPage4
        '
        Me.TabPage4.Controls.Add(Me.DGrid_PG4)
        Me.TabPage4.Location = New System.Drawing.Point(4, 28)
        Me.TabPage4.Name = "TabPage4"
        Me.TabPage4.Size = New System.Drawing.Size(1278, 744)
        Me.TabPage4.TabIndex = 4
        Me.TabPage4.Text = "MMI Page 4"
        Me.TabPage4.UseVisualStyleBackColor = True
        '
        'DGrid_PG4
        '
        Me.DGrid_PG4.BackgroundColor = System.Drawing.SystemColors.ActiveCaption
        Me.DGrid_PG4.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DGrid_PG4.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.DataGridViewTextBoxColumn8, Me.DataGridViewTextBoxColumn9, Me.DataGridViewTextBoxColumn10, Me.DataGridViewTextBoxColumn11, Me.DataGridViewTextBoxColumn12, Me.DataGridViewTextBoxColumn13, Me.DataGridViewTextBoxColumn14, Me.C_H, Me.C_I, Me.C_J, Me.C_K, Me.C_L, Me.C_M, Me.C_N})
        Me.DGrid_PG4.Location = New System.Drawing.Point(0, 7)
        Me.DGrid_PG4.Name = "DGrid_PG4"
        Me.DGrid_PG4.RowHeadersWidth = 51
        Me.DGrid_PG4.RowTemplate.Height = 24
        Me.DGrid_PG4.Size = New System.Drawing.Size(1179, 615)
        Me.DGrid_PG4.TabIndex = 2
        '
        'DataGridViewTextBoxColumn8
        '
        Me.DataGridViewTextBoxColumn8.HeaderText = "Col_A"
        Me.DataGridViewTextBoxColumn8.MinimumWidth = 6
        Me.DataGridViewTextBoxColumn8.Name = "DataGridViewTextBoxColumn8"
        Me.DataGridViewTextBoxColumn8.Width = 125
        '
        'DataGridViewTextBoxColumn9
        '
        Me.DataGridViewTextBoxColumn9.HeaderText = "Col_B"
        Me.DataGridViewTextBoxColumn9.MinimumWidth = 6
        Me.DataGridViewTextBoxColumn9.Name = "DataGridViewTextBoxColumn9"
        Me.DataGridViewTextBoxColumn9.Width = 125
        '
        'DataGridViewTextBoxColumn10
        '
        Me.DataGridViewTextBoxColumn10.HeaderText = "Col_C"
        Me.DataGridViewTextBoxColumn10.MinimumWidth = 6
        Me.DataGridViewTextBoxColumn10.Name = "DataGridViewTextBoxColumn10"
        Me.DataGridViewTextBoxColumn10.Width = 125
        '
        'DataGridViewTextBoxColumn11
        '
        Me.DataGridViewTextBoxColumn11.HeaderText = "Col_D"
        Me.DataGridViewTextBoxColumn11.MinimumWidth = 6
        Me.DataGridViewTextBoxColumn11.Name = "DataGridViewTextBoxColumn11"
        Me.DataGridViewTextBoxColumn11.Width = 125
        '
        'DataGridViewTextBoxColumn12
        '
        Me.DataGridViewTextBoxColumn12.HeaderText = "Col_E"
        Me.DataGridViewTextBoxColumn12.MinimumWidth = 6
        Me.DataGridViewTextBoxColumn12.Name = "DataGridViewTextBoxColumn12"
        Me.DataGridViewTextBoxColumn12.Width = 125
        '
        'DataGridViewTextBoxColumn13
        '
        Me.DataGridViewTextBoxColumn13.HeaderText = "Col_F"
        Me.DataGridViewTextBoxColumn13.MinimumWidth = 6
        Me.DataGridViewTextBoxColumn13.Name = "DataGridViewTextBoxColumn13"
        Me.DataGridViewTextBoxColumn13.Width = 125
        '
        'DataGridViewTextBoxColumn14
        '
        Me.DataGridViewTextBoxColumn14.HeaderText = "Col_G"
        Me.DataGridViewTextBoxColumn14.MinimumWidth = 6
        Me.DataGridViewTextBoxColumn14.Name = "DataGridViewTextBoxColumn14"
        Me.DataGridViewTextBoxColumn14.Width = 125
        '
        'C_H
        '
        Me.C_H.HeaderText = "C_H"
        Me.C_H.MinimumWidth = 6
        Me.C_H.Name = "C_H"
        Me.C_H.Width = 125
        '
        'C_I
        '
        Me.C_I.HeaderText = "C_I"
        Me.C_I.MinimumWidth = 6
        Me.C_I.Name = "C_I"
        Me.C_I.Width = 125
        '
        'C_J
        '
        Me.C_J.HeaderText = "C_J"
        Me.C_J.MinimumWidth = 6
        Me.C_J.Name = "C_J"
        Me.C_J.Width = 125
        '
        'C_K
        '
        Me.C_K.HeaderText = "C_K"
        Me.C_K.MinimumWidth = 6
        Me.C_K.Name = "C_K"
        Me.C_K.Width = 125
        '
        'C_L
        '
        Me.C_L.HeaderText = "C_L"
        Me.C_L.MinimumWidth = 6
        Me.C_L.Name = "C_L"
        Me.C_L.Width = 125
        '
        'C_M
        '
        Me.C_M.HeaderText = "C_M"
        Me.C_M.MinimumWidth = 6
        Me.C_M.Name = "C_M"
        Me.C_M.Width = 125
        '
        'C_N
        '
        Me.C_N.HeaderText = "C_N"
        Me.C_N.MinimumWidth = 6
        Me.C_N.Name = "C_N"
        Me.C_N.Width = 125
        '
        'TabPage3
        '
        Me.TabPage3.Controls.Add(Me.DGrid_PG1)
        Me.TabPage3.Location = New System.Drawing.Point(4, 28)
        Me.TabPage3.Name = "TabPage3"
        Me.TabPage3.Size = New System.Drawing.Size(1278, 744)
        Me.TabPage3.TabIndex = 3
        Me.TabPage3.Text = "MMI Page 1"
        Me.TabPage3.UseVisualStyleBackColor = True
        '
        'DGrid_PG1
        '
        Me.DGrid_PG1.BackgroundColor = System.Drawing.SystemColors.ActiveCaption
        Me.DGrid_PG1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DGrid_PG1.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.DataGridViewTextBoxColumn1, Me.DataGridViewTextBoxColumn2, Me.DataGridViewTextBoxColumn3, Me.DataGridViewTextBoxColumn4, Me.DataGridViewTextBoxColumn5, Me.DataGridViewTextBoxColumn6, Me.DataGridViewTextBoxColumn7, Me.CL_H, Me.CL_I, Me.CLL_J, Me.CLL_K, Me.CLL_L, Me.CLL_M, Me.CLL_N})
        Me.DGrid_PG1.Location = New System.Drawing.Point(7, 1)
        Me.DGrid_PG1.Name = "DGrid_PG1"
        Me.DGrid_PG1.RowHeadersWidth = 51
        Me.DGrid_PG1.RowTemplate.Height = 24
        Me.DGrid_PG1.Size = New System.Drawing.Size(1179, 615)
        Me.DGrid_PG1.TabIndex = 1
        '
        'DataGridViewTextBoxColumn1
        '
        Me.DataGridViewTextBoxColumn1.HeaderText = "Col_A"
        Me.DataGridViewTextBoxColumn1.MinimumWidth = 6
        Me.DataGridViewTextBoxColumn1.Name = "DataGridViewTextBoxColumn1"
        Me.DataGridViewTextBoxColumn1.Width = 125
        '
        'DataGridViewTextBoxColumn2
        '
        Me.DataGridViewTextBoxColumn2.HeaderText = "Col_B"
        Me.DataGridViewTextBoxColumn2.MinimumWidth = 6
        Me.DataGridViewTextBoxColumn2.Name = "DataGridViewTextBoxColumn2"
        Me.DataGridViewTextBoxColumn2.Width = 125
        '
        'DataGridViewTextBoxColumn3
        '
        Me.DataGridViewTextBoxColumn3.HeaderText = "Col_C"
        Me.DataGridViewTextBoxColumn3.MinimumWidth = 6
        Me.DataGridViewTextBoxColumn3.Name = "DataGridViewTextBoxColumn3"
        Me.DataGridViewTextBoxColumn3.Width = 125
        '
        'DataGridViewTextBoxColumn4
        '
        Me.DataGridViewTextBoxColumn4.HeaderText = "Col_D"
        Me.DataGridViewTextBoxColumn4.MinimumWidth = 6
        Me.DataGridViewTextBoxColumn4.Name = "DataGridViewTextBoxColumn4"
        Me.DataGridViewTextBoxColumn4.Width = 125
        '
        'DataGridViewTextBoxColumn5
        '
        Me.DataGridViewTextBoxColumn5.HeaderText = "Col_E"
        Me.DataGridViewTextBoxColumn5.MinimumWidth = 6
        Me.DataGridViewTextBoxColumn5.Name = "DataGridViewTextBoxColumn5"
        Me.DataGridViewTextBoxColumn5.Width = 125
        '
        'DataGridViewTextBoxColumn6
        '
        Me.DataGridViewTextBoxColumn6.HeaderText = "Col_F"
        Me.DataGridViewTextBoxColumn6.MinimumWidth = 6
        Me.DataGridViewTextBoxColumn6.Name = "DataGridViewTextBoxColumn6"
        Me.DataGridViewTextBoxColumn6.Width = 125
        '
        'DataGridViewTextBoxColumn7
        '
        Me.DataGridViewTextBoxColumn7.HeaderText = "Col_G"
        Me.DataGridViewTextBoxColumn7.MinimumWidth = 6
        Me.DataGridViewTextBoxColumn7.Name = "DataGridViewTextBoxColumn7"
        Me.DataGridViewTextBoxColumn7.Width = 125
        '
        'CL_H
        '
        Me.CL_H.HeaderText = "CL_H"
        Me.CL_H.MinimumWidth = 6
        Me.CL_H.Name = "CL_H"
        Me.CL_H.Width = 125
        '
        'CL_I
        '
        Me.CL_I.HeaderText = "CL_I"
        Me.CL_I.MinimumWidth = 6
        Me.CL_I.Name = "CL_I"
        Me.CL_I.Width = 125
        '
        'CLL_J
        '
        Me.CLL_J.HeaderText = "CLL_J"
        Me.CLL_J.MinimumWidth = 6
        Me.CLL_J.Name = "CLL_J"
        Me.CLL_J.Width = 125
        '
        'CLL_K
        '
        Me.CLL_K.HeaderText = "CLL_K"
        Me.CLL_K.MinimumWidth = 6
        Me.CLL_K.Name = "CLL_K"
        Me.CLL_K.Width = 125
        '
        'CLL_L
        '
        Me.CLL_L.HeaderText = "CLL_L"
        Me.CLL_L.MinimumWidth = 6
        Me.CLL_L.Name = "CLL_L"
        Me.CLL_L.Width = 125
        '
        'CLL_M
        '
        Me.CLL_M.HeaderText = "CLL_M"
        Me.CLL_M.MinimumWidth = 6
        Me.CLL_M.Name = "CLL_M"
        Me.CLL_M.Width = 125
        '
        'CLL_N
        '
        Me.CLL_N.HeaderText = "CLL_N"
        Me.CLL_N.MinimumWidth = 6
        Me.CLL_N.Name = "CLL_N"
        Me.CLL_N.Width = 125
        '
        'TabPage1
        '
        Me.TabPage1.Controls.Add(Me.Button1)
        Me.TabPage1.Controls.Add(Me.MMI_Message)
        Me.TabPage1.Controls.Add(Me.DGridOverview)
        Me.TabPage1.Location = New System.Drawing.Point(4, 28)
        Me.TabPage1.Name = "TabPage1"
        Me.TabPage1.Size = New System.Drawing.Size(1278, 744)
        Me.TabPage1.TabIndex = 2
        Me.TabPage1.Text = "MMI Overview"
        Me.TabPage1.UseVisualStyleBackColor = True
        '
        'Button1
        '
        Me.Button1.BackColor = System.Drawing.SystemColors.Control
        Me.Button1.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button1.Location = New System.Drawing.Point(913, 613)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(257, 67)
        Me.Button1.TabIndex = 2
        Me.Button1.Text = "EXPORT MMI Overview"
        Me.Button1.UseVisualStyleBackColor = False
        '
        'MMI_Message
        '
        Me.MMI_Message.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.MMI_Message.Location = New System.Drawing.Point(3, 586)
        Me.MMI_Message.Name = "MMI_Message"
        Me.MMI_Message.Size = New System.Drawing.Size(855, 139)
        Me.MMI_Message.TabIndex = 1
        Me.MMI_Message.Text = ""
        '
        'DGridOverview
        '
        Me.DGridOverview.BackgroundColor = System.Drawing.SystemColors.ActiveCaption
        Me.DGridOverview.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DGridOverview.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.Col_A, Me.COL_O, Me.COL_P, Me.Col_B, Me.Col_C, Me.Col_D, Me.Col_E, Me.Col_F, Me.Col_G, Me.Col_H, Me.Col_I, Me.Col_J, Me.Cl_J, Me.Col_k, Me.Col_L, Me.Col_M, Me.Col_N})
        Me.DGridOverview.Location = New System.Drawing.Point(3, 0)
        Me.DGridOverview.Name = "DGridOverview"
        Me.DGridOverview.RowHeadersWidth = 51
        Me.DGridOverview.RowTemplate.Height = 24
        Me.DGridOverview.Size = New System.Drawing.Size(1182, 579)
        Me.DGridOverview.TabIndex = 0
        '
        'Col_A
        '
        Me.Col_A.HeaderText = "Col_A"
        Me.Col_A.MinimumWidth = 6
        Me.Col_A.Name = "Col_A"
        Me.Col_A.Width = 125
        '
        'COL_O
        '
        Me.COL_O.HeaderText = "COL_O"
        Me.COL_O.MinimumWidth = 6
        Me.COL_O.Name = "COL_O"
        Me.COL_O.Width = 125
        '
        'COL_P
        '
        Me.COL_P.HeaderText = "COL_P"
        Me.COL_P.MinimumWidth = 6
        Me.COL_P.Name = "COL_P"
        Me.COL_P.Width = 125
        '
        'Col_B
        '
        Me.Col_B.HeaderText = "Col_B"
        Me.Col_B.MinimumWidth = 6
        Me.Col_B.Name = "Col_B"
        Me.Col_B.Width = 125
        '
        'Col_C
        '
        Me.Col_C.HeaderText = "Col_C"
        Me.Col_C.MinimumWidth = 6
        Me.Col_C.Name = "Col_C"
        Me.Col_C.Width = 125
        '
        'Col_D
        '
        Me.Col_D.HeaderText = "Col_D"
        Me.Col_D.MinimumWidth = 6
        Me.Col_D.Name = "Col_D"
        Me.Col_D.Width = 125
        '
        'Col_E
        '
        Me.Col_E.HeaderText = "Col_E"
        Me.Col_E.MinimumWidth = 6
        Me.Col_E.Name = "Col_E"
        Me.Col_E.Width = 125
        '
        'Col_F
        '
        Me.Col_F.HeaderText = "Col_F"
        Me.Col_F.MinimumWidth = 6
        Me.Col_F.Name = "Col_F"
        Me.Col_F.Width = 125
        '
        'Col_G
        '
        Me.Col_G.HeaderText = "Col_G"
        Me.Col_G.MinimumWidth = 6
        Me.Col_G.Name = "Col_G"
        Me.Col_G.Width = 125
        '
        'Col_H
        '
        Me.Col_H.HeaderText = "Col_H"
        Me.Col_H.MinimumWidth = 6
        Me.Col_H.Name = "Col_H"
        Me.Col_H.Width = 125
        '
        'Col_I
        '
        Me.Col_I.HeaderText = "Col_I"
        Me.Col_I.MinimumWidth = 6
        Me.Col_I.Name = "Col_I"
        Me.Col_I.Width = 125
        '
        'Col_J
        '
        Me.Col_J.HeaderText = "Col_J"
        Me.Col_J.MinimumWidth = 6
        Me.Col_J.Name = "Col_J"
        Me.Col_J.Width = 125
        '
        'Cl_J
        '
        Me.Cl_J.HeaderText = "Cl_J"
        Me.Cl_J.MinimumWidth = 6
        Me.Cl_J.Name = "Cl_J"
        Me.Cl_J.Width = 125
        '
        'Col_k
        '
        Me.Col_k.HeaderText = "Col_k"
        Me.Col_k.MinimumWidth = 6
        Me.Col_k.Name = "Col_k"
        Me.Col_k.Width = 125
        '
        'Col_L
        '
        Me.Col_L.HeaderText = "Col_L"
        Me.Col_L.MinimumWidth = 6
        Me.Col_L.Name = "Col_L"
        Me.Col_L.Width = 125
        '
        'Col_M
        '
        Me.Col_M.HeaderText = "Col_M"
        Me.Col_M.MinimumWidth = 6
        Me.Col_M.Name = "Col_M"
        Me.Col_M.Width = 125
        '
        'Col_N
        '
        Me.Col_N.HeaderText = "Col_N"
        Me.Col_N.MinimumWidth = 6
        Me.Col_N.Name = "Col_N"
        Me.Col_N.Width = 125
        '
        'TabPage2
        '
        Me.TabPage2.Controls.Add(Me.btnCEIExport)
        Me.TabPage2.Controls.Add(Me.DGridMInfo)
        Me.TabPage2.Controls.Add(Me.DGridCEI)
        Me.TabPage2.Location = New System.Drawing.Point(4, 28)
        Me.TabPage2.Name = "TabPage2"
        Me.TabPage2.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage2.Size = New System.Drawing.Size(1278, 744)
        Me.TabPage2.TabIndex = 1
        Me.TabPage2.Text = "CEI Details"
        Me.TabPage2.UseVisualStyleBackColor = True
        '
        'btnCEIExport
        '
        Me.btnCEIExport.Font = New System.Drawing.Font("Microsoft Sans Serif", 11.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnCEIExport.Location = New System.Drawing.Point(647, 682)
        Me.btnCEIExport.Name = "btnCEIExport"
        Me.btnCEIExport.Size = New System.Drawing.Size(192, 40)
        Me.btnCEIExport.TabIndex = 17
        Me.btnCEIExport.Text = "Export"
        Me.btnCEIExport.UseVisualStyleBackColor = True
        '
        'DGridMInfo
        '
        Me.DGridMInfo.BackgroundColor = System.Drawing.SystemColors.ActiveCaption
        Me.DGridMInfo.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DGridMInfo.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.MLast_Name, Me.MAddress, Me.PayLoc_Eng})
        Me.DGridMInfo.FilterAndSortEnabled = True
        Me.DGridMInfo.FilterStringChangedInvokeBeforeDatasourceUpdate = True
        Me.DGridMInfo.Location = New System.Drawing.Point(18, 458)
        Me.DGridMInfo.Name = "DGridMInfo"
        Me.DGridMInfo.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.DGridMInfo.RowHeadersWidth = 51
        Me.DGridMInfo.RowTemplate.Height = 24
        Me.DGridMInfo.Size = New System.Drawing.Size(1231, 177)
        Me.DGridMInfo.SortStringChangedInvokeBeforeDatasourceUpdate = True
        Me.DGridMInfo.TabIndex = 1
        '
        'MLast_Name
        '
        Me.MLast_Name.HeaderText = "MLast_Name"
        Me.MLast_Name.MinimumWidth = 22
        Me.MLast_Name.Name = "MLast_Name"
        Me.MLast_Name.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Programmatic
        Me.MLast_Name.Width = 125
        '
        'MAddress
        '
        Me.MAddress.HeaderText = "MAddress"
        Me.MAddress.MinimumWidth = 22
        Me.MAddress.Name = "MAddress"
        Me.MAddress.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Programmatic
        Me.MAddress.Width = 125
        '
        'PayLoc_Eng
        '
        Me.PayLoc_Eng.HeaderText = "PayLoc_Eng"
        Me.PayLoc_Eng.MinimumWidth = 22
        Me.PayLoc_Eng.Name = "PayLoc_Eng"
        Me.PayLoc_Eng.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Programmatic
        Me.PayLoc_Eng.Width = 125
        '
        'DGridCEI
        '
        Me.DGridCEI.BackgroundColor = System.Drawing.SystemColors.ActiveCaption
        Me.DGridCEI.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DGridCEI.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.PFName, Me.PRelation, Me.DOB, Me.End_DT})
        Me.DGridCEI.FilterAndSortEnabled = True
        Me.DGridCEI.FilterStringChangedInvokeBeforeDatasourceUpdate = True
        Me.DGridCEI.Location = New System.Drawing.Point(18, 7)
        Me.DGridCEI.Name = "DGridCEI"
        Me.DGridCEI.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.DGridCEI.RowHeadersWidth = 51
        Me.DGridCEI.RowTemplate.Height = 24
        Me.DGridCEI.Size = New System.Drawing.Size(1231, 432)
        Me.DGridCEI.SortStringChangedInvokeBeforeDatasourceUpdate = True
        Me.DGridCEI.TabIndex = 0
        '
        'PFName
        '
        Me.PFName.HeaderText = "PFName"
        Me.PFName.MinimumWidth = 22
        Me.PFName.Name = "PFName"
        Me.PFName.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Programmatic
        Me.PFName.Width = 125
        '
        'PRelation
        '
        Me.PRelation.HeaderText = "PRelation"
        Me.PRelation.MinimumWidth = 22
        Me.PRelation.Name = "PRelation"
        Me.PRelation.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Programmatic
        Me.PRelation.Width = 125
        '
        'DOB
        '
        Me.DOB.HeaderText = "DOB"
        Me.DOB.MinimumWidth = 22
        Me.DOB.Name = "DOB"
        Me.DOB.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Programmatic
        Me.DOB.Width = 125
        '
        'End_DT
        '
        Me.End_DT.HeaderText = "End_DT"
        Me.End_DT.MinimumWidth = 22
        Me.End_DT.Name = "End_DT"
        Me.End_DT.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Programmatic
        Me.End_DT.Width = 125
        '
        'MHI
        '
        Me.MHI.Controls.Add(Me.Button2)
        Me.MHI.Controls.Add(Me.tblMHI)
        Me.MHI.Location = New System.Drawing.Point(4, 28)
        Me.MHI.Name = "MHI"
        Me.MHI.Padding = New System.Windows.Forms.Padding(3)
        Me.MHI.Size = New System.Drawing.Size(1278, 744)
        Me.MHI.TabIndex = 0
        Me.MHI.Text = "MHI"
        Me.MHI.UseVisualStyleBackColor = True
        '
        'Button2
        '
        Me.Button2.Font = New System.Drawing.Font("Microsoft Sans Serif", 11.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button2.Location = New System.Drawing.Point(366, 698)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(192, 40)
        Me.Button2.TabIndex = 16
        Me.Button2.Text = "Export"
        Me.Button2.UseVisualStyleBackColor = True
        '
        'tblMHI
        '
        Me.tblMHI.BackgroundColor = System.Drawing.SystemColors.ActiveCaption
        Me.tblMHI.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.tblMHI.FilterAndSortEnabled = True
        Me.tblMHI.FilterStringChangedInvokeBeforeDatasourceUpdate = True
        Me.tblMHI.Location = New System.Drawing.Point(0, 6)
        Me.tblMHI.Name = "tblMHI"
        Me.tblMHI.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.tblMHI.RowHeadersWidth = 51
        Me.tblMHI.RowTemplate.Height = 24
        Me.tblMHI.Size = New System.Drawing.Size(1242, 686)
        Me.tblMHI.SortStringChangedInvokeBeforeDatasourceUpdate = True
        Me.tblMHI.TabIndex = 0
        '
        'TabControl1
        '
        Me.TabControl1.Appearance = System.Windows.Forms.TabAppearance.FlatButtons
        Me.TabControl1.Controls.Add(Me.MHI)
        Me.TabControl1.Controls.Add(Me.TabPage2)
        Me.TabControl1.Controls.Add(Me.TabPage1)
        Me.TabControl1.Controls.Add(Me.TabPage3)
        Me.TabControl1.Controls.Add(Me.TabPage4)
        Me.TabControl1.Controls.Add(Me.TabPage5)
        Me.TabControl1.Controls.Add(Me.TabPage8)
        Me.TabControl1.Controls.Add(Me.TabPage10)
        Me.TabControl1.Controls.Add(Me.TabPage12)
        Me.TabControl1.Controls.Add(Me.TabPage11)
        Me.TabControl1.Controls.Add(Me.TabPage13)
        Me.TabControl1.Controls.Add(Me.TabPage6)
        Me.TabControl1.Location = New System.Drawing.Point(376, 64)
        Me.TabControl1.Name = "TabControl1"
        Me.TabControl1.SelectedIndex = 0
        Me.TabControl1.Size = New System.Drawing.Size(1286, 776)
        Me.TabControl1.TabIndex = 15
        '
        'TabPage6
        '
        Me.TabPage6.Controls.Add(Me.DGridMHI_II)
        Me.TabPage6.Location = New System.Drawing.Point(4, 28)
        Me.TabPage6.Name = "TabPage6"
        Me.TabPage6.Size = New System.Drawing.Size(1278, 744)
        Me.TabPage6.TabIndex = 14
        Me.TabPage6.Text = "MHI_Support"
        Me.TabPage6.UseVisualStyleBackColor = True
        '
        'DGridMHI_II
        '
        Me.DGridMHI_II.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DGridMHI_II.Location = New System.Drawing.Point(3, 3)
        Me.DGridMHI_II.Name = "DGridMHI_II"
        Me.DGridMHI_II.RowHeadersWidth = 51
        Me.DGridMHI_II.RowTemplate.Height = 24
        Me.DGridMHI_II.Size = New System.Drawing.Size(1024, 594)
        Me.DGridMHI_II.TabIndex = 0
        '
        'Form1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.WhiteSmoke
        Me.ClientSize = New System.Drawing.Size(1924, 903)
        Me.Controls.Add(Me.TabControl1)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.RichTextBox1)
        Me.Controls.Add(Me.Panel2)
        Me.Controls.Add(Me.Panel1)
        Me.Controls.Add(Me.GroupBox2)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.MenuStrip1)
        Me.MainMenuStrip = Me.MenuStrip1
        Me.Name = "Form1"
        Me.Text = "UNET Out of Pocket Calculator V1.0.6"
        Me.MenuStrip1.ResumeLayout(False)
        Me.MenuStrip1.PerformLayout()
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        Me.TabPage13.ResumeLayout(False)
        CType(Me.DGridMHI, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TabPage11.ResumeLayout(False)
        CType(Me.DGridADJ, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TabPage12.ResumeLayout(False)
        CType(Me.tblOOP, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TabPage10.ResumeLayout(False)
        CType(Me.tblCopay, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TabPage8.ResumeLayout(False)
        CType(Me.DGrid_PG10, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TabPage5.ResumeLayout(False)
        CType(Me.DGrid_PG5, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TabPage4.ResumeLayout(False)
        CType(Me.DGrid_PG4, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TabPage3.ResumeLayout(False)
        CType(Me.DGrid_PG1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TabPage1.ResumeLayout(False)
        CType(Me.DGridOverview, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TabPage2.ResumeLayout(False)
        CType(Me.DGridMInfo, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.DGridCEI, System.ComponentModel.ISupportInitialize).EndInit()
        Me.MHI.ResumeLayout(False)
        CType(Me.tblMHI, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TabControl1.ResumeLayout(False)
        Me.TabPage6.ResumeLayout(False)
        CType(Me.DGridMHI_II, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents BackgroundWorker1 As System.ComponentModel.BackgroundWorker
    Friend WithEvents MenuStrip1 As MenuStrip
    Friend WithEvents MHIHistoryToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents CalculateOOPToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents GatherProvInfoToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ResetViewToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ClearToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents MHIShortOptionToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents PaitentNameAndDOSToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents SortByICNToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ProviderTinAndSuffixToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents DateOfServiceToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ProviderTinToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents DeductibleIndicatorToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents SortByPercentToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents PatientAndProcessedDateToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ProcessedDateAndDraftToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ProcessedDateOnlyToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents FormatMHIToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents CleanClaimToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents FormatMHISheetToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ELGSLetterToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents InstructionsToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents AboutToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents GroupBox1 As GroupBox
    Friend WithEvents btnMain As Button
    Friend WithEvents endSelect As DateTimePicker
    Friend WithEvents startSelect As DateTimePicker
    Friend WithEvents txt_SSN As TextBox
    Friend WithEvents txt_Policy As TextBox
    Friend WithEvents Label5 As Label
    Friend WithEvents Label4 As Label
    Friend WithEvents Label3 As Label
    Friend WithEvents memberList As CheckedListBox
    Friend WithEvents GroupBox2 As GroupBox
    Friend WithEvents txtUserPass As TextBox
    Friend WithEvents txtUserID As TextBox
    Friend WithEvents Label2 As Label
    Friend WithEvents Label1 As Label
    Friend WithEvents Panel1 As Panel
    Friend WithEvents Panel2 As Panel
    Friend WithEvents btnClaimInfo As Button
    Friend WithEvents yearList As ComboBox
    Friend WithEvents RichTextBox1 As RichTextBox
    Friend WithEvents Label6 As Label
    Friend WithEvents TabPage13 As TabPage
    Friend WithEvents DGridMHI As Zuby.ADGV.AdvancedDataGridView
    Friend WithEvents From As DataGridViewTextBoxColumn
    Friend WithEvents Thru As DataGridViewTextBoxColumn
    Friend WithEvents Svc As DataGridViewTextBoxColumn
    Friend WithEvents PS As DataGridViewTextBoxColumn
    Friend WithEvents Nbr As DataGridViewTextBoxColumn
    Friend WithEvents OV As DataGridViewTextBoxColumn
    Friend WithEvents P As DataGridViewTextBoxColumn
    Friend WithEvents N As DataGridViewTextBoxColumn
    Friend WithEvents RC As DataGridViewTextBoxColumn
    Friend WithEvents Charge As DataGridViewTextBoxColumn
    Friend WithEvents NotCov As DataGridViewTextBoxColumn
    Friend WithEvents BM As DataGridViewTextBoxColumn
    Friend WithEvents Covered As DataGridViewTextBoxColumn
    Friend WithEvents Deduct As DataGridViewTextBoxColumn
    Friend WithEvents D As DataGridViewTextBoxColumn
    Friend WithEvents Perc As DataGridViewTextBoxColumn
    Friend WithEvents Paid As DataGridViewTextBoxColumn
    Friend WithEvents S As DataGridViewTextBoxColumn
    Friend WithEvents DC As DataGridViewTextBoxColumn
    Friend WithEvents Sanc As DataGridViewTextBoxColumn
    Friend WithEvents CauseCode As DataGridViewTextBoxColumn
    Friend WithEvents P1 As DataGridViewTextBoxColumn
    Friend WithEvents Tin As DataGridViewTextBoxColumn
    Friend WithEvents Suffix As DataGridViewTextBoxColumn
    Friend WithEvents ClaimNumber As DataGridViewTextBoxColumn
    Friend WithEvents Draft As DataGridViewTextBoxColumn
    Friend WithEvents ProcDate As DataGridViewTextBoxColumn
    Friend WithEvents AdjNo As DataGridViewTextBoxColumn
    Friend WithEvents TotalBilled As DataGridViewTextBoxColumn
    Friend WithEvents TotalPaid As DataGridViewTextBoxColumn
    Friend WithEvents ICN As DataGridViewTextBoxColumn
    Friend WithEvents Suf As DataGridViewTextBoxColumn
    Friend WithEvents FLN As DataGridViewTextBoxColumn
    Friend WithEvents PRS As DataGridViewTextBoxColumn
    Friend WithEvents SI As DataGridViewTextBoxColumn
    Friend WithEvents PT_Name As DataGridViewTextBoxColumn
    Friend WithEvents Blank As DataGridViewTextBoxColumn
    Friend WithEvents PTRel As DataGridViewTextBoxColumn
    Friend WithEvents PTName As DataGridViewTextBoxColumn
    Friend WithEvents INN_Ded As DataGridViewTextBoxColumn
    Friend WithEvents INN_OOP As DataGridViewTextBoxColumn
    Friend WithEvents OON_Ded As DataGridViewTextBoxColumn
    Friend WithEvents OON_OOP As DataGridViewTextBoxColumn
    Friend WithEvents INNDed As DataGridViewTextBoxColumn
    Friend WithEvents INNOOP As DataGridViewTextBoxColumn
    Friend WithEvents OONDed As DataGridViewTextBoxColumn
    Friend WithEvents OONOOP As DataGridViewTextBoxColumn
    Friend WithEvents OI_OIM As DataGridViewTextBoxColumn
    Friend WithEvents OOPCalcRun As DataGridViewTextBoxColumn
    Friend WithEvents ICNandSuffix As DataGridViewTextBoxColumn
    Friend WithEvents Inp_Facility As DataGridViewTextBoxColumn
    Friend WithEvents ProviderName As DataGridViewTextBoxColumn
    Friend WithEvents ProviderType As DataGridViewTextBoxColumn
    Friend WithEvents M1 As DataGridViewTextBoxColumn
    Friend WithEvents M2 As DataGridViewTextBoxColumn
    Friend WithEvents M3 As DataGridViewTextBoxColumn
    Friend WithEvents M4 As DataGridViewTextBoxColumn
    Friend WithEvents TabPage11 As TabPage
    Friend WithEvents DGridADJ As DataGridView
    Friend WithEvents TabPage12 As TabPage
    Friend WithEvents btnOOPExport As Button
    Friend WithEvents tblOOP As DataGridView
    Friend WithEvents TabPage10 As TabPage
    Friend WithEvents tblCopay As DataGridView
    Friend WithEvents Policy As DataGridViewTextBoxColumn
    Friend WithEvents Plan_CodeReportingCode_PlanVar As DataGridViewTextBoxColumn
    Friend WithEvents Year As DataGridViewTextBoxColumn
    Friend WithEvents PatientName As DataGridViewTextBoxColumn
    Friend WithEvents CauseCodes As DataGridViewTextBoxColumn
    Friend WithEvents PlaceofService As DataGridViewTextBoxColumn
    Friend WithEvents MajorMedCalcCode As DataGridViewTextBoxColumn
    Friend WithEvents SpecialProcCode As DataGridViewTextBoxColumn
    Friend WithEvents CopaySet As DataGridViewTextBoxColumn
    Friend WithEvents DataGridViewTextBoxColumn110 As DataGridViewTextBoxColumn
    Friend WithEvents TabPage8 As TabPage
    Friend WithEvents DGrid_PG10 As DataGridView
    Friend WithEvents TabPage5 As TabPage
    Friend WithEvents DGrid_PG5 As DataGridView
    Friend WithEvents TabPage4 As TabPage
    Friend WithEvents DGrid_PG4 As DataGridView
    Friend WithEvents TabPage3 As TabPage
    Friend WithEvents DGrid_PG1 As DataGridView
    Friend WithEvents TabPage1 As TabPage
    Friend WithEvents DGridOverview As DataGridView
    Friend WithEvents TabPage2 As TabPage
    Friend WithEvents btnCEIExport As Button
    Friend WithEvents DGridMInfo As Zuby.ADGV.AdvancedDataGridView
    Friend WithEvents MLast_Name As DataGridViewTextBoxColumn
    Friend WithEvents MAddress As DataGridViewTextBoxColumn
    Friend WithEvents PayLoc_Eng As DataGridViewTextBoxColumn
    Friend WithEvents DGridCEI As Zuby.ADGV.AdvancedDataGridView
    Friend WithEvents PFName As DataGridViewTextBoxColumn
    Friend WithEvents PRelation As DataGridViewTextBoxColumn
    Friend WithEvents DOB As DataGridViewTextBoxColumn
    Friend WithEvents MHI As TabPage
    Friend WithEvents Button2 As Button
    Friend WithEvents tblMHI As Zuby.ADGV.AdvancedDataGridView
    Friend WithEvents TabControl1 As TabControl
    Friend WithEvents GetMMIToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents TabPage6 As TabPage
    Friend WithEvents DGridMHI_II As DataGridView
    Friend WithEvents End_DT As DataGridViewTextBoxColumn
    Friend WithEvents Chk_select_memlist As CheckBox
    Friend WithEvents Col_A As DataGridViewTextBoxColumn
    Friend WithEvents COL_O As DataGridViewTextBoxColumn
    Friend WithEvents COL_P As DataGridViewTextBoxColumn
    Friend WithEvents Col_B As DataGridViewTextBoxColumn
    Friend WithEvents Col_C As DataGridViewTextBoxColumn
    Friend WithEvents Col_D As DataGridViewTextBoxColumn
    Friend WithEvents Col_E As DataGridViewTextBoxColumn
    Friend WithEvents Col_F As DataGridViewTextBoxColumn
    Friend WithEvents Col_G As DataGridViewTextBoxColumn
    Friend WithEvents Col_H As DataGridViewTextBoxColumn
    Friend WithEvents Col_I As DataGridViewTextBoxColumn
    Friend WithEvents Col_J As DataGridViewTextBoxColumn
    Friend WithEvents Cl_J As DataGridViewTextBoxColumn
    Friend WithEvents Col_k As DataGridViewTextBoxColumn
    Friend WithEvents Col_L As DataGridViewTextBoxColumn
    Friend WithEvents Col_M As DataGridViewTextBoxColumn
    Friend WithEvents Col_N As DataGridViewTextBoxColumn
    Friend WithEvents DataGridViewTextBoxColumn8 As DataGridViewTextBoxColumn
    Friend WithEvents DataGridViewTextBoxColumn9 As DataGridViewTextBoxColumn
    Friend WithEvents DataGridViewTextBoxColumn10 As DataGridViewTextBoxColumn
    Friend WithEvents DataGridViewTextBoxColumn11 As DataGridViewTextBoxColumn
    Friend WithEvents DataGridViewTextBoxColumn12 As DataGridViewTextBoxColumn
    Friend WithEvents DataGridViewTextBoxColumn13 As DataGridViewTextBoxColumn
    Friend WithEvents DataGridViewTextBoxColumn14 As DataGridViewTextBoxColumn
    Friend WithEvents C_H As DataGridViewTextBoxColumn
    Friend WithEvents C_I As DataGridViewTextBoxColumn
    Friend WithEvents C_J As DataGridViewTextBoxColumn
    Friend WithEvents C_K As DataGridViewTextBoxColumn
    Friend WithEvents C_L As DataGridViewTextBoxColumn
    Friend WithEvents C_M As DataGridViewTextBoxColumn
    Friend WithEvents C_N As DataGridViewTextBoxColumn
    Friend WithEvents DataGridViewTextBoxColumn1 As DataGridViewTextBoxColumn
    Friend WithEvents DataGridViewTextBoxColumn2 As DataGridViewTextBoxColumn
    Friend WithEvents DataGridViewTextBoxColumn3 As DataGridViewTextBoxColumn
    Friend WithEvents DataGridViewTextBoxColumn4 As DataGridViewTextBoxColumn
    Friend WithEvents DataGridViewTextBoxColumn5 As DataGridViewTextBoxColumn
    Friend WithEvents DataGridViewTextBoxColumn6 As DataGridViewTextBoxColumn
    Friend WithEvents DataGridViewTextBoxColumn7 As DataGridViewTextBoxColumn
    Friend WithEvents CL_H As DataGridViewTextBoxColumn
    Friend WithEvents CL_I As DataGridViewTextBoxColumn
    Friend WithEvents CLL_J As DataGridViewTextBoxColumn
    Friend WithEvents CLL_K As DataGridViewTextBoxColumn
    Friend WithEvents CLL_L As DataGridViewTextBoxColumn
    Friend WithEvents CLL_M As DataGridViewTextBoxColumn
    Friend WithEvents CLL_N As DataGridViewTextBoxColumn
    Friend WithEvents DataGridViewTextBoxColumn15 As DataGridViewTextBoxColumn
    Friend WithEvents DataGridViewTextBoxColumn16 As DataGridViewTextBoxColumn
    Friend WithEvents DataGridViewTextBoxColumn17 As DataGridViewTextBoxColumn
    Friend WithEvents DataGridViewTextBoxColumn18 As DataGridViewTextBoxColumn
    Friend WithEvents DataGridViewTextBoxColumn19 As DataGridViewTextBoxColumn
    Friend WithEvents DataGridViewTextBoxColumn20 As DataGridViewTextBoxColumn
    Friend WithEvents DataGridViewTextBoxColumn21 As DataGridViewTextBoxColumn
    Friend WithEvents Colu_H As DataGridViewTextBoxColumn
    Friend WithEvents Colu_I As DataGridViewTextBoxColumn
    Friend WithEvents Colu_J As DataGridViewTextBoxColumn
    Friend WithEvents Colu_K As DataGridViewTextBoxColumn
    Friend WithEvents Colu_L As DataGridViewTextBoxColumn
    Friend WithEvents Colu_M As DataGridViewTextBoxColumn
    Friend WithEvents Colu_N As DataGridViewTextBoxColumn
    Friend WithEvents Colu_O As DataGridViewTextBoxColumn
    Friend WithEvents Colu_P As DataGridViewTextBoxColumn
    Friend WithEvents DataGridViewTextBoxColumn29 As DataGridViewTextBoxColumn
    Friend WithEvents DataGridViewTextBoxColumn30 As DataGridViewTextBoxColumn
    Friend WithEvents DataGridViewTextBoxColumn31 As DataGridViewTextBoxColumn
    Friend WithEvents DataGridViewTextBoxColumn32 As DataGridViewTextBoxColumn
    Friend WithEvents DataGridViewTextBoxColumn33 As DataGridViewTextBoxColumn
    Friend WithEvents DataGridViewTextBoxColumn34 As DataGridViewTextBoxColumn
    Friend WithEvents DataGridViewTextBoxColumn35 As DataGridViewTextBoxColumn
    Friend WithEvents Column_H As DataGridViewTextBoxColumn
    Friend WithEvents Column_I As DataGridViewTextBoxColumn
    Friend WithEvents Column_J As DataGridViewTextBoxColumn
    Friend WithEvents Column_K As DataGridViewTextBoxColumn
    Friend WithEvents Column_L As DataGridViewTextBoxColumn
    Friend WithEvents Column_M As DataGridViewTextBoxColumn
    Friend WithEvents Column_N As DataGridViewTextBoxColumn
    Friend WithEvents Column_O As DataGridViewTextBoxColumn
    Friend WithEvents Column_P As DataGridViewTextBoxColumn
    Friend WithEvents MMI_Message As RichTextBox
    Friend WithEvents Button1 As Button
    Friend WithEvents ClearAllFilterToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ApplyColorToolStripMenuItem As ToolStripMenuItem
End Class
