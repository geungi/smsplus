<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmPartRtn
    Inherits System.Windows.Forms.Form

    'Form은 Dispose를 재정의하여 구성 요소 목록을 정리합니다.
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

    'Windows Form 디자이너에 필요합니다.
    Private components As System.ComponentModel.IContainer

    '참고: 다음 프로시저는 Windows Form 디자이너에 필요합니다.
    '수정하려면 Windows Form 디자이너를 사용하십시오.
    '코드 편집기를 사용하여 수정하지 마십시오.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(FrmPartRtn))
        Me.FrmGuiMgr = New DevComponents.DotNetBar.DotNetBarManager(Me.components)
        Me.DockSite4 = New DevComponents.DotNetBar.DockSite
        Me.Bar3 = New DevComponents.DotNetBar.Bar
        Me.PanelDockContainer3 = New DevComponents.DotNetBar.PanelDockContainer
        Me.FpSpread1 = New FarPoint.Win.Spread.FpSpread
        Me.FpSpread1_Sheet1 = New FarPoint.Win.Spread.SheetView
        Me.DockContainerItem3 = New DevComponents.DotNetBar.DockContainerItem
        Me.DockSite9 = New DevComponents.DotNetBar.DockSite
        Me.Bar5 = New DevComponents.DotNetBar.Bar
        Me.FpSpread3 = New FarPoint.Win.Spread.FpSpread
        Me.FpSpread3_Sheet1 = New FarPoint.Win.Spread.SheetView
        Me.Bar8 = New DevComponents.DotNetBar.Bar
        Me.LabelItem4 = New DevComponents.DotNetBar.LabelItem
        Me.Bar4 = New DevComponents.DotNetBar.Bar
        Me.Splitter1 = New System.Windows.Forms.Splitter
        Me.FpSpread2 = New FarPoint.Win.Spread.FpSpread
        Me.FpSpread2_Sheet1 = New FarPoint.Win.Spread.SheetView
        Me.Bar7 = New DevComponents.DotNetBar.Bar
        Me.LabelX4 = New DevComponents.DotNetBar.LabelX
        Me.PtWhCb = New DevComponents.DotNetBar.Controls.ComboBoxEx
        Me.PtMdCb = New DevComponents.DotNetBar.Controls.ComboBoxEx
        Me.LabelX1 = New DevComponents.DotNetBar.LabelX
        Me.LabelItem1 = New DevComponents.DotNetBar.LabelItem
        Me.ControlContainerItem3 = New DevComponents.DotNetBar.ControlContainerItem
        Me.ControlContainerItem4 = New DevComponents.DotNetBar.ControlContainerItem
        Me.DockSite1 = New DevComponents.DotNetBar.DockSite
        Me.DockSite2 = New DevComponents.DotNetBar.DockSite
        Me.DockSite8 = New DevComponents.DotNetBar.DockSite
        Me.DockSite5 = New DevComponents.DotNetBar.DockSite
        Me.DockSite6 = New DevComponents.DotNetBar.DockSite
        Me.DockSite3 = New DevComponents.DotNetBar.DockSite
        Me.Bar1 = New DevComponents.DotNetBar.Bar
        Me.PanelDockContainer1 = New DevComponents.DotNetBar.PanelDockContainer
        Me.CtlBar = New DevComponents.DotNetBar.Bar
        Me.FindBtn = New DevComponents.DotNetBar.ButtonItem
        Me.SaveBtn = New DevComponents.DotNetBar.ButtonItem
        Me.DelBtn = New DevComponents.DotNetBar.ButtonItem
        Me.PrtBtn = New DevComponents.DotNetBar.ButtonItem
        Me.Excel = New DevComponents.DotNetBar.ButtonItem
        Me.DockContainerItem1 = New DevComponents.DotNetBar.DockContainerItem
        Me.Bar2 = New DevComponents.DotNetBar.Bar
        Me.PanelDockContainer2 = New DevComponents.DotNetBar.PanelDockContainer
        Me.LabelX6 = New DevComponents.DotNetBar.LabelX
        Me.PartNoTxt = New DevComponents.DotNetBar.Controls.TextBoxX
        Me.ToCb = New DevComponents.DotNetBar.Controls.ComboBoxEx
        Me.LabelX7 = New DevComponents.DotNetBar.LabelX
        Me.FromCb = New DevComponents.DotNetBar.Controls.ComboBoxEx
        Me.POStDate = New DevComponents.Editors.DateTimeAdv.DateTimeInput
        Me.POEdDate = New DevComponents.Editors.DateTimeAdv.DateTimeInput
        Me.LabelX2 = New DevComponents.DotNetBar.LabelX
        Me.DockContainerItem2 = New DevComponents.DotNetBar.DockContainerItem
        Me.DockContainerItem5 = New DevComponents.DotNetBar.DockContainerItem
        Me.DockContainerItem6 = New DevComponents.DotNetBar.DockContainerItem
        Me.DockContainerItem4 = New DevComponents.DotNetBar.DockContainerItem
        Me.ControlContainerItem1 = New DevComponents.DotNetBar.ControlContainerItem
        Me.ControlContainerItem2 = New DevComponents.DotNetBar.ControlContainerItem
        Me.ContextMenuBar1 = New DevComponents.DotNetBar.ContextMenuBar
        Me.CtxSp = New DevComponents.DotNetBar.ButtonItem
        Me.bDel = New DevComponents.DotNetBar.ButtonItem
        Me.ButtonItem1 = New DevComponents.DotNetBar.ButtonItem
        Me.CtxSp2 = New DevComponents.DotNetBar.ButtonItem
        Me.ButtonItem2 = New DevComponents.DotNetBar.ButtonItem
        Me.cPrint = New DevComponents.DotNetBar.ButtonItem
        Me.cExcel = New DevComponents.DotNetBar.ButtonItem
        Me.bSave = New DevComponents.DotNetBar.ButtonItem
        Me.bPrint = New DevComponents.DotNetBar.ButtonItem
        Me.bExcel = New DevComponents.DotNetBar.ButtonItem
        Me.SaveFileDialog1 = New System.Windows.Forms.SaveFileDialog
        Me.LabelItem2 = New DevComponents.DotNetBar.LabelItem
        Me.XlsBtn1 = New DevComponents.DotNetBar.ButtonX
        Me.PrtBtn1 = New DevComponents.DotNetBar.ButtonX
        Me.DelBtn1 = New DevComponents.DotNetBar.ButtonX
        Me.SaveBtn1 = New DevComponents.DotNetBar.ButtonX
        Me.NewBtn1 = New DevComponents.DotNetBar.ButtonX
        Me.DockSite4.SuspendLayout()
        CType(Me.Bar3, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Bar3.SuspendLayout()
        Me.PanelDockContainer3.SuspendLayout()
        CType(Me.FpSpread1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.FpSpread1_Sheet1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.DockSite9.SuspendLayout()
        CType(Me.Bar5, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Bar5.SuspendLayout()
        CType(Me.FpSpread3, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.FpSpread3_Sheet1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.Bar8, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.Bar4, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Bar4.SuspendLayout()
        CType(Me.FpSpread2, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.FpSpread2_Sheet1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.Bar7, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Bar7.SuspendLayout()
        Me.DockSite3.SuspendLayout()
        CType(Me.Bar1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Bar1.SuspendLayout()
        Me.PanelDockContainer1.SuspendLayout()
        CType(Me.CtlBar, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.Bar2, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Bar2.SuspendLayout()
        Me.PanelDockContainer2.SuspendLayout()
        CType(Me.POStDate, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.POEdDate, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.ContextMenuBar1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'FrmGuiMgr
        '
        Me.FrmGuiMgr.AutoDispatchShortcuts.Add(DevComponents.DotNetBar.eShortcut.F1)
        Me.FrmGuiMgr.AutoDispatchShortcuts.Add(DevComponents.DotNetBar.eShortcut.CtrlC)
        Me.FrmGuiMgr.AutoDispatchShortcuts.Add(DevComponents.DotNetBar.eShortcut.CtrlA)
        Me.FrmGuiMgr.AutoDispatchShortcuts.Add(DevComponents.DotNetBar.eShortcut.CtrlV)
        Me.FrmGuiMgr.AutoDispatchShortcuts.Add(DevComponents.DotNetBar.eShortcut.CtrlX)
        Me.FrmGuiMgr.AutoDispatchShortcuts.Add(DevComponents.DotNetBar.eShortcut.CtrlZ)
        Me.FrmGuiMgr.AutoDispatchShortcuts.Add(DevComponents.DotNetBar.eShortcut.CtrlY)
        Me.FrmGuiMgr.AutoDispatchShortcuts.Add(DevComponents.DotNetBar.eShortcut.Del)
        Me.FrmGuiMgr.AutoDispatchShortcuts.Add(DevComponents.DotNetBar.eShortcut.Ins)
        Me.FrmGuiMgr.BottomDockSite = Me.DockSite4
        Me.FrmGuiMgr.FillDockSite = Me.DockSite9
        Me.FrmGuiMgr.LeftDockSite = Me.DockSite1
        Me.FrmGuiMgr.LicenseKey = "F962CEC7-CD8F-4911-A9E9-CAB39962FC1F"
        Me.FrmGuiMgr.ParentForm = Me
        Me.FrmGuiMgr.RightDockSite = Me.DockSite2
        Me.FrmGuiMgr.Style = DevComponents.DotNetBar.eDotNetBarStyle.Office2007
        Me.FrmGuiMgr.ToolbarBottomDockSite = Me.DockSite8
        Me.FrmGuiMgr.ToolbarLeftDockSite = Me.DockSite5
        Me.FrmGuiMgr.ToolbarRightDockSite = Me.DockSite6
        Me.FrmGuiMgr.TopDockSite = Me.DockSite3
        Me.FrmGuiMgr.UseHook = True
        '
        'DockSite4
        '
        Me.DockSite4.AccessibleRole = System.Windows.Forms.AccessibleRole.Window
        Me.DockSite4.Controls.Add(Me.Bar3)
        Me.DockSite4.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.DockSite4.DocumentDockContainer = New DevComponents.DotNetBar.DocumentDockContainer(New DevComponents.DotNetBar.DocumentBaseContainer() {CType(New DevComponents.DotNetBar.DocumentBarContainer(Me.Bar3, 1016, 151), DevComponents.DotNetBar.DocumentBaseContainer)}, DevComponents.DotNetBar.eOrientation.Vertical)
        Me.DockSite4.Location = New System.Drawing.Point(0, 578)
        Me.DockSite4.Name = "DockSite4"
        Me.DockSite4.Size = New System.Drawing.Size(1016, 154)
        Me.DockSite4.TabIndex = 3
        Me.DockSite4.TabStop = False
        '
        'Bar3
        '
        Me.Bar3.AccessibleDescription = "DotNetBar Bar (Bar3)"
        Me.Bar3.AccessibleName = "DotNetBar Bar"
        Me.Bar3.AccessibleRole = System.Windows.Forms.AccessibleRole.ToolBar
        Me.Bar3.AutoSyncBarCaption = True
        Me.Bar3.CloseSingleTab = True
        Me.Bar3.Controls.Add(Me.PanelDockContainer3)
        Me.Bar3.Font = New System.Drawing.Font("Verdana", 8.0!)
        Me.Bar3.GrabHandleStyle = DevComponents.DotNetBar.eGrabHandleStyle.Caption
        Me.Bar3.Items.AddRange(New DevComponents.DotNetBar.BaseItem() {Me.DockContainerItem3})
        Me.Bar3.LayoutType = DevComponents.DotNetBar.eLayoutType.DockContainer
        Me.Bar3.Location = New System.Drawing.Point(0, 3)
        Me.Bar3.Name = "Bar3"
        Me.Bar3.Size = New System.Drawing.Size(1016, 151)
        Me.Bar3.Stretch = True
        Me.Bar3.Style = DevComponents.DotNetBar.eDotNetBarStyle.Office2007
        Me.Bar3.TabIndex = 0
        Me.Bar3.TabStop = False
        Me.Bar3.Text = "PART I/O HISTORY"
        '
        'PanelDockContainer3
        '
        Me.PanelDockContainer3.ColorSchemeStyle = DevComponents.DotNetBar.eDotNetBarStyle.Office2007
        Me.PanelDockContainer3.Controls.Add(Me.FpSpread1)
        Me.PanelDockContainer3.Location = New System.Drawing.Point(3, 23)
        Me.PanelDockContainer3.Name = "PanelDockContainer3"
        Me.PanelDockContainer3.Size = New System.Drawing.Size(1010, 125)
        Me.PanelDockContainer3.Style.Alignment = System.Drawing.StringAlignment.Center
        Me.PanelDockContainer3.Style.BackColor1.ColorSchemePart = DevComponents.DotNetBar.eColorSchemePart.BarBackground
        Me.PanelDockContainer3.Style.BackColor2.ColorSchemePart = DevComponents.DotNetBar.eColorSchemePart.BarBackground2
        Me.PanelDockContainer3.Style.BorderColor.ColorSchemePart = DevComponents.DotNetBar.eColorSchemePart.BarDockedBorder
        Me.PanelDockContainer3.Style.ForeColor.ColorSchemePart = DevComponents.DotNetBar.eColorSchemePart.ItemText
        Me.PanelDockContainer3.Style.GradientAngle = 90
        Me.PanelDockContainer3.TabIndex = 0
        '
        'FpSpread1
        '
        Me.FpSpread1.AccessibleDescription = ""
        Me.FpSpread1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.FpSpread1.Location = New System.Drawing.Point(0, 0)
        Me.FpSpread1.Name = "FpSpread1"
        Me.FpSpread1.Sheets.AddRange(New FarPoint.Win.Spread.SheetView() {Me.FpSpread1_Sheet1})
        Me.FpSpread1.Size = New System.Drawing.Size(1010, 125)
        Me.FpSpread1.TabIndex = 0
        Me.FpSpread1.SetActiveViewport(0, -1, -1)
        '
        'FpSpread1_Sheet1
        '
        Me.FpSpread1_Sheet1.Reset()
        Me.FpSpread1_Sheet1.SheetName = "Sheet1"
        'Formulas and custom names must be loaded with R1C1 reference style
        Me.FpSpread1_Sheet1.ReferenceStyle = FarPoint.Win.Spread.Model.ReferenceStyle.R1C1
        Me.FpSpread1_Sheet1.ColumnCount = 0
        Me.FpSpread1_Sheet1.RowCount = 0
        Me.FpSpread1_Sheet1.ActiveColumnIndex = -1
        Me.FpSpread1_Sheet1.ActiveRowIndex = -1
        Me.FpSpread1_Sheet1.ReferenceStyle = FarPoint.Win.Spread.Model.ReferenceStyle.A1
        '
        'DockContainerItem3
        '
        Me.DockContainerItem3.Control = Me.PanelDockContainer3
        Me.DockContainerItem3.Name = "DockContainerItem3"
        Me.DockContainerItem3.Text = "PART I/O HISTORY"
        '
        'DockSite9
        '
        Me.DockSite9.AccessibleRole = System.Windows.Forms.AccessibleRole.Window
        Me.DockSite9.Controls.Add(Me.Bar5)
        Me.DockSite9.Controls.Add(Me.Bar4)
        Me.DockSite9.Dock = System.Windows.Forms.DockStyle.Fill
        Me.DockSite9.DocumentDockContainer = New DevComponents.DotNetBar.DocumentDockContainer(New DevComponents.DotNetBar.DocumentBaseContainer() {CType(New DevComponents.DotNetBar.DocumentBarContainer(Me.Bar4, 575, 510), DevComponents.DotNetBar.DocumentBaseContainer), CType(New DevComponents.DotNetBar.DocumentBarContainer(Me.Bar5, 438, 510), DevComponents.DotNetBar.DocumentBaseContainer)}, DevComponents.DotNetBar.eOrientation.Horizontal)
        Me.DockSite9.Location = New System.Drawing.Point(0, 68)
        Me.DockSite9.Name = "DockSite9"
        Me.DockSite9.Size = New System.Drawing.Size(1016, 510)
        Me.DockSite9.TabIndex = 8
        Me.DockSite9.TabStop = False
        '
        'Bar5
        '
        Me.Bar5.AccessibleDescription = "DotNetBar Bar (Bar5)"
        Me.Bar5.AccessibleName = "DotNetBar Bar"
        Me.Bar5.AccessibleRole = System.Windows.Forms.AccessibleRole.ToolBar
        Me.Bar5.CanDockBottom = False
        Me.Bar5.CanDockDocument = True
        Me.Bar5.CanDockLeft = False
        Me.Bar5.CanDockRight = False
        Me.Bar5.CanDockTab = False
        Me.Bar5.CanDockTop = False
        Me.Bar5.CanHide = True
        Me.Bar5.CanReorderTabs = False
        Me.Bar5.CanUndock = False
        Me.Bar5.Controls.Add(Me.FpSpread3)
        Me.Bar5.Controls.Add(Me.Bar8)
        Me.Bar5.DockTabAlignment = DevComponents.DotNetBar.eTabStripAlignment.Top
        Me.Bar5.Font = New System.Drawing.Font("Verdana", 8.0!)
        Me.Bar5.LayoutType = DevComponents.DotNetBar.eLayoutType.DockContainer
        Me.Bar5.Location = New System.Drawing.Point(578, 0)
        Me.Bar5.Name = "Bar5"
        Me.Bar5.Size = New System.Drawing.Size(438, 510)
        Me.Bar5.Stretch = True
        Me.Bar5.Style = DevComponents.DotNetBar.eDotNetBarStyle.Office2007
        Me.Bar5.TabIndex = 1
        Me.Bar5.TabStop = False
        '
        'FpSpread3
        '
        Me.FpSpread3.AccessibleDescription = ""
        Me.FpSpread3.AllowDragDrop = True
        Me.FpSpread3.AllowDragFill = True
        Me.FpSpread3.AllowDrop = True
        Me.FpSpread3.Dock = System.Windows.Forms.DockStyle.Fill
        Me.FpSpread3.EditModeReplace = True
        Me.FpSpread3.Location = New System.Drawing.Point(0, 24)
        Me.FpSpread3.Name = "FpSpread3"
        Me.FpSpread3.Sheets.AddRange(New FarPoint.Win.Spread.SheetView() {Me.FpSpread3_Sheet1})
        Me.FpSpread3.Size = New System.Drawing.Size(438, 486)
        Me.FpSpread3.TabIndex = 5
        Me.FpSpread3.SetActiveViewport(0, -1, -1)
        '
        'FpSpread3_Sheet1
        '
        Me.FpSpread3_Sheet1.Reset()
        Me.FpSpread3_Sheet1.SheetName = "Sheet1"
        'Formulas and custom names must be loaded with R1C1 reference style
        Me.FpSpread3_Sheet1.ReferenceStyle = FarPoint.Win.Spread.Model.ReferenceStyle.R1C1
        Me.FpSpread3_Sheet1.ColumnCount = 0
        Me.FpSpread3_Sheet1.RowCount = 0
        Me.FpSpread3_Sheet1.ActiveColumnIndex = -1
        Me.FpSpread3_Sheet1.ActiveRowIndex = -1
        Me.FpSpread3_Sheet1.ReferenceStyle = FarPoint.Win.Spread.Model.ReferenceStyle.A1
        '
        'Bar8
        '
        Me.Bar8.Dock = System.Windows.Forms.DockStyle.Top
        Me.Bar8.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Bar8.Items.AddRange(New DevComponents.DotNetBar.BaseItem() {Me.LabelItem4})
        Me.Bar8.Location = New System.Drawing.Point(0, 0)
        Me.Bar8.Name = "Bar8"
        Me.Bar8.PaddingBottom = 4
        Me.Bar8.PaddingLeft = 2
        Me.Bar8.PaddingRight = 2
        Me.Bar8.PaddingTop = 4
        Me.Bar8.Size = New System.Drawing.Size(438, 24)
        Me.Bar8.Stretch = True
        Me.Bar8.Style = DevComponents.DotNetBar.eDotNetBarStyle.Office2007
        Me.Bar8.TabIndex = 4
        Me.Bar8.TabStop = False
        Me.Bar8.Text = "TRANSFERING PART"
        '
        'LabelItem4
        '
        Me.LabelItem4.Name = "LabelItem4"
        Me.LabelItem4.Text = "이동 품목 리스트"
        '
        'Bar4
        '
        Me.Bar4.AccessibleDescription = "DotNetBar Bar (Bar4)"
        Me.Bar4.AccessibleName = "DotNetBar Bar"
        Me.Bar4.AccessibleRole = System.Windows.Forms.AccessibleRole.MenuBar
        Me.Bar4.AutoCreateCaptionMenu = False
        Me.Bar4.AutoSyncBarCaption = True
        Me.Bar4.CanAutoHide = False
        Me.Bar4.CanCustomize = False
        Me.Bar4.CanDockBottom = False
        Me.Bar4.CanDockDocument = True
        Me.Bar4.CanDockLeft = False
        Me.Bar4.CanDockRight = False
        Me.Bar4.CanDockTab = False
        Me.Bar4.CanDockTop = False
        Me.Bar4.CanReorderTabs = False
        Me.Bar4.CanUndock = False
        Me.Bar4.Controls.Add(Me.Splitter1)
        Me.Bar4.Controls.Add(Me.FpSpread2)
        Me.Bar4.Controls.Add(Me.Bar7)
        Me.Bar4.DisplayMoreItemsOnMenu = True
        Me.Bar4.DockTabAlignment = DevComponents.DotNetBar.eTabStripAlignment.Top
        Me.Bar4.EqualButtonSize = True
        Me.Bar4.Font = New System.Drawing.Font("Verdana", 8.0!)
        Me.Bar4.LayoutType = DevComponents.DotNetBar.eLayoutType.DockContainer
        Me.Bar4.Location = New System.Drawing.Point(0, 0)
        Me.Bar4.MenuBar = True
        Me.Bar4.Name = "Bar4"
        Me.Bar4.Size = New System.Drawing.Size(575, 510)
        Me.Bar4.Stretch = True
        Me.Bar4.Style = DevComponents.DotNetBar.eDotNetBarStyle.Office2007
        Me.Bar4.TabIndex = 0
        Me.Bar4.TabStop = False
        Me.Bar4.Text = "PART ROOM"
        '
        'Splitter1
        '
        Me.Splitter1.Location = New System.Drawing.Point(0, 26)
        Me.Splitter1.Name = "Splitter1"
        Me.Splitter1.Size = New System.Drawing.Size(3, 484)
        Me.Splitter1.TabIndex = 5
        Me.Splitter1.TabStop = False
        '
        'FpSpread2
        '
        Me.FpSpread2.AccessibleDescription = ""
        Me.FpSpread2.AllowDragDrop = True
        Me.FpSpread2.AllowDragFill = True
        Me.FpSpread2.AllowDrop = True
        Me.FpSpread2.Dock = System.Windows.Forms.DockStyle.Fill
        Me.FpSpread2.Location = New System.Drawing.Point(0, 26)
        Me.FpSpread2.Name = "FpSpread2"
        Me.FpSpread2.Sheets.AddRange(New FarPoint.Win.Spread.SheetView() {Me.FpSpread2_Sheet1})
        Me.FpSpread2.Size = New System.Drawing.Size(575, 484)
        Me.FpSpread2.TabIndex = 4
        Me.FpSpread2.SetActiveViewport(0, -1, -1)
        '
        'FpSpread2_Sheet1
        '
        Me.FpSpread2_Sheet1.Reset()
        Me.FpSpread2_Sheet1.SheetName = "Sheet1"
        'Formulas and custom names must be loaded with R1C1 reference style
        Me.FpSpread2_Sheet1.ReferenceStyle = FarPoint.Win.Spread.Model.ReferenceStyle.R1C1
        Me.FpSpread2_Sheet1.ColumnCount = 0
        Me.FpSpread2_Sheet1.RowCount = 0
        Me.FpSpread2_Sheet1.ActiveColumnIndex = -1
        Me.FpSpread2_Sheet1.ActiveRowIndex = -1
        Me.FpSpread2_Sheet1.SelectionPolicy = FarPoint.Win.Spread.Model.SelectionPolicy.MultiRange
        Me.FpSpread2_Sheet1.ReferenceStyle = FarPoint.Win.Spread.Model.ReferenceStyle.A1
        '
        'Bar7
        '
        Me.Bar7.Controls.Add(Me.LabelX4)
        Me.Bar7.Controls.Add(Me.PtWhCb)
        Me.Bar7.Controls.Add(Me.PtMdCb)
        Me.Bar7.Controls.Add(Me.LabelX1)
        Me.Bar7.Dock = System.Windows.Forms.DockStyle.Top
        Me.Bar7.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Bar7.Items.AddRange(New DevComponents.DotNetBar.BaseItem() {Me.LabelItem1, Me.ControlContainerItem3, Me.ControlContainerItem4})
        Me.Bar7.Location = New System.Drawing.Point(0, 0)
        Me.Bar7.Name = "Bar7"
        Me.Bar7.Size = New System.Drawing.Size(575, 26)
        Me.Bar7.Stretch = True
        Me.Bar7.Style = DevComponents.DotNetBar.eDotNetBarStyle.Office2007
        Me.Bar7.TabIndex = 2
        Me.Bar7.TabStop = False
        Me.Bar7.Text = "Bar7"
        '
        'LabelX4
        '
        Me.LabelX4.BackColor = System.Drawing.Color.Transparent
        '
        '
        '
        Me.LabelX4.BackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square
        Me.LabelX4.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LabelX4.Location = New System.Drawing.Point(223, 7)
        Me.LabelX4.Name = "LabelX4"
        Me.LabelX4.Size = New System.Drawing.Size(63, 12)
        Me.LabelX4.TabIndex = 10001
        Me.LabelX4.Text = "출고창고"
        Me.LabelX4.TextAlignment = System.Drawing.StringAlignment.Center
        '
        'PtWhCb
        '
        Me.PtWhCb.DisplayMember = "Text"
        Me.PtWhCb.DrawMode = System.Windows.Forms.DrawMode.OwnerDrawFixed
        Me.PtWhCb.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.PtWhCb.FormattingEnabled = True
        Me.PtWhCb.ItemHeight = 15
        Me.PtWhCb.Location = New System.Drawing.Point(287, 2)
        Me.PtWhCb.Name = "PtWhCb"
        Me.PtWhCb.Size = New System.Drawing.Size(171, 21)
        Me.PtWhCb.TabIndex = 10000
        '
        'PtMdCb
        '
        Me.PtMdCb.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.PtMdCb.DisplayMember = "Text"
        Me.PtMdCb.DrawMode = System.Windows.Forms.DrawMode.OwnerDrawFixed
        Me.PtMdCb.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.PtMdCb.FormattingEnabled = True
        Me.PtMdCb.ItemHeight = 15
        Me.PtMdCb.Location = New System.Drawing.Point(685, 2)
        Me.PtMdCb.Name = "PtMdCb"
        Me.PtMdCb.Size = New System.Drawing.Size(107, 21)
        Me.PtMdCb.TabIndex = 5
        '
        'LabelX1
        '
        Me.LabelX1.AutoSize = True
        Me.LabelX1.BackColor = System.Drawing.Color.Transparent
        '
        '
        '
        Me.LabelX1.BackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square
        Me.LabelX1.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LabelX1.Location = New System.Drawing.Point(75, 4)
        Me.LabelX1.Name = "LabelX1"
        Me.LabelX1.Size = New System.Drawing.Size(31, 17)
        Me.LabelX1.TabIndex = 7
        Me.LabelX1.Text = "모델"
        Me.LabelX1.TextAlignment = System.Drawing.StringAlignment.Center
        '
        'LabelItem1
        '
        Me.LabelItem1.Name = "LabelItem1"
        Me.LabelItem1.Text = "품목 리스트"
        '
        'ControlContainerItem3
        '
        Me.ControlContainerItem3.AllowItemResize = False
        Me.ControlContainerItem3.Control = Me.LabelX1
        Me.ControlContainerItem3.MenuVisibility = DevComponents.DotNetBar.eMenuVisibility.VisibleAlways
        Me.ControlContainerItem3.Name = "ControlContainerItem3"
        '
        'ControlContainerItem4
        '
        Me.ControlContainerItem4.AllowItemResize = False
        Me.ControlContainerItem4.Control = Me.PtMdCb
        Me.ControlContainerItem4.MenuVisibility = DevComponents.DotNetBar.eMenuVisibility.VisibleAlways
        Me.ControlContainerItem4.Name = "ControlContainerItem4"
        '
        'DockSite1
        '
        Me.DockSite1.AccessibleRole = System.Windows.Forms.AccessibleRole.Window
        Me.DockSite1.Dock = System.Windows.Forms.DockStyle.Left
        Me.DockSite1.DocumentDockContainer = New DevComponents.DotNetBar.DocumentDockContainer
        Me.DockSite1.Location = New System.Drawing.Point(0, 68)
        Me.DockSite1.Name = "DockSite1"
        Me.DockSite1.Size = New System.Drawing.Size(0, 510)
        Me.DockSite1.TabIndex = 0
        Me.DockSite1.TabStop = False
        '
        'DockSite2
        '
        Me.DockSite2.AccessibleRole = System.Windows.Forms.AccessibleRole.Window
        Me.DockSite2.Dock = System.Windows.Forms.DockStyle.Right
        Me.DockSite2.DocumentDockContainer = New DevComponents.DotNetBar.DocumentDockContainer
        Me.DockSite2.Location = New System.Drawing.Point(1016, 68)
        Me.DockSite2.Name = "DockSite2"
        Me.DockSite2.Size = New System.Drawing.Size(0, 510)
        Me.DockSite2.TabIndex = 1
        Me.DockSite2.TabStop = False
        '
        'DockSite8
        '
        Me.DockSite8.AccessibleRole = System.Windows.Forms.AccessibleRole.Window
        Me.DockSite8.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.DockSite8.Location = New System.Drawing.Point(0, 732)
        Me.DockSite8.Name = "DockSite8"
        Me.DockSite8.Size = New System.Drawing.Size(1016, 0)
        Me.DockSite8.TabIndex = 7
        Me.DockSite8.TabStop = False
        '
        'DockSite5
        '
        Me.DockSite5.AccessibleRole = System.Windows.Forms.AccessibleRole.Window
        Me.DockSite5.Dock = System.Windows.Forms.DockStyle.Left
        Me.DockSite5.Location = New System.Drawing.Point(0, 0)
        Me.DockSite5.Name = "DockSite5"
        Me.DockSite5.Size = New System.Drawing.Size(0, 732)
        Me.DockSite5.TabIndex = 4
        Me.DockSite5.TabStop = False
        '
        'DockSite6
        '
        Me.DockSite6.AccessibleRole = System.Windows.Forms.AccessibleRole.Window
        Me.DockSite6.Dock = System.Windows.Forms.DockStyle.Right
        Me.DockSite6.Location = New System.Drawing.Point(1016, 0)
        Me.DockSite6.Name = "DockSite6"
        Me.DockSite6.Size = New System.Drawing.Size(0, 732)
        Me.DockSite6.TabIndex = 5
        Me.DockSite6.TabStop = False
        '
        'DockSite3
        '
        Me.DockSite3.AccessibleRole = System.Windows.Forms.AccessibleRole.Window
        Me.DockSite3.Controls.Add(Me.Bar1)
        Me.DockSite3.Controls.Add(Me.Bar2)
        Me.DockSite3.Dock = System.Windows.Forms.DockStyle.Top
        Me.DockSite3.DocumentDockContainer = New DevComponents.DotNetBar.DocumentDockContainer(New DevComponents.DotNetBar.DocumentBaseContainer() {CType(New DevComponents.DotNetBar.DocumentBarContainer(Me.Bar1, 221, 65), DevComponents.DotNetBar.DocumentBaseContainer), CType(New DevComponents.DotNetBar.DocumentBarContainer(Me.Bar2, 792, 65), DevComponents.DotNetBar.DocumentBaseContainer)}, DevComponents.DotNetBar.eOrientation.Horizontal)
        Me.DockSite3.Location = New System.Drawing.Point(0, 0)
        Me.DockSite3.Name = "DockSite3"
        Me.DockSite3.Size = New System.Drawing.Size(1016, 68)
        Me.DockSite3.TabIndex = 2
        Me.DockSite3.TabStop = False
        '
        'Bar1
        '
        Me.Bar1.AccessibleDescription = "DotNetBar Bar (Bar1)"
        Me.Bar1.AccessibleName = "DotNetBar Bar"
        Me.Bar1.AccessibleRole = System.Windows.Forms.AccessibleRole.ToolBar
        Me.Bar1.AutoSyncBarCaption = True
        Me.Bar1.CanDockTab = False
        Me.Bar1.CanReorderTabs = False
        Me.Bar1.CanUndock = False
        Me.Bar1.Controls.Add(Me.PanelDockContainer1)
        Me.Bar1.DockOffset = 34
        Me.Bar1.DockTabAlignment = DevComponents.DotNetBar.eTabStripAlignment.Left
        Me.Bar1.Font = New System.Drawing.Font("Verdana", 8.0!)
        Me.Bar1.GrabHandleStyle = DevComponents.DotNetBar.eGrabHandleStyle.Caption
        Me.Bar1.Items.AddRange(New DevComponents.DotNetBar.BaseItem() {Me.DockContainerItem1})
        Me.Bar1.LayoutType = DevComponents.DotNetBar.eLayoutType.DockContainer
        Me.Bar1.Location = New System.Drawing.Point(0, 0)
        Me.Bar1.Name = "Bar1"
        Me.Bar1.Size = New System.Drawing.Size(221, 65)
        Me.Bar1.Stretch = True
        Me.Bar1.Style = DevComponents.DotNetBar.eDotNetBarStyle.Office2007
        Me.Bar1.TabIndex = 0
        Me.Bar1.TabStop = False
        Me.Bar1.Text = "실행 메뉴"
        '
        'PanelDockContainer1
        '
        Me.PanelDockContainer1.ColorSchemeStyle = DevComponents.DotNetBar.eDotNetBarStyle.Office2007
        Me.PanelDockContainer1.Controls.Add(Me.CtlBar)
        Me.PanelDockContainer1.Location = New System.Drawing.Point(3, 23)
        Me.PanelDockContainer1.Name = "PanelDockContainer1"
        Me.PanelDockContainer1.Size = New System.Drawing.Size(215, 39)
        Me.PanelDockContainer1.Style.Alignment = System.Drawing.StringAlignment.Center
        Me.PanelDockContainer1.Style.BackColor1.ColorSchemePart = DevComponents.DotNetBar.eColorSchemePart.BarBackground
        Me.PanelDockContainer1.Style.BackColor2.ColorSchemePart = DevComponents.DotNetBar.eColorSchemePart.BarBackground2
        Me.PanelDockContainer1.Style.BorderColor.ColorSchemePart = DevComponents.DotNetBar.eColorSchemePart.BarDockedBorder
        Me.PanelDockContainer1.Style.ForeColor.ColorSchemePart = DevComponents.DotNetBar.eColorSchemePart.ItemText
        Me.PanelDockContainer1.Style.GradientAngle = 90
        Me.PanelDockContainer1.TabIndex = 0
        '
        'CtlBar
        '
        Me.CtlBar.BarType = DevComponents.DotNetBar.eBarType.DockWindow
        Me.CtlBar.Dock = System.Windows.Forms.DockStyle.Fill
        Me.CtlBar.Font = New System.Drawing.Font("Verdana", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CtlBar.Items.AddRange(New DevComponents.DotNetBar.BaseItem() {Me.FindBtn, Me.SaveBtn, Me.DelBtn, Me.PrtBtn, Me.Excel})
        Me.CtlBar.Location = New System.Drawing.Point(0, 0)
        Me.CtlBar.Name = "CtlBar"
        Me.CtlBar.Size = New System.Drawing.Size(215, 41)
        Me.CtlBar.Stretch = True
        Me.CtlBar.Style = DevComponents.DotNetBar.eDotNetBarStyle.Office2007
        Me.CtlBar.TabIndex = 4
        Me.CtlBar.TabStop = False
        Me.CtlBar.Text = "실행 메뉴"
        '
        'FindBtn
        '
        Me.FindBtn.Image = CType(resources.GetObject("FindBtn.Image"), System.Drawing.Image)
        Me.FindBtn.ImageFixedSize = New System.Drawing.Size(32, 32)
        Me.FindBtn.Name = "FindBtn"
        Me.FindBtn.Text = "ButtonItem1"
        '
        'SaveBtn
        '
        Me.SaveBtn.Image = Global.SMSPLUS.My.Resources.Resources.document_save
        Me.SaveBtn.Name = "SaveBtn"
        Me.SaveBtn.Shortcuts.Add(DevComponents.DotNetBar.eShortcut.CtrlS)
        Me.SaveBtn.Text = "Save"
        Me.SaveBtn.Tooltip = "Save data"
        '
        'DelBtn
        '
        Me.DelBtn.Image = Global.SMSPLUS.My.Resources.Resources.edit_delete
        Me.DelBtn.Name = "DelBtn"
        Me.DelBtn.Shortcuts.Add(DevComponents.DotNetBar.eShortcut.CtrlDel)
        Me.DelBtn.Text = "Delete"
        Me.DelBtn.Tooltip = "Delete Data"
        '
        'PrtBtn
        '
        Me.PrtBtn.Image = Global.SMSPLUS.My.Resources.Resources.document_print
        Me.PrtBtn.Name = "PrtBtn"
        Me.PrtBtn.Shortcuts.Add(DevComponents.DotNetBar.eShortcut.CtrlP)
        Me.PrtBtn.Text = "Print"
        Me.PrtBtn.Tooltip = "Print Document"
        '
        'Excel
        '
        Me.Excel.Image = Global.SMSPLUS.My.Resources.Resources.Excel
        Me.Excel.Name = "Excel"
        Me.Excel.Shortcuts.Add(DevComponents.DotNetBar.eShortcut.CtrlE)
        Me.Excel.Text = "Excel"
        Me.Excel.Tooltip = "Download To Excel"
        '
        'DockContainerItem1
        '
        Me.DockContainerItem1.Control = Me.PanelDockContainer1
        Me.DockContainerItem1.Name = "DockContainerItem1"
        Me.DockContainerItem1.Text = "실행 메뉴"
        '
        'Bar2
        '
        Me.Bar2.AccessibleDescription = "DotNetBar Bar (Bar2)"
        Me.Bar2.AccessibleName = "DotNetBar Bar"
        Me.Bar2.AccessibleRole = System.Windows.Forms.AccessibleRole.ToolBar
        Me.Bar2.AutoSyncBarCaption = True
        Me.Bar2.CanDockTab = False
        Me.Bar2.CanReorderTabs = False
        Me.Bar2.CanUndock = False
        Me.Bar2.Controls.Add(Me.PanelDockContainer2)
        Me.Bar2.DockTabAlignment = DevComponents.DotNetBar.eTabStripAlignment.Left
        Me.Bar2.Font = New System.Drawing.Font("Verdana", 8.0!)
        Me.Bar2.GrabHandleStyle = DevComponents.DotNetBar.eGrabHandleStyle.Caption
        Me.Bar2.Items.AddRange(New DevComponents.DotNetBar.BaseItem() {Me.DockContainerItem2})
        Me.Bar2.LayoutType = DevComponents.DotNetBar.eLayoutType.DockContainer
        Me.Bar2.Location = New System.Drawing.Point(224, 0)
        Me.Bar2.Name = "Bar2"
        Me.Bar2.Size = New System.Drawing.Size(792, 65)
        Me.Bar2.Stretch = True
        Me.Bar2.Style = DevComponents.DotNetBar.eDotNetBarStyle.Office2007
        Me.Bar2.TabIndex = 1
        Me.Bar2.TabStop = False
        Me.Bar2.Text = "조회 조건"
        '
        'PanelDockContainer2
        '
        Me.PanelDockContainer2.ColorSchemeStyle = DevComponents.DotNetBar.eDotNetBarStyle.Office2007
        Me.PanelDockContainer2.Controls.Add(Me.LabelX6)
        Me.PanelDockContainer2.Controls.Add(Me.PartNoTxt)
        Me.PanelDockContainer2.Controls.Add(Me.ToCb)
        Me.PanelDockContainer2.Controls.Add(Me.LabelX7)
        Me.PanelDockContainer2.Controls.Add(Me.FromCb)
        Me.PanelDockContainer2.Controls.Add(Me.POStDate)
        Me.PanelDockContainer2.Controls.Add(Me.POEdDate)
        Me.PanelDockContainer2.Controls.Add(Me.LabelX2)
        Me.PanelDockContainer2.Location = New System.Drawing.Point(3, 23)
        Me.PanelDockContainer2.Name = "PanelDockContainer2"
        Me.PanelDockContainer2.Size = New System.Drawing.Size(786, 39)
        Me.PanelDockContainer2.Style.Alignment = System.Drawing.StringAlignment.Center
        Me.PanelDockContainer2.Style.BackColor1.ColorSchemePart = DevComponents.DotNetBar.eColorSchemePart.BarBackground
        Me.PanelDockContainer2.Style.BackColor2.ColorSchemePart = DevComponents.DotNetBar.eColorSchemePart.BarBackground2
        Me.PanelDockContainer2.Style.BorderColor.ColorSchemePart = DevComponents.DotNetBar.eColorSchemePart.BarDockedBorder
        Me.PanelDockContainer2.Style.ForeColor.ColorSchemePart = DevComponents.DotNetBar.eColorSchemePart.ItemText
        Me.PanelDockContainer2.Style.GradientAngle = 90
        Me.PanelDockContainer2.TabIndex = 0
        '
        'LabelX6
        '
        Me.LabelX6.BackColor = System.Drawing.Color.Transparent
        '
        '
        '
        Me.LabelX6.BackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square
        Me.LabelX6.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LabelX6.Location = New System.Drawing.Point(577, 9)
        Me.LabelX6.Name = "LabelX6"
        Me.LabelX6.Size = New System.Drawing.Size(59, 20)
        Me.LabelX6.TabIndex = 29
        Me.LabelX6.Text = "품목번호"
        Me.LabelX6.TextAlignment = System.Drawing.StringAlignment.Center
        Me.LabelX6.TextLineAlignment = System.Drawing.StringAlignment.Far
        '
        'PartNoTxt
        '
        '
        '
        '
        Me.PartNoTxt.Border.Class = "TextBoxBorder"
        Me.PartNoTxt.Border.CornerType = DevComponents.DotNetBar.eCornerType.Square
        Me.PartNoTxt.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.PartNoTxt.Location = New System.Drawing.Point(641, 9)
        Me.PartNoTxt.Name = "PartNoTxt"
        Me.PartNoTxt.Size = New System.Drawing.Size(129, 21)
        Me.PartNoTxt.TabIndex = 28
        '
        'ToCb
        '
        Me.ToCb.DisplayMember = "Text"
        Me.ToCb.DrawMode = System.Windows.Forms.DrawMode.OwnerDrawFixed
        Me.ToCb.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ToCb.FormattingEnabled = True
        Me.ToCb.ItemHeight = 15
        Me.ToCb.Location = New System.Drawing.Point(444, 9)
        Me.ToCb.Name = "ToCb"
        Me.ToCb.Size = New System.Drawing.Size(128, 21)
        Me.ToCb.TabIndex = 23
        '
        'LabelX7
        '
        Me.LabelX7.BackColor = System.Drawing.Color.Transparent
        '
        '
        '
        Me.LabelX7.BackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square
        Me.LabelX7.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LabelX7.Location = New System.Drawing.Point(179, 12)
        Me.LabelX7.Name = "LabelX7"
        Me.LabelX7.Size = New System.Drawing.Size(15, 15)
        Me.LabelX7.TabIndex = 22
        Me.LabelX7.Text = "~"
        Me.LabelX7.TextAlignment = System.Drawing.StringAlignment.Center
        '
        'FromCb
        '
        Me.FromCb.DisplayMember = "Text"
        Me.FromCb.DrawMode = System.Windows.Forms.DrawMode.OwnerDrawFixed
        Me.FromCb.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FromCb.FormattingEnabled = True
        Me.FromCb.ItemHeight = 15
        Me.FromCb.Location = New System.Drawing.Point(312, 9)
        Me.FromCb.Name = "FromCb"
        Me.FromCb.Size = New System.Drawing.Size(128, 21)
        Me.FromCb.TabIndex = 20
        '
        'POStDate
        '
        '
        '
        '
        Me.POStDate.BackgroundStyle.Class = "DateTimeInputBackground"
        Me.POStDate.BackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square
        Me.POStDate.BackgroundStyle.Font = New System.Drawing.Font("Verdana", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.POStDate.ButtonDropDown.Visible = True
        Me.POStDate.CustomFormat = "yyyy-MM-dd"
        Me.POStDate.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.POStDate.Format = DevComponents.Editors.eDateTimePickerFormat.Custom
        Me.POStDate.IsPopupCalendarOpen = False
        Me.POStDate.Location = New System.Drawing.Point(67, 9)
        '
        '
        '
        Me.POStDate.MonthCalendar.AnnuallyMarkedDates = New Date(-1) {}
        '
        '
        '
        Me.POStDate.MonthCalendar.BackgroundStyle.BackColor = System.Drawing.SystemColors.Window
        Me.POStDate.MonthCalendar.BackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square
        Me.POStDate.MonthCalendar.CalendarDimensions = New System.Drawing.Size(1, 1)
        Me.POStDate.MonthCalendar.ClearButtonVisible = True
        '
        '
        '
        Me.POStDate.MonthCalendar.CommandsBackgroundStyle.BackColor2SchemePart = DevComponents.DotNetBar.eColorSchemePart.BarBackground2
        Me.POStDate.MonthCalendar.CommandsBackgroundStyle.BackColorGradientAngle = 90
        Me.POStDate.MonthCalendar.CommandsBackgroundStyle.BackColorSchemePart = DevComponents.DotNetBar.eColorSchemePart.BarBackground
        Me.POStDate.MonthCalendar.CommandsBackgroundStyle.BorderTop = DevComponents.DotNetBar.eStyleBorderType.Solid
        Me.POStDate.MonthCalendar.CommandsBackgroundStyle.BorderTopColorSchemePart = DevComponents.DotNetBar.eColorSchemePart.BarDockedBorder
        Me.POStDate.MonthCalendar.CommandsBackgroundStyle.BorderTopWidth = 1
        Me.POStDate.MonthCalendar.CommandsBackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square
        Me.POStDate.MonthCalendar.DisplayMonth = New Date(2009, 5, 1, 0, 0, 0, 0)
        Me.POStDate.MonthCalendar.MarkedDates = New Date(-1) {}
        Me.POStDate.MonthCalendar.MonthlyMarkedDates = New Date(-1) {}
        '
        '
        '
        Me.POStDate.MonthCalendar.NavigationBackgroundStyle.BackColor2SchemePart = DevComponents.DotNetBar.eColorSchemePart.PanelBackground2
        Me.POStDate.MonthCalendar.NavigationBackgroundStyle.BackColorGradientAngle = 90
        Me.POStDate.MonthCalendar.NavigationBackgroundStyle.BackColorSchemePart = DevComponents.DotNetBar.eColorSchemePart.PanelBackground
        Me.POStDate.MonthCalendar.NavigationBackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square
        Me.POStDate.MonthCalendar.TodayButtonVisible = True
        Me.POStDate.MonthCalendar.WeeklyMarkedDays = New System.DayOfWeek(-1) {}
        Me.POStDate.Name = "POStDate"
        Me.POStDate.Size = New System.Drawing.Size(112, 21)
        Me.POStDate.TabIndex = 18
        '
        'POEdDate
        '
        '
        '
        '
        Me.POEdDate.BackgroundStyle.Class = "DateTimeInputBackground"
        Me.POEdDate.BackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square
        Me.POEdDate.ButtonDropDown.Visible = True
        Me.POEdDate.CustomFormat = "yyyy-MM-dd"
        Me.POEdDate.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.POEdDate.Format = DevComponents.Editors.eDateTimePickerFormat.Custom
        Me.POEdDate.IsPopupCalendarOpen = False
        Me.POEdDate.Location = New System.Drawing.Point(195, 9)
        '
        '
        '
        Me.POEdDate.MonthCalendar.AnnuallyMarkedDates = New Date(-1) {}
        '
        '
        '
        Me.POEdDate.MonthCalendar.BackgroundStyle.BackColor = System.Drawing.SystemColors.Window
        Me.POEdDate.MonthCalendar.BackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square
        Me.POEdDate.MonthCalendar.CalendarDimensions = New System.Drawing.Size(1, 1)
        Me.POEdDate.MonthCalendar.ClearButtonVisible = True
        '
        '
        '
        Me.POEdDate.MonthCalendar.CommandsBackgroundStyle.BackColor2SchemePart = DevComponents.DotNetBar.eColorSchemePart.BarBackground2
        Me.POEdDate.MonthCalendar.CommandsBackgroundStyle.BackColorGradientAngle = 90
        Me.POEdDate.MonthCalendar.CommandsBackgroundStyle.BackColorSchemePart = DevComponents.DotNetBar.eColorSchemePart.BarBackground
        Me.POEdDate.MonthCalendar.CommandsBackgroundStyle.BorderTop = DevComponents.DotNetBar.eStyleBorderType.Solid
        Me.POEdDate.MonthCalendar.CommandsBackgroundStyle.BorderTopColorSchemePart = DevComponents.DotNetBar.eColorSchemePart.BarDockedBorder
        Me.POEdDate.MonthCalendar.CommandsBackgroundStyle.BorderTopWidth = 1
        Me.POEdDate.MonthCalendar.CommandsBackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square
        Me.POEdDate.MonthCalendar.DisplayMonth = New Date(2009, 5, 1, 0, 0, 0, 0)
        Me.POEdDate.MonthCalendar.MarkedDates = New Date(-1) {}
        Me.POEdDate.MonthCalendar.MonthlyMarkedDates = New Date(-1) {}
        '
        '
        '
        Me.POEdDate.MonthCalendar.NavigationBackgroundStyle.BackColor2SchemePart = DevComponents.DotNetBar.eColorSchemePart.PanelBackground2
        Me.POEdDate.MonthCalendar.NavigationBackgroundStyle.BackColorGradientAngle = 90
        Me.POEdDate.MonthCalendar.NavigationBackgroundStyle.BackColorSchemePart = DevComponents.DotNetBar.eColorSchemePart.PanelBackground
        Me.POEdDate.MonthCalendar.NavigationBackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square
        Me.POEdDate.MonthCalendar.TodayButtonVisible = True
        Me.POEdDate.MonthCalendar.WeeklyMarkedDays = New System.DayOfWeek(-1) {}
        Me.POEdDate.Name = "POEdDate"
        Me.POEdDate.Size = New System.Drawing.Size(112, 21)
        Me.POEdDate.TabIndex = 19
        '
        'LabelX2
        '
        Me.LabelX2.AutoSize = True
        Me.LabelX2.BackColor = System.Drawing.Color.Transparent
        '
        '
        '
        Me.LabelX2.BackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square
        Me.LabelX2.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LabelX2.Location = New System.Drawing.Point(8, 11)
        Me.LabelX2.Name = "LabelX2"
        Me.LabelX2.Size = New System.Drawing.Size(56, 17)
        Me.LabelX2.TabIndex = 7
        Me.LabelX2.Text = "조회기간"
        Me.LabelX2.TextAlignment = System.Drawing.StringAlignment.Center
        '
        'DockContainerItem2
        '
        Me.DockContainerItem2.Control = Me.PanelDockContainer2
        Me.DockContainerItem2.Name = "DockContainerItem2"
        Me.DockContainerItem2.Text = "조회 조건"
        '
        'DockContainerItem5
        '
        Me.DockContainerItem5.Name = "DockContainerItem5"
        Me.DockContainerItem5.Text = "RETURN PART "
        '
        'DockContainerItem6
        '
        Me.DockContainerItem6.Name = "DockContainerItem6"
        Me.DockContainerItem6.Text = "PART ROOM"
        '
        'DockContainerItem4
        '
        Me.DockContainerItem4.Name = "DockContainerItem4"
        Me.DockContainerItem4.Text = "PART ROOM"
        '
        'ControlContainerItem1
        '
        Me.ControlContainerItem1.AllowItemResize = False
        Me.ControlContainerItem1.MenuVisibility = DevComponents.DotNetBar.eMenuVisibility.VisibleAlways
        Me.ControlContainerItem1.Name = "ControlContainerItem1"
        '
        'ControlContainerItem2
        '
        Me.ControlContainerItem2.AllowItemResize = True
        Me.ControlContainerItem2.MenuVisibility = DevComponents.DotNetBar.eMenuVisibility.VisibleAlways
        Me.ControlContainerItem2.Name = "ControlContainerItem2"
        '
        'ContextMenuBar1
        '
        Me.ContextMenuBar1.Font = New System.Drawing.Font("Malgun Gothic", 9.0!)
        Me.ContextMenuBar1.Items.AddRange(New DevComponents.DotNetBar.BaseItem() {Me.CtxSp, Me.CtxSp2})
        Me.ContextMenuBar1.Location = New System.Drawing.Point(471, 355)
        Me.ContextMenuBar1.Name = "ContextMenuBar1"
        Me.ContextMenuBar1.Size = New System.Drawing.Size(185, 27)
        Me.ContextMenuBar1.Stretch = True
        Me.ContextMenuBar1.Style = DevComponents.DotNetBar.eDotNetBarStyle.Office2003
        Me.ContextMenuBar1.TabIndex = 9
        Me.ContextMenuBar1.TabStop = False
        Me.ContextMenuBar1.Text = "ContextMenuBar1"
        '
        'CtxSp
        '
        Me.CtxSp.AutoExpandOnClick = True
        Me.CtxSp.Name = "CtxSp"
        Me.CtxSp.SubItems.AddRange(New DevComponents.DotNetBar.BaseItem() {Me.bDel, Me.ButtonItem1})
        Me.CtxSp.Text = "FPspread3"
        '
        'bDel
        '
        Me.bDel.Image = Global.SMSPLUS.My.Resources.Resources._14_layer_deletelayer
        Me.bDel.Name = "bDel"
        Me.bDel.Text = "&Delete"
        '
        'ButtonItem1
        '
        Me.ButtonItem1.Name = "ButtonItem1"
        Me.ButtonItem1.Text = "Clear"
        '
        'CtxSp2
        '
        Me.CtxSp2.AutoExpandOnClick = True
        Me.CtxSp2.Name = "CtxSp2"
        Me.CtxSp2.SubItems.AddRange(New DevComponents.DotNetBar.BaseItem() {Me.ButtonItem2, Me.cPrint, Me.cExcel})
        Me.CtxSp2.Text = "Fpspread2"
        '
        'ButtonItem2
        '
        Me.ButtonItem2.Name = "ButtonItem2"
        Me.ButtonItem2.Text = "Selected rows transfer"
        '
        'cPrint
        '
        Me.cPrint.Image = Global.SMSPLUS.My.Resources.Resources.agt_print
        Me.cPrint.Name = "cPrint"
        Me.cPrint.Text = "Print"
        '
        'cExcel
        '
        Me.cExcel.Image = Global.SMSPLUS.My.Resources.Resources.sExcel
        Me.cExcel.Name = "cExcel"
        Me.cExcel.Text = "Excel"
        '
        'bSave
        '
        Me.bSave.Image = Global.SMSPLUS.My.Resources.Resources.filesave
        Me.bSave.Name = "bSave"
        Me.bSave.Text = "&Save"
        '
        'bPrint
        '
        Me.bPrint.Image = Global.SMSPLUS.My.Resources.Resources.agt_print
        Me.bPrint.Name = "bPrint"
        Me.bPrint.Text = "&Print"
        '
        'bExcel
        '
        Me.bExcel.Image = Global.SMSPLUS.My.Resources.Resources.sExcel
        Me.bExcel.Name = "bExcel"
        Me.bExcel.Text = "E&xcel"
        Me.bExcel.Tooltip = "Download to Excel"
        '
        'LabelItem2
        '
        Me.LabelItem2.Name = "LabelItem2"
        Me.LabelItem2.Text = "TRANSFERING LIST"
        '
        'XlsBtn1
        '
        Me.XlsBtn1.AccessibleRole = System.Windows.Forms.AccessibleRole.PushButton
        Me.XlsBtn1.ColorTable = DevComponents.DotNetBar.eButtonColor.OrangeWithBackground
        Me.XlsBtn1.Location = New System.Drawing.Point(471, 414)
        Me.XlsBtn1.Name = "XlsBtn1"
        Me.XlsBtn1.Size = New System.Drawing.Size(75, 23)
        Me.XlsBtn1.TabIndex = 23
        Me.XlsBtn1.Text = "E&Xcel"
        '
        'PrtBtn1
        '
        Me.PrtBtn1.AccessibleRole = System.Windows.Forms.AccessibleRole.PushButton
        Me.PrtBtn1.ColorTable = DevComponents.DotNetBar.eButtonColor.OrangeWithBackground
        Me.PrtBtn1.Location = New System.Drawing.Point(471, 385)
        Me.PrtBtn1.Name = "PrtBtn1"
        Me.PrtBtn1.Size = New System.Drawing.Size(75, 23)
        Me.PrtBtn1.TabIndex = 22
        Me.PrtBtn1.Text = "&Print"
        '
        'DelBtn1
        '
        Me.DelBtn1.AccessibleRole = System.Windows.Forms.AccessibleRole.PushButton
        Me.DelBtn1.ColorTable = DevComponents.DotNetBar.eButtonColor.OrangeWithBackground
        Me.DelBtn1.Location = New System.Drawing.Point(471, 356)
        Me.DelBtn1.Name = "DelBtn1"
        Me.DelBtn1.Size = New System.Drawing.Size(75, 23)
        Me.DelBtn1.TabIndex = 21
        Me.DelBtn1.Text = "&Delete"
        '
        'SaveBtn1
        '
        Me.SaveBtn1.AccessibleRole = System.Windows.Forms.AccessibleRole.PushButton
        Me.SaveBtn1.ColorTable = DevComponents.DotNetBar.eButtonColor.OrangeWithBackground
        Me.SaveBtn1.Location = New System.Drawing.Point(471, 327)
        Me.SaveBtn1.Name = "SaveBtn1"
        Me.SaveBtn1.Size = New System.Drawing.Size(75, 23)
        Me.SaveBtn1.TabIndex = 20
        Me.SaveBtn1.Text = "&Save"
        '
        'NewBtn1
        '
        Me.NewBtn1.AccessibleRole = System.Windows.Forms.AccessibleRole.PushButton
        Me.NewBtn1.ColorTable = DevComponents.DotNetBar.eButtonColor.OrangeWithBackground
        Me.NewBtn1.Location = New System.Drawing.Point(471, 298)
        Me.NewBtn1.Name = "NewBtn1"
        Me.NewBtn1.Size = New System.Drawing.Size(75, 23)
        Me.NewBtn1.TabIndex = 19
        Me.NewBtn1.Text = "&New"
        '
        'FrmPartRtn
        '
        Me.ClientSize = New System.Drawing.Size(1016, 732)
        Me.Controls.Add(Me.ContextMenuBar1)
        Me.Controls.Add(Me.DockSite9)
        Me.Controls.Add(Me.DockSite2)
        Me.Controls.Add(Me.DockSite1)
        Me.Controls.Add(Me.DockSite3)
        Me.Controls.Add(Me.DockSite4)
        Me.Controls.Add(Me.DockSite5)
        Me.Controls.Add(Me.DockSite6)
        Me.Controls.Add(Me.DockSite8)
        Me.Controls.Add(Me.XlsBtn1)
        Me.Controls.Add(Me.PrtBtn1)
        Me.Controls.Add(Me.DelBtn1)
        Me.Controls.Add(Me.SaveBtn1)
        Me.Controls.Add(Me.NewBtn1)
        Me.Font = New System.Drawing.Font("Verdana", 8.0!)
        Me.Name = "FrmPartRtn"
        Me.Text = "FrmPartRtn"
        Me.DockSite4.ResumeLayout(False)
        CType(Me.Bar3, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Bar3.ResumeLayout(False)
        Me.PanelDockContainer3.ResumeLayout(False)
        CType(Me.FpSpread1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.FpSpread1_Sheet1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.DockSite9.ResumeLayout(False)
        CType(Me.Bar5, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Bar5.ResumeLayout(False)
        CType(Me.FpSpread3, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.FpSpread3_Sheet1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.Bar8, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.Bar4, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Bar4.ResumeLayout(False)
        CType(Me.FpSpread2, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.FpSpread2_Sheet1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.Bar7, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Bar7.ResumeLayout(False)
        Me.Bar7.PerformLayout()
        Me.DockSite3.ResumeLayout(False)
        CType(Me.Bar1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Bar1.ResumeLayout(False)
        Me.PanelDockContainer1.ResumeLayout(False)
        CType(Me.CtlBar, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.Bar2, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Bar2.ResumeLayout(False)
        Me.PanelDockContainer2.ResumeLayout(False)
        Me.PanelDockContainer2.PerformLayout()
        CType(Me.POStDate, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.POEdDate, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.ContextMenuBar1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents FrmGuiMgr As DevComponents.DotNetBar.DotNetBarManager
    Friend WithEvents DockSite4 As DevComponents.DotNetBar.DockSite
    Friend WithEvents DockSite1 As DevComponents.DotNetBar.DockSite
    Friend WithEvents DockSite2 As DevComponents.DotNetBar.DockSite
    Friend WithEvents DockSite3 As DevComponents.DotNetBar.DockSite
    Friend WithEvents Bar1 As DevComponents.DotNetBar.Bar
    Friend WithEvents DockContainerItem1 As DevComponents.DotNetBar.DockContainerItem
    Friend WithEvents Bar2 As DevComponents.DotNetBar.Bar
    Friend WithEvents PanelDockContainer2 As DevComponents.DotNetBar.PanelDockContainer
    Friend WithEvents DockContainerItem2 As DevComponents.DotNetBar.DockContainerItem
    Friend WithEvents DockSite5 As DevComponents.DotNetBar.DockSite
    Friend WithEvents DockSite6 As DevComponents.DotNetBar.DockSite
    Friend WithEvents DockSite8 As DevComponents.DotNetBar.DockSite
    Friend WithEvents DockSite9 As DevComponents.DotNetBar.DockSite
    Friend WithEvents DockContainerItem4 As DevComponents.DotNetBar.DockContainerItem
    Friend WithEvents DockContainerItem5 As DevComponents.DotNetBar.DockContainerItem
    Friend WithEvents Bar5 As DevComponents.DotNetBar.Bar
    Friend WithEvents Bar4 As DevComponents.DotNetBar.Bar
    Friend WithEvents Bar7 As DevComponents.DotNetBar.Bar
    Friend WithEvents LabelItem1 As DevComponents.DotNetBar.LabelItem
    Friend WithEvents DockContainerItem6 As DevComponents.DotNetBar.DockContainerItem
    Friend WithEvents PanelDockContainer1 As DevComponents.DotNetBar.PanelDockContainer
    Friend WithEvents LabelX2 As DevComponents.DotNetBar.LabelX
    Friend WithEvents LabelX7 As DevComponents.DotNetBar.LabelX
    Friend WithEvents FromCb As DevComponents.DotNetBar.Controls.ComboBoxEx
    Friend WithEvents POStDate As DevComponents.Editors.DateTimeAdv.DateTimeInput
    Friend WithEvents POEdDate As DevComponents.Editors.DateTimeAdv.DateTimeInput
    Friend WithEvents CtlBar As DevComponents.DotNetBar.Bar
    Friend WithEvents SaveBtn As DevComponents.DotNetBar.ButtonItem
    Friend WithEvents DelBtn As DevComponents.DotNetBar.ButtonItem
    Friend WithEvents PrtBtn As DevComponents.DotNetBar.ButtonItem
    Friend WithEvents Excel As DevComponents.DotNetBar.ButtonItem
    Friend WithEvents FpSpread3 As FarPoint.Win.Spread.FpSpread
    Friend WithEvents FpSpread3_Sheet1 As FarPoint.Win.Spread.SheetView
    Friend WithEvents Bar3 As DevComponents.DotNetBar.Bar
    Friend WithEvents PanelDockContainer3 As DevComponents.DotNetBar.PanelDockContainer
    Friend WithEvents FpSpread1 As FarPoint.Win.Spread.FpSpread
    Friend WithEvents FpSpread1_Sheet1 As FarPoint.Win.Spread.SheetView
    Friend WithEvents DockContainerItem3 As DevComponents.DotNetBar.DockContainerItem
    Friend WithEvents FpSpread2 As FarPoint.Win.Spread.FpSpread
    Friend WithEvents FpSpread2_Sheet1 As FarPoint.Win.Spread.SheetView
    Friend WithEvents Splitter1 As System.Windows.Forms.Splitter
    Friend WithEvents PtMdCb As DevComponents.DotNetBar.Controls.ComboBoxEx
    Friend WithEvents LabelX1 As DevComponents.DotNetBar.LabelX
    Friend WithEvents ControlContainerItem4 As DevComponents.DotNetBar.ControlContainerItem
    Friend WithEvents ControlContainerItem3 As DevComponents.DotNetBar.ControlContainerItem
    Friend WithEvents ControlContainerItem1 As DevComponents.DotNetBar.ControlContainerItem
    Friend WithEvents ControlContainerItem2 As DevComponents.DotNetBar.ControlContainerItem
    Friend WithEvents ContextMenuBar1 As DevComponents.DotNetBar.ContextMenuBar
    Friend WithEvents CtxSp As DevComponents.DotNetBar.ButtonItem
    Friend WithEvents bSave As DevComponents.DotNetBar.ButtonItem
    Friend WithEvents bDel As DevComponents.DotNetBar.ButtonItem
    Friend WithEvents bPrint As DevComponents.DotNetBar.ButtonItem
    Friend WithEvents bExcel As DevComponents.DotNetBar.ButtonItem
    Friend WithEvents SaveFileDialog1 As System.Windows.Forms.SaveFileDialog
    Friend WithEvents CtxSp2 As DevComponents.DotNetBar.ButtonItem
    Friend WithEvents cPrint As DevComponents.DotNetBar.ButtonItem
    Friend WithEvents cExcel As DevComponents.DotNetBar.ButtonItem
    Friend WithEvents ToCb As DevComponents.DotNetBar.Controls.ComboBoxEx
    Friend WithEvents Bar8 As DevComponents.DotNetBar.Bar
    Friend WithEvents LabelItem2 As DevComponents.DotNetBar.LabelItem
    Friend WithEvents PtWhCb As DevComponents.DotNetBar.Controls.ComboBoxEx
    Friend WithEvents LabelItem4 As DevComponents.DotNetBar.LabelItem
    Friend WithEvents LabelX4 As DevComponents.DotNetBar.LabelX
    Friend WithEvents LabelX6 As DevComponents.DotNetBar.LabelX
    Friend WithEvents PartNoTxt As DevComponents.DotNetBar.Controls.TextBoxX
    Friend WithEvents XlsBtn1 As DevComponents.DotNetBar.ButtonX
    Friend WithEvents PrtBtn1 As DevComponents.DotNetBar.ButtonX
    Friend WithEvents DelBtn1 As DevComponents.DotNetBar.ButtonX
    Friend WithEvents SaveBtn1 As DevComponents.DotNetBar.ButtonX
    Friend WithEvents NewBtn1 As DevComponents.DotNetBar.ButtonX
    Friend WithEvents ButtonItem1 As DevComponents.DotNetBar.ButtonItem
    Friend WithEvents ButtonItem2 As DevComponents.DotNetBar.ButtonItem
    Friend WithEvents FindBtn As DevComponents.DotNetBar.ButtonItem

End Class
