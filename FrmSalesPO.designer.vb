<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class FrmSalesPO
    Inherits System.Windows.Forms.Form

    'Form은 Dispose를 재정의하여 구성 요소 목록을 정리합니다.
    <System.Diagnostics.DebuggerNonUserCode()>
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
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim EnhancedScrollBarRenderer1 As FarPoint.Win.Spread.EnhancedScrollBarRenderer = New FarPoint.Win.Spread.EnhancedScrollBarRenderer()
        Dim EnhancedScrollBarRenderer2 As FarPoint.Win.Spread.EnhancedScrollBarRenderer = New FarPoint.Win.Spread.EnhancedScrollBarRenderer()
        Dim EnhancedScrollBarRenderer3 As FarPoint.Win.Spread.EnhancedScrollBarRenderer = New FarPoint.Win.Spread.EnhancedScrollBarRenderer()
        Dim EnhancedScrollBarRenderer4 As FarPoint.Win.Spread.EnhancedScrollBarRenderer = New FarPoint.Win.Spread.EnhancedScrollBarRenderer()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(FrmSalesPO))
        Me.FrmGuiMgr = New DevComponents.DotNetBar.DotNetBarManager(Me.components)
        Me.DockSite4 = New DevComponents.DotNetBar.DockSite()
        Me.DockSite9 = New DevComponents.DotNetBar.DockSite()
        Me.Bar12 = New DevComponents.DotNetBar.Bar()
        Me.PanelDockContainer7 = New DevComponents.DotNetBar.PanelDockContainer()
        Me.FpSpread1 = New FarPoint.Win.Spread.FpSpread()
        Me.FpSpread1_Sheet1 = New FarPoint.Win.Spread.SheetView()
        Me.DockContainerItem4 = New DevComponents.DotNetBar.DockContainerItem()
        Me.Bar3 = New DevComponents.DotNetBar.Bar()
        Me.PanelDockContainer8 = New DevComponents.DotNetBar.PanelDockContainer()
        Me.FpSpread2 = New FarPoint.Win.Spread.FpSpread()
        Me.FpSpread2_Sheet1 = New FarPoint.Win.Spread.SheetView()
        Me.DockContainerItem5 = New DevComponents.DotNetBar.DockContainerItem()
        Me.DockSite1 = New DevComponents.DotNetBar.DockSite()
        Me.Bar2 = New DevComponents.DotNetBar.Bar()
        Me.PanelDockContainer2 = New DevComponents.DotNetBar.PanelDockContainer()
        Me.ContextMenuBar1 = New DevComponents.DotNetBar.ContextMenuBar()
        Me.CtxSp = New DevComponents.DotNetBar.ButtonItem()
        Me.bDel = New DevComponents.DotNetBar.ButtonItem()
        Me.bClear = New DevComponents.DotNetBar.ButtonItem()
        Me.bClosing = New DevComponents.DotNetBar.ButtonItem()
        Me.bsClose = New DevComponents.DotNetBar.ButtonItem()
        Me.bsCancled = New DevComponents.DotNetBar.ButtonItem()
        Me.ButtonItem1 = New DevComponents.DotNetBar.ButtonItem()
        Me.ButtonItem2 = New DevComponents.DotNetBar.ButtonItem()
        Me.ButtonItem5 = New DevComponents.DotNetBar.ButtonItem()
        Me.CtxSp2 = New DevComponents.DotNetBar.ButtonItem()
        Me.ButtonItem4 = New DevComponents.DotNetBar.ButtonItem()
        Me.cDel = New DevComponents.DotNetBar.ButtonItem()
        Me.cClear = New DevComponents.DotNetBar.ButtonItem()
        Me.CtxSp3 = New DevComponents.DotNetBar.ButtonItem()
        Me.dAllDown = New DevComponents.DotNetBar.ButtonItem()
        Me.dSelDown = New DevComponents.DotNetBar.ButtonItem()
        Me.dStatus = New DevComponents.DotNetBar.ButtonItem()
        Me.dsClosed = New DevComponents.DotNetBar.ButtonItem()
        Me.dsCancled = New DevComponents.DotNetBar.ButtonItem()
        Me.ButtonItem3 = New DevComponents.DotNetBar.ButtonItem()
        Me.PartList = New DevComponents.DotNetBar.Controls.ListViewEx()
        Me.ColumnHeader1 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.ColumnHeader2 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.ColumnHeader3 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.ColumnHeader4 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.ColumnHeader5 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.ColumnHeader6 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.Bar4 = New DevComponents.DotNetBar.Bar()
        Me.TextBoxX1 = New DevComponents.DotNetBar.Controls.TextBoxX()
        Me.LabelX1 = New DevComponents.DotNetBar.LabelX()
        Me.PtMdCb = New DevComponents.DotNetBar.Controls.ComboBoxEx()
        Me.DockContainerItem2 = New DevComponents.DotNetBar.DockContainerItem()
        Me.DockSite2 = New DevComponents.DotNetBar.DockSite()
        Me.DockSite8 = New DevComponents.DotNetBar.DockSite()
        Me.DockSite5 = New DevComponents.DotNetBar.DockSite()
        Me.DockSite6 = New DevComponents.DotNetBar.DockSite()
        Me.DockSite7 = New DevComponents.DotNetBar.DockSite()
        Me.DockSite3 = New DevComponents.DotNetBar.DockSite()
        Me.Bar1 = New DevComponents.DotNetBar.Bar()
        Me.PanelDockContainer1 = New DevComponents.DotNetBar.PanelDockContainer()
        Me.CtlBar = New DevComponents.DotNetBar.Bar()
        Me.FindBtn = New DevComponents.DotNetBar.ButtonItem()
        Me.NewBtn = New DevComponents.DotNetBar.ButtonItem()
        Me.SaveBtn = New DevComponents.DotNetBar.ButtonItem()
        Me.DelBtn = New DevComponents.DotNetBar.ButtonItem()
        Me.PrtBtn = New DevComponents.DotNetBar.ButtonItem()
        Me.Excel = New DevComponents.DotNetBar.ButtonItem()
        Me.DockContainerItem1 = New DevComponents.DotNetBar.DockContainerItem()
        Me.Bar8 = New DevComponents.DotNetBar.Bar()
        Me.PanelDockContainer3 = New DevComponents.DotNetBar.PanelDockContainer()
        Me.ComboBoxEx1 = New DevComponents.DotNetBar.Controls.ComboBoxEx()
        Me.LabelX7 = New DevComponents.DotNetBar.LabelX()
        Me.LabelX5 = New DevComponents.DotNetBar.LabelX()
        Me.LabelX2 = New DevComponents.DotNetBar.LabelX()
        Me.ModelCb = New DevComponents.DotNetBar.Controls.ComboBoxEx()
        Me.POStDate = New DevComponents.Editors.DateTimeAdv.DateTimeInput()
        Me.POEdDate = New DevComponents.Editors.DateTimeAdv.DateTimeInput()
        Me.SupCb = New DevComponents.DotNetBar.Controls.ComboBoxEx()
        Me.StatCb = New DevComponents.DotNetBar.Controls.ComboBoxEx()
        Me.LabelX3 = New DevComponents.DotNetBar.LabelX()
        Me.DockContainerItem6 = New DevComponents.DotNetBar.DockContainerItem()
        Me.DockContainerItem10 = New DevComponents.DotNetBar.DockContainerItem()
        Me.DockContainerItem3 = New DevComponents.DotNetBar.DockContainerItem()
        Me.SaveFileDialog1 = New System.Windows.Forms.SaveFileDialog()
        Me.bNew = New DevComponents.DotNetBar.ButtonItem()
        Me.bSave = New DevComponents.DotNetBar.ButtonItem()
        Me.bPrint = New DevComponents.DotNetBar.ButtonItem()
        Me.bExcel = New DevComponents.DotNetBar.ButtonItem()
        Me.XlsBtn1 = New DevComponents.DotNetBar.ButtonX()
        Me.PrtBtn1 = New DevComponents.DotNetBar.ButtonX()
        Me.DelBtn1 = New DevComponents.DotNetBar.ButtonX()
        Me.SaveBtn1 = New DevComponents.DotNetBar.ButtonX()
        Me.NewBtn1 = New DevComponents.DotNetBar.ButtonX()
        Me.FpSpread4 = New FarPoint.Win.Spread.FpSpread()
        Me.FpSpread4_Sheet1 = New FarPoint.Win.Spread.SheetView()
        Me.LabelItem2 = New DevComponents.DotNetBar.LabelItem()
        Me.LabelItem3 = New DevComponents.DotNetBar.LabelItem()
        Me.LabelItem4 = New DevComponents.DotNetBar.LabelItem()
        Me.StyleManager1 = New DevComponents.DotNetBar.StyleManager(Me.components)
        Me.DockSite9.SuspendLayout()
        CType(Me.Bar12, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Bar12.SuspendLayout()
        Me.PanelDockContainer7.SuspendLayout()
        CType(Me.FpSpread1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.FpSpread1_Sheet1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.Bar3, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Bar3.SuspendLayout()
        Me.PanelDockContainer8.SuspendLayout()
        CType(Me.FpSpread2, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.FpSpread2_Sheet1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.DockSite1.SuspendLayout()
        CType(Me.Bar2, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Bar2.SuspendLayout()
        Me.PanelDockContainer2.SuspendLayout()
        CType(Me.ContextMenuBar1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.Bar4, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Bar4.SuspendLayout()
        Me.DockSite3.SuspendLayout()
        CType(Me.Bar1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Bar1.SuspendLayout()
        Me.PanelDockContainer1.SuspendLayout()
        CType(Me.CtlBar, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.Bar8, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Bar8.SuspendLayout()
        Me.PanelDockContainer3.SuspendLayout()
        CType(Me.POStDate, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.POEdDate, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.FpSpread4, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.FpSpread4_Sheet1, System.ComponentModel.ISupportInitialize).BeginInit()
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
        Me.FrmGuiMgr.ToolbarTopDockSite = Me.DockSite7
        Me.FrmGuiMgr.TopDockSite = Me.DockSite3
        '
        'DockSite4
        '
        Me.DockSite4.AccessibleRole = System.Windows.Forms.AccessibleRole.Window
        Me.DockSite4.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.DockSite4.DocumentDockContainer = New DevComponents.DotNetBar.DocumentDockContainer()
        Me.DockSite4.Location = New System.Drawing.Point(0, 732)
        Me.DockSite4.Name = "DockSite4"
        Me.DockSite4.Size = New System.Drawing.Size(1270, 0)
        Me.DockSite4.TabIndex = 3
        Me.DockSite4.TabStop = False
        '
        'DockSite9
        '
        Me.DockSite9.AccessibleRole = System.Windows.Forms.AccessibleRole.Window
        Me.DockSite9.Controls.Add(Me.Bar12)
        Me.DockSite9.Controls.Add(Me.Bar3)
        Me.DockSite9.Dock = System.Windows.Forms.DockStyle.Fill
        Me.DockSite9.DocumentDockContainer = New DevComponents.DotNetBar.DocumentDockContainer(New DevComponents.DotNetBar.DocumentBaseContainer() {CType(New DevComponents.DotNetBar.DocumentBarContainer(Me.Bar12, 609, 158), DevComponents.DotNetBar.DocumentBaseContainer), CType(New DevComponents.DotNetBar.DocumentBarContainer(Me.Bar3, 609, 158), DevComponents.DotNetBar.DocumentBaseContainer)}, DevComponents.DotNetBar.eOrientation.Vertical)
        Me.DockSite9.Location = New System.Drawing.Point(407, 66)
        Me.DockSite9.Name = "DockSite9"
        Me.DockSite9.Size = New System.Drawing.Size(863, 666)
        Me.DockSite9.TabIndex = 8
        Me.DockSite9.TabStop = False
        '
        'Bar12
        '
        Me.Bar12.AccessibleDescription = "DotNetBar Bar (Bar12)"
        Me.Bar12.AccessibleName = "DotNetBar Bar"
        Me.Bar12.AccessibleRole = System.Windows.Forms.AccessibleRole.ToolBar
        Me.Bar12.AlwaysDisplayDockTab = True
        Me.Bar12.AutoCreateCaptionMenu = False
        Me.Bar12.AutoHide = True
        Me.Bar12.AutoHideTabTextAlwaysVisible = True
        Me.Bar12.AutoSyncBarCaption = True
        Me.Bar12.BarType = DevComponents.DotNetBar.eBarType.MenuBar
        Me.Bar12.CanDockDocument = True
        Me.Bar12.CanHide = True
        Me.Bar12.CloseSingleTab = True
        Me.Bar12.Controls.Add(Me.PanelDockContainer7)
        Me.Bar12.DockTabAlignment = DevComponents.DotNetBar.eTabStripAlignment.Top
        Me.Bar12.Font = New System.Drawing.Font("Verdana", 8.0!)
        Me.Bar12.Items.AddRange(New DevComponents.DotNetBar.BaseItem() {Me.DockContainerItem4})
        Me.Bar12.LayoutType = DevComponents.DotNetBar.eLayoutType.DockContainer
        Me.Bar12.Location = New System.Drawing.Point(0, 0)
        Me.Bar12.Name = "Bar12"
        Me.Bar12.SelectedDockTab = 0
        Me.Bar12.Size = New System.Drawing.Size(863, 329)
        Me.Bar12.Stretch = True
        Me.Bar12.Style = DevComponents.DotNetBar.eDotNetBarStyle.Office2007
        Me.Bar12.TabIndex = 1
        Me.Bar12.TabNavigation = True
        Me.Bar12.TabStop = False
        Me.Bar12.Text = "수주 개요"
        '
        'PanelDockContainer7
        '
        Me.PanelDockContainer7.ColorSchemeStyle = DevComponents.DotNetBar.eDotNetBarStyle.Office2007
        Me.PanelDockContainer7.Controls.Add(Me.FpSpread1)
        Me.PanelDockContainer7.Location = New System.Drawing.Point(3, 28)
        Me.PanelDockContainer7.Name = "PanelDockContainer7"
        Me.PanelDockContainer7.Size = New System.Drawing.Size(857, 298)
        Me.PanelDockContainer7.Style.Alignment = System.Drawing.StringAlignment.Center
        Me.PanelDockContainer7.Style.BackColor1.ColorSchemePart = DevComponents.DotNetBar.eColorSchemePart.BarBackground
        Me.PanelDockContainer7.Style.ForeColor.ColorSchemePart = DevComponents.DotNetBar.eColorSchemePart.ItemText
        Me.PanelDockContainer7.Style.GradientAngle = 90
        Me.PanelDockContainer7.TabIndex = 1
        '
        'FpSpread1
        '
        Me.FpSpread1.AccessibleDescription = ""
        Me.FpSpread1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.FpSpread1.HorizontalScrollBar.Buttons = New FarPoint.Win.Spread.FpScrollBarButtonCollection("BackwardLineButton,ThumbTrack,ForwardLineButton")
        Me.FpSpread1.HorizontalScrollBar.Name = ""
        EnhancedScrollBarRenderer1.ArrowColor = System.Drawing.Color.DarkSlateGray
        EnhancedScrollBarRenderer1.ArrowHoveredColor = System.Drawing.Color.DarkSlateGray
        EnhancedScrollBarRenderer1.ArrowSelectedColor = System.Drawing.Color.DarkSlateGray
        EnhancedScrollBarRenderer1.ButtonBackgroundColor = System.Drawing.Color.CadetBlue
        EnhancedScrollBarRenderer1.ButtonBorderColor = System.Drawing.Color.SlateGray
        EnhancedScrollBarRenderer1.ButtonHoveredBackgroundColor = System.Drawing.Color.SlateGray
        EnhancedScrollBarRenderer1.ButtonHoveredBorderColor = System.Drawing.Color.DarkGray
        EnhancedScrollBarRenderer1.ButtonSelectedBackgroundColor = System.Drawing.Color.DarkGray
        EnhancedScrollBarRenderer1.ButtonSelectedBorderColor = System.Drawing.Color.CadetBlue
        EnhancedScrollBarRenderer1.TrackBarBackgroundColor = System.Drawing.Color.CadetBlue
        EnhancedScrollBarRenderer1.TrackBarSelectedBackgroundColor = System.Drawing.Color.SlateGray
        Me.FpSpread1.HorizontalScrollBar.Renderer = EnhancedScrollBarRenderer1
        Me.FpSpread1.HorizontalScrollBar.TabIndex = 4
        Me.FpSpread1.Location = New System.Drawing.Point(0, 0)
        Me.FpSpread1.Name = "FpSpread1"
        Me.FpSpread1.Sheets.AddRange(New FarPoint.Win.Spread.SheetView() {Me.FpSpread1_Sheet1})
        Me.FpSpread1.Size = New System.Drawing.Size(857, 298)
        Me.FpSpread1.Skin = FarPoint.Win.Spread.DefaultSpreadSkins.Seashell
        Me.FpSpread1.TabIndex = 5
        Me.FpSpread1.VerticalScrollBar.Buttons = New FarPoint.Win.Spread.FpScrollBarButtonCollection("BackwardLineButton,ThumbTrack,ForwardLineButton")
        Me.FpSpread1.VerticalScrollBar.Name = ""
        EnhancedScrollBarRenderer2.ArrowColor = System.Drawing.Color.DarkSlateGray
        EnhancedScrollBarRenderer2.ArrowHoveredColor = System.Drawing.Color.DarkSlateGray
        EnhancedScrollBarRenderer2.ArrowSelectedColor = System.Drawing.Color.DarkSlateGray
        EnhancedScrollBarRenderer2.ButtonBackgroundColor = System.Drawing.Color.CadetBlue
        EnhancedScrollBarRenderer2.ButtonBorderColor = System.Drawing.Color.SlateGray
        EnhancedScrollBarRenderer2.ButtonHoveredBackgroundColor = System.Drawing.Color.SlateGray
        EnhancedScrollBarRenderer2.ButtonHoveredBorderColor = System.Drawing.Color.DarkGray
        EnhancedScrollBarRenderer2.ButtonSelectedBackgroundColor = System.Drawing.Color.DarkGray
        EnhancedScrollBarRenderer2.ButtonSelectedBorderColor = System.Drawing.Color.CadetBlue
        EnhancedScrollBarRenderer2.TrackBarBackgroundColor = System.Drawing.Color.CadetBlue
        EnhancedScrollBarRenderer2.TrackBarSelectedBackgroundColor = System.Drawing.Color.SlateGray
        Me.FpSpread1.VerticalScrollBar.Renderer = EnhancedScrollBarRenderer2
        Me.FpSpread1.VerticalScrollBar.TabIndex = 5
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
        Me.FpSpread1_Sheet1.ColumnFooter.DefaultStyle.NoteIndicatorColor = System.Drawing.Color.Red
        Me.FpSpread1_Sheet1.ColumnFooter.DefaultStyle.Parent = "ColumnHeaderSeashell"
        Me.FpSpread1_Sheet1.ColumnFooterSheetCornerStyle.NoteIndicatorColor = System.Drawing.Color.Red
        Me.FpSpread1_Sheet1.ColumnFooterSheetCornerStyle.Parent = "CornerSeashell"
        Me.FpSpread1_Sheet1.ColumnHeader.DefaultStyle.NoteIndicatorColor = System.Drawing.Color.Red
        Me.FpSpread1_Sheet1.ColumnHeader.DefaultStyle.Parent = "ColumnHeaderSeashell"
        Me.FpSpread1_Sheet1.RowHeader.Columns.Default.Resizable = False
        Me.FpSpread1_Sheet1.RowHeader.DefaultStyle.NoteIndicatorColor = System.Drawing.Color.Red
        Me.FpSpread1_Sheet1.RowHeader.DefaultStyle.Parent = "RowHeaderSeashell"
        Me.FpSpread1_Sheet1.SheetCornerStyle.NoteIndicatorColor = System.Drawing.Color.Red
        Me.FpSpread1_Sheet1.SheetCornerStyle.Parent = "CornerSeashell"
        Me.FpSpread1_Sheet1.ReferenceStyle = FarPoint.Win.Spread.Model.ReferenceStyle.A1
        '
        'DockContainerItem4
        '
        Me.DockContainerItem4.Control = Me.PanelDockContainer7
        Me.DockContainerItem4.Name = "DockContainerItem4"
        Me.DockContainerItem4.Text = "수주 개요"
        '
        'Bar3
        '
        Me.Bar3.AccessibleDescription = "DotNetBar Bar (Bar3)"
        Me.Bar3.AccessibleName = "DotNetBar Bar"
        Me.Bar3.AccessibleRole = System.Windows.Forms.AccessibleRole.ToolBar
        Me.Bar3.AlwaysDisplayDockTab = True
        Me.Bar3.CanDockDocument = True
        Me.Bar3.CanHide = True
        Me.Bar3.Controls.Add(Me.PanelDockContainer8)
        Me.Bar3.DockTabAlignment = DevComponents.DotNetBar.eTabStripAlignment.Top
        Me.Bar3.Font = New System.Drawing.Font("Verdana", 8.0!)
        Me.Bar3.Items.AddRange(New DevComponents.DotNetBar.BaseItem() {Me.DockContainerItem5})
        Me.Bar3.LayoutType = DevComponents.DotNetBar.eLayoutType.DockContainer
        Me.Bar3.Location = New System.Drawing.Point(0, 332)
        Me.Bar3.Name = "Bar3"
        Me.Bar3.SelectedDockTab = 0
        Me.Bar3.Size = New System.Drawing.Size(863, 334)
        Me.Bar3.Stretch = True
        Me.Bar3.Style = DevComponents.DotNetBar.eDotNetBarStyle.Office2007
        Me.Bar3.TabIndex = 2
        Me.Bar3.TabStop = False
        '
        'PanelDockContainer8
        '
        Me.PanelDockContainer8.ColorSchemeStyle = DevComponents.DotNetBar.eDotNetBarStyle.Office2007
        Me.PanelDockContainer8.Controls.Add(Me.FpSpread2)
        Me.PanelDockContainer8.Location = New System.Drawing.Point(3, 28)
        Me.PanelDockContainer8.Name = "PanelDockContainer8"
        Me.PanelDockContainer8.Size = New System.Drawing.Size(857, 303)
        Me.PanelDockContainer8.Style.Alignment = System.Drawing.StringAlignment.Center
        Me.PanelDockContainer8.Style.BackColor1.ColorSchemePart = DevComponents.DotNetBar.eColorSchemePart.BarBackground
        Me.PanelDockContainer8.Style.ForeColor.ColorSchemePart = DevComponents.DotNetBar.eColorSchemePart.ItemText
        Me.PanelDockContainer8.Style.GradientAngle = 90
        Me.PanelDockContainer8.TabIndex = 3
        '
        'FpSpread2
        '
        Me.FpSpread2.AccessibleDescription = ""
        Me.FpSpread2.AllowDragDrop = True
        Me.FpSpread2.AllowDragFill = True
        Me.FpSpread2.AllowDrop = True
        Me.FpSpread2.Dock = System.Windows.Forms.DockStyle.Fill
        Me.FpSpread2.EditModeReplace = True
        Me.FpSpread2.HorizontalScrollBar.Buttons = New FarPoint.Win.Spread.FpScrollBarButtonCollection("BackwardLineButton,ThumbTrack,ForwardLineButton")
        Me.FpSpread2.HorizontalScrollBar.Name = ""
        EnhancedScrollBarRenderer3.ArrowColor = System.Drawing.Color.DarkSlateGray
        EnhancedScrollBarRenderer3.ArrowHoveredColor = System.Drawing.Color.DarkSlateGray
        EnhancedScrollBarRenderer3.ArrowSelectedColor = System.Drawing.Color.DarkSlateGray
        EnhancedScrollBarRenderer3.ButtonBackgroundColor = System.Drawing.Color.CadetBlue
        EnhancedScrollBarRenderer3.ButtonBorderColor = System.Drawing.Color.SlateGray
        EnhancedScrollBarRenderer3.ButtonHoveredBackgroundColor = System.Drawing.Color.SlateGray
        EnhancedScrollBarRenderer3.ButtonHoveredBorderColor = System.Drawing.Color.DarkGray
        EnhancedScrollBarRenderer3.ButtonSelectedBackgroundColor = System.Drawing.Color.DarkGray
        EnhancedScrollBarRenderer3.ButtonSelectedBorderColor = System.Drawing.Color.CadetBlue
        EnhancedScrollBarRenderer3.TrackBarBackgroundColor = System.Drawing.Color.CadetBlue
        EnhancedScrollBarRenderer3.TrackBarSelectedBackgroundColor = System.Drawing.Color.SlateGray
        Me.FpSpread2.HorizontalScrollBar.Renderer = EnhancedScrollBarRenderer3
        Me.FpSpread2.HorizontalScrollBar.TabIndex = 4
        Me.FpSpread2.Location = New System.Drawing.Point(0, 0)
        Me.FpSpread2.Name = "FpSpread2"
        Me.FpSpread2.Sheets.AddRange(New FarPoint.Win.Spread.SheetView() {Me.FpSpread2_Sheet1})
        Me.FpSpread2.Size = New System.Drawing.Size(857, 303)
        Me.FpSpread2.Skin = FarPoint.Win.Spread.DefaultSpreadSkins.Seashell
        Me.FpSpread2.TabIndex = 4
        Me.FpSpread2.VerticalScrollBar.Buttons = New FarPoint.Win.Spread.FpScrollBarButtonCollection("BackwardLineButton,ThumbTrack,ForwardLineButton")
        Me.FpSpread2.VerticalScrollBar.Name = ""
        EnhancedScrollBarRenderer4.ArrowColor = System.Drawing.Color.DarkSlateGray
        EnhancedScrollBarRenderer4.ArrowHoveredColor = System.Drawing.Color.DarkSlateGray
        EnhancedScrollBarRenderer4.ArrowSelectedColor = System.Drawing.Color.DarkSlateGray
        EnhancedScrollBarRenderer4.ButtonBackgroundColor = System.Drawing.Color.CadetBlue
        EnhancedScrollBarRenderer4.ButtonBorderColor = System.Drawing.Color.SlateGray
        EnhancedScrollBarRenderer4.ButtonHoveredBackgroundColor = System.Drawing.Color.SlateGray
        EnhancedScrollBarRenderer4.ButtonHoveredBorderColor = System.Drawing.Color.DarkGray
        EnhancedScrollBarRenderer4.ButtonSelectedBackgroundColor = System.Drawing.Color.DarkGray
        EnhancedScrollBarRenderer4.ButtonSelectedBorderColor = System.Drawing.Color.CadetBlue
        EnhancedScrollBarRenderer4.TrackBarBackgroundColor = System.Drawing.Color.CadetBlue
        EnhancedScrollBarRenderer4.TrackBarSelectedBackgroundColor = System.Drawing.Color.SlateGray
        Me.FpSpread2.VerticalScrollBar.Renderer = EnhancedScrollBarRenderer4
        Me.FpSpread2.VerticalScrollBar.TabIndex = 5
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
        Me.FpSpread2_Sheet1.ColumnFooter.DefaultStyle.NoteIndicatorColor = System.Drawing.Color.Red
        Me.FpSpread2_Sheet1.ColumnFooter.DefaultStyle.Parent = "ColumnHeaderSeashell"
        Me.FpSpread2_Sheet1.ColumnFooterSheetCornerStyle.NoteIndicatorColor = System.Drawing.Color.Red
        Me.FpSpread2_Sheet1.ColumnFooterSheetCornerStyle.Parent = "CornerSeashell"
        Me.FpSpread2_Sheet1.ColumnHeader.DefaultStyle.NoteIndicatorColor = System.Drawing.Color.Red
        Me.FpSpread2_Sheet1.ColumnHeader.DefaultStyle.Parent = "ColumnHeaderSeashell"
        Me.FpSpread2_Sheet1.RowHeader.DefaultStyle.NoteIndicatorColor = System.Drawing.Color.Red
        Me.FpSpread2_Sheet1.RowHeader.DefaultStyle.Parent = "RowHeaderSeashell"
        Me.FpSpread2_Sheet1.SheetCornerStyle.NoteIndicatorColor = System.Drawing.Color.Red
        Me.FpSpread2_Sheet1.SheetCornerStyle.Parent = "CornerSeashell"
        Me.FpSpread2_Sheet1.ReferenceStyle = FarPoint.Win.Spread.Model.ReferenceStyle.A1
        '
        'DockContainerItem5
        '
        Me.DockContainerItem5.Control = Me.PanelDockContainer8
        Me.DockContainerItem5.Name = "DockContainerItem5"
        Me.DockContainerItem5.Text = "수주 상세"
        '
        'DockSite1
        '
        Me.DockSite1.AccessibleRole = System.Windows.Forms.AccessibleRole.Window
        Me.DockSite1.Controls.Add(Me.Bar2)
        Me.DockSite1.Dock = System.Windows.Forms.DockStyle.Left
        Me.DockSite1.DocumentDockContainer = New DevComponents.DotNetBar.DocumentDockContainer(New DevComponents.DotNetBar.DocumentBaseContainer() {CType(New DevComponents.DotNetBar.DocumentBarContainer(Me.Bar2, 404, 666), DevComponents.DotNetBar.DocumentBaseContainer)}, DevComponents.DotNetBar.eOrientation.Horizontal)
        Me.DockSite1.Location = New System.Drawing.Point(0, 66)
        Me.DockSite1.Name = "DockSite1"
        Me.DockSite1.Size = New System.Drawing.Size(407, 666)
        Me.DockSite1.TabIndex = 0
        Me.DockSite1.TabStop = False
        Me.DockSite1.Text = "PART LIST"
        '
        'Bar2
        '
        Me.Bar2.AccessibleDescription = "DotNetBar Bar (Bar2)"
        Me.Bar2.AccessibleName = "DotNetBar Bar"
        Me.Bar2.AccessibleRole = System.Windows.Forms.AccessibleRole.ToolBar
        Me.Bar2.AutoSyncBarCaption = True
        Me.Bar2.CloseSingleTab = True
        Me.Bar2.Controls.Add(Me.PanelDockContainer2)
        Me.Bar2.Font = New System.Drawing.Font("Verdana", 8.0!)
        Me.Bar2.GrabHandleStyle = DevComponents.DotNetBar.eGrabHandleStyle.Caption
        Me.Bar2.Items.AddRange(New DevComponents.DotNetBar.BaseItem() {Me.DockContainerItem2})
        Me.Bar2.LayoutType = DevComponents.DotNetBar.eLayoutType.DockContainer
        Me.Bar2.Location = New System.Drawing.Point(0, 0)
        Me.Bar2.Name = "Bar2"
        Me.Bar2.SingleLineColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.Bar2.Size = New System.Drawing.Size(404, 666)
        Me.Bar2.Stretch = True
        Me.Bar2.Style = DevComponents.DotNetBar.eDotNetBarStyle.Office2007
        Me.Bar2.TabIndex = 0
        Me.Bar2.TabStop = False
        Me.Bar2.Text = "DockContainerItem2"
        '
        'PanelDockContainer2
        '
        Me.PanelDockContainer2.ColorSchemeStyle = DevComponents.DotNetBar.eDotNetBarStyle.Office2007
        Me.PanelDockContainer2.Controls.Add(Me.ContextMenuBar1)
        Me.PanelDockContainer2.Controls.Add(Me.PartList)
        Me.PanelDockContainer2.Controls.Add(Me.Bar4)
        Me.PanelDockContainer2.Location = New System.Drawing.Point(3, 23)
        Me.PanelDockContainer2.Name = "PanelDockContainer2"
        Me.PanelDockContainer2.Size = New System.Drawing.Size(398, 640)
        Me.PanelDockContainer2.Style.Alignment = System.Drawing.StringAlignment.Center
        Me.PanelDockContainer2.Style.BackColor1.ColorSchemePart = DevComponents.DotNetBar.eColorSchemePart.BarBackground
        Me.PanelDockContainer2.Style.ForeColor.ColorSchemePart = DevComponents.DotNetBar.eColorSchemePart.ItemText
        Me.PanelDockContainer2.Style.GradientAngle = 90
        Me.PanelDockContainer2.TabIndex = 0
        Me.PanelDockContainer2.Text = "PART LIST"
        '
        'ContextMenuBar1
        '
        Me.ContextMenuBar1.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ContextMenuBar1.Items.AddRange(New DevComponents.DotNetBar.BaseItem() {Me.CtxSp, Me.CtxSp2, Me.CtxSp3})
        Me.ContextMenuBar1.Location = New System.Drawing.Point(30, 56)
        Me.ContextMenuBar1.Name = "ContextMenuBar1"
        Me.ContextMenuBar1.Size = New System.Drawing.Size(288, 25)
        Me.ContextMenuBar1.Stretch = True
        Me.ContextMenuBar1.Style = DevComponents.DotNetBar.eDotNetBarStyle.Office2007
        Me.ContextMenuBar1.TabIndex = 9
        Me.ContextMenuBar1.TabStop = False
        Me.ContextMenuBar1.Text = "ContextMenuBar1"
        '
        'CtxSp
        '
        Me.CtxSp.AutoExpandOnClick = True
        Me.CtxSp.Name = "CtxSp"
        Me.CtxSp.SubItems.AddRange(New DevComponents.DotNetBar.BaseItem() {Me.bDel, Me.bClear, Me.bClosing, Me.ButtonItem1, Me.ButtonItem2, Me.ButtonItem5})
        Me.CtxSp.Text = "FPspread2"
        '
        'bDel
        '
        Me.bDel.Image = Global.SMSPLUS.My.Resources.Resources._14_layer_deletelayer
        Me.bDel.Name = "bDel"
        Me.bDel.Text = "&Delete"
        '
        'bClear
        '
        Me.bClear.Image = Global.SMSPLUS.My.Resources.Resources.sclear
        Me.bClear.Name = "bClear"
        Me.bClear.Text = "Clear"
        '
        'bClosing
        '
        Me.bClosing.Name = "bClosing"
        Me.bClosing.SubItems.AddRange(New DevComponents.DotNetBar.BaseItem() {Me.bsClose, Me.bsCancled})
        Me.bClosing.Text = "Status Modify"
        '
        'bsClose
        '
        Me.bsClose.Name = "bsClose"
        Me.bsClose.Text = "CLOSED"
        '
        'bsCancled
        '
        Me.bsCancled.Name = "bsCancled"
        Me.bsCancled.Text = "CANCLED"
        '
        'ButtonItem1
        '
        Me.ButtonItem1.Name = "ButtonItem1"
        Me.ButtonItem1.Text = "Open &Rceiving"
        '
        'ButtonItem2
        '
        Me.ButtonItem2.ForeColor = System.Drawing.Color.Silver
        Me.ButtonItem2.Name = "ButtonItem2"
        Me.ButtonItem2.Text = "Modify Part's Order"
        '
        'ButtonItem5
        '
        Me.ButtonItem5.ForeColor = System.Drawing.Color.Silver
        Me.ButtonItem5.Name = "ButtonItem5"
        Me.ButtonItem5.Text = "Insert Part's Order"
        '
        'CtxSp2
        '
        Me.CtxSp2.AutoExpandOnClick = True
        Me.CtxSp2.Name = "CtxSp2"
        Me.CtxSp2.SubItems.AddRange(New DevComponents.DotNetBar.BaseItem() {Me.ButtonItem4, Me.cDel, Me.cClear})
        Me.CtxSp2.Text = "FpSpread3"
        '
        'ButtonItem4
        '
        Me.ButtonItem4.Name = "ButtonItem4"
        Me.ButtonItem4.Text = "Auto Order No"
        '
        'cDel
        '
        Me.cDel.Image = Global.SMSPLUS.My.Resources.Resources._14_layer_deletelayer
        Me.cDel.Name = "cDel"
        Me.cDel.Text = "Delete"
        '
        'cClear
        '
        Me.cClear.Image = Global.SMSPLUS.My.Resources.Resources.sclear
        Me.cClear.Name = "cClear"
        Me.cClear.Text = "Clear"
        '
        'CtxSp3
        '
        Me.CtxSp3.AutoExpandOnClick = True
        Me.CtxSp3.Name = "CtxSp3"
        Me.CtxSp3.SubItems.AddRange(New DevComponents.DotNetBar.BaseItem() {Me.dAllDown, Me.dSelDown, Me.dStatus, Me.ButtonItem3})
        Me.CtxSp3.Text = "FpSpread"
        '
        'dAllDown
        '
        Me.dAllDown.Name = "dAllDown"
        Me.dAllDown.Text = "ALL P/O Download"
        '
        'dSelDown
        '
        Me.dSelDown.Name = "dSelDown"
        Me.dSelDown.Text = "Selected P/O Download"
        '
        'dStatus
        '
        Me.dStatus.Name = "dStatus"
        Me.dStatus.SubItems.AddRange(New DevComponents.DotNetBar.BaseItem() {Me.dsClosed, Me.dsCancled})
        Me.dStatus.Text = "Status modify"
        '
        'dsClosed
        '
        Me.dsClosed.Name = "dsClosed"
        Me.dsClosed.Text = "CLOSED"
        '
        'dsCancled
        '
        Me.dsCancled.Name = "dsCancled"
        Me.dsCancled.Text = "CANCLED"
        '
        'ButtonItem3
        '
        Me.ButtonItem3.ForeColor = System.Drawing.Color.Silver
        Me.ButtonItem3.Name = "ButtonItem3"
        Me.ButtonItem3.Text = "Selected P/O Delete"
        '
        'PartList
        '
        Me.PartList.AllowColumnReorder = True
        Me.PartList.AllowDrop = True
        Me.PartList.BackColor = System.Drawing.Color.White
        '
        '
        '
        Me.PartList.Border.Class = "ListViewBorder"
        Me.PartList.Border.CornerType = DevComponents.DotNetBar.eCornerType.Square
        Me.PartList.Columns.AddRange(New System.Windows.Forms.ColumnHeader() {Me.ColumnHeader1, Me.ColumnHeader2, Me.ColumnHeader3, Me.ColumnHeader4, Me.ColumnHeader5, Me.ColumnHeader6})
        Me.PartList.Dock = System.Windows.Forms.DockStyle.Fill
        Me.PartList.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.PartList.ForeColor = System.Drawing.Color.Black
        Me.PartList.HoverSelection = True
        Me.PartList.Location = New System.Drawing.Point(0, 25)
        Me.PartList.Name = "PartList"
        Me.PartList.Size = New System.Drawing.Size(398, 615)
        Me.PartList.Sorting = System.Windows.Forms.SortOrder.Ascending
        Me.PartList.TabIndex = 4
        Me.PartList.UseCompatibleStateImageBehavior = False
        Me.PartList.View = System.Windows.Forms.View.Details
        '
        'ColumnHeader1
        '
        Me.ColumnHeader1.Text = "품목번호"
        Me.ColumnHeader1.Width = 113
        '
        'ColumnHeader2
        '
        Me.ColumnHeader2.Text = "품목명"
        Me.ColumnHeader2.Width = 85
        '
        'ColumnHeader3
        '
        Me.ColumnHeader3.Text = "모델"
        Me.ColumnHeader3.Width = 73
        '
        'ColumnHeader4
        '
        Me.ColumnHeader4.Text = "레벨"
        Me.ColumnHeader4.Width = 58
        '
        'ColumnHeader5
        '
        Me.ColumnHeader5.Text = "단가"
        '
        'ColumnHeader6
        '
        Me.ColumnHeader6.Text = "LG 품목번호"
        Me.ColumnHeader6.Width = 59
        '
        'Bar4
        '
        Me.Bar4.AccessibleDescription = "Bar4 (Bar4)"
        Me.Bar4.AccessibleName = "Bar4"
        Me.Bar4.AccessibleRole = System.Windows.Forms.AccessibleRole.ToolBar
        Me.Bar4.Controls.Add(Me.TextBoxX1)
        Me.Bar4.Controls.Add(Me.LabelX1)
        Me.Bar4.Controls.Add(Me.PtMdCb)
        Me.Bar4.Dock = System.Windows.Forms.DockStyle.Top
        Me.Bar4.Font = New System.Drawing.Font("Verdana", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Bar4.Location = New System.Drawing.Point(0, 0)
        Me.Bar4.Name = "Bar4"
        Me.Bar4.Size = New System.Drawing.Size(398, 25)
        Me.Bar4.Stretch = True
        Me.Bar4.Style = DevComponents.DotNetBar.eDotNetBarStyle.Office2007
        Me.Bar4.TabIndex = 0
        Me.Bar4.TabStop = False
        Me.Bar4.Text = "Bar4"
        '
        'TextBoxX1
        '
        Me.TextBoxX1.BackColor = System.Drawing.SystemColors.Control
        '
        '
        '
        Me.TextBoxX1.Border.Class = "TextBoxBorder"
        Me.TextBoxX1.Border.CornerType = DevComponents.DotNetBar.eCornerType.Square
        Me.TextBoxX1.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TextBoxX1.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TextBoxX1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.TextBoxX1.Location = New System.Drawing.Point(239, 0)
        Me.TextBoxX1.Name = "TextBoxX1"
        Me.TextBoxX1.Size = New System.Drawing.Size(150, 21)
        Me.TextBoxX1.TabIndex = 2
        '
        'LabelX1
        '
        Me.LabelX1.BackColor = System.Drawing.Color.Transparent
        '
        '
        '
        Me.LabelX1.BackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square
        Me.LabelX1.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LabelX1.Location = New System.Drawing.Point(4, 3)
        Me.LabelX1.Name = "LabelX1"
        Me.LabelX1.Size = New System.Drawing.Size(30, 19)
        Me.LabelX1.TabIndex = 1
        Me.LabelX1.Text = "SKU"
        Me.LabelX1.TextAlignment = System.Drawing.StringAlignment.Center
        '
        'PtMdCb
        '
        Me.PtMdCb.DisplayMember = "Text"
        Me.PtMdCb.DrawMode = System.Windows.Forms.DrawMode.OwnerDrawFixed
        Me.PtMdCb.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.PtMdCb.FormattingEnabled = True
        Me.PtMdCb.ItemHeight = 15
        Me.PtMdCb.Location = New System.Drawing.Point(35, 0)
        Me.PtMdCb.Name = "PtMdCb"
        Me.PtMdCb.Size = New System.Drawing.Size(198, 21)
        Me.PtMdCb.TabIndex = 0
        '
        'DockContainerItem2
        '
        Me.DockContainerItem2.Control = Me.PanelDockContainer2
        Me.DockContainerItem2.Name = "DockContainerItem2"
        Me.DockContainerItem2.Text = "DockContainerItem2"
        '
        'DockSite2
        '
        Me.DockSite2.AccessibleRole = System.Windows.Forms.AccessibleRole.Window
        Me.DockSite2.Dock = System.Windows.Forms.DockStyle.Right
        Me.DockSite2.DocumentDockContainer = New DevComponents.DotNetBar.DocumentDockContainer()
        Me.DockSite2.Location = New System.Drawing.Point(1270, 66)
        Me.DockSite2.Name = "DockSite2"
        Me.DockSite2.Size = New System.Drawing.Size(0, 666)
        Me.DockSite2.TabIndex = 1
        Me.DockSite2.TabStop = False
        '
        'DockSite8
        '
        Me.DockSite8.AccessibleRole = System.Windows.Forms.AccessibleRole.Window
        Me.DockSite8.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.DockSite8.Location = New System.Drawing.Point(0, 732)
        Me.DockSite8.Name = "DockSite8"
        Me.DockSite8.Size = New System.Drawing.Size(1270, 0)
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
        Me.DockSite6.Location = New System.Drawing.Point(1270, 0)
        Me.DockSite6.Name = "DockSite6"
        Me.DockSite6.Size = New System.Drawing.Size(0, 732)
        Me.DockSite6.TabIndex = 5
        Me.DockSite6.TabStop = False
        '
        'DockSite7
        '
        Me.DockSite7.AccessibleRole = System.Windows.Forms.AccessibleRole.Window
        Me.DockSite7.Dock = System.Windows.Forms.DockStyle.Top
        Me.DockSite7.Location = New System.Drawing.Point(0, 0)
        Me.DockSite7.Name = "DockSite7"
        Me.DockSite7.Size = New System.Drawing.Size(1270, 0)
        Me.DockSite7.TabIndex = 6
        Me.DockSite7.TabStop = False
        '
        'DockSite3
        '
        Me.DockSite3.AccessibleRole = System.Windows.Forms.AccessibleRole.Window
        Me.DockSite3.Controls.Add(Me.Bar1)
        Me.DockSite3.Controls.Add(Me.Bar8)
        Me.DockSite3.Dock = System.Windows.Forms.DockStyle.Top
        Me.DockSite3.DocumentDockContainer = New DevComponents.DotNetBar.DocumentDockContainer(New DevComponents.DotNetBar.DocumentBaseContainer() {CType(New DevComponents.DotNetBar.DocumentBarContainer(Me.Bar1, 260, 63), DevComponents.DotNetBar.DocumentBaseContainer), CType(New DevComponents.DotNetBar.DocumentBarContainer(Me.Bar8, 1007, 63), DevComponents.DotNetBar.DocumentBaseContainer)}, DevComponents.DotNetBar.eOrientation.Horizontal)
        Me.DockSite3.Location = New System.Drawing.Point(0, 0)
        Me.DockSite3.Name = "DockSite3"
        Me.DockSite3.Size = New System.Drawing.Size(1270, 66)
        Me.DockSite3.TabIndex = 2
        Me.DockSite3.TabStop = False
        Me.DockSite3.Text = "실행 메뉴"
        '
        'Bar1
        '
        Me.Bar1.AccessibleDescription = "DotNetBar Bar (Bar1)"
        Me.Bar1.AccessibleName = "DotNetBar Bar"
        Me.Bar1.AccessibleRole = System.Windows.Forms.AccessibleRole.ToolBar
        Me.Bar1.AutoSyncBarCaption = True
        Me.Bar1.CloseSingleTab = True
        Me.Bar1.Controls.Add(Me.PanelDockContainer1)
        Me.Bar1.Font = New System.Drawing.Font("Verdana", 8.0!)
        Me.Bar1.GrabHandleStyle = DevComponents.DotNetBar.eGrabHandleStyle.Caption
        Me.Bar1.Items.AddRange(New DevComponents.DotNetBar.BaseItem() {Me.DockContainerItem1})
        Me.Bar1.LayoutType = DevComponents.DotNetBar.eLayoutType.DockContainer
        Me.Bar1.Location = New System.Drawing.Point(0, 0)
        Me.Bar1.Name = "Bar1"
        Me.Bar1.Size = New System.Drawing.Size(260, 63)
        Me.Bar1.Stretch = True
        Me.Bar1.Style = DevComponents.DotNetBar.eDotNetBarStyle.Office2007
        Me.Bar1.TabIndex = 0
        Me.Bar1.TabStop = False
        Me.Bar1.Text = "DockContainerItem1"
        '
        'PanelDockContainer1
        '
        Me.PanelDockContainer1.ColorSchemeStyle = DevComponents.DotNetBar.eDotNetBarStyle.Office2007
        Me.PanelDockContainer1.Controls.Add(Me.CtlBar)
        Me.PanelDockContainer1.Location = New System.Drawing.Point(3, 23)
        Me.PanelDockContainer1.Name = "PanelDockContainer1"
        Me.PanelDockContainer1.Size = New System.Drawing.Size(254, 37)
        Me.PanelDockContainer1.Style.Alignment = System.Drawing.StringAlignment.Center
        Me.PanelDockContainer1.Style.BackColor1.ColorSchemePart = DevComponents.DotNetBar.eColorSchemePart.BarBackground
        Me.PanelDockContainer1.Style.ForeColor.ColorSchemePart = DevComponents.DotNetBar.eColorSchemePart.ItemText
        Me.PanelDockContainer1.Style.GradientAngle = 90
        Me.PanelDockContainer1.TabIndex = 0
        Me.PanelDockContainer1.Text = "실행 메뉴"
        '
        'CtlBar
        '
        Me.CtlBar.Dock = System.Windows.Forms.DockStyle.Fill
        Me.CtlBar.Font = New System.Drawing.Font("Malgun Gothic", 9.0!)
        Me.CtlBar.Items.AddRange(New DevComponents.DotNetBar.BaseItem() {Me.FindBtn, Me.NewBtn, Me.SaveBtn, Me.DelBtn, Me.PrtBtn, Me.Excel})
        Me.CtlBar.Location = New System.Drawing.Point(0, 0)
        Me.CtlBar.Name = "CtlBar"
        Me.CtlBar.Size = New System.Drawing.Size(254, 41)
        Me.CtlBar.Stretch = True
        Me.CtlBar.Style = DevComponents.DotNetBar.eDotNetBarStyle.Office2007
        Me.CtlBar.TabIndex = 2
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
        'NewBtn
        '
        Me.NewBtn.Image = Global.SMSPLUS.My.Resources.Resources.document_new
        Me.NewBtn.Name = "NewBtn"
        Me.NewBtn.Shortcuts.Add(DevComponents.DotNetBar.eShortcut.CtrlN)
        Me.NewBtn.Text = "New"
        Me.NewBtn.Tooltip = "Add new row"
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
        Me.DockContainerItem1.Text = "DockContainerItem1"
        '
        'Bar8
        '
        Me.Bar8.AccessibleDescription = "DotNetBar Bar (Bar8)"
        Me.Bar8.AccessibleName = "DotNetBar Bar"
        Me.Bar8.AccessibleRole = System.Windows.Forms.AccessibleRole.ToolBar
        Me.Bar8.AutoSyncBarCaption = True
        Me.Bar8.Controls.Add(Me.PanelDockContainer3)
        Me.Bar8.Font = New System.Drawing.Font("Verdana", 8.0!)
        Me.Bar8.GrabHandleStyle = DevComponents.DotNetBar.eGrabHandleStyle.Caption
        Me.Bar8.Items.AddRange(New DevComponents.DotNetBar.BaseItem() {Me.DockContainerItem6})
        Me.Bar8.LayoutType = DevComponents.DotNetBar.eLayoutType.DockContainer
        Me.Bar8.Location = New System.Drawing.Point(263, 0)
        Me.Bar8.Name = "Bar8"
        Me.Bar8.Size = New System.Drawing.Size(1007, 63)
        Me.Bar8.Stretch = True
        Me.Bar8.Style = DevComponents.DotNetBar.eDotNetBarStyle.Office2007
        Me.Bar8.TabIndex = 1
        Me.Bar8.TabStop = False
        Me.Bar8.Text = "DockContainerItem6"
        '
        'PanelDockContainer3
        '
        Me.PanelDockContainer3.ColorSchemeStyle = DevComponents.DotNetBar.eDotNetBarStyle.Office2007
        Me.PanelDockContainer3.Controls.Add(Me.ComboBoxEx1)
        Me.PanelDockContainer3.Controls.Add(Me.LabelX7)
        Me.PanelDockContainer3.Controls.Add(Me.LabelX5)
        Me.PanelDockContainer3.Controls.Add(Me.LabelX2)
        Me.PanelDockContainer3.Controls.Add(Me.ModelCb)
        Me.PanelDockContainer3.Controls.Add(Me.POStDate)
        Me.PanelDockContainer3.Controls.Add(Me.POEdDate)
        Me.PanelDockContainer3.Controls.Add(Me.SupCb)
        Me.PanelDockContainer3.Controls.Add(Me.StatCb)
        Me.PanelDockContainer3.Controls.Add(Me.LabelX3)
        Me.PanelDockContainer3.Location = New System.Drawing.Point(3, 23)
        Me.PanelDockContainer3.Name = "PanelDockContainer3"
        Me.PanelDockContainer3.Size = New System.Drawing.Size(1001, 37)
        Me.PanelDockContainer3.Style.Alignment = System.Drawing.StringAlignment.Center
        Me.PanelDockContainer3.Style.BackColor1.ColorSchemePart = DevComponents.DotNetBar.eColorSchemePart.BarBackground
        Me.PanelDockContainer3.Style.ForeColor.ColorSchemePart = DevComponents.DotNetBar.eColorSchemePart.ItemText
        Me.PanelDockContainer3.Style.GradientAngle = 90
        Me.PanelDockContainer3.TabIndex = 2
        '
        'ComboBoxEx1
        '
        Me.ComboBoxEx1.DisplayMember = "Text"
        Me.ComboBoxEx1.DrawMode = System.Windows.Forms.DrawMode.OwnerDrawFixed
        Me.ComboBoxEx1.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ComboBoxEx1.FormattingEnabled = True
        Me.ComboBoxEx1.ItemHeight = 15
        Me.ComboBoxEx1.Location = New System.Drawing.Point(640, 8)
        Me.ComboBoxEx1.Name = "ComboBoxEx1"
        Me.ComboBoxEx1.Size = New System.Drawing.Size(173, 21)
        Me.ComboBoxEx1.TabIndex = 17
        '
        'LabelX7
        '
        Me.LabelX7.BackColor = System.Drawing.Color.Transparent
        '
        '
        '
        Me.LabelX7.BackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square
        Me.LabelX7.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LabelX7.Location = New System.Drawing.Point(184, 11)
        Me.LabelX7.Name = "LabelX7"
        Me.LabelX7.Size = New System.Drawing.Size(15, 15)
        Me.LabelX7.TabIndex = 16
        Me.LabelX7.Text = "~"
        Me.LabelX7.TextAlignment = System.Drawing.StringAlignment.Center
        '
        'LabelX5
        '
        Me.LabelX5.AutoSize = True
        Me.LabelX5.BackColor = System.Drawing.Color.Transparent
        '
        '
        '
        Me.LabelX5.BackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square
        Me.LabelX5.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LabelX5.Location = New System.Drawing.Point(320, 10)
        Me.LabelX5.Name = "LabelX5"
        Me.LabelX5.Size = New System.Drawing.Size(44, 17)
        Me.LabelX5.TabIndex = 14
        Me.LabelX5.Text = "수주처"
        Me.LabelX5.TextAlignment = System.Drawing.StringAlignment.Center
        Me.LabelX5.TextLineAlignment = System.Drawing.StringAlignment.Far
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
        Me.LabelX2.Location = New System.Drawing.Point(12, 10)
        Me.LabelX2.Name = "LabelX2"
        Me.LabelX2.Size = New System.Drawing.Size(56, 17)
        Me.LabelX2.TabIndex = 6
        Me.LabelX2.Text = "조회기간"
        Me.LabelX2.TextAlignment = System.Drawing.StringAlignment.Center
        '
        'ModelCb
        '
        Me.ModelCb.DisplayMember = "Text"
        Me.ModelCb.DrawMode = System.Windows.Forms.DrawMode.OwnerDrawFixed
        Me.ModelCb.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ModelCb.FormattingEnabled = True
        Me.ModelCb.ItemHeight = 15
        Me.ModelCb.Location = New System.Drawing.Point(366, 8)
        Me.ModelCb.Name = "ModelCb"
        Me.ModelCb.Size = New System.Drawing.Size(121, 21)
        Me.ModelCb.TabIndex = 13
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
        Me.POStDate.Location = New System.Drawing.Point(71, 8)
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
        Me.POStDate.TabIndex = 7
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
        Me.POEdDate.Location = New System.Drawing.Point(200, 8)
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
        Me.POEdDate.TabIndex = 8
        '
        'SupCb
        '
        Me.SupCb.DisplayMember = "Text"
        Me.SupCb.DrawMode = System.Windows.Forms.DrawMode.OwnerDrawFixed
        Me.SupCb.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.SupCb.FormattingEnabled = True
        Me.SupCb.ItemHeight = 15
        Me.SupCb.Location = New System.Drawing.Point(490, 8)
        Me.SupCb.Name = "SupCb"
        Me.SupCb.Size = New System.Drawing.Size(147, 21)
        Me.SupCb.TabIndex = 11
        '
        'StatCb
        '
        Me.StatCb.DisplayMember = "Text"
        Me.StatCb.DrawMode = System.Windows.Forms.DrawMode.OwnerDrawFixed
        Me.StatCb.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.StatCb.FormattingEnabled = True
        Me.StatCb.ItemHeight = 15
        Me.StatCb.Location = New System.Drawing.Point(879, 8)
        Me.StatCb.Name = "StatCb"
        Me.StatCb.Size = New System.Drawing.Size(121, 21)
        Me.StatCb.TabIndex = 9
        '
        'LabelX3
        '
        Me.LabelX3.AutoSize = True
        Me.LabelX3.BackColor = System.Drawing.Color.Transparent
        '
        '
        '
        Me.LabelX3.BackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square
        Me.LabelX3.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LabelX3.Location = New System.Drawing.Point(822, 10)
        Me.LabelX3.Name = "LabelX3"
        Me.LabelX3.Size = New System.Drawing.Size(56, 17)
        Me.LabelX3.TabIndex = 10
        Me.LabelX3.Text = "진행상태"
        Me.LabelX3.TextAlignment = System.Drawing.StringAlignment.Center
        Me.LabelX3.TextLineAlignment = System.Drawing.StringAlignment.Far
        '
        'DockContainerItem6
        '
        Me.DockContainerItem6.Control = Me.PanelDockContainer3
        Me.DockContainerItem6.Name = "DockContainerItem6"
        Me.DockContainerItem6.Text = "DockContainerItem6"
        '
        'DockContainerItem10
        '
        Me.DockContainerItem10.Name = "DockContainerItem10"
        Me.DockContainerItem10.Text = "입고 내역"
        '
        'DockContainerItem3
        '
        Me.DockContainerItem3.Name = "DockContainerItem3"
        Me.DockContainerItem3.Text = "DockContainerItem3"
        '
        'bNew
        '
        Me.bNew.Image = Global.SMSPLUS.My.Resources.Resources.filenew
        Me.bNew.Name = "bNew"
        Me.bNew.Text = "&New"
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
        'FpSpread4
        '
        Me.FpSpread4.AccessibleDescription = ""
        Me.FpSpread4.Location = New System.Drawing.Point(675, 0)
        Me.FpSpread4.Name = "FpSpread4"
        Me.FpSpread4.Sheets.AddRange(New FarPoint.Win.Spread.SheetView() {Me.FpSpread4_Sheet1})
        Me.FpSpread4.Size = New System.Drawing.Size(200, 100)
        Me.FpSpread4.TabIndex = 24
        Me.FpSpread4.Visible = False
        '
        'FpSpread4_Sheet1
        '
        Me.FpSpread4_Sheet1.Reset()
        Me.FpSpread4_Sheet1.SheetName = "Sheet1"
        'Formulas and custom names must be loaded with R1C1 reference style
        Me.FpSpread4_Sheet1.ReferenceStyle = FarPoint.Win.Spread.Model.ReferenceStyle.R1C1
        Me.FpSpread4_Sheet1.ColumnHeader.RowCount = 0
        Me.FpSpread4_Sheet1.ColumnHeader.Visible = False
        Me.FpSpread4_Sheet1.ReferenceStyle = FarPoint.Win.Spread.Model.ReferenceStyle.A1
        '
        'LabelItem2
        '
        Me.LabelItem2.Name = "LabelItem2"
        Me.LabelItem2.Text = "P/O HEADER"
        '
        'LabelItem3
        '
        Me.LabelItem3.Name = "LabelItem3"
        Me.LabelItem3.Text = "P/O DETAILS"
        '
        'LabelItem4
        '
        Me.LabelItem4.Name = "LabelItem4"
        Me.LabelItem4.Text = "RECEIVING DETAILS"
        '
        'StyleManager1
        '
        Me.StyleManager1.ManagerStyle = DevComponents.DotNetBar.eStyle.Windows7Blue
        Me.StyleManager1.MetroColorParameters = New DevComponents.DotNetBar.Metro.ColorTables.MetroColorGeneratorParameters(System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer)), System.Drawing.Color.FromArgb(CType(CType(128, Byte), Integer), CType(CType(57, Byte), Integer), CType(CType(123, Byte), Integer)))
        '
        'FrmSalesPO
        '
        Me.ClientSize = New System.Drawing.Size(1270, 732)
        Me.Controls.Add(Me.DockSite9)
        Me.Controls.Add(Me.DockSite2)
        Me.Controls.Add(Me.DockSite1)
        Me.Controls.Add(Me.DockSite3)
        Me.Controls.Add(Me.DockSite4)
        Me.Controls.Add(Me.DockSite5)
        Me.Controls.Add(Me.DockSite6)
        Me.Controls.Add(Me.DockSite7)
        Me.Controls.Add(Me.DockSite8)
        Me.Controls.Add(Me.XlsBtn1)
        Me.Controls.Add(Me.PrtBtn1)
        Me.Controls.Add(Me.DelBtn1)
        Me.Controls.Add(Me.SaveBtn1)
        Me.Controls.Add(Me.NewBtn1)
        Me.Controls.Add(Me.FpSpread4)
        Me.Font = New System.Drawing.Font("Verdana", 8.0!)
        Me.Name = "FrmSalesPO"
        Me.Text = "수주관리"
        Me.DockSite9.ResumeLayout(False)
        CType(Me.Bar12, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Bar12.ResumeLayout(False)
        Me.PanelDockContainer7.ResumeLayout(False)
        CType(Me.FpSpread1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.FpSpread1_Sheet1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.Bar3, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Bar3.ResumeLayout(False)
        Me.PanelDockContainer8.ResumeLayout(False)
        CType(Me.FpSpread2, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.FpSpread2_Sheet1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.DockSite1.ResumeLayout(False)
        CType(Me.Bar2, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Bar2.ResumeLayout(False)
        Me.PanelDockContainer2.ResumeLayout(False)
        CType(Me.ContextMenuBar1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.Bar4, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Bar4.ResumeLayout(False)
        Me.DockSite3.ResumeLayout(False)
        CType(Me.Bar1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Bar1.ResumeLayout(False)
        Me.PanelDockContainer1.ResumeLayout(False)
        CType(Me.CtlBar, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.Bar8, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Bar8.ResumeLayout(False)
        Me.PanelDockContainer3.ResumeLayout(False)
        Me.PanelDockContainer3.PerformLayout()
        CType(Me.POStDate, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.POEdDate, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.FpSpread4, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.FpSpread4_Sheet1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents FrmGuiMgr As DevComponents.DotNetBar.DotNetBarManager
    Friend WithEvents DockSite4 As DevComponents.DotNetBar.DockSite
    Friend WithEvents DockSite1 As DevComponents.DotNetBar.DockSite
    Friend WithEvents DockSite2 As DevComponents.DotNetBar.DockSite
    Friend WithEvents DockSite3 As DevComponents.DotNetBar.DockSite
    Friend WithEvents DockSite5 As DevComponents.DotNetBar.DockSite
    Friend WithEvents DockSite6 As DevComponents.DotNetBar.DockSite
    Friend WithEvents DockSite7 As DevComponents.DotNetBar.DockSite
    Friend WithEvents DockSite8 As DevComponents.DotNetBar.DockSite
    Friend WithEvents DockSite9 As DevComponents.DotNetBar.DockSite
    Friend WithEvents DockContainerItem3 As DevComponents.DotNetBar.DockContainerItem
    Friend WithEvents Bar2 As DevComponents.DotNetBar.Bar
    Friend WithEvents PanelDockContainer2 As DevComponents.DotNetBar.PanelDockContainer
    Friend WithEvents DockContainerItem2 As DevComponents.DotNetBar.DockContainerItem
    Friend WithEvents Bar1 As DevComponents.DotNetBar.Bar
    Friend WithEvents PanelDockContainer1 As DevComponents.DotNetBar.PanelDockContainer
    Friend WithEvents DockContainerItem1 As DevComponents.DotNetBar.DockContainerItem
    Friend WithEvents CtlBar As DevComponents.DotNetBar.Bar
    Friend WithEvents NewBtn As DevComponents.DotNetBar.ButtonItem
    Friend WithEvents SaveBtn As DevComponents.DotNetBar.ButtonItem
    Friend WithEvents DelBtn As DevComponents.DotNetBar.ButtonItem
    Friend WithEvents PrtBtn As DevComponents.DotNetBar.ButtonItem
    Friend WithEvents Excel As DevComponents.DotNetBar.ButtonItem
    Friend WithEvents Bar4 As DevComponents.DotNetBar.Bar
    Friend WithEvents PartList As DevComponents.DotNetBar.Controls.ListViewEx
    Friend WithEvents POEdDate As DevComponents.Editors.DateTimeAdv.DateTimeInput
    Friend WithEvents POStDate As DevComponents.Editors.DateTimeAdv.DateTimeInput
    Friend WithEvents LabelX2 As DevComponents.DotNetBar.LabelX
    Friend WithEvents LabelX1 As DevComponents.DotNetBar.LabelX
    Friend WithEvents PtMdCb As DevComponents.DotNetBar.Controls.ComboBoxEx
    Friend WithEvents LabelX3 As DevComponents.DotNetBar.LabelX
    Friend WithEvents StatCb As DevComponents.DotNetBar.Controls.ComboBoxEx
    Friend WithEvents LabelX5 As DevComponents.DotNetBar.LabelX
    Friend WithEvents ModelCb As DevComponents.DotNetBar.Controls.ComboBoxEx
    Friend WithEvents SupCb As DevComponents.DotNetBar.Controls.ComboBoxEx
    Friend WithEvents LabelX7 As DevComponents.DotNetBar.LabelX
    Friend WithEvents ColumnHeader1 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader2 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader3 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader4 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader5 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader6 As System.Windows.Forms.ColumnHeader
    Friend WithEvents Bar8 As DevComponents.DotNetBar.Bar
    Friend WithEvents PanelDockContainer3 As DevComponents.DotNetBar.PanelDockContainer
    Friend WithEvents DockContainerItem6 As DevComponents.DotNetBar.DockContainerItem
    Friend WithEvents SaveFileDialog1 As System.Windows.Forms.SaveFileDialog
    Friend WithEvents ContextMenuBar1 As DevComponents.DotNetBar.ContextMenuBar
    Friend WithEvents CtxSp As DevComponents.DotNetBar.ButtonItem
    Friend WithEvents bNew As DevComponents.DotNetBar.ButtonItem
    Friend WithEvents bSave As DevComponents.DotNetBar.ButtonItem
    Friend WithEvents bDel As DevComponents.DotNetBar.ButtonItem
    Friend WithEvents bPrint As DevComponents.DotNetBar.ButtonItem
    Friend WithEvents bExcel As DevComponents.DotNetBar.ButtonItem
    Friend WithEvents bClear As DevComponents.DotNetBar.ButtonItem
    Friend WithEvents CtxSp2 As DevComponents.DotNetBar.ButtonItem
    Friend WithEvents cDel As DevComponents.DotNetBar.ButtonItem
    Friend WithEvents cClear As DevComponents.DotNetBar.ButtonItem
    Friend WithEvents CtxSp3 As DevComponents.DotNetBar.ButtonItem
    Friend WithEvents dAllDown As DevComponents.DotNetBar.ButtonItem
    Friend WithEvents dSelDown As DevComponents.DotNetBar.ButtonItem
    Friend WithEvents XlsBtn1 As DevComponents.DotNetBar.ButtonX
    Friend WithEvents PrtBtn1 As DevComponents.DotNetBar.ButtonX
    Friend WithEvents DelBtn1 As DevComponents.DotNetBar.ButtonX
    Friend WithEvents SaveBtn1 As DevComponents.DotNetBar.ButtonX
    Friend WithEvents NewBtn1 As DevComponents.DotNetBar.ButtonX
    Friend WithEvents bClosing As DevComponents.DotNetBar.ButtonItem
    Friend WithEvents dStatus As DevComponents.DotNetBar.ButtonItem
    Friend WithEvents bsClose As DevComponents.DotNetBar.ButtonItem
    Friend WithEvents bsCancled As DevComponents.DotNetBar.ButtonItem
    Friend WithEvents dsClosed As DevComponents.DotNetBar.ButtonItem
    Friend WithEvents dsCancled As DevComponents.DotNetBar.ButtonItem
    Friend WithEvents FpSpread4 As FarPoint.Win.Spread.FpSpread
    Friend WithEvents FpSpread4_Sheet1 As FarPoint.Win.Spread.SheetView
    Friend WithEvents LabelItem2 As DevComponents.DotNetBar.LabelItem
    Friend WithEvents LabelItem3 As DevComponents.DotNetBar.LabelItem
    Friend WithEvents Bar12 As DevComponents.DotNetBar.Bar
    Friend WithEvents PanelDockContainer7 As DevComponents.DotNetBar.PanelDockContainer
    Friend WithEvents FpSpread1 As FarPoint.Win.Spread.FpSpread
    Friend WithEvents FpSpread1_Sheet1 As FarPoint.Win.Spread.SheetView
    Friend WithEvents DockContainerItem4 As DevComponents.DotNetBar.DockContainerItem
    Friend WithEvents Bar3 As DevComponents.DotNetBar.Bar
    Friend WithEvents PanelDockContainer8 As DevComponents.DotNetBar.PanelDockContainer
    Friend WithEvents FpSpread2 As FarPoint.Win.Spread.FpSpread
    Friend WithEvents FpSpread2_Sheet1 As FarPoint.Win.Spread.SheetView
    Friend WithEvents DockContainerItem5 As DevComponents.DotNetBar.DockContainerItem
    Friend WithEvents DockContainerItem10 As DevComponents.DotNetBar.DockContainerItem
    Friend WithEvents LabelItem4 As DevComponents.DotNetBar.LabelItem
    Friend WithEvents ButtonItem1 As DevComponents.DotNetBar.ButtonItem
    Friend WithEvents ButtonItem2 As DevComponents.DotNetBar.ButtonItem
    Friend WithEvents ButtonItem3 As DevComponents.DotNetBar.ButtonItem
    Friend WithEvents ButtonItem4 As DevComponents.DotNetBar.ButtonItem
    Friend WithEvents TextBoxX1 As DevComponents.DotNetBar.Controls.TextBoxX
    Friend WithEvents ButtonItem5 As DevComponents.DotNetBar.ButtonItem
    Friend WithEvents FindBtn As DevComponents.DotNetBar.ButtonItem
    Friend WithEvents ComboBoxEx1 As DevComponents.DotNetBar.Controls.ComboBoxEx
    Friend WithEvents StyleManager1 As DevComponents.DotNetBar.StyleManager

End Class
