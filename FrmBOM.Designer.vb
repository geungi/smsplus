<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmBOM
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
        Dim EnhancedRowHeaderRenderer1 As FarPoint.Win.Spread.CellType.EnhancedRowHeaderRenderer = New FarPoint.Win.Spread.CellType.EnhancedRowHeaderRenderer
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(FrmBOM))
        Me.FrmGUImgr = New DevComponents.DotNetBar.DotNetBarManager(Me.components)
        Me.DockSite4 = New DevComponents.DotNetBar.DockSite
        Me.DockSite9 = New DevComponents.DotNetBar.DockSite
        Me.Bar2 = New DevComponents.DotNetBar.Bar
        Me.PartList = New DevComponents.DotNetBar.Controls.ListViewEx
        Me.PartColHd1 = New System.Windows.Forms.ColumnHeader
        Me.PartColHd2 = New System.Windows.Forms.ColumnHeader
        Me.PartColHd3 = New System.Windows.Forms.ColumnHeader
        Me.PartColHd4 = New System.Windows.Forms.ColumnHeader
        Me.PartColHd5 = New System.Windows.Forms.ColumnHeader
        Me.Bar3 = New DevComponents.DotNetBar.Bar
        Me.LabelX1 = New DevComponents.DotNetBar.LabelX
        Me.TextBoxX1 = New DevComponents.DotNetBar.Controls.TextBoxX
        Me.Bar1 = New DevComponents.DotNetBar.Bar
        Me.FpSpread1 = New FarPoint.Win.Spread.FpSpread
        Me.FpSpread1_Sheet1 = New FarPoint.Win.Spread.SheetView
        Me.bar4 = New DevComponents.DotNetBar.Bar
        Me.LabelX2 = New DevComponents.DotNetBar.LabelX
        Me.DockSite1 = New DevComponents.DotNetBar.DockSite
        Me.DockSite2 = New DevComponents.DotNetBar.DockSite
        Me.ModelArea = New DevComponents.DotNetBar.Bar
        Me.DockSite8 = New DevComponents.DotNetBar.DockSite
        Me.DockSite5 = New DevComponents.DotNetBar.DockSite
        Me.DockSite6 = New DevComponents.DotNetBar.DockSite
        Me.DockSite7 = New DevComponents.DotNetBar.DockSite
        Me.DockSite3 = New DevComponents.DotNetBar.DockSite
        Me.Bar5 = New DevComponents.DotNetBar.Bar
        Me.PanelDockContainer2 = New DevComponents.DotNetBar.PanelDockContainer
        Me.ModelLbl = New DevComponents.DotNetBar.LabelX
        Me.ActCb = New DevComponents.DotNetBar.Controls.ComboBoxEx
        Me.ALL = New DevComponents.Editors.ComboItem
        Me.YES = New DevComponents.Editors.ComboItem
        Me.NO = New DevComponents.Editors.ComboItem
        Me.ActLbl = New DevComponents.DotNetBar.LabelX
        Me.ModelCb = New DevComponents.DotNetBar.Controls.ComboBoxEx
        Me.CheckBoxX1 = New DevComponents.DotNetBar.Controls.CheckBoxX
        Me.DockContainerItem10 = New DevComponents.DotNetBar.DockContainerItem
        Me.TopDock = New DevComponents.DotNetBar.Bar
        Me.MenuPanel = New DevComponents.DotNetBar.PanelDockContainer
        Me.CtlBar = New DevComponents.DotNetBar.Bar
        Me.FindBtn = New DevComponents.DotNetBar.ButtonItem
        Me.SaveBtn = New DevComponents.DotNetBar.ButtonItem
        Me.DelBtn = New DevComponents.DotNetBar.ButtonItem
        Me.PrtBtn = New DevComponents.DotNetBar.ButtonItem
        Me.Excel = New DevComponents.DotNetBar.ButtonItem
        Me.DockContainerItem1 = New DevComponents.DotNetBar.DockContainerItem
        Me.PanelDockContainer1 = New DevComponents.DotNetBar.PanelDockContainer
        Me.DockContainerItem5 = New DevComponents.DotNetBar.DockContainerItem
        Me.CtxSp = New DevComponents.DotNetBar.ButtonItem
        Me.bDel = New DevComponents.DotNetBar.ButtonItem
        Me.bSave = New DevComponents.DotNetBar.ButtonItem
        Me.bPrint = New DevComponents.DotNetBar.ButtonItem
        Me.bExcel = New DevComponents.DotNetBar.ButtonItem
        Me.DockContainerItem2 = New DevComponents.DotNetBar.DockContainerItem
        Me.SaveFileDialog1 = New System.Windows.Forms.SaveFileDialog
        Me.ContextMenuBar1 = New DevComponents.DotNetBar.ContextMenuBar
        Me.DockContainerItem4 = New DevComponents.DotNetBar.DockContainerItem
        Me.DockContainerItem3 = New DevComponents.DotNetBar.DockContainerItem
        Me.DockContainerItem6 = New DevComponents.DotNetBar.DockContainerItem
        Me.DockContainerItem7 = New DevComponents.DotNetBar.DockContainerItem
        Me.DockContainerItem9 = New DevComponents.DotNetBar.DockContainerItem
        Me.DockContainerItem8 = New DevComponents.DotNetBar.DockContainerItem
        Me.ControlContainerItem2 = New DevComponents.DotNetBar.ControlContainerItem
        Me.XlsBtn1 = New DevComponents.DotNetBar.ButtonX
        Me.PrtBtn1 = New DevComponents.DotNetBar.ButtonX
        Me.DelBtn1 = New DevComponents.DotNetBar.ButtonX
        Me.SaveBtn1 = New DevComponents.DotNetBar.ButtonX
        Me.NewBtn1 = New DevComponents.DotNetBar.ButtonX
        Me.OpenFileDialog1 = New System.Windows.Forms.OpenFileDialog
        Me.DockSite9.SuspendLayout()
        CType(Me.Bar2, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Bar2.SuspendLayout()
        CType(Me.Bar3, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Bar3.SuspendLayout()
        CType(Me.Bar1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Bar1.SuspendLayout()
        CType(Me.FpSpread1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.FpSpread1_Sheet1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.bar4, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.bar4.SuspendLayout()
        Me.DockSite2.SuspendLayout()
        CType(Me.ModelArea, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.DockSite3.SuspendLayout()
        CType(Me.Bar5, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Bar5.SuspendLayout()
        Me.PanelDockContainer2.SuspendLayout()
        CType(Me.TopDock, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TopDock.SuspendLayout()
        Me.MenuPanel.SuspendLayout()
        CType(Me.CtlBar, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.ContextMenuBar1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        EnhancedRowHeaderRenderer1.Name = "EnhancedRowHeaderRenderer1"
        EnhancedRowHeaderRenderer1.TextRotationAngle = 0
        '
        'FrmGUImgr
        '
        Me.FrmGUImgr.AutoDispatchShortcuts.Add(DevComponents.DotNetBar.eShortcut.F1)
        Me.FrmGUImgr.AutoDispatchShortcuts.Add(DevComponents.DotNetBar.eShortcut.CtrlC)
        Me.FrmGUImgr.AutoDispatchShortcuts.Add(DevComponents.DotNetBar.eShortcut.CtrlA)
        Me.FrmGUImgr.AutoDispatchShortcuts.Add(DevComponents.DotNetBar.eShortcut.CtrlV)
        Me.FrmGUImgr.AutoDispatchShortcuts.Add(DevComponents.DotNetBar.eShortcut.CtrlX)
        Me.FrmGUImgr.AutoDispatchShortcuts.Add(DevComponents.DotNetBar.eShortcut.CtrlZ)
        Me.FrmGUImgr.AutoDispatchShortcuts.Add(DevComponents.DotNetBar.eShortcut.CtrlY)
        Me.FrmGUImgr.AutoDispatchShortcuts.Add(DevComponents.DotNetBar.eShortcut.CtrlDel)
        Me.FrmGUImgr.AutoDispatchShortcuts.Add(DevComponents.DotNetBar.eShortcut.Ins)
        Me.FrmGUImgr.BottomDockSite = Me.DockSite4
        Me.FrmGUImgr.EnableFullSizeDock = False
        Me.FrmGUImgr.FillDockSite = Me.DockSite9
        Me.FrmGUImgr.LeftDockSite = Me.DockSite1
        Me.FrmGUImgr.LicenseKey = "F962CEC7-CD8F-4911-A9E9-CAB39962FC1F"
        Me.FrmGUImgr.ParentForm = Me
        Me.FrmGUImgr.RightDockSite = Me.DockSite2
        Me.FrmGUImgr.Style = DevComponents.DotNetBar.eDotNetBarStyle.Office2007
        Me.FrmGUImgr.ToolbarBottomDockSite = Me.DockSite8
        Me.FrmGUImgr.ToolbarLeftDockSite = Me.DockSite5
        Me.FrmGUImgr.ToolbarRightDockSite = Me.DockSite6
        Me.FrmGUImgr.ToolbarTopDockSite = Me.DockSite7
        Me.FrmGUImgr.TopDockSite = Me.DockSite3
        '
        'DockSite4
        '
        Me.DockSite4.AccessibleRole = System.Windows.Forms.AccessibleRole.Window
        Me.DockSite4.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.DockSite4.Location = New System.Drawing.Point(0, 680)
        Me.DockSite4.Name = "DockSite4"
        Me.DockSite4.Size = New System.Drawing.Size(984, 0)
        Me.DockSite4.TabIndex = 3
        Me.DockSite4.TabStop = False
        '
        'DockSite9
        '
        Me.DockSite9.AccessibleRole = System.Windows.Forms.AccessibleRole.Window
        Me.DockSite9.Controls.Add(Me.Bar2)
        Me.DockSite9.Controls.Add(Me.Bar1)
        Me.DockSite9.Dock = System.Windows.Forms.DockStyle.Fill
        Me.DockSite9.DocumentDockContainer = New DevComponents.DotNetBar.DocumentDockContainer(New DevComponents.DotNetBar.DocumentBaseContainer() {CType(New DevComponents.DotNetBar.DocumentBarContainer(Me.Bar2, 422, 665), DevComponents.DotNetBar.DocumentBaseContainer), CType(New DevComponents.DotNetBar.DocumentBarContainer(Me.Bar1, 591, 665), DevComponents.DotNetBar.DocumentBaseContainer)}, DevComponents.DotNetBar.eOrientation.Horizontal)
        Me.DockSite9.Location = New System.Drawing.Point(0, 67)
        Me.DockSite9.Name = "DockSite9"
        Me.DockSite9.Size = New System.Drawing.Size(984, 613)
        Me.DockSite9.TabIndex = 8
        Me.DockSite9.TabStop = False
        '
        'Bar2
        '
        Me.Bar2.AccessibleDescription = "DotNetBar Bar (Bar2)"
        Me.Bar2.AccessibleName = "DotNetBar Bar"
        Me.Bar2.AccessibleRole = System.Windows.Forms.AccessibleRole.ToolBar
        Me.Bar2.AlwaysDisplayDockTab = True
        Me.Bar2.CanCustomize = False
        Me.Bar2.CanDockBottom = False
        Me.Bar2.CanDockDocument = True
        Me.Bar2.CanDockLeft = False
        Me.Bar2.CanDockRight = False
        Me.Bar2.CanDockTop = False
        Me.Bar2.CanHide = True
        Me.Bar2.CanUndock = False
        Me.Bar2.Controls.Add(Me.PartList)
        Me.Bar2.Controls.Add(Me.Bar3)
        Me.Bar2.DockTabAlignment = DevComponents.DotNetBar.eTabStripAlignment.Top
        Me.Bar2.Font = New System.Drawing.Font("Verdana", 8.0!)
        Me.Bar2.LayoutType = DevComponents.DotNetBar.eLayoutType.DockContainer
        Me.Bar2.Location = New System.Drawing.Point(0, 0)
        Me.Bar2.Name = "Bar2"
        Me.Bar2.Size = New System.Drawing.Size(408, 613)
        Me.Bar2.Stretch = True
        Me.Bar2.Style = DevComponents.DotNetBar.eDotNetBarStyle.Office2007
        Me.Bar2.TabIndex = 0
        Me.Bar2.TabNavigation = True
        Me.Bar2.TabStop = False
        '
        'PartList
        '
        Me.PartList.AllowDrop = True
        '
        '
        '
        Me.PartList.Border.Class = "ListViewBorder"
        Me.PartList.Border.CornerType = DevComponents.DotNetBar.eCornerType.Square
        Me.PartList.Columns.AddRange(New System.Windows.Forms.ColumnHeader() {Me.PartColHd1, Me.PartColHd2, Me.PartColHd3, Me.PartColHd4, Me.PartColHd5})
        Me.PartList.Dock = System.Windows.Forms.DockStyle.Fill
        Me.PartList.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.PartList.FullRowSelect = True
        Me.PartList.HeaderStyle = System.Windows.Forms.ColumnHeaderStyle.Nonclickable
        Me.PartList.HoverSelection = True
        Me.PartList.Location = New System.Drawing.Point(0, 25)
        Me.PartList.MultiSelect = False
        Me.PartList.Name = "PartList"
        Me.PartList.Size = New System.Drawing.Size(408, 588)
        Me.PartList.Sorting = System.Windows.Forms.SortOrder.Ascending
        Me.PartList.TabIndex = 3
        Me.PartList.UseCompatibleStateImageBehavior = False
        Me.PartList.View = System.Windows.Forms.View.Details
        '
        'PartColHd1
        '
        Me.PartColHd1.Text = "품목번호"
        Me.PartColHd1.Width = 90
        '
        'PartColHd2
        '
        Me.PartColHd2.Text = "품목명"
        Me.PartColHd2.Width = 119
        '
        'PartColHd3
        '
        Me.PartColHd3.Text = "LG 품목번호"
        Me.PartColHd3.Width = 139
        '
        'PartColHd4
        '
        Me.PartColHd4.Text = "단가"
        Me.PartColHd4.Width = 86
        '
        'PartColHd5
        '
        Me.PartColHd5.Text = "품목구분"
        '
        'Bar3
        '
        Me.Bar3.Controls.Add(Me.LabelX1)
        Me.Bar3.Controls.Add(Me.TextBoxX1)
        Me.Bar3.Dock = System.Windows.Forms.DockStyle.Top
        Me.Bar3.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Bar3.Location = New System.Drawing.Point(0, 0)
        Me.Bar3.Name = "Bar3"
        Me.Bar3.Size = New System.Drawing.Size(408, 25)
        Me.Bar3.Stretch = True
        Me.Bar3.Style = DevComponents.DotNetBar.eDotNetBarStyle.Office2007
        Me.Bar3.TabIndex = 2
        Me.Bar3.TabStop = False
        Me.Bar3.Text = "Bar3"
        '
        'LabelX1
        '
        Me.LabelX1.BackColor = System.Drawing.Color.Transparent
        '
        '
        '
        Me.LabelX1.BackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square
        Me.LabelX1.Font = New System.Drawing.Font("Verdana", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LabelX1.Location = New System.Drawing.Point(15, 3)
        Me.LabelX1.Name = "LabelX1"
        Me.LabelX1.Size = New System.Drawing.Size(90, 18)
        Me.LabelX1.TabIndex = 1
        Me.LabelX1.Text = "품목 리스트"
        '
        'TextBoxX1
        '
        Me.TextBoxX1.BackColor = System.Drawing.Color.White
        '
        '
        '
        Me.TextBoxX1.Border.Class = "TextBoxBorder"
        Me.TextBoxX1.Border.CornerType = DevComponents.DotNetBar.eCornerType.Square
        Me.TextBoxX1.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TextBoxX1.FocusHighlightEnabled = True
        Me.TextBoxX1.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TextBoxX1.Location = New System.Drawing.Point(103, 1)
        Me.TextBoxX1.Name = "TextBoxX1"
        Me.TextBoxX1.Size = New System.Drawing.Size(186, 21)
        Me.TextBoxX1.TabIndex = 0
        '
        'Bar1
        '
        Me.Bar1.AccessibleDescription = "DotNetBar Bar (Bar1)"
        Me.Bar1.AccessibleName = "DotNetBar Bar"
        Me.Bar1.AccessibleRole = System.Windows.Forms.AccessibleRole.ToolBar
        Me.Bar1.AlwaysDisplayDockTab = True
        Me.Bar1.CanDockBottom = False
        Me.Bar1.CanDockDocument = True
        Me.Bar1.CanDockLeft = False
        Me.Bar1.CanDockRight = False
        Me.Bar1.CanDockTop = False
        Me.Bar1.CanHide = True
        Me.Bar1.CanUndock = False
        Me.Bar1.Controls.Add(Me.FpSpread1)
        Me.Bar1.Controls.Add(Me.bar4)
        Me.Bar1.DockTabAlignment = DevComponents.DotNetBar.eTabStripAlignment.Top
        Me.Bar1.Font = New System.Drawing.Font("Verdana", 8.0!)
        Me.Bar1.LayoutType = DevComponents.DotNetBar.eLayoutType.DockContainer
        Me.Bar1.Location = New System.Drawing.Point(411, 0)
        Me.Bar1.Name = "Bar1"
        Me.Bar1.Size = New System.Drawing.Size(573, 613)
        Me.Bar1.Stretch = True
        Me.Bar1.Style = DevComponents.DotNetBar.eDotNetBarStyle.Office2007
        Me.Bar1.TabIndex = 1
        Me.Bar1.TabStop = False
        '
        'FpSpread1
        '
        Me.FpSpread1.AccessibleDescription = ""
        Me.FpSpread1.AllowDragDrop = True
        Me.FpSpread1.AllowDragFill = True
        Me.FpSpread1.AllowDrop = True
        Me.FpSpread1.ClipboardPasteToFill = True
        Me.FpSpread1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.FpSpread1.Font = New System.Drawing.Font("Verdana", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FpSpread1.Location = New System.Drawing.Point(0, 25)
        Me.FpSpread1.Name = "FpSpread1"
        Me.FpSpread1.Sheets.AddRange(New FarPoint.Win.Spread.SheetView() {Me.FpSpread1_Sheet1})
        Me.FpSpread1.Size = New System.Drawing.Size(573, 588)
        Me.FpSpread1.TabIndex = 4
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
        'bar4
        '
        Me.bar4.Controls.Add(Me.LabelX2)
        Me.bar4.Dock = System.Windows.Forms.DockStyle.Top
        Me.bar4.Font = New System.Drawing.Font("Verdana", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.bar4.Location = New System.Drawing.Point(0, 0)
        Me.bar4.Name = "bar4"
        Me.bar4.Size = New System.Drawing.Size(573, 25)
        Me.bar4.Stretch = True
        Me.bar4.Style = DevComponents.DotNetBar.eDotNetBarStyle.Office2007
        Me.bar4.TabIndex = 1
        Me.bar4.TabStop = False
        '
        'LabelX2
        '
        Me.LabelX2.AutoSize = True
        Me.LabelX2.BackColor = System.Drawing.Color.Transparent
        '
        '
        '
        Me.LabelX2.BackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square
        Me.LabelX2.Font = New System.Drawing.Font("Verdana", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LabelX2.Location = New System.Drawing.Point(16, 3)
        Me.LabelX2.Name = "LabelX2"
        Me.LabelX2.Size = New System.Drawing.Size(71, 20)
        Me.LabelX2.TabIndex = 2
        Me.LabelX2.Text = "모델 번호"
        '
        'DockSite1
        '
        Me.DockSite1.AccessibleRole = System.Windows.Forms.AccessibleRole.Window
        Me.DockSite1.Dock = System.Windows.Forms.DockStyle.Left
        Me.DockSite1.DocumentDockContainer = New DevComponents.DotNetBar.DocumentDockContainer
        Me.DockSite1.Location = New System.Drawing.Point(0, 67)
        Me.DockSite1.Name = "DockSite1"
        Me.DockSite1.Size = New System.Drawing.Size(0, 613)
        Me.DockSite1.TabIndex = 0
        Me.DockSite1.TabStop = False
        '
        'DockSite2
        '
        Me.DockSite2.AccessibleRole = System.Windows.Forms.AccessibleRole.Window
        Me.DockSite2.Controls.Add(Me.ModelArea)
        Me.DockSite2.Dock = System.Windows.Forms.DockStyle.Right
        Me.DockSite2.DocumentDockContainer = New DevComponents.DotNetBar.DocumentDockContainer(New DevComponents.DotNetBar.DocumentBaseContainer() {CType(New DevComponents.DotNetBar.DocumentBarContainer(Me.ModelArea, -3, 613), DevComponents.DotNetBar.DocumentBaseContainer)}, DevComponents.DotNetBar.eOrientation.Horizontal)
        Me.DockSite2.Location = New System.Drawing.Point(984, 67)
        Me.DockSite2.Name = "DockSite2"
        Me.DockSite2.Size = New System.Drawing.Size(0, 613)
        Me.DockSite2.TabIndex = 1
        Me.DockSite2.TabStop = False
        '
        'ModelArea
        '
        Me.ModelArea.AccessibleDescription = "DotNetBar Bar (ModelArea)"
        Me.ModelArea.AccessibleName = "DotNetBar Bar"
        Me.ModelArea.AccessibleRole = System.Windows.Forms.AccessibleRole.ToolBar
        Me.ModelArea.AutoSyncBarCaption = True
        Me.ModelArea.CloseSingleTab = True
        Me.ModelArea.Font = New System.Drawing.Font("Verdana", 8.0!)
        Me.ModelArea.GrabHandleStyle = DevComponents.DotNetBar.eGrabHandleStyle.Caption
        Me.ModelArea.LayoutType = DevComponents.DotNetBar.eLayoutType.DockContainer
        Me.ModelArea.Location = New System.Drawing.Point(3, 0)
        Me.ModelArea.Name = "ModelArea"
        Me.ModelArea.Size = New System.Drawing.Size(2, 613)
        Me.ModelArea.Stretch = True
        Me.ModelArea.Style = DevComponents.DotNetBar.eDotNetBarStyle.Office2007
        Me.ModelArea.TabIndex = 0
        Me.ModelArea.TabStop = False
        Me.ModelArea.Text = "Model Name"
        '
        'DockSite8
        '
        Me.DockSite8.AccessibleRole = System.Windows.Forms.AccessibleRole.Window
        Me.DockSite8.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.DockSite8.Location = New System.Drawing.Point(0, 680)
        Me.DockSite8.Name = "DockSite8"
        Me.DockSite8.Size = New System.Drawing.Size(984, 0)
        Me.DockSite8.TabIndex = 7
        Me.DockSite8.TabStop = False
        '
        'DockSite5
        '
        Me.DockSite5.AccessibleRole = System.Windows.Forms.AccessibleRole.Window
        Me.DockSite5.Dock = System.Windows.Forms.DockStyle.Left
        Me.DockSite5.Location = New System.Drawing.Point(0, 0)
        Me.DockSite5.Name = "DockSite5"
        Me.DockSite5.Size = New System.Drawing.Size(0, 680)
        Me.DockSite5.TabIndex = 4
        Me.DockSite5.TabStop = False
        '
        'DockSite6
        '
        Me.DockSite6.AccessibleRole = System.Windows.Forms.AccessibleRole.Window
        Me.DockSite6.Dock = System.Windows.Forms.DockStyle.Right
        Me.DockSite6.Location = New System.Drawing.Point(984, 0)
        Me.DockSite6.Name = "DockSite6"
        Me.DockSite6.Size = New System.Drawing.Size(0, 680)
        Me.DockSite6.TabIndex = 5
        Me.DockSite6.TabStop = False
        '
        'DockSite7
        '
        Me.DockSite7.AccessibleRole = System.Windows.Forms.AccessibleRole.Window
        Me.DockSite7.Dock = System.Windows.Forms.DockStyle.Top
        Me.DockSite7.Location = New System.Drawing.Point(0, 0)
        Me.DockSite7.Name = "DockSite7"
        Me.DockSite7.Size = New System.Drawing.Size(984, 0)
        Me.DockSite7.TabIndex = 6
        Me.DockSite7.TabStop = False
        '
        'DockSite3
        '
        Me.DockSite3.AccessibleRole = System.Windows.Forms.AccessibleRole.Window
        Me.DockSite3.Controls.Add(Me.Bar5)
        Me.DockSite3.Controls.Add(Me.TopDock)
        Me.DockSite3.Dock = System.Windows.Forms.DockStyle.Top
        Me.DockSite3.DocumentDockContainer = New DevComponents.DotNetBar.DocumentDockContainer(New DevComponents.DotNetBar.DocumentBaseContainer() {CType(New DevComponents.DotNetBar.DocumentBarContainer(Me.TopDock, 229, 64), DevComponents.DotNetBar.DocumentBaseContainer), CType(New DevComponents.DotNetBar.DocumentBarContainer(Me.Bar5, 752, 64), DevComponents.DotNetBar.DocumentBaseContainer)}, DevComponents.DotNetBar.eOrientation.Horizontal)
        Me.DockSite3.Location = New System.Drawing.Point(0, 0)
        Me.DockSite3.Name = "DockSite3"
        Me.DockSite3.Size = New System.Drawing.Size(984, 67)
        Me.DockSite3.TabIndex = 2
        Me.DockSite3.TabStop = False
        '
        'Bar5
        '
        Me.Bar5.AccessibleDescription = "DotNetBar Bar (Bar5)"
        Me.Bar5.AccessibleName = "DotNetBar Bar"
        Me.Bar5.AccessibleRole = System.Windows.Forms.AccessibleRole.ToolBar
        Me.Bar5.AutoSyncBarCaption = True
        Me.Bar5.Controls.Add(Me.PanelDockContainer2)
        Me.Bar5.Font = New System.Drawing.Font("Verdana", 8.0!)
        Me.Bar5.GrabHandleStyle = DevComponents.DotNetBar.eGrabHandleStyle.Caption
        Me.Bar5.Items.AddRange(New DevComponents.DotNetBar.BaseItem() {Me.DockContainerItem10})
        Me.Bar5.LayoutType = DevComponents.DotNetBar.eLayoutType.DockContainer
        Me.Bar5.Location = New System.Drawing.Point(232, 0)
        Me.Bar5.Name = "Bar5"
        Me.Bar5.Size = New System.Drawing.Size(752, 64)
        Me.Bar5.Stretch = True
        Me.Bar5.Style = DevComponents.DotNetBar.eDotNetBarStyle.Office2007
        Me.Bar5.TabIndex = 1
        Me.Bar5.TabStop = False
        Me.Bar5.Text = "DockContainerItem10"
        '
        'PanelDockContainer2
        '
        Me.PanelDockContainer2.ColorSchemeStyle = DevComponents.DotNetBar.eDotNetBarStyle.Office2007
        Me.PanelDockContainer2.Controls.Add(Me.ModelLbl)
        Me.PanelDockContainer2.Controls.Add(Me.ActCb)
        Me.PanelDockContainer2.Controls.Add(Me.ActLbl)
        Me.PanelDockContainer2.Controls.Add(Me.ModelCb)
        Me.PanelDockContainer2.Controls.Add(Me.CheckBoxX1)
        Me.PanelDockContainer2.Location = New System.Drawing.Point(3, 23)
        Me.PanelDockContainer2.Name = "PanelDockContainer2"
        Me.PanelDockContainer2.Size = New System.Drawing.Size(746, 38)
        Me.PanelDockContainer2.Style.Alignment = System.Drawing.StringAlignment.Center
        Me.PanelDockContainer2.Style.BackColor1.ColorSchemePart = DevComponents.DotNetBar.eColorSchemePart.BarBackground
        Me.PanelDockContainer2.Style.BackColor2.ColorSchemePart = DevComponents.DotNetBar.eColorSchemePart.BarBackground2
        Me.PanelDockContainer2.Style.BorderColor.ColorSchemePart = DevComponents.DotNetBar.eColorSchemePart.BarDockedBorder
        Me.PanelDockContainer2.Style.ForeColor.ColorSchemePart = DevComponents.DotNetBar.eColorSchemePart.ItemText
        Me.PanelDockContainer2.Style.GradientAngle = 90
        Me.PanelDockContainer2.TabIndex = 2
        '
        'ModelLbl
        '
        Me.ModelLbl.AutoSize = True
        Me.ModelLbl.BackColor = System.Drawing.Color.Transparent
        '
        '
        '
        Me.ModelLbl.BackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square
        Me.ModelLbl.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ModelLbl.Location = New System.Drawing.Point(24, 10)
        Me.ModelLbl.Name = "ModelLbl"
        Me.ModelLbl.Size = New System.Drawing.Size(31, 17)
        Me.ModelLbl.TabIndex = 0
        Me.ModelLbl.Text = "모델"
        Me.ModelLbl.TextAlignment = System.Drawing.StringAlignment.Center
        '
        'ActCb
        '
        Me.ActCb.DisplayMember = "Text"
        Me.ActCb.DrawMode = System.Windows.Forms.DrawMode.OwnerDrawFixed
        Me.ActCb.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ActCb.FormattingEnabled = True
        Me.ActCb.ItemHeight = 15
        Me.ActCb.Items.AddRange(New Object() {Me.ALL, Me.YES, Me.NO})
        Me.ActCb.Location = New System.Drawing.Point(301, 8)
        Me.ActCb.Name = "ActCb"
        Me.ActCb.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.ActCb.Size = New System.Drawing.Size(56, 21)
        Me.ActCb.TabIndex = 3
        Me.ActCb.Text = "ALL"
        '
        'ALL
        '
        Me.ALL.Text = "ALL"
        '
        'YES
        '
        Me.YES.Text = "YES"
        '
        'NO
        '
        Me.NO.Text = "NO"
        '
        'ActLbl
        '
        Me.ActLbl.AutoSize = True
        Me.ActLbl.BackColor = System.Drawing.Color.Transparent
        '
        '
        '
        Me.ActLbl.BackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square
        Me.ActLbl.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ActLbl.Location = New System.Drawing.Point(242, 11)
        Me.ActLbl.Name = "ActLbl"
        Me.ActLbl.Size = New System.Drawing.Size(56, 17)
        Me.ActLbl.TabIndex = 1
        Me.ActLbl.Text = "사용여부"
        Me.ActLbl.TextAlignment = System.Drawing.StringAlignment.Center
        '
        'ModelCb
        '
        Me.ModelCb.DisplayMember = "Text"
        Me.ModelCb.DrawMode = System.Windows.Forms.DrawMode.OwnerDrawFixed
        Me.ModelCb.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ModelCb.FormattingEnabled = True
        Me.ModelCb.ItemHeight = 15
        Me.ModelCb.Location = New System.Drawing.Point(59, 8)
        Me.ModelCb.Name = "ModelCb"
        Me.ModelCb.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.ModelCb.Size = New System.Drawing.Size(177, 21)
        Me.ModelCb.TabIndex = 2
        '
        'CheckBoxX1
        '
        Me.CheckBoxX1.BackColor = System.Drawing.Color.Transparent
        '
        '
        '
        Me.CheckBoxX1.BackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square
        Me.CheckBoxX1.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CheckBoxX1.Location = New System.Drawing.Point(323, 8)
        Me.CheckBoxX1.Name = "CheckBoxX1"
        Me.CheckBoxX1.Size = New System.Drawing.Size(112, 21)
        Me.CheckBoxX1.TabIndex = 6
        Me.CheckBoxX1.Text = "Upload Excel"
        Me.CheckBoxX1.Visible = False
        '
        'DockContainerItem10
        '
        Me.DockContainerItem10.Control = Me.PanelDockContainer2
        Me.DockContainerItem10.Name = "DockContainerItem10"
        Me.DockContainerItem10.Text = "DockContainerItem10"
        '
        'TopDock
        '
        Me.TopDock.AccessibleDescription = "DotNetBar Bar (TopDock)"
        Me.TopDock.AccessibleName = "DotNetBar Bar"
        Me.TopDock.AccessibleRole = System.Windows.Forms.AccessibleRole.ToolBar
        Me.TopDock.AutoSyncBarCaption = True
        Me.TopDock.CloseSingleTab = True
        Me.TopDock.Controls.Add(Me.MenuPanel)
        Me.TopDock.Font = New System.Drawing.Font("Verdana", 8.0!)
        Me.TopDock.GrabHandleStyle = DevComponents.DotNetBar.eGrabHandleStyle.Caption
        Me.TopDock.Items.AddRange(New DevComponents.DotNetBar.BaseItem() {Me.DockContainerItem1})
        Me.TopDock.LayoutType = DevComponents.DotNetBar.eLayoutType.DockContainer
        Me.TopDock.Location = New System.Drawing.Point(0, 0)
        Me.TopDock.Name = "TopDock"
        Me.TopDock.Size = New System.Drawing.Size(229, 64)
        Me.TopDock.Stretch = True
        Me.TopDock.Style = DevComponents.DotNetBar.eDotNetBarStyle.Office2007
        Me.TopDock.TabIndex = 0
        Me.TopDock.TabStop = False
        Me.TopDock.Text = "DockContainerItem1"
        '
        'MenuPanel
        '
        Me.MenuPanel.ColorSchemeStyle = DevComponents.DotNetBar.eDotNetBarStyle.Office2007
        Me.MenuPanel.Controls.Add(Me.CtlBar)
        Me.MenuPanel.Location = New System.Drawing.Point(3, 23)
        Me.MenuPanel.Name = "MenuPanel"
        Me.MenuPanel.Size = New System.Drawing.Size(223, 38)
        Me.MenuPanel.Style.Alignment = System.Drawing.StringAlignment.Center
        Me.MenuPanel.Style.BackColor1.ColorSchemePart = DevComponents.DotNetBar.eColorSchemePart.BarBackground
        Me.MenuPanel.Style.BackColor2.ColorSchemePart = DevComponents.DotNetBar.eColorSchemePart.BarBackground2
        Me.MenuPanel.Style.BorderColor.ColorSchemePart = DevComponents.DotNetBar.eColorSchemePart.BarDockedBorder
        Me.MenuPanel.Style.ForeColor.ColorSchemePart = DevComponents.DotNetBar.eColorSchemePart.ItemText
        Me.MenuPanel.Style.GradientAngle = 90
        Me.MenuPanel.TabIndex = 0
        '
        'CtlBar
        '
        Me.CtlBar.Dock = System.Windows.Forms.DockStyle.Fill
        Me.CtlBar.Font = New System.Drawing.Font("Malgun Gothic", 9.0!)
        Me.CtlBar.Items.AddRange(New DevComponents.DotNetBar.BaseItem() {Me.FindBtn, Me.SaveBtn, Me.DelBtn, Me.PrtBtn, Me.Excel})
        Me.CtlBar.Location = New System.Drawing.Point(0, 0)
        Me.CtlBar.Name = "CtlBar"
        Me.CtlBar.Size = New System.Drawing.Size(223, 41)
        Me.CtlBar.Stretch = True
        Me.CtlBar.Style = DevComponents.DotNetBar.eDotNetBarStyle.Office2007
        Me.CtlBar.TabIndex = 1
        Me.CtlBar.TabStop = False
        Me.CtlBar.Text = "Bar2"
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
        Me.DockContainerItem1.Control = Me.MenuPanel
        Me.DockContainerItem1.Name = "DockContainerItem1"
        Me.DockContainerItem1.Text = "DockContainerItem1"
        '
        'PanelDockContainer1
        '
        Me.PanelDockContainer1.ColorSchemeStyle = DevComponents.DotNetBar.eDotNetBarStyle.Office2007
        Me.PanelDockContainer1.Location = New System.Drawing.Point(3, 23)
        Me.PanelDockContainer1.Name = "PanelDockContainer1"
        Me.PanelDockContainer1.Size = New System.Drawing.Size(434, 615)
        Me.PanelDockContainer1.Style.Alignment = System.Drawing.StringAlignment.Center
        Me.PanelDockContainer1.Style.BackColor1.ColorSchemePart = DevComponents.DotNetBar.eColorSchemePart.BarBackground
        Me.PanelDockContainer1.Style.BackColor2.ColorSchemePart = DevComponents.DotNetBar.eColorSchemePart.BarBackground2
        Me.PanelDockContainer1.Style.BorderColor.ColorSchemePart = DevComponents.DotNetBar.eColorSchemePart.BarDockedBorder
        Me.PanelDockContainer1.Style.ForeColor.ColorSchemePart = DevComponents.DotNetBar.eColorSchemePart.ItemText
        Me.PanelDockContainer1.Style.GradientAngle = 90
        Me.PanelDockContainer1.TabIndex = 0
        '
        'DockContainerItem5
        '
        Me.DockContainerItem5.Name = "DockContainerItem5"
        Me.DockContainerItem5.Text = "DockContainerItem5"
        '
        'CtxSp
        '
        Me.CtxSp.AutoExpandOnClick = True
        Me.CtxSp.Name = "CtxSp"
        Me.CtxSp.SubItems.AddRange(New DevComponents.DotNetBar.BaseItem() {Me.bDel})
        Me.CtxSp.Text = "FPspread"
        '
        'bDel
        '
        Me.bDel.Image = Global.SMSPLUS.My.Resources.Resources._14_layer_deletelayer
        Me.bDel.Name = "bDel"
        Me.bDel.Text = "&Delete"
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
        'DockContainerItem2
        '
        Me.DockContainerItem2.Name = "DockContainerItem2"
        Me.DockContainerItem2.Text = "DockContainerItem2"
        '
        'ContextMenuBar1
        '
        Me.ContextMenuBar1.Dock = System.Windows.Forms.DockStyle.Top
        Me.ContextMenuBar1.DockSide = DevComponents.DotNetBar.eDockSide.Top
        Me.ContextMenuBar1.Font = New System.Drawing.Font("Malgun Gothic", 9.0!)
        Me.ContextMenuBar1.Items.AddRange(New DevComponents.DotNetBar.BaseItem() {Me.CtxSp})
        Me.ContextMenuBar1.Location = New System.Drawing.Point(0, 0)
        Me.ContextMenuBar1.Name = "ContextMenuBar1"
        Me.ContextMenuBar1.Size = New System.Drawing.Size(75, 27)
        Me.ContextMenuBar1.Stretch = True
        Me.ContextMenuBar1.Style = DevComponents.DotNetBar.eDotNetBarStyle.Office2007
        Me.ContextMenuBar1.TabIndex = 4
        Me.ContextMenuBar1.TabStop = False
        Me.ContextMenuBar1.Text = "ContextMenuBar1"
        '
        'DockContainerItem4
        '
        Me.DockContainerItem4.Name = "DockContainerItem4"
        Me.DockContainerItem4.Text = "DockContainerItem4"
        '
        'DockContainerItem3
        '
        Me.DockContainerItem3.Name = "DockContainerItem3"
        Me.DockContainerItem3.Text = "DockContainerItem3"
        '
        'DockContainerItem6
        '
        Me.DockContainerItem6.Name = "DockContainerItem6"
        Me.DockContainerItem6.Text = "DockContainerItem6"
        '
        'DockContainerItem7
        '
        Me.DockContainerItem7.Control = Me.PanelDockContainer1
        Me.DockContainerItem7.Name = "DockContainerItem7"
        Me.DockContainerItem7.Text = "DockContainerItem7"
        '
        'DockContainerItem9
        '
        Me.DockContainerItem9.Name = "DockContainerItem9"
        Me.DockContainerItem9.Text = "DockContainerItem9"
        '
        'DockContainerItem8
        '
        Me.DockContainerItem8.Name = "DockContainerItem8"
        Me.DockContainerItem8.Text = "DockContainerItem8"
        '
        'ControlContainerItem2
        '
        Me.ControlContainerItem2.AllowItemResize = False
        Me.ControlContainerItem2.MenuVisibility = DevComponents.DotNetBar.eMenuVisibility.VisibleAlways
        Me.ControlContainerItem2.Name = "ControlContainerItem2"
        '
        'XlsBtn1
        '
        Me.XlsBtn1.AccessibleRole = System.Windows.Forms.AccessibleRole.PushButton
        Me.XlsBtn1.ColorTable = DevComponents.DotNetBar.eButtonColor.OrangeWithBackground
        Me.XlsBtn1.Location = New System.Drawing.Point(386, 382)
        Me.XlsBtn1.Name = "XlsBtn1"
        Me.XlsBtn1.Size = New System.Drawing.Size(66, 21)
        Me.XlsBtn1.TabIndex = 23
        Me.XlsBtn1.Text = "E&Xcel"
        '
        'PrtBtn1
        '
        Me.PrtBtn1.AccessibleRole = System.Windows.Forms.AccessibleRole.PushButton
        Me.PrtBtn1.ColorTable = DevComponents.DotNetBar.eButtonColor.OrangeWithBackground
        Me.PrtBtn1.Location = New System.Drawing.Point(386, 355)
        Me.PrtBtn1.Name = "PrtBtn1"
        Me.PrtBtn1.Size = New System.Drawing.Size(66, 21)
        Me.PrtBtn1.TabIndex = 22
        Me.PrtBtn1.Text = "&Print"
        '
        'DelBtn1
        '
        Me.DelBtn1.AccessibleRole = System.Windows.Forms.AccessibleRole.PushButton
        Me.DelBtn1.ColorTable = DevComponents.DotNetBar.eButtonColor.OrangeWithBackground
        Me.DelBtn1.Location = New System.Drawing.Point(386, 328)
        Me.DelBtn1.Name = "DelBtn1"
        Me.DelBtn1.Size = New System.Drawing.Size(66, 21)
        Me.DelBtn1.TabIndex = 21
        Me.DelBtn1.Text = "&Delete"
        '
        'SaveBtn1
        '
        Me.SaveBtn1.AccessibleRole = System.Windows.Forms.AccessibleRole.PushButton
        Me.SaveBtn1.ColorTable = DevComponents.DotNetBar.eButtonColor.OrangeWithBackground
        Me.SaveBtn1.Location = New System.Drawing.Point(386, 301)
        Me.SaveBtn1.Name = "SaveBtn1"
        Me.SaveBtn1.Size = New System.Drawing.Size(66, 21)
        Me.SaveBtn1.TabIndex = 20
        Me.SaveBtn1.Text = "&Save"
        '
        'NewBtn1
        '
        Me.NewBtn1.AccessibleRole = System.Windows.Forms.AccessibleRole.PushButton
        Me.NewBtn1.ColorTable = DevComponents.DotNetBar.eButtonColor.OrangeWithBackground
        Me.NewBtn1.Location = New System.Drawing.Point(386, 274)
        Me.NewBtn1.Name = "NewBtn1"
        Me.NewBtn1.Size = New System.Drawing.Size(66, 21)
        Me.NewBtn1.TabIndex = 19
        Me.NewBtn1.Text = "&New"
        '
        'OpenFileDialog1
        '
        Me.OpenFileDialog1.FileName = "OpenFileDialog1"
        '
        'FrmBOM
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(984, 680)
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
        Me.Font = New System.Drawing.Font("Verdana", 8.0!)
        Me.Name = "FrmBOM"
        Me.Text = "FrmBOM"
        Me.DockSite9.ResumeLayout(False)
        CType(Me.Bar2, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Bar2.ResumeLayout(False)
        CType(Me.Bar3, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Bar3.ResumeLayout(False)
        CType(Me.Bar1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Bar1.ResumeLayout(False)
        CType(Me.FpSpread1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.FpSpread1_Sheet1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.bar4, System.ComponentModel.ISupportInitialize).EndInit()
        Me.bar4.ResumeLayout(False)
        Me.bar4.PerformLayout()
        Me.DockSite2.ResumeLayout(False)
        CType(Me.ModelArea, System.ComponentModel.ISupportInitialize).EndInit()
        Me.DockSite3.ResumeLayout(False)
        CType(Me.Bar5, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Bar5.ResumeLayout(False)
        Me.PanelDockContainer2.ResumeLayout(False)
        Me.PanelDockContainer2.PerformLayout()
        CType(Me.TopDock, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TopDock.ResumeLayout(False)
        Me.MenuPanel.ResumeLayout(False)
        CType(Me.CtlBar, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.ContextMenuBar1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents FrmGUImgr As DevComponents.DotNetBar.DotNetBarManager
    Friend WithEvents DockSite4 As DevComponents.DotNetBar.DockSite
    Friend WithEvents DockSite1 As DevComponents.DotNetBar.DockSite
    Friend WithEvents DockSite2 As DevComponents.DotNetBar.DockSite
    Friend WithEvents DockSite3 As DevComponents.DotNetBar.DockSite
    Friend WithEvents DockSite5 As DevComponents.DotNetBar.DockSite
    Friend WithEvents DockSite6 As DevComponents.DotNetBar.DockSite
    Friend WithEvents DockSite7 As DevComponents.DotNetBar.DockSite
    Friend WithEvents DockSite8 As DevComponents.DotNetBar.DockSite
    Friend WithEvents NO As DevComponents.Editors.ComboItem
    Friend WithEvents YES As DevComponents.Editors.ComboItem
    Friend WithEvents ALL As DevComponents.Editors.ComboItem
    Friend WithEvents SaveBtn As DevComponents.DotNetBar.ButtonItem
    Friend WithEvents DelBtn As DevComponents.DotNetBar.ButtonItem
    Friend WithEvents PrtBtn As DevComponents.DotNetBar.ButtonItem
    Friend WithEvents TopDock As DevComponents.DotNetBar.Bar
    Friend WithEvents MenuPanel As DevComponents.DotNetBar.PanelDockContainer
    Friend WithEvents CtlBar As DevComponents.DotNetBar.Bar
    Friend WithEvents DockContainerItem1 As DevComponents.DotNetBar.DockContainerItem
    Friend WithEvents DockContainerItem2 As DevComponents.DotNetBar.DockContainerItem
    Friend WithEvents ModelLbl As DevComponents.DotNetBar.LabelX
    Friend WithEvents ActCb As DevComponents.DotNetBar.Controls.ComboBoxEx
    Friend WithEvents ModelCb As DevComponents.DotNetBar.Controls.ComboBoxEx
    Friend WithEvents ActLbl As DevComponents.DotNetBar.LabelX
    Friend WithEvents CtxSp As DevComponents.DotNetBar.ButtonItem
    Friend WithEvents bSave As DevComponents.DotNetBar.ButtonItem
    Friend WithEvents bDel As DevComponents.DotNetBar.ButtonItem
    Friend WithEvents bPrint As DevComponents.DotNetBar.ButtonItem
    Friend WithEvents Excel As DevComponents.DotNetBar.ButtonItem
    Friend WithEvents bExcel As DevComponents.DotNetBar.ButtonItem
    Friend WithEvents SaveFileDialog1 As System.Windows.Forms.SaveFileDialog
    Friend WithEvents ContextMenuBar1 As DevComponents.DotNetBar.ContextMenuBar
    Friend WithEvents DockContainerItem5 As DevComponents.DotNetBar.DockContainerItem
    Friend WithEvents DockContainerItem4 As DevComponents.DotNetBar.DockContainerItem
    Friend WithEvents DockContainerItem3 As DevComponents.DotNetBar.DockContainerItem
    Friend WithEvents DockContainerItem6 As DevComponents.DotNetBar.DockContainerItem
    Friend WithEvents DockSite9 As DevComponents.DotNetBar.DockSite
    Friend WithEvents Bar2 As DevComponents.DotNetBar.Bar
    Friend WithEvents DockContainerItem9 As DevComponents.DotNetBar.DockContainerItem
    Friend WithEvents DockContainerItem7 As DevComponents.DotNetBar.DockContainerItem
    Friend WithEvents ModelArea As DevComponents.DotNetBar.Bar
    Friend WithEvents PanelDockContainer1 As DevComponents.DotNetBar.PanelDockContainer
    Friend WithEvents Bar3 As DevComponents.DotNetBar.Bar
    Friend WithEvents Bar1 As DevComponents.DotNetBar.Bar
    Friend WithEvents bar4 As DevComponents.DotNetBar.Bar
    Friend WithEvents DockContainerItem8 As DevComponents.DotNetBar.DockContainerItem
    Friend WithEvents PartList As DevComponents.DotNetBar.Controls.ListViewEx
    Friend WithEvents PartColHd1 As System.Windows.Forms.ColumnHeader
    Friend WithEvents PartColHd2 As System.Windows.Forms.ColumnHeader
    Friend WithEvents PartColHd3 As System.Windows.Forms.ColumnHeader
    Friend WithEvents PartColHd4 As System.Windows.Forms.ColumnHeader
    Friend WithEvents FpSpread1 As FarPoint.Win.Spread.FpSpread
    Friend WithEvents FpSpread1_Sheet1 As FarPoint.Win.Spread.SheetView
    Friend WithEvents ControlContainerItem2 As DevComponents.DotNetBar.ControlContainerItem
    Friend WithEvents Bar5 As DevComponents.DotNetBar.Bar
    Friend WithEvents PanelDockContainer2 As DevComponents.DotNetBar.PanelDockContainer
    Friend WithEvents DockContainerItem10 As DevComponents.DotNetBar.DockContainerItem
    Friend WithEvents CheckBoxX1 As DevComponents.DotNetBar.Controls.CheckBoxX
    ' Friend WithEvents UcXlsUpload1 As EMS.UCXlsUpload
    Friend WithEvents XlsBtn1 As DevComponents.DotNetBar.ButtonX
    Friend WithEvents PrtBtn1 As DevComponents.DotNetBar.ButtonX
    Friend WithEvents DelBtn1 As DevComponents.DotNetBar.ButtonX
    Friend WithEvents SaveBtn1 As DevComponents.DotNetBar.ButtonX
    Friend WithEvents NewBtn1 As DevComponents.DotNetBar.ButtonX
    Friend WithEvents TextBoxX1 As DevComponents.DotNetBar.Controls.TextBoxX
    Friend WithEvents LabelX1 As DevComponents.DotNetBar.LabelX
    Friend WithEvents LabelX2 As DevComponents.DotNetBar.LabelX
    Friend WithEvents PartColHd5 As System.Windows.Forms.ColumnHeader
    Friend WithEvents OpenFileDialog1 As System.Windows.Forms.OpenFileDialog
    Friend WithEvents FindBtn As DevComponents.DotNetBar.ButtonItem
End Class
