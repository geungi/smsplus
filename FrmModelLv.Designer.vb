﻿<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmModelLv
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
        Me.FrmGUImgr = New DevComponents.DotNetBar.DotNetBarManager(Me.components)
        Me.DockSite4 = New DevComponents.DotNetBar.DockSite
        Me.DockSite9 = New DevComponents.DotNetBar.DockSite
        Me.MainContets = New DevComponents.DotNetBar.Bar
        Me.ContextMenuBar1 = New DevComponents.DotNetBar.ContextMenuBar
        Me.CtxSp = New DevComponents.DotNetBar.ButtonItem
        Me.bDel = New DevComponents.DotNetBar.ButtonItem
        Me.FpSpread1 = New FarPoint.Win.Spread.FpSpread
        Me.FpSpread1_Sheet1 = New FarPoint.Win.Spread.SheetView
        Me.DockSite1 = New DevComponents.DotNetBar.DockSite
        Me.DockSite2 = New DevComponents.DotNetBar.DockSite
        Me.DockSite8 = New DevComponents.DotNetBar.DockSite
        Me.DockSite5 = New DevComponents.DotNetBar.DockSite
        Me.DockSite6 = New DevComponents.DotNetBar.DockSite
        Me.DockSite7 = New DevComponents.DotNetBar.DockSite
        Me.DockSite3 = New DevComponents.DotNetBar.DockSite
        Me.Bar1 = New DevComponents.DotNetBar.Bar
        Me.PanelDockContainer1 = New DevComponents.DotNetBar.PanelDockContainer
        Me.ClassLbl = New DevComponents.DotNetBar.LabelX
        Me.ClassCb = New DevComponents.DotNetBar.Controls.ComboBoxEx
        Me.FindBtn = New DevComponents.DotNetBar.ButtonX
        Me.ModelLbl = New DevComponents.DotNetBar.LabelX
        Me.ModelLvCb = New DevComponents.DotNetBar.Controls.ComboBoxEx
        Me.DockContainerItem3 = New DevComponents.DotNetBar.DockContainerItem
        Me.TopDock = New DevComponents.DotNetBar.Bar
        Me.MenuPanel = New DevComponents.DotNetBar.PanelDockContainer
        Me.CtlBar = New DevComponents.DotNetBar.Bar
        Me.NewBtn = New DevComponents.DotNetBar.ButtonItem
        Me.SaveBtn = New DevComponents.DotNetBar.ButtonItem
        Me.DelBtn = New DevComponents.DotNetBar.ButtonItem
        Me.PrtBtn = New DevComponents.DotNetBar.ButtonItem
        Me.Excel = New DevComponents.DotNetBar.ButtonItem
        Me.DockContainerItem1 = New DevComponents.DotNetBar.DockContainerItem
        Me.bNew = New DevComponents.DotNetBar.ButtonItem
        Me.bSave = New DevComponents.DotNetBar.ButtonItem
        Me.bPrint = New DevComponents.DotNetBar.ButtonItem
        Me.bExcel = New DevComponents.DotNetBar.ButtonItem
        Me.DockContainerItem2 = New DevComponents.DotNetBar.DockContainerItem
        Me.SaveFileDialog1 = New System.Windows.Forms.SaveFileDialog
        Me.XlsBtn1 = New DevComponents.DotNetBar.ButtonX
        Me.PrtBtn1 = New DevComponents.DotNetBar.ButtonX
        Me.DelBtn1 = New DevComponents.DotNetBar.ButtonX
        Me.SaveBtn1 = New DevComponents.DotNetBar.ButtonX
        Me.NewBtn1 = New DevComponents.DotNetBar.ButtonX
        Me.DockSite9.SuspendLayout()
        CType(Me.MainContets, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.MainContets.SuspendLayout()
        CType(Me.ContextMenuBar1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.FpSpread1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.FpSpread1_Sheet1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.DockSite3.SuspendLayout()
        CType(Me.Bar1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Bar1.SuspendLayout()
        Me.PanelDockContainer1.SuspendLayout()
        CType(Me.TopDock, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TopDock.SuspendLayout()
        Me.MenuPanel.SuspendLayout()
        CType(Me.CtlBar, System.ComponentModel.ISupportInitialize).BeginInit()
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
        Me.FrmGUImgr.UseHook = True
        '
        'DockSite4
        '
        Me.DockSite4.AccessibleRole = System.Windows.Forms.AccessibleRole.Window
        Me.DockSite4.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.DockSite4.Location = New System.Drawing.Point(0, 676)
        Me.DockSite4.Name = "DockSite4"
        Me.DockSite4.Size = New System.Drawing.Size(977, 0)
        Me.DockSite4.TabIndex = 3
        Me.DockSite4.TabStop = False
        '
        'DockSite9
        '
        Me.DockSite9.AccessibleRole = System.Windows.Forms.AccessibleRole.Window
        Me.DockSite9.Controls.Add(Me.MainContets)
        Me.DockSite9.Dock = System.Windows.Forms.DockStyle.Fill
        Me.DockSite9.DocumentDockContainer = New DevComponents.DotNetBar.DocumentDockContainer(New DevComponents.DotNetBar.DocumentBaseContainer() {CType(New DevComponents.DotNetBar.DocumentBarContainer(Me.MainContets, 977, 611), DevComponents.DotNetBar.DocumentBaseContainer)}, DevComponents.DotNetBar.eOrientation.Horizontal)
        Me.DockSite9.Location = New System.Drawing.Point(0, 65)
        Me.DockSite9.Name = "DockSite9"
        Me.DockSite9.Size = New System.Drawing.Size(977, 611)
        Me.DockSite9.TabIndex = 9
        Me.DockSite9.TabStop = False
        '
        'MainContets
        '
        Me.MainContets.AccessibleDescription = "DotNetBar Bar (MainContets)"
        Me.MainContets.AccessibleName = "DotNetBar Bar"
        Me.MainContets.AccessibleRole = System.Windows.Forms.AccessibleRole.ToolBar
        Me.MainContets.AlwaysDisplayDockTab = True
        Me.MainContets.CanCustomize = False
        Me.MainContets.CanDockBottom = False
        Me.MainContets.CanDockDocument = True
        Me.MainContets.CanDockLeft = False
        Me.MainContets.CanDockRight = False
        Me.MainContets.CanDockTop = False
        Me.MainContets.CanHide = True
        Me.MainContets.CanUndock = False
        Me.MainContets.Controls.Add(Me.ContextMenuBar1)
        Me.MainContets.Controls.Add(Me.FpSpread1)
        Me.MainContets.DockTabAlignment = DevComponents.DotNetBar.eTabStripAlignment.Top
        Me.MainContets.Font = New System.Drawing.Font("Verdana", 8.0!)
        Me.MainContets.LayoutType = DevComponents.DotNetBar.eLayoutType.DockContainer
        Me.MainContets.Location = New System.Drawing.Point(0, 0)
        Me.MainContets.Name = "MainContets"
        Me.MainContets.Size = New System.Drawing.Size(977, 611)
        Me.MainContets.Stretch = True
        Me.MainContets.Style = DevComponents.DotNetBar.eDotNetBarStyle.Office2007
        Me.MainContets.TabIndex = 0
        Me.MainContets.TabNavigation = True
        Me.MainContets.TabStop = False
        '
        'ContextMenuBar1
        '
        Me.ContextMenuBar1.Font = New System.Drawing.Font("Malgun Gothic", 9.0!)
        Me.ContextMenuBar1.Items.AddRange(New DevComponents.DotNetBar.BaseItem() {Me.CtxSp})
        Me.ContextMenuBar1.Location = New System.Drawing.Point(0, 0)
        Me.ContextMenuBar1.Name = "ContextMenuBar1"
        Me.ContextMenuBar1.Size = New System.Drawing.Size(66, 27)
        Me.ContextMenuBar1.Stretch = True
        Me.ContextMenuBar1.Style = DevComponents.DotNetBar.eDotNetBarStyle.Office2003
        Me.ContextMenuBar1.TabIndex = 4
        Me.ContextMenuBar1.TabStop = False
        Me.ContextMenuBar1.Text = "ContextMenuBar1"
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
        'FpSpread1
        '
        Me.FpSpread1.AccessibleDescription = ""
        Me.FpSpread1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.FpSpread1.Font = New System.Drawing.Font("Verdana", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FpSpread1.Location = New System.Drawing.Point(0, 0)
        Me.FpSpread1.Name = "FpSpread1"
        Me.FpSpread1.Sheets.AddRange(New FarPoint.Win.Spread.SheetView() {Me.FpSpread1_Sheet1})
        Me.FpSpread1.Size = New System.Drawing.Size(977, 611)
        Me.FpSpread1.TabIndex = 6
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
        'DockSite1
        '
        Me.DockSite1.AccessibleRole = System.Windows.Forms.AccessibleRole.Window
        Me.DockSite1.Dock = System.Windows.Forms.DockStyle.Left
        Me.DockSite1.Location = New System.Drawing.Point(0, 65)
        Me.DockSite1.Name = "DockSite1"
        Me.DockSite1.Size = New System.Drawing.Size(0, 611)
        Me.DockSite1.TabIndex = 0
        Me.DockSite1.TabStop = False
        '
        'DockSite2
        '
        Me.DockSite2.AccessibleRole = System.Windows.Forms.AccessibleRole.Window
        Me.DockSite2.Dock = System.Windows.Forms.DockStyle.Right
        Me.DockSite2.Location = New System.Drawing.Point(977, 65)
        Me.DockSite2.Name = "DockSite2"
        Me.DockSite2.Size = New System.Drawing.Size(0, 611)
        Me.DockSite2.TabIndex = 1
        Me.DockSite2.TabStop = False
        '
        'DockSite8
        '
        Me.DockSite8.AccessibleRole = System.Windows.Forms.AccessibleRole.Window
        Me.DockSite8.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.DockSite8.Location = New System.Drawing.Point(0, 676)
        Me.DockSite8.Name = "DockSite8"
        Me.DockSite8.Size = New System.Drawing.Size(977, 0)
        Me.DockSite8.TabIndex = 7
        Me.DockSite8.TabStop = False
        '
        'DockSite5
        '
        Me.DockSite5.AccessibleRole = System.Windows.Forms.AccessibleRole.Window
        Me.DockSite5.Dock = System.Windows.Forms.DockStyle.Left
        Me.DockSite5.Location = New System.Drawing.Point(0, 0)
        Me.DockSite5.Name = "DockSite5"
        Me.DockSite5.Size = New System.Drawing.Size(0, 676)
        Me.DockSite5.TabIndex = 4
        Me.DockSite5.TabStop = False
        '
        'DockSite6
        '
        Me.DockSite6.AccessibleRole = System.Windows.Forms.AccessibleRole.Window
        Me.DockSite6.Dock = System.Windows.Forms.DockStyle.Right
        Me.DockSite6.Location = New System.Drawing.Point(977, 0)
        Me.DockSite6.Name = "DockSite6"
        Me.DockSite6.Size = New System.Drawing.Size(0, 676)
        Me.DockSite6.TabIndex = 5
        Me.DockSite6.TabStop = False
        '
        'DockSite7
        '
        Me.DockSite7.AccessibleRole = System.Windows.Forms.AccessibleRole.Window
        Me.DockSite7.Dock = System.Windows.Forms.DockStyle.Top
        Me.DockSite7.Location = New System.Drawing.Point(0, 0)
        Me.DockSite7.Name = "DockSite7"
        Me.DockSite7.Size = New System.Drawing.Size(977, 0)
        Me.DockSite7.TabIndex = 6
        Me.DockSite7.TabStop = False
        '
        'DockSite3
        '
        Me.DockSite3.AccessibleRole = System.Windows.Forms.AccessibleRole.Window
        Me.DockSite3.Controls.Add(Me.Bar1)
        Me.DockSite3.Controls.Add(Me.TopDock)
        Me.DockSite3.Dock = System.Windows.Forms.DockStyle.Top
        Me.DockSite3.DocumentDockContainer = New DevComponents.DotNetBar.DocumentDockContainer(New DevComponents.DotNetBar.DocumentBaseContainer() {CType(New DevComponents.DotNetBar.DocumentBarContainer(Me.TopDock, 256, 62), DevComponents.DotNetBar.DocumentBaseContainer), CType(New DevComponents.DotNetBar.DocumentBarContainer(Me.Bar1, 718, 62), DevComponents.DotNetBar.DocumentBaseContainer)}, DevComponents.DotNetBar.eOrientation.Horizontal)
        Me.DockSite3.Location = New System.Drawing.Point(0, 0)
        Me.DockSite3.Name = "DockSite3"
        Me.DockSite3.Size = New System.Drawing.Size(977, 65)
        Me.DockSite3.TabIndex = 2
        Me.DockSite3.TabStop = False
        '
        'Bar1
        '
        Me.Bar1.AccessibleDescription = "DotNetBar Bar (Bar1)"
        Me.Bar1.AccessibleName = "DotNetBar Bar"
        Me.Bar1.AccessibleRole = System.Windows.Forms.AccessibleRole.ToolBar
        Me.Bar1.AutoSyncBarCaption = True
        Me.Bar1.Controls.Add(Me.PanelDockContainer1)
        Me.Bar1.Font = New System.Drawing.Font("Verdana", 8.0!)
        Me.Bar1.GrabHandleStyle = DevComponents.DotNetBar.eGrabHandleStyle.Caption
        Me.Bar1.Items.AddRange(New DevComponents.DotNetBar.BaseItem() {Me.DockContainerItem3})
        Me.Bar1.LayoutType = DevComponents.DotNetBar.eLayoutType.DockContainer
        Me.Bar1.Location = New System.Drawing.Point(259, 0)
        Me.Bar1.Name = "Bar1"
        Me.Bar1.Size = New System.Drawing.Size(718, 62)
        Me.Bar1.Stretch = True
        Me.Bar1.Style = DevComponents.DotNetBar.eDotNetBarStyle.Office2007
        Me.Bar1.TabIndex = 1
        Me.Bar1.TabStop = False
        Me.Bar1.Text = "DockContainerItem3"
        '
        'PanelDockContainer1
        '
        Me.PanelDockContainer1.ColorSchemeStyle = DevComponents.DotNetBar.eDotNetBarStyle.Office2007
        Me.PanelDockContainer1.Controls.Add(Me.ClassLbl)
        Me.PanelDockContainer1.Controls.Add(Me.ClassCb)
        Me.PanelDockContainer1.Controls.Add(Me.FindBtn)
        Me.PanelDockContainer1.Controls.Add(Me.ModelLbl)
        Me.PanelDockContainer1.Controls.Add(Me.ModelLvCb)
        Me.PanelDockContainer1.Location = New System.Drawing.Point(3, 23)
        Me.PanelDockContainer1.Name = "PanelDockContainer1"
        Me.PanelDockContainer1.Size = New System.Drawing.Size(712, 36)
        Me.PanelDockContainer1.Style.Alignment = System.Drawing.StringAlignment.Center
        Me.PanelDockContainer1.Style.BackColor1.ColorSchemePart = DevComponents.DotNetBar.eColorSchemePart.BarBackground
        Me.PanelDockContainer1.Style.BackColor2.ColorSchemePart = DevComponents.DotNetBar.eColorSchemePart.BarBackground2
        Me.PanelDockContainer1.Style.BorderColor.ColorSchemePart = DevComponents.DotNetBar.eColorSchemePart.BarDockedBorder
        Me.PanelDockContainer1.Style.ForeColor.ColorSchemePart = DevComponents.DotNetBar.eColorSchemePart.ItemText
        Me.PanelDockContainer1.Style.GradientAngle = 90
        Me.PanelDockContainer1.TabIndex = 2
        '
        'ClassLbl
        '
        Me.ClassLbl.AutoSize = True
        Me.ClassLbl.BackColor = System.Drawing.Color.Transparent
        '
        '
        '
        Me.ClassLbl.BackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square
        Me.ClassLbl.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ClassLbl.Location = New System.Drawing.Point(10, 8)
        Me.ClassLbl.Name = "ClassLbl"
        Me.ClassLbl.Size = New System.Drawing.Size(56, 17)
        Me.ClassLbl.TabIndex = 8
        Me.ClassLbl.Text = "모델번호"
        Me.ClassLbl.TextAlignment = System.Drawing.StringAlignment.Center
        '
        'ClassCb
        '
        Me.ClassCb.DisplayMember = "class_id"
        Me.ClassCb.DrawMode = System.Windows.Forms.DrawMode.OwnerDrawFixed
        Me.ClassCb.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ClassCb.FormattingEnabled = True
        Me.ClassCb.ItemHeight = 15
        Me.ClassCb.Location = New System.Drawing.Point(68, 6)
        Me.ClassCb.Name = "ClassCb"
        Me.ClassCb.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.ClassCb.Size = New System.Drawing.Size(234, 21)
        Me.ClassCb.TabIndex = 7
        Me.ClassCb.ValueMember = "class_id"
        '
        'FindBtn
        '
        Me.FindBtn.AccessibleRole = System.Windows.Forms.AccessibleRole.PushButton
        Me.FindBtn.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink
        Me.FindBtn.ColorTable = DevComponents.DotNetBar.eButtonColor.OrangeWithBackground
        Me.FindBtn.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FindBtn.Location = New System.Drawing.Point(494, 6)
        Me.FindBtn.Name = "FindBtn"
        Me.FindBtn.Size = New System.Drawing.Size(66, 21)
        Me.FindBtn.TabIndex = 4
        Me.FindBtn.Text = "&FIND"
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
        Me.ModelLbl.Location = New System.Drawing.Point(310, 8)
        Me.ModelLbl.Name = "ModelLbl"
        Me.ModelLbl.Size = New System.Drawing.Size(56, 17)
        Me.ModelLbl.TabIndex = 0
        Me.ModelLbl.Text = "모델레벨"
        Me.ModelLbl.TextAlignment = System.Drawing.StringAlignment.Center
        '
        'ModelLvCb
        '
        Me.ModelLvCb.DisplayMember = "Text"
        Me.ModelLvCb.DrawMode = System.Windows.Forms.DrawMode.OwnerDrawFixed
        Me.ModelLvCb.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ModelLvCb.FormattingEnabled = True
        Me.ModelLvCb.ItemHeight = 15
        Me.ModelLvCb.Location = New System.Drawing.Point(369, 6)
        Me.ModelLvCb.Name = "ModelLvCb"
        Me.ModelLvCb.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.ModelLvCb.Size = New System.Drawing.Size(78, 21)
        Me.ModelLvCb.TabIndex = 2
        '
        'DockContainerItem3
        '
        Me.DockContainerItem3.Control = Me.PanelDockContainer1
        Me.DockContainerItem3.Name = "DockContainerItem3"
        Me.DockContainerItem3.Text = "DockContainerItem3"
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
        Me.TopDock.Size = New System.Drawing.Size(256, 62)
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
        Me.MenuPanel.Size = New System.Drawing.Size(250, 36)
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
        Me.CtlBar.Items.AddRange(New DevComponents.DotNetBar.BaseItem() {Me.NewBtn, Me.SaveBtn, Me.DelBtn, Me.PrtBtn, Me.Excel})
        Me.CtlBar.Location = New System.Drawing.Point(0, 0)
        Me.CtlBar.Name = "CtlBar"
        Me.CtlBar.Size = New System.Drawing.Size(250, 41)
        Me.CtlBar.Stretch = True
        Me.CtlBar.Style = DevComponents.DotNetBar.eDotNetBarStyle.Office2007
        Me.CtlBar.TabIndex = 1
        Me.CtlBar.TabStop = False
        Me.CtlBar.Text = "Bar2"
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
        Me.DockContainerItem1.Control = Me.MenuPanel
        Me.DockContainerItem1.Name = "DockContainerItem1"
        Me.DockContainerItem1.Text = "DockContainerItem1"
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
        'DockContainerItem2
        '
        Me.DockContainerItem2.Name = "DockContainerItem2"
        Me.DockContainerItem2.Text = "DockContainerItem2"
        '
        'XlsBtn1
        '
        Me.XlsBtn1.AccessibleRole = System.Windows.Forms.AccessibleRole.PushButton
        Me.XlsBtn1.ColorTable = DevComponents.DotNetBar.eButtonColor.OrangeWithBackground
        Me.XlsBtn1.Location = New System.Drawing.Point(351, 382)
        Me.XlsBtn1.Name = "XlsBtn1"
        Me.XlsBtn1.Size = New System.Drawing.Size(66, 21)
        Me.XlsBtn1.TabIndex = 23
        Me.XlsBtn1.Text = "E&Xcel"
        '
        'PrtBtn1
        '
        Me.PrtBtn1.AccessibleRole = System.Windows.Forms.AccessibleRole.PushButton
        Me.PrtBtn1.ColorTable = DevComponents.DotNetBar.eButtonColor.OrangeWithBackground
        Me.PrtBtn1.Location = New System.Drawing.Point(351, 355)
        Me.PrtBtn1.Name = "PrtBtn1"
        Me.PrtBtn1.Size = New System.Drawing.Size(66, 21)
        Me.PrtBtn1.TabIndex = 22
        Me.PrtBtn1.Text = "&Print"
        '
        'DelBtn1
        '
        Me.DelBtn1.AccessibleRole = System.Windows.Forms.AccessibleRole.PushButton
        Me.DelBtn1.ColorTable = DevComponents.DotNetBar.eButtonColor.OrangeWithBackground
        Me.DelBtn1.Location = New System.Drawing.Point(351, 328)
        Me.DelBtn1.Name = "DelBtn1"
        Me.DelBtn1.Size = New System.Drawing.Size(66, 21)
        Me.DelBtn1.TabIndex = 21
        Me.DelBtn1.Text = "&Delete"
        '
        'SaveBtn1
        '
        Me.SaveBtn1.AccessibleRole = System.Windows.Forms.AccessibleRole.PushButton
        Me.SaveBtn1.ColorTable = DevComponents.DotNetBar.eButtonColor.OrangeWithBackground
        Me.SaveBtn1.Location = New System.Drawing.Point(351, 301)
        Me.SaveBtn1.Name = "SaveBtn1"
        Me.SaveBtn1.Size = New System.Drawing.Size(66, 21)
        Me.SaveBtn1.TabIndex = 20
        Me.SaveBtn1.Text = "&Save"
        '
        'NewBtn1
        '
        Me.NewBtn1.AccessibleRole = System.Windows.Forms.AccessibleRole.PushButton
        Me.NewBtn1.ColorTable = DevComponents.DotNetBar.eButtonColor.OrangeWithBackground
        Me.NewBtn1.Location = New System.Drawing.Point(351, 274)
        Me.NewBtn1.Name = "NewBtn1"
        Me.NewBtn1.Size = New System.Drawing.Size(66, 21)
        Me.NewBtn1.TabIndex = 19
        Me.NewBtn1.Text = "&New"
        '
        'FrmModelLv
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(977, 676)
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
        Me.Name = "FrmModelLv"
        Me.Text = "FrmModelLv"
        Me.DockSite9.ResumeLayout(False)
        CType(Me.MainContets, System.ComponentModel.ISupportInitialize).EndInit()
        Me.MainContets.ResumeLayout(False)
        CType(Me.ContextMenuBar1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.FpSpread1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.FpSpread1_Sheet1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.DockSite3.ResumeLayout(False)
        CType(Me.Bar1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Bar1.ResumeLayout(False)
        Me.PanelDockContainer1.ResumeLayout(False)
        Me.PanelDockContainer1.PerformLayout()
        CType(Me.TopDock, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TopDock.ResumeLayout(False)
        Me.MenuPanel.ResumeLayout(False)
        CType(Me.CtlBar, System.ComponentModel.ISupportInitialize).EndInit()
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
    Friend WithEvents NewBtn As DevComponents.DotNetBar.ButtonItem
    Friend WithEvents SaveBtn As DevComponents.DotNetBar.ButtonItem
    Friend WithEvents DelBtn As DevComponents.DotNetBar.ButtonItem
    Friend WithEvents PrtBtn As DevComponents.DotNetBar.ButtonItem
    Friend WithEvents TopDock As DevComponents.DotNetBar.Bar
    Friend WithEvents MenuPanel As DevComponents.DotNetBar.PanelDockContainer
    Friend WithEvents CtlBar As DevComponents.DotNetBar.Bar
    Friend WithEvents DockContainerItem1 As DevComponents.DotNetBar.DockContainerItem
    Friend WithEvents DockSite9 As DevComponents.DotNetBar.DockSite
    Friend WithEvents MainContets As DevComponents.DotNetBar.Bar
    Friend WithEvents DockContainerItem2 As DevComponents.DotNetBar.DockContainerItem
    Friend WithEvents ModelLbl As DevComponents.DotNetBar.LabelX
    Friend WithEvents ModelLvCb As DevComponents.DotNetBar.Controls.ComboBoxEx
    Friend WithEvents FindBtn As DevComponents.DotNetBar.ButtonX
    Friend WithEvents ContextMenuBar1 As DevComponents.DotNetBar.ContextMenuBar
    Friend WithEvents CtxSp As DevComponents.DotNetBar.ButtonItem
    Friend WithEvents bNew As DevComponents.DotNetBar.ButtonItem
    Friend WithEvents bSave As DevComponents.DotNetBar.ButtonItem
    Friend WithEvents bDel As DevComponents.DotNetBar.ButtonItem
    Friend WithEvents bPrint As DevComponents.DotNetBar.ButtonItem
    Friend WithEvents Excel As DevComponents.DotNetBar.ButtonItem
    Friend WithEvents bExcel As DevComponents.DotNetBar.ButtonItem
    Friend WithEvents SaveFileDialog1 As System.Windows.Forms.SaveFileDialog
    Friend WithEvents FpSpread1 As FarPoint.Win.Spread.FpSpread
    Friend WithEvents FpSpread1_Sheet1 As FarPoint.Win.Spread.SheetView
    Friend WithEvents Bar1 As DevComponents.DotNetBar.Bar
    Friend WithEvents PanelDockContainer1 As DevComponents.DotNetBar.PanelDockContainer
    Friend WithEvents DockContainerItem3 As DevComponents.DotNetBar.DockContainerItem
    Friend WithEvents XlsBtn1 As DevComponents.DotNetBar.ButtonX
    Friend WithEvents PrtBtn1 As DevComponents.DotNetBar.ButtonX
    Friend WithEvents DelBtn1 As DevComponents.DotNetBar.ButtonX
    Friend WithEvents SaveBtn1 As DevComponents.DotNetBar.ButtonX
    Friend WithEvents NewBtn1 As DevComponents.DotNetBar.ButtonX
    Friend WithEvents ClassLbl As DevComponents.DotNetBar.LabelX
    Friend WithEvents ClassCb As DevComponents.DotNetBar.Controls.ComboBoxEx
End Class
