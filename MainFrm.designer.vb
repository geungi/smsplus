<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class MainFrm
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
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(MainFrm))
        Me.DotNetBarManager1 = New DevComponents.DotNetBar.DotNetBarManager(Me.components)
        Me.BottomDock = New DevComponents.DotNetBar.DockSite()
        Me.StatusBar1 = New DevComponents.DotNetBar.Bar()
        Me.LabelItem1 = New DevComponents.DotNetBar.LabelItem()
        Me.ComboBoxItem1 = New DevComponents.DotNetBar.ComboBoxItem()
        Me.ProgressBarItem1 = New DevComponents.DotNetBar.ProgressBarItem()
        Me.LabelItem2 = New DevComponents.DotNetBar.LabelItem()
        Me.DockSite1 = New DevComponents.DotNetBar.DockSite()
        Me.TabFrm = New DevComponents.DotNetBar.Bar()
        Me.RibbonBar1 = New DevComponents.DotNetBar.RibbonBar()
        Me.ContextMenuBar2 = New DevComponents.DotNetBar.ContextMenuBar()
        Me.ButtonItem3 = New DevComponents.DotNetBar.ButtonItem()
        Me.ButtonItem4 = New DevComponents.DotNetBar.ButtonItem()
        Me.ContextMenuBar1 = New DevComponents.DotNetBar.ContextMenuBar()
        Me.ButtonItem1 = New DevComponents.DotNetBar.ButtonItem()
        Me.ButtonItem2 = New DevComponents.DotNetBar.ButtonItem()
        Me.MainMenu = New DevComponents.DotNetBar.Bar()
        Me.PanelDockContainer1 = New DevComponents.DotNetBar.PanelDockContainer()
        Me.MenuList = New DevComponents.DotNetBar.SideBar()
        Me.DockContainerItem2 = New DevComponents.DotNetBar.DockContainerItem()
        Me.ImageList2 = New System.Windows.Forms.ImageList(Me.components)
        Me.LeftDock = New DevComponents.DotNetBar.DockSite()
        Me.RightDock = New DevComponents.DotNetBar.DockSite()
        Me.ToolBarBottomDock = New DevComponents.DotNetBar.DockSite()
        Me.ToolBarLeftDock = New DevComponents.DotNetBar.DockSite()
        Me.ToolBarRightDock = New DevComponents.DotNetBar.DockSite()
        Me.ToolBarTopDock = New DevComponents.DotNetBar.DockSite()
        Me.TopDock = New DevComponents.DotNetBar.DockSite()
        Me.SliderItem1 = New DevComponents.DotNetBar.SliderItem()
        Me.DockContainerItem4 = New DevComponents.DotNetBar.DockContainerItem()
        Me.ComboItem1 = New DevComponents.Editors.ComboItem()
        Me.DockContainerItem1 = New DevComponents.DotNetBar.DockContainerItem()
        Me.ImageList1 = New System.Windows.Forms.ImageList(Me.components)
        Me.DockContainerItem3 = New DevComponents.DotNetBar.DockContainerItem()
        Me.DockContainerItem5 = New DevComponents.DotNetBar.DockContainerItem()
        Me.BottomDock.SuspendLayout()
        CType(Me.StatusBar1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.DockSite1.SuspendLayout()
        CType(Me.TabFrm, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TabFrm.SuspendLayout()
        CType(Me.ContextMenuBar2, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.ContextMenuBar1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.MainMenu, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.MainMenu.SuspendLayout()
        Me.PanelDockContainer1.SuspendLayout()
        Me.LeftDock.SuspendLayout()
        Me.SuspendLayout()
        '
        'DotNetBarManager1
        '
        Me.DotNetBarManager1.AutoDispatchShortcuts.Add(DevComponents.DotNetBar.eShortcut.Ins)
        Me.DotNetBarManager1.AutoDispatchShortcuts.Add(DevComponents.DotNetBar.eShortcut.Del)
        Me.DotNetBarManager1.AutoDispatchShortcuts.Add(DevComponents.DotNetBar.eShortcut.F1)
        Me.DotNetBarManager1.AutoDispatchShortcuts.Add(DevComponents.DotNetBar.eShortcut.CtrlA)
        Me.DotNetBarManager1.AutoDispatchShortcuts.Add(DevComponents.DotNetBar.eShortcut.CtrlB)
        Me.DotNetBarManager1.AutoDispatchShortcuts.Add(DevComponents.DotNetBar.eShortcut.CtrlC)
        Me.DotNetBarManager1.AutoDispatchShortcuts.Add(DevComponents.DotNetBar.eShortcut.CtrlV)
        Me.DotNetBarManager1.AutoDispatchShortcuts.Add(DevComponents.DotNetBar.eShortcut.CtrlX)
        Me.DotNetBarManager1.AutoDispatchShortcuts.Add(DevComponents.DotNetBar.eShortcut.CtrlY)
        Me.DotNetBarManager1.AutoDispatchShortcuts.Add(DevComponents.DotNetBar.eShortcut.CtrlZ)
        Me.DotNetBarManager1.BottomDockSite = Me.BottomDock
        Me.DotNetBarManager1.DispatchShortcuts = True
        Me.DotNetBarManager1.EnableFullSizeDock = False
        Me.DotNetBarManager1.FillDockSite = Me.DockSite1
        Me.DotNetBarManager1.LeftDockSite = Me.LeftDock
        Me.DotNetBarManager1.LicenseKey = "F962CEC7-CD8F-4911-A9E9-CAB39962FC1F"
        Me.DotNetBarManager1.ParentForm = Me
        Me.DotNetBarManager1.RightDockSite = Me.RightDock
        Me.DotNetBarManager1.ShowCustomizeContextMenu = False
        Me.DotNetBarManager1.Style = DevComponents.DotNetBar.eDotNetBarStyle.Metro
        Me.DotNetBarManager1.ToolbarBottomDockSite = Me.ToolBarBottomDock
        Me.DotNetBarManager1.ToolbarLeftDockSite = Me.ToolBarLeftDock
        Me.DotNetBarManager1.ToolbarRightDockSite = Me.ToolBarRightDock
        Me.DotNetBarManager1.ToolbarTopDockSite = Me.ToolBarTopDock
        Me.DotNetBarManager1.TopDockSite = Me.TopDock
        Me.DotNetBarManager1.UseCustomCustomizeDialog = True
        '
        'BottomDock
        '
        Me.BottomDock.AccessibleRole = System.Windows.Forms.AccessibleRole.Window
        Me.BottomDock.Controls.Add(Me.StatusBar1)
        Me.BottomDock.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.BottomDock.DocumentDockContainer = New DevComponents.DotNetBar.DocumentDockContainer(New DevComponents.DotNetBar.DocumentBaseContainer() {CType(New DevComponents.DotNetBar.DocumentBarContainer(Me.StatusBar1, 1008, 22), DevComponents.DotNetBar.DocumentBaseContainer)}, DevComponents.DotNetBar.eOrientation.Vertical)
        Me.BottomDock.Location = New System.Drawing.Point(0, 707)
        Me.BottomDock.Name = "BottomDock"
        Me.BottomDock.Size = New System.Drawing.Size(1008, 25)
        Me.BottomDock.TabIndex = 4
        Me.BottomDock.TabStop = False
        '
        'StatusBar1
        '
        Me.StatusBar1.AccessibleDescription = "StatusBar1 (StatusBar1)"
        Me.StatusBar1.AccessibleName = "StatusBar1"
        Me.StatusBar1.AccessibleRole = System.Windows.Forms.AccessibleRole.ToolBar
        Me.StatusBar1.AutoSyncBarCaption = True
        Me.StatusBar1.CloseSingleTab = True
        Me.ContextMenuBar1.SetContextMenuEx(Me.StatusBar1, Me.ButtonItem1)
        Me.StatusBar1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.StatusBar1.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.StatusBar1.Items.AddRange(New DevComponents.DotNetBar.BaseItem() {Me.LabelItem1, Me.ComboBoxItem1, Me.ProgressBarItem1, Me.LabelItem2})
        Me.StatusBar1.Location = New System.Drawing.Point(0, 0)
        Me.StatusBar1.Name = "StatusBar1"
        Me.StatusBar1.Size = New System.Drawing.Size(1008, 25)
        Me.StatusBar1.Stretch = True
        Me.StatusBar1.Style = DevComponents.DotNetBar.eDotNetBarStyle.Metro
        Me.StatusBar1.TabIndex = 0
        Me.StatusBar1.TabStop = False
        Me.StatusBar1.Text = "StatusBar1"
        '
        'LabelItem1
        '
        Me.LabelItem1.BorderSide = DevComponents.DotNetBar.eBorderSide.None
        Me.LabelItem1.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LabelItem1.ForeColor = System.Drawing.Color.MidnightBlue
        Me.LabelItem1.Name = "LabelItem1"
        Me.LabelItem1.TextAlignment = System.Drawing.StringAlignment.Center
        '
        'ComboBoxItem1
        '
        Me.ComboBoxItem1.ComboWidth = 200
        Me.ComboBoxItem1.DropDownHeight = 106
        Me.ComboBoxItem1.DropDownWidth = 242
        Me.ComboBoxItem1.Name = "ComboBoxItem1"
        Me.ComboBoxItem1.Visible = False
        '
        'ProgressBarItem1
        '
        '
        '
        '
        Me.ProgressBarItem1.BackStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square
        Me.ProgressBarItem1.ChunkGradientAngle = 0!
        Me.ProgressBarItem1.ItemAlignment = DevComponents.DotNetBar.eItemAlignment.Far
        Me.ProgressBarItem1.MenuVisibility = DevComponents.DotNetBar.eMenuVisibility.VisibleAlways
        Me.ProgressBarItem1.Name = "ProgressBarItem1"
        Me.ProgressBarItem1.RecentlyUsed = False
        Me.ProgressBarItem1.TextVisible = True
        Me.ProgressBarItem1.Width = 200
        '
        'LabelItem2
        '
        Me.LabelItem2.Name = "LabelItem2"
        '
        'DockSite1
        '
        Me.DockSite1.AccessibleRole = System.Windows.Forms.AccessibleRole.Window
        Me.DockSite1.Controls.Add(Me.TabFrm)
        Me.DockSite1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.DockSite1.DocumentDockContainer = New DevComponents.DotNetBar.DocumentDockContainer(New DevComponents.DotNetBar.DocumentBaseContainer() {CType(New DevComponents.DotNetBar.DocumentBarContainer(Me.TabFrm, 816, 707), DevComponents.DotNetBar.DocumentBaseContainer)}, DevComponents.DotNetBar.eOrientation.Horizontal)
        Me.DockSite1.Location = New System.Drawing.Point(192, 0)
        Me.DockSite1.Name = "DockSite1"
        Me.DockSite1.Size = New System.Drawing.Size(816, 707)
        Me.DockSite1.TabIndex = 9
        Me.DockSite1.TabStop = False
        '
        'TabFrm
        '
        Me.TabFrm.AccessibleDescription = "DotNetBar Bar (TabFrm)"
        Me.TabFrm.AccessibleName = "DotNetBar Bar"
        Me.TabFrm.AccessibleRole = System.Windows.Forms.AccessibleRole.MenuBar
        Me.TabFrm.AlwaysDisplayDockTab = True
        Me.TabFrm.AutoHide = True
        Me.TabFrm.AutoHideTabTextAlwaysVisible = True
        Me.TabFrm.AutoSyncBarCaption = True
        Me.TabFrm.CanCustomize = False
        Me.TabFrm.CanDockDocument = True
        Me.TabFrm.CanHide = True
        Me.TabFrm.CanUndock = False
        Me.TabFrm.Controls.Add(Me.RibbonBar1)
        Me.TabFrm.Controls.Add(Me.ContextMenuBar2)
        Me.TabFrm.Controls.Add(Me.ContextMenuBar1)
        Me.TabFrm.DisplayMoreItemsOnMenu = True
        Me.TabFrm.DockTabAlignment = DevComponents.DotNetBar.eTabStripAlignment.Top
        Me.TabFrm.DockTabCloseButtonVisible = True
        Me.TabFrm.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TabFrm.Images = Me.ImageList2
        Me.TabFrm.LayoutType = DevComponents.DotNetBar.eLayoutType.DockContainer
        Me.TabFrm.Location = New System.Drawing.Point(0, 0)
        Me.TabFrm.MenuBar = True
        Me.TabFrm.Name = "TabFrm"
        Me.TabFrm.Size = New System.Drawing.Size(816, 707)
        Me.TabFrm.Stretch = True
        Me.TabFrm.Style = DevComponents.DotNetBar.eDotNetBarStyle.Metro
        Me.TabFrm.TabIndex = 0
        Me.TabFrm.TabNavigation = True
        Me.TabFrm.TabStop = False
        Me.TabFrm.Text = "MAIN"
        '
        'RibbonBar1
        '
        Me.RibbonBar1.AutoOverflowEnabled = True
        '
        '
        '
        Me.RibbonBar1.BackgroundMouseOverStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square
        '
        '
        '
        Me.RibbonBar1.BackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square
        Me.RibbonBar1.ContainerControlProcessDialogKey = True
        Me.ContextMenuBar2.SetContextMenuEx(Me.RibbonBar1, Me.ButtonItem3)
        Me.RibbonBar1.Dock = System.Windows.Forms.DockStyle.Top
        Me.RibbonBar1.ItemSpacing = 10
        Me.RibbonBar1.LicenseKey = "F962CEC7-CD8F-4911-A9E9-CAB39962FC1F"
        Me.RibbonBar1.Location = New System.Drawing.Point(0, 0)
        Me.RibbonBar1.Name = "RibbonBar1"
        Me.RibbonBar1.Size = New System.Drawing.Size(816, 28)
        Me.RibbonBar1.Style = DevComponents.DotNetBar.eDotNetBarStyle.Windows7
        Me.RibbonBar1.TabIndex = 23
        Me.RibbonBar1.Text = "RibbonBar1"
        '
        '
        '
        Me.RibbonBar1.TitleStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square
        '
        '
        '
        Me.RibbonBar1.TitleStyleMouseOver.CornerType = DevComponents.DotNetBar.eCornerType.Square
        Me.RibbonBar1.TitleVisible = False
        '
        'ContextMenuBar2
        '
        Me.ContextMenuBar2.Font = New System.Drawing.Font("Malgun Gothic", 9.0!)
        Me.ContextMenuBar2.Items.AddRange(New DevComponents.DotNetBar.BaseItem() {Me.ButtonItem3})
        Me.ContextMenuBar2.Location = New System.Drawing.Point(252, 236)
        Me.ContextMenuBar2.Name = "ContextMenuBar2"
        Me.ContextMenuBar2.Size = New System.Drawing.Size(108, 27)
        Me.ContextMenuBar2.Stretch = True
        Me.ContextMenuBar2.Style = DevComponents.DotNetBar.eDotNetBarStyle.Office2003
        Me.ContextMenuBar2.TabIndex = 19
        Me.ContextMenuBar2.TabStop = False
        Me.ContextMenuBar2.Text = "ContextMenuBar2"
        '
        'ButtonItem3
        '
        Me.ButtonItem3.AutoExpandOnClick = True
        Me.ButtonItem3.Name = "ButtonItem3"
        Me.ButtonItem3.SubItems.AddRange(New DevComponents.DotNetBar.BaseItem() {Me.ButtonItem4})
        Me.ButtonItem3.Text = "ButtonItem3"
        '
        'ButtonItem4
        '
        Me.ButtonItem4.Name = "ButtonItem4"
        Me.ButtonItem4.Text = "Show Side Menu"
        '
        'ContextMenuBar1
        '
        Me.ContextMenuBar1.Font = New System.Drawing.Font("Malgun Gothic", 9.0!)
        Me.ContextMenuBar1.Items.AddRange(New DevComponents.DotNetBar.BaseItem() {Me.ButtonItem1})
        Me.ContextMenuBar1.Location = New System.Drawing.Point(79, 124)
        Me.ContextMenuBar1.Name = "ContextMenuBar1"
        Me.ContextMenuBar1.Size = New System.Drawing.Size(102, 27)
        Me.ContextMenuBar1.Stretch = True
        Me.ContextMenuBar1.Style = DevComponents.DotNetBar.eDotNetBarStyle.Office2003
        Me.ContextMenuBar1.TabIndex = 17
        Me.ContextMenuBar1.TabStop = False
        Me.ContextMenuBar1.Text = "ContextMenuBar1"
        '
        'ButtonItem1
        '
        Me.ButtonItem1.AutoExpandOnClick = True
        Me.ButtonItem1.Name = "ButtonItem1"
        Me.ButtonItem1.SubItems.AddRange(New DevComponents.DotNetBar.BaseItem() {Me.ButtonItem2})
        Me.ButtonItem1.Text = "ButtonItem1"
        '
        'ButtonItem2
        '
        Me.ButtonItem2.Name = "ButtonItem2"
        Me.ButtonItem2.Text = "Show Top Menu"
        '
        'MainMenu
        '
        Me.MainMenu.AccessibleDescription = "DotNetBar Bar (MainMenu)"
        Me.MainMenu.AccessibleName = "DotNetBar Bar"
        Me.MainMenu.AccessibleRole = System.Windows.Forms.AccessibleRole.ToolBar
        Me.MainMenu.AutoSyncBarCaption = True
        Me.MainMenu.CloseSingleTab = True
        Me.ContextMenuBar1.SetContextMenuEx(Me.MainMenu, Me.ButtonItem1)
        Me.MainMenu.Controls.Add(Me.PanelDockContainer1)
        Me.MainMenu.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.MainMenu.GrabHandleStyle = DevComponents.DotNetBar.eGrabHandleStyle.Caption
        Me.MainMenu.Items.AddRange(New DevComponents.DotNetBar.BaseItem() {Me.DockContainerItem2})
        Me.MainMenu.LayoutType = DevComponents.DotNetBar.eLayoutType.DockContainer
        Me.MainMenu.Location = New System.Drawing.Point(0, 0)
        Me.MainMenu.Name = "MainMenu"
        Me.MainMenu.Size = New System.Drawing.Size(189, 707)
        Me.MainMenu.Stretch = True
        Me.MainMenu.Style = DevComponents.DotNetBar.eDotNetBarStyle.Metro
        Me.MainMenu.TabIndex = 1
        Me.MainMenu.TabStop = False
        Me.MainMenu.Text = "DockContainerItem2"
        '
        'PanelDockContainer1
        '
        Me.PanelDockContainer1.ColorSchemeStyle = DevComponents.DotNetBar.eDotNetBarStyle.Metro
        Me.PanelDockContainer1.Controls.Add(Me.MenuList)
        Me.PanelDockContainer1.Location = New System.Drawing.Point(3, 23)
        Me.PanelDockContainer1.Name = "PanelDockContainer1"
        Me.PanelDockContainer1.Size = New System.Drawing.Size(183, 681)
        Me.PanelDockContainer1.Style.Alignment = System.Drawing.StringAlignment.Center
        Me.PanelDockContainer1.Style.BackColor1.ColorSchemePart = DevComponents.DotNetBar.eColorSchemePart.BarBackground
        Me.PanelDockContainer1.Style.BackColor2.ColorSchemePart = DevComponents.DotNetBar.eColorSchemePart.BarBackground2
        Me.PanelDockContainer1.Style.BorderColor.ColorSchemePart = DevComponents.DotNetBar.eColorSchemePart.BarDockedBorder
        Me.PanelDockContainer1.Style.ForeColor.ColorSchemePart = DevComponents.DotNetBar.eColorSchemePart.ItemText
        Me.PanelDockContainer1.Style.GradientAngle = 90
        Me.PanelDockContainer1.TabIndex = 0
        '
        'MenuList
        '
        Me.MenuList.AccessibleRole = System.Windows.Forms.AccessibleRole.ToolBar
        Me.MenuList.BackColor = System.Drawing.Color.FromArgb(CType(CType(128, Byte), Integer), CType(CType(128, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.MenuList.BorderStyle = DevComponents.DotNetBar.eBorderType.None
        Me.MenuList.Dock = System.Windows.Forms.DockStyle.Fill
        Me.MenuList.ExpandedPanel = Nothing
        Me.MenuList.Font = New System.Drawing.Font("Malgun Gothic", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.MenuList.Location = New System.Drawing.Point(0, 0)
        Me.MenuList.Name = "MenuList"
        Me.MenuList.Size = New System.Drawing.Size(183, 681)
        Me.MenuList.Style = DevComponents.DotNetBar.eDotNetBarStyle.Office2007
        Me.MenuList.TabIndex = 2
        Me.MenuList.Text = "Menu List"
        '
        'DockContainerItem2
        '
        Me.DockContainerItem2.Control = Me.PanelDockContainer1
        Me.DockContainerItem2.Name = "DockContainerItem2"
        Me.DockContainerItem2.Text = "DockContainerItem2"
        '
        'ImageList2
        '
        Me.ImageList2.ImageStream = CType(resources.GetObject("ImageList2.ImageStream"), System.Windows.Forms.ImageListStreamer)
        Me.ImageList2.TransparentColor = System.Drawing.Color.Transparent
        Me.ImageList2.Images.SetKeyName(0, "document-new.png")
        Me.ImageList2.Images.SetKeyName(1, "document-open.png")
        Me.ImageList2.Images.SetKeyName(2, "filesave.png")
        Me.ImageList2.Images.SetKeyName(3, "edit-delete.png")
        Me.ImageList2.Images.SetKeyName(4, "fileprint.png")
        Me.ImageList2.Images.SetKeyName(5, "gnome-mime-application-vnd.ms-excel.png")
        '
        'LeftDock
        '
        Me.LeftDock.AccessibleRole = System.Windows.Forms.AccessibleRole.Window
        Me.LeftDock.Controls.Add(Me.MainMenu)
        Me.LeftDock.Dock = System.Windows.Forms.DockStyle.Left
        Me.LeftDock.DocumentDockContainer = New DevComponents.DotNetBar.DocumentDockContainer(New DevComponents.DotNetBar.DocumentBaseContainer() {CType(New DevComponents.DotNetBar.DocumentBarContainer(Me.MainMenu, 189, 707), DevComponents.DotNetBar.DocumentBaseContainer)}, DevComponents.DotNetBar.eOrientation.Horizontal)
        Me.LeftDock.Location = New System.Drawing.Point(0, 0)
        Me.LeftDock.Name = "LeftDock"
        Me.LeftDock.Size = New System.Drawing.Size(192, 707)
        Me.LeftDock.TabIndex = 1
        Me.LeftDock.TabStop = False
        Me.LeftDock.Text = "LeftDock"
        '
        'RightDock
        '
        Me.RightDock.AccessibleRole = System.Windows.Forms.AccessibleRole.Window
        Me.RightDock.Dock = System.Windows.Forms.DockStyle.Right
        Me.RightDock.DocumentDockContainer = New DevComponents.DotNetBar.DocumentDockContainer()
        Me.RightDock.Location = New System.Drawing.Point(1008, 0)
        Me.RightDock.Name = "RightDock"
        Me.RightDock.Size = New System.Drawing.Size(0, 707)
        Me.RightDock.TabIndex = 2
        Me.RightDock.TabStop = False
        '
        'ToolBarBottomDock
        '
        Me.ToolBarBottomDock.AccessibleRole = System.Windows.Forms.AccessibleRole.Window
        Me.ToolBarBottomDock.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.ToolBarBottomDock.Location = New System.Drawing.Point(0, 732)
        Me.ToolBarBottomDock.Name = "ToolBarBottomDock"
        Me.ToolBarBottomDock.Size = New System.Drawing.Size(1008, 0)
        Me.ToolBarBottomDock.TabIndex = 8
        Me.ToolBarBottomDock.TabStop = False
        '
        'ToolBarLeftDock
        '
        Me.ToolBarLeftDock.AccessibleRole = System.Windows.Forms.AccessibleRole.Window
        Me.ToolBarLeftDock.Dock = System.Windows.Forms.DockStyle.Left
        Me.ToolBarLeftDock.Location = New System.Drawing.Point(0, 0)
        Me.ToolBarLeftDock.Name = "ToolBarLeftDock"
        Me.ToolBarLeftDock.Size = New System.Drawing.Size(0, 732)
        Me.ToolBarLeftDock.TabIndex = 5
        Me.ToolBarLeftDock.TabStop = False
        '
        'ToolBarRightDock
        '
        Me.ToolBarRightDock.AccessibleRole = System.Windows.Forms.AccessibleRole.Window
        Me.ToolBarRightDock.Dock = System.Windows.Forms.DockStyle.Right
        Me.ToolBarRightDock.Location = New System.Drawing.Point(1008, 0)
        Me.ToolBarRightDock.Name = "ToolBarRightDock"
        Me.ToolBarRightDock.Size = New System.Drawing.Size(0, 732)
        Me.ToolBarRightDock.TabIndex = 6
        Me.ToolBarRightDock.TabStop = False
        '
        'ToolBarTopDock
        '
        Me.ToolBarTopDock.AccessibleRole = System.Windows.Forms.AccessibleRole.Window
        Me.ToolBarTopDock.Dock = System.Windows.Forms.DockStyle.Top
        Me.ToolBarTopDock.Location = New System.Drawing.Point(0, 0)
        Me.ToolBarTopDock.Name = "ToolBarTopDock"
        Me.ToolBarTopDock.Size = New System.Drawing.Size(1008, 0)
        Me.ToolBarTopDock.TabIndex = 7
        Me.ToolBarTopDock.TabStop = False
        Me.ToolBarTopDock.Text = "ToolBarTopDock"
        '
        'TopDock
        '
        Me.TopDock.AccessibleRole = System.Windows.Forms.AccessibleRole.Window
        Me.TopDock.Dock = System.Windows.Forms.DockStyle.Top
        Me.TopDock.DocumentDockContainer = New DevComponents.DotNetBar.DocumentDockContainer()
        Me.TopDock.Location = New System.Drawing.Point(0, 0)
        Me.TopDock.Name = "TopDock"
        Me.TopDock.Size = New System.Drawing.Size(1008, 0)
        Me.TopDock.TabIndex = 3
        Me.TopDock.TabStop = False
        '
        'SliderItem1
        '
        Me.SliderItem1.ItemAlignment = DevComponents.DotNetBar.eItemAlignment.Far
        Me.SliderItem1.Name = "SliderItem1"
        Me.SliderItem1.Text = "SliderItem1"
        Me.SliderItem1.Value = 0
        Me.SliderItem1.Visible = False
        '
        'DockContainerItem4
        '
        Me.DockContainerItem4.Name = "DockContainerItem4"
        Me.DockContainerItem4.Text = "MAIN"
        '
        'ComboItem1
        '
        Me.ComboItem1.Text = "ComboItem1"
        '
        'DockContainerItem1
        '
        Me.DockContainerItem1.Name = "DockContainerItem1"
        Me.DockContainerItem1.Text = "Main Menu"
        '
        'ImageList1
        '
        Me.ImageList1.ImageStream = CType(resources.GetObject("ImageList1.ImageStream"), System.Windows.Forms.ImageListStreamer)
        Me.ImageList1.TransparentColor = System.Drawing.Color.Transparent
        Me.ImageList1.Images.SetKeyName(0, "filenew.png")
        Me.ImageList1.Images.SetKeyName(1, "fileopen.png")
        Me.ImageList1.Images.SetKeyName(2, "filesave.png")
        Me.ImageList1.Images.SetKeyName(3, "filequickprint.png")
        Me.ImageList1.Images.SetKeyName(4, "gtk-undo-ltr.png")
        Me.ImageList1.Images.SetKeyName(5, "gtk-redo-ltr.png")
        Me.ImageList1.Images.SetKeyName(6, "editcut.png")
        Me.ImageList1.Images.SetKeyName(7, "edit-copy.png")
        Me.ImageList1.Images.SetKeyName(8, "gtk-paste.png")
        Me.ImageList1.Images.SetKeyName(9, "edit-select-all.png")
        '
        'DockContainerItem3
        '
        Me.DockContainerItem3.Name = "DockContainerItem3"
        Me.DockContainerItem3.Text = "MAIN"
        '
        'DockContainerItem5
        '
        Me.DockContainerItem5.Name = "DockContainerItem5"
        Me.DockContainerItem5.Text = "DockContainerItem5"
        '
        'MainFrm
        '
        Me.BackColor = System.Drawing.Color.White
        Me.ClientSize = New System.Drawing.Size(1008, 732)
        Me.ContextMenuBar1.SetContextMenuEx(Me, Me.ButtonItem1)
        Me.Controls.Add(Me.DockSite1)
        Me.Controls.Add(Me.RightDock)
        Me.Controls.Add(Me.LeftDock)
        Me.Controls.Add(Me.TopDock)
        Me.Controls.Add(Me.BottomDock)
        Me.Controls.Add(Me.ToolBarLeftDock)
        Me.Controls.Add(Me.ToolBarRightDock)
        Me.Controls.Add(Me.ToolBarTopDock)
        Me.Controls.Add(Me.ToolBarBottomDock)
        Me.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.HelpButton = True
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.KeyPreview = True
        Me.Name = "MainFrm"
        Me.Text = "SCOTTII KOREA 통합관리시스템"
        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
        Me.BottomDock.ResumeLayout(False)
        CType(Me.StatusBar1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.DockSite1.ResumeLayout(False)
        CType(Me.TabFrm, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TabFrm.ResumeLayout(False)
        CType(Me.ContextMenuBar2, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.ContextMenuBar1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.MainMenu, System.ComponentModel.ISupportInitialize).EndInit()
        Me.MainMenu.ResumeLayout(False)
        Me.PanelDockContainer1.ResumeLayout(False)
        Me.LeftDock.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents DotNetBarManager1 As DevComponents.DotNetBar.DotNetBarManager
    Friend WithEvents BottomDock As DevComponents.DotNetBar.DockSite
    Friend WithEvents LeftDock As DevComponents.DotNetBar.DockSite
    Friend WithEvents RightDock As DevComponents.DotNetBar.DockSite
    Friend WithEvents TopDock As DevComponents.DotNetBar.DockSite
    Friend WithEvents ToolBarLeftDock As DevComponents.DotNetBar.DockSite
    Friend WithEvents ToolBarRightDock As DevComponents.DotNetBar.DockSite
    Friend WithEvents ToolBarTopDock As DevComponents.DotNetBar.DockSite
    Friend WithEvents ToolBarBottomDock As DevComponents.DotNetBar.DockSite
    Friend WithEvents MainMenu As DevComponents.DotNetBar.Bar
    Friend WithEvents DockSite1 As DevComponents.DotNetBar.DockSite
    Friend WithEvents TabFrm As DevComponents.DotNetBar.Bar
    Friend WithEvents ImageList1 As System.Windows.Forms.ImageList
    Friend WithEvents StatusBar1 As DevComponents.DotNetBar.Bar
    Friend WithEvents ImageList2 As System.Windows.Forms.ImageList
    Friend WithEvents ComboItem1 As DevComponents.Editors.ComboItem
    Friend WithEvents ProgressBarItem1 As DevComponents.DotNetBar.ProgressBarItem
    Friend WithEvents DockContainerItem1 As DevComponents.DotNetBar.DockContainerItem
    Friend WithEvents PanelDockContainer1 As DevComponents.DotNetBar.PanelDockContainer
    Friend WithEvents MenuList As DevComponents.DotNetBar.SideBar
    Friend WithEvents DockContainerItem2 As DevComponents.DotNetBar.DockContainerItem
    Friend WithEvents LabelItem1 As DevComponents.DotNetBar.LabelItem
    Friend WithEvents ComboBoxItem1 As DevComponents.DotNetBar.ComboBoxItem
    Friend WithEvents LabelItem2 As DevComponents.DotNetBar.LabelItem
    Friend WithEvents DockContainerItem4 As DevComponents.DotNetBar.DockContainerItem
    Friend WithEvents DockContainerItem3 As DevComponents.DotNetBar.DockContainerItem
    Friend WithEvents DockContainerItem5 As DevComponents.DotNetBar.DockContainerItem
    Friend WithEvents ContextMenuBar1 As DevComponents.DotNetBar.ContextMenuBar
    Friend WithEvents ButtonItem1 As DevComponents.DotNetBar.ButtonItem
    Friend WithEvents ButtonItem2 As DevComponents.DotNetBar.ButtonItem
    Friend WithEvents ContextMenuBar2 As DevComponents.DotNetBar.ContextMenuBar
    Friend WithEvents ButtonItem3 As DevComponents.DotNetBar.ButtonItem
    Friend WithEvents ButtonItem4 As DevComponents.DotNetBar.ButtonItem
    Friend WithEvents SliderItem1 As DevComponents.DotNetBar.SliderItem
    Friend WithEvents RibbonBar1 As DevComponents.DotNetBar.RibbonBar
End Class
