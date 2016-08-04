Imports DevComponents.DotNetBar
Imports System.Data.SqlClient


Public Class MainFrm

    Protected Conn As New SqlConnection()
    '    Protected MenuDs As New MenuDataSet()
    '   Protected NewMenuTadt As New MenuDataSetTableAdapters.NewmenuTapt
    Public login_yn As Boolean

    Public Sub New()
        MyBase.New()
        'This call is required by the Windows Form Designer.
        InitializeComponent()
 
    End Sub

    Private Sub MainFrm_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        Dim ipaddr As String
        Dim aa() As System.Net.IPAddress

        ipaddr = ""
        aa = System.Net.Dns.GetHostAddresses(My.Computer.Name)

        For Each ip As System.Net.IPAddress In aa
            ipaddr = ip.ToString
            If InStr(ipaddr, ".") > 0 Then
                Exit For
            End If
        Next

        If Site_id IsNot Nothing Then
            If Insert_Data("EXEC SP_COMMON_LOGSTATUS_C '" & Site_id & "','" & Emp_No & "','" & ipaddr & "','" & CStr(My.Application.Info.Version.Major) & "." & CStr(My.Application.Info.Version.Minor) & "." & CStr(My.Application.Info.Version.Build) & "." & CStr(My.Application.Info.Version.Revision) & "','" & Now & "'") = True Then
            End If
        End If

    End Sub

    Private Sub MainFrm_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing

        If MessageBox.Show("시스템을 종료하시겠습니까? ", "시스템 종료", MessageBoxButtons.YesNo) = Windows.Forms.DialogResult.No Then
            e.Cancel = True
        End If

    End Sub

    Private Sub MainFrm_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        SliderItem1.Value = 100

        Try

            LoginFrm.ShowDialog()
            If login_yn <> True Then
                Me.Close()
                Exit Sub
            End If

            i_conn.Open(db_conn(1))
            i_cmd.ActiveConnection = i_conn

            Me.Text = "SCOTTII KOREA 통합관리시스템"
            If Site_id <> "" And Emp_No <> "" Then
                If SVR <> "" And dbid <> "" And dbpw <> "" Then
                    'NewMenuTadt.Connection.ConnectionString = db_conn(5)
                End If

                Me.LabelItem1.Text = "[" & Site_id & "] " & Site_nm
                Me.DockContainerItem2.Text = "SCOTTII KOREA"

                If Query_Combo(ComboBoxItem1.ComboBoxEx, "sELECT '[' + EMP_NO + ']' + EMP_NM FROM TBL_EMPMASTER WHERE dept_cd = (select dept_cd from tbl_empmaster where emp_no = '" & Emp_No & "') and RETIRE_YN = 'N' ORDER BY EMP_NO") = True Then
                Else
                    MessageBox.Show("error")
                End If
                ComboBoxItem1.ComboBoxEx.Text = "[" & Emp_No & "]" & Emp_Nm





                Me.Text = Me.Text & "(Version " & CStr(My.Application.Info.Version.Major) & "." & CStr(My.Application.Info.Version.Minor) & "." & CStr(My.Application.Info.Version.Build) & "." & CStr(My.Application.Info.Version.Revision) & ")"
            End If

            RibbonBar1.Visible = False

            If ShowSideMenu() = False Then
                MessageBox.Show("ShowSideMenu Show error")
            End If

            LabelItem2.Text = "   Logined by [" & Emp_No & "] " & Emp_Nm
            MainMenu.Refresh()

            disp_main()

        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message, "ERROR")
        End Try

    End Sub

    Private Function ShowMenuBarM(ByVal MenuName As String) As Boolean '상단메뉴바에 메뉴표시
        Try
            Dim MN As ButtonItem

            MN = New ButtonItem(MenuName)
            MN.Text = MenuName + "(&" + Mid(MenuName, 1, 1) + ")"

        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message, "ERROR")
        End Try
    End Function

    Private Function ShowMenuBarS(ByVal MainMenu As String, ByVal SubName As String, ByVal imgidx As Integer, ByVal ShortKey As eShortcut) As Boolean '상단메뉴바에 메뉴표시
        Try
            Dim MM As New ButtonItem
            Dim SN As ButtonItem

            SN = New ButtonItem(SubName, SubName)
            SN.Text = "&" + SubName
            If imgidx < 99 Then
                SN.Image = ImageList1.Images(imgidx)
            End If
            If ShortKey <> eShortcut.None Then
                SN.Shortcuts.Add(ShortKey)
            End If

        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message, "ERROR")
        End Try
    End Function

    Private Function ShowToolBars(ByVal ToolMenu As String, ByVal ToolMenuNM As String, ByVal imgidx As Integer, ByVal ShortKey As eShortcut) As Boolean
        Try
            Dim TM As New ButtonItem
            TM = New ButtonItem(ToolMenu, ToolMenuNM)
            TM.Tooltip = ToolMenuNM
            If imgidx < 99 Then
                TM.Image = ImageList2.Images(imgidx)
            End If
            If ShortKey <> eShortcut.None Then
                TM.Shortcuts.Add(ShortKey)
            End If

        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message, "ERROR")
        End Try
    End Function

    Private Function ShowSideMenu() As Boolean
        'Try
        ShowSideMenu = False

        '좌측에 메뉴표시
        Dim MenuDate As New ADODB.Recordset
        Dim M1 As New SideBarPanelItem
        Dim M2 As New ButtonItem
        Dim M3 As New ButtonItem
        Dim BM1 As New ButtonItem
        Dim BM2 As New ButtonItem
        Dim BM3 As New ButtonItem
        Dim i As Integer

        MenuDate = Query_RS_ALL("select menu1, isnull(menu2,''), isnull(menu3,''), callform from tbl_menu order by orderdisp")

        Dim PreMenu1, PreMenu2, PreMenu3 As String

        PreMenu1 = ""
        PreMenu2 = ""
        PreMenu3 = ""

        For i = 0 To MenuDate.RecordCount - 1
            If PreMenu1 <> MenuDate(0).Value Then
                M1 = New SideBarPanelItem(MenuDate(0).Value, MenuDate(0).Value)
                M1.FontBold = True
                MenuList.Panels.Add(M1)
                PreMenu1 = MenuDate(0).Value

                BM1 = New ButtonItem(MenuDate(0).Value, MenuDate(0).Value)
                BM1.GlobalName = MenuDate(3).Value
                BM1.Text = MenuDate(0).Value
                RibbonBar1.Items.Add(BM1)
            End If

            If PreMenu1 = MenuDate(0).Value And PreMenu2 <> MenuDate(1).Value Then
                If Menu_Authority(MenuDate(3).Value) = True Then
                    M2 = New ButtonItem(MenuDate(1).Value, MenuDate(1).Value)
                    M2.GlobalName = MenuDate(3).Value
                    M2.ImagePaddingHorizontal = 13
                    M1.SubItems.Add(M2)
                    PreMenu2 = MenuDate(1).Value

                    BM2 = New ButtonItem(MenuDate(1).Value, MenuDate(1).Value)
                    BM2.GlobalName = MenuDate(3).Value

                    BM2.Text = MenuDate(1).Value
                    BM1.SubItems.Add(BM2)
                End If
            End If

            If PreMenu1 = MenuDate(0).Value And PreMenu2 = MenuDate(1).Value And MenuDate(2).Value <> "" Then
                If Menu_Authority(MenuDate(3).Value) = True Then
                    M3 = New ButtonItem(MenuDate(2).Value, MenuDate(2).Value)
                    M3.GlobalName = MenuDate(3).Value
                    M3.ImagePaddingHorizontal = 13
                    M3.ButtonStyle = eButtonStyle.TextOnlyAlways
                    M2.SubItems.Add(M3)
                    PreMenu3 = MenuDate(2).Value

                    BM3 = New ButtonItem(MenuDate(2).Value, MenuDate(2).Value)
                    BM3.GlobalName = MenuDate(3).Value
                    BM3.Text = MenuDate(2).Value
                    BM2.SubItems.Add(BM3)

                End If
            End If
            MenuDate.MoveNext()
        Next

        ShowSideMenu = True

        RibbonBar1.LayoutOrientation = eOrientation.Horizontal

        'Catch ex As Exception
        '    MessageBox.Show("Error: " & ex.Message, "ERROR")
        'End Try
    End Function

    Private Sub menulist_ItemClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles MenuList.ItemClick, RibbonBar1.ItemClick
        'Try
        '사이드 메뉴 클릭시 해당되는 Form을 호출

        Dim objItem As BaseItem = CType(sender, BaseItem)
        Dim i As Integer
        If objItem.GlobalName.Length > 0 Then

            Dim ActFrmItem As DevComponents.DotNetBar.DockContainerItem = New DevComponents.DotNetBar.DockContainerItem()

            If objItem.SubItems.Count > 0 Then
                ActFrmItem.Text = objItem.SubItems.Item(0).Text
                ActFrmItem.Name = objItem.GlobalName
            Else
                ActFrmItem.Text = objItem.Name
                ActFrmItem.Name = objItem.GlobalName
            End If

            If Query_RS("select isnull(type,'0') from tbl_menu where callform ='" & objItem.GlobalName & "'") = "1" Then

                If UCase(ActFrmItem.Text) = "MANUAL" Then
                    VwHelp_Load()
                ElseIf UCase(ActFrmItem.Text) = "Change Password" Then
                    FrmUSPwChg.ShowDialog()
                End If

            Else

                Dim actForm As New Form
                'actForm = (Activator.CreateInstance(Type.GetType("KEMS." + ActFrmItem.Name)))
                'Form 호출
                Try
                    If objItem.GlobalName = "FrmUSMontoring1" Or objItem.GlobalName = "FrmSAOrder1" Or objItem.GlobalName = "FrmMCMaster1" Or objItem.GlobalName = "FrmPRResult1" Or objItem.GlobalName = "FrmEQSum1" Then
                        actForm = (Activator.CreateInstance(Type.GetType("SMSPLUS." + Mid(objItem.GlobalName, 1, Len(objItem.GlobalName) - 1))))
                    Else
                        actForm = (Activator.CreateInstance(Type.GetType("SMSPLUS." + objItem.GlobalName)))
                    End If
                Catch ex As Exception
                    MessageBox.Show("Under Construction!")
                    Exit Sub
                End Try

                actForm.TopLevel = False 'Manfrm의 control로 자식Form을 보여주기 위해서는 False 
                actForm.Location = New System.Drawing.Point(0, 0)
                actForm.ShowInTaskbar = False
                actForm.FormBorderStyle = Windows.Forms.FormBorderStyle.None
                actForm.Dock = DockStyle.Fill

                For i = 0 To TabFrm.VisibleItemCount - 1

                    If TabFrm.Items(i).Name = ActFrmItem.Name Then
                        TabFrm.SelectedDockContainerItem = TabFrm.Items(i)
                        '                        MsgBox("Already Menu loaded")
                        Exit Sub '이미 로딩된 form인 경우 로딩이 되지 않도록 Sub 프로시져를 빠져나감
                    End If
                Next

                'Tab으로 Form 표시
                ActFrmItem.Control = actForm
                TabFrm.Items.Add(ActFrmItem)

                actForm.Show()
                actForm.Update()

                TabFrm.SelectedDockContainerItem = ActFrmItem

                If Not TabFrm.Visible Then
                    TabFrm.Visible = True
                Else
                    TabFrm.RecalcLayout()
                End If

            End If

            If MainMenu.Visible = True Then
                MainMenu.AutoHide = True
            End If

        End If


        'Catch ex As Exception
        '    MessageBox.Show("Error: " & ex.Message, "ERROR")
        'End Try
    End Sub

    Private Sub ActFormClosing(ByVal sender As Object, ByVal e As DevComponents.DotNetBar.DockTabClosingEventArgs) Handles DotNetBarManager1.DockTabClosing
        'MsgBox(TabFrm.VisibleItemCount.ToString & "," & TabFrm.Items.Count)
        '        Dim r As DialogResult = MessageBox.Show(e.DockContainerItem.Text + " 화면을 종료하시겠습니까?", "단위 화면 종료", MessageBoxButtons.YesNo)
        '       If r = DialogResult.Yes Then
        'e.DockContainerItem.Dispose()
        e.RemoveDockTab = True
        'Else
        'e.Cancel = True
        'End If
    End Sub

    Private Sub ComboBoxItem1_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBoxItem1.SelectedIndexChanged
        OP_No = Mid(ComboBoxItem1.SelectedItem, 2, InStr(ComboBoxItem1.Text, "]") - 2)
        OP_Nm = Query_RS("select emp_nm from tbl_empmaster where emp_no = '" & OP_No & "'")
    End Sub


    Private Sub ButtonItem2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonItem2.Click
        RibbonBar1.Visible = True
        MainMenu.Visible = False
        SliderItem1.Orientation = eOrientation.Horizontal
    End Sub

    Private Sub ButtonItem4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonItem4.Click
        RibbonBar1.Visible = False
        MainMenu.Visible = True
    End Sub

    'Private Sub SliderItem1_ValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles SliderItem1.ValueChanged
    '    Me.Opacity = SliderItem1.Value / 100
    'End Sub

    Private Sub disp_main()

        '로그인시 메인화면에 최초로 로딩되는 Form을 DB에서 읽어서 화면에 표시

        Dim ActFrmItem As DevComponents.DotNetBar.DockContainerItem = New DevComponents.DotNetBar.DockContainerItem()
        Dim actForm As New Form
        Dim FrmRs As ADODB.Recordset

        FrmRs = Query_RS_ALL("select isnull(a.menu3,a.menu2),a.callform from tbl_menu a, tbl_empmaster b where b.site_id = '" & Site_id & "' and a.callform = b.init_form and b.emp_no = '" & Emp_No & "'")

        If FrmRs Is Nothing Then
            Exit Sub
        End If

        actForm = Activator.CreateInstance(Type.GetType("KEMS." + FrmRs(1).Value))
        actForm.TopLevel = False
        actForm.Location = New System.Drawing.Point(0, 0)
        actForm.ShowInTaskbar = False
        actForm.FormBorderStyle = Windows.Forms.FormBorderStyle.None
        actForm.Dock = DockStyle.Fill

        ActFrmItem.Text = FrmRs(0).Value
        ActFrmItem.Name = actForm.Name
        ActFrmItem.Control = actForm
        TabFrm.Items.Add(ActFrmItem)

        actForm.Show()
        actForm.Update()
        TabFrm.RecalcLayout()
        TabFrm.SelectedDockContainerItem = ActFrmItem
    End Sub

End Class