Imports FarPoint.Win.Spread

Public Class FrmShipping


    Private PkRow As New Integer
    Private p_cnt, s_cnt As New Integer


    Private Sub FrmShipping_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        DockContainerItem1.Text = "실행 메뉴"
        DockContainerItem2.Text = "조회 조건"
        DockContainerItem3.Text = "출하 현황"
        DockContainerItem4.Text = "출하 대기 현황"
        DockContainerItem5.Text = "선적서류"
        DockContainerItem6.Text = "선적번호 부여"

        ButtonItem12.Visible = False
        LabelX4.Visible = False

        Bar4.Visible = False
        Bar5.Visible = False
        Bar6.Visible = True

        If Spread_Setting(FpSpread1, Me.Name) = True Then
            FpSpread1.ActiveSheet.RowCount = 0
            Spread_AutoCol(FpSpread1)
        End If

        If Spread_Setting(FpSpread2, Me.Name) = True Then
            FpSpread2.ActiveSheet.RowCount = 0
            FpSpread2.ActiveSheet.ColumnCount = FpSpread2.ActiveSheet.ColumnCount + 1
            FpSpread2.ActiveSheet.ColumnHeader.Columns(FpSpread2.ActiveSheet.ColumnCount - 1).Label = "선적번호"
            FpSpread2.ActiveSheet.Columns(FpSpread2.ActiveSheet.ColumnCount - 1).CellType = textcell
            'FpSpread2.ActiveSheet.ColumnCount = FpSpread2.ActiveSheet.ColumnCount + 1
            'FpSpread2.ActiveSheet.ColumnHeader.Columns(FpSpread2.ActiveSheet.ColumnCount - 1).Label = "모델"
            'FpSpread2.ActiveSheet.Columns(FpSpread2.ActiveSheet.ColumnCount - 1).CellType = textcell
            'FpSpread2.ActiveSheet.ColumnCount = FpSpread2.ActiveSheet.ColumnCount + 1
            'FpSpread2.ActiveSheet.ColumnHeader.Columns(FpSpread2.ActiveSheet.ColumnCount - 1).Label = "상태"
            'FpSpread2.ActiveSheet.Columns(FpSpread2.ActiveSheet.ColumnCount - 1).CellType = textcell

            Spread_AutoCol(FpSpread2)
        End If

        DateTimeInput1.Value = Now
        DateTimeInput2.Value = Now

        If Query_Combo(Me.ComboBoxEx4, "SELECT CODE_ID FROM tbl_CODEmaster WHERE site_id = '" & Site_id & "' and CLASS_ID = 'R0007' AND active = 'Y' ORDER BY DIS_ORDER") = True Then
        End If
        Me.ComboBoxEx4.Items.Add("ALL")
        Me.ComboBoxEx4.Text = "ALL"

        If Query_Combo(Me.ComboBoxEx3, "SELECT model_no FROM tbl_modelmaster WHERE site_id = '" & Site_id & "' and active = 'Y' ORDER BY model_no") = True Then
        End If
        Me.ComboBoxEx3.Items.Add("ALL")
        Me.ComboBoxEx3.Text = "ALL"


        If Query_Combo(Me.ComboBoxEx1, "SELECT CUS_nm FROM tbl_CUSTOMER WHERE site_id = '" & Site_id & "' and CUS_TYPE = '1차' ORDER BY CUS_NM") = True Then
        End If
        Me.ComboBoxEx1.Items.Add("ALL")
        Me.ComboBoxEx1.Text = "ALL"


        With FpSpread5.ActiveSheet
            .RowCount = 0
            .ColumnCount = 0
        End With

        Formbim_Authority(Me.NewBtn, Me.Name, "NEW")
        Formbtn_Authority(Me.NewBtn1, Me.Name, "NEW")
        Formbim_Authority(Me.SaveBtn, Me.Name, "SAVE")
        Formbtn_Authority(Me.SaveBtn1, Me.Name, "SAVE")
        Formbim_Authority(Me.DelBtn, Me.Name, "DELETE")
        Formbtn_Authority(Me.DelBtn1, Me.Name, "DELETE")
        Formbim_Authority(Me.PrtBtn, Me.Name, "PRINT")
        Formbtn_Authority(Me.PrtBtn1, Me.Name, "PRINT")
        Formbim_Authority(Me.XlsBtn, Me.Name, "EXCEL")
        Formbtn_Authority(Me.XlsBtn1, Me.Name, "EXCEL")
        Formbim_Authority(Me.FindBtn, Me.Name, "FIND")

        SaveBtn.Enabled = False
        SaveBtn1.Enabled = False

        RichTextBoxEx1.Text = ""

        DockContainerItem3.Selected = True

    End Sub


    Private Sub ListViewEx1_ItemChecked(ByVal sender As Object, ByVal e As System.Windows.Forms.ItemCheckedEventArgs) Handles ListViewEx1.ItemChecked
        Dim I As New Integer

        If e.Item.Checked = True Then

            With FpSpread1.ActiveSheet
                If .RowCount > 0 Then

                    'MessageBox.Show(e.Item.SubItems(1).Text)
                    'MessageBox.Show(.GetValue(0, 4))

                    If e.Item.SubItems(4).Text <> .GetValue(0, 5) Then
                        Modal_Error("발주번호가 틀립니다.!!")
                        e.Item.Checked = False
                        Exit Sub
                    End If
                End If

            End With

            If e.Item.SubItems(2).Text <> e.Item.SubItems(3).Text Then
                Modal_Error("박스의 포장 수량이 완료되지 않았습니다!!")
                e.Item.Checked = False
                Exit Sub
            End If


            If e.Item.Text = "TOTAL" Then
                For I = 0 To ListViewEx1.Items.Count - 2
                    ListViewEx1.Items(I).Checked = True
                Next
            Else
                Dim QRY As String

                QRY = "select LOT_NO, MODEL, STATUS, O_QTY, (SELECT COUNT(PROD_NO) FROM TBL_PRODMASTER_B WHERE LOT_NO = A.LOT_NO), S_PONO, '" & FpSpread2.ActiveSheet.GetValue(0, 1) & "','','' FROM tbl_LOTmaster A" & vbNewLine
                QRY = QRY & "WHERE SITE_ID = '" & Site_id & "'" & vbNewLine
                QRY = QRY & "  AND LOT_NO = '" & e.Item.Text & "'" & vbNewLine
                '      QRY = QRY & "ORDER BY CONVERT(INT,IN_POS) DESC" & vbNewLine

                If Query_Spread_LTD(FpSpread1, QRY, 0, 7) = True Then
                    FpSpread1.ActiveSheet.Rows(0, FpSpread1.ActiveSheet.RowCount - 1).ForeColor = Color.OrangeRed
                    Spread_AutoCol(FpSpread1)
                End If
            End If
        Else
            If e.Item.Text = "TOTAL" Then
                For I = 0 To ListViewEx1.Items.Count - 2
                    ListViewEx1.Items(I).Checked = False
                Next
            Else
                For I = 0 To FpSpread1.ActiveSheet.RowCount - 1
                    If e.Item.Text = FpSpread1.ActiveSheet.GetValue(I, 0) Then
                        FpSpread1.ActiveSheet.RemoveRows(I, 1)
                        Exit For
                    End If
                Next
            End If
        End If

    End Sub


    Private Sub ButtonItem7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonItem7.Click, FpSpread2.CellDoubleClick

        FpSpread1.ActiveSheet.RowCount = 0

        FpSpread1.ActiveSheet.FrozenColumnCount = 3

        If FpSpread2.ActiveSheet.RowCount = 0 Then
            MessageBox.Show("출하번호를 선택하세요.")
            Exit Sub
        End If

        If FpSpread2.ActiveSheet.Rows(FpSpread2.ActiveSheet.ActiveRowIndex).ForeColor = Color.OrangeRed Then
            MessageBox.Show("생산출하된 출하번호가 아닙니다.")
        Else
            MainFrm.ProgressBarItem1.ProgressType = DevComponents.DotNetBar.eProgressItemType.Marquee

            Dim QRY As String
            QRY = ""
            QRY += "SELECT LOT_NO, MODEL, '',O_QTY, (SELECT COUNT(PROD_NO) FROM TBL_PRODMASTER_I WHERE SHIP_NO = A.P_PONO), S_PONO, P_PONO, RESERV1, RESERV2" & vbNewLine 'RESERV1은 PALLET ID, RESERV2는 리비전
            QRY += "FROM TBL_LOTMASTER_B A " & vbNewLine
            QRY += "WHERE P_PONO = '" & FpSpread2.ActiveSheet.GetValue(FpSpread2.ActiveSheet.ActiveRowIndex, 1) & "'" & vbNewLine
            QRY += "order BY  S_PONO, LOT_NO, MODEL " & vbNewLine

            If Query_Spread(FpSpread1, QRY, 1) = True Then
                Spread_AutoCol(FpSpread1)
                LabelX4.Text = FpSpread2.ActiveSheet.GetValue(FpSpread2.ActiveSheet.ActiveRowIndex, 1)

            End If
            MainFrm.ProgressBarItem1.ProgressType = DevComponents.DotNetBar.eProgressItemType.Standard
            MessageBox.Show("조회 완료되었습니다.")

        End If
    End Sub

    Function Refresh_List() As Boolean
        Dim qry As String = ""
        Dim i, O_qty, P_QTY As New Integer

        qry += "SELECT SUBSTRING(SHIP_NO,3,8), SHIP_NO, COUNT(SHIP_NO), '', ''" & vbNewLine ' RESERV1은 PALLET ID, RESERV2는 리비전
        qry += "FROM TBL_PRODmaster_I a" & vbNewLine
        qry += "WHERE SUBSTRING(SHIP_NO,3,8) BETWEEN CONVERT(VARCHAR(8), '" & DateTimeInput1.Text & "',112) AND  CONVERT(VARCHAR(8), '" & DateTimeInput2.Text & "',112)" & vbNewLine
        If ComboBoxEx3.Text <> "ALL" Then
            qry += "  AND MODEL = '" & ComboBoxEx3.Text & "'" & vbNewLine
        End If
        'If ComboBoxEx4.Text <> "ALL" Then
        '    qry += "  AND RETURN_DV = '" & ComboBoxEx4.Text & "'" & vbNewLine
        'End If
        qry += "GROUP BY SUBSTRING(SHIP_NO,3,8), SHIP_NO" & vbNewLine
        qry += "order BY  SUBSTRING(SHIP_NO,3,8), SHIP_NO" & vbNewLine

        If Query_Spread(FpSpread2, qry, 1) = True Then
            'If Query_Spread(FpSpread2, "EXEC SP_FRMSHIPPING_GETSHIPPEDSUMMARY '" & Site_id & "','" & DateTimeInput1.Text & "','" & DateTimeInput2.Text & "'", 1) = True Then
            FpSpread2.ActiveSheet.RowCount = FpSpread2.ActiveSheet.RowCount + 1
            FpSpread2.ActiveSheet.Rows(FpSpread2.ActiveSheet.RowCount - 1).BackColor = Color.Yellow
            FpSpread2.ActiveSheet.SetValue(FpSpread2.ActiveSheet.RowCount - 1, 0, "TOTAL")

            FpSpread2.AllowUserFormulas = True

            FpSpread2.ActiveSheet.Cells(FpSpread2.ActiveSheet.RowCount - 1, 2).Formula = "SUM(C1:C" & FpSpread2.ActiveSheet.RowCount - 1 & ")"

            Spread_AutoCol(FpSpread2)
        End If

        qry = ""
        qry += "SELECT LOT_NO, MODEL, O_QTY, (SELECT COUNT(PROD_NO) FROM TBL_PRODMASTER_B WHERE LOT_NO = A.LOT_NO), S_PONO, (SELECT CUS_NM FROM TBL_CUSTOMER WHERE CUS_NO = A.C_NO)" & vbNewLine
        qry += "FROM TBL_LOTMASTER A " & vbNewLine
        qry += "WHERE C_PRC = 'W3000'" & vbNewLine
        If ComboBoxEx3.Text <> "ALL" Then
            qry += "  AND MODEL = '" & ComboBoxEx3.Text & "'" & vbNewLine
        End If
        'If ComboBoxEx4.Text <> "ALL" Then
        '    qry += "  AND RETURN_DV = '" & ComboBoxEx4.Text & "'" & vbNewLine
        'End If
        qry += "order BY  S_PONO, LOT_NO, MODEL " & vbNewLine


        If Query_Listview(ListViewEx1, qry, True) = True Then
            '            If Query_Listview(ListViewEx1, "exec SP_FRMshipping_GETCURBOX '" & Site_id & "', 'all'", True) = True Then

            For i = 0 To ListViewEx1.Items.Count - 1
                O_qty = O_qty + CInt(ListViewEx1.Items(i).SubItems(2).Text)
                P_QTY = P_QTY + CInt(ListViewEx1.Items(i).SubItems(3).Text)
            Next

            ListViewEx1.Items.Add("TOTAL")
            ListViewEx1.Items(ListViewEx1.Items.Count - 1).SubItems.Add("")
            ListViewEx1.Items(ListViewEx1.Items.Count - 1).SubItems.Add(O_qty)
            ListViewEx1.Items(ListViewEx1.Items.Count - 1).SubItems.Add(P_QTY)
            ListViewEx1.Items(ListViewEx1.Items.Count - 1).SubItems.Add("")
            ListViewEx1.Items(ListViewEx1.Items.Count - 1).SubItems.Add("")
            ListViewEx1.Items(ListViewEx1.Items.Count - 1).BackColor = Color.Yellow
            ListViewEx1.Items(ListViewEx1.Items.Count - 1).SubItems(1).BackColor = Color.Yellow
            ListViewEx1.Items(ListViewEx1.Items.Count - 1).SubItems(2).BackColor = Color.Yellow
            ListViewEx1.Items(ListViewEx1.Items.Count - 1).SubItems(3).BackColor = Color.Yellow
            ListViewEx1.Items(ListViewEx1.Items.Count - 1).SubItems(4).BackColor = Color.Yellow
            ListViewEx1.Items(ListViewEx1.Items.Count - 1).SubItems(5).BackColor = Color.Yellow


            ListViewEx1.AutoResizeColumn(0, ColumnHeaderAutoResizeStyle.ColumnContent)
            'Bar4.AutoHide = False
        End If

    End Function


    Private Sub FindBtn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FindBtn.Click
        Dim i, i_qty As New Integer

        LabelX4.Text = ""
        FpSpread1.ActiveSheet.RowCount = 0
        FpSpread2.ActiveSheet.RowCount = 0
        Bar5.Visible = False

        'ButtonItem16.Visible = True

        For i = 0 To ListViewEx1.Items.Count - 1
            ListViewEx1.Items(i).Checked = False
        Next

        ListViewEx1.CheckBoxes = False

        'Bar5.AutoHide = True

        Refresh_List()

        SaveBtn.Enabled = False
        SaveBtn1.Enabled = False

        DockContainerItem3.Selected = True

    End Sub

    Private Sub ButtonItem8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonItem8.Click, ButtonX1.Click
        Try

            If MessageBox.Show("선적서류를 조회하시겠습니까?", "선적서류 조회", MessageBoxButtons.YesNo) <> Windows.Forms.DialogResult.Yes Then
                Exit Sub
            End If

            Dim PK_RS As New ADODB.Recordset

            If FpSpread2.ActiveSheet.Rows(FpSpread2.ActiveSheet.ActiveRowIndex).ForeColor = Color.OrangeRed Then
                MessageBox.Show("출하번호 오류!")
            Else
                MainFrm.ProgressBarItem1.ProgressType = DevComponents.DotNetBar.eProgressItemType.Marquee

                Dim S3 = Me.FpSpread3_Sheet1
                Dim S4 = Me.FpSpread3_Sheet2
                Dim S5 = Me.FpSpread3_Sheet3
                Dim pmd1 As String = ""
                Dim pmd2 As String = ""
                Dim pmd3 As String = ""

                S3.Visible = True
                S4.Visible = True
                S5.Visible = True

                S3.Cells(3, 7).Value = FpSpread2.ActiveSheet.GetValue(FpSpread2.ActiveSheet.ActiveRowIndex, 1)
                S3.Cells(3, 12).Value = Query_RS("select CONVERT(VARCHAR(10), GETDATE(), 21)")
                S3.Cells(21, 4).Value = Query_RS("select CONVERT(VARCHAR(10), GETDATE(), 21)")

                S4.Cells(3, 7).Value = FpSpread2.ActiveSheet.GetValue(FpSpread2.ActiveSheet.ActiveRowIndex, 1)
                S4.Cells(3, 12).Value = Query_RS("select CONVERT(VARCHAR(10), GETDATE(), 21)")
                S4.Cells(21, 4).Value = Query_RS("select CONVERT(VARCHAR(10), GETDATE(), 21)")

                S5.Cells(6, 7).Value = "From : " & Query_RS("Select CONVERT(VARCHAR(10), GETDATE(), 102)") & " To" & Query_RS("Select CONVERT(VARCHAR(10), GETDATE()+365, 102)")

                curcell.DecimalPlaces = 2
                curcell.CurrencySymbol = "$"
                curcell.ShowSeparator = True
                curcell.Separator = ","

                Dim QRY As String = ""
                With FpSpread2.ActiveSheet
                    QRY = QRY & "select ISNULL((SELECT MODEL_NAME FROM TBL_MODELMASTER WHERE MODEL_NO = MODEL),MODEL)," & vbNewLine
                    QRY = QRY & "       COUNT(SHIP_NO)," & vbNewLine
                    QRY = QRY & "       ISNULL((SELECT PRICE FROM TBL_PACKINGDET WHERE LOT_NO = A.LOT_NO AND MODEL = A.MODEL),0)," & vbNewLine
                    QRY = QRY & "       MODEL" & vbNewLine
                    QRY = QRY & "FROM TBL_PRODMASTER_I A" & vbNewLine
                    QRY = QRY & "WHERE SHIP_NO = '" & .GetValue(.ActiveRowIndex, 1) & "'" & vbNewLine
                    QRY = QRY & "GROUP BY SHIP_NO,LOT_NO, MODEL" & vbNewLine
                    QRY = QRY & "ORDER BY SHIP_NO,LOT_NO, MODEL" & vbNewLine
                End With

                PK_RS = Query_RS_ALL(QRY)


                If PK_RS Is Nothing Then
                    Exit Sub
                Else
                    For I As Integer = 0 To PK_RS.RecordCount - 1
                        S3.Cells(27 + I, 4).Value = PK_RS(0).Value
                        S3.Cells(27 + I, 8).Value = PK_RS(1).Value
                        S3.Cells(27 + I, 10).Value = "USD" 'CDec(PK_RS(2).Value)
                        S3.Cells(27 + I, 12).Value = "USD" 'S3.Cells(27 + I, 8).Value * S3.Cells(27 + I, 10).Value

                        S3.Cells(27 + I, 9).Value = "EA"
                        S3.Cells(27 + I, 11).Value = PK_RS(2).Value '"48.00"
                        S3.Cells(27 + I, 13).Value = S3.Cells(27 + I, 8).Value * S3.Cells(27 + I, 11).Value


                        S4.Cells(27 + I, 4).Value = PK_RS(0).Value
                        S4.Cells(27 + I, 8).Value = PK_RS(1).Value
                        S4.Cells(27 + I, 10).Value = CDec(PK_RS(2).Value)
                        ' S4.Cells(27 + I, 12).CellType = curcell
                        S4.Cells(27 + I, 12).Value = S4.Cells(27 + I, 8).Value * S4.Cells(27 + I, 10).Value

                        S4.Cells(27 + I, 9).Value = "EA"
                        S4.Cells(27 + I, 11).Value = "KG"
                        S4.Cells(27 + I, 13).Value = "KG"

                        S5.Cells(21 + I, 0).Value = I + 1
                        S5.Cells(21 + I, 1).Value = PK_RS(0).Value
                        S5.Cells(21 + I, 4).Value = PK_RS(1).Value

                        PK_RS.MoveNext()
                    Next
                End If

                S3.Cells(47, 8).Formula = "SUM(I28:I46)"
                'S3.Cells(35, 11).Formula = "SUM(L28:L35)"
                S3.Cells(47, 12).Formula = "SUM(N28:N46)"

                S3.Cells(47, 12).CellType = curcell

                S4.Cells(47, 8).Formula = "SUM(I28:I46)"
                S4.Cells(47, 10).Formula = "SUM(K28:K46)"
                S4.Cells(47, 12).Formula = "SUM(M28:M46)"


                S4.Cells(30, 2).Value = Query_RS("SELECT COUNT(LOT_NO) FROM TBL_LOTMASTER_B WHERE P_PONO = '" & FpSpread2.ActiveSheet.GetValue(FpSpread2.ActiveSheet.ActiveRowIndex, 1) & "'")
                S4.Cells(32, 2).Value = S4.Cells(46, 12).Value




                Bar5.Visible = True
            End If


            MainFrm.ProgressBarItem1.ProgressType = DevComponents.DotNetBar.eProgressItemType.Standard
            'Bar5.AutoHide = False
            '            Bar5.Left = FpSpread1.Left
            DockContainerItem5.Width = 800

        Catch ex As Exception
            MessageBox.Show("Error : " & ex.Message, "Error")
        End Try
    End Sub

    Private Sub NewBtn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NewBtn.Click, NewBtn1.Click
        Dim i, O_qty, P_QTY As New Integer
        LabelX4.Text = ""

        FpSpread1.ActiveSheet.RowCount = 0
        FpSpread2.ActiveSheet.RowCount = 0

        FpSpread2.ActiveSheet.RowCount = 1
        FpSpread2.ActiveSheet.Rows(0).ForeColor = Color.OrangeRed
        ListViewEx1.CheckBoxes = True

        ListViewEx1.Items.Clear()

        If ComboBoxEx1.Text = "ALL" Then
            Modal_Error("출고할 고객을 선택하세요.")
            Exit Sub
        End If


        'Bar6.Visible = True

        FpSpread2.ActiveSheet.SetValue(0, 0, Query_RS("select convert(varchar(8), getdate(), 112)"))

        Dim FILESEQ As String
        FILESEQ = Microsoft.VisualBasic.Right("00000" & CStr(CInt(Query_RS("SELECT CODE_NAME FROM TBL_CODEMASTER WHERE CLASS_ID = 'R0055' AND CODE_ID = 'A1000'")) + 1), 5)
        FpSpread2.ActiveSheet.SetValue(0, 1, "SK" & Query_RS("select convert(varchar(8), getdate(), 112)") & "-" & FILESEQ)


        Dim qry As String
        qry = ""
        qry += "SELECT LOT_NO, MODEL, O_QTY, (SELECT COUNT(PROD_NO) FROM TBL_PRODMASTER_B WHERE LOT_NO = A.LOT_NO), S_PONO, (SELECT CUS_NM FROM TBL_CUSTOMER WHERE CUS_NO = A.C_NO)" & vbNewLine
        qry += "FROM TBL_LOTMASTER A " & vbNewLine
        qry += "WHERE C_PRC = 'W3000'" & vbNewLine
        If ComboBoxEx3.Text <> "ALL" Then
            qry += "  AND MODEL = '" & ComboBoxEx3.Text & "'" & vbNewLine
        End If
        qry += "  and c_no in (SELECT CUS_NO FROM TBL_CUSTOMER WHERE CUS_NM = '" & ComboBoxEx1.Text & "')" & vbNewLine

        qry += "order BY  S_PONO, LOT_NO, MODEL " & vbNewLine



        If Query_Listview(ListViewEx1, qry, True) = True Then

            For i = 0 To ListViewEx1.Items.Count - 1
                O_qty = O_qty + CInt(ListViewEx1.Items(i).SubItems(2).Text)
                P_QTY = P_QTY + CInt(ListViewEx1.Items(i).SubItems(3).Text)
            Next

            ListViewEx1.Items.Add("TOTAL")
            ListViewEx1.Items(ListViewEx1.Items.Count - 1).SubItems.Add("")
            ListViewEx1.Items(ListViewEx1.Items.Count - 1).SubItems.Add(O_qty)
            ListViewEx1.Items(ListViewEx1.Items.Count - 1).SubItems.Add(P_QTY)
            ListViewEx1.Items(ListViewEx1.Items.Count - 1).SubItems.Add("")
            ListViewEx1.Items(ListViewEx1.Items.Count - 1).SubItems.Add("")
            ListViewEx1.Items(ListViewEx1.Items.Count - 1).BackColor = Color.Yellow
            ListViewEx1.Items(ListViewEx1.Items.Count - 1).SubItems(1).BackColor = Color.Yellow
            ListViewEx1.Items(ListViewEx1.Items.Count - 1).SubItems(2).BackColor = Color.Yellow
            ListViewEx1.Items(ListViewEx1.Items.Count - 1).SubItems(3).BackColor = Color.Yellow
            ListViewEx1.Items(ListViewEx1.Items.Count - 1).SubItems(4).BackColor = Color.Yellow
            ListViewEx1.Items(ListViewEx1.Items.Count - 1).SubItems(5).BackColor = Color.Yellow


            ListViewEx1.AutoResizeColumn(0, ColumnHeaderAutoResizeStyle.ColumnContent)
            'Bar4.AutoHide = False
        End If
        DockContainerItem4.Selected = True


        SaveBtn.Enabled = True
        SaveBtn1.Enabled = True

        TextBoxX1.Focus()

    End Sub

    Private Sub SaveBtn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SaveBtn.Click, SaveBtn1.Click
        Dim i As New Integer

        Me.Cursor = Cursors.WaitCursor

        If ComboBoxEx1.Text = "ALL" Then
            Modal_Error("출고할 고객을 선택하세요.")
            Exit Sub
        End If

        With FpSpread1.ActiveSheet
            For X As Integer = 0 To .RowCount - 1
                If .GetText(X, 7) = "" Then
                    System.Console.Beep(3000, 400)
                    System.Console.Beep(3000, 400)
                    Modal_Error("PALLET 구성이 되지 않았습니다. : " & .GetValue(X, 0) & "!!")
                    Exit Sub
                End If
            Next
        End With


        If MessageBox.Show("출하번호 " & FpSpread2.ActiveSheet.GetValue(0, 1) & "로 출하 하시겠습니까?", "출하", MessageBoxButtons.YesNo) = Windows.Forms.DialogResult.Yes Then
            SaveBtn.Enabled = False
            MainFrm.ProgressBarItem1.Maximum = FpSpread1.ActiveSheet.RowCount

            ' 기사용한 Shipping No 인지 Check
            If Query_RS("select count(P_POno) from TBL_LOTmaster_B where P_POno = '" & FpSpread2.ActiveSheet.GetValue(0, 1) & "'") > 0 Then
                System.Console.Beep(3000, 400)
                System.Console.Beep(3000, 400)
                Modal_Error("이미 사용된 출하번호 입니다. : " & FpSpread2.ActiveSheet.GetValue(0, 1) & "!!")

                FindBtn_Click(sender, e)
                Exit Sub
            End If

            Dim SHIP_NO As String = FpSpread2.ActiveSheet.GetValue(0, 1)

            For i = 0 To FpSpread1.ActiveSheet.RowCount - 1

                MainFrm.ProgressBarItem1.Value = i + 1

                If FpSpread1.ActiveSheet.Rows(i).ForeColor = Color.OrangeRed Then
                    If Insert_Data("EXEC SP_FRMSHIPPING_SAVE '" & Site_id & "','" & FpSpread1.ActiveSheet.GetValue(i, 0) & "','" & Emp_No & "','W9000','" & FpSpread2.ActiveSheet.GetValue(0, 1) & "','" & FpSpread1.ActiveSheet.GetValue(i, 7) & "'") = True Then
                        FpSpread1.ActiveSheet.Rows(i).ForeColor = Color.Black
                    End If
                End If
            Next


            FpSpread1.ActiveSheet.RowCount = 0
            MainFrm.ProgressBarItem1.Value = 0
            Bar4.AutoHide = True

            ListViewEx1.CheckBoxes = False
        Else

        End If

        SaveBtn.Enabled = False
        SaveBtn1.Enabled = False

        MessageBox.Show("저장되었습니다")

        Refresh_List()


        'Bar6.Visible = False
        Me.Cursor = Cursors.Default
        SaveBtn.Enabled = True
        DockContainerItem3.Selected = True

    End Sub

    Private Sub XlsBtn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles XlsBtn.Click, XlsBtn1.Click

        If save_excel = "" Then
            MessageBox.Show("Select Spread for saving!!")
            Exit Sub
        End If


        If save_excel = "FpSpread1" Then
            SaveFileDialog1.FileName = LabelX4.Text
            'File_XML(SaveFileDialog1, FpSpread1)
            File_Save(SaveFileDialog1, FpSpread1)
            '            MAKE_XML()
            SaveFileDialog1.FileName = ""

        ElseIf save_excel = "FpSpread2" Then
            File_Save(SaveFileDialog1, FpSpread2)
        ElseIf save_excel = "FpSpread3" Then
            File_Save2(SaveFileDialog1, FpSpread3, PkRow)
        Else
            MessageBox.Show("Select Spread for Save!!")
        End If

    End Sub




    Private Sub PrtBtn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PrtBtn.Click, PrtBtn1.Click, ButtonItem6.Click
        '      Try
        Dim i As New Integer

        If save_excel = "" Then
            MessageBox.Show("Select Spread for Printing!!")
            Exit Sub
        End If

        If save_excel = "FpSpread1" Then
            If Spread_Print(Me.FpSpread1, "Repair Upload", 0) = False Then
                MsgBox("Fail to Print")
            End If
        ElseIf save_excel = "FpSpread2" Then
            If Spread_Print(Me.FpSpread2, "Shipping Summary", 0) = False Then
                MsgBox("Fail to Print")
            End If
        ElseIf save_excel = "FpSpread3" Then

            Spread_Print(FpSpread3, "INVOICE", 0)
        Else
            MessageBox.Show("No Printing Object!!")
        End If

        'Catch ex As Exception
        '    MessageBox.Show("Error : " & ex.Message, "Error")
        'End Try
    End Sub

    Private Sub sheet_print2()
        Try

            Dim pi As New FarPoint.Win.Spread.PrintInfo()
            Dim pm As New FarPoint.Win.Spread.PrintMargin

            pm.Right = 0
            pm.Top = 40
            pm.Bottom = 20
            pm.Left = 30

            pi.ShowBorder = False
            pi.ShowColor = False
            pi.ShowGrid = False
            pi.ShowShadows = False
            pi.UseMax = False
            pi.Margin = pm

            pi.Orientation = FarPoint.Win.Spread.PrintOrientation.Landscape

            pi.Centering = FarPoint.Win.Spread.Centering.Both

            pi.ColStart = 0
            pi.RowStart = 0
            pi.ColEnd = 14
            pi.RowEnd = 44
            pi.PrintType = FarPoint.Win.Spread.PrintType.PageRange
            pi.PageStart = 1
            pi.PageEnd = 2
            pi.Preview = True
            'pi.ShowPrintDialog = True

            pi.UseSmartPrint = False
            pi.PrintType = FarPoint.Win.Spread.PrintType.PageRange

            With FpSpread3
                .Sheets(0).PrintInfo = pi
                .PrintSheet(0)
            End With

            'With FpSpread5
            '    If .ActiveSheet.GetValue(12, 1) <> "" Then
            '        .Sheets(0).PrintInfo = pi
            '        .PrintSheet(0)
            '    End If
            'End With

            'With FpSpread6
            '    If .ActiveSheet.GetValue(12, 1) <> "" Then
            '        .Sheets(0).PrintInfo = pi
            '        .PrintSheet(0)
            '    End If
            'End With

            pi = Nothing
            pm = Nothing

        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message, "ERROR")
        End Try
    End Sub


    Private Sub FpSpread2_CellClick(ByVal sender As System.Object, ByVal e As FarPoint.Win.Spread.CellClickEventArgs) Handles FpSpread2.CellClick
        save_excel = "FpSpread2"
    End Sub

    Private Sub FpSpread1_CellClick(ByVal sender As System.Object, ByVal e As FarPoint.Win.Spread.CellClickEventArgs) Handles FpSpread1.CellClick
        save_excel = "FpSpread1"
    End Sub

    Private Sub FpSpread3_CellClick(ByVal sender As System.Object, ByVal e As FarPoint.Win.Spread.CellClickEventArgs) Handles FpSpread3.CellClick
        save_excel = "FpSpread3"
    End Sub

    Private Sub TextBoxX1_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TextBoxX1.KeyDown


        If e.KeyValue = Keys.Enter Then

            If ListViewEx1.FindItemWithText(TextBoxX1.Text) Is Nothing Then
                Modal_Error(TextBoxX1.Text & vbNewLine & "Wrong BOXID !!")
            Else
                ListViewEx1.FindItemWithText(TextBoxX1.Text).Checked = True
            End If
            TextBoxX1.Text = ""
        End If

    End Sub


    Private Sub ButtonItem3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonItem3.Click
        If Me.ButtonItem3.Text = "Go to Next page" Then
            'Me.FpSpread3.ActiveSheet = Me.FpSpread3_Sheet2
            Me.ButtonItem3.Text = "Go to First page"
        Else
            Me.FpSpread3.ActiveSheet = Me.FpSpread3_Sheet1
            Me.ButtonItem3.Text = "Go to Next page"
        End If

    End Sub


    Private Sub FpSpread3_CellDoubleClick(ByVal sender As Object, ByVal e As FarPoint.Win.Spread.CellClickEventArgs) Handles FpSpread3.CellDoubleClick

        Me.FpSpread3.ActiveSheet.Cells(1, 11).Locked = False

    End Sub

    Private Sub FpSpread3_Change(ByVal sender As Object, ByVal e As FarPoint.Win.Spread.ChangeEventArgs) Handles FpSpread3.Change
        If Me.FpSpread3.ActiveSheetIndex = 0 Then
            Me.FpSpread3.ActiveSheet.Cells(1, 11).Locked = False
            Me.FpSpread3.Sheets(1).Cells(1, 11).Text = Me.FpSpread3.Sheets(0).Cells(1, 11).Text
        End If
        If Me.FpSpread3.ActiveSheetIndex = 1 Then
            Me.FpSpread3.ActiveSheet.Cells(1, 11).Locked = False
            Me.FpSpread3.Sheets(0).Cells(1, 11).Text = Me.FpSpread3.Sheets(1).Cells(1, 11).Text
        End If
    End Sub

    Private Sub ButtonItem11_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonItem11.Click
        Try

            With FpSpread1.ActiveSheet

                If .RowCount = 0 Then
                    Exit Sub
                End If

                For I As Integer = 0 To .RowCount - 1
                    Print_PalletBarcode(.GetValue(I, 6), .GetValue(I, 7), .GetValue(I, 8), RichTextBox1)
                Next

            End With



            'Dim S1 = Me.FpSpread1.ActiveSheet
            'Dim s2 = Me.FpSpread2.ActiveSheet
            'Dim S4 = Me.FpSpread4.ActiveSheet
            'Dim i As Integer
            'If S1.RowCount < 1 Then
            '    Exit Sub
            'End If

            'If S4.RowCount > 0 Then
            '    S4.Rows.Clear()
            'End If

            'For i = 0 To S1.RowCount - 1
            '    S4.RowCount += 1
            '    S4.SetValue(i, 0, S1.GetValue(i, 2))
            'Next

            'SaveFileDialog1.FileName = s2.GetValue(s2.ActiveRowIndex, 1) & "-DMCHK"
            'File_Save(SaveFileDialog1, FpSpread4)

        Catch ex As Exception
            MessageBox.Show("Error : " & ex.Message, "Error")
        End Try
    End Sub


    Private Sub ButtonItem12_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonItem12.Click

        If MessageBox.Show("Are you sure to fix Error ?", "Fix Error", MessageBoxButtons.YesNo) = Windows.Forms.DialogResult.Yes Then
            MessageBox.Show(Query_RS("EXEC SP_FRMSHIPPING_SHIPERROR"))
            FindBtn_Click(sender, e)
            CheckBoxX1.Checked = False
        End If

    End Sub

    Private Sub CheckBoxX1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBoxX1.CheckedChanged
        If CheckBoxX1.Checked = True Then
            ButtonItem12.Visible = True
        Else
            ButtonItem12.Visible = False

        End If
    End Sub

    Private Sub ButtonItem14_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonItem14.Click

        With ListViewEx1
            If .Items.Count > 0 Then

                'PALLET 정보 바코드 출력


                'Dim r As DialogResult = MessageBox.Show("Are you sure to cancel box " & .SelectedItems(0).SubItems(3).Text & "?", "Cancel Packing BOX", MessageBoxButtons.YesNo)
                'If r = Windows.Forms.DialogResult.Yes Then
                '    Insert_Data("EXEC SP_FRMPACKING_CANCELPACKING_PQC '" & Site_id & "','" & .SelectedItems(0).SubItems(3).Text & "','" & .SelectedItems(0).Text & "','" & Emp_No & "'")
                '    '                    Refresh_Current(ComboBoxEx3.Text)
                'End If

            End If
        End With


    End Sub

    Private Sub TextBoxX2_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TextBoxX2.KeyDown




    End Sub




    Private Sub ButtonItem16_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonItem16.Click

        With FpSpread2.ActiveSheet


            If .RowCount = 0 Then
                Exit Sub
            End If

            Dim QRY As String = "SELECT DISTINCT (SELECT POORDER FROM TBL_MODELMASTER WHERE MODEL_NO = (SELECT TOP 1 MODEL FROM TBL_ESNMASTER_B WHERE SHIP_NO = A.SHIP_NO)), " & .GetValue(.ActiveRowIndex, 2) & ", INV_NO,(select count(distinct inboxid) from tbl_esnmaster_b where ship_no = a.ship_no group by ship_no),(SELECT RESERV4 FROM TBL_MODELMASTER WHERE MODEL_NO = A.MODEL)--INBOXID " & vbNewLine
            QRY = QRY & "FROM TBL_fESNMASTER_B A" & vbNewLine
            QRY = QRY & "WHERE pSHIP_NO = '" & .GetValue(.ActiveRowIndex, 1) & "'" & vbNewLine
            QRY = QRY & "GROUP BY SHIP_NO, INV_NO, MODEL--INBOXID" & vbNewLine

            Dim CB_RS As ADODB.Recordset = Query_RS_ALL(QRY)

            If CB_RS Is Nothing Then
                MessageBox.Show("No Data!!")
                Exit Sub
            Else
                Print_Barcode_Pallet(CB_RS(0).Value, CB_RS(1).Value, CB_RS(3).Value, CB_RS(2).Value, CB_RS(4).Value)
            End If

        End With

    End Sub


    Private Sub ButtonItem18_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonItem18.Click, ButtonX4.Click

        If MessageBox.Show("Are you sure to make invoice ?", "Make Invoice", MessageBoxButtons.YesNo) = Windows.Forms.DialogResult.No Then
            Exit Sub
        End If

        With FpSpread5.ActiveSheet
            .RowCount = 0
            .ColumnCount = 21

            .ColumnHeader.Columns(0).Label = "Repair_Year"
            .ColumnHeader.Columns(1).Label = "Repair_Month"
            .ColumnHeader.Columns(2).Label = "Invoice_Date"
            .ColumnHeader.Columns(3).Label = "Invoice_Number"
            .ColumnHeader.Columns(4).Label = "PO_Number"
            .ColumnHeader.Columns(5).Label = "Invoice_Status"
            .ColumnHeader.Columns(6).Label = "Brand"
            .ColumnHeader.Columns(7).Label = "VENDOR"
            .ColumnHeader.Columns(8).Label = "MANF"
            .ColumnHeader.Columns(9).Label = "SKU"
            .ColumnHeader.Columns(10).Label = "Warr_DESC"
            .ColumnHeader.Columns(11).Label = "CODE_LEVEL"
            .ColumnHeader.Columns(12).Label = "Charge_Description"
            .ColumnHeader.Columns(13).Label = "Qty"
            .ColumnHeader.Columns(14).Label = "Labor"
            .ColumnHeader.Columns(15).Label = "Parts"
            .ColumnHeader.Columns(16).Label = "Reclaim_Cost"
            .ColumnHeader.Columns(17).Label = "Additional_Parts_Cost"
            .ColumnHeader.Columns(18).Label = "Other_Cost"
            .ColumnHeader.Columns(19).Label = "Total"
            .ColumnHeader.Columns(20).Label = "Charge_Reconciliation"

            deccell.DecimalPlaces = 2
            deccell.Separator = ","
            deccell.ShowSeparator = True
            .Columns(14, 20).CellType = deccell


            Dim QRY As String = "EXEC SP_FRMSHIPPING_INVOICE '" & FpSpread2.ActiveSheet.GetValue(FpSpread2.ActiveSheet.ActiveRowIndex, 1) & "'"

            If Query_Spread(FpSpread5, QRY, 1) = True Then
                For I As Integer = 0 To .RowCount - 1
                    .Cells(I, 19).Formula = "SUM(O" & I + 1 & ":S" & I + 1 & ")"
                Next

                Spread_AutoCol(FpSpread5)
            End If




            Dim f As SaveFileDialog = SaveFileDialog1

            f.DefaultExt = "XLS"
            f.InitialDirectory = My.Computer.FileSystem.SpecialDirectories.MyDocuments
            f.FileName = Mid(FpSpread2.ActiveSheet.GetValue(FpSpread2.ActiveSheet.ActiveRowIndex, 1), 3, Len(FpSpread2.ActiveSheet.GetValue(FpSpread2.ActiveSheet.ActiveRowIndex, 1)) - 2) '& ".csv"
            f.Filter = "Microsoft Office Excel (*.xls)|*.xls*|All Files(*.*)|*.*"
            f.FileName = Mid(FpSpread2.ActiveSheet.GetValue(FpSpread2.ActiveSheet.ActiveRowIndex, 1), 3, Len(FpSpread2.ActiveSheet.GetValue(FpSpread2.ActiveSheet.ActiveRowIndex, 1)) - 2) '& ".csv"
            SaveFileDialog1.FileName = Mid(FpSpread2.ActiveSheet.GetValue(FpSpread2.ActiveSheet.ActiveRowIndex, 1), 6, 1) & Mid(FpSpread2.ActiveSheet.GetValue(FpSpread2.ActiveSheet.ActiveRowIndex, 1), 12, 6) '& ".csv"

            If f.ShowDialog() = Windows.Forms.DialogResult.OK Then
                .Protect = False

                'If FpSpread5.Save(f.FileName, False) = True Then
                'RichTextBoxEx1.SaveFile(f.FileName, RichTextBoxStreamType.PlainText)

                If FpSpread5.SaveExcel(f.FileName, FarPoint.Excel.ExcelSaveFlags.SaveCustomColumnHeaders) = True Then
                    f.FileName = ""
                    SaveFileDialog1.FileName = ""
                End If
            End If

            QRY = "UPDATE TBL_CODEMASTER SET CODE_NAME = CONVERT(INT,CODE_NAME) + 1 WHERE CLASS_ID = 'R3002' AND CODE_ID = 'INVNO'" & vbNewLine & vbNewLine

            Insert_Data(QRY)

        End With





    End Sub

    Private Sub ButtonX5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonX5.Click

        With FpSpread2.ActiveSheet

            If TextBoxX2.Text = "" Then
                MessageBox.Show("Type a tracking no!")
                Exit Sub
            End If

            If MessageBox.Show("Are you sure to set Tracking No to Shipping(" & .GetValue(.ActiveRowIndex, 1) & ") ?", "set Tracking No", MessageBoxButtons.YesNo) = Windows.Forms.DialogResult.No Then
                Exit Sub
            End If


            Dim qry As String = ""
            qry = "update tbl_esnmaster_b set inrano = '" & TextBoxX2.Text & "' where ship_no = '" & .GetValue(.ActiveRowIndex, 1) & "'"

            If Insert_Data(qry) <> True Then
                MessageBox.Show("Error to Set!!")
            Else
                TextBoxX2.Text = ""
                MessageBox.Show("Complete to Set!!")
            End If

        End With


    End Sub

    Private Sub FpSpread1_CellDoubleClick(sender As Object, e As CellClickEventArgs) Handles FpSpread1.CellDoubleClick

        With FpSpread1.ActiveSheet

            If .RowCount = 0 Then
                Exit Sub
            End If


            If e.Column = 7 Or e.Column = 8 Then
                .Cells(e.Row, e.Column).Locked = False
            End If


        End With

    End Sub

End Class