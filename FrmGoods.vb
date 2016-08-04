Public Class FrmGoods


    Private PkRow As New Integer
    Private p_cnt, s_cnt As New Integer
    Const CtrlMask As Byte = 8


    Private Sub FrmShipping_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        DockContainerItem1.Text = "실행 메뉴"
        DockContainerItem2.Text = "조회 조건"
        DockContainerItem3.Text = "출하 현황"

        DockContainerItem5.Text = "선적서류"
        DockContainerItem6.Text = "선적번호 부여"

        LabelX4.Visible = False

        Bar4.Visible = False
        Bar5.Visible = False
        Bar6.Visible = False


        If Spread_Setting(FpSpread1, Me.Name) = True Then
            FpSpread1.ActiveSheet.RowCount = 0
            Spread_AutoCol(FpSpread1)
        End If

        If Spread_Setting(FpSpread2, "FrmShipping") = True Then
            FpSpread2.ActiveSheet.RowCount = 0
            FpSpread2.ActiveSheet.ColumnCount = FpSpread2.ActiveSheet.ColumnCount + 1
            FpSpread2.ActiveSheet.ColumnHeader.Columns(FpSpread2.ActiveSheet.ColumnCount - 1).Label = "선적번호"
            FpSpread2.ActiveSheet.Columns(FpSpread2.ActiveSheet.ColumnCount - 1).CellType = textcell
            FpSpread2.ActiveSheet.ColumnCount = FpSpread2.ActiveSheet.ColumnCount + 1
            FpSpread2.ActiveSheet.ColumnHeader.Columns(FpSpread2.ActiveSheet.ColumnCount - 1).Label = "모델"
            FpSpread2.ActiveSheet.Columns(FpSpread2.ActiveSheet.ColumnCount - 1).CellType = textcell
            FpSpread2.ActiveSheet.ColumnCount = FpSpread2.ActiveSheet.ColumnCount + 1
            FpSpread2.ActiveSheet.ColumnHeader.Columns(FpSpread2.ActiveSheet.ColumnCount - 1).Label = "상태"
            FpSpread2.ActiveSheet.Columns(FpSpread2.ActiveSheet.ColumnCount - 1).CellType = textcell

            Spread_AutoCol(FpSpread2)
        End If

        DateTimeInput1.Value = Now
        DateTimeInput2.Value = Now

        Me.ComboBoxEx4.Items.Add("제품")
        Me.ComboBoxEx4.Items.Add("품목")
        Me.ComboBoxEx4.Text = "제품"

        If Query_Combo(Me.ComboBoxEx3, "SELECT model_no FROM tbl_modelmaster WHERE site_id = '" & Site_id & "' and active = 'Y' ORDER BY model_no") = True Then
        End If
        Me.ComboBoxEx3.Items.Add("ALL")
        Me.ComboBoxEx3.Text = "ALL"

        If Query_Combo(Me.ComboBoxEx1, "SELECT CODE_ID FROM tbl_CODEmaster WHERE site_id = '" & Site_id & "' and CLASS_ID = 'R0007' AND active = 'Y' ORDER BY DIS_ORDER") = True Then
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

    Private Sub PartList_ItemDrag(ByVal sender As System.Object, ByVal e As System.Windows.Forms.ItemDragEventArgs) Handles PartList.ItemDrag
        If e.Button = Windows.Forms.MouseButtons.Left Then
            'invoke the drag and drop operation
            DoDragDrop(e.Item, DragDropEffects.Move Or DragDropEffects.Copy)
        End If
    End Sub

    Private Sub PartList_DragEnter(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles PartList.DragEnter, FpSpread1.DragEnter
        If (e.Data.GetDataPresent("System.Windows.Forms.ListViewItem")) Then
            ' If the Ctrl key was pressed during the drag operation then perform
            ' a Copy. If not, perform a Move.
            If (e.KeyState And CtrlMask) = CtrlMask Then
                e.Effect = DragDropEffects.Copy
            Else
                e.Effect = DragDropEffects.Move
            End If
        Else
            e.Effect = DragDropEffects.None
        End If

    End Sub

    Private Sub TextBoxX1_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TextBoxX1.KeyUp
        Try
            If e.KeyCode <> Keys.Enter Then
                Exit Sub
            End If

            Dim QRY As String

            QRY = "SELECT PART_NO, PART_NAME, oringinal_no, PRICE, ISNULL((SELECT CODE_NAME FROM TBL_CODEMASTER WHERE SITE_ID = A.SITE_ID AND CLASS_ID = 'R0038' AND CODE_ID = A.PART_DV),'')" & vbNewLine
            QRY = QRY & "FROM TBL_PARTMASTER A" & vbNewLine
            QRY = QRY & "WHERE SITE_ID = '" & Site_id & "'" & vbNewLine
            QRY = QRY & "  AND PART_NO LIKE '%" & TextBoxX1.Text & "%'" & vbNewLine
            QRY = QRY & "  and active = 'Y'" & vbNewLine
            QRY = QRY & "ORDER BY PART_NO" & vbNewLine

            If Query_Listview(PartList, QRY, True) = True Then
            End If

            If TextBoxX1.Text <> "" Then
                Dim item1 As ListViewItem = PartList.FindItemWithText(TextBoxX1.Text) '파트검색
                If (item1 IsNot Nothing) Then
                    'MsgBox(item1.ToString)
                    item1.EnsureVisible()  '해당위치로 스크롤 이동
                    item1.Selected = True
                    PartList.Focus()
                End If
            End If

        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message, "ERROR")
        End Try
    End Sub

    Private Sub PartList_DragDrop(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles PartList.DragDrop, FpSpread1.DragDrop

        If e.Data.GetDataPresent("System.Windows.Forms.ListViewItem", False) Then
            FpSpread_Fill(Me.PartList)
        End If
    End Sub

    Private Sub PartList_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles PartList.DoubleClick

        If SPREAD_DUP_CHECK(FpSpread1, PartList.SelectedItems(0).Text, 2) = False Then
            MessageBox.Show("Already exists.")
            Exit Sub
        End If
        FpSpread_Fill(Me.PartList)


        SaveBtn.Enabled = True
        SaveBtn1.Enabled = True

    End Sub


    Private Sub ButtonItem7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonItem7.Click, FpSpread2.CellDoubleClick

        FpSpread1.ActiveSheet.RowCount = 0

        FpSpread1.ActiveSheet.FrozenColumnCount = 3

        If FpSpread2.ActiveSheet.RowCount = 0 Then
            MessageBox.Show("출하번호를 선택하세요.")
            Exit Sub
        End If

        If FpSpread2.ActiveSheet.Rows(FpSpread2.ActiveSheet.ActiveRowIndex).ForeColor = Color.OrangeRed Then
            MessageBox.Show("상품 출하된 출하번호가 아닙니다.")
        Else
            MainFrm.ProgressBarItem1.ProgressType = DevComponents.DotNetBar.eProgressItemType.Marquee

            Dim QRY As String

            QRY = "select MODEL, ISNULL(RETURN_DV,'GOOD'), CASE WHEN (SELECT COUNT(MODEL) FROM TBL_FESNMASTER_B WHERE CSHIP_NO = A.SHIP_NO) > 0  THEN '생산출하' WHEN (SELECT COUNT(MODEL_NO) FROM TBL_MODELMASTER WHERE MODEL_NO = A.MODEL) > 0 THEN '제품'  ELSE '품목' END, CHARGE, QTY, CHARGE*QTY FROM VIEW_SHIPSUMMARY A" & vbNewLine
            QRY = QRY & "WHERE SHIP_NO = '" & FpSpread2.ActiveSheet.GetValue(FpSpread2.ActiveSheet.ActiveRowIndex, 1) & "'" & vbNewLine
            QRY = QRY & "ORDER BY MODEL" & vbNewLine


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
        Dim i, i_qty As New Integer

        'qry += "SELECT SUBSTRING(CSHIP_NO,3,8)+ ' ' + convert(varchar(5), (select max(U_date) from tbl_Fesnmaster_B where Cship_no = a.Cship_no),108), CSHIP_NO, COUNT(CSHIP_NO)," & vbNewLine
        'qry += "INV_NO ,'', MODEL, RETURN_DV" & vbNewLine
        'qry += "FROM TBL_Fesnmaster_B a" & vbNewLine
        'qry += "WHERE SUBSTRING(CSHIP_NO,3,8) BETWEEN CONVERT(VARCHAR(8), '" & DateTimeInput1.Text & "',112) AND  CONVERT(VARCHAR(8), '" & DateTimeInput2.Text & "',112)" & vbNewLine
        'If ComboBoxEx3.Text <> "ALL" Then
        '    qry += "  AND MODEL = '" & ComboBoxEx3.Text & "'" & vbNewLine
        'End If
        'If ComboBoxEx1.Text <> "ALL" Then
        '    qry += "  AND RETURN_DV = '" & ComboBoxEx1.Text & "'" & vbNewLine
        'End If
        'qry += "GROUP BY RIGHT(LEFT(CSHIP_NO,10),8), CSHIP_NO, INV_NO , MODEL, RETURN_DV" & vbNewLine
        'qry += "UNION" & vbNewLine
        qry += "SELECT SHIP_DATE, SHIP_NO, QTY," & vbNewLine
        qry += "INV_NO, AA, MODEL, RETURN_DV" & vbNewLine
        qry += "FROM VIEW_SHIPSUMMARY" & vbNewLine
        qry += "WHERE SHIP_DATE BETWEEN CONVERT(VARCHAR(8), '" & DateTimeInput1.Text & "',112) AND  CONVERT(VARCHAR(8), '" & DateTimeInput2.Text & "',112)" & vbNewLine
        If ComboBoxEx3.Text <> "ALL" Then
            qry += "  AND MODEL = '" & ComboBoxEx3.Text & "'" & vbNewLine
        End If
        If ComboBoxEx1.Text <> "ALL" Then
            qry += "  AND RETURN_DV = '" & ComboBoxEx1.Text & "'" & vbNewLine
        End If
        qry += "order BY  SHIP_NO" & vbNewLine


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
        qry += "SELECT MODEL, RETURN_DV, '', INBOX_NO, COUNT(MODEL)" & vbNewLine
        qry += "FROM TBL_FESNMASTER_K " & vbNewLine
        qry += "WHERE INBOX_NO IS NOT NULL" & vbNewLine
        qry += "and CSHIP_NO is null" & vbNewLine
        If ComboBoxEx3.Text <> "ALL" Then
            qry += "  AND MODEL = '" & ComboBoxEx3.Text & "'" & vbNewLine
        End If
        If ComboBoxEx4.Text <> "ALL" Then
            qry += "  AND RETURN_DV = '" & ComboBoxEx4.Text & "'" & vbNewLine
        End If
        qry += "GROUP BY MODEL, RETURN_DV, INBOX_NO" & vbNewLine
        qry += "order BY MODEL, RETURN_DV, INBOX_NO" & vbNewLine



    End Function


    Private Sub FindBtn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FindBtn.Click
        Dim i, i_qty As New Integer

        LabelX4.Text = ""
        FpSpread1.ActiveSheet.RowCount = 0
        FpSpread2.ActiveSheet.RowCount = 0
        Bar5.Visible = False


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
                Dim pmd1 As String = ""
                Dim pmd2 As String = ""
                Dim pmd3 As String = ""

                S3.Visible = True
                S4.Visible = True

                S3.Cells(3, 7).Value = FpSpread2.ActiveSheet.GetValue(FpSpread2.ActiveSheet.ActiveRowIndex, 1)
                S3.Cells(3, 12).Value = Query_RS("select CONVERT(VARCHAR(10), GETDATE(), 21)")
                S3.Cells(21, 4).Value = Query_RS("select CONVERT(VARCHAR(10), GETDATE(), 21)")

                S4.Cells(3, 7).Value = FpSpread2.ActiveSheet.GetValue(FpSpread2.ActiveSheet.ActiveRowIndex, 1)
                S4.Cells(3, 12).Value = Query_RS("select CONVERT(VARCHAR(10), GETDATE(), 21)")
                S4.Cells(21, 4).Value = Query_RS("select CONVERT(VARCHAR(10), GETDATE(), 21)")

                Dim QRY As String = ""
                'QRY = QRY & "select MODEL, count(MODEL), (SELECT RESERV6 FROM TBL_MODELMASTER WHERE MODEL_NO = A.MODEL) FROM TBL_FESNMASTER_B A" & vbNewLine
                'QRY = QRY & "WHERE CSHIP_NO = '" & S3.Cells(3, 7).Value & "' GROUP BY MODEL" & vbNewLine
                'QRY = QRY & "UNION" & vbNewLine
                QRY = QRY & "select MODEL, QTY, ISNULL((SELECT RESERV6 FROM TBL_MODELMASTER WHERE MODEL_NO = B.MODEL),0) FROM VIEW_SHIPSUMMARY B" & vbNewLine
                QRY = QRY & "WHERE SHIP_NO = '" & S3.Cells(3, 7).Value & "' " & vbNewLine
                QRY = QRY & "ORDER BY MODEL" & vbNewLine

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
                        S3.Cells(27 + I, 11).Value = Query_RS("SELECT ISNULL((SELECT CHARGE FROM TBL_SHIPSUMMARY WHERE SHIP_NO = '" & S3.Cells(3, 7).Value & "' AND MODEL = '" & PK_RS(0).Value & "'), (SELECT ISNULL(CHARGE,0) FROM TBL_SHIPSUMMARY_GOODS WHERE SHIP_NO = '" & S3.Cells(3, 7).Value & "' AND MODEL = '" & PK_RS(0).Value & "'))")
                        S3.Cells(27 + I, 13).Value = S3.Cells(27 + I, 8).Value * S3.Cells(27 + I, 11).Value


                        S4.Cells(27 + I, 4).Value = PK_RS(0).Value
                        S4.Cells(27 + I, 8).Value = PK_RS(1).Value
                        S4.Cells(27 + I, 10).Value = CDec(PK_RS(2).Value)
                        S4.Cells(27 + I, 12).Value = S4.Cells(27 + I, 8).Value * S4.Cells(27 + I, 10).Value

                        S4.Cells(27 + I, 9).Value = "EA"
                        S4.Cells(27 + I, 11).Value = "KG"
                        S4.Cells(27 + I, 13).Value = "KG"

                        PK_RS.MoveNext()
                    Next
                End If

                S3.Cells(35, 8).Formula = "SUM(I28:I35)"
                'S3.Cells(35, 11).Formula = "SUM(L28:L35)"
                S3.Cells(35, 13).Formula = "SUM(N28:N35)"

                S4.Cells(35, 8).Formula = "SUM(I28:I35)"
                S4.Cells(35, 10).Formula = "SUM(K28:K35)"
                S4.Cells(35, 12).Formula = "SUM(M28:M35)"


                S4.Cells(31, 2).Value = S4.Cells(35, 10).Value
                S4.Cells(32, 2).Value = S4.Cells(35, 12).Value

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

    Private Sub FpSpread_Fill(ByVal ListVw As ListView)
        'BOM의 신규파트 추가
        Try
            Dim i As Integer

            For i = 0 To ListVw.SelectedItems.Count - 1

                With Me.FpSpread1.ActiveSheet
                    .RowCount = .RowCount + 1
                    .SetValue(.RowCount - 1, 0, PartList.SelectedItems(0).Text)
                    .SetValue(.RowCount - 1, 1, ComboBoxEx1.Text)
                    .SetValue(.RowCount - 1, 2, ComboBoxEx4.Text)
                    .SetValue(.RowCount - 1, 3, Query_RS("SELECT ISNULL((SELECT PRICE FROM TBL_MODELLEVEL WHERE MODEL = '" & PartList.SelectedItems(0).Text & "' AND SERVICE_TYPE = 'L1'),0)"))
                    .SetValue(.RowCount - 1, 4, 0)
                    .SetValue(.RowCount - 1, 5, 0)
                    .Cells(.RowCount - 1, 3, .RowCount - 1, 5).Locked = False
                    .Rows(.RowCount - 1).ForeColor = Color.OrangeRed
                    Spread_AutoCol(FpSpread1)


                    '.RowCount += 1
                    '.Cells(.RowCount - 1, 0).Text = Site_id
                    '.Cells(.RowCount - 1, 1).Text = LabelX2.Text
                    '.Cells(.RowCount - 1, 2).Text = ListVw.SelectedItems(i).SubItems(0).Text
                    '.Cells(.RowCount - 1, 8).Text = 0
                    'Chg_ComboCell(FpSpread1, .RowCount - 1, 5, New String() {"Y", "N"})
                    'Chg_ComboCell(FpSpread1, .RowCount - 1, 6, New String() {"Y", "N"})
                    'Chg_ComboCell(FpSpread1, .RowCount - 1, 7, New String() {"Y", "N"})
                    'Chg_ComboCell(FpSpread1, .RowCount - 1, 9, New String() {"Y", "N"})
                    'Select Case ListVw.SelectedItems(i).SubItems(4).Text
                    '    Case "LEVEL 1&2" '코스메틱
                    '        .Cells(.RowCount - 1, 5).Text = "Y"
                    '        .Cells(.RowCount - 1, 7).Text = "N"
                    '        .Cells(.RowCount - 1, 12).Text = "N"
                    '    Case "Technician" '텍
                    '        .Cells(.RowCount - 1, 5).Text = "N"
                    '        .Cells(.RowCount - 1, 7).Text = "N"
                    '        .Cells(.RowCount - 1, 12).Text = "N"
                    '    Case "FOC" 'FOC
                    '        .Cells(.RowCount - 1, 5).Text = "Y"
                    '        .Cells(.RowCount - 1, 7).Text = "Y"
                    '        .Cells(.RowCount - 1, 12).Text = "N"
                    '    Case "ETC" '시설재
                    '        .Cells(.RowCount - 1, 5).Text = "N"
                    '        .Cells(.RowCount - 1, 7).Text = "N"
                    '        .Cells(.RowCount - 1, 12).Text = "Y"
                    'End Select
                    '.Cells(.RowCount - 1, 6).Text = "N"
                    '.Cells(.RowCount - 1, 9).Text = "N"
                    '.Cells(.RowCount - 1, 10).Text = Query_RS("select isnull(repair_cd,'') from tbl_partmaster where site_id = '" & Site_id & "' and part_no = '" & ListVw.SelectedItems(i).SubItems(0).Text & "'")
                    '.Cells(.RowCount - 1, 11).Text = Query_RS("select top def_cd from tbl_repset where site_id = '" & Site_id & "' and part_no = '" & ListVw.SelectedItems(i).SubItems(0).Text & "' and isnull(rep_cd,'') = '" & .Cells(.RowCount - 1, 10).Text & "'")
                    '.Cells(.RowCount - 1, 13).Text = "Y"
                    '.Cells(.RowCount - 1, 14).Text = Emp_No
                    '.Cells(.RowCount - 1, 15).Text = Now
                    '.Cells(.RowCount - 1, 16).Text = Emp_No
                    '.Cells(.RowCount - 1, 17).Text = Now
                    '.Cells(.RowCount - 1, 18).Text = ListVw.SelectedItems(i).SubItems(1).Text
                    '.Cells(.RowCount - 1, 19).Text = ListVw.SelectedItems(i).SubItems(2).Text
                    '.Cells(.RowCount - 1, 20).Text = ListVw.SelectedItems(i).SubItems(3).Text
                    '.Rows(.RowCount - 1).ForeColor = Color.OrangeRed
                    'For j = 1 To .ColumnCount - 8
                    '    .Cells(.RowCount - 1, j).Locked = False
                    'Next
                End With

            Next

            Spread_AutoCol(Me.FpSpread1)

            For i = ListVw.SelectedItems.Count - 1 To 0 Step -1 'BOM에 추가된 파트는 리스트에서 삭제
                ListVw.Items.Remove(ListVw.SelectedItems.Item(i))
            Next

            FpSpread1.ActiveSheet.SetActiveCell(FpSpread1.ActiveSheet.RowCount - 1, 0)
            FpSpread1.ShowActiveCell(FarPoint.Win.Spread.VerticalPosition.Bottom, FarPoint.Win.Spread.HorizontalPosition.Left) 'Activecell위치로 스크롤 자동이동

        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message, "ERROR")
        End Try
    End Sub

    Private Sub NewBtn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NewBtn.Click, NewBtn1.Click
        Dim i, i_qty As New Integer
        LabelX4.Text = ""


        If ComboBoxEx3.Text = "ALL" Then
            Modal_Error("모델을 선택하십시오.")
            Exit Sub
        End If

        If ComboBoxEx1.Text = "ALL" Then
            Modal_Error("상태를 선택하십시오.")
            Exit Sub
        End If

        With FpSpread1.ActiveSheet
            If FpSpread2.ActiveSheet.RowCount = 0 Then
                If MessageBox.Show("새로운 출하번호를 발행하시겠습니까?", "신규 출하번호 발행", MessageBoxButtons.YesNo) = Windows.Forms.DialogResult.Yes Then
                    .RowCount = 0

                    FpSpread2.ActiveSheet.RowCount = 1
                    FpSpread2.ActiveSheet.Rows(0).ForeColor = Color.OrangeRed

                    FpSpread2.ActiveSheet.SetValue(0, 0, Query_RS("select convert(varchar(8), getdate(), 112)"))

                    Dim FILESEQ As String
                    FILESEQ = Microsoft.VisualBasic.Right("00000" & CStr(CInt(Query_RS("SELECT CODE_NAME FROM TBL_CODEMASTER WHERE CLASS_ID = 'R0055' AND CODE_ID = 'A1000'")) + 1), 5)

                    FpSpread2.ActiveSheet.SetValue(0, 1, "EK" & Query_RS("select convert(varchar(8), getdate(), 112)") & "-" & FILESEQ)
                    FpSpread2.ActiveSheet.SetActiveCell(0, 0)
                Else
                End If


            Else
                If FpSpread2.ActiveSheet.RowCount = 1 And FpSpread2.ActiveSheet.Rows(0).ForeColor = Color.OrangeRed Then '출하번호 신규 발행한 경우
                Else
                    If .RowCount > 0 Then
                        If .Rows(0).ForeColor <> Color.OrangeRed Then
                            If MessageBox.Show("새로운 출하번호를 발행하시겠습니까?", "신규 출하번호 발행", MessageBoxButtons.YesNo) = Windows.Forms.DialogResult.Yes Then
                                .RowCount = 0

                                FpSpread2.ActiveSheet.RowCount = 1
                                FpSpread2.ActiveSheet.Rows(0).ForeColor = Color.OrangeRed

                                FpSpread2.ActiveSheet.SetValue(0, 0, Query_RS("select convert(varchar(8), getdate(), 112)"))

                                Dim FILESEQ As String
                                FILESEQ = Microsoft.VisualBasic.Right("00000" & CStr(CInt(Query_RS("SELECT CODE_NAME FROM TBL_CODEMASTER WHERE CLASS_ID = 'R0055' AND CODE_ID = 'A1000'")) + 1), 6)

                                FpSpread2.ActiveSheet.SetValue(0, 1, "EK" & Query_RS("select convert(varchar(8), getdate(), 112)") & "-" & FILESEQ)
                                FpSpread2.ActiveSheet.SetActiveCell(0, 0)
                            Else
                            End If
                        Else
                        End If
                    Else
                        If MessageBox.Show("새로운 출하번호를 발행하시겠습니까?", "신규 출하번호 발행", MessageBoxButtons.YesNo) = Windows.Forms.DialogResult.Yes Then
                            .RowCount = 0

                            FpSpread2.ActiveSheet.RowCount = 1
                            FpSpread2.ActiveSheet.Rows(0).ForeColor = Color.OrangeRed

                            FpSpread2.ActiveSheet.SetValue(0, 0, Query_RS("select convert(varchar(8), getdate(), 112)"))

                            Dim FILESEQ As String
                            FILESEQ = Microsoft.VisualBasic.Right("00000" & CStr(CInt(Query_RS("SELECT CODE_NAME FROM TBL_CODEMASTER WHERE CLASS_ID = 'R0055' AND CODE_ID = 'A1000'")) + 1), 6)

                            FpSpread2.ActiveSheet.SetValue(0, 1, "EK" & Query_RS("select convert(varchar(8), getdate(), 112)") & "-" & FILESEQ)
                            FpSpread2.ActiveSheet.SetActiveCell(0, 0)
                        Else
                        End If
                        
                    End If
                End If


            End If

        End With



        With FpSpread1.ActiveSheet

            If ComboBoxEx4.Text = "품목" Then
                SaveBtn.Enabled = True
                SaveBtn1.Enabled = True
                Exit Sub
            End If

            If ComboBoxEx4.Text = "제품" And SPREAD_DUP_CHECK(FpSpread1, ComboBoxEx3.Text, 2) = False Then
                Modal_Error("동일한 모델이 이미 추가되었습니다.")
                Exit Sub
            End If

            .RowCount = .RowCount + 1
            .SetValue(.RowCount - 1, 0, ComboBoxEx3.Text)
            .SetValue(.RowCount - 1, 1, ComboBoxEx1.Text)
            .SetValue(.RowCount - 1, 2, ComboBoxEx4.Text)
            .SetValue(.RowCount - 1, 3, Query_RS("SELECT ISNULL((SELECT PRICE FROM tbl_modellevel WHERE MODEL = '" & ComboBoxEx3.Text & "' AND SERVICE_TYPE = 'L1'), 0)"))
            .SetValue(.RowCount - 1, 4, 0)
            .SetValue(.RowCount - 1, 5, 0)
            .Cells(.RowCount - 1, 3, .RowCount - 1, 5).Locked = False
            .Rows(.RowCount - 1).ForeColor = Color.OrangeRed
            Spread_AutoCol(FpSpread1)
        End With


        SaveBtn.Enabled = True
        SaveBtn1.Enabled = True


    End Sub

    Private Sub SaveBtn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SaveBtn.Click, SaveBtn1.Click
        Dim i As New Integer

        Me.Cursor = Cursors.WaitCursor

        If MessageBox.Show("출하번호 " & FpSpread2.ActiveSheet.GetValue(FpSpread2.ActiveSheet.ActiveRowIndex, 1) & "로 상품출하 하시겠습니까?", "상품출하", MessageBoxButtons.YesNo) = Windows.Forms.DialogResult.Yes Then
            SaveBtn.Enabled = False
            MainFrm.ProgressBarItem1.Maximum = FpSpread1.ActiveSheet.RowCount

            If FpSpread2.ActiveSheet.RowCount = 1 And FpSpread2.ActiveSheet.Rows(0).ForeColor = Color.OrangeRed Then
                ' 기사용한 Shipping No 인지 Check
                If Query_RS("select count(Cship_no) from view_Fesnmaster where Cship_no = '" & FpSpread2.ActiveSheet.GetValue(0, 1) & "'") > 0 Then
                    System.Console.Beep(3000, 400)
                    System.Console.Beep(3000, 400)
                    Modal_Error("이미 사용된 출하번호 입니다. : " & FpSpread2.ActiveSheet.GetValue(0, 1) & "!!")

                    FindBtn_Click(sender, e)
                    Exit Sub
                End If

                If Query_RS("select count(ship_no) from TBL_SHIPSUMMARY_goods where ship_no = '" & FpSpread2.ActiveSheet.GetValue(0, 1) & "'") > 0 Then
                    System.Console.Beep(3000, 400)
                    System.Console.Beep(3000, 400)
                    Modal_Error("이미 사용된 출하번호 입니다. : " & FpSpread2.ActiveSheet.GetValue(0, 1) & "!!")

                    FindBtn_Click(sender, e)
                    Exit Sub
                End If
            End If


            Dim SHIP_NO As String = FpSpread2.ActiveSheet.GetValue(FpSpread2.ActiveSheet.ActiveRowIndex, 1)
            Dim SHIP_dt As String = Mid(FpSpread2.ActiveSheet.GetValue(FpSpread2.ActiveSheet.ActiveRowIndex, 0), 1, 8)

            With FpSpread1.ActiveSheet
                For i = 0 To FpSpread1.ActiveSheet.RowCount - 1

                    MainFrm.ProgressBarItem1.Value = i + 1

                    If .Rows(i).ForeColor = Color.OrangeRed Then

                        Dim qry1 As String = ""

                        qry1 = qry1 & "insert into tbl_shipsummary_goods values ('" & Site_id & "','" & SHIP_dt & "','" & SHIP_NO & "','"
                        qry1 = qry1 & .GetValue(i, 0) & "','"
                        qry1 = qry1 & .GetValue(i, 1) & "','"
                        qry1 = qry1 & .GetValue(i, 2) & "',"
                        qry1 = qry1 & .GetValue(i, 3) & ","
                        qry1 = qry1 & .GetValue(i, 4) & ","
                        qry1 = qry1 & .GetValue(i, 5) & ",'"
                        qry1 = qry1 & Emp_No & "', getdate(),'" & Emp_No & "', getdate())" & vbNewLine & vbNewLine

                        If Query_RS("SELECT isnull(QTY,0) FROM TBL_PARTINV WHERE WH_CD = 'W1000' AND PART_NO = '" & .GetValue(i, 0) & "'") >= .GetValue(i, 4) Then
                            qry1 = qry1 & "     UPDATE TBL_PARTINV SET QTY = QTY - " & .GetValue(i, 4) & " WHERE WH_CD = 'W1000' AND PART_NO = '" & .GetValue(i, 0) & "'" & vbNewLine
                            qry1 = qry1 & "     INSERT INTO TBL_PARTIO VALUES ('" & Site_id & "','" & .GetValue(i, 0) & "','" & FpSpread2.ActiveSheet.GetValue(0, 1) & "'," & .GetValue(i, 4) & ",'W1000','U2014-0001','상품출하:" & FpSpread2.ActiveSheet.GetValue(0, 1) & "','" & Emp_No & "', GETDATE(),'" & Emp_No & "',GETDATE())" & vbNewLine
                        Else
                            Modal_Error("자재창고에 수량이 부족합니다.")
                            Exit Sub
                        End If

                        If Insert_Data(qry1) = True Then
                            FpSpread1.ActiveSheet.Rows(i).ForeColor = Color.Black
                        End If
                    End If
                Next

                FpSpread1.ActiveSheet.RowCount = 0

            End With

            'MessageBox.Show("저장되었습니다")
            MainFrm.ProgressBarItem1.Value = 0

            Bar4.AutoHide = True

            Dim qry As String

            If FpSpread2.ActiveSheet.RowCount = 1 And FpSpread2.ActiveSheet.Rows(0).ForeColor = Color.OrangeRed Then
                qry = "UPDATE TBL_CODEMASTER SET CODE_NAME = CONVERT(INT,CODE_NAME) + 1 WHERE CLASS_ID = 'R0055' AND CODE_ID = 'A1000'" & vbNewLine & vbNewLine

                Insert_Data(qry)

            End If

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


    Private Sub PrtBtn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PrtBtn.Click, PrtBtn1.Click
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

            If PkRow < 16 Then
                s_cnt = 1
            ElseIf PkRow > 15 And PkRow < 31 Then
                s_cnt = 2
            ElseIf PkRow > 30 Then
                s_cnt = 3

            End If

            FpSpread3.ActiveSheet.ColumnHeaderVisible = False
            FpSpread3.ActiveSheet.RowHeaderVisible = False

            sheet_print2()

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

    Private Sub ButtonItem11_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Try


            Dim S1 = Me.FpSpread1.ActiveSheet
            Dim s2 = Me.FpSpread2.ActiveSheet
            Dim S4 = Me.FpSpread4.ActiveSheet
            Dim i As Integer
            If S1.RowCount < 1 Then
                Exit Sub
            End If

            If S4.RowCount > 0 Then
                S4.Rows.Clear()
            End If

            For i = 0 To S1.RowCount - 1
                S4.RowCount += 1
                S4.SetValue(i, 0, S1.GetValue(i, 2))
            Next

            SaveFileDialog1.FileName = s2.GetValue(s2.ActiveRowIndex, 1) & "-DMCHK"
            File_Save(SaveFileDialog1, FpSpread4)

        Catch ex As Exception
            MessageBox.Show("Error : " & ex.Message, "Error")
        End Try
    End Sub

    Private Sub TextBoxX2_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TextBoxX2.KeyDown



    End Sub


    Private Sub ButtonItem18_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonX4.Click

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

    Private Sub LabelX3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles LabelX3.Click

    End Sub

    Private Sub ComboBoxEx1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBoxEx1.SelectedIndexChanged

    End Sub

    Private Sub FpSpread1_Change(ByVal sender As Object, ByVal e As FarPoint.Win.Spread.ChangeEventArgs) Handles FpSpread1.Change

        With FpSpread1.ActiveSheet
            .SetValue(e.Row, 5, .GetValue(e.Row, 3) * .GetValue(e.Row, 4))
            Spread_AutoCol(FpSpread1)
        End With


    End Sub

    Private Sub ComboBoxEx4_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBoxEx4.SelectedIndexChanged

        If ComboBoxEx4.Text = "품목" Then
            Bar6.Visible = True
        Else
            PartList.Items.Clear()
            Bar6.Visible = False
        End If

    End Sub
End Class