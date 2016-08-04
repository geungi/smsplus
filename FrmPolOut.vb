Public Class FrmPolOut
    Const CtrlMask As Byte = 8
    Private NowPoNo As String = ""
    Private TempSno As String = ""
    Private TotRec As Integer = 0
    Private TotRec2 As Integer = 0
    Private FwhCd As String = ""
    Private TwhCd As String = ""
    Private EmpNmTXT As String = ""
    Private PartAuth As Boolean = False
    Private CosWH As String = ""

    Private s As String
    Private cr As FarPoint.Win.Spread.Model.CellRange

    Private Sub FrmPartTransfer_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        'If Partauth_yn = "Y" Then
        '    PartAuth = True
        'End If

        DockContainerItem2.Text = "조회 조건"

        If Spread_Setting(FpSpread2, Me.Name) = True Then
            FpSpread2.ActiveSheet.FrozenColumnCount = 3
            Spread_AutoCol(FpSpread2)
            FpSpread2.ActiveSheet.OperationMode = FarPoint.Win.Spread.OperationMode.Normal
        End If

        If Spread_Setting(FpSpread3, Me.Name) = True Then
            Spread_AutoCol(FpSpread3)
        End If


        If Query_Combo(Me.ComboBoxEx3, "SELECT model_no FROM tbl_modelmaster WHERE site_id = '" & Site_id & "' and active = 'Y' ORDER BY model_no") = True Then
        End If
        Me.ComboBoxEx3.Items.Add("ALL")
        Me.ComboBoxEx3.Text = "ALL"


        DateTimeInput1.Value = Now
        DateTimeInput2.Value = Now

        Formbtn_Authority(Me.NewBtn1, Me.Name, "NEW")
        Formbim_Authority(Me.SaveBtn, Me.Name, "SAVE")
        Formbtn_Authority(Me.SaveBtn1, Me.Name, "SAVE")
        Formbim_Authority(Me.DelBtn, Me.Name, "DELETE")
        Formbtn_Authority(Me.DelBtn1, Me.Name, "DELETE")
        Formbim_Authority(Me.PrtBtn, Me.Name, "PRINT")
        Formbtn_Authority(Me.PrtBtn1, Me.Name, "PRINT")
        Formbim_Authority(Me.Excel, Me.Name, "EXCEL")
        Formbtn_Authority(Me.XlsBtn1, Me.Name, "EXCEL")
        Formbim_Authority(Me.FindBtn, Me.Name, "FIND")

    End Sub


    Private Sub FindBtn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FindBtn.Click

        Dim qry As String = ""

        FpSpread2.ActiveSheet.RowCount = 0
        FpSpread3.ActiveSheet.RowCount = 0

        qry = qry & "select LOT_NO, MODEL, C_DATE, INIT_QTY, ACT_QTY  " & vbNewLine
        qry = qry & "from TBL_LOTMASTER A " & vbNewLine
        qry = qry & "WHERE C_PRC = 'K3000'" & vbNewLine
        If ComboBoxEx3.Text <> "ALL" Then
            qry = qry & "  AND MODEL = '" & ComboBoxEx3.Text & "'" & vbNewLine
        End If
        qry = qry & "AND LOT_NO NOT LIKE 'POL%'" & vbNewLine
        qry = qry & "AND C_DATE BETWEEN '" & DateTimeInput1.Text & "' AND '" & DateTimeInput2.Text & " 23:59:59'" & vbNewLine
        qry = qry & "order by LOT_NO" & vbNewLine

        If Query_Spread(FpSpread2, qry, 1) = True Then
            Spread_AutoCol(FpSpread2)
        End If


    End Sub

    Private Sub PtMdCb_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Try
            Me.FpSpread3.ActiveSheet.Rows.Clear()

            disp_part()
        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message, "ERROR")
        End Try
    End Sub

    Private Sub PtWhCb_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Try
            Me.FpSpread3.ActiveSheet.Rows.Clear()

            disp_part()
        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message, "ERROR")
        End Try
    End Sub

    Private Sub FpSpread2_CellDoublsClick(ByVal sender As Object, ByVal e As FarPoint.Win.Spread.CellClickEventArgs) Handles FpSpread2.CellDoubleClick
        If Me.FpSpread2.ActiveSheet.RowCount > 0 Then
            FpSpread_Fill(e.Row)
        End If
    End Sub

    Private Sub FpSpread_Fill(ByVal rowidx As Integer) '파트리스트에서 선택한 파트를 해당 스프레드시트에 넣고, 넣은 파트는 리스트에서 삭제
        Try
            Dim S2 = Me.FpSpread2.ActiveSheet
            Dim S3 = Me.FpSpread3.ActiveSheet

            If S3.RowCount > 0 Then
                If S3.Rows(S3.RowCount - 1).ForeColor = Color.Black Then
                    S3.Rows.Clear()
                End If
            End If

            'Dim WHary As String()
            Dim J As String() = SPREAD_SEARCH(Me.FpSpread3, 0, S2.Cells(rowidx, 0).Text, 0, 0, S2.RowCount - 1, 0, False)

            If J(0) >= 0 And J(1) >= 0 Then
                Modal_Error("이미 선택된 LOT 번호 입니다 !!!")
                Exit Sub
            End If

            S3.RowCount += 1
            S3.Cells(S3.RowCount - 1, 1).Locked = False
            S3.Cells(S3.RowCount - 1, 2).Locked = False
            S3.Cells(S3.RowCount - 1, 3).Locked = False

            S3.Cells(S3.RowCount - 1, 0).Text = S2.Cells(rowidx, 0).Text
            S3.Cells(S3.RowCount - 1, 1).Text = 0
            S3.Cells(S3.RowCount - 1, 2).Text = 0
            S3.Cells(S3.RowCount - 1, 3).Text = 0


            Me.FpSpread3.ActiveSheet.SetActiveCell(S3.RowCount - 1, 2)
            Me.FpSpread3.ShowActiveCell(S3.RowCount - 1, 2)
            Spread_Change(Me.FpSpread3, S3.RowCount - 1)
            Spread_AutoCol(Me.FpSpread3)
            '            S3.Columns(4).Width = S3.Columns(2).Width + 100

        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message, "ERROR")
        End Try
    End Sub


    Private Sub FpSpread3_CellClick(ByVal sender As Object, ByVal e As FarPoint.Win.Spread.CellClickEventArgs) Handles FpSpread3.CellClick
        Try
            save_excel = "FpSpread3"

            If Me.FpSpread3.ActiveSheet.RowCount > 0 Then
                If Me.FpSpread2.ActiveSheet.RowCount = 0 Then
                    MessageBox.Show("선택된 LOT가 없습니다!!", "validation Error")
                    Exit Sub
                End If
                Me.FpSpread3.ActiveSheet.Cells(e.Row, 1).Locked = False
                Me.FpSpread3.ActiveSheet.Cells(e.Row, 2).Locked = False
                Me.FpSpread3.ActiveSheet.Cells(e.Row, 3).Locked = False

                Dim J As String() = SPREAD_SEARCH(Me.FpSpread2, 0, Me.FpSpread3.ActiveSheet.Cells(e.Row, 0).Text, 0, 0, Me.FpSpread2.ActiveSheet.RowCount - 1, 0, False)
                If J(0) > -1 Then
                    Me.FpSpread2.ActiveSheet.SetActiveCell(J(0), 1)
                    Me.FpSpread2.ShowActiveCell(J(0), 1)
                End If
            End If
        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message, "ERROR")
        End Try
    End Sub

    Private Sub FpSpread3_CellDoubleClick(ByVal sender As Object, ByVal e As FarPoint.Win.Spread.CellClickEventArgs) Handles FpSpread3.CellDoubleClick
        Me.FpSpread3.ActiveSheet.Cells(e.Row, 1).Locked = False
        Me.FpSpread3.ActiveSheet.Cells(e.Row, 2).Locked = False
        Me.FpSpread3.ActiveSheet.Cells(e.Row, 3).Locked = False
        'Me.FpSpread3.ActiveSheet.Cells(e.Row, 4).Locked = False
        'Me.FpSpread3.ActiveSheet.Cells(e.Row, 5).Locked = False

    End Sub

    Private Sub FpSpread3_LeaveCell(ByVal sender As Object, ByVal e As FarPoint.Win.Spread.LeaveCellEventArgs) Handles FpSpread3.LeaveCell
        Try
            Dim S3 = Me.FpSpread3.ActiveSheet
            Dim S2 = Me.FpSpread2.ActiveSheet

            Select Case e.Column
                'Case 1
                '    If S3.Cells(e.Row, 1).Text <> "" Then
                '        Dim J As String() = SPREAD_SEARCH(Me.FpSpread2, 0, S3.Cells(e.Row, 0).Text, 0, 0, S2.RowCount - 1, 0, False)
                '        If CInt(S3.Cells(e.Row, 1).Text) > CInt(S2.Cells(J(0), 4).Text) Then
                '            MessageBox.Show("ERROR : [PartList's QTY] < [Transferrig's T_Qty] !!! ", "validation Error")
                '            FpSpread3.UndoManager.Undo(1)

                '            Exit Sub
                '        End If
                '        If Cal_Qty(Me.FpSpread3, 0, S3.Cells(e.Row, 0).Text, 1) > CInt(S2.Cells(J(0), 4).Text) Then
                '            MessageBox.Show("ERROR : [PartList's Qtyk] < [[Transferrig's Total T_Qty] !!! ", "validation Error")
                '            FpSpread3.UndoManager.Undo(1)
                '            Exit Sub
                '        End If
                '    Else
                '    End If
                'Case 3
                '    If S3.Cells(e.Row, 3).Text <> "" Then
                '        If S3.Cells(e.Row, 3).Text = S3.Cells(e.Row, 2).Text Then
                '            MessageBox.Show("ERROR : [From WH] = [To WH] !!! ", "validation Error")
                '            FpSpread3.UndoManager.Undo(1)
                '            'e.NewRow = e.Row
                '            'e.NewColumn = e.Column
                '        End If
                '    Else
                '        MessageBox.Show("[" & S3.ColumnHeader.Columns(3).Label & "]'s is Empty!!", "validation Error")
                '    End If

            End Select
        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message, "ERROR")
        End Try
    End Sub

    Private Sub FpSpread3_Change(ByVal sender As Object, ByVal e As FarPoint.Win.Spread.ChangeEventArgs) Handles FpSpread3.Change
        Try
            Spread_Change(sender, e.Row)
        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message, "ERROR")
        End Try
    End Sub

    Private Sub SaveBtn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SaveBtn1.Click, SaveBtn.Click, bSave.Click
        Try
            Dim i As Integer
            Dim S2 = Me.FpSpread2.ActiveSheet
            Dim S3 = Me.FpSpread3.ActiveSheet


            If S3.RowCount > 0 Then

                S3.SetActiveCell(0, 0)   '마지막 변경셀에서 다른 셀로 포커스 이동하지 않아도 저장되도록 자동 포커스 이동

                For i = 0 To S3.RowCount - 1
                    If S3.Rows(i).ForeColor <> Color.OrangeRed Then
                        Continue For
                    End If

                    If Query_RS("SELECT ACT_QTY FROM TBL_LOTMASTER WHERE LOT_NO = '" & S3.GetValue(i, 0) & "'") < S3.GetValue(i, 1) + S3.GetValue(i, 2) Then
                        Modal_Error("LOT 현재 수량을 초과합니다.")
                        Exit Sub
                    End If


                    'If S2.GetValue(SPREAD_DUP_ROW(FpSpread2, S3.GetValue(i, 0), 0), 7) < S3.GetValue(i, 2) + S3.GetValue(i, 3) Then
                    '    Modal_Error("검사 대기 수량을 초과합니다.")
                    '    Exit Sub
                    'End If

                    Dim QRY As String = "EXEC SP_FRMPOLOUT_LOTSAVE '" & Site_id & "','" & S3.GetValue(i, 0) & "'," & S3.GetValue(i, 1) & "," & S3.GetValue(i, 2) & ",'" & Emp_No & "'"

                    If Insert_Data(QRY) = True Then
                        S3.Rows(i).ForeColor = Color.Black
                    End If

                Next

                MessageBox.Show("저장되었습니다", "Message")

            End If
        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message, "ERROR")
        End Try
    End Sub

    Private Sub DelBtn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DelBtn1.Click, DelBtn.Click, bDel.Click
        Dim S3 = Me.FpSpread3.ActiveSheet

        Dim r As DialogResult = MessageBox.Show("Selected rows delete now?", "Selected Rows Delete", MessageBoxButtons.YesNo)
        If r = Windows.Forms.DialogResult.Yes Then
            S3.RemoveRows(S3.ActiveRowIndex, 1)
            MessageBox.Show("삭제되었습니다", "Message")
        End If

    End Sub

    Private Sub PrtBtn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PrtBtn1.Click, PrtBtn.Click, bPrint.Click
        Dim S3 = Me.FpSpread3.ActiveSheet
        If S3.RowCount > 0 Then


            If Spread_Print(Me.FpSpread3, "", 1) = False Then
                MsgBox("Fail to Print")
            End If
        End If
    End Sub

    Private Sub Excel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles XlsBtn1.Click, Excel.Click, bExcel.Click
        Me.FpSpread3.ActiveSheet.Columns(0).Visible = True
        If File_Save(SaveFileDialog1, FpSpread3) = True Then
            Me.FpSpread3.ActiveSheet.Columns(0).Visible = False
        End If
    End Sub
    Private Sub PrtBtn2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cPrint.Click
        Dim S2 = Me.FpSpread2.ActiveSheet
        If S2.RowCount > 0 Then
            If Spread_Print(Me.FpSpread2, "", 1) = False Then
                MsgBox("Fail to Print")
            End If
        End If
    End Sub

    Private Sub Excel2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cExcel.Click
        Me.FpSpread2.ActiveSheet.Columns(0).Visible = True
        If File_Save(SaveFileDialog1, FpSpread2) = True Then
            Me.FpSpread2.ActiveSheet.Columns(0).Visible = False
        End If
    End Sub

    Private Sub disp_part()
        Try
            Query_Spread(Me.FpSpread2, "", 1)

            Spread_AutoCol(FpSpread2)
        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message, "ERROR")
        End Try
    End Sub





    Private Sub FpSpread2_CellClick(ByVal sender As System.Object, ByVal e As FarPoint.Win.Spread.CellClickEventArgs) Handles FpSpread2.CellClick
        save_excel = "FpSpread2"
    End Sub


    Private Sub ButtonItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonItem1.Click
        Dim S3 = Me.FpSpread3.ActiveSheet

        If S3.RowCount > 0 Then
            S3.Rows.Clear()
        End If
    End Sub

    Private Sub FpSpread2_MouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles FpSpread2.MouseUp
        Try
            cr = FpSpread2.ActiveSheet.GetSelection(0)

            s = FpSpread2.ActiveSheet.GetClip(cr.Row, 0, cr.RowCount, 1)
        Catch ex As Exception
            'MessageBox.Show("You didn't make a selection!!")
            Return
        End Try
    End Sub

    Private Sub ButtonItem2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonItem2.Click
        Dim S2 = Me.FpSpread2.ActiveSheet
        Dim S3 = Me.FpSpread3.ActiveSheet
        Dim i As Integer = 0

        If cr.RowCount < 1 Then
            Exit Sub
        End If

        For i = cr.Row To cr.Row + cr.RowCount - 1
            FpSpread_Fill(i)
        Next

        cr = Nothing
    End Sub




End Class