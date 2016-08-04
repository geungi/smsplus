Public Class FrmPartRtn
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

        Condi_Disp() '콤보박스의 조건데이터 출력

        If Spread_Setting(FpSpread1, "FrmPartTransfer") = True Then
            Spread_AutoCol(FpSpread1)
        End If

        If Spread_Setting(FpSpread2, "FrmPartTransfer") = True Then
            FpSpread2.ActiveSheet.FrozenColumnCount = 1
            Spread_AutoCol(FpSpread2)
            FpSpread2.ActiveSheet.OperationMode = FarPoint.Win.Spread.OperationMode.Normal
        End If

        If Spread_Setting(FpSpread3, "FrmPartTransfer") = True Then
            Spread_AutoCol(FpSpread3)
        End If

        Me.ContextMenuBar1.SetContextMenuEx(Me.FpSpread3, CtxSp)
        Me.ContextMenuBar1.SetContextMenuEx(Me.FpSpread2, CtxSp2)

        Me.Bar2.AutoHide = True
        Me.Bar3.AutoHide = True

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

    Private Sub Condi_Disp() 'CONTROL PANEL 및 파트리스트뷰에 있는 콤보박스에 데이터 출력

        Me.POStDate.Value = Now
        Me.POEdDate.Value = Now

        'Part List의 model
        Me.PtMdCb.Text = "ALL"
        Query_Combo(Me.PtMdCb, "select model_no from tbl_modelmaster where site_id = '" & Site_id & "' and active='Y' ORDER BY model_no")
        Me.PtMdCb.Items.Add("ALL")

        'From WH
        Query_WHCombo(Me.FromCb, Site_id, True)
        'To WH
        Query_WHCombo(Me.ToCb, Site_id, True)
        Query_WHCombo(Me.PtWhCb, Site_id, True)

    End Sub

    Private Sub FindBtn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FindBtn.Click

        disp_io()

    End Sub

    Private Sub PtMdCb_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PtMdCb.SelectedIndexChanged
        Try
            Me.FpSpread3.ActiveSheet.Rows.Clear()

            disp_part()
        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message, "ERROR")
        End Try
    End Sub

    Private Sub PtWhCb_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PtWhCb.SelectedIndexChanged
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

            If Me.PtWhCb.Text <> "" Then

                'Dim WHary As String()
                Dim J As String() = SPREAD_SEARCH(Me.FpSpread3, 0, S2.Cells(rowidx, 0).Text, 0, 0, S2.RowCount - 1, 0, False)

                If J(0) >= 0 And J(1) >= 0 Then
                    MessageBox.Show("이미 선택된 품목입니다 !!!")
                    Exit Sub
                End If

                If CInt(S2.Cells(rowidx, 4).Text) > 0 Then
                    S3.RowCount += 1
                    S3.Cells(S3.RowCount - 1, 1).Locked = False
                    S3.Cells(S3.RowCount - 1, 3).Locked = False
                    S3.Cells(S3.RowCount - 1, 4).Locked = False

                    S3.Cells(S3.RowCount - 1, 0).Text = S2.Cells(rowidx, 0).Text
                    If Site_id = "S1000" Then
                        S3.Cells(S3.RowCount - 1, 1).Text = S2.Cells(rowidx, 4).Text
                    End If
                    '                    S3.Cells(S3.RowCount - 1, 2).Text = S2.Cells(rowidx, 5).Text
                    S3.Cells(S3.RowCount - 1, 2).Text = Me.PtWhCb.Text

                    Chg_ComboCell(FpSpread3, S3.RowCount - 1, 3, Query_Cell_Code2("select '['+sup_no+'] ' +sup_nm from tbl_supmst where sup_no like 'S%' union select '[S2014-0000] 자체처리'"))


                    Me.FpSpread3.ActiveSheet.SetActiveCell(S3.RowCount - 1, 2)
                    Me.FpSpread3.ShowActiveCell(S3.RowCount - 1, 4)
                    Spread_Change(Me.FpSpread3, S3.RowCount - 1)
                    Spread_AutoCol(Me.FpSpread3)
                    S3.Columns(3).Width = S3.Columns(2).Width + 10
                Else
                    MessageBox.Show("수량 < 1 !!", "validation Error")
                End If
            Else
                MessageBox.Show("타입을 선택하세요!!", "validation Error")
            End If
        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message, "ERROR")
        End Try
    End Sub


    Private Sub FpSpread3_CellClick(ByVal sender As Object, ByVal e As FarPoint.Win.Spread.CellClickEventArgs) Handles FpSpread3.CellClick
        Try
            save_excel = "FpSpread3"

            If Me.FpSpread3.ActiveSheet.RowCount > 0 Then
                If Me.FpSpread2.ActiveSheet.RowCount = 0 Then
                    MessageBox.Show("품목리스트가 없습니다!!", "validation Error")
                    Exit Sub
                End If
                Me.FpSpread3.ActiveSheet.Cells(e.Row, 1).Locked = False
                Me.FpSpread3.ActiveSheet.Cells(e.Row, 2).Locked = False
                Me.FpSpread3.ActiveSheet.Cells(e.Row, 3).Locked = False
                Me.FpSpread3.ActiveSheet.Cells(e.Row, 4).Locked = False

                Dim J As String() = SPREAD_SEARCH(Me.FpSpread2, 0, Me.FpSpread3.ActiveSheet.Cells(e.Row, 0).Text, 0, 0, Me.FpSpread2.ActiveSheet.RowCount - 1, 0, False)
                If J(0) > -1 Then
                    Me.FpSpread2.ActiveSheet.SetActiveCell(J(0), 4)
                    Me.FpSpread2.ShowActiveCell(J(0), 4)
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
        Me.FpSpread3.ActiveSheet.Cells(e.Row, 4).Locked = False

    End Sub

    Private Sub FpSpread3_LeaveCell(ByVal sender As Object, ByVal e As FarPoint.Win.Spread.LeaveCellEventArgs) Handles FpSpread3.LeaveCell
        Try
            Dim S3 = Me.FpSpread3.ActiveSheet
            Dim S2 = Me.FpSpread2.ActiveSheet

            Select Case e.Column
                Case 1
                    If S3.Cells(e.Row, 1).Text <> "" Then
                        Dim J As String() = SPREAD_SEARCH(Me.FpSpread2, 0, S3.Cells(e.Row, 0).Text, 0, 0, S2.RowCount - 1, 0, False)
                        If CInt(S3.Cells(e.Row, 1).Text) > CInt(S2.Cells(J(0), 4).Text) Then
                            MessageBox.Show("ERROR : [PartList's QTY] < [Transferrig's T_Qty] !!! ", "validation Error")
                            FpSpread3.UndoManager.Undo(1)

                            Exit Sub
                        End If
                        If Cal_Qty(Me.FpSpread3, 0, S3.Cells(e.Row, 0).Text, 1) > CInt(S2.Cells(J(0), 4).Text) Then
                            MessageBox.Show("ERROR : [PartList's Qtyk] < [[Transferrig's Total T_Qty] !!! ", "validation Error")
                            FpSpread3.UndoManager.Undo(1)
                            Exit Sub
                        End If
                    Else
                    End If
                Case 3
                    If S3.Cells(e.Row, 3).Text <> "" Then
                        If S3.Cells(e.Row, 3).Text = S3.Cells(e.Row, 2).Text Then
                            MessageBox.Show("ERROR : [From WH] = [To WH] !!! ", "validation Error")
                            FpSpread3.UndoManager.Undo(1)
                            'e.NewRow = e.Row
                            'e.NewColumn = e.Column
                        End If
                    Else
                        MessageBox.Show("[" & S3.ColumnHeader.Columns(3).Label & "]'s is Empty!!", "validation Error")
                    End If

            End Select
        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message, "ERROR")
        End Try
    End Sub

    Private Sub FpSpread3_Change(ByVal sender As Object, ByVal e As FarPoint.Win.Spread.ChangeEventArgs) Handles FpSpread3.Change
        Try
            Spread_Change(sender, e.Row)
            Spread_AutoCol(FpSpread3)
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


                If Save_Vallidate(Me.FpSpread3, "PART TRANSFER", New String() {1, 3}) = False Then
                    Exit Sub
                End If

                For i = 0 To S3.RowCount - 1
                    Dim J As String() = SPREAD_SEARCH(Me.FpSpread2, 0, S3.Cells(i, 0).Text, 0, 0, S2.RowCount - 1, 0, False)
                    If J(0) = -1 Then
                        MessageBox.Show("[" & S3.Cells(i, 0).Text & "] is not exist!!", "Validation Error")
                        Exit Sub
                    End If
                    If S3.Cells(i, 3).Text = S3.Cells(i, 2).Text Then
                        MessageBox.Show("ERROR : [From WH] = [To WH] !!! ", "validation Error")
                        FpSpread3.UndoManager.Undo(1)
                        'e.NewRow = e.Row
                        'e.NewColumn = e.Column
                    End If
                    If CInt(S3.Cells(i, 1).Text) > CInt(S2.Cells(J(0), 4).Text) Then
                        MessageBox.Show("ERROR : [PartList's QTY] < [Transferrig's T_Qty] !!! ", "validation Error")
                        S3.SetActiveCell(i, 1)
                        Exit Sub
                    End If
                    If Cal_Qty(Me.FpSpread3, 0, S3.Cells(i, 0).Text, 1) > CInt(S2.Cells(J(0), 4).Text) Then
                        MessageBox.Show("ERROR : [PartList's Qty] < [[Transferrig's Total T_Qty] !!! ", "validation Error")
                        S3.SetActiveCell(i, 1)
                        Exit Sub
                    End If

                    If S3.Rows(i).ForeColor = Color.OrangeRed Then
                        '@SITE_ID,@PART_NO,@S_NO,@QTY,@FROM_WAREHOUSE,@TO_WHAREHOUSE,@REMARK,@C_PERSON
                        If Insert_Data("EXEC SP_FRMPARTRTN_INSINV '" & Site_id & "','" & S3.Cells(i, 0).Text & "','자재반품_" & Mid(S3.Cells(i, 2).Text, 2, 5) & Mid(S3.Cells(i, 3).Text, 2, 10) & "'," & S3.Cells(i, 1).Value & ", '" & Mid(S3.Cells(i, 2).Text, 2, 5) & "','" & Mid(S3.Cells(i, 3).Text, 2, 10) & "','" & S3.Cells(i, 4).Text & "','" & Emp_No & "'") = False Then
                            MessageBox.Show("Failed to Save")
                            Exit Sub
                        End If
                        S3.Rows(i).ForeColor = Color.Black
                    End If
                Next

                disp_part()
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


            If Spread_Print(Me.FpSpread3, Me.PtWhCb.Text & "'s Part Transfer", 1) = False Then
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
            If Spread_Print(Me.FpSpread2, Me.PtWhCb.Text & "'s Part list", 1) = False Then
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
            Query_Spread(Me.FpSpread2, "exec SP_FRMPARTTRANSFER_PARTLIST3 '" & Site_id & "','" & Me.PtMdCb.Text & "','" & Mid(Me.PtWhCb.Text, 2, 5) & "'", 1)

            Spread_AutoCol(FpSpread2)
        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message, "ERROR")
        End Try
    End Sub


    Private Sub disp_io()
        Try
            Dim QueryPo As String = ""
            Dim i As Integer

            Me.Bar3.AutoHide = False

            'PART I/O 출력
            '@SITE VARCHAR(10),
            '@FROM_WAREHOUSE VARCHAR(25),
            '@TO_WAREHOUSE VARCHAR(25),
            '@ST_DT DATETIME,
            '@ED_DT DATETIME,
            '@PART_NO	VARCHAR(50)
            'QueryPo = "EXEC SP_PART_IOLIST2 '" & Site_id & "', '" & Me.FromCb.SelectedValue.ToString & "', '" & Me.ToCb.SelectedValue.ToString & "', '" & Me.POStDate.Text & "','00:00:00','" & Me.POEdDate.Text & "','23:59:59','" & Me.PartNoTxt.Text & "',''"

            QueryPo = "SELECT C_DATE, PART_NO," & vbNewLine
            QueryPo = QueryPo & "	  isnull((SELECT '['+ CODE_ID +'] '+CODE_NAME FROM TBL_CODEMASTER WHERE SITE_ID='" & Site_id & "' AND CLASS_ID = 'R0003' AND CODE_ID = A.F_WH),(SELECT '['+ CODE_ID +'] '+CODE_NAME FROM TBL_CODEMASTER WHERE SITE_ID='" & Site_id & "' AND CLASS_ID = '10024' AND CODE_ID = A.F_WH))," & vbNewLine
            QueryPo = QueryPo & "	  (SELECT '['+ CODE_ID +'] '+CODE_NAME FROM TBL_CODEMASTER WHERE SITE_ID='" & Site_id & "' AND CLASS_ID = '10024' AND CODE_ID = A.T_WH)," & vbNewLine
            QueryPo = QueryPo & "	   QTY, SRC_NO," & vbNewLine
            QueryPo = QueryPo & "	   (SELECT MAX(S_NO) FROM TBL_PARTRCV WHERE SITE_ID = A.SITE_ID AND PO_NO = A.SRC_NO AND PART_NO = A.PART_NO)," & vbNewLine
            QueryPo = QueryPo & "	   REMARK,(SELECT EMP_NM FROM TBL_EMPMASTER WHERE SITE_ID = A.SITE_ID AND EMP_NO = A.C_PERSON) " & vbNewLine
            QueryPo = QueryPo & "FROM TBL_PARTIO A" & vbNewLine
            QueryPo = QueryPo & "WHERE SITE_ID = '" & Site_id & "'" & vbNewLine
            If FromCb.Text <> "ALL" Then
                QueryPo = QueryPo & "	 AND F_WH LIKE '%" & Mid(FromCb.Text, 2, 5) & "%'" & vbNewLine
            End If
            If ToCb.Text <> "ALL" Then
                QueryPo = QueryPo & "	 AND T_WH LIKE '%" & Mid(ToCb.Text, 2, 5) & "%'" & vbNewLine
            End If
            QueryPo = QueryPo & "	 AND C_DATE BETWEEN '" & POStDate.Text & "' AND '" & POEdDate.Text & " 23:59:59'" & vbNewLine
            QueryPo = QueryPo & "	 AND PART_NO LIKE '%" & PartNoTxt.Text & "%'" & vbNewLine
            QueryPo = QueryPo & "ORDER BY C_DATE DESC" & vbNewLine


            If Query_Spread(Me.FpSpread1, QueryPo, 1) = True Then
                For i = 0 To Me.FpSpread1.ActiveSheet.RowCount - 1
                    Me.FpSpread1.ActiveSheet.Rows(i).Locked = False
                Next
            End If
            Spread_AutoCol(Me.FpSpread1)

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