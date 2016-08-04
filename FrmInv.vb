Public Class FrmInv
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

        DockContainerItem2.Text = "조회 조건"

        If Spread_Setting(FpSpread2, Me.Name) = True Then
            FpSpread2.ActiveSheet.FrozenColumnCount = 1
            FpSpread2.ActiveSheet.RowCount = 0
            Spread_AutoCol(FpSpread2)
            FpSpread2.ActiveSheet.OperationMode = FarPoint.Win.Spread.OperationMode.Normal
        End If

        If Spread_Setting(FpSpread3, Me.Name) = True Then
            FpSpread3.ActiveSheet.RowCount = 0
            Spread_AutoCol(FpSpread3)
        End If

        If Spread_Setting(FpSpread1, Me.Name) = True Then
            FpSpread1.ActiveSheet.RowCount = 0
            Spread_AutoCol(FpSpread1)
        End If

        'If Query_Combo(Me.ComboBoxEx1, "SELECT model_no FROM tbl_modelmaster WHERE site_id = '" & Site_id & "' and active = 'Y' ORDER BY model_no") = True Then
        'End If
        'Me.ComboBoxEx1.Items.Add("ALL")
        'Me.ComboBoxEx1.Text = "ALL"
        If Query_CheckList(Me.CheckedListBox1, "SELECT model_no FROM tbl_modelmaster WHERE site_id = '" & Site_id & "' and active = 'Y' ORDER BY model_no") = True Then
            Me.CheckedListBox1.Items.Add("ALL")
            Me.CheckedListBox1.Text = "ALL"
        End If

        If Query_Combo(Me.ComboBoxEx2, "SELECT CODE_ID FROM tbl_CODEmaster WHERE site_id = '" & Site_id & "' and CLASS_ID = 'R0007' AND active = 'Y' ORDER BY DIS_ORDER") = True Then
        End If
        Me.ComboBoxEx2.Items.Add("ALL")
        Me.ComboBoxEx2.Text = "ALL"

        DateTimeInput1.Value = Now
        DateTimeInput2.Value = Now
        DateTimeInput4.Value = Now

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


        If Query_RS("select isnull(insa_yn,'N') from tbl_empmaster where emp_no = '" & Emp_No & "'") = "N" Then
            FpSpread1.ActiveSheet.Columns(4).Visible = False
            FpSpread2.ActiveSheet.Columns(3).Visible = False
            FpSpread2.ActiveSheet.Columns(6).Visible = False
            FpSpread2.ActiveSheet.Columns(7).Visible = False
            FpSpread3.ActiveSheet.Columns(3).Visible = False
            FpSpread3.ActiveSheet.Columns(7).Visible = False
        End If


    End Sub

    Private Sub FindBtn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FindBtn.Click

        FpSpread2.ActiveSheet.RowCount = 0
        FpSpread3.ActiveSheet.RowCount = 0
        FpSpread1.ActiveSheet.RowCount = 0
        Dim m_qry As String = ""
        m_qry = MODEL_SELECTED(CheckedListBox1)

        With FpSpread2.ActiveSheet

            curcell.CurrencySymbol = "$"

            .RowCount = 0

            Dim QRY As String = ""

            QRY = QRY & "SELECT SHIP_DATE, SHIP_NO, SUM(QTY) , SUM(QTY*CHARGE) , RETURN_DV , (SELECT COUNT(SHIP_NO) FROM TBL_INVSUMMARY WHERE SHIP_NO = A.SHIP_NO), (SELECT SUM(INV_CHARGE) FROM TBL_INVSUMMARY WHERE SHIP_NO = A.SHIP_NO), (SELECT SUM(BILL_CHARGE) FROM TBL_BILLING WHERE SHIP_NO = A.SHIP_NO)" & vbNewLine
            QRY = QRY & "FROM VIEW_SHIPSUMMARY A" & vbNewLine
            QRY = QRY & "WHERE SHIP_DATE BETWEEN CONVERT(VARCHAR(8), CONVERT(DATETIME,'" & DateTimeInput1.Text & "'),112) AND CONVERT(VARCHAR(8), CONVERT(DATETIME,'" & DateTimeInput2.Text & "'),112)" & vbNewLine
            'If ComboBoxEx1.Text <> "ALL" Then
            '    QRY = QRY & "AND MODEL = '" & ComboBoxEx1.Text & "'" & vbNewLine
            'End If
            If m_qry <> "" Then
                QRY = QRY & "AND MODEL IN (" & m_qry & ")" & vbNewLine
            End If

            If ComboBoxEx2.Text <> "ALL" Then
                QRY = QRY & "AND RETURN_DV = '" & ComboBoxEx2.Text & "'" & vbNewLine
            End If
            QRY = QRY & "GROUP BY SHIP_NO, SHIP_DATE , RETURN_DV " & vbNewLine
            QRY = QRY & "ORDER BY SHIP_NO, SHIP_DATE , RETURN_DV " & vbNewLine
            QRY = QRY & "" & vbNewLine

            If Query_Spread(FpSpread2, QRY, 1) = True Then
                .RowCount = FpSpread2.ActiveSheet.RowCount + 1
                .Rows(FpSpread2.ActiveSheet.RowCount - 1).BackColor = Color.Yellow
                .SetValue(FpSpread2.ActiveSheet.RowCount - 1, 0, "TOTAL")

                FpSpread2.AllowUserFormulas = True

                .Cells(.RowCount - 1, 2).Formula = "SUM(C1:C" & .RowCount - 1 & ")"
                .Cells(.RowCount - 1, 3).Formula = "SUM(D1:D" & .RowCount - 1 & ")"
                .Cells(.RowCount - 1, 6).Formula = "SUM(G1:G" & .RowCount - 1 & ")"
                .Cells(.RowCount - 1, 7).Formula = "SUM(H1:H" & .RowCount - 1 & ")"

                .ColumnFooterVisible = True
                .ColumnFooter.Cells(0, 0).Text = "출하 합계"
                .ColumnFooter.Cells(0, 1).CellType = intcell
                .ColumnFooter.Cells(0, 1).Text = .GetValue(.RowCount - 1, 2)
                .ColumnFooter.Cells(0, 2).CellType = curcell
                .ColumnFooter.Cells(0, 2).Text = .GetValue(.RowCount - 1, 3)

                .ColumnFooter.Cells(0, 4).ColumnSpan = 2
                .ColumnFooter.Cells(0, 4).Text = "청구/수금 합계"

                .ColumnFooter.Cells(0, 6).CellType = curcell
                .ColumnFooter.Cells(0, 6).Text = .GetValue(.RowCount - 1, 6)
                .ColumnFooter.Cells(0, 7).CellType = curcell
                .ColumnFooter.Cells(0, 7).Text = .GetValue(.RowCount - 1, 7)

                Spread_AutoCol(FpSpread2)
            End If

        End With
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

            FpSpread1.ActiveSheet.RowCount = 0


            If S3.RowCount > 0 Then
                If S3.Rows(S3.RowCount - 1).ForeColor = Color.Black Then
                    S3.Rows.Clear()
                End If
                Dim J As String() = SPREAD_SEARCH(Me.FpSpread3, 0, S2.Cells(rowidx, 1).Text, 0, 1, S3.RowCount - 1, 1, False)
                If J(0) >= 0 And J(1) >= 0 Then
                    Modal_Error("이미 선택된 출하건입니다 !!!")
                    Exit Sub
                End If
            End If

            Dim qry As String = "SELECT DISTINCT A.SHIP_DATE, A.SHIP_NO, sum(A.QTY), SUM(A.QTY*A.CHARGE), A.RETURN_DV,  B.INV_NO, B.INV_DATE, B.INV_CHARGE, B.INV_RMK "
            qry = qry & "FROM VIEW_SHIPSUMMARY A, TBL_INVSUMMARY B " & vbNewLine
            qry = qry & "WHERE A.SHIP_NO = '" & S2.GetValue(S2.ActiveRowIndex, 1) & "'" & vbNewLine
            qry = qry & "AND A.SHIP_NO = B.SHIP_NO" & vbNewLine
            qry = qry & "GROUP BY A.SHIP_DATE, A.SHIP_NO, A.RETURN_DV,  B.INV_NO, B.INV_DATE, B.INV_CHARGE, B.INV_RMK" & vbNewLine
            qry = qry & "ORDER BY b.inv_no" & vbNewLine

            If Query_Spread_LTD_ROW(FpSpread3, qry, 0, S3.ColumnCount - 1, S3.RowCount) = True Then
                Spread_AutoCol(FpSpread3)
            End If


            If DateTimeInput4.Text = "" Then
                Modal_Error("인보이스 일자를 선택하십시오 !!!")
                Exit Sub
            End If



            If S2.GetValue(S2.ActiveRowIndex, 3) > S2.GetValue(S2.ActiveRowIndex, 6) Then

                If TextBoxX1.Text = "" Then
                    Modal_Error("인보이스 번호를 입력하십시오 !!!")
                    Exit Sub
                End If

                S3.RowCount += 1
                S3.Cells(S3.RowCount - 1, 7, S3.RowCount - 1, 8).Locked = False

                S3.Cells(S3.RowCount - 1, 0).Text = S2.Cells(rowidx, 0).Text
                S3.Cells(S3.RowCount - 1, 1).Text = S2.Cells(rowidx, 1).Text
                S3.Cells(S3.RowCount - 1, 2).Text = S2.Cells(rowidx, 2).Text
                S3.Cells(S3.RowCount - 1, 3).Text = S2.Cells(rowidx, 3).Text
                S3.Cells(S3.RowCount - 1, 4).Text = S2.Cells(rowidx, 4).Text

                S3.Cells(S3.RowCount - 1, 5).Text = TextBoxX1.Text
                S3.Cells(S3.RowCount - 1, 6).Text = DateTimeInput4.Text

                S3.Cells(S3.RowCount - 1, 7).Text = S2.Cells(rowidx, 3).Text - Query_RS("SELECT ISNULL(SUM(INV_CHARGE),0) FROM TBL_INVSUMMARY WHERE SHIP_NO = '" & S3.GetValue(S3.RowCount - 1, 1) & "'")


                Me.FpSpread3.ActiveSheet.SetActiveCell(S3.RowCount - 1, 7)
                Me.FpSpread3.ShowActiveCell(S3.RowCount - 1, 7)
                Spread_Change(Me.FpSpread3, S3.RowCount - 1)
                Spread_AutoCol(Me.FpSpread3)

                'Else
                '    Modal_Error("이미 안보이스된 출하건입니다 !!!")
                '    Exit Sub
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
                Me.FpSpread3.ActiveSheet.Cells(e.Row, 7, e.Row, 8).Locked = False

            End If
        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message, "ERROR")
        End Try
    End Sub

    Private Sub FpSpread3_CellDoubleClick(ByVal sender As Object, ByVal e As FarPoint.Win.Spread.CellClickEventArgs) Handles FpSpread3.CellDoubleClick

        Dim S3 = Me.FpSpread3.ActiveSheet
        Dim S1 = Me.FpSpread1.ActiveSheet

        FpSpread1.ActiveSheet.RowCount = 0

        With FpSpread3.ActiveSheet
            If .RowCount = 0 Then
                Exit Sub
            End If


            If e.Column = 7 Or e.Column = 8 Then
                Me.FpSpread3.ActiveSheet.Cells(e.Row, 7, e.Row, 8).Locked = False
            End If

            Dim qry As String = "SELECT DISTINCT A.SHIP_NO, a.inv_no, b.bill_seq, B.BILL_DATE, b.bill_CHARGE, B.BILL_RMK "
            qry = qry & "FROM TBL_INVSUMMARY A, TBL_BILLING B " & vbNewLine
            qry = qry & "WHERE A.SHIP_NO = '" & .GetValue(.ActiveRowIndex, 1) & "'" & vbNewLine
            qry = qry & "AND A.INV_NO = '" & .GetValue(.ActiveRowIndex, 5) & "'" & vbNewLine
            qry = qry & "AND A.SHIP_NO = B.SHIP_NO" & vbNewLine
            qry = qry & "AND A.INV_NO = B.INV_NO" & vbNewLine
            qry = qry & "GROUP BY A.SHIP_NO, A.INV_NO, B.BILL_DATE, B.BILL_CHARGE, B.BILL_RMK, B.BILL_SEQ" & vbNewLine
            qry = qry & "ORDER BY A.SHIP_NO, A.INV_NO, B.BILL_DATE, B.BILL_SEQ" & vbNewLine

            If Query_Spread(FpSpread1, qry, 1) = True Then
                Spread_AutoCol(FpSpread1)
            End If

            'Dim AA As Decimal = .GetValue(e.Row, 7)

            If CDec(.GetValue(e.Row, 7)) > CDec(Query_RS("SELECT ISNULL(SUM(BILL_CHARGE),0) FROM TBL_BILLING WHERE SHIP_NO = '" & .GetValue(e.Row, 1) & "' AND INV_NO = '" & .GetValue(e.Row, 5) & "'")) Then

                S1.RowCount += 1
                S1.Cells(S1.RowCount - 1, 3, S1.RowCount - 1, 5).Locked = False

                S1.Cells(S1.RowCount - 1, 0).Text = .Cells(.ActiveRowIndex, 1).Text
                S1.Cells(S1.RowCount - 1, 1).Text = .Cells(.ActiveRowIndex, 5).Text
                S1.Cells(S1.RowCount - 1, 2).Text = S1.RowCount
                S1.Cells(S1.RowCount - 1, 3).Text = Now
                S1.Cells(S1.RowCount - 1, 4).Text = .GetValue(e.Row, 7) - Query_RS("SELECT ISNULL(SUM(BILL_CHARGE),0) FROM TBL_BILLING WHERE SHIP_NO = '" & .GetValue(e.Row, 1) & "' AND INV_NO = '" & .GetValue(e.Row, 5) & "'")

                S1.Cells(S1.RowCount - 1, 5).Text = ""


                S1.SetActiveCell(S1.RowCount - 1, 3)
                FpSpread1.ShowActiveCell(S1.RowCount - 1, 3)
                Spread_Change(Me.FpSpread1, S1.RowCount - 1)
                Spread_AutoCol(Me.FpSpread1)

                'Else
                '    Modal_Error("이미 안보이스된 출하건입니다 !!!")
                '    Exit Sub
            End If



        End With


    End Sub

    Private Sub FpSpread3_LeaveCell(ByVal sender As Object, ByVal e As FarPoint.Win.Spread.LeaveCellEventArgs) Handles FpSpread3.LeaveCell
        Try
            Dim S3 = Me.FpSpread3.ActiveSheet
            Dim S2 = Me.FpSpread2.ActiveSheet


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
            Dim S1 = Me.FpSpread1.ActiveSheet

            If S3.RowCount > 0 Then

                S3.SetActiveCell(0, 0)   '마지막 변경셀에서 다른 셀로 포커스 이동하지 않아도 저장되도록 자동 포커스 이동

                For i = 0 To S3.RowCount - 1
                    If S3.Rows(i).ForeColor = Color.OrangeRed Then
                        If Query_RS("SELECT SUM(QTY*CHARGE) FROM VIEW_SHIPSUMMARY WHERE SHIP_NO = '" & S3.GetValue(i, 1) & "'") < Query_RS("SELECT ISNULL(SUM(INV_CHARGE),0) FROM TBL_INVSUMMARY WHERE SHIP_NO = '" & S3.GetValue(i, 1) & "'") + S3.GetValue(i, 7) Then
                            Modal_Error("인보이스 금액이 출하금액을 초과합니다.")
                            Exit Sub
                        End If


                        Dim QRY As String = ""

                        QRY = QRY & "IF EXISTS (SELECT INV_NO FROM TBL_INVSUMMARY WHERE SHIP_NO = '" & S3.GetValue(i, 1) & "' AND INV_NO = '" & S3.GetValue(i, 5) & "')" & vbNewLine
                        QRY = QRY & "BEGIN" & vbNewLine
                        QRY = QRY & "   UPDATE TBL_INVSUMMARY SET INV_DATE = '" & S3.GetValue(i, 6) & "', INV_CHARGE = " & S3.GetValue(i, 7) & ", INV_RMK = '" & S3.GetValue(i, 8) & "', U_DATE = GETDATE(), U_PERSON = '" & Emp_No & "'" & vbNewLine
                        QRY = QRY & "   WHERE SHIP_NO = '" & S3.GetValue(i, 1) & "' AND INV_NO = '" & S3.GetValue(i, 5) & "'" & vbNewLine
                        QRY = QRY & "END" & vbNewLine
                        QRY = QRY & "ELSE" & vbNewLine
                        QRY = QRY & "BEGIN" & vbNewLine
                        QRY = QRY & "INSERT INTO TBL_INVSUMMARY VALUES ('" & Site_id & "','"
                        QRY = QRY & S3.GetValue(i, 1) & "','" & S3.GetValue(i, 5) & "'," & S3.GetValue(i, 7) & ",'" & S3.GetValue(i, 6) & "','" & S3.GetValue(i, 8) & "','11111', GETDATE(),'11111', GETDATE())" & vbNewLine
                        QRY = QRY & "END" & vbNewLine


                        If Insert_Data(QRY) = False Then
                            MessageBox.Show("저장 오류.")
                            Exit Sub
                        End If
                        S3.Rows(i).ForeColor = Color.Black
                    End If
                Next
            End If


            If S1.RowCount > 0 Then

                S1.SetActiveCell(0, 0)   '마지막 변경셀에서 다른 셀로 포커스 이동하지 않아도 저장되도록 자동 포커스 이동

                For i = 0 To S1.RowCount - 1
                    If S1.Rows(i).ForeColor = Color.OrangeRed Then
                        If Query_RS("SELECT ISNULL(SUM(INV_CHARGE),0) FROM TBL_INVSUMMARY WHERE SHIP_NO = '" & S1.GetValue(i, 0) & "' AND INV_NO = '" & S1.GetValue(i, 1) & "'") < Query_RS("SELECT ISNULL(SUM(BILL_CHARGE),0) FROM TBL_BILLING WHERE SHIP_NO = '" & S1.GetValue(i, 0) & "' AND INV_NO = '" & S1.GetValue(i, 1) & "'") + S1.GetValue(i, 4) Then
                            Modal_Error("수금 금액이 인보이스 금액을 초과합니다.")
                            Exit Sub
                        End If

                        Dim QRY As String = ""

                        QRY = QRY & "IF EXISTS (SELECT INV_NO FROM TBL_BILLING WHERE SHIP_NO = '" & S1.GetValue(i, 0) & "' AND INV_NO = '" & S1.GetValue(i, 1) & "' AND BILL_SEQ = " & S1.GetValue(i, 2) & ")" & vbNewLine
                        QRY = QRY & "BEGIN" & vbNewLine
                        QRY = QRY & "   UPDATE TBL_BILLING SET BILL_DATE = '" & S1.GetValue(i, 3) & "', BILL_CHARGE = " & S1.GetValue(i, 4) & ", BILL_RMK = '" & S1.GetValue(i, 5) & "', U_DATE = GETDATE(), U_PERSON = '" & Emp_No & "'" & vbNewLine
                        QRY = QRY & "   WHERE SHIP_NO = '" & S1.GetValue(i, 0) & "' AND INV_NO = '" & S1.GetValue(i, 1) & "' AND BILL_SEQ = " & S1.GetValue(i, 2) & vbNewLine
                        QRY = QRY & "END" & vbNewLine
                        QRY = QRY & "ELSE" & vbNewLine
                        QRY = QRY & "BEGIN" & vbNewLine
                        QRY = QRY & "   INSERT INTO TBL_BILLING VALUES ('" & Site_id & "','"
                        QRY = QRY & S1.GetValue(i, 0) & "','" & S1.GetValue(i, 1) & "'," & S1.GetValue(i, 2) & "," & S1.GetValue(i, 4) & ",'" & S1.GetValue(i, 3) & "','" & S1.GetValue(i, 5) & "','11111', GETDATE(),'11111', GETDATE())" & vbNewLine
                        QRY = QRY & "END" & vbNewLine


                        If Insert_Data(QRY) = False Then
                            MessageBox.Show("저장 오류.")
                            Exit Sub
                        End If
                        S1.Rows(i).ForeColor = Color.Black
                    End If
                Next
            End If

            MessageBox.Show("저장이 완료되었습니다.", "Message")

        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message, "ERROR")
        End Try
    End Sub

    Private Sub DelBtn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DelBtn1.Click, DelBtn.Click
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

        If save_excel = "FpSpread1" Then
            File_Save(SaveFileDialog1, FpSpread1)
        ElseIf save_excel = "FpSpread2" Then
            File_Save(SaveFileDialog1, FpSpread2)
        ElseIf save_excel = "FpSpread3" Then
            File_Save(SaveFileDialog1, FpSpread3)
            'ElseIf save_excel = "FpSpread4" Then
            '    File_Save(SaveFileDialog1, FpSpread4)
        Else
            MessageBox.Show("Select Spread for Saving!!")
        End If
    End Sub

    Private Sub disp_io()
        Try



        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message, "ERROR")
        End Try
    End Sub

    Private Sub FpSpread2_CellClick(ByVal sender As System.Object, ByVal e As FarPoint.Win.Spread.CellClickEventArgs) Handles FpSpread2.CellClick
        save_excel = "FpSpread2"
    End Sub


    Private Sub CheckedListBox1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles CheckedListBox1.LostFocus
        TextBoxDropDown1.Text = MODEL_SELECTED(CheckedListBox1)
    End Sub

End Class