Public Class FrmReceiving

    'Private ChkNew As Boolean
    'Private RecCnt As Integer

    Private Sub FrmReceiving_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        DockContainerItem1.Text = "실행 메뉴"
        DockContainerItem2.Text = "조회 조건"
        DockContainerItem3.Text = "입고대기 현황"


        If Spread_Setting(FpSpread1, "FrmOpenra") = True Then
            Spread_AutoCol(FpSpread1)
            Me.FpSpread1.ActiveSheet.Protect = False
            Me.FpSpread1.ActiveSheet.Columns(0, Me.FpSpread1.ActiveSheet.ColumnCount - 1).Locked = False

        End If

        If Spread_Setting(FpSpread2, "FrmOpenra") = True Then
            FpSpread2.ActiveSheet.RowCount = 0
            Spread_AutoCol(FpSpread2)
        End If

        If Spread_Setting(FpSpread3, "FrmOpenra") = True Then
            FpSpread3.ActiveSheet.RowCount = 0
            Spread_AutoCol(FpSpread3)
        End If

        Formbim_Authority(Me.NewBtn, Me.Name, "NEW")
        Formbim_Authority(Me.SaveBtn, Me.Name, "SAVE")
        Formbim_Authority(Me.DelBtn, Me.Name, "DELETE")
        Formbim_Authority(Me.PrtBtn, Me.Name, "PRINT")
        Formbim_Authority(Me.XlsBtn, Me.Name, "EXCEL")
        Formbim_Authority(Me.FindBtn, Me.Name, "FIND")
        Me.SaveBtn.Enabled = False
        Me.NewBtn.Enabled = False
        Me.DelBtn.Enabled = False

    End Sub

    Private Sub FrmAssy_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.GotFocus
        Me.TextBoxX1.Focus()
        TextBoxX1.SelectAll()
    End Sub

    Private Sub TextBoxX1_click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBoxX1.Click
        Me.TextBoxX1.Text = ""
    End Sub

    Private Sub TextBoxX1_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TextBoxX1.KeyDown

        If e.KeyCode <> Keys.Enter Then
            Exit Sub
        End If

        If ComboBoxEx1.Text = "플립" Then
            If FESN_VERIFY(Me.TextBoxX1, e) = False Then
                TextBoxX1.Focus()
                TextBoxX1.SelectAll()
                Modal_Error("Wrong Format Flip ID")
                Exit Sub
            End If

            If Check_Valid_FEsn(TextBoxX1.Text, Me.Name) = False Then
                TextBoxX1.Focus()
                TextBoxX1.SelectAll()
                '            Modal_Error("Wrong Flip")
                Exit Sub
            End If


            Dim rcv_rs As New ADODB.Recordset
            rcv_rs = Query_RS_ALL("EXEC SP_COMMON_CheckValidEsn '" & Site_id & "','" & TextBoxX1.Text & "','2'")

            If rcv_rs(7).Value = "Y" Then
                System.Console.Beep(3000, 400)
                System.Console.Beep(3000, 400)
                Modal_Error(TextBoxX1.Text & vbNewLine & "이미 제품입고 완료된 플립입니다!!")
                
                TextBoxX1.Focus()
                TextBoxX1.SelectAll()
                
                Exit Sub
            Else
                If Insert_Data("update tbl_fesnmaster_k set chk_rcv = 'Y', RCV_DATE = GETDATE(), C_PRC = 'K1000', U_PERSON = '" & Emp_No & "', U_DATE= GETDATE() WHERE OUT_ESN = '" & TextBoxX1.Text & "'") = True Then

                    With FpSpread2.ActiveSheet
                        .AddRows(0, 1)
                        For I As Integer = 0 To rcv_rs.Fields.Count - 1
                            If I = 7 Then
                                .SetValue(0, I, "Y")
                            Else
                                .SetValue(0, I, rcv_rs(I).Value)
                            End If
                        Next
                    End With
                    Spread_AutoCol(FpSpread2)
                End If
            End If
        Else

        End If

        FpSpread1.ActiveSheet.RowCount = 0
        FpSpread3.ActiveSheet.RowCount = 0

        If Query_Spread(FpSpread1, "exec KSP_FRMOPENRA_GETDAILY '" & Site_id & "', '" & Mid(FpSpread2.ActiveSheet.GetValue(0, 3), 4, 8) & "', '" & Mid(FpSpread2.ActiveSheet.GetValue(0, 3), 4, 8) & "','" & FpSpread2.ActiveSheet.GetValue(0, 1) & "'", 1) = True Then

            With FpSpread1.ActiveSheet
                .RowCount = .RowCount + 1
                .Rows(.RowCount - 1).BackColor = Color.Yellow
                .Rows(.RowCount - 1).ForeColor = Color.Black
                .Rows(.RowCount - 1).Locked = True
                .SetValue(.RowCount - 1, 0, "TOTAL")

                SPREAD_ROW_TOTAL(FpSpread1, 3, 5, 1)
            End With

            Spread_AutoCol(FpSpread1)
        End If

        TextBoxX1.Focus()
        TextBoxX1.SelectAll()


    End Sub


    Private Sub FindBtn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FindBtn.Click

        FpSpread1.ActiveSheet.RowCount = 0
        'FpSpread2.ActiveSheet.RowCount = 0
        FpSpread3.ActiveSheet.RowCount = 0



    End Sub

    Private Sub ExcelBtn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles XlsBtn.Click

        If save_excel = "" Then
            MessageBox.Show("Select Spread for Saving!!")
            Exit Sub
        End If

        If save_excel = "FpSpread1" Then
            'File_Save(SaveFileDialog1, FpSpread1)
        ElseIf save_excel = "FpSpread1" Then
            File_Save(SaveFileDialog1, FpSpread1)
        ElseIf save_excel = "FpSpread2" Then
            File_Save(SaveFileDialog1, FpSpread2)
        ElseIf save_excel = "FpSpread3" Then
            File_Save(SaveFileDialog1, FpSpread2)
        Else
            MessageBox.Show("Select Spread for Saving!!")
        End If

    End Sub

    Private Sub SaveBtn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SaveBtn.Click
        Me.Cursor = Cursors.WaitCursor

        Me.Cursor = Cursors.Default
    End Sub

    Private Sub DelBtn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DelBtn.Click


    End Sub

    Private Sub PrtBtn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PrtBtn.Click

        If save_excel = "" Then
            MessageBox.Show("Select Spread for Printing!!")
            Exit Sub
        End If

        If save_excel = "FpSpread1" Then
            'If Spread_Print(Me.FpSpread1, DockContainerItem1.Text, 0) = False Then
            '    MsgBox("Fail to Print")
            'End If
        ElseIf save_excel = "FpSpread1" Then
            If Spread_Print(Me.FpSpread1, "Receiving Summary", 0) = False Then
                MsgBox("Fail to Print")
            End If
        ElseIf save_excel = "FpSpread2" Then
            If Spread_Print(Me.FpSpread2, "Receiving details", 0) = False Then
                MsgBox("Fail to Print")
            End If
        ElseIf save_excel = "FpSpread3" Then
            If Spread_Print(Me.FpSpread3, "Receiving details", 0) = False Then
                MsgBox("Fail to Print")
            End If
        Else
            MessageBox.Show("Select Spread for Printing!!")
        End If


    End Sub

    Private Sub FpSpread1_CellDoubleClick(ByVal sender As Object, ByVal e As FarPoint.Win.Spread.CellClickEventArgs) Handles FpSpread1.CellDoubleClick
        save_excel = "FpSpread1"

        'FpSpread2.ActiveSheet.RowCount = 0
        FpSpread3.ActiveSheet.RowCount = 0


        If Query_Spread(FpSpread3, "exec KSP_FRMOPENRA_GETSUMMARY '" & Site_id & "', '" & FpSpread1.ActiveSheet.GetValue(e.Row, 0) & "', '" & FpSpread1.ActiveSheet.GetValue(e.Row, 0) & "','" & FpSpread1.ActiveSheet.GetValue(e.Row, 1) & "',''", 1) = True Then
            With FpSpread3.ActiveSheet
                .RowCount = .RowCount + 1
                .Rows(.RowCount - 1).BackColor = Color.Yellow
                .Rows(.RowCount - 1).ForeColor = Color.Black
                .Rows(.RowCount - 1).Locked = True
                .SetValue(.RowCount - 1, 0, "TOTAL")

                SPREAD_ROW_TOTAL(FpSpread3, 3, 5, 1)
            End With

            Spread_AutoCol(FpSpread3)
        End If







    End Sub


    Private Sub FpSpread3_CellDoubleClick(ByVal sender As Object, ByVal e As FarPoint.Win.Spread.CellClickEventArgs) Handles FpSpread3.CellDoubleClick
        save_excel = "FpSpread3"

        'FpSpread2.ActiveSheet.RowCount = 0

        'Dim qry As String = ""

        'qry = qry & "Select OUT_esn, model,qareject, PSHIP_NO, OUTBOX_NO, CONVERT(INT,OUT_POS), OUT_DEF_CD, CHK_RCV, '' " & vbNewLine
        'qry = qry & "from TBL_Fesnmaster_K" & vbNewLine
        'qry = qry & "where site_id = '" & Site_id & "'" & vbNewLine
        'qry = qry & "and PSHIP_NO = '" & FpSpread3.ActiveSheet.GetValue(e.Row, 0) & "'" & vbNewLine
        'qry = qry & "AND MODEL = '" & FpSpread3.ActiveSheet.GetValue(e.Row, 1) & "'" & vbNewLine
        'qry = qry & "AND OUTBOX_NO = '" & FpSpread3.ActiveSheet.GetValue(e.Row, 6) & "'" & vbNewLine
        'qry = qry & "ORDER BY CONVERT(INT,OUT_POS)" & vbNewLine
        'qry = qry & "" & vbNewLine

        'If Query_Spread(FpSpread2, qry, 1) = True Then
        '    Spread_AutoCol(FpSpread2)
        'End If

    End Sub



    Private Sub FpSpread1_CellClick(ByVal sender As System.Object, ByVal e As FarPoint.Win.Spread.CellClickEventArgs) Handles FpSpread1.CellClick
        save_excel = "FpSpread1"
    End Sub
    Private Sub FpSpread2_CellClick(ByVal sender As System.Object, ByVal e As FarPoint.Win.Spread.CellClickEventArgs) Handles FpSpread2.CellClick
        save_excel = "FpSpread2"
    End Sub
    Private Sub FpSpread3_CellClick(ByVal sender As System.Object, ByVal e As FarPoint.Win.Spread.CellClickEventArgs) Handles FpSpread3.CellClick
        save_excel = "FpSpread3"
    End Sub


    Private Sub ComboBoxEx1_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBoxEx1.SelectedIndexChanged

        LabelX1.Text = ComboBoxEx1.Text & " ID"
        TextBoxX1.Text = ""
        TextBoxX1.Focus()

    End Sub


End Class