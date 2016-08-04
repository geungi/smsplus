Public Class FrmOpenra

    'Private ChkNew As Boolean
    'Private RecCnt As Integer
    Private Sub CheckedListBox1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles CheckedListBox1.LostFocus
        TextBoxDropDown1.Text = MODEL_SELECTED(CheckedListBox1)
    End Sub

    Private Sub FrmOpenra_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        DockContainerItem1.Text = "실행 메뉴"
        DockContainerItem2.Text = "조회 조건"
        DockContainerItem3.Text = "입고대기 현황"

        DateTimeInput1.Text = Now
        DateTimeInput2.Text = Now

        'If Query_Combo(Me.ComboBoxEx1, "SELECT model_no FROM tbl_modelmaster WHERE site_id = '" & Site_id & "' and active = 'Y' ORDER BY model_no") = True Then
        'End If
        'Me.ComboBoxEx1.Items.Add("ALL")
        'Me.ComboBoxEx1.Text = "ALL"
        If Query_CheckList(Me.CheckedListBox1, "SELECT model_no FROM tbl_modelmaster WHERE site_id = '" & Site_id & "' and active = 'Y' ORDER BY model_no") = True Then
            Me.CheckedListBox1.Items.Add("ALL")
            Me.CheckedListBox1.Text = "ALL"
        End If

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
        Me.NewBtn.Enabled = False
        Formbim_Authority(Me.SaveBtn, Me.Name, "SAVE")
        Formbim_Authority(Me.DelBtn, Me.Name, "DELETE")
        Formbim_Authority(Me.PrtBtn, Me.Name, "PRINT")
        Formbim_Authority(Me.XlsBtn, Me.Name, "EXCEL")
        Formbim_Authority(Me.FindBtn, Me.Name, "FIND")

    End Sub

    Private Sub FindBtn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FindBtn.Click

        FpSpread1.ActiveSheet.RowCount = 0
        FpSpread2.ActiveSheet.RowCount = 0
        FpSpread3.ActiveSheet.RowCount = 0

        Dim m_qry As String = ""
        m_qry = MODEL_SELECTED(CheckedListBox1)
        Dim qry As String = ""


        qry = qry & "select SUBSTRING(PSHIP_NO,4,8), model,QAREJECT, COUNT(MODEL), " & vbNewLine
        qry = qry & "(SELECT COUNT(MODEL) FROM VIEW_FESNMASTER WHERE SUBSTRING(PSHIP_NO,4,8) = SUBSTRING(a.PSHIP_NO,4,8)	 AND MODEL = A.MODEL AND CHK_RCV ='Y'  AND QAREJECT = A.qareject) AS RCVCNT," & vbNewLine
        qry = qry & "(SELECT COUNT(MODEL) FROM VIEW_FESNMASTER WHERE SUBSTRING(PSHIP_NO,4,8) = SUBSTRING(a.PSHIP_NO,4,8)	 AND MODEL = A.MODEL AND CHK_RCV ='N'  AND QAREJECT = A.qareject) AS NORCVCNT" & vbNewLine
        qry = qry & "from VIEW_FESNMASTER  a" & vbNewLine
        qry = qry & "where site_id = '" & Site_id & "'" & vbNewLine
        If m_qry <> "" Then
            qry = qry & "AND MODEL IN (" & m_qry & ")" & vbNewLine
        End If
        qry = qry & "	and SUBSTRING(PSHIP_NO,4,8) between CONVERT(VARCHAR(8), CONVERT(DATETIME,'" & DateTimeInput1.Text & "'),112) and CONVERT(VARCHAR(8), CONVERT(DATETIME,'" & DateTimeInput2.Text & "'),112)	" & vbNewLine
        qry = qry & "GROUP BY SITE_ID, SUBSTRING(PSHIP_NO,4,8), model,QAREJECT" & vbNewLine
        qry = qry & "order by SITE_ID, SUBSTRING(PSHIP_NO,4,8), model" & vbNewLine
        qry = qry & "" & vbNewLine
        qry = qry & "" & vbNewLine
        qry = qry & "" & vbNewLine
        qry = qry & "" & vbNewLine
        qry = qry & "" & vbNewLine

        If Query_Spread(FpSpread1, qry, 1) = True Then
            'If Query_Spread(FpSpread1, "exec KSP_FRMOPENRA_GETDAILY '" & Site_id & "', '" & DateTimeInput1.Text & "', '" & DateTimeInput2.Text & "','" & ComboBoxEx1.Text & "'", 1) = True Then
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


        Me.FpSpread1.ActiveSheet.Protect = False
        Me.FpSpread1.ActiveSheet.Columns(0, Me.FpSpread1.ActiveSheet.ColumnCount - 1).Locked = False
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

        FpSpread2.ActiveSheet.RowCount = 0
        FpSpread3.ActiveSheet.RowCount = 0


        If Query_Spread(FpSpread3, "exec KSP_FRMOPENRA_GETSUMMARY '" & Site_id & "', '" & FpSpread1.ActiveSheet.GetValue(e.Row, 0) & "', '" & FpSpread1.ActiveSheet.GetValue(e.Row, 0) & "','" & FpSpread1.ActiveSheet.GetValue(e.Row, 1) & "','','" & FpSpread1.ActiveSheet.GetValue(e.Row, 2) & "'", 1) = True Then
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

        FpSpread2.ActiveSheet.RowCount = 0

        Dim qry As String = ""

        qry = qry & "Select OUT_esn, model,qareject, PSHIP_NO, OUTBOX_NO, CONVERT(INT,OUT_POS), OUT_DEF_CD, CHK_RCV, '' " & vbNewLine
        qry = qry & "from TBL_Fesnmaster_K" & vbNewLine
        qry = qry & "where site_id = '" & Site_id & "'" & vbNewLine
        qry = qry & "and PSHIP_NO = '" & FpSpread3.ActiveSheet.GetValue(e.Row, 0) & "'" & vbNewLine
        qry = qry & "AND MODEL = '" & FpSpread3.ActiveSheet.GetValue(e.Row, 1) & "'" & vbNewLine
        qry = qry & "AND OUTBOX_NO = '" & FpSpread3.ActiveSheet.GetValue(e.Row, 6) & "'" & vbNewLine
        qry = qry & "ORDER BY CONVERT(INT,OUT_POS)" & vbNewLine
        qry = qry & "" & vbNewLine

        If Query_Spread(FpSpread2, qry, 1) = True Then
            Spread_AutoCol(FpSpread2)
        End If

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


End Class