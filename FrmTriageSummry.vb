Public Class FrmTriageSummary

    Private Sub CheckedListBox1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles CheckedListBox1.LostFocus
        TextBoxDropDown1.Text = MODEL_SELECTED(CheckedListBox1)
    End Sub

    Private Sub FrmTriageSummary_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim rs As ADODB.Recordset = Nothing
        Dim i As Integer = 0

        DockContainerItem1.Text = "실행 메뉴"
        DockContainerItem2.Text = "조회 조건"

        'If Query_Combo(Me.ComboBoxEx1, "SELECT model_no FROM tbl_modelmaster WHERE site_id = '" & Site_id & "' and active = 'Y' ORDER BY model_no") = True Then
        'End If
        'Me.ComboBoxEx1.Items.Add("ALL")
        'Me.ComboBoxEx1.Text = "ALL"
        If Query_CheckList(Me.CheckedListBox1, "SELECT model_no FROM tbl_modelmaster WHERE site_id = '" & Site_id & "' and active = 'Y' ORDER BY model_no") = True Then
            Me.CheckedListBox1.Items.Add("ALL")
            Me.CheckedListBox1.Text = "ALL"
        End If

        DateTimeInput1.Value = Now
        DateTimeInput2.Value = Now

        If Spread_Setting(FpSpread1, Me.Name) = True Then
            FpSpread1.ActiveSheet.RowCount = 0
            Spread_AutoCol(FpSpread1)
        End If

        If Spread_Setting(FpSpread2, Me.Name) = True Then
            FpSpread2.ActiveSheet.RowCount = 0
            Spread_AutoCol(FpSpread2)
        End If

        Formbim_Authority(Me.ButtonItem12, Me.Name, "PRINT")
        Formbim_Authority(Me.ButtonItem13, Me.Name, "EXCEL")
        Formbim_Authority(Me.FindBtn, Me.Name, "FIND")

    End Sub

    Private Sub FindBtn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FindBtn.Click
        Dim qry As String = ""
        'Dim str_model As String = ""
        Dim i As Integer = 0
        Dim J As Integer = 0
        Dim rs As ADODB.Recordset = Nothing
        Dim m_qry As String = ""
        m_qry = MODEL_SELECTED(CheckedListBox1)

        With FpSpread1.ActiveSheet
            .SetColumnMerge(-1, FarPoint.Win.Spread.Model.MergePolicy.None)

            .RowCount = 0
            FpSpread2.ActiveSheet.RowCount = 0

            .FrozenColumnCount = 1


            qry = "select A.MODEL, A.TRG1, (SELECT CODE_NAME FROM TBL_DEFECT WHERE CODE_ID = A.TRG1 ), SUM(A.INIT_QTY ), SUM(A.INIT_QTY )*1.0/(SELECT SUM(INIT_QTY) FROM VIEW_TRIAGE WHERE MODEL = A.MODEL AND TRG_DT BETWEEN '" & DateTimeInput1.Text & "' AND '" & DateTimeInput2.Text & "' + ' 23:59:59') " & vbNewLine
            qry = qry & "from VIEW_TRIAGE A " & vbNewLine
            qry = qry & "where A.TRG_DT BETWEEN '" & DateTimeInput1.Text & "' AND '" & DateTimeInput2.Text & "' + ' 23:59:59'" & vbNewLine
            'If ComboBoxEx1.Text <> "ALL" Then
            '    qry = qry & "AND B.MODEL = '" & ComboBoxEx1.Text & "'" & vbNewLine
            'End If
            If m_qry <> "" Then
                qry = qry & "AND a.MODEL IN (" & m_qry & ")" & vbNewLine
            End If
            qry = qry & "GROUP BY A.MODEL, A.TRG1" & vbNewLine
            qry = qry & "ORDER BY A.MODEL, A.TRG1" & vbNewLine
            qry = qry & "" & vbNewLine
            qry = qry & "" & vbNewLine

            If Query_Spread(FpSpread1, qry, 1) = True Then
                .RowCount = .RowCount + 1
                .Cells(.RowCount - 1, 1, .RowCount - 1, .ColumnCount - 1).CellType = intcell
                .Rows(.RowCount - 1).BackColor = Color.Yellow
                .SetText(.RowCount - 1, 0, "TOTAL")
                SPREAD_ROW_TOTAL(FpSpread1, 3, 3, 1)
                Spread_AutoCol(FpSpread1)

                If .RowCount > 1 Then
                    i = 1
                    J = 0
                    While i <= .RowCount - 1
                        If .GetValue(i, 0) <> .GetValue(i - 1, 0) Then
                            If .GetValue(i, 1) <> "소계" Then
                                .AddRows(i, 1)
                                .SetValue(i, 0, .GetValue(i - 1, 0))
                                .SetValue(i, 1, "소계")
                                J = J + .GetValue(i - 1, 3)
                                .SetValue(i, 3, J)
                                J = 0
                                .Rows(i).BackColor = Color.YellowGreen
                            End If
                            i = i + 2
                        Else
                            J = J + .GetValue(i - 1, 3)
                            i = i + 1
                        End If

                    End While
                End If
                '              .SetRowMerge(-1, FarPoint.Win.Spread.Model.MergePolicy.Always)
                .SetColumnMerge(0, FarPoint.Win.Spread.Model.MergePolicy.Always)
            End If
        End With
    End Sub

    Private Sub FpSpread1_CellDoubleClick(ByVal sender As Object, ByVal e As FarPoint.Win.Spread.CellClickEventArgs) Handles FpSpread1.CellDoubleClick
        Dim qry As String = ""
        Dim str_temp As String = ""
        Dim i As Integer = 0

        FpSpread2.ActiveSheet.RowCount = 0

        qry = "select E.LOT_NO, E.MODEL, ISNULL(E.TRG1,''), ISNULL(E.TRG2,''), ISNULL(E.TRG3,''), ISNULL(E.TRG4,''), ISNULL(E.TRG5,''), E.TRG_DT, (SELECT EMP_NM FROM TBL_EMPMASTER WHERE EMP_NO = E.TRG_OP ), INIT_QTY" & vbNewLine
        qry += " from VIEW_TRIAGE E" & vbNewLine
        qry += " where E.MODEL = '" & FpSpread1.ActiveSheet.GetValue(FpSpread1.ActiveSheet.ActiveRowIndex, 0) & "'" & vbNewLine
        qry += " and E.TRG1 = '" & FpSpread1.ActiveSheet.GetValue(FpSpread1.ActiveSheet.ActiveRowIndex, 1) & "'" & vbNewLine
        qry += " and TRG_DT BETWEEN '" & DateTimeInput1.Text & "' AND '" & DateTimeInput2.Text & "' + ' 23:59:59'" & vbNewLine
        qry += " ORDER BY TRG_DT" & vbNewLine

        Query_Spread(FpSpread2, qry, 1)
        Spread_AutoCol(FpSpread2)
    End Sub

    Private Sub ButtonItem12_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonItem12.Click
        If save_excel = "" Then
            MessageBox.Show("Select Spread for Printing!!")
            Exit Sub
        End If

        If save_excel = "FpSpread1" Then
            If Spread_Print(Me.FpSpread1, "Triage Report", 0) = False Then
                MsgBox("Fail to Print")
            End If
        ElseIf save_excel = "FpSpread2" Then
            If Spread_Print(Me.FpSpread2, "Triage Report", 0) = False Then
                MsgBox("Fail to Print")
            End If
        Else
            MessageBox.Show("Select Spread for Printing!!")
        End If
    End Sub

    Private Sub ButtonItem13_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonItem13.Click
        If save_excel = "" Then
            MessageBox.Show("Select Spread for Saving!!")
            Exit Sub
        End If

        If save_excel = "FpSpread1" Then
            File_Save(SaveFileDialog1, FpSpread1)
        ElseIf save_excel = "FpSpread2" Then
            File_Save(SaveFileDialog1, FpSpread2)
        Else
            MessageBox.Show("Select Spread for Saving!!")
        End If
    End Sub

    Private Sub FpSpread1_CellClick(ByVal sender As System.Object, ByVal e As FarPoint.Win.Spread.CellClickEventArgs) Handles FpSpread1.CellClick
        save_excel = "FpSpread1"
    End Sub

    Private Sub FpSpread2_CellClick(ByVal sender As System.Object, ByVal e As FarPoint.Win.Spread.CellClickEventArgs) Handles FpSpread2.CellClick
        save_excel = "FpSpread2"
    End Sub
End Class