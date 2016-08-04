Public Class FrmResult

    Private Sub ComboBoxEx99_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub FrmResult_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        DockContainerItem1.Text = "실행 메뉴"
        DockContainerItem2.Text = "조회 조건"

        DateTimeInput1.Value = Now
        DateTimeInput2.Value = Now

        curcell.Separator = ","
        curcell.ShowSeparator = True
        curcell.ShowCurrencySymbol = True
        curcell.CurrencySymbol = "$"

        'If Query_Combo(Me.ComboBoxEx1, "SELECT model_no FROM tbl_modelmaster WHERE site_id = '" & Site_id & "' and active = 'Y' ORDER BY model_no") = True Then
        'End If
        'Me.ComboBoxEx1.Items.Add("ALL")
        'Me.ComboBoxEx1.Text = "ALL"


        If Query_CheckList(Me.CheckedListBox1, "SELECT model_no FROM tbl_modelmaster WHERE site_id = '" & Site_id & "' and active = 'Y' ORDER BY model_no") = True Then
            Me.CheckedListBox1.Items.Add("ALL")
            Me.CheckedListBox1.Text = "ALL"
        End If

        With FpSpread1.ActiveSheet

            .ColumnHeaderRowCount = 2
            .ColumnHeader.Cells(0, 0).RowSpan = 2
            .ColumnHeader.Cells(0, 0).Text = "모델"

            .ColumnHeader.Cells(0, 1).ColumnSpan = 3
            .ColumnHeader.Cells(0, 1).Text = "입고"
            .ColumnHeader.Cells(0, 5).Text = "입고"

            .ColumnHeader.Cells(0, 4).ColumnSpan = 4
            .ColumnHeader.Cells(0, 4).Text = "출하"
            .ColumnHeader.Cells(0, 7).Text = "출하"

            .ColumnHeader.Cells(0, 8).ColumnSpan = 3
            .ColumnHeader.Cells(0, 8).Text = "수금"

            If Spread_Setting(FpSpread1, Me.Name) = True Then
                .RowCount = 0
                Spread_AutoCol(FpSpread1)

                .FrozenColumnCount = 1

                '.Columns(1, 2).Visible = False
                '.Columns(4, 5).Visible = False
                '.Columns(8, 9).Visible = False

                .Columns(3).BackColor = Color.LightGreen
                .Columns(6, 7).BackColor = Color.LightPink
                .Columns(9, 10).BackColor = Color.LightYellow

                .Columns(8, 10).Visible = False
            End If

            If Query_RS("SELECT ISNULL(EMP_TYPE,'') FROM TBL_EMPMASTER WHERE SITE_ID = '" & Site_id & "' AND EMP_NO = '" & Emp_No & "'") = "10001" Then

            End If

            curcell.DecimalPlaces = 2
            curcell.ShowSeparator = True

            .Columns(7).CellType = curcell
            .Columns(10).CellType = curcell

            'End If
        End With

        Me.Chart1.Visible = False

        Me.ContextMenuBar1.SetContextMenuEx(Me.Chart1, ChartM)

        Formbim_Authority(Me.ButtonItem2, Me.Name, "NEW")
        Formbtn_Authority(Me.NewBtn1, Me.Name, "NEW")
        Formbim_Authority(Me.ButtonItem3, Me.Name, "SAVE")
        Formbtn_Authority(Me.SaveBtn1, Me.Name, "SAVE")
        Formbim_Authority(Me.ButtonItem4, Me.Name, "DELETE")
        Formbtn_Authority(Me.DelBtn1, Me.Name, "DELETE")
        Formbim_Authority(Me.ButtonItem5, Me.Name, "PRINT")
        Formbtn_Authority(Me.PrtBtn1, Me.Name, "PRINT")
        Formbim_Authority(Me.ButtonItem6, Me.Name, "EXCEL")
        Formbtn_Authority(Me.XlsBtn1, Me.Name, "EXCEL")
        Formbim_Authority(Me.FindBtn, Me.Name, "FIND")

        ButtonItem2.Enabled = False
        ButtonItem3.Enabled = False
        ButtonItem4.Enabled = False


        If Query_RS("select isnull(insa_yn,'N') from tbl_empmaster where emp_no = '" & Emp_No & "'") = "N" Then
            FpSpread1.ActiveSheet.Columns(7).Visible = False
        End If

    End Sub


    Private Sub FindBtn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FindBtn.Click

        Dim Qry As String
        Dim m_qry As String = ""
        m_qry = MODEL_SELECTED(CheckedListBox1)

        FpSpread1.ActiveSheet.RowCount = 0

        Qry = "set nocount on" & vbNewLine & vbNewLine

        Qry = Qry & "SELECT CASE WHEN (SELECT COUNT(MODEL_NO) FROM TBL_MODELMASTER WHERE MODEL_NO = A.MODEL) > 0 THEN A.MODEL ELSE (SELECT PART_NO + '('+PART_NAME +')' FROM TBL_PARTMASTER WHERE PART_NO = A.MODEL) END," & vbNewLine
        Qry = Qry & "			ISNULL((SELECT COUNT(MODEL) FROM VIEW_FESNMASTER WHERE MODEL = A.MODEL AND CHK_RCV = 'N' AND SUBSTRING(PSHIP_NO,4,8) BETWEEN CONVERT(VARCHAR(8), (CONVERT(DATETIME, '" & DateTimeInput1.Text & "')), 112) AND CONVERT(VARCHAR(8), (CONVERT(DATETIME, '" & DateTimeInput2.Text & "')), 112)),0) AS 입고대기," & vbNewLine
        Qry = Qry & "			ISNULL((SELECT COUNT(MODEL) FROM VIEW_FESNMASTER WHERE MODEL = A.MODEL AND CHK_RCV = 'Y' AND SUBSTRING(PSHIP_NO,4,8) BETWEEN CONVERT(VARCHAR(8), (CONVERT(DATETIME, '" & DateTimeInput1.Text & "')), 112) AND CONVERT(VARCHAR(8), (CONVERT(DATETIME, '" & DateTimeInput2.Text & "')), 112)),0) AS 입고완료," & vbNewLine
        Qry = Qry & "			ISNULL((SELECT COUNT(MODEL) FROM VIEW_FESNMASTER WHERE MODEL = A.MODEL AND SUBSTRING(PSHIP_NO,4,8) BETWEEN CONVERT(VARCHAR(8), (CONVERT(DATETIME, '" & DateTimeInput1.Text & "')), 112) AND CONVERT(VARCHAR(8), (CONVERT(DATETIME, '" & DateTimeInput2.Text & "')), 112)),0) AS 입고합계," & vbNewLine

        Qry = Qry & "			ISNULL((SELECT SUM(QTY) FROM tbl_shipsummary  WHERE MODEL = A.MODEL  AND SHIP_DATE BETWEEN CONVERT(VARCHAR(8), (CONVERT(DATETIME, '" & DateTimeInput1.Text & "')), 112) AND CONVERT(VARCHAR(8), (CONVERT(DATETIME, '" & DateTimeInput2.Text & "')), 112)),0) AS 생산출하," & vbNewLine
        Qry = Qry & "			ISNULL((SELECT SUM(QTY) FROM tbl_shipsummary_GOODS  WHERE MODEL = A.MODEL  AND SHIP_DATE BETWEEN CONVERT(VARCHAR(8), (CONVERT(DATETIME, '" & DateTimeInput1.Text & "')), 112) AND CONVERT(VARCHAR(8), (CONVERT(DATETIME, '" & DateTimeInput2.Text & "')), 112)),0) AS 상품출하," & vbNewLine
        Qry = Qry & "			ISNULL((SELECT SUM(QTY) FROM VIEW_shipsummary  WHERE MODEL = A.MODEL  AND SHIP_DATE BETWEEN CONVERT(VARCHAR(8), (CONVERT(DATETIME, '" & DateTimeInput1.Text & "')), 112) AND CONVERT(VARCHAR(8), (CONVERT(DATETIME, '" & DateTimeInput2.Text & "')), 112)),0) AS 출하합계," & vbNewLine
        Qry = Qry & "			ISNULL((SELECT SUM(QTY*CHARGE) FROM VIEW_shipsummary  WHERE MODEL = A.MODEL  AND SHIP_DATE BETWEEN CONVERT(VARCHAR(8), (CONVERT(DATETIME, '" & DateTimeInput1.Text & "')), 112) AND CONVERT(VARCHAR(8), (CONVERT(DATETIME, '" & DateTimeInput2.Text & "')), 112)),0) AS 출하금액," & vbNewLine

        Qry = Qry & "0,0,0" & vbNewLine
        Qry = Qry & "FROM VIEW_SHIPSUMMARY A" & vbNewLine
        Qry = Qry & "WHERE SHIP_DATE IS NOT NULL" & vbNewLine

        If m_qry <> "" Then
            Qry = Qry & "AND MODEL IN (" & m_qry & ")" & vbNewLine
        End If
        'If ComboBoxEx1.Text <> "ALL" Then
        '    Qry = Qry & "AND MODEL  = '" & ComboBoxEx1.Text & "'" & vbNewLine
        'End If
        Qry = Qry & "GROUP BY MODEL" & vbNewLine
        Qry = Qry & "ORDER BY MODEL" & vbNewLine


        If Query_Spread(FpSpread1, Qry, 1) = True Then


            '        If Query_Spread(FpSpread1, Qry & " '" & Site_id & "','" & ComboBoxEx1.Text & "','" & DateTimeInput1.Text & "','" & DateTimeInput2.Text & "'", 1) = True Then
            With FpSpread1.ActiveSheet
                .RowCount = .RowCount + 1
                .Cells(.RowCount - 1, 1, .RowCount - 1, .ColumnCount - 1).CellType = intcell
                .Cells(.RowCount - 1, .ColumnCount - 1).CellType = deccell
                .Rows(.RowCount - 1).BackColor = Color.Yellow
                .SetText(.RowCount - 1, 0, "TOTAL")

                .Cells(.RowCount - 1, 7, .RowCount - 1, 7).CellType = curcell
                '.Cells(.RowCount - 1, 9, .RowCount - 1, 9).CellType = curcell
                .Cells(.RowCount - 1, 10, .RowCount - 1, 10).CellType = curcell


                SPREAD_ROW_TOTAL(FpSpread1, 1, .ColumnCount - 1, 1)
                .Cells(.RowCount - 1, 7, .RowCount - 1, 7).Value = 0
                .Cells(.RowCount - 1, 10, .RowCount - 1, 10).Value = 0
                SPREAD_ROW_TOTAL_DEC(FpSpread1, 7, 7, 1)
                'SPREAD_ROW_TOTAL_DEC(FpSpread1, 9, 9, 1)
                SPREAD_ROW_TOTAL_DEC(FpSpread1, 10, 10, 1)

            End With

            curcell.CurrencySymbol = "$"
            curcell.DecimalPlaces = 2
            curcell.ShowCurrencySymbol = True
            curcell.ShowSeparator = True

            Spread_AutoCol(FpSpread1)
            FpSpread1.ActiveSheet.Columns(1, 10).Width = 80

            disp_chart(-1)

        End If

        'End If

    End Sub


    Private Sub ButtonItem2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonItem2.Click

    End Sub

    Private Sub ButtonItem3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonItem3.Click

    End Sub

    Private Sub ButtonItem4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonItem4.Click

    End Sub

    Private Sub ButtonItem5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonItem5.Click, PrtBtn1.Click
        If save_excel = "" Then
            MessageBox.Show("Select Spread for Printing!!")
            Exit Sub
        End If

        If save_excel = "FpSpread1" Then
            If Spread_Print(Me.FpSpread1, "NO SHIP(Board) Summary", 0) = False Then
                MsgBox("Fail to Print")
            End If
            'ElseIf save_excel = "FpSpread2" Then
            '    If Spread_Print(Me.FpSpread2, "NO SHIP(Board) Details", 0) = False Then
            '        MsgBox("Fail to Print")
            '    End If
        Else
            MessageBox.Show("Select Spread for Printing!!")
        End If
    End Sub

    Private Sub ButtonItem6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonItem6.Click, XlsBtn1.Click

        If save_excel = "" Then
            MessageBox.Show("Select Spread for Saving!!")
            Exit Sub
        End If

        If save_excel = "FpSpread1" Then
            File_Save(SaveFileDialog1, FpSpread1)
            'ElseIf save_excel = "FpSpread2" Then
            '    File_Save(SaveFileDialog1, FpSpread2)
        Else
            MessageBox.Show("Select Spread for Saving!!")
        End If

    End Sub

    Private Sub FpSpread1_CellClick(ByVal sender As System.Object, ByVal e As FarPoint.Win.Spread.CellClickEventArgs) Handles FpSpread1.CellClick
        save_excel = "FpSpread1"

        '  disp_chart(e.Row)

    End Sub

    Private Sub disp_chart(ByVal rowidx As Integer)

        Me.Chart1.Visible = True

        If rowidx < 0 Then

            'Chart_Spread_All_1(Me.Chart1, False, 2, DataVisualization.Charting.SeriesChartType.Column, Me.FpSpread1, New String() {"RECEIVING", "SHIPPING", "CLAIMED", "INVOICED"}, 0, New Integer() {5, 12, 21, 31}, True)
            Chart_Spread_All_1(Me.Chart1, False, 2, DataVisualization.Charting.SeriesChartType.Column, Me.FpSpread1, New String() {"입고", "출하", "청구"}, 0, New Integer() {3, 6, 9}, True)

        Else
            Chart_Spread_Col(Me.Chart1, False, 2, DataVisualization.Charting.SeriesChartType.Line, Me.FpSpread1, Me.FpSpread1.ActiveSheet.GetValue(rowidx, 0), rowidx, New Integer() {1, 2, 4, 5, 7, 8}, 1)
            'Chart_Spread_Col(Me.Chart1, False, 2, DataVisualization.Charting.SeriesChartType.Line, Me.FpSpread1, Me.FpSpread1.ActiveSheet.GetValue(rowidx, 0), rowidx, New Integer() {13, 14, 15, 16, 17, 18, 19, 20, 24, 25, 26, 27, 28, 29, 30}, 1)

        End If
    End Sub

    Private Sub ButtonItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonItem1.Click
        Chart_Print(Me.Chart1, True)
    End Sub


    Private Sub CheckedListBox1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles CheckedListBox1.LostFocus
        TextBoxDropDown1.Text = MODEL_SELECTED(CheckedListBox1)
    End Sub

End Class