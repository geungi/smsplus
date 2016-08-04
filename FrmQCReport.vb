Imports System
Imports System.Drawing
Imports System.Collections
Imports System.ComponentModel
Imports System.Windows.Forms
Imports FarPoint.Win.Spread
Imports FarPoint.Win

Public Class FrmQCReport

    Private tot_smp As Integer = 0  'total sample size
    Private xpos As Integer = 101  'checkbox 초기 X축 위치 

    Private Sub CheckedListBox1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles CheckedListBox1.LostFocus
        TextBoxDropDown1.Text = MODEL_SELECTED(CheckedListBox1)
    End Sub

    Private Sub FrmQCReport_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try
            DockContainerItem1.Text = "실행 메뉴"
            DockContainerItem2.Text = "조회 조건"

            DockContainerItem5.Text = "불량 상세"
            Bar5.AutoHide = True

            DateTimeInput1.Value = Now
            DateTimeInput2.Value = Now

            'If Query_Combo(Me.ComboBoxEx1, "SELECT model_no FROM tbl_modelmaster WHERE site_id = '" & Site_id & "' and active = 'Y' ORDER BY model_no") = True Then
            'End If
            'Me.ComboBoxEx1.Items.Add("ALL")
            'Me.ComboBoxEx1.Text = "ALL"
            If Query_CheckList(Me.CheckedListBox1, "SELECT model_no FROM tbl_modelmaster WHERE site_id = '" & Site_id & "' and active = 'Y' ORDER BY model_no") = True Then
                Me.CheckedListBox1.Items.Add("ALL")
                Me.CheckedListBox1.Text = "ALL"
            End If

            'If Query_Combo(Me.ComboBoxEx2, "SELECT CODE_NAME FROM tbl_CODEmaster WHERE site_id = '" & Site_id & "' AND CLASS_ID = 'R0001' and CODE_ID IN ('K6000','K3900','K4900') ORDER BY CODE_ID") = True Then
            'End If
            'Me.ComboBoxEx2.Text = "FINAL QC"
            If Query_CheckList(Me.CheckedListBox2, "SELECT CODE_NAME FROM tbl_CODEmaster WHERE site_id = '" & Site_id & "' AND CLASS_ID = 'R0001' and CODE_ID IN ('K6000','K3900','K4900') ORDER BY CODE_ID") = True Then

            End If

            curcell.Separator = ","
            curcell.ShowSeparator = True
            curcell.ShowCurrencySymbol = True
            curcell.CurrencySymbol = "$"

            FpSpread2.ActiveSheet.RowCount = 0

            showHD() 'Display Spread Header
            'create_btn() 'excluded item에 표시할 Checkbox  
            FpSpread2.ActiveSheet.FrozenColumnCount = 4

            ContextMenuBar2.SetContextMenuEx(Me.Chart1, Me.ButtonItem7)
            ' ContextMenuBar2.SetContextMenuEx(Me.FpSpread2, Me.ButtonItem15)

            Formbim_Authority(Me.ButtonItem2, Me.Name, "NEW")
            Formbim_Authority(Me.ButtonItem3, Me.Name, "SAVE")
            Formbim_Authority(Me.ButtonItem4, Me.Name, "DELETE")
            Formbim_Authority(Me.ButtonItem5, Me.Name, "PRINT")
            Formbim_Authority(Me.ButtonItem6, Me.Name, "EXCEL")
            Formbim_Authority(Me.FindBtn, Me.Name, "FIND")

            ButtonItem2.Visible = False
            ButtonItem3.Visible = False
            ButtonItem4.Visible = False

        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message, "ERROR")
        End Try
    End Sub

    Private Sub FindBtn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FindBtn.Click
        Dim rs As ADODB.Recordset
        Dim i, j As Integer
        Dim S2 = Me.FpSpread2.ActiveSheet

        Dim list_qry As String
        Dim fd, td As DateTime
        Dim S3 = Me.FpSpread3.ActiveSheet

        Dim m_qry As String = ""
        m_qry = MODEL_SELECTED(CheckedListBox1)
        Dim T_qry As String = ""
        T_qry = WK_SELECTED(CheckedListBox2)

        fd = DateTimeInput1.Text
        td = DateTimeInput2.Text

        S2.RowCount = 0

        S3.RowCount = 0
        tot_smp = 0

        If T_qry = "" Then
            Modal_Error("공정을 선택하세요.")
            Exit Sub
        End If

        Dim WK As String = Query_RS("SELECT CODE_ID FROM TBL_CODEMASTER WHERE CLASS_ID = 'R0001' AND CODE_NAME = '" & ComboBoxEx2.Text & "'")

        list_qry = "select model_no, " & vbNewLine
        list_qry = list_qry & "(select isnull(sum(CONVERT(INT,ACT_QTY)),0)  from VIEW_LREPAIRMASTER where PROCESS IN (" & T_qry & ") AND SRC_NO IN (SELECT LOT_NO FROM TBL_LOTMASTER WHERE MODEL = A.MODEL_no) AND C_DATE BETWEEN '" & Me.DateTimeInput1.Text & "' and '" & Me.DateTimeInput2.Text & " 23:59:59'  and result like '%합격%') +" & vbNewLine
        list_qry = list_qry & "(select isnull(sum(CONVERT(INT,out_QTY)),0)  from VIEW_LREPAIRMASTER where PROCESS IN (" & T_qry & ") AND SRC_NO IN (SELECT LOT_NO FROM TBL_LOTMASTER WHERE MODEL = A.MODEL_no) AND C_DATE BETWEEN '" & Me.DateTimeInput1.Text & "' and '" & Me.DateTimeInput2.Text & " 23:59:59')," & vbNewLine
        list_qry = list_qry & "(select isnull(sum(CONVERT(INT,OUT_QTY)),0)  from VIEW_LREPAIRMASTER where PROCESS IN (" & T_qry & ") AND SRC_NO IN (SELECT LOT_NO FROM TBL_LOTMASTER WHERE MODEL = A.MODEL_no) AND C_DATE BETWEEN '" & Me.DateTimeInput1.Text & "' and '" & Me.DateTimeInput2.Text & " 23:59:59')," & vbNewLine
        list_qry = list_qry & "0," & vbNewLine

        With FpSpread2.ActiveSheet
            For i = 4 To .ColumnCount - 1
                If i < .ColumnCount - 1 Then
                    list_qry = list_qry & "(select isnull(sum(CONVERT(INT,OUT_QTY)),0)  from VIEW_LREPAIRMASTER where PROCESS IN (" & T_qry & ") AND DEF_CD = '" & .ColumnHeader.Cells(1, i).Text & "' AND SRC_NO IN (SELECT LOT_NO FROM TBL_LOTMASTER WHERE MODEL = A.MODEL_no) AND C_DATE BETWEEN '" & Me.DateTimeInput1.Text & "' and '" & Me.DateTimeInput2.Text & " 23:59:59')," & vbNewLine
                Else
                    list_qry = list_qry & "(select isnull(sum(CONVERT(INT,OUT_QTY)),0)  from VIEW_LREPAIRMASTER where PROCESS IN (" & T_qry & ") AND DEF_CD = '" & .ColumnHeader.Cells(1, i).Text & "' AND SRC_NO IN (SELECT LOT_NO FROM TBL_LOTMASTER WHERE MODEL = A.MODEL_no) AND C_DATE BETWEEN '" & Me.DateTimeInput1.Text & "' and '" & Me.DateTimeInput2.Text & " 23:59:59')" & vbNewLine
                End If
            Next

        End With

        list_qry = list_qry & "from TBL_modelmaster A" & vbNewLine
        list_qry = list_qry & "where site_id = '" & Site_id & "'" & vbNewLine
        'If ComboBoxEx1.Text <> "ALL" Then
        '    list_qry = list_qry & "AND model_no = '" & ComboBoxEx1.Text & "'" & vbNewLine
        'End If
        If m_qry <> "" Then
            list_qry = list_qry & "AND a.MODEL_no IN (" & m_qry & ")" & vbNewLine
        End If



        list_qry = list_qry & "AND active = 'Y'" & vbNewLine
        list_qry = list_qry & "ORDER BY MODEL_no" & vbNewLine


        rs = Query_RS_ALL(list_qry)


        If rs Is Nothing Then
            Exit Sub
        End If

        For i = 0 To rs.RecordCount - 1
            S2.RowCount += 1
            S2.Rows(S2.RowCount - 1).Locked = True
            For j = 0 To S2.ColumnCount - 1
                If j = 0 Then
                    Spread_TxtType(Me.FpSpread2, i, j, True, 99)
                ElseIf j = 3 Then
                    Spread_PercentType(Me.FpSpread2, i, j)
                Else
                    Spread_NumType(Me.FpSpread2, i, j)
                End If
                S2.SetValue(i, j, rs(j).Value)
            Next
            S2.Cells(i, 3).Formula = "C" & i + 1 & "/B" & i + 1
            rs.MoveNext()
        Next

        With FpSpread2.ActiveSheet
            Spread_AutoCol(FpSpread2)
            .Columns(1, .ColumnCount - 1).Width = 50
            .Columns(1, 3).Width = 80
        End With

        total_rej()   'QA Reject result의 Total Summary를 마지막 Row를 추가하여 표시
        chart_tot()   'Chart에 모델별 QA reject 수량 표시

        rs.Close()
        'query_qcr()

    End Sub

    Private Sub ButtonItem6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonItem6.Click
        If save_excel = "" Then
            MessageBox.Show("Select Spread for Saving!!")
            Exit Sub
        End If

        If save_excel = "FpSpread1" Then

        ElseIf save_excel = "FpSpread2" Then
            File_Save(SaveFileDialog1, FpSpread2)
        ElseIf save_excel = "FpSpread4" Then

        ElseIf save_excel = "FpSpread3" Then
            File_Save(SaveFileDialog1, FpSpread3)
        Else
            MessageBox.Show("Select Spread for Saving!!")
        End If

    End Sub

    Private Sub showHD()
        Dim i As Integer
        Dim def_rs As New ADODB.Recordset

        'Spread에 HEADER 출력
        Try
            With Me.FpSpread2.ActiveSheet
                .ColumnHeader.RowCount = 2
                .ColumnHeader.Columns.Count = 4
                .ColumnHeader.Cells(0, 0).RowSpan = 2
                .ColumnHeader.Cells(0, 0).Text = "모델"
                .ColumnHeader.Cells(0, 1).RowSpan = 2
                .ColumnHeader.Cells(0, 1).Text = "검사수량"
                .ColumnHeader.Cells(0, 2).ColumnSpan = 2
                .ColumnHeader.Cells(0, 2).Text = "불량"
                .ColumnHeader.Cells(1, 2).Text = "불량수량"
                .ColumnHeader.Cells(1, 3).Text = "불량율"

                def_rs = Query_RS_ALL("select DISTINCT CODE_ID from tbl_DEFECT where ACTIVE = 'Y' order by CODE_ID")

                For i = 0 To def_rs.RecordCount - 1
                    .ColumnCount = .ColumnCount + 1
                    .ColumnHeader.Cells(1, i + 4).Text = def_rs(0).Value
                    def_rs.MoveNext()
                Next

                .ColumnHeader.Cells(0, 4).ColumnSpan = def_rs.RecordCount
                .ColumnHeader.Cells(0, 4).Text = "불량코드"

            End With



            With FpSpread3.ActiveSheet
                .ColumnCount = 7
                .Columns(0).Label = "모델"
                .Columns(1).Label = "생산LOT 번호"
                .Columns(1).CellType = textcell

                .Columns(2).Label = "불량코드"
                .Columns(3).Label = "불량명"
                .Columns(4).Label = "불량수량"
                .Columns(4).CellType = intcell
                intcell.DecimalPlaces = 0

                .Columns(5).Label = "검사자"
                .Columns(6).Label = "불량판정일"

                .RowCount = 0

            End With

        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message, "ERROR")
        End Try
    End Sub

    Sub total_rej()
        '마지막 Row에 Reject result Total 표시
        Try
            Dim i As Integer
            With Me.FpSpread2.ActiveSheet
                .RowCount += 1
                .Rows(.RowCount - 1).BackColor = Color.Yellow
                .Rows(.RowCount - 1).ForeColor = Color.DarkBlue
                .SetText(.RowCount - 1, 0, "TOTAL")
                Spread_NumType(Me.FpSpread2, .RowCount - 1, 1)
                .Cells(.RowCount - 1, 1).Formula = "sum(B1:B" & .RowCount - 1 & ")"
                tot_smp = .GetValue(.RowCount - 1, 1)
                Spread_NumType(Me.FpSpread2, .RowCount - 1, 2)
                .Cells(.RowCount - 1, 2).Formula = "sum(C1:C" & .RowCount - 1 & ")"
                Spread_PercentType(Me.FpSpread2, .RowCount - 1, 3)
                .Cells(.RowCount - 1, 3).Formula = "C" & .RowCount & "/B" & .RowCount

                For i = 4 To .ColumnCount - 1
                    Spread_NumType(Me.FpSpread2, .RowCount - 1, i)
                    If i < 26 Then
                        .Cells(.RowCount - 1, i).Formula = "sum(" & Chr(65 + i) & "1:" & Chr(65 + i) & .RowCount - 1 & ")"
                    Else
                        .Cells(.RowCount - 1, i).Formula = "sum(A" & Chr(65 + i - 26) & "1:A" & Chr(65 + i - 26) & .RowCount - 1 & ")"
                    End If
                Next
            End With
        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message, "ERROR")
        End Try
    End Sub

    Private Sub chart_tot()
        '모델별 QA Reject total 을 차트에 표시
        Dim S2 = Me.FpSpread2.ActiveSheet
        Chart_Spread_All_1(Chart1, False, 2, DataVisualization.Charting.SeriesChartType.Column, Me.FpSpread2, New String() {"Total"}, 0, New Integer() {2}, True)
    End Sub

    Private Sub chart_row(ByVal rowidx As Integer)
        'Row 클릭시 선택모델의 Reject code 별 수량을 차트에 표시
        Try
            Dim S2 = Me.FpSpread2.ActiveSheet
            Dim colidx(S2.ColumnCount - 4 - 1), i As Integer
            Dim seriestxt As String

            seriestxt = S2.GetValue(rowidx, 0)

            For i = 0 To S2.ColumnCount - 4 - 1
                colidx(i) = i + 4
            Next

            Chart_Spread_Col(Chart1, False, 2, DataVisualization.Charting.SeriesChartType.Column, Me.FpSpread2, seriestxt, rowidx, colidx, 1)

        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message, "ERROR")
        End Try
    End Sub
    Private Sub FpSpread1_CellClick(ByVal sender As System.Object, ByVal e As FarPoint.Win.Spread.CellClickEventArgs)
        save_excel = "FpSpread1"
    End Sub

    Private Sub FpSpread4_CellClick(ByVal sender As System.Object, ByVal e As FarPoint.Win.Spread.CellClickEventArgs)
        save_excel = "FpSpread4"
    End Sub
    Private Sub FpSpread3_CellClick(ByVal sender As System.Object, ByVal e As FarPoint.Win.Spread.CellClickEventArgs) Handles FpSpread3.CellClick
        save_excel = "FpSpread3"
    End Sub

    Private Sub FpSpread2_CellClick(ByVal sender As Object, ByVal e As FarPoint.Win.Spread.CellClickEventArgs) Handles FpSpread2.CellClick
        save_excel = "FpSpread2"
        Dim S2 = Me.FpSpread2.ActiveSheet
        chart_row(e.Row) 'Row 클릭시 선택모델의 Reject code 별 수량을 차트에 표시
    End Sub

    Private Sub ButtonItem12_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        chart_tot() '모델별 QA Reject total 을 차트에 표시
    End Sub

    Private Sub ButtonItem13_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonItem13.Click
        Chart_Print(Me.Chart1, True)
    End Sub

    'Private Sub ButtonItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonItem1.Click
    '    If Spread_Print(Me.FpSpread4, "QC Fail Summary", 0) = False Then
    '        MsgBox("Fail to Print")
    '    End If
    'End Sub

    'Private Sub ButtonItem10_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonItem10.Click
    '    File_Save(SaveFileDialog1, FpSpread4)
    'End Sub

    Private Sub ButtonItem5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonItem5.Click
        If Spread_Print(Me.FpSpread2, "QA Reject Rate", 0) = False Then
            MsgBox("Fail to Print")
        End If
    End Sub

    Private Sub FpSpread2_CellDoubleClick(ByVal sender As Object, ByVal e As FarPoint.Win.Spread.CellClickEventArgs) Handles FpSpread2.CellDoubleClick

        Dim DEF As String

        Dim S2 = Me.FpSpread2.ActiveSheet
        'Dim S4 = Me.FpSpread4.ActiveSheet
        Dim S3 = Me.FpSpread3.ActiveSheet
        save_excel = "FpSpread2"

        S3.RowCount = 0

        Dim T_qry As String = ""
        T_qry = WK_SELECTED(CheckedListBox2)


        With FpSpread2.ActiveSheet
            Dim qry As String
            If e.Column = 2 Then
                DEF = ""
            ElseIf e.Column >= 4 Then
                DEF = .ColumnHeader.Cells(1, e.Column).Text
            Else
                Exit Sub
            End If

            Dim WK As String = Query_RS("SELECT CODE_ID FROM TBL_CODEMASTER WHERE CLASS_ID = 'R0001' AND CODE_NAME = '" & ComboBoxEx2.Text & "'")

            qry = "select B.MODEL, B.LOT_NO, A.def_cd, (SELECT CODE_NAME FROM TBL_DEFECT WHERE CODE_ID = a.DEF_CD), A.OUT_QTY, (SELECT EMP_NM FROM TBL_EMPMASTER WHERE SITE_ID = A.SITE_ID AND EMP_NO = A.C_PERSON) AS EMP, a.C_DATE	" & vbNewLine
            qry = qry & "from view_Lrepairmaster A, TBL_LOTmaster b" & vbNewLine
            qry = qry & "where a.c_date between '" & DateTimeInput1.Text & "' and '" & DateTimeInput2.Text & "' + ' 23:59:59'" & vbNewLine
            qry = qry & "and A.PROCESS IN (" & T_qry & ") " & vbNewLine
            qry = qry & "and isnull(a.OUT_QTY,0) > 0" & vbNewLine
            qry = qry & "and a.src_no = b.LOT_NO" & vbNewLine
            If .GetValue(e.Row, 0) <> "TOTAL" Then
                qry = qry & "AND a.DEF_CD LIKE '" & DEF & "%'" & vbNewLine
                qry = qry & "and b.model = '" & .GetValue(e.Row, 0) & "'" & vbNewLine
                '                QRY = QRY & "and a.src_no in (select obid from " & vw_ens & " where model = '" & ComboBoxEx1.Text & "')" & vbNewLine
            End If
            qry = qry & "ORDER BY a.DEF_CD, a.C_DATE" & vbNewLine

            If Query_Spread(FpSpread3, qry, 1) = True Then
                '                If Query_Spread(FpSpread3, "EXEC SP_FRMLQCREPORT1_DETAIL '" & Site_id & "','" & .GetValue(e.Row, 0) & "','" & DateTimeInput1.Text & "','" & DateTimeInput2.Text & "','" & DEF & "'", 1) = True Then
                .Cells(0, 0, .RowCount - 1, .ColumnCount - 1).Locked = True
                Spread_AutoCol(FpSpread3)
            End If

        End With

        '   query_qcr()


        If Bar5.AutoHide = True Then
            Bar5.AutoHide = False
        End If

    End Sub

    'Function query_qcr() As Boolean
    '    '        Dim RS As ADODB.Recordset
    '    Dim I, J As Integer
    '    Dim fd, td As DateTime
    '    Dim fdt, tdt As String


    '    Dim S2 = Me.FpSpread2.ActiveSheet
    '    Dim S4 = Me.FpSpread4.ActiveSheet
    '    Dim S3 = Me.FpSpread3.ActiveSheet

    '    S1.RowCount = 0
    '    'S3.RowCount = 0
    '    S4.RowCount = 0

    '    fd = DateTimeInput1.Text
    '    td = DateTimeInput2.Text

    '    fdt = Query_RS("select convert(varchar(8),convert(datetime, '" & DateTimeInput1.Text & "'),112)")
    '    tdt = Query_RS("select convert(varchar(8),convert(datetime, '" & DateTimeInput2.Text & "'),112)")

    '    Dim cnt As Integer
    '    Dim qry As String

    '    cnt = Query_RS("select datediff(day, '" & DateTimeInput1.Text & "', '" & DateTimeInput2.Text & "')")

    '    With S1
    '        .ColumnCount = 4
    '        .FrozenColumnCount = 4

    '        .Columns(3).CellType = intcell

    '        qry = "select '', CODE_ID, isnull((select code_name from tbl_defect where code_id = a.DEF_CD),DEF_CD)," & vbNewLine
    '        qry = qry & "isnull((select sum(qty) from tbl_qcfail where model = '" & S2.GetValue(S2.ActiveRowIndex, 0) & "' and qcdate between convert(varchar(8),dateadd(day,0,'" & fd & "'),112) and convert(varchar(8),dateadd(day, " & cnt & ", '" & fd & "'),112) and def_cd = a.DEF_CD),0)," & vbNewLine

    '        For I = 0 To cnt
    '            .ColumnCount = .ColumnCount + 1
    '            .Columns(.ColumnCount - 1).CellType = intcell
    '            .ColumnHeader.Columns(.ColumnCount - 1).Label = Query_RS("select convert(varchar(10), dateadd(day, " & I & ",'" & fd & "'), 101)") ' Query_RS("select left(convert(varchar(10), dateadd(day, " & i & ",'" & fd & "'), 101),5)")
    '            If I < cnt Then
    '                qry = qry & "isnull((select sum(qty) from tbl_qcfail where model = '" & S2.GetValue(S2.ActiveRowIndex, 0) & "' and qcdate = convert(varchar(10), dateadd(day, " & I & ",'" & fd & "'), 112) and def_cd = a.DEF_CD),0), " & vbNewLine
    '            Else
    '                qry = qry & "isnull((select sum(qty) from tbl_qcfail where model = '" & S2.GetValue(S2.ActiveRowIndex, 0) & "' and qcdate = convert(varchar(10), dateadd(day, " & I & ",'" & fd & "'), 112) and def_cd = a.DEF_CD),0) " & vbNewLine
    '            End If
    '        Next
    '        qry = qry & "from tbl_DEFECT a" & vbNewLine
    '        qry = qry & "where ACTIVE = 'Y'" & vbNewLine
    '        qry = qry & "order by CODE_ID"

    '        If Query_Spread(FpSpread1, qry, 1) = True Then
    '            .Cells(0, 0).RowSpan = .RowCount
    '            .Cells(0, 0).Text = S2.GetValue(S2.ActiveRowIndex, 0) '"LINE"
    '            .Cells(0, 0).HorizontalAlignment = FarPoint.Win.Spread.CellHorizontalAlignment.Center
    '            .Cells(0, 0).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
    '            .SheetName = S2.GetValue(S2.ActiveRowIndex, 0)

    '            .RowCount = .RowCount + 1
    '            .Cells(.RowCount - 1, 0).Text = "Reject Qty."
    '            .Cells(.RowCount - 1, 0).ColumnSpan = 3
    '            .Cells(.RowCount - 1, 0).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
    '            .Cells(.RowCount - 1, 0).HorizontalAlignment = FarPoint.Win.Spread.CellHorizontalAlignment.Center
    '            .Rows(.RowCount - 1).BackColor = Color.DarkOrange

    '            For I = 3 To .ColumnCount - 1
    '                For J = 0 To .RowCount - 2
    '                    .SetValue(.RowCount - 1, I, .GetValue(J, I) + .GetValue(.RowCount - 1, I))
    '                Next
    '            Next

    '            .RowCount = .RowCount + 1
    '            .Cells(.RowCount - 1, 0).Text = "Total Qty."
    '            .Cells(.RowCount - 1, 0).ColumnSpan = 3
    '            .Cells(.RowCount - 1, 0).HorizontalAlignment = FarPoint.Win.Spread.CellHorizontalAlignment.Center
    '            .Rows(.RowCount - 1).BackColor = Color.SkyBlue

    '            .SetValue(.RowCount - 1, 3, Query_RS("select isnull(sum(pass_qty),0) + isnull(sum(Fail_qty),0) from tbl_qcresult where model = '" & S2.GetValue(S2.ActiveRowIndex, 0) & "' and c_date between '" & .ColumnHeader.Columns(4).Label & "' and '" & .ColumnHeader.Columns(.ColumnCount - 1).Label & "'"))
    '            For I = 4 To .ColumnCount - 1
    '                .SetValue(.RowCount - 1, I, Query_RS("select isnull(pass_qty,0) + isnull(Fail_qty,0) from tbl_qcresult where model = '" & S2.GetValue(S2.ActiveRowIndex, 0) & "' and c_date = '" & .ColumnHeader.Columns(I).Label & "'"))
    '                If .GetValue(.RowCount - 1, I) = "" Then
    '                    .SetValue(.RowCount - 1, I, 0)
    '                End If
    '            Next

    '            .RowCount = .RowCount + 1
    '            .Cells(.RowCount - 1, 0).Text = "Reject Rate"
    '            .Cells(.RowCount - 1, 0).ColumnSpan = 3
    '            .Cells(.RowCount - 1, 0).HorizontalAlignment = FarPoint.Win.Spread.CellHorizontalAlignment.Center
    '            .Rows(.RowCount - 1).BackColor = Color.DarkOrange
    '            .Cells(.RowCount - 1, 3, .RowCount - 1, .ColumnCount - 1).CellType = percell

    '            For I = 3 To .ColumnCount - 1
    '                If .GetValue(.RowCount - 2, I) > 0 Then
    '                    .SetValue(.RowCount - 1, I, .GetValue(.RowCount - 3, I) / .GetValue(.RowCount - 2, I))
    '                End If
    '            Next

    '            .Cells(0, 0, .RowCount - 1, .ColumnCount - 1).Border = New LineBorder(Color.Black, 1, True, True, True, True)
    '            .ColumnHeader.Cells(0, 0, .ColumnHeader.RowCount - 1, .ColumnCount - 1).Border = New LineBorder(Color.Black, 1, True, True, True, True)


    '            Spread_AutoCol(FpSpread1)
    '        End If
    '    End With

    'End Function


End Class