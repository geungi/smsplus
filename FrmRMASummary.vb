Imports System
Imports System.Drawing
Imports System.Collections
Imports System.ComponentModel
Imports System.Windows.Forms
Imports FarPoint.Win.Spread
Imports FarPoint.Win

Public Class FrmRMASummary

    Private tot_smp As Integer = 0  'total sample size
    Private xpos As Integer = 101  'checkbox 초기 X축 위치 

    Private Sub FrmQCReport_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try
            DockContainerItem1.Text = "Form Menu"
            DockContainerItem2.Text = "Retrieve Condition"
            DockContainerItem3.Text = "RMA Fail Detail"

            DateTimeInput1.Value = Now
            DateTimeInput2.Value = Now

            If Query_Combo(Me.ComboBoxEx1, "SELECT model_no FROM tbl_modelmaster WHERE site_id = '" & Site_id & "' and active = 'Y' ORDER BY model_no") = True Then
            End If
            Me.ComboBoxEx1.Items.Add("ALL")
            Me.ComboBoxEx1.Text = "ALL"

            curcell.Separator = ","
            curcell.ShowSeparator = True
            curcell.ShowCurrencySymbol = True
            curcell.CurrencySymbol = "\"

            FpSpread2.ActiveSheet.RowCount = 0
            FpSpread1.ActiveSheet.RowCount = 0

            showHD() 'Display Spread Header

            FpSpread2.ActiveSheet.FrozenColumnCount = 4

            ContextMenuBar2.SetContextMenuEx(Me.Chart1, Me.ButtonItem7)
            ContextMenuBar2.SetContextMenuEx(Me.FpSpread2, Me.ButtonItem15)

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

        Dim S1 = Me.FpSpread1.ActiveSheet
        Dim S2 = Me.FpSpread2.ActiveSheet
        Dim list_qry As String
        Dim fd, td As DateTime

        fd = DateTimeInput1.Text
        td = DateTimeInput2.Text

        S1.RowCount = 0
        S2.RowCount = 0
 

        list_qry = "SELECT RESERV5, MODEL," & vbNewLine
        list_qry = list_qry & "ISNULL((SELECT COUNT(RESERV5) FROM VIEW_FREPAIRMASTER WHERE RESERV5 = A.RESERV5 /*AND RESERV1 = 'W9400'*/ AND ISNULL(DEF_CD,'') NOT IN ('N/A','') AND RESERV4 = A.MODEL and RESULT = 'RMA INSPECTION'),0)," & vbNewLine
        list_qry = list_qry & "		COUNT(RESERV5) ," & vbNewLine
        list_qry = list_qry & "		ISNULL((SELECT COUNT(RESERV5) FROM VIEW_FREPAIRMASTER WHERE RESERV5 = A.RESERV5 /*AND RESERV1 = 'W9400'*/ AND ISNULL(DEF_CD,'') NOT IN ('N/A','') AND RESERV4 = A.MODEL and RESULT = 'RMA INSPECTION' ),0)*1.00/COUNT(RESERV5) ," & vbNewLine
        list_qry = list_qry & "		ISNULL((SELECT COUNT(RESERV5) FROM VIEW_FREPAIRMASTER WHERE RESERV5 = A.RESERV5 /*AND RESERV1 = 'W9400'*/ AND DEF_CD = 'D2' AND ISNULL(DEF_CD,'') NOT IN ('N/A','') AND RESERV4 = A.MODEL and RESULT = 'RMA INSPECTION'),0)," & vbNewLine
        list_qry = list_qry & "		ISNULL((SELECT COUNT(RESERV5) FROM VIEW_FREPAIRMASTER WHERE RESERV5 = A.RESERV5 /*AND RESERV1 = 'W9400'*/ AND DEF_CD = 'D3' AND ISNULL(DEF_CD,'') NOT IN ('N/A','') AND RESERV4 = A.MODEL and RESULT = 'RMA INSPECTION'),0)," & vbNewLine
        list_qry = list_qry & "		ISNULL((SELECT COUNT(RESERV5) FROM VIEW_FREPAIRMASTER WHERE RESERV5 = A.RESERV5 /*AND RESERV1 = 'W9400'*/ AND DEF_CD = 'D8' AND ISNULL(DEF_CD,'') NOT IN ('N/A','') AND RESERV4 = A.MODEL and RESULT = 'RMA INSPECTION'),0)," & vbNewLine
        list_qry = list_qry & "		ISNULL((SELECT COUNT(RESERV5) FROM VIEW_FREPAIRMASTER WHERE RESERV5 = A.RESERV5 /*AND RESERV1 = 'W9400'*/ AND DEF_CD = 'DL' AND ISNULL(DEF_CD,'') NOT IN ('N/A','') AND RESERV4 = A.MODEL and RESULT = 'RMA INSPECTION'),0)," & vbNewLine
        list_qry = list_qry & "		ISNULL((SELECT COUNT(RESERV5) FROM VIEW_FREPAIRMASTER WHERE RESERV5 = A.RESERV5 /*AND RESERV1 = 'W9400'*/ AND DEF_CD = 'DT' AND ISNULL(DEF_CD,'') NOT IN ('N/A','') AND RESERV4 = A.MODEL and RESULT = 'RMA INSPECTION'),0)," & vbNewLine
        list_qry = list_qry & "		ISNULL((SELECT COUNT(RESERV5) FROM VIEW_FREPAIRMASTER WHERE RESERV5 = A.RESERV5 /*AND RESERV1 = 'W9400'*/ AND DEF_CD = 'FJ' AND ISNULL(DEF_CD,'') NOT IN ('N/A','') AND RESERV4 = A.MODEL and RESULT = 'RMA INSPECTION'),0)," & vbNewLine
        list_qry = list_qry & "		ISNULL((SELECT COUNT(RESERV5) FROM VIEW_FREPAIRMASTER WHERE RESERV5 = A.RESERV5 /*AND RESERV1 = 'W9400'*/ AND DEF_CD = 'W3' AND ISNULL(DEF_CD,'') NOT IN ('N/A','') AND RESERV4 = A.MODEL and RESULT = 'RMA INSPECTION'),0)," & vbNewLine
        list_qry = list_qry & "		ISNULL((SELECT COUNT(RESERV5) FROM VIEW_FREPAIRMASTER WHERE RESERV5 = A.RESERV5 /*AND RESERV1 = 'W9400'*/ AND DEF_CD = 'WA' AND ISNULL(DEF_CD,'') NOT IN ('N/A','') AND RESERV4 = A.MODEL and RESULT = 'RMA INSPECTION'),0)," & vbNewLine
        list_qry = list_qry & "		ISNULL((SELECT COUNT(RESERV5) FROM VIEW_FREPAIRMASTER WHERE RESERV5 = A.RESERV5 /*AND RESERV1 = 'W9400'*/ AND DEF_CD = 'X2' AND ISNULL(DEF_CD,'') NOT IN ('N/A','') AND RESERV4 = A.MODEL and RESULT = 'RMA INSPECTION'),0)" & vbNewLine
        list_qry = list_qry & "from VIEW_FESNMASTER A" & vbNewLine
        list_qry = list_qry & "where site_id = '" & Site_id & "'" & vbNewLine
        list_qry = list_qry & "AND SUBSTRING(RESERV5,3,8) BETWEEN '" & Format(DateTimeInput1.Value, "yyyyMMdd") & "' AND '" & Format(DateTimeInput2.Value, "yyyyMMdd") & "'" & vbNewLine

        If ComboBoxEx1.Text <> "ALL" Then
            list_qry = list_qry & "AND model = '" & ComboBoxEx1.Text & "'" & vbNewLine
        End If
        list_qry = list_qry & "GROUP BY RESERV5, MODEL" & vbNewLine
        list_qry = list_qry & "ORDER BY RESERV5, MODEL" & vbNewLine


        If Query_Spread_USA(FpSpread2, list_qry, 1) = True Then
            With FpSpread2.ActiveSheet
                .RowCount = .RowCount + 1
                .Cells(.RowCount - 1, 5, .RowCount - 1, .ColumnCount - 1).CellType = intcell
                .Cells(.RowCount - 1, 2, .RowCount - 1, 3).CellType = intcell
                .Cells(.RowCount - 1, 4, .RowCount - 1, 4).CellType = percell
                .Rows(.RowCount - 1).BackColor = Color.Yellow
                .SetText(.RowCount - 1, 0, "TOTAL")
                SPREAD_ROW_TOTAL(FpSpread2, 2, 3, 1)
                SPREAD_ROW_TOTAL(FpSpread2, 5, .ColumnCount - 1, 1)
                .SetValue(.RowCount - 1, 4, .GetValue(.RowCount - 1, 2) / .GetValue(.RowCount - 1, 3))


                Spread_AutoCol(FpSpread2)

                FpSpread2.ActiveSheet.Columns(2, 13).Width = 60
            End With
        End If



        chart_tot()   'Chart에 모델별 QA reject 수량 표시


    End Sub

    Private Sub ButtonItem6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonItem6.Click
        If save_excel = "" Then
            MessageBox.Show("Select Spread for Saving!!")
            Exit Sub
        End If

        If save_excel = "FpSpread1" Then
            File_Save(SaveFileDialog1, FpSpread1)
        ElseIf save_excel = "FpSpread2" Then
            File_Save(SaveFileDialog1, FpSpread2)
        ElseIf save_excel = "FpSpread4" Then
            '            File_Save(SaveFileDialog1, FpSpread4)
        ElseIf save_excel = "FpSpread3" Then
            '           File_Save(SaveFileDialog1, FpSpread3)
        Else
            MessageBox.Show("Select Spread for Saving!!")
        End If

    End Sub

    Private Sub showHD()

        Dim def_rs As New ADODB.Recordset

        'Spread에 HEADER 출력
        Try
            With Me.FpSpread2.ActiveSheet
                .ColumnHeader.RowCount = 2
                .ColumnHeader.Columns.Count = 14
                .ColumnHeader.Cells(0, 0).RowSpan = 2
                .ColumnHeader.Cells(0, 0).Text = "SHIP NO"

                .ColumnHeader.Cells(0, 1).RowSpan = 2
                .ColumnHeader.Cells(0, 1).Text = "MODEL"


                .ColumnHeader.Cells(0, 2).RowSpan = 2
                .ColumnHeader.Cells(0, 2).Text = "Fail TOTAL"
                .ColumnHeader.Cells(0, 3).RowSpan = 2
                .ColumnHeader.Cells(0, 3).Text = "Ship TOTAL"
                .ColumnHeader.Cells(0, 4).RowSpan = 2
                .ColumnHeader.Cells(0, 4).Text = "FAIL RATE"

                .ColumnHeader.Cells(0, 5).ColumnSpan = 9
                .ColumnHeader.Cells(0, 5).Text = "Defect Code"
                .ColumnHeader.Cells(1, 5).Text = "D2"
                .ColumnHeader.Cells(1, 6).Text = "D3"
                .ColumnHeader.Cells(1, 7).Text = "D8"
                .ColumnHeader.Cells(1, 8).Text = "DL"
                .ColumnHeader.Cells(1, 9).Text = "DT"
                .ColumnHeader.Cells(1, 10).Text = "FJ"
                .ColumnHeader.Cells(1, 11).Text = "W3"
                .ColumnHeader.Cells(1, 12).Text = "WA"
                .ColumnHeader.Cells(1, 13).Text = "X2"

                .Columns(2, 3).CellType = intcell
                .Columns(4).CellType = percell
                .Columns(5, 13).CellType = intcell

                .RowCount = 0

            End With


            With FpSpread1.ActiveSheet
                .ColumnCount = 5
                .Columns(0).Label = "MODEL"
                .Columns(1).Label = "SHIP NO"
                .Columns(2).Label = "FLIP ID"
                .Columns(3).Label = "DEFECT"
                .Columns(4).Label = "DESC"

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
    Private Sub FpSpread1_CellClick(ByVal sender As System.Object, ByVal e As FarPoint.Win.Spread.CellClickEventArgs) Handles FpSpread1.CellClick
        save_excel = "FpSpread1"
    End Sub

    Private Sub FpSpread4_CellClick(ByVal sender As System.Object, ByVal e As FarPoint.Win.Spread.CellClickEventArgs)
        save_excel = "FpSpread4"
    End Sub
    Private Sub FpSpread3_CellClick(ByVal sender As System.Object, ByVal e As FarPoint.Win.Spread.CellClickEventArgs)
        save_excel = "FpSpread3"
    End Sub

    Private Sub FpSpread2_CellClick(ByVal sender As Object, ByVal e As FarPoint.Win.Spread.CellClickEventArgs) Handles FpSpread2.CellClick
        save_excel = "FpSpread2"
        Dim S2 = Me.FpSpread2.ActiveSheet
        chart_row(e.Row) 'Row 클릭시 선택모델의 Reject code 별 수량을 차트에 표시
    End Sub

    Private Sub ButtonItem12_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonItem12.Click
        chart_tot() '모델별 QA Reject total 을 차트에 표시
    End Sub

    Private Sub ButtonItem13_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonItem13.Click
        Chart_Print(Me.Chart1, True)
    End Sub



    Private Sub ButtonItem5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonItem5.Click
        If Spread_Print(Me.FpSpread2, "QA Reject Rate", 0) = False Then
            MsgBox("Fail to Print")
        End If
    End Sub

 

    Private Sub FpSpread2_CellDoubleClick(ByVal sender As Object, ByVal e As FarPoint.Win.Spread.CellClickEventArgs) Handles FpSpread2.CellDoubleClick

        Dim DEF As String
        Dim S1 = Me.FpSpread1.ActiveSheet
        Dim S2 = Me.FpSpread2.ActiveSheet
        save_excel = "FpSpread2"


        With FpSpread2.ActiveSheet
            Dim qry As String
            If e.Column = 2 Then
                DEF = ""
            ElseIf e.Column >= 4 Then
                DEF = .ColumnHeader.Cells(1, e.Column).Text
            Else
                Exit Sub
            End If

            qry = "select RESERV4, RESERV5, ESN, DEF_CD, (select code_name from tbl_defect where code_id = a.def_cd)	" & vbNewLine
            qry = qry & "from view_Frepairmaster A" & vbNewLine
            qry = qry & "where SITE_ID = '" & Site_id & "'" & vbNewLine

            'qry = qry & "and a.RESERV1 = 'W9400'" & vbNewLine
            qry = qry & "and a.RESULT = 'RMA INSPECTION'" & vbNewLine
            If .ActiveColumnIndex > 4 Then
                qry = qry & "and isnull(a.DEF_cd,'N/A') = '" & .ColumnHeader.Columns(.ActiveColumnIndex).Label & "'" & vbNewLine
            ElseIf .ActiveColumnIndex = 2 Then
                qry = qry & "and isnull(a.DEF_cd,'N/A') NOT IN ('N/A', '')" & vbNewLine
            End If


            If .GetValue(e.Row, 0) <> "TOTAL" Then
                qry = qry & "and RESERV5 = '" & .GetValue(.ActiveRowIndex, 0) & "'" & vbNewLine
                qry = qry & "AND RESERV4 = '" & .GetValue(.ActiveRowIndex, 1) & "'" & vbNewLine
            Else
                Dim MODEL As String = ""
                For I As Integer = 0 To .RowCount - 2
                    MODEL = MODEL & "'" & .GetValue(I, 0) & "',"
                Next
                qry = qry & "AND RESERV5 IN (" & Mid(MODEL, 1, Len(MODEL) - 1) & ")" & vbNewLine

                If ComboBoxEx1.Text <> "ALL" Then
                    qry = qry & "AND RESERV4 = '" & ComboBoxEx1.Text & "'" & vbNewLine
                End If
            End If

            qry = qry & "ORDER BY RESERV4, RESERV5, DEF_CD, ESN" & vbNewLine

            If Query_Spread_USA(FpSpread1, qry, 1) = True Then
                '                If Query_Spread(FpSpread3, "EXEC SP_FRMLQCREPORT1_DETAIL '" & Site_id & "','" & .GetValue(e.Row, 0) & "','" & DateTimeInput1.Text & "','" & DateTimeInput2.Text & "','" & DEF & "'", 1) = True Then
                .Cells(0, 0, .RowCount - 1, .ColumnCount - 1).Locked = True
                Spread_AutoCol(FpSpread1)
            End If

        End With


    End Sub

 


End Class