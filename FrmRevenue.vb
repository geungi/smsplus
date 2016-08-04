Public Class FrmRevenue

    Private Sub CheckedListBox1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles CheckedListBox1.LostFocus
        TextBoxDropDown1.Text = MODEL_SELECTED(CheckedListBox1)
        FpSpread1.ActiveSheet.RowCount = 0
    End Sub

    Private Sub FrmRevenue_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        DockContainerItem2.Text = "실행 메뉴"
        DockContainerItem1.Text = "조회 조건"

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

        Me.ComboBoxEx2.Items.Add("YEAR")
        Me.ComboBoxEx2.Items.Add("MONTH")
        Me.ComboBoxEx2.Items.Add("WEEK")
        Me.ComboBoxEx2.Text = "YEAR"

        With FpSpread1.ActiveSheet
            .RowCount = 0
            .Protect = False
        End With

        Me.Chart1.Visible = False

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


    End Sub

    Private Sub ComboBoxEx2_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBoxEx2.SelectedIndexChanged
        Try

            Dim totrowcnt As Integer = Me.FpSpread1.ActiveSheet.RowCount
            If Me.FpSpread1.ActiveSheet.RowCount > 0 Then
                Me.FpSpread1.ActiveSheet.Reset()
                Me.FpSpread1.ActiveSheet.Rows.Clear()
            End If

            DateTimeInput1.Value = Now
            DateTimeInput2.Value = Now
            Create_HD()

        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message, "ERROR")
        End Try
    End Sub

    Private Sub FindBtn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FindBtn.Click
        '     Try
        Dim rev_rs As ADODB.Recordset
        Dim i, j As Integer
        Dim QrySrt As String = ""
        Dim intcell_r As New FarPoint.Win.Spread.CellType.NumberCellType()
        Dim deccell_r As New FarPoint.Win.Spread.CellType.NumberCellType()
        Dim fWeek As Integer

        Dim Qry As String = ""
        Dim m_qry As String = ""
        m_qry = MODEL_SELECTED(CheckedListBox1)

        curcell.CurrencySymbol = "$"

        Select Case ComboBoxEx2.Text
            Case "YEAR"
                DateTimeInput1.Value = CDate(Year(DateTimeInput1.Value) & "-01-01")
                DateTimeInput2.Value = CDate(Year(DateTimeInput1.Value) & "-12-31")
            Case "MONTH"
                DateTimeInput1.Value = CDate(Month_Day(DateTimeInput1.Value)(0))
                DateTimeInput2.Value = CDate(Month_Day(DateTimeInput1.Value)(1))
            Case "WEEK"
        End Select

        With FpSpread1.ActiveSheet

            .ClearRange(0, 0, .RowCount, .ColumnCount, True)
            '.RowCount = 0


            Qry = Qry & "select model" & vbNewLine
            Qry = Qry & "from view_SHIPsummary" & vbNewLine
            Qry = Qry & "where ship_date between convert(varchar(8),CONVERT(DATETIME,'" & Format(DateTimeInput1.Value, "yyyy-MM-dd") & "'),112) and convert(varchar(8),CONVERT(DATETIME,'" & Format(DateTimeInput2.Value, "yyyy-MM-dd") & "'),112)" & vbNewLine
            If m_qry <> "" Then
                Qry = Qry & "AND MODEL IN (" & m_qry & ")" & vbNewLine
            End If
            Qry = Qry & "group by  model" & vbNewLine
            Qry = Qry & "order by model" & vbNewLine
            Qry = Qry & "" & vbNewLine

            rev_rs = Query_RS_ALL(Qry)
            '            rev_rs = Query_RS_ALL("exec SP_FRMREVENUE_GETMODEL '" & Site_id & "','" & ComboBoxEx1.Text & "','" & Format(DateTimeInput1.Value, "yyyy-MM-dd") & "','" & Format(DateTimeInput2.Value, "yyyy-MM-dd") & "'")

            If rev_rs Is Nothing Then
                Exit Sub
            End If

            If rev_rs.RecordCount > 0 Then
                .RowCount = rev_rs.RecordCount

                For i = 0 To .RowCount - 1
                    .SetValue(i, 0, rev_rs(0).Value)

                    .SetValue(i, 1, "출하금액")
                    rev_rs.MoveNext()
                Next

                .Cells(0, 0, .RowCount - 1, 1).HorizontalAlignment = FarPoint.Win.Spread.CellHorizontalAlignment.Center
                .Cells(0, 0, .RowCount - 1, 1).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center

            End If

            rev_rs = Nothing

            ' SUB TOTAL CALCULATE
            For i = 0 To .RowCount - 1
                fWeek = Month_Week(DateTimeInput1.Value)(0)
                For j = 2 To .ColumnCount - 1

                    Select Case ComboBoxEx2.Text
                        Case "YEAR"
                            QrySrt = "EXEC SP_FRMREVENUE_GETdetail_y '" & Site_id & "','" & .GetValue(i, 0) & "','" & DateTimeInput1.Text & "','" & .ColumnHeader.Cells(0, j).Text & "','GOOD'"
                        Case "MONTH"
                            QrySrt = "EXEC SP_FRMREVENUE_GETdetail_m '" & Site_id & "','" & .GetValue(i, 0) & "'," & fWeek & ",'" & Format(DateTimeInput1.Value, "yyyy-MM-dd") & "','" & Format(DateTimeInput2.Value, "yyyy-MM-dd") & "','GOOD'"
                        Case "WEEK"
                            QrySrt = "EXEC SP_FRMREVENUE_GETdetail_w '" & Site_id & "','" & .GetValue(i, 0) & "','" & Format(DateAdd(DateInterval.Day, ((j - 2) / 4), DateTimeInput1.Value), "yyyy-MM-dd") & "','GOOD'"
                    End Select

                    If .GetValue(i, 1) <> "Sub Total" Then

                        rev_rs = Query_RS_ALL(QrySrt)

                        If rev_rs Is Nothing Then
                        Else
                            .SetValue(i, j, rev_rs(0).Value)
                            .SetValue(i, j + 1, rev_rs(1).Value)
                            .SetText(i, j + 1, rev_rs(1).Value)

                            .SetValue(i, j + 3, rev_rs(2).Value)
                            .SetText(i, j + 3, rev_rs(2).Value)

                            '.SetValue(i, j + 4, rev_rs(3).Value)
                            '.SetText(i, j + 4, rev_rs(3).Value)


                        End If
                    Else
                        'For m = i - 6 To (i - 6 + 5)

                        '    .SetValue(i, j + 1, CInt(.GetValue(m, j + 1)) + CInt(.GetValue(i, j + 1)))
                        '    .SetValue(i, j + 3, CDec(.GetValue(m, j + 3)) + CDec(.GetValue(i, j + 3)))
                        'Next

                    End If
                    fWeek += 1
                    j = j + 3
                Next

            Next


            'GRAND TOTAL(ROW)
            .RowCount = .RowCount + 1
            .Cells(.RowCount - 1, 0).Text = "TOTAL"
            '            .Cells(.RowCount - 1, 2, .RowCount - 1, .ColumnCount - 1).Value = 0
            .Rows(.RowCount - 1).BackColor = Color.Pink
            '.Columns(.ColumnCount - 4).CellType = intcell
            '.Columns(.ColumnCount - 3).CellType = deccell
            '.Columns(.ColumnCount - 2).CellType = percell
            '.Columns(.ColumnCount - 1).CellType = deccell

            If SPREAD_ROW_TOTAL_DEC(FpSpread1, 2, .ColumnCount - 1, 1) = True Then

            End If


            ' RATIO CALCULATE
            'For i = 0 To .RowCount - 1
            '    For j = 4 To .ColumnCount - 1
            '        If .GetValue(i, 1) = "Sub Total" Or .GetValue(i, 1) = "TOTAL" Then

            '            If CInt(.GetValue(i, j - 1)) > 0 Then

            '                'GRAND TOTAL SUM(ROW)

            '                If j < .ColumnCount - 6 And i < .RowCount - 7 Then

            '                    .SetValue(.RowCount - 7, j - 1, CInt(.GetValue(i - 6, j - 1)) + CInt(.GetValue(.RowCount - 7, j - 1)))
            '                    .SetValue(.RowCount - 7, j + 1, CInt(.GetValue(i - 6, j + 1)) + CInt(.GetValue(.RowCount - 7, j + 1)))

            '                    .SetValue(.RowCount - 6, j - 1, CInt(.GetValue(i - 5, j - 1)) + CInt(.GetValue(.RowCount - 6, j - 1)))
            '                    .SetValue(.RowCount - 6, j + 1, CInt(.GetValue(i - 5, j + 1)) + CInt(.GetValue(.RowCount - 6, j + 1)))

            '                    .SetValue(.RowCount - 5, j - 1, CInt(.GetValue(i - 4, j - 1)) + CInt(.GetValue(.RowCount - 5, j - 1)))
            '                    .SetValue(.RowCount - 5, j + 1, CInt(.GetValue(i - 4, j + 1)) + CInt(.GetValue(.RowCount - 5, j + 1)))

            '                    .SetValue(.RowCount - 4, j - 1, CInt(.GetValue(i - 3, j - 1)) + CInt(.GetValue(.RowCount - 4, j - 1)))
            '                    .SetValue(.RowCount - 4, j + 1, CInt(.GetValue(i - 3, j + 1)) + CInt(.GetValue(.RowCount - 4, j + 1)))

            '                    .SetValue(.RowCount - 3, j - 1, CInt(.GetValue(i - 2, j - 1)) + CInt(.GetValue(.RowCount - 3, j - 1)))
            '                    .SetValue(.RowCount - 3, j + 1, CInt(.GetValue(i - 2, j + 1)) + CInt(.GetValue(.RowCount - 3, j + 1)))

            '                    .SetValue(.RowCount - 2, j - 1, CInt(.GetValue(i - 1, j - 1)) + CInt(.GetValue(.RowCount - 2, j - 1)))
            '                    .SetValue(.RowCount - 2, j + 1, CDec(.GetValue(i - 1, j + 1)) + CDec(.GetValue(.RowCount - 2, j + 1)))

            '                    .SetValue(.RowCount - 1, j - 1, CInt(.GetValue(i, j - 1)) + CInt(.GetValue(.RowCount - 1, j - 1)))
            '                    .SetValue(.RowCount - 1, j + 1, CDec(.GetValue(i, j + 1)) + CDec(.GetValue(.RowCount - 1, j + 1)))

            '                End If

            '                'GRAND TOTAL SUM(COLUMN)
            '                If j < .ColumnCount - 6 Then
            '                    .SetValue(i, .ColumnCount - 3, CInt(.GetValue(i, j - 1)) + CInt(.GetValue(i, .ColumnCount - 3)))
            '                    .SetValue(i, .ColumnCount - 1, CInt(.GetValue(i, j + 1)) + CInt(.GetValue(i, .ColumnCount - 1)))

            '                    .SetValue(i - 6, .ColumnCount - 3, CInt(.GetValue(i - 6, j - 1)) + CInt(.GetValue(i - 6, .ColumnCount - 3)))
            '                    .SetValue(i - 6, .ColumnCount - 1, CInt(.GetValue(i - 6, j + 1)) + CInt(.GetValue(i - 6, .ColumnCount - 1)))

            '                    .SetValue(i - 5, .ColumnCount - 3, CInt(.GetValue(i - 5, j - 1)) + CInt(.GetValue(i - 5, .ColumnCount - 3)))
            '                    .SetValue(i - 5, .ColumnCount - 1, CInt(.GetValue(i - 5, j + 1)) + CInt(.GetValue(i - 5, .ColumnCount - 1)))



            '                    .SetValue(i - 4, .ColumnCount - 3, CInt(.GetValue(i - 4, j - 1)) + CInt(.GetValue(i - 4, .ColumnCount - 3)))
            '                    .SetValue(i - 4, .ColumnCount - 1, CInt(.GetValue(i - 4, j + 1)) + CInt(.GetValue(i - 4, .ColumnCount - 1)))

            '                    .SetValue(i - 3, .ColumnCount - 3, CInt(.GetValue(i - 3, j - 1)) + CInt(.GetValue(i - 3, .ColumnCount - 3)))
            '                    .SetValue(i - 3, .ColumnCount - 1, CInt(.GetValue(i - 3, j + 1)) + CInt(.GetValue(i - 3, .ColumnCount - 1)))

            '                    .SetValue(i - 2, .ColumnCount - 3, CInt(.GetValue(i - 2, j - 1)) + CInt(.GetValue(i - 2, .ColumnCount - 3)))
            '                    .SetValue(i - 2, .ColumnCount - 1, CInt(.GetValue(i - 2, j + 1)) + CInt(.GetValue(i - 2, .ColumnCount - 1)))

            '                    .SetValue(i - 1, .ColumnCount - 3, CInt(.GetValue(i - 1, j - 1)) + CInt(.GetValue(i - 1, .ColumnCount - 3)))
            '                    .SetValue(i - 1, .ColumnCount - 1, CInt(.GetValue(i - 1, j + 1)) + CInt(.GetValue(i - 1, .ColumnCount - 1)))
            '                End If

            '                .SetValue(i - 6, j, CInt(.GetValue(i - 6, j - 1)) / CInt(.GetValue(i, j - 1)))
            '                .SetValue(i - 5, j, CInt(.GetValue(i - 5, j - 1)) / CInt(.GetValue(i, j - 1)))
            '                .SetValue(i - 4, j, CInt(.GetValue(i - 4, j - 1)) / CInt(.GetValue(i, j - 1)))
            '                .SetValue(i - 3, j, CInt(.GetValue(i - 3, j - 1)) / CInt(.GetValue(i, j - 1)))
            '                .SetValue(i - 2, j, CInt(.GetValue(i - 2, j - 1)) / CInt(.GetValue(i, j - 1)))
            '                .SetValue(i - 1, j, CInt(.GetValue(i - 1, j - 1)) / CInt(.GetValue(i, j - 1)))

            '            End If
            '        End If
            '        j = j + 3
            '    Next

            'Next

            intcell_r.DecimalPlaces = 0
            intcell_r.ShowSeparator = True
            intcell_r.Separator = ","

            deccell_r.DecimalPlaces = 2
            deccell_r.ShowSeparator = True
            deccell_r.Separator = ","

            For i = 2 To .ColumnCount - 4
                .Columns(i).CellType = intcell_r

                .Cells(0, i + 1, .RowCount - 1, i + 1).CellType = deccell_r
                If .Cells(.RowCount - 1, i + 1).Text = "" Then
                    .Cells(.RowCount - 1, i + 1).Value = 0
                End If

                .Columns(i + 2).CellType = percell
                '                    .Columns(i + 3).CellType = deccell_r
                .Columns(i + 3).CellType = curcell
                If .Cells(.RowCount - 1, i + 3).Text = "" Then
                    .Cells(.RowCount - 1, i + 3).Value = 0
                End If

                i = i + 3
            Next

            disp_chart(.RowCount - 1)
        End With
        Spread_AutoCol(FpSpread1)

        If ComboBoxEx2.Text = "YEAR" Then
            FpSpread1.ActiveSheet.Columns(2, FpSpread1.ActiveSheet.ColumnCount - 1).Width = 80
        ElseIf ComboBoxEx2.Text = "YEAR" Then
            FpSpread1.ActiveSheet.Columns(2, FpSpread1.ActiveSheet.ColumnCount - 1).Width = 120
        Else
            FpSpread1.ActiveSheet.Columns(2, FpSpread1.ActiveSheet.ColumnCount - 1).Width = 100
        End If

        MessageBox.Show("조회되었습니다!")

        'Catch ex As Exception
        '    MessageBox.Show("Error: " & ex.Message, "ERROR")
        'End Try
    End Sub

    Private Sub disp_chart(ByVal rowidx As Integer)

        Me.Chart1.Visible = True

        Dim cc
        Dim bb
        Dim i As Integer

        cc = ""
        bb = ""

        With FpSpread1.ActiveSheet
            For i = 2 To .ColumnCount - 1
                cc = cc & .ColumnHeader.Cells(0, i).Text & ","
                i = i + 3
            Next

            For i = 3 To .ColumnCount - 1
                bb = bb & "," & i
                i = i + 3
            Next

        End With

        Select ComboBoxEx2.Text
            Case "YEAR"
                'If rowidx < 0 Then
                '            Chart_Spread_All(Me.Chart1, True, 2, DataVisualization.Charting.SeriesChartType.Column, Me.FpSpread1, New String() {"JAN", "FEB", "MAR", "APR", "MAY", "JUN", "JUL", "AUG", "SEP", "OCT", "NOV", "DEC"}, 0, New Integer() {3, 7, 11, 15, 19, 23, 27, 31, 35, 39, 43, 47}, True)
                'Chart_Spread_All_1(Me.Chart1, False, 2, DataVisualization.Charting.SeriesChartType.Column, Me.FpSpread1, New String() {"JAN", "FEB", "MAR", "APR", "MAY", "JUN", "JUL", "AUG", "SEP", "OCT", "NOV", "DEC"}, 41, New Integer() {3, 7, 11, 15, 19, 23, 27, 31, 35, 39, 43, 47}, False)

                'Else
                '    '                    Chart_Spread_Col(Me.Chart1, False, 2, DataVisualization.Charting.SeriesChartType.Column, Me.FpSpread1, Me.FpSpread1.ActiveSheet.GetValue(rowidx, 0), rowidx, New Integer() {3, 7, 11, 15, 19, 23, 27, 31, 35, 39, 43, 47}, 0)
                Chart_Spread_Col(Me.Chart1, False, 2, DataVisualization.Charting.SeriesChartType.Column, Me.FpSpread1, "출하금액", rowidx, New Integer() {3, 7, 11, 15, 19, 23, 27, 31, 35, 39, 43, 47}, 0)
                'End If
            Case "MONTH"

                Dim stWeek, edWeek, totWeek As Integer
                Dim Xlbl As String() = {}
                Dim Ylbl As Integer() = {}

                stWeek = Month_Week(DateTimeInput1.Value)(0)
                edWeek = Month_Week(DateTimeInput1.Value)(1)
                totWeek = edWeek - stWeek + 1

                Select Case totWeek
                    Case 4
                        Xlbl = New String() {"1st", "2st", "3st", "4st"}
                        Ylbl = New Integer() {3, 7, 11, 15}
                    Case 5
                        Xlbl = New String() {"1st", "2st", "3st", "4st", "5st"}
                        Ylbl = New Integer() {3, 7, 11, 15, 19}
                    Case 6
                        Xlbl = New String() {"1st", "2st", "3st", "4st", "5st", "6st"}
                        Ylbl = New Integer() {3, 7, 11, 15, 19, 23, 27}
                End Select

                If rowidx < 0 Then
                    Chart_Spread_All(Me.Chart1, False, 2, DataVisualization.Charting.SeriesChartType.Column, Me.FpSpread1, Xlbl, 0, Ylbl, True)
                Else
                    Chart_Spread_Col(Me.Chart1, False, 2, DataVisualization.Charting.SeriesChartType.Column, Me.FpSpread1, "출하금액", rowidx, Ylbl, 0)
                End If

            Case "WEEK"

                Dim Xlbl As String() = {}
                Dim Ylbl As Integer() = {}

                Xlbl = New String() {DateAdd(DateInterval.Day, 0, DateTimeInput1.Value), DateAdd(DateInterval.Day, 1, DateTimeInput1.Value), DateAdd(DateInterval.Day, 3, DateTimeInput1.Value), DateAdd(DateInterval.Day, 4, DateTimeInput1.Value), DateAdd(DateInterval.Day, 5, DateTimeInput1.Value), DateAdd(DateInterval.Day, 6, DateTimeInput1.Value)}
                Ylbl = New Integer() {3, 7, 11, 15, 19, 23, 27}


                If rowidx < 0 Then
                    Chart_Spread_All(Me.Chart1, False, 2, DataVisualization.Charting.SeriesChartType.Column, Me.FpSpread1, Xlbl, 0, Ylbl, True)
                Else
                    Chart_Spread_Col(Me.Chart1, False, 2, DataVisualization.Charting.SeriesChartType.Column, Me.FpSpread1, "출하금액", rowidx, Ylbl, 0)
                End If

        End Select
    End Sub

    Private Sub Create_HD()
        Try
            Dim I As Integer
            DateTimeInput1.Format = DevComponents.Editors.eDateTimePickerFormat.Custom
            DateTimeInput2.Visible = False
            LabelX2.Visible = False

            curcell.ShowCurrencySymbol = True
            curcell.CurrencySymbol = "$"
            curcell.ShowSeparator = True

            FpSpread1.ActiveSheet.ClearRange(0, 0, FpSpread1.ActiveSheet.RowCount - 1, FpSpread1.ActiveSheet.ColumnCount - 1, True)

            If ComboBoxEx2.Text = "YEAR" Then
                DateTimeInput1.CustomFormat = "yyyy"
                DateTimeInput2.CustomFormat = "yyyy"

                With FpSpread1.ActiveSheet
                    'If .RowCount > 0 Then
                    '    .Rows(0, .RowCount - 1).CellType = textcell
                    'End If
                    .ClearRange(0, 0, .RowCount - 1, .ColumnCount - 1, True)

                    '.ColumnCount = 0
                    '.RowCount = 0

                    .ColumnCount = 50
                    .FrozenColumnCount = 2


                    .ColumnHeaderRowCount = 2
                    .ColumnHeader.Cells(0, 0).RowSpan = 2
                    .ColumnHeader.Cells(0, 0).Text = "모델"
                    .ColumnHeader.Cells(0, 1).RowSpan = 2
                    .ColumnHeader.Cells(0, 1).Text = "매출유형"

                    .ColumnHeader.Cells(0, 2).ColumnSpan = 4
                    .ColumnHeader.Cells(0, 2).Text = "01"
                    .ColumnHeader.Cells(0, 3).Text = "01"
                    .ColumnHeader.Cells(0, 6).ColumnSpan = 4
                    .ColumnHeader.Cells(0, 6).Text = "02"
                    .ColumnHeader.Cells(0, 7).Text = "02"
                    .ColumnHeader.Cells(0, 10).ColumnSpan = 4
                    .ColumnHeader.Cells(0, 10).Text = "03"
                    .ColumnHeader.Cells(0, 11).Text = "03"
                    .ColumnHeader.Cells(0, 14).ColumnSpan = 4
                    .ColumnHeader.Cells(0, 14).Text = "04"
                    .ColumnHeader.Cells(0, 15).Text = "04"
                    .ColumnHeader.Cells(0, 18).ColumnSpan = 4
                    .ColumnHeader.Cells(0, 18).Text = "05"
                    .ColumnHeader.Cells(0, 19).Text = "05"
                    .ColumnHeader.Cells(0, 22).ColumnSpan = 4
                    .ColumnHeader.Cells(0, 22).Text = "06"
                    .ColumnHeader.Cells(0, 23).Text = "06"
                    .ColumnHeader.Cells(0, 26).ColumnSpan = 4
                    .ColumnHeader.Cells(0, 26).Text = "07"
                    .ColumnHeader.Cells(0, 27).Text = "07"
                    .ColumnHeader.Cells(0, 30).ColumnSpan = 4
                    .ColumnHeader.Cells(0, 30).Text = "08"
                    .ColumnHeader.Cells(0, 31).Text = "08"
                    .ColumnHeader.Cells(0, 34).ColumnSpan = 4
                    .ColumnHeader.Cells(0, 34).Text = "09"
                    .ColumnHeader.Cells(0, 35).Text = "09"
                    .ColumnHeader.Cells(0, 38).ColumnSpan = 4
                    .ColumnHeader.Cells(0, 38).Text = "10"
                    .ColumnHeader.Cells(0, 39).Text = "10"
                    .ColumnHeader.Cells(0, 42).ColumnSpan = 4
                    .ColumnHeader.Cells(0, 42).Text = "11"
                    .ColumnHeader.Cells(0, 43).Text = "11"
                    .ColumnHeader.Cells(0, 46).ColumnSpan = 4
                    .ColumnHeader.Cells(0, 46).Text = "12"
                    .ColumnHeader.Cells(0, 47).Text = "12"

                    intcell.DecimalPlaces = 0
                    deccell.DecimalPlaces = 2

                    For I = 2 To .ColumnCount - 1
                        .ColumnHeader.Cells(1, I).Text = "출하수량"
                        .Columns(I).CellType = intcell
                        .ColumnHeader.Cells(1, I + 1).Text = "출하금액"
                        .Columns(I + 1).CellType = deccell

                        .ColumnHeader.Cells(1, I + 2).Text = "비율"
                        .Columns(I + 2).CellType = percell
                        .Columns(I + 2).Visible = False
                        .ColumnHeader.Cells(1, I + 3).Text = "매출금액"
                        .Columns(I + 3).CellType = curcell
                        .Columns(I + 3).Visible = False

                        I = I + 3
                    Next


                End With

            ElseIf ComboBoxEx2.Text = "MONTH" Then
                DateTimeInput1.CustomFormat = "MM/yyyy"
                DateTimeInput2.CustomFormat = "MM/yyyy"

                DateTimeInput1.Value = CDate(Month_Day(DateTimeInput1.Value)(0))
                DateTimeInput2.Value = CDate(Month_Day(DateTimeInput1.Value)(1))

                With FpSpread1.ActiveSheet
                    Dim stWeek, edWeek, totWeek As Integer

                    stWeek = Month_Week(DateTimeInput1.Value)(0)
                    edWeek = Month_Week(DateTimeInput1.Value)(1)
                    totWeek = edWeek - stWeek + 1

                    .ClearRange(0, 0, .RowCount - 1, .ColumnCount - 1, True)

                    .ColumnCount = 2 + totWeek * 4
                    .FrozenColumnCount = 2

                    .ColumnHeaderRowCount = 2
                    .ColumnHeader.Cells(0, 0).RowSpan = 2
                    .ColumnHeader.Cells(0, 0).Text = "모델"
                    .ColumnHeader.Cells(0, 1).RowSpan = 2
                    .ColumnHeader.Cells(0, 1).Text = "매출유형"

                    For I = 0 To (totWeek - 1) * 4 Step 4
                        .ColumnHeader.Cells(0, I + 2).ColumnSpan = 4
                        .ColumnHeader.Cells(0, I + 2).Text = (I / 4 + 1) & "st(" & Week_Day(DateTimeInput1.Value, I / 4 + 1, totWeek)(0) & "~" & Week_Day(DateTimeInput1.Value, I / 4 + 1, totWeek)(1) & ")"
                        .ColumnHeader.Cells(0, I + 3).Text = (I / 4 + 1) & "st(" & Week_Day(DateTimeInput1.Value, I / 4 + 1, totWeek)(0) & "~" & Week_Day(DateTimeInput1.Value, I / 4 + 1, totWeek)(1) & ")"
                    Next

                    For I = 2 To .ColumnCount - 1
                        .ColumnHeader.Cells(1, I).Text = "출하수량"
                        .Columns(I).CellType = intcell
                        .ColumnHeader.Cells(1, I + 1).Text = "출하금액"
                        .Columns(I + 1).CellType = deccell

                        .ColumnHeader.Cells(1, I + 2).Text = "RATIO"
                        .Columns(I + 2).CellType = percell
                        .Columns(I + 2).Visible = False
                        .ColumnHeader.Cells(1, I + 3).Text = "AMT"
                        .Columns(I + 3).CellType = curcell
                        .Columns(I + 3).Visible = False

                        I = I + 3
                    Next


                End With


            ElseIf ComboBoxEx2.Text = "WEEK" Then
                DateTimeInput1.CustomFormat = "dd/MM/yyyy"
                DateTimeInput2.CustomFormat = "dd/MM/yyyy"

                DateTimeInput1.Value = CDate(Week_FEday(DateTimeInput1.Value)(0))
                DateTimeInput2.Value = CDate(Week_FEday(DateTimeInput1.Value)(1))

                With FpSpread1.ActiveSheet

                    .ClearRange(0, 0, .RowCount - 1, .ColumnCount - 1, True)

                    .ColumnCount = 30
                    .FrozenColumnCount = 2


                    .ColumnHeaderRowCount = 2
                    .ColumnHeader.Cells(0, 0).RowSpan = 2
                    .ColumnHeader.Cells(0, 0).Text = "모델"
                    .ColumnHeader.Cells(0, 1).RowSpan = 2
                    .ColumnHeader.Cells(0, 1).Text = "매출유형"

                    For I = 0 To 24 Step 4
                        .ColumnHeader.Cells(0, I + 2).ColumnSpan = 4
                        .ColumnHeader.Cells(0, I + 2).Text = Format(DateAdd(DateInterval.Day, I / 4, DateTimeInput1.Value), "d")
                        .ColumnHeader.Cells(0, I + 3).Text = Format(DateAdd(DateInterval.Day, I / 4, DateTimeInput1.Value), "d")
                    Next

                    For I = 2 To .ColumnCount - 1
                        .ColumnHeader.Cells(1, I).Text = "출하수량"
                        .Columns(I).CellType = intcell
                        .ColumnHeader.Cells(1, I + 1).Text = "출하금액"
                        .Columns(I + 1).CellType = deccell

                        .ColumnHeader.Cells(1, I + 2).Text = "RATIO"
                        .Columns(I + 2).CellType = percell
                        .Columns(I + 2).Visible = False

                        .ColumnHeader.Cells(1, I + 3).Text = "AMT"
                        .Columns(I + 3).CellType = curcell
                        .Columns(I + 3).Visible = False

                        I = I + 3
                    Next


                End With

            End If
        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message, "ERROR")
        End Try
    End Sub

    Private Sub DateTimeInput1_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DateTimeInput1.TextChanged
        Create_HD()
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
            If Spread_Print(Me.FpSpread1, "Revenue Summary", 0) = False Then
                MsgBox("Fail to Print")
            End If
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
        Else
            MessageBox.Show("Select Spread for Saving!!")
        End If

    End Sub


    Private Sub FpSpread1_CellClick(ByVal sender As System.Object, ByVal e As FarPoint.Win.Spread.CellClickEventArgs) Handles FpSpread1.CellClick
        save_excel = "FpSpread1"
    End Sub

End Class