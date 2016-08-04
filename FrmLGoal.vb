Public Class FrmLGoal



    Private Sub FrmLGoal_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        DockContainerItem1.Text = "실행 메뉴"
        DockContainerItem2.Text = "조회 조건"

        DateTimeInput1.Value = Now

        If Query_Combo(Me.ComboBoxEx1, "SELECT model_no FROM tbl_modelmaster WHERE site_id = '" & Site_id & "' and active = 'Y' ORDER BY model_no") = True Then
        End If
        Me.ComboBoxEx1.Items.Add("ALL")
        Me.ComboBoxEx1.Text = "ALL"

        If Spread_Setting_ByCode(FpSpread1, "SP_COMMON_GETCODEMASTER", "R2000", "INT") = True Then
            With FpSpread1.ActiveSheet
                .RowCount = 0
                .AddColumns(0, 1)
                .ColumnHeader.Columns(0).Label = "LINE"
                .AddColumns(0, 1)
                .ColumnHeader.Columns(0).Label = "MODEL"
                .ColumnCount = .ColumnCount + 1
                .ColumnHeader.Columns(.ColumnCount - 1).Label = "TOTAL"
                .Columns(.ColumnCount - 1).CellType = intcell
                .Columns(.ColumnCount - 1).BackColor = Color.Yellow
                Spread_AutoCol(FpSpread1)
                .FrozenColumnCount = 1
            End With
        End If

        'If Spread_Setting_ByCode(FpSpread2, "SP_COMMON_GETCODEMASTER", "R2000", "INT") = True Then
        '    With FpSpread2.ActiveSheet
        '        .RowCount = 0
        '        .AddColumns(0, 1)
        '        .ColumnHeader.Columns(0).Label = "GROUP"
        '        .AddColumns(0, 1)
        '        .ColumnHeader.Columns(0).Label = "MODEL"
        '        .ColumnCount = .ColumnCount + 1
        '        .ColumnHeader.Columns(.ColumnCount - 1).Label = "TOTAL"
        '        .Columns(.ColumnCount - 1).CellType = intcell
        '        .Columns(.ColumnCount - 1).BackColor = Color.Yellow
        '        Spread_AutoCol(FpSpread2)
        '        .FrozenColumnCount = 1
        '        .Protect = False
        '    End With
        'End If

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


    Function Refresh_Result() As Boolean
        Dim I, J As Integer

        If Query_Spread(FpSpread1, "EXEC SP_FRMLGOAL_GETPLAN '" & Site_id & "','" & ComboBoxEx1.Text & "','" & DateTimeInput1.Text & "'", 1) = True Then

            With FpSpread1.ActiveSheet

                For I = 0 To .RowCount - 1
                    For J = 2 To .ColumnCount - 1
                        If .GetValue(I, J) = 0 Then
                            .SetText(I, J, "")
                        End If
                    Next
                Next


                .RowCount = .RowCount + 4
                .Cells(.RowCount - 4, 2, .RowCount - 1, .ColumnCount - 1).CellType = intcell
                .Rows(.RowCount - 4, .RowCount - 1).BackColor = Color.Yellow
                .SetText(.RowCount - 4, 0, "TOTAL")
                .SetText(.RowCount - 3, 0, "TOTAL")
                .SetText(.RowCount - 2, 0, "TOTAL")
                .SetText(.RowCount - 1, 0, "TOTAL")
                .SetText(.RowCount - 4, 1, "수입검사")
                .SetText(.RowCount - 3, 1, "합지")
                .SetText(.RowCount - 2, 1, "조립")
                .SetText(.RowCount - 1, 1, "출하검사")

                For I = 0 To .RowCount - 1
                    .Cells(I, 0).RowSpan = 4

                    If I < .RowCount - 4 Then
                        For J = 2 To .ColumnCount - 1
                            .SetValue(.RowCount - 4, J, .GetValue(I, J) + .GetValue(.RowCount - 4, J))
                            .SetValue(.RowCount - 3, J, .GetValue(I, J) + .GetValue(.RowCount - 3, J))
                            .SetValue(.RowCount - 2, J, .GetValue(I + 1, J) + .GetValue(.RowCount - 2, J))
                            .SetValue(.RowCount - 1, J, .GetValue(I + 2, J) + .GetValue(.RowCount - 1, J))
                        Next
                    End If

                    I = I + 3
                Next


                '                SPREAD_ROW_TOTAL(FpSpread1, 2, .ColumnCount - 1, 1)

                Spread_AutoCol(FpSpread1)
                '                .Columns(1, .ColumnCount - 1).Width = 60
                .Protect = True
            End With
        End If

    End Function

    Private Sub FindBtn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FindBtn.Click
        FpSpread1.ActiveSheet.ClearRange(0, 0, FpSpread1.ActiveSheet.RowCount - 1, FpSpread1.ActiveSheet.ColumnCount - 1, False)

        FpSpread1.ActiveSheet.RowCount = 0
        '        FpSpread2.ActiveSheet.RowCount = 0

        Refresh_Result()
    End Sub


    Private Sub ButtonItem2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonItem2.Click, NewBtn1.Click

        'With FpSpread2.ActiveSheet

        '    FpSpread2.ActiveSheet.Protect = False
        '    .AddRows(0, 1)
        '    .Rows(0).ForeColor = Color.OrangeRed
        '    .Cells(0, 1).CellType = combocell
        '    .Cells(0, 2, 0, .ColumnCount - 1).CellType = intcell
        '    .Columns(0).Width = 100
        '    .SetValue(0, 0, FpSpread1.ActiveSheet.GetValue(FpSpread1.ActiveSheet.ActiveRowIndex, 0))
        '    .SetValue(0, 2, 0)
        '    .SetValue(0, 3, 0)
        '    .SetValue(0, 4, 0)
        '    .SetValue(0, 5, 0)
        '    .SetValue(0, 6, 0)
        '    .SetValue(0, 7, 0)


        '    If UCase(FpSpread1.ActiveSheet.GetValue(FpSpread1.ActiveSheet.ActiveRowIndex, 1)) = "TRIAGE" Then
        '        Chg_ComboCell(FpSpread2, 0, 1, Query_Cell_Code1("code_name", "R0053"))
        '    ElseIf UCase(FpSpread1.ActiveSheet.GetValue(FpSpread1.ActiveSheet.ActiveRowIndex, 1)) = "DATA ENTRY" Then
        '        Chg_ComboCell(FpSpread2, 0, 1, Query_Cell_Code1("code_name", "R0051"))
        '    ElseIf UCase(FpSpread1.ActiveSheet.GetValue(FpSpread1.ActiveSheet.ActiveRowIndex, 1)) = "QC" Then
        '        Chg_ComboCell(FpSpread2, 0, 1, Query_Cell_Code1("code_name", "R0052"))

        '    End If

        'End With

    End Sub

    Private Sub ButtonItem3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonItem3.Click, SaveBtn1.Click
        Dim i, j, qty As Integer
        Dim Gtime, model, wc As String

        With FpSpread1.ActiveSheet
            .SetActiveCell(.ActiveRowIndex, 0)

            .Protect = False

            If .RowCount > 0 Then
                For i = 0 To .RowCount - 1
                    model = .GetValue(i, 0)
                    wc = Query_RS("select code_id from tbl_codemaster where site_id ='" & Site_id & "' and class_id = 'R0001' and code_name = '" & .GetValue(i, 1) & "'")

                    For j = 2 To .ColumnCount - 2
                        If .Cells(i, j).ForeColor = Color.OrangeRed Then
                            If .GetValue(i, j) = 0 Then
                                qty = 0
                            Else
                                qty = .GetValue(i, j)
                            End If
                            Gtime = Query_RS("select code_id from tbl_codemaster where site_id ='" & Site_id & "' and class_id = 'R2000' and code_name = '" & .ColumnHeader.Columns(j).Label & "'")
                            If Insert_Data("exec SP_FRMLGOAL_SETPLAN '" & Site_id & "','" & DateTimeInput1.Text & "','" & Gtime & "','" & model & "','" & wc & "',''," & qty & ",'" & Emp_No & "'") = True Then
                                .Rows(i).ForeColor = Color.Black
                            End If
                        End If
                    Next

                Next

            Else
                MessageBox.Show("No Rows")
            End If
        End With

        FindBtn_Click(sender, e)

    End Sub

    Private Sub ButtonItem4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonItem4.Click, DelBtn1.Click

    End Sub

    Private Sub ButtonItem5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonItem5.Click, PrtBtn1.Click
        If save_excel = "" Then
            MessageBox.Show("Select Spread for Printing!!")
            Exit Sub
        End If

        If save_excel = "FpSpread1" Then
            If Spread_Print(Me.FpSpread1, "WIP Inventory(Board) Summary", 0) = False Then
                MsgBox("Fail to Print")
            End If
        ElseIf save_excel = "FpSpread2" Then
            'If Spread_Print(Me.FpSpread2, "WIP Inventory(Board) Details", 0) = False Then
            '    MsgBox("Fail to Print")
            'End If
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
        ElseIf save_excel = "FpSpread2" Then
            'File_Save(SaveFileDialog1, FpSpread2)
        Else
            MessageBox.Show("Select Spread for Saving!!")
        End If

    End Sub

    Private Sub FpSpread1_CellClick(ByVal sender As System.Object, ByVal e As FarPoint.Win.Spread.CellClickEventArgs) Handles FpSpread1.CellClick
        'Dim MODEL, WC As String


        save_excel = "FpSpread1"

        'With FpSpread1.ActiveSheet
        '    model = .GetValue(.ActiveRowIndex, 0)
        '    wc = Query_RS("select code_id from tbl_codemaster where site_id ='" & Site_id & "' and class_id = 'R0001' and code_name = '" & .GetValue(.ActiveRowIndex, 1) & "'")
        'End With

        'With FpSpread2.ActiveSheet
        '    .RowCount = 0

        '    If Query_Spread(FpSpread2, "EXEC SP_FRMLGOAL_GETPLANDETAILS '" & Site_id & "','" & DateTimeInput1.Text & "','" & MODEL & "','" & WC & "'", 1) = True Then
        '        Spread_AutoCol(FpSpread2)
        '    End If
        'End With

    End Sub

    Private Sub FpSpread2_CellClick(ByVal sender As System.Object, ByVal e As FarPoint.Win.Spread.CellClickEventArgs)
        save_excel = "FpSpread2"


    End Sub

    Private Sub FpSpread1_CellDoubleClick(ByVal sender As Object, ByVal e As FarPoint.Win.Spread.CellClickEventArgs) Handles FpSpread1.CellDoubleClick

        With FpSpread1.ActiveSheet

            If e.Column = 0 Then
                .Cells(e.Row, e.Column).Locked = True
                Exit Sub
            ElseIf e.Column = 1 Then
                .Cells(e.Row, e.Column).Locked = True
                Exit Sub
            ElseIf e.Column = .ColumnCount - 1 Then
                .Cells(e.Row, e.Column).Locked = True
                Exit Sub
            Else
                If .Rows(e.Row).BackColor = Color.Yellow Then
                    .Cells(e.Row, e.Column).Locked = True
                    Exit Sub
                Else
                    .Cells(e.Row, e.Column).ForeColor = Color.OrangeRed
                    .Cells(e.Row, e.Column).Locked = False
                End If
            End If

        End With

    End Sub

    'Private Sub FpSpread2_Change(ByVal sender As Object, ByVal e As FarPoint.Win.Spread.ChangeEventArgs)
    '    With FpSpread2.ActiveSheet
    '        '.Cells(e.Row, .ColumnCount - 1).Formula = True
    '        If .GetValue(e.Row, 1) = "" Then
    '            MessageBox.Show("Select Group!!!")
    '            Exit Sub
    '        End If

    '        FpSpread2.AllowUserFormulas = True

    '        .Cells(e.Row, .ColumnCount - 1).Formula = "SUM(C" & e.Row + 1 & ":H" & e.Row + 1 & ")"
    '        .Rows(e.Row).ForeColor = Color.OrangeRed

    '    End With
    'End Sub

    Private Sub FpSpread1_SelectionChanged(ByVal sender As Object, ByVal e As FarPoint.Win.Spread.SelectionChangedEventArgs) Handles FpSpread1.SelectionChanged
        Dim MODEL, WC As String


        save_excel = "FpSpread1"

        With FpSpread1.ActiveSheet
            MODEL = .GetValue(.ActiveRowIndex, 0)
            WC = Query_RS("select code_id from tbl_codemaster where site_id ='" & Site_id & "' and class_id = 'R0001' and code_name = '" & .GetValue(.ActiveRowIndex, 1) & "'")
        End With

        'With FpSpread2.ActiveSheet
        '    .RowCount = 0

        '    If Query_Spread(FpSpread2, "EXEC SP_FRMLGOAL_GETPLANDETAILS '" & Site_id & "','" & DateTimeInput1.Text & "','" & MODEL & "','" & WC & "'", 1) = True Then
        '        Spread_AutoCol(FpSpread2)
        '    End If
        'End With
    End Sub
End Class