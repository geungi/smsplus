Public Class FrmWKRESULT1

    Private Sub FrmWKRESULT1_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        'Dim i As Integer

        DockContainerItem1.Text = "Form Menu"
        DockContainerItem2.Text = "Retrieve Condition"

        DateTimeInput1.Value = Now
        DateTimeInput2.Value = Now

        If Query_Combo(Me.ComboBoxEx1, "SELECT model_no FROM tbl_modelmaster WHERE site_id = '" & Site_id & "' and active = 'Y' ORDER BY model_no") = True Then
        End If
        Me.ComboBoxEx1.Items.Add("ALL")
        Me.ComboBoxEx1.Text = "ALL"

        If Spread_Setting_ByCode(FpSpread1, "SP_COMMON_GETCODEMASTER", "R0001", "INT") = True Then
            With FpSpread1.ActiveSheet
                .RowCount = 0


                .AddColumns(0, 1)
                .ColumnHeader.Rows(0).Height = 60
                .ColumnHeader.Columns(0).Label = "WORK DATE"

                .AddColumns(0, 1)
                .ColumnHeader.Rows(0).Height = 40
                .ColumnHeader.Columns(0).Label = "MODEL"


                .ColumnCount = FpSpread1.ActiveSheet.ColumnCount + 1
                .ColumnHeader.Columns(FpSpread1.ActiveSheet.ColumnCount - 1).Label = "TOTAL"

                'For i = 0 To .ColumnCount - 1
                '    If .ColumnHeader.Columns(i).Label = "PD/SCRAP" Then
                '        .ColumnHeader.Columns(i).Label = "PD" & vbNewLine & "SCRAP"
                '    ElseIf .ColumnHeader.Columns(i).Label = "LEVEL 4" Then
                '        .ColumnHeader.Columns(i).Label = "LEVEL 4"
                '    ElseIf .ColumnHeader.Columns(i).Label = "LEVEL 3" Then
                '        .ColumnHeader.Columns(i).Label = "LEVEL 3"
                '    ElseIf .ColumnHeader.Columns(i).Label = "DATA ENTRY" Then
                '        .ColumnHeader.Columns(i).Label = "DATA" & vbNewLine & "ENTRY"
                '    ElseIf .ColumnHeader.Columns(i).Label = "DM&INFO" Then
                '        .ColumnHeader.Columns(i).Label = "DM&" & vbNewLine & "INFO"
                '    ElseIf .ColumnHeader.Columns(i).Label = "DOWNLOAD" Then
                '        .ColumnHeader.Columns(i).Label = "DOWN" & vbNewLine & "LOAD"
                '    ElseIf .ColumnHeader.Columns(i).Label = "RECEIVING" Then
                '        .ColumnHeader.Columns(i).Label = "RCV"
                '    ElseIf .ColumnHeader.Columns(i).Label = "LEVEL 1&2" Then
                '        .ColumnHeader.Columns(i).Label = "LEVEL 1&2"
                '    ElseIf .ColumnHeader.Columns(i).Label = "CAL(RAD)" Then
                '        .ColumnHeader.Columns(i).Label = "RF TEST" & vbNewLine & "(RAD)"
                '    ElseIf .ColumnHeader.Columns(i).Label = "PACKING" Then
                '        .ColumnHeader.Columns(i).Label = "PACK"
                '    End If

                'Next

                .Columns(.ColumnCount - 1).CellType = intcell
                .Columns(.ColumnCount - 1).BackColor = Color.Yellow
                Spread_AutoCol(FpSpread1)
                .FrozenColumnCount = 2
                .Columns(2, FpSpread1.ActiveSheet.ColumnCount - 1).Width = 50
            End With
        End If

        If Spread_Setting(FpSpread2, "FrmWKRESULT") = True Then
            FpSpread2.ActiveSheet.RowCount = 0
            Spread_AutoCol(FpSpread2)
        End If

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
        Formbtn_Authority(Me.FindBtn, Me.Name, "FIND")

    End Sub


    Function Refresh_Result() As Boolean
        Dim qry As String

        Insert_Data("EXEC SP_COMMON_GETWKRESULT '" & Site_id & "'")


        qry = "select MODEL, wk_date," & vbNewLine
        qry = qry & Generate_Spread_Query_1(FpSpread1, "ISNULL((SELECT sum(cnt) FROM TBL_WKRESULT WHERE SITE_ID = A.SITE_ID AND WK_DATE = A.WK_DATE and model = a.model AND C_PRC = (SELECT CODE_ID FROM TBL_CODEMASTER WHERE SITE_ID = A.SITE_ID AND CLASS_ID = 'R0001' AND CODE_NAME = '", "')),0),", 2)
        qry = qry & "ISNULL((SELECT sum(cnt) FROM TBL_wkresult WHERE SITE_ID = A.SITE_ID AND WK_DATE = A.WK_DATE and model = a.model),0) " & vbNewLine
        qry = qry & "FROM TBL_WKresult A "
        qry = qry & "WHERE SITE_ID = '" & Site_id & "'" & vbNewLine
        qry = qry & "  and wk_date between '" & DateTimeInput1.Text & "' and '" & DateTimeInput2.Text & "' " & vbNewLine
        If ComboBoxEx1.Text <> "ALL" Then
            qry = qry & "  and model = '" & ComboBoxEx1.Text & "'" & vbNewLine
        End If
        qry = qry & "GROUP BY SITE_ID, MODEL, wk_date" & vbNewLine
        qry = qry & "ORDER BY SITE_ID, MODEL, wk_date"

        If Query_Spread(FpSpread1, qry, 1) = True Then
            FpSpread1.ActiveSheet.RowCount = FpSpread1.ActiveSheet.RowCount + 1
            FpSpread1.ActiveSheet.Cells(FpSpread1.ActiveSheet.RowCount - 1, 1, FpSpread1.ActiveSheet.RowCount - 1, FpSpread1.ActiveSheet.ColumnCount - 1).CellType = intcell
            FpSpread1.ActiveSheet.Cells(FpSpread1.ActiveSheet.RowCount - 1, 0, FpSpread1.ActiveSheet.RowCount - 1, FpSpread1.ActiveSheet.ColumnCount - 1).BackColor = Color.Yellow
            FpSpread1.ActiveSheet.SetText(FpSpread1.ActiveSheet.RowCount - 1, 0, "TOTAL")

            SPREAD_ROW_TOTAL(FpSpread1, 2, FpSpread1.ActiveSheet.ColumnCount - 1, 1)
            '    Spread_AutoCol(FpSpread1)
            'With FpSpread1.ActiveSheet
            '    For I = 0 To .ColumnCount - 1
            '        If .ColumnHeader.Columns(I).Label = "PD/SCRAP" Then
            '            .ColumnHeader.Columns(I).Label = "PD" & vbNewLine & "SCRAP"
            '        ElseIf .ColumnHeader.Columns(I).Label = "LEVEL 4" Then
            '            .ColumnHeader.Columns(I).Label = "TECH" & vbNewLine & "(L3)"
            '        ElseIf .ColumnHeader.Columns(I).Label = "LEVEL 3" Then
            '            .ColumnHeader.Columns(I).Label = "TECH" & vbNewLine & "(L2)"
            '        ElseIf .ColumnHeader.Columns(I).Label = "LINE QC" Then
            '            .ColumnHeader.Columns(I).Label = "LINE QC"
            '        ElseIf .ColumnHeader.Columns(I).Label = "DM&INFO" Then
            '            .ColumnHeader.Columns(I).Label = "DM&" & vbNewLine & "INFO"
            '        ElseIf .ColumnHeader.Columns(I).Label = "DOWNLOAD" Then
            '            .ColumnHeader.Columns(I).Label = "DOWN" & vbNewLine & "LOAD"
            '        ElseIf .ColumnHeader.Columns(I).Label = "RECEIVING" Then
            '            .ColumnHeader.Columns(I).Label = "RCV"
            '        ElseIf .ColumnHeader.Columns(I).Label = "LEVEL 1&2" Then
            '            .ColumnHeader.Columns(I).Label = "COSM"
            '        ElseIf .ColumnHeader.Columns(I).Label = "PACKING" Then
            '            .ColumnHeader.Columns(I).Label = "PACK"
            '        End If

            '    Next

            'End With

            FpSpread1.ActiveSheet.Columns(2, FpSpread1.ActiveSheet.ColumnCount - 1).Width = 50

        End If


    End Function

    Private Sub FindBtn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FindBtn.Click
        FpSpread1.ActiveSheet.RowCount = 0
        FpSpread2.ActiveSheet.RowCount = 0

        Refresh_Result()
    End Sub

    Private Sub FpSpread1_CellDoubleClick(ByVal sender As Object, ByVal e As FarPoint.Win.Spread.CellClickEventArgs) Handles FpSpread1.CellDoubleClick
        Dim AA, qry As String
        'Dim i As Integer
        AA = ""
        qry = ""

        FpSpread2.ActiveSheet.RowCount = 0
        FpSpread2.ActiveSheet.ColumnCount = 0

        With FpSpread2.ActiveSheet
            .ColumnCount = 4
            .RowCount = 0

            .ColumnHeader.Columns(0).Label = "모델"
            .ColumnHeader.Columns(1).Label = "일반"
            .ColumnHeader.Columns(2).Label = "폴리싱"
            .ColumnHeader.Columns(3).Label = "합계"
            .Columns(1, 3).CellType = intcell
            qry = "SELECT MODEL, " & vbNewLine
            qry = qry & "ISNULL((SELECT SUM(CNT) FROM TBL_WKRESULT WHERE wk_date between '" & FpSpread1.ActiveSheet.GetValue(e.Row, 1) & "' AND '" & FpSpread1.ActiveSheet.GetValue(e.Row, 1) & "' AND C_PRC = (SELECT CODE_ID FROM TBL_CODEMASTER WHERE CLASS_ID = 'R0001' AND CODE_NAME = '" & FpSpread1.ActiveSheet.ColumnHeader.Columns(e.Column).Label & "') AND MODEL = A.MODEL AND OP_CD = 'GEN'),0)," & vbNewLine
            qry = qry & "ISNULL((SELECT SUM(CNT) FROM TBL_WKRESULT WHERE wk_date between '" & FpSpread1.ActiveSheet.GetValue(e.Row, 1) & "' AND '" & FpSpread1.ActiveSheet.GetValue(e.Row, 1) & "' AND C_PRC = (SELECT CODE_ID FROM TBL_CODEMASTER WHERE CLASS_ID = 'R0001' AND CODE_NAME = '" & FpSpread1.ActiveSheet.ColumnHeader.Columns(e.Column).Label & "') AND MODEL = A.MODEL AND OP_CD = 'POL'),0)," & vbNewLine
            qry = qry & "ISNULL((SELECT SUM(CNT) FROM TBL_WKRESULT WHERE wk_date between '" & FpSpread1.ActiveSheet.GetValue(e.Row, 1) & "' AND '" & FpSpread1.ActiveSheet.GetValue(e.Row, 1) & "' AND C_PRC = (SELECT CODE_ID FROM TBL_CODEMASTER WHERE CLASS_ID = 'R0001' AND CODE_NAME = '" & FpSpread1.ActiveSheet.ColumnHeader.Columns(e.Column).Label & "') AND MODEL = A.MODEL),0)" & vbNewLine


            qry = qry & "from tbl_wkresult a" & vbNewLine
            qry = qry & "where model = '" & FpSpread1.ActiveSheet.GetValue(e.Row, 0) & "'" & vbNewLine
            qry = qry & "  and wk_date = '" & FpSpread1.ActiveSheet.GetValue(e.Row, 1) & "'" & vbNewLine
            qry = qry & "group by model, wk_date, hour order by model, wk_date, hour"

        End With

        If e.Row = FpSpread1.ActiveSheet.RowCount - 1 Then
            'If Query_Spread(FpSpread2, qry, 1) = True Then
            '    Spread_AutoCol(FpSpread2)
            '    'FpSpread2.ActiveSheet.Protect = False
            '    FpSpread2.ActiveSheet.Columns(0, FpSpread2.ActiveSheet.ColumnCount - 1).Locked = False
            'End If
        Else
            With FpSpread2.ActiveSheet
                If Query_Spread(FpSpread2, qry, 1) = True Then
                    If .RowCount > 0 And .ColumnCount > 0 Then

                        .RowCount += 1
                        .Cells(.RowCount - 1, 1, .RowCount - 1, .ColumnCount - 1).CellType = intcell
                        .Cells(.RowCount - 1, 0, .RowCount - 1, .ColumnCount - 1).BackColor = Color.Yellow
                        .SetText(.RowCount - 1, 0, "TOTAL")

                        SPREAD_ROW_TOTAL(FpSpread2, 1, .ColumnCount - 1, 1)
                        'Spread_AutoCol(FpSpread2)
                        'FpSpread2.ActiveSheet.Protect = False
                        .Columns(0, .ColumnCount - 1).Locked = False
                    End If
                End If

            End With

        End If
    End Sub
    Function Header_disp(ByVal aa As String) As String
        Dim qry As String
        Dim i As Integer
        Dim OP_RS As ADODB.Recordset

        '        qry = "SELECT DISTINCT (SELECT EMP_NM FROM TBL_EMPMASTER WHERE EMP_NO = OP_CD) AS EMP_NM FROM tbl_wkresult WHERE SITE_ID = '" & Site_id & "' AND WK_DATE BETWEEN '" & DateTimeInput1.Text & "' AND '" & DateTimeInput2.Text & "' AND C_PRC = (SELECT CODE_ID FROM TBL_CODEMASTER WHERE CLASS_ID = 'R0001' AND CODE_NAME ='" & aa & "') ORDER BY EMP_NM"

        OP_RS = Query_RS_ALL("SELECT DISTINCT (SELECT EMP_NM FROM TBL_EMPMASTER WHERE EMP_NO = OP_CD) AS EMP_NM FROM tbl_wkresult WHERE SITE_ID = '" & Site_id & "' AND WK_DATE BETWEEN '" & DateTimeInput1.Text & "' AND '" & DateTimeInput2.Text & "' AND C_PRC = (SELECT CODE_ID FROM TBL_CODEMASTER WHERE CLASS_ID = 'R0001' AND CODE_NAME ='" & aa & "') ORDER BY EMP_NM")

        With FpSpread2.ActiveSheet
            If OP_RS Is Nothing Then
                MessageBox.Show("NO DATA!")
                Header_disp = ""
            Else
                .RowCount = 0
                .ColumnCount = OP_RS.RecordCount

                .AddColumns(0, 1)
                .ColumnHeader.Rows(0).Height = 60
                .ColumnHeader.Columns(0).Label = "WORK HOUR"

                For i = 1 To OP_RS.RecordCount
                    .ColumnHeader.Columns(i).Label = OP_RS(0).Value
                    .Columns(i).CellType = intcell
                    OP_RS.MoveNext()
                Next

                .ColumnCount = .ColumnCount + 1
                .ColumnHeader.Columns(.ColumnCount - 1).Label = "TOTAL"
                .Columns(.ColumnCount - 1).CellType = intcell
                .Columns(.ColumnCount - 1).BackColor = Color.Yellow
                Spread_AutoCol(FpSpread2)
                .FrozenColumnCount = 2
                .Columns(1, .ColumnCount - 1).Width = 50
            End If
        End With

        qry = ""

        qry = "select hour," & vbNewLine & qry
        For i = 1 To FpSpread2.ActiveSheet.ColumnCount - 2
            qry = qry & "		isnull((select sum(cnt) from tbl_wkresult where model = a.model and wk_date = a.wk_date and hour = a.hour and c_prc = (select code_id from tbl_codemaster where class_id = 'R0001' and code_name = '" & aa & "' ) and op_cd = (select emp_no from tbl_empmaster where emp_nm = '" & FpSpread2.ActiveSheet.ColumnHeader.Columns(i).Label & "' )),0)," & vbNewLine
        Next
        qry = qry & "		isnull((select sum(cnt) from tbl_wkresult where model = a.model and wk_date = a.wk_date and hour = a.hour and c_prc = (select code_id from tbl_codemaster where class_id = 'R0001' and code_name = '" & aa & "' )),0)" & vbNewLine
        Header_disp = qry

    End Function

    Private Sub ButtonItem2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonItem2.Click, NewBtn1.Click

    End Sub

    Private Sub ButtonItem3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonItem3.Click, SaveBtn1.Click

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
            If Spread_Print(Me.FpSpread2, "WIP Inventory(Board) Details", 0) = False Then
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