Public Class FrmWKRESULT

    Private Sub FrmWKRESULT_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        'Dim i As Integer

        DockContainerItem1.Text = "Form Menu"
        DockContainerItem2.Text = "Retrieve Condition"

        DateTimeInput1.Value = Now
        DateTimeInput2.Value = Now

        If Query_Combo(Me.ComboBoxEx1, "SELECT model_no FROM tbl_modelmaster WHERE site_id = '" & Site_id & "' and active = 'Y' ORDER BY model_no") = True Then
        End If
        Me.ComboBoxEx1.Items.Add("ALL")
        Me.ComboBoxEx1.Text = "ALL"


        Me.ComboBoxEx2.Items.Add("00")
        Me.ComboBoxEx2.Items.Add("01")
        Me.ComboBoxEx2.Items.Add("02")
        Me.ComboBoxEx2.Items.Add("03")
        Me.ComboBoxEx2.Items.Add("04")
        Me.ComboBoxEx2.Items.Add("05")
        Me.ComboBoxEx2.Items.Add("06")
        Me.ComboBoxEx2.Items.Add("07")
        Me.ComboBoxEx2.Items.Add("08")
        Me.ComboBoxEx2.Items.Add("09")
        Me.ComboBoxEx2.Items.Add("10")
        Me.ComboBoxEx2.Items.Add("11")
        Me.ComboBoxEx2.Items.Add("12")
        Me.ComboBoxEx2.Items.Add("13")
        Me.ComboBoxEx2.Items.Add("14")
        Me.ComboBoxEx2.Items.Add("15")
        Me.ComboBoxEx2.Items.Add("16")
        Me.ComboBoxEx2.Items.Add("17")
        Me.ComboBoxEx2.Items.Add("18")
        Me.ComboBoxEx2.Items.Add("19")
        Me.ComboBoxEx2.Items.Add("20")
        Me.ComboBoxEx2.Items.Add("21")
        Me.ComboBoxEx2.Items.Add("22")
        Me.ComboBoxEx2.Items.Add("23")
        Me.ComboBoxEx2.Text = "00"


        Me.ComboBoxEx3.Items.Add("00")
        Me.ComboBoxEx3.Items.Add("01")
        Me.ComboBoxEx3.Items.Add("02")
        Me.ComboBoxEx3.Items.Add("03")
        Me.ComboBoxEx3.Items.Add("04")
        Me.ComboBoxEx3.Items.Add("05")
        Me.ComboBoxEx3.Items.Add("06")
        Me.ComboBoxEx3.Items.Add("07")
        Me.ComboBoxEx3.Items.Add("08")
        Me.ComboBoxEx3.Items.Add("09")
        Me.ComboBoxEx3.Items.Add("10")
        Me.ComboBoxEx3.Items.Add("11")
        Me.ComboBoxEx3.Items.Add("12")
        Me.ComboBoxEx3.Items.Add("13")
        Me.ComboBoxEx3.Items.Add("14")
        Me.ComboBoxEx3.Items.Add("15")
        Me.ComboBoxEx3.Items.Add("16")
        Me.ComboBoxEx3.Items.Add("17")
        Me.ComboBoxEx3.Items.Add("18")
        Me.ComboBoxEx3.Items.Add("19")
        Me.ComboBoxEx3.Items.Add("20")
        Me.ComboBoxEx3.Items.Add("21")
        Me.ComboBoxEx3.Items.Add("22")
        Me.ComboBoxEx3.Items.Add("23")
        Me.ComboBoxEx3.Text = "23"


        If Spread_Setting_ByCode(FpSpread1, "SP_COMMON_GETCODEMASTER", "R0001", "INT") = True Then
            With FpSpread1.ActiveSheet
                .RowCount = 0
                .AddColumns(0, 1)
                .ColumnHeader.Rows(0).Height = 40

                .ColumnHeader.Columns(0).Label = "WORK DATE"
                '                .ColumnHeader.Cells(0, 0).Text = "AA" & vbNewLine & "bbb"
                .AddColumns(1, 1)
                .ColumnHeader.Columns(1).Label = "HOUR"
                'FpSpread1.ActiveSheet.AddColumns(2, 1)
                'FpSpread1.ActiveSheet.ColumnHeader.Columns(2).Label = "HOUR"
                .ColumnCount = FpSpread1.ActiveSheet.ColumnCount + 1
                .ColumnHeader.Columns(FpSpread1.ActiveSheet.ColumnCount - 1).Label = "TOTAL"

                'For i = 0 To .ColumnCount - 1
                '    If .ColumnHeader.Columns(i).Label = "PD/SCRAP" Then
                '        .ColumnHeader.Columns(i).Label = "PD" & vbNewLine & "SCRAP"
                '    ElseIf .ColumnHeader.Columns(i).Label = "LEVEL 4" Then
                '        .ColumnHeader.Columns(i).Label = "TECH" & vbNewLine & "(L3)"
                '    ElseIf .ColumnHeader.Columns(i).Label = "LEVEL 3" Then
                '        .ColumnHeader.Columns(i).Label = "TECH" & vbNewLine & "(L2)"
                '    ElseIf .ColumnHeader.Columns(i).Label = "DATA ENTRY" Then
                '        .ColumnHeader.Columns(i).Label = "DATA" & vbNewLine & "ENTRY"
                '    ElseIf .ColumnHeader.Columns(i).Label = "DM&INFO" Then
                '        .ColumnHeader.Columns(i).Label = "DM&" & vbNewLine & "INFO"
                '    ElseIf .ColumnHeader.Columns(i).Label = "DOWNLOAD" Then
                '        .ColumnHeader.Columns(i).Label = "DOWN" & vbNewLine & "LOAD"
                '    ElseIf .ColumnHeader.Columns(i).Label = "RECEIVING" Then
                '        .ColumnHeader.Columns(i).Label = "RCV"
                '    ElseIf .ColumnHeader.Columns(i).Label = "LEVEL 1&2" Then
                '        .ColumnHeader.Columns(i).Label = "COSM"
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
                .Columns(1, FpSpread1.ActiveSheet.ColumnCount - 1).Width = 50
            End With
        End If

        If Spread_Setting(FpSpread2, Me.Name) = True Then
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
        Formbim_Authority(Me.FindBtn, Me.Name, "FIND")

    End Sub


    Function Refresh_Result() As Boolean
        Dim qry As String

        Insert_Data("EXEC SP_COMMON_GETWKRESULT '" & Site_id & "'")

        qry = "select WK_DATE, HOUR, " & vbNewLine
        qry = qry & Generate_Spread_Query_1(FpSpread1, "ISNULL((SELECT sum(cnt) FROM TBL_WKRESULT WHERE SITE_ID = A.SITE_ID AND WK_DATE = A.WK_DATE AND HOUR = A.HOUR AND C_PRC = (SELECT CODE_ID FROM TBL_CODEMASTER WHERE SITE_ID = A.SITE_ID AND CLASS_ID = 'R0001' AND CODE_NAME = '", "')),0),", 2)
        qry = qry & "ISNULL((SELECT sum(cnt) FROM TBL_wkresult WHERE SITE_ID = A.SITE_ID AND WK_DATE = A.WK_DATE AND HOUR = A.HOUR),0) " & vbNewLine
        qry = qry & "FROM TBL_WKresult A "
        qry = qry & "WHERE SITE_ID = '" & Site_id & "'" & vbNewLine
        qry = qry & "  and CAST(CONVERT(VARCHAR(10), wk_date, 101) + ' ' + CAST(HOUR AS VARCHAR(2)) + ':00:00' AS DATETIME) between CAST('" & DateTimeInput1.Text & " " & ComboBoxEx2.Text & ":00:00' AS DATETIME) and CAST('" & DateTimeInput2.Text & " " & ComboBoxEx3.Text & ":23:59' AS DATETIME)" & vbNewLine

        If ComboBoxEx1.Text <> "ALL" Then
            qry = qry & "  and model = '" & ComboBoxEx1.Text & "'" & vbNewLine
        End If

        qry = qry & "GROUP BY SITE_ID, WK_DATE, HOUR" & vbNewLine
        qry = qry & "ORDER BY SITE_ID, WK_DATE, HOUR"

        If Query_Spread(FpSpread1, qry, 1) = True Then
            FpSpread1.ActiveSheet.RowCount = FpSpread1.ActiveSheet.RowCount + 1
            FpSpread1.ActiveSheet.Cells(FpSpread1.ActiveSheet.RowCount - 1, 1, FpSpread1.ActiveSheet.RowCount - 1, FpSpread1.ActiveSheet.ColumnCount - 1).CellType = intcell
            FpSpread1.ActiveSheet.Cells(FpSpread1.ActiveSheet.RowCount - 1, 0, FpSpread1.ActiveSheet.RowCount - 1, FpSpread1.ActiveSheet.ColumnCount - 1).BackColor = Color.Yellow
            FpSpread1.ActiveSheet.SetText(FpSpread1.ActiveSheet.RowCount - 1, 0, "TOTAL")

            SPREAD_ROW_TOTAL(FpSpread1, 2, FpSpread1.ActiveSheet.ColumnCount - 1, 1)
            'FpSpread1.ActiveSheet.Columns(1, FpSpread1.ActiveSheet.ColumnCount - 1).Width = 50
            Spread_AutoCol(FpSpread1)
        End If


    End Function

    Private Sub FindBtn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FindBtn.Click

        FpSpread1.ActiveSheet.RowCount = 0
        FpSpread2.ActiveSheet.RowCount = 0

        Refresh_Result()
    End Sub

    Private Sub FpSpread1_CellDoubleClick(ByVal sender As Object, ByVal e As FarPoint.Win.Spread.CellClickEventArgs) Handles FpSpread1.CellDoubleClick
        Dim AA As String
        AA = ""

        If AA = "TOTAL" Then
            MessageBox.Show("SELECT WORKCENTER!")
            Exit Sub
        End If

        Dim QRY As String

        With FpSpread2.ActiveSheet
            .ColumnCount = 4
            .RowCount = 0

            .ColumnHeader.Columns(0).Label = "모델"
            .ColumnHeader.Columns(1).Label = "일반"
            .ColumnHeader.Columns(2).Label = "폴리싱"
            .ColumnHeader.Columns(3).Label = "합계"
            .Columns(1, 3).CellType = intcell

            'OP_RS = Query_RS_ALL("SELECT DISTINCT (SELECT EMP_NM FROM TBL_EMPMASTER WHERE EMP_NO = OP_CD) AS EMP_NM FROM tbl_wkresult WHERE SITE_ID = '" & Site_id & "' AND WK_DATE BETWEEN '" & DateTimeInput1.Text & "' AND '" & DateTimeInput2.Text & "' AND C_PRC = (SELECT CODE_ID FROM TBL_CODEMASTER WHERE CLASS_ID = 'R0001' AND CODE_NAME ='" & FpSpread1.ActiveSheet.ColumnHeader.Columns(e.Column).Label & "') ORDER BY EMP_NM")

            'If OP_RS Is Nothing Then
            '    MessageBox.Show("NO DATA!")
            '    Exit Sub
            'Else
            '    .ColumnCount = OP_RS.RecordCount + 1
            '    .ColumnHeader.Columns(I).Label = "MODEL"

            '    For I = 1 To OP_RS.RecordCount
            '        .ColumnHeader.Columns(I).Label = OP_RS(0).Value
            '        OP_RS.MoveNext()
            '    Next

            '    .ColumnCount = .ColumnCount + 1
            '    .ColumnHeader.Columns(.ColumnCount - 1).Label = "TOTAL"

            'End If

            QRY = "SELECT MODEL, " & vbNewLine
            'QRY = QRY & "(SELECT ISNULL(SUM(CAST(ACT_QTY AS INT)) ,0)" & vbNewLine
            'QRY = QRY & "FROM" & vbNewLine
            'QRY = QRY & "(" & vbNewLine
            'QRY = QRY & "SELECT DISTINCT ESN, CASE WHEN CAST(OUT_QTY AS INT) > 0 THEN OUT_QTY ELSE ACT_QTY END AS act_qty FROM tbl_Lrepairmaster " & vbNewLine
            'QRY = QRY & "WHERE process = (SELECT CODE_ID FROM TBL_CODEMASTER WHERE CLASS_ID = 'R0001' AND CODE_NAME = '" & FpSpread1.ActiveSheet.Columns(e.Column).Label & "')" & vbNewLine
            'QRY = QRY & "AND C_DATE BETWEEN '" & FpSpread1.ActiveSheet.GetValue(e.Row, 0) & " " & FpSpread1.ActiveSheet.GetValue(e.Row, 1) & ":00:00' AND '" & FpSpread1.ActiveSheet.GetValue(e.Row, 0) & " " & FpSpread1.ActiveSheet.GetValue(e.Row, 1) & ":59:59'" & vbNewLine
            'QRY = QRY & "AND ESN NOT LIKE 'S%'" & vbNewLine
            'QRY = QRY & "AND ESN IN (SELECT LOT_NO FROM TBL_LOTMASTER WHERE MODEL = A.MODEL)" & vbNewLine
            'QRY = QRY & "AND PROCESS+RESERV1 <> 'K8000K8000'" & vbNewLine
            'QRY = QRY & ") B), " & vbNewLine
            'QRY = QRY & "(SELECT ISNULL(SUM(CAST(ACT_QTY AS INT)) ,0)" & vbNewLine
            'QRY = QRY & "FROM" & vbNewLine
            'QRY = QRY & "(" & vbNewLine
            'QRY = QRY & "SELECT DISTINCT ESN, CASE WHEN CAST(OUT_QTY AS INT) > 0 THEN OUT_QTY ELSE ACT_QTY END AS act_qty FROM tbl_Lrepairmaster " & vbNewLine
            'QRY = QRY & "WHERE process = (SELECT CODE_ID FROM TBL_CODEMASTER WHERE CLASS_ID = 'R0001' AND CODE_NAME = '" & FpSpread1.ActiveSheet.Columns(e.Column).Label & "')" & vbNewLine
            'QRY = QRY & "AND C_DATE BETWEEN '" & FpSpread1.ActiveSheet.GetValue(e.Row, 0) & " " & FpSpread1.ActiveSheet.GetValue(e.Row, 1) & ":00:00' AND '" & FpSpread1.ActiveSheet.GetValue(e.Row, 0) & " " & FpSpread1.ActiveSheet.GetValue(e.Row, 1) & ":59:59'" & vbNewLine
            'QRY = QRY & "AND ESN LIKE 'S%'" & vbNewLine
            'QRY = QRY & "AND ESN IN (SELECT LOT_NO FROM TBL_LOTMASTER WHERE MODEL = A.MODEL)" & vbNewLine
            'QRY = QRY & "AND PROCESS+RESERV1 <> 'K8000K8000'" & vbNewLine
            'QRY = QRY & ") C), " & vbNewLine
            'QRY = QRY & "(SELECT ISNULL(SUM(CAST(ACT_QTY AS INT)) ,0)" & vbNewLine
            'QRY = QRY & "FROM" & vbNewLine
            'QRY = QRY & "(" & vbNewLine
            'QRY = QRY & "SELECT DISTINCT ESN, CASE WHEN CAST(OUT_QTY AS INT) > 0 THEN OUT_QTY ELSE ACT_QTY END AS act_qty FROM tbl_Lrepairmaster " & vbNewLine
            'QRY = QRY & "WHERE process = (SELECT CODE_ID FROM TBL_CODEMASTER WHERE CLASS_ID = 'R0001' AND CODE_NAME = '" & FpSpread1.ActiveSheet.Columns(e.Column).Label & "')" & vbNewLine
            'QRY = QRY & "AND C_DATE BETWEEN '" & FpSpread1.ActiveSheet.GetValue(e.Row, 0) & " " & FpSpread1.ActiveSheet.GetValue(e.Row, 1) & ":00:00' AND '" & FpSpread1.ActiveSheet.GetValue(e.Row, 0) & " " & FpSpread1.ActiveSheet.GetValue(e.Row, 1) & ":59:59'" & vbNewLine
            'QRY = QRY & "AND ESN IN (SELECT LOT_NO FROM TBL_LOTMASTER WHERE MODEL = A.MODEL)" & vbNewLine
            'QRY = QRY & "AND PROCESS+RESERV1 <> 'K8000K8000'" & vbNewLine
            'QRY = QRY & ") D) " & vbNewLine

            If FpSpread1.ActiveSheet.GetText(FpSpread1.ActiveSheet.ActiveRowIndex, 0) = "TOTAL" Then

                QRY = QRY & "ISNULL((SELECT SUM(CNT) FROM TBL_WKRESULT WHERE wk_date between '" & DateTimeInput1.Text & "' AND '" & DateTimeInput2.Text & "' AND C_PRC = (SELECT CODE_ID FROM TBL_CODEMASTER WHERE CLASS_ID = 'R0001' AND CODE_NAME = '" & FpSpread1.ActiveSheet.ColumnHeader.Columns(e.Column).Label & "') AND MODEL = A.MODEL AND OP_CD = 'GEN'),0)," & vbNewLine
                QRY = QRY & "ISNULL((SELECT SUM(CNT) FROM TBL_WKRESULT WHERE wk_date between '" & DateTimeInput1.Text & "' AND '" & DateTimeInput2.Text & "' AND C_PRC = (SELECT CODE_ID FROM TBL_CODEMASTER WHERE CLASS_ID = 'R0001' AND CODE_NAME = '" & FpSpread1.ActiveSheet.ColumnHeader.Columns(e.Column).Label & "') AND MODEL = A.MODEL AND OP_CD = 'POL'),0)," & vbNewLine
                QRY = QRY & "ISNULL((SELECT SUM(CNT) FROM TBL_WKRESULT WHERE wk_date between '" & DateTimeInput1.Text & "' AND '" & DateTimeInput2.Text & "' AND C_PRC = (SELECT CODE_ID FROM TBL_CODEMASTER WHERE CLASS_ID = 'R0001' AND CODE_NAME = '" & FpSpread1.ActiveSheet.ColumnHeader.Columns(e.Column).Label & "') AND MODEL = A.MODEL),0)" & vbNewLine

            Else
                QRY = QRY & "ISNULL((SELECT SUM(CNT) FROM TBL_WKRESULT WHERE wk_date between '" & FpSpread1.ActiveSheet.GetText(FpSpread1.ActiveSheet.ActiveRowIndex, 0) & "' AND '" & FpSpread1.ActiveSheet.GetText(FpSpread1.ActiveSheet.ActiveRowIndex, 0) & "' AND HOUR = '" & FpSpread1.ActiveSheet.GetText(FpSpread1.ActiveSheet.ActiveRowIndex, 1) & "' AND C_PRC = (SELECT CODE_ID FROM TBL_CODEMASTER WHERE CLASS_ID = 'R0001' AND CODE_NAME = '" & FpSpread1.ActiveSheet.ColumnHeader.Columns(e.Column).Label & "') AND MODEL = A.MODEL AND OP_CD = 'GEN'),0)," & vbNewLine
                QRY = QRY & "ISNULL((SELECT SUM(CNT) FROM TBL_WKRESULT WHERE wk_date between '" & FpSpread1.ActiveSheet.GetText(FpSpread1.ActiveSheet.ActiveRowIndex, 0) & "' AND '" & FpSpread1.ActiveSheet.GetText(FpSpread1.ActiveSheet.ActiveRowIndex, 0) & "' AND HOUR = '" & FpSpread1.ActiveSheet.GetText(FpSpread1.ActiveSheet.ActiveRowIndex, 1) & "' AND C_PRC = (SELECT CODE_ID FROM TBL_CODEMASTER WHERE CLASS_ID = 'R0001' AND CODE_NAME = '" & FpSpread1.ActiveSheet.ColumnHeader.Columns(e.Column).Label & "') AND MODEL = A.MODEL AND OP_CD = 'POL'),0)," & vbNewLine
                QRY = QRY & "ISNULL((SELECT SUM(CNT) FROM TBL_WKRESULT WHERE wk_date between '" & FpSpread1.ActiveSheet.GetText(FpSpread1.ActiveSheet.ActiveRowIndex, 0) & "' AND '" & FpSpread1.ActiveSheet.GetText(FpSpread1.ActiveSheet.ActiveRowIndex, 0) & "' AND HOUR = '" & FpSpread1.ActiveSheet.GetText(FpSpread1.ActiveSheet.ActiveRowIndex, 1) & "' AND C_PRC = (SELECT CODE_ID FROM TBL_CODEMASTER WHERE CLASS_ID = 'R0001' AND CODE_NAME = '" & FpSpread1.ActiveSheet.ColumnHeader.Columns(e.Column).Label & "') AND MODEL = A.MODEL),0)" & vbNewLine
            End If

            QRY = QRY & "FROM TBL_WKRESULT A " & vbNewLine
            QRY = QRY & "WHERE wk_date between '" & DateTimeInput1.Text & "' AND '" & DateTimeInput2.Text & "' AND C_PRC = (SELECT CODE_ID FROM TBL_CODEMASTER WHERE CLASS_ID = 'R0001' AND CODE_NAME = '" & FpSpread1.ActiveSheet.ColumnHeader.Columns(e.Column).Label & "')" & vbNewLine
            QRY = QRY & "GROUP BY MODEL" & vbNewLine
            QRY = QRY & "ORDER BY MODEL"

            If Query_Spread(FpSpread2, QRY, 1) = True Then
                Spread_AutoCol(FpSpread2)
                'FpSpread2.ActiveSheet.Protect = False
                .RowCount = .RowCount + 1
                .Rows(.RowCount - 1).BackColor = Color.Yellow
                .SetValue(.RowCount - 1, 0, "TOTAL")
                .Columns(1, .ColumnCount - 1).CellType = intcell

                SPREAD_ROW_TOTAL_LTD(FpSpread2, 1, .ColumnCount - 1, .RowCount - 1)
                FpSpread2.ActiveSheet.Columns(0, FpSpread2.ActiveSheet.ColumnCount - 1).Locked = False
            End If
        End With

    End Sub

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