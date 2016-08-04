Public Class FrmPolSummary

    Private Sub CheckedListBox1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles CheckedListBox1.LostFocus
        TextBoxDropDown1.Text = MODEL_SELECTED(CheckedListBox1)
    End Sub


    Private Sub FrmBWIP_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        DockContainerItem1.Text = "실행 메뉴"
        DockContainerItem2.Text = "조회 조건"

        GroupPanel2.Visible = False
        ExpandableSplitter1.Visible = False
        GroupPanel1.Dock = DockStyle.Fill

        FpSpread1.ActiveSheet.RowCount = 0
        FpSpread2.ActiveSheet.RowCount = 0

        DateTimeInput1.Value = Now
        DateTimeInput2.Value = Now

        If Spread_Setting(FpSpread1, Me.Name) = True Then
            FpSpread1.ActiveSheet.RowCount = 0
            Spread_AutoCol(FpSpread1)
        End If

        With FpSpread3.ActiveSheet
            .RowCount = 0
            .ColumnCount = 1
        End With

        If Query_CheckList(Me.CheckedListBox1, "SELECT model_no FROM tbl_modelmaster WHERE site_id = '" & Site_id & "' and active = 'Y' ORDER BY model_no") = True Then
            Me.CheckedListBox1.Items.Add("ALL")
            Me.CheckedListBox1.Text = "ALL"
        End If

        '       If Spread_Setting(FpSpread2, Me.Name) = True Then
        FpSpread2.ActiveSheet.RowCount = 0
        Spread_AutoCol(FpSpread2)
        '        End If

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
        Dim qry As String = ""
        Dim I, J As Integer
        Dim m_qry As String = ""
        m_qry = MODEL_SELECTED(CheckedListBox1)



        qry = "select A.MODEL," & vbNewLine
        qry = qry & "ISNULL((SELECT COUNT(MODEL) FROM  VIEW_fESNMASTER  WHERE SITE_ID = A.SITE_ID AND MODEL = A.MODEL and PSHIP_NO LIKE 'EXP%' and out_esn like 'S%' AND SUBSTRING(PSHIP_NO,4, 8) BETWEEN '" & DateTimeInput1.Text & "' AND '" & DateTimeInput2.Text & "' GROUP BY MODEL ),0)," & vbNewLine
        qry = qry & "ISNULL((SELECT SUM(ACT_QTY) FROM  TBL_LOTMASTER  WHERE SITE_ID = A.SITE_ID AND LOT_NO LIKE 'S%' AND MODEL = A.MODEL GROUP BY MODEL ),0)," & vbNewLine
        qry = qry & "ISNULL((SELECT COUNT(MODEL) FROM  TBL_fESNMASTER_B  WHERE SITE_ID = A.SITE_ID AND MODEL = A.MODEL AND PSHIP_NO LIKE 'EXP%' and out_esn like 'S%'  AND SUBSTRING(CSHIP_NO,4, 8) BETWEEN '" & DateTimeInput1.Text & "' AND '" & DateTimeInput2.Text & "' GROUP BY MODEL ),0)," & vbNewLine
        qry = qry & "isnull((SELECT SUM(CAST(OUT_QTY AS INT)) FROM  TBL_LREPAIRMASTER  WHERE SITE_ID = A.SITE_ID and RESULT = '폴리싱반품' AND ESN IN (SELECT LOT_NO FROM TBL_LOTMASTER WHERE MODEL = A.MODEL AND LOT_NO LIKE 'S%') AND CONVERT(VARCHAR(8), C_DATE, 112) BETWEEN '" & DateTimeInput1.Text & "' AND '" & DateTimeInput2.Text & "'),0)" & vbNewLine
        qry = qry & "FROM  VIEW_FESNMASTER  A" & vbNewLine
        qry = qry & "WHERE SITE_ID = '" & Site_id & "'" & vbNewLine
        '       qry = qry & "and b.active = 'Y'" & vbNewLine

        'If ComboBoxEx1.Text <> "ALL" Then
        '    qry = qry & "  and model_NO = '" & ComboBoxEx1.Text & "'" & vbNewLine
        'End If
        If m_qry <> "" Then
            qry = qry & "AND A.MODEL IN (" & m_qry & ")" & vbNewLine
        End If
        qry = qry & "GROUP BY SITE_ID, MODEL" & vbNewLine
        qry = qry & "ORDER BY MODEL"

        If Query_Spread(FpSpread1, qry, 1) = True Then

            With FpSpread1.ActiveSheet

                For I = 0 To .RowCount - 1
                    For J = 1 To .ColumnCount - 1
                        If .GetValue(I, J) = 0 Then
                            .SetText(I, J, "")
                        End If

                    Next
                Next

                .RowCount = .RowCount + 1
                .Cells(.RowCount - 1, 1, .RowCount - 1, .ColumnCount - 1).CellType = intcell
                .Rows(.RowCount - 1).BackColor = Color.Yellow
                .SetText(.RowCount - 1, 0, "TOTAL")

                SPREAD_ROW_TOTAL(FpSpread1, 1, .ColumnCount - 1, 1)

                Spread_AutoCol(FpSpread1)
                .Columns(1, .ColumnCount - 1).Width = 80
            End With
        End If

    End Function

    Private Sub FindBtn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FindBtn.Click
        FpSpread1.ActiveSheet.RowCount = 0
        FpSpread2.ActiveSheet.RowCount = 0

        GroupPanel2.Visible = False
        ExpandableSplitter1.Visible = False
        GroupPanel1.Dock = DockStyle.Fill

        Refresh_Result()
    End Sub

    Private Sub FpSpread1_CellDoubleClick(ByVal sender As Object, ByVal e As FarPoint.Win.Spread.CellClickEventArgs) Handles FpSpread1.CellDoubleClick

        GroupPanel2.Visible = True
        ExpandableSplitter1.Visible = True
        GroupPanel1.Dock = DockStyle.Top

        Dim QRY As String = ""
        Dim m_qry As String = ""
        m_qry = MODEL_SELECTED(CheckedListBox1)

        With FpSpread2.ActiveSheet
            'If FpSpread1.ActiveSheet.ColumnHeader.Columns(e.Column).Label = "입고수량" Then
            .ColumnCount = 6
            .ColumnHeader.Columns(0).Label = "모델"
            .ColumnHeader.Columns(1).Label = "입고번호"
            .ColumnHeader.Columns(2).Label = "입고수량"
            .ColumnHeader.Columns(3).Label = "WIP 수량"
            .ColumnHeader.Columns(4).Label = "출하수량"
            .ColumnHeader.Columns(5).Label = "반품수량"
            .Columns(2, 5).CellType = intcell


            QRY = "SELECT MODEL, PSHIP_NO, COUNT(MODEL), " & vbNewLine
            QRY = QRY & "	isnull((select sum(act_qty) from TBL_LOTMASTER where model = a.model and LOT_NO in (select PBOX_NO from VIEW_FESNMASTER where pship_no = a.pship_no ) ),0)+ISNULL((SELECT COUNT(MODEL) FROM VIEW_FESNMASTER WHERE PSHIP_NO = A.pship_no AND PBOX_NO IS NULL and out_esn like 'S%'),0)," & vbNewLine
            QRY = QRY & "	isnull((select count(model) from TBL_FESNMASTER_b where model = a.model and pship_no = a.pship_no ),0)," & vbNewLine
            QRY = QRY & "	isnull((select sum(cast(out_qty as int)) from TBL_LrepairMASTER where model = a.model and esn in (select PBOX_NO from VIEW_FESNMASTER where pship_no = a.pship_no ) ),0)" & vbNewLine
            QRY = QRY & "FROM VIEW_FESNMASTER A" & vbNewLine
            QRY = QRY & "WHERE PSHIP_NO LIKE 'EXP%'" & vbNewLine
            QRY = QRY & "and out_esn like 'S%'" & vbNewLine

            QRY = QRY & "AND SUBSTRING(PSHIP_NO,4,8) BETWEEN '" & DateTimeInput1.Text & "' AND '" & DateTimeInput2.Text & "'" & vbNewLine
            QRY = QRY & "AND MODEL = '" & FpSpread1.ActiveSheet.GetValue(e.Row, 0) & "'" & vbNewLine
            QRY = QRY & "GROUP BY MODEL, PSHIP_NO" & vbNewLine
            QRY = QRY & "ORDER BY MODEL, PSHIP_NO" & vbNewLine
            QRY = QRY & "" & vbNewLine


            'ElseIf FpSpread1.ActiveSheet.ColumnHeader.Columns(e.Column).Label = "WIP 수량" Then
            '.ColumnCount = 3
            '.ColumnHeader.Columns(0).Label = "모델"
            '.ColumnHeader.Columns(1).Label = "공정"
            '.ColumnHeader.Columns(2).Label = "공정수량"
            '.Columns(2).CellType = intcell


            'QRY = "SELECT MODEL, (SELECT CODE_NAME FROM TBL_CODEMASTER WHERE CLASS_ID = 'R0001' AND CODE_ID = A.C_PRC), SUM(ACT_QTY) FROM TBL_LOTMASTER A" & vbNewLine
            'QRY = QRY & "WHERE LOT_NO LIKE 'S%'" & vbNewLine
            'QRY = QRY & "AND A.MODEL = '" & FpSpread1.ActiveSheet.GetValue(e.Row, 0) & "'" & vbNewLine
            'QRY = QRY & "GROUP BY MODEL, C_PRC" & vbNewLine
            'QRY = QRY & "ORDER BY MODEL, C_PRC" & vbNewLine
            'QRY = QRY & "" & vbNewLine

            'ElseIf FpSpread1.ActiveSheet.ColumnHeader.Columns(e.Column).Label = "출하수량" Then
            '.ColumnCount = 3
            '.ColumnHeader.Columns(0).Label = "모델"
            '.ColumnHeader.Columns(1).Label = "출하번호"
            '.ColumnHeader.Columns(2).Label = "출하수량"
            '.Columns(2).CellType = intcell


            'QRY = "SELECT MODEL, CSHIP_NO, COUNT(MODEL) FROM TBL_FESNMASTER_B " & vbNewLine
            'QRY = QRY & "WHERE SUBSTRING(CSHIP_NO,4,8) BETWEEN '" & DateTimeInput1.Text & "' AND '" & DateTimeInput2.Text & "'" & vbNewLine
            'QRY = QRY & "AND MODEL = '" & FpSpread1.ActiveSheet.GetValue(e.Row, 0) & "'" & vbNewLine
            'QRY = QRY & "AND PSHIP_NO LIKE 'EXP%'" & vbNewLine
            'QRY = QRY & "GROUP BY MODEL, CSHIP_NO" & vbNewLine
            'QRY = QRY & "ORDER BY MODEL, CSHIP_NO" & vbNewLine
            'QRY = QRY & "" & vbNewLine

            'ElseIf FpSpread1.ActiveSheet.ColumnHeader.Columns(e.Column).Label = "반품수량" Then
            '.ColumnCount = 3
            '.ColumnHeader.Columns(0).Label = "모델"
            '.ColumnHeader.Columns(1).Label = "반품일시"
            '.ColumnHeader.Columns(2).Label = "반품수량"
            '.Columns(2).CellType = intcell


            'QRY = "SELECT B.MODEL, A.C_DATE, OUT_QTY " & vbNewLine
            'QRY = QRY & "FROM TBL_LREPAIRMASTER A, TBL_LOTMASTER B" & vbNewLine
            'QRY = QRY & "WHERE B.LOT_NO LIKE 'S%'" & vbNewLine
            'QRY = QRY & "AND CONVERT(VARCHAR(8), A.C_DATE,112) BETWEEN '" & DateTimeInput1.Text & "' AND '" & DateTimeInput2.Text & "'" & vbNewLine
            'QRY = QRY & "AND A.ESN = B.LOT_NO" & vbNewLine
            'QRY = QRY & "AND A.RESULT = '폴리싱반품'" & vbNewLine
            'QRY = QRY & "AND B.MODEL = '" & FpSpread1.ActiveSheet.GetValue(e.Row, 0) & "'" & vbNewLine
            'QRY = QRY & "ORDER BY B.MODEL, A.C_DATE" & vbNewLine
            'QRY = QRY & "" & vbNewLine
            'Else
            'End If
            If Query_Spread(FpSpread2, QRY, 1) = True Then
                Spread_AutoCol(FpSpread2)
                .Columns(2, 5).Width = 80
                ' FpSpread2.ActiveSheet.Columns(0, FpSpread2.ActiveSheet.ColumnCount - 1).Locked = False
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

    Private Sub ButtonItem7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonItem7.Click
        With FpSpread2.ActiveSheet
            '            Print_BBarcode(.GetValue(.ActiveRowIndex, 0), .GetValue(.ActiveRowIndex, 1), .GetValue(.ActiveRowIndex, 6))
            'Print_USABarcode(.GetValue(.ActiveRowIndex, 0), .GetValue(.ActiveRowIndex, 1), RichTextBox1)
        End With
    End Sub

    Private Sub ButtonItem8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonItem8.Click
        Dim i As Integer
        With FpSpread2.ActiveSheet
            For i = 0 To .RowCount - 1 'exf20140115-01까지 완료, 20140311-01부터 출력하면 됨
                'Print_USABarcode(.GetValue(i, 0), .GetValue(i, 1), RichTextBox1)
                '                Print_BBarcode(.GetValue(i, 0), .GetValue(i, 1), .GetValue(i, 6))
                System.Threading.Thread.Sleep(500)
            Next
        End With
    End Sub

    Private Sub FpSpread2_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles FpSpread2.KeyDown
        If e.Control = True And e.KeyCode = Keys.C Then
            If Me.FpSpread2.ActiveSheet.RowCount > 0 Then
                If Me.FpSpread2.ActiveSheet.ActiveColumnIndex = 0 Then
                    Clipboard.SetText(Me.FpSpread2.ActiveSheet.GetText(Me.FpSpread2.ActiveSheet.ActiveRowIndex, 0))
                    'MessageBox.Show("Copy " & Me.FpSpread1.ActiveSheet.GetValue(Me.FpSpread1.ActiveSheet.ActiveRowIndex, 0), "COPY MESSAGE")
                End If
            End If
        End If
    End Sub



    Private Sub ButtonItem9_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonItem9.Click

        With FpSpread3.ActiveSheet
            .RowCount = 0

            .RowCount = FpSpread2.ActiveSheet.RowCount

            For i As Integer = 0 To FpSpread2.ActiveSheet.RowCount - 1
                .SetValue(i, 0, FpSpread2.ActiveSheet.GetValue(i, 7))
            Next

            File_Save_meiddet(SaveFileDialog1, FpSpread3)
        End With


    End Sub

    Private Sub ButtonItem10_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonItem10.Click
        With FpSpread2.ActiveSheet
            '            Print_BBarcode(.GetValue(.ActiveRowIndex, 0), .GetValue(.ActiveRowIndex, 1), .GetValue(.ActiveRowIndex, 6))
            Print_LOTBarcode(.GetValue(.ActiveRowIndex, 0), .GetValue(.ActiveRowIndex, 1), .GetValue(.ActiveRowIndex, 3), RichTextBox1, .GetValue(.ActiveRowIndex, 7))
        End With

    End Sub
End Class