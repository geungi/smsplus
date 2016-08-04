Public Class FrmBWIP

    Private Sub CheckedListBox1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles CheckedListBox1.LostFocus
        TextBoxDropDown1.Text = MODEL_SELECTED(CheckedListBox1)
    End Sub


    Private Sub FrmBWIP_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        DockContainerItem1.Text = "실행 메뉴"
        DockContainerItem2.Text = "조회 조건"

        'With FpSpread1.ActiveSheet
        '    If Spread_Setting_ByCode(FpSpread1, "SP_COMMON_GETCODEMASTER", "R0001", "INT") = True Then
        '        .RowCount = 0
        '        .AddColumns(0, 1)
        '        .ColumnHeader.Columns(0).Label = "MODEL"
        '        .ColumnCount = .ColumnCount + 1
        '        .ColumnHeader.Columns(.ColumnCount - 1).Label = "TOTAL"
        '        .Columns(0).CellType = textcell
        '        .Columns(1, .ColumnCount - 1).CellType = intcell
        '        .Columns(.ColumnCount - 1).BackColor = Color.Yellow
        '        Spread_AutoCol(FpSpread1)
        '        FpSpread1.ActiveSheet.FrozenColumnCount = 1
        '    End If
        'End With

        GroupPanel2.Visible = False
        ExpandableSplitter1.Visible = False
        GroupPanel1.Dock = DockStyle.Fill

        FpSpread1.ActiveSheet.RowCount = 0
        FpSpread2.ActiveSheet.RowCount = 0

        With FpSpread1.ActiveSheet
            If Spread_Setting_ByCode(FpSpread1, "SP_COMMON_GETCODEMASTER", "R0001", "INT") = True Then
                .RowCount = 0
                .AddColumns(0, 1)
                .ColumnHeader.Columns(0).Label = "MODEL NAME"
                .AddColumns(0, 1)
                .ColumnHeader.Columns(0).Label = "MODEL"
                .ColumnCount = .ColumnCount + 1
                .ColumnHeader.Columns(.ColumnCount - 1).Label = "TOTAL"
                .Columns(0, 1).CellType = textcell
                .Columns(2, .ColumnCount - 1).CellType = intcell
                .Columns(.ColumnCount - 1).BackColor = Color.Yellow
                Spread_AutoCol(FpSpread1)
                FpSpread1.ActiveSheet.FrozenColumnCount = 1
            End If
        End With

        'If Query_Combo(Me.ComboBoxEx1, "SELECT model_no FROM TBL_MODELMASTER WHERE site_id = '" & Site_id & "' and active = 'Y' ORDER BY model_no") = True Then
        'End If
        'Me.ComboBoxEx1.Items.Add("ALL")
        'Me.ComboBoxEx1.Text = "ALL"

        With FpSpread3.ActiveSheet
            .RowCount = 0
            .ColumnCount = 1
        End With

        If Query_CheckList(Me.CheckedListBox1, "SELECT model_no FROM tbl_modelmaster WHERE site_id = '" & Site_id & "' and active = 'Y' ORDER BY model_no") = True Then
            Me.CheckedListBox1.Items.Add("ALL")
            Me.CheckedListBox1.Text = "ALL"
        End If


        If Spread_Setting(FpSpread2, "FrmBWIP") = True Then
            FpSpread2.ActiveSheet.RowCount = 0
            Spread_AutoCol(FpSpread2)
        End If


        ComboBoxEx2.Items.Add("ALL")
        ComboBoxEx2.Items.Add("일반")
        ComboBoxEx2.Items.Add("폴리싱")
        ComboBoxEx2.Text = "ALL"

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
        Dim qry, tb_esn, TB_MOD As String
        Dim I, J As Integer
        Dim m_qry As String = ""
        m_qry = MODEL_SELECTED(CheckedListBox1)
        Dim P_QRY As String = ""
        Dim L_QRY As String = ""


        tb_esn = "TBL_PRODMASTER"
        tb_esn = "TBL_PRODMASTER_B"
        TB_MOD = "TBL_MODELMASTER"

        qry = "select B.MODEL_NO,B.MODEL_DESC," & vbNewLine
        qry = qry & "isnull((SELECT count(prod_no) FROM tbl_prodMASTER WHERE SITE_ID = B.SITE_ID AND MODEL = B.MODEL_NO  AND C_PRC = 'W1000'),0), " & vbNewLine
        qry = qry & "isnull((SELECT count(prod_no) FROM tbl_prodMASTER_b WHERE SITE_ID = B.SITE_ID AND MODEL = B.MODEL_NO AND C_PRC = 'W2000'),0), " & vbNewLine
        qry = qry & "isnull((SELECT count(prod_no) FROM tbl_prodMASTER_b WHERE SITE_ID = B.SITE_ID AND MODEL = B.MODEL_NO AND C_PRC = 'W3000'),0), " & vbNewLine

        '        qry = qry & Generate_Spread_Query_1(FpSpread1, "(Select COUNT(PROD_NO) FROM  TBL_PRODMASTER  WHERE SITE_ID = B.SITE_ID And MODEL = B.MODEL_NO And C_PRC = (Select CODE_ID FROM TBL_CODEMASTER WHERE SITE_ID = B.SITE_ID And CLASS_ID = 'R0001' AND CODE_NAME = '", "') " & L_QRY & "),", 2)
        qry = qry & "isnull((SELECT count(prod_no) FROM  view_prodMASTER WHERE SITE_ID = B.SITE_ID AND MODEL = B.MODEL_NO  " & L_QRY & "),0) " & vbNewLine
        qry = qry & "FROM  " & TB_MOD & "  B "
        qry = qry & "WHERE SITE_ID = '" & Site_id & "'" & vbNewLine
        qry = qry & "and b.active = 'Y'" & vbNewLine

        'If ComboBoxEx1.Text <> "ALL" Then
        '    qry = qry & "  and model_NO = '" & ComboBoxEx1.Text & "'" & vbNewLine
        'End If
        If m_qry <> "" Then
            qry = qry & "AND b.MODEL_no IN (" & m_qry & ")" & vbNewLine
        End If


        qry = qry & "ORDER BY SITE_ID, MODEL_NO, RESERV1"

        If Query_Spread(FpSpread1, qry, 1) = True Then

            With FpSpread1.ActiveSheet

                For I = 0 To .RowCount - 1
                    For J = 2 To .ColumnCount - 1
                        If .GetValue(I, J) = 0 Then
                            .SetText(I, J, "")
                        End If

                    Next
                Next

                .RowCount = .RowCount + 1
                .Cells(.RowCount - 1, 2, .RowCount - 1, .ColumnCount - 1).CellType = intcell
                .Rows(.RowCount - 1).BackColor = Color.Yellow
                .SetText(.RowCount - 1, 0, "TOTAL")

                SPREAD_ROW_TOTAL(FpSpread1, 2, .ColumnCount - 1, 1)

                Spread_AutoCol(FpSpread1)
                .Columns(2, .ColumnCount - 1).Width = 80
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

        Dim QRY, TB_ESN, TB_MOD As String
        Dim m_qry As String = ""
        m_qry = MODEL_SELECTED(CheckedListBox1)

        TB_ESN = "TBL_FESNMASTER_K"
        TB_MOD = "TBL_MODELMASTER"

        'QRY = "select esn, model, accountno, qareject, R_DATE, datediff(day, c_date, getdate()),/*CASE datediff(day, convert(datetime, LEFT(OUTRANO,8),101) , GETDATE()) - (SELECT COUNT(DT) FROM TBL_HOLIDAY WHERE DT BETWEEN convert(datetime, LEFT(OUTRANO,8),101) AND convert(datetime, GETDATE(),101)) + 1 WHEN 0 THEN 1 ELSE datediff(day, convert(datetime, LEFT(OUTRANO,8),101) , GETDATE()) - (SELECT COUNT(DT) FROM TBL_HOLIDAY WHERE DT BETWEEN convert(datetime, LEFT(OUTRANO,8),101) AND convert(datetime, GETDATE(),101)) + 1 END,*/" & vbNewLine
        'QRY = QRY & "isnull(t_def_cd,''), "
        ''        QRY = QRY & "   (select ISNULL(TRG01,'') + ':' + ISNULL(TRG02,'') FROM VIEW_LGTRIAGE WHERE OBID = (select OBID from tbl_esnmaster where site_id = A.SITE_ID and esn = A.ESN))," & vbNewLine
        'QRY = QRY & "   ISNULL(RESERV5,''), ISNULL(REP_CD,'')," & vbNewLine
        'QRY = QRY & "   (select code_name from TBL_CODEMASTER WHERE SITE_ID = A.SITE_ID AND CLASS_ID = 'R0001' AND CODE_id = a.c_prc)," & vbNewLine
        'QRY = QRY & "   u_date, ISNULL((SELECT EMP_NM FROM TBL_EMPMASTER WHERE SITE_ID = A.SITE_ID AND EMP_NO = A.u_person), 'SYSADMIN'), ISNULL(INBOXID,'')" & vbNewLine
        'QRY = QRY & "FROM " & TB_ESN & " A" & vbNewLine
        'QRY = QRY & "WHERE SITE_ID = '" & Site_id & "'" & vbNewLine

        'If ComboBoxEx1.Text <> "ALL" Then
        '    QRY = QRY & "  AND MODEL = '" & ComboBoxEx1.Text & "'" & vbNewLine
        'End If

        'With FpSpread1.ActiveSheet
        '    If .ColumnHeader.Columns(e.Column).Label <> "TOTAL" Then
        '        QRY = QRY & "  AND C_PRC = (SELECT CODE_ID FROM TBL_CODEMASTER WHERE SITE_ID = A.SITE_ID AND CLASS_ID = 'R0001' AND CODE_NAME = '" & .ColumnHeader.Columns(e.Column).Label & "')" & vbNewLine
        '    End If

        '    If .GetValue(e.Row, 0) <> "TOTAL" Then
        '        QRY = QRY & "  AND MODEL = '" & .GetValue(e.Row, 0) & "'" & vbNewLine
        '    End If
        'End With

        'QRY = QRY & "ORDER BY U_DATE" & vbNewLine
        With FpSpread1.ActiveSheet
            'If .ColumnHeader.Columns(e.Column).Label = "PREINSPECTION" Then
            '    QRY = "select out_esn AS LOT_NO, model, 1, 1, pship_no,NULL,NULL,NULL,NULL," & vbNewLine
            '    QRY = QRY & "   (select code_name from TBL_CODEMASTER WHERE SITE_ID = A.SITE_ID AND CLASS_ID = 'R0001' AND CODE_id = a.c_prc)," & vbNewLine
            '    QRY = QRY & "   u_date, ISNULL((SELECT EMP_NM FROM TBL_EMPMASTER WHERE SITE_ID = A.SITE_ID AND EMP_NO = A.u_person), 'SYSADMIN')" & vbNewLine
            '    QRY = QRY & "FROM TBL_fesnMASTER_k A" & vbNewLine
            'Else
            QRY = "select LOT_NO, model, INIT_QTY, ACT_QTY, NULL,NULL,T_DEF_CD,(SELECT CODE_NAME FROM TBL_DEFECT WHERE CODE_ID = A.T_DEF_CD),NULL," & vbNewLine
            QRY = QRY & "   (select code_name from TBL_CODEMASTER WHERE SITE_ID = A.SITE_ID AND CLASS_ID = '10024' AND CODE_id = a.c_prc)," & vbNewLine
            QRY = QRY & "   u_date, ISNULL((SELECT EMP_NM FROM TBL_EMPMASTER WHERE SITE_ID = A.SITE_ID AND EMP_NO = A.u_person), 'SYSADMIN')" & vbNewLine
            QRY = QRY & "FROM TBL_LOTMASTER A" & vbNewLine
            'End If

            QRY = QRY & "WHERE SITE_ID = '" & Site_id & "'" & vbNewLine

            'If ComboBoxEx1.Text <> "ALL" Then
            '    QRY = QRY & "  AND MODEL = '" &  & "'" & vbNewLine
            'End If
            If m_qry <> "" Then
                QRY = QRY & "AND MODEL IN (" & m_qry & ")" & vbNewLine
            End If

            If .ColumnHeader.Columns(e.Column).Label <> "TOTAL" Then
                QRY = QRY & "  AND C_PRC = (SELECT CODE_ID FROM TBL_CODEMASTER WHERE SITE_ID = A.SITE_ID AND CLASS_ID = '10024' AND CODE_NAME = '" & .ColumnHeader.Columns(e.Column).Label & "')" & vbNewLine
            End If

            If .GetValue(e.Row, 0) <> "TOTAL" Then
                QRY = QRY & "  AND MODEL = '" & .GetValue(e.Row, 0) & "'" & vbNewLine
            End If
        End With

        QRY = QRY & "ORDER BY LOT_NO" & vbNewLine

        If Query_Spread(FpSpread2, QRY, 1) = True Then
            Spread_AutoCol(FpSpread2)
            FpSpread2.ActiveSheet.Columns(0, FpSpread2.ActiveSheet.ColumnCount - 1).Locked = False
        End If


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