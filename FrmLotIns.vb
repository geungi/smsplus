Public Class FrmLotIns

    Private Sub FrmFIns_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.GotFocus
        Me.ComboBoxEx2.Focus()
    End Sub

    Private Sub FrmFIns_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        DockContainerItem1.Text = "실행 메뉴"
        DockContainerItem2.Text = "입력 조건"
        DockContainerItem3.Text = "라인 자재 재고 현황"

        DateTimeInput1.Value = Now

        'TYPE 
        If Query_Combo(Me.ComboBoxEx2, "SELECT '[' + CODE_ID + '] '+CODE_NAME FROM tbl_CODEmaster WHERE site_id = '" & Site_id & "' and CLASS_ID = 'R0043' AND active = 'Y' ORDER BY CODE_ID") = True Then
            ComboBoxEx2.Text = ComboBoxEx2.Items(0)
        End If
        'ComboBoxEx2.Enabled = False
        '공정 
        If Query_Combo(Me.ComboBoxEx3, "SELECT CODE_NAME FROM tbl_CODEmaster WHERE site_id = '" & Site_id & "' and CLASS_ID = 'R0001' AND active = 'Y' ORDER BY CODE_ID") = True Then
            ComboBoxEx3.Text = ComboBoxEx3.Items(0)
        End If


        'MODEL 
        If Query_Combo(Me.ComboBoxEx1, "SELECT model_no FROM tbl_modelmaster WHERE site_id = '" & Site_id & "' and active = 'Y' ORDER BY model_no") = True Then
        End If

        '거래처 
        If Query_Combo(Me.ComboBoxEx4, "SELECT '['+sup_no+'] ' + sup_nm  FROM tbl_supmst WHERE site_id = '" & Site_id & "' and sup_no like 'P%' ORDER BY sup_no") = True Then
        End If

        If Spread_Setting(FpSpread1, Me.Name) = True Then
            FpSpread1.ActiveSheet.RowCount = 0
            Spread_AutoCol(FpSpread1)
        End If

        Formbim_Authority(Me.ButtonItem2, Me.Name, "NEW")
        Formbtn_Authority(Me.NewBtn1, Me.Name, "NEW")
        Formbim_Authority(Me.ButtonItem3, Me.Name, "SAVE")
        Formbim_Authority(Me.SaveBtn, Me.Name, "SAVE")
        Formbim_Authority(Me.ButtonItem4, Me.Name, "DELETE")
        Formbtn_Authority(Me.DelBtn1, Me.Name, "DELETE")
        Formbtn_Authority(Me.PrtBtn1, Me.Name, "PRINT")
        Formbtn_Authority(Me.XlsBtn1, Me.Name, "EXCEL")



    End Sub

    Private Sub ButtonItem4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonItem4.Click, DelBtn1.Click

        If ComboBoxEx1.Text = "" Then
            System.Console.Beep(3000, 400)
            System.Console.Beep(3000, 400)
            System.Windows.Forms.MessageBox.Show("모델을 입력하세요!!")
            'Me.TextBoxX1.Text = ""
            Exit Sub
        End If

        If ComboBoxEx2.Text = "" Then
            System.Console.Beep(3000, 400)
            System.Console.Beep(3000, 400)
            System.Windows.Forms.MessageBox.Show("입력 유형을 선택하세요!!")
            'Me.TextBoxX1.Text = ""
            Exit Sub
        End If

        If ComboBoxEx4.Text = "" Then
            System.Console.Beep(3000, 400)
            System.Console.Beep(3000, 400)
            System.Windows.Forms.MessageBox.Show("입고처를 선택하세요!!")
            'Me.TextBoxX1.Text = ""
            Exit Sub
        End If



        With FpSpread1.ActiveSheet
            Dim QRY = "SELECT TOP " & IntegerInput1.Text & " MODEL, OUT_ESN, PSHIP_NO, RCV_DATE, (SELECT CODE_NAME FROM TBL_CODEMASTER WHERE CLASS_ID = 'R0001' AND CODE_ID = A.C_PRC), PBOX_NO FROM TBL_FESNMASTER_K A" & vbNewLine
            If ComboBoxEx2.Text = "[BRK] 브로큰입고" Then
                QRY = QRY & "where OUT_ESN LIKE 'B%'" & vbNewLine
            ElseIf ComboBoxEx2.Text = "[POL] 폴리싱입고" Then
                QRY = QRY & "where OUT_ESN LIKE 'S%'" & vbNewLine
            ElseIf ComboBoxEx2.Text = "[QAR] RMA입고" Then
                QRY = QRY & "where OUT_ESN LIKE 'Q%'" & vbNewLine
            ElseIf ComboBoxEx2.Text = "[CHN] 구매입고" Then
                QRY = QRY & "where OUT_ESN LIKE 'C%'" & vbNewLine
            ElseIf ComboBoxEx2.Text = "[PRD] 정상입고" Then
                QRY = QRY & "where SUBSTRING(OUT_ESN,1,1) NOT IN ('B','S','Q','C')" & vbNewLine
            End If
            QRY = QRY & "AND C_PRC = 'K1000'" & vbNewLine
            QRY = QRY & "AND CHK_RCV = 'Y'" & vbNewLine
            QRY = QRY & "AND PBOX_NO IS NULL" & vbNewLine
            QRY = QRY & "AND IN_ESN IS NULL" & vbNewLine
            QRY = QRY & "AND MODEL = '" & ComboBoxEx1.Text & "'" & vbNewLine
            If ComboBoxEx2.Text = "[PRD] 정상입고" Then
                QRY = QRY & "ORDER BY C_DATE" & vbNewLine
            Else
                QRY = QRY & "ORDER BY OUT_ESN" & vbNewLine
            End If

            If Query_Spread(FpSpread1, QRY, 1) = True Then
                .Rows(0, .RowCount - 1).ForeColor = Color.OrangeRed
                Spread_AutoCol(FpSpread1)
                .SetActiveCell(.RowCount - 1, 0)
                Me.FpSpread1.ShowActiveCell(FarPoint.Win.Spread.VerticalPosition.Bottom, FarPoint.Win.Spread.HorizontalPosition.Left)
            End If


            If Query_Listview(ListViewEX1, "EXEC SP_COMMON_GETPARTINV '" & Site_id & "','" & ComboBoxEx1.Text & "','" & "C1000" & "',''", True) = True Then
                ListViewEX1.AutoResizeColumn(0, ColumnHeaderAutoResizeStyle.ColumnContent)
            End If


        End With

    End Sub

    Private Sub ListViewEx1_ItemChecked(ByVal sender As Object, ByVal e As System.Windows.Forms.ItemCheckedEventArgs) Handles ListViewEX1.ItemChecked



        If e.Item.Checked = True Then
            If IntegerInput1.Text = "0" Then
                Modal_Error("생산LOT 수량을 입력하십시오.")
                Exit Sub
            End If

            If e.Item.SubItems(3).Text = 0 Then
                System.Console.Beep(3000, 400)
                System.Console.Beep(3000, 400)
                Modal_Error("재고수량이 없습니다!!")
                e.Item.Checked = False
                Exit Sub
            End If
        Else

        End If

    End Sub


    Private Sub SaveBtn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SaveBtn.Click

        If ComboBoxEx1.Text = "" Then
            System.Console.Beep(3000, 400)
            System.Console.Beep(3000, 400)
            System.Windows.Forms.MessageBox.Show("모델을 입력하세요!!")
            'Me.TextBoxX1.Text = ""
            Exit Sub
        End If

        If ComboBoxEx2.Text = "" Then
            System.Console.Beep(3000, 400)
            System.Console.Beep(3000, 400)
            System.Windows.Forms.MessageBox.Show("입력 유형을 선택하세요!!")
            'Me.TextBoxX1.Text = ""
            Exit Sub
        End If

        If ComboBoxEx4.Text = "" Then
            System.Console.Beep(3000, 400)
            System.Console.Beep(3000, 400)
            System.Windows.Forms.MessageBox.Show("입고처를 선택하세요!!")
            'Me.TextBoxX1.Text = ""
            Exit Sub
        End If

        With FpSpread1.ActiveSheet
            MainFrm.ProgressBarItem1.Maximum = .RowCount

            If .RowCount = 0 Then
                System.Console.Beep(3000, 400)
                System.Console.Beep(3000, 400)
                System.Windows.Forms.MessageBox.Show("입고시킬 수량을 입력하세요!!")
                'Me.TextBoxX1.Text = ""
                Exit Sub
            End If

            Dim NO As String = ""
            Dim dt As String = ""

            dt = Query_RS("SELECT CONVERT(VARCHAR(8), GETDATE(),112)")
            NO = Query_RS("SELECT MAX(LOT_NO) FROM tbl_LOTMASTER WHERE LOT_NO LIKE '%" & dt & "%' AND LEFT(LOT_NO,1) = '" & Mid(ComboBoxEx2.Text, 2, 1) & "'")

            Dim SN As Integer
            If NO = "" Then
                SN = 0
            Else
                SN = CInt(Microsoft.VisualBasic.Right(NO, 4))
            End If

            SN = SN + 1
            Dim PBOX As String = Mid(ComboBoxEx2.Text, 2, 1) & dt & "-" & Microsoft.VisualBasic.Right("0000" & CStr(SN), 4)
            Dim QRY As String = "EXEC SP_COMMON_LOTCREATE '" & Site_id & "','" & PBOX & "','NT01','" & ComboBoxEx1.Text & "','" & ComboBoxEx2.Text & " LOT 생성','" & Emp_No & "','PREINSPECTION','" & ComboBoxEx3.Text & "','" & Mid(ComboBoxEx4.Text, 2, 10) & "','" & DateTimeInput1.Text & "'" & vbNewLine

            'If Mid(ComboBoxEx4.Text, 2, 10) <> "P2014-0000" Then
            '    Dim QRY1 As String = "EXEC SP_FrmBUYING_SAVE '" & Site_id & "','',"
            '    QRY1 = QRY1 & Mid(ComboBoxEx4.Text, 14, Len(ComboBoxEx4.Text) - 13) & "','" & DateTimeInput1.Text & "','"
            '    If ComboBoxEx2.Text = "" Then
            '    ElseIf ComboBoxEx2.Text = "" Then
            '    ElseIf ComboBoxEx2.Text = "" Then
            '    End If
            '    QRY1 = QRY1 & "" & "','"
            '    QRY1 = QRY1 & "" & "','"
            '    QRY1 = QRY1 & "" & "','"
            '    QRY1 = QRY1 & "" & "','"
            '    QRY1 = QRY1 & "" & "','"
            '    QRY1 = QRY1 & "" & "','"

            '    If Insert_Data(QRY1) Then

            '        .Rows(i).ForeColor = Color.Black
            '    End If
            'End If



            QRY = QRY & "UPDATE TBL_LOTMASTER SET INIT_QTY = " & .RowCount & ", ACT_QTY = " & .RowCount & " WHERE LOT_NO = '" & PBOX & "'" & vbNewLine

            For i As Integer = 0 To ListViewEX1.Items.Count - 1
                If ListViewEX1.Items(i).Checked = True Then
                    QRY = QRY & "EXEC SP_COMMON_LOTSAVE '" & Site_id & "','" & PBOX & "','','','" & ListViewEX1.Items(i).Text & "'," & .RowCount & ",'LOT 생성 파트입력','" & Emp_No & "','PREINSPECTION','" & ComboBoxEx3.Text & "','" & OP_No & "','C1000', 0" & vbNewLine
                End If
            Next

            If Insert_Data(QRY) = True Then
                Print_LOTBarcode(PBOX, ComboBoxEx1.Text, IntegerInput1.Text, RichTextBox1, "NT01")

                For I As Integer = 0 To .RowCount - 1
                    If .Rows(I).ForeColor = Color.OrangeRed Then
                        If Insert_Data("UPDATE tbl_fesnmaster_K SET PBOX_NO = '" & PBOX & "', C_PRC = '" & Query_RS("SELECT CODE_ID FROM TBL_CODEMASTER WHERE CLASS_ID = 'R0001' AND CODE_NAME = '" & ComboBoxEx3.Text & "'") & "',U_PERSON = '" & Emp_No & "',U_DATE =  GETDATE() WHERE OUT_ESN = '" & .GetValue(I, 1) & "'") = True Then
                            .Rows(I).ForeColor = Color.Black
                        End If
                    End If
                    MainFrm.ProgressBarItem1.Value = I
                Next
            End If


            System.Windows.Forms.MessageBox.Show("생성 완료되었습니다!!")

        End With

    End Sub

    Private Sub ButtonItem13_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PrtBtn1.Click, ButtonItem13.Click
        If save_excel = "" Then
            MessageBox.Show("Select Spread for Printing!!")
            Exit Sub
        End If

        If save_excel = "FpSpread1" Then
            If Spread_Print(Me.FpSpread1, "NO SHIP(Board) Summary", 0) = False Then
                MsgBox("Fail to Print")
            End If

        Else
            MessageBox.Show("Select Spread for Printing!!")
        End If
    End Sub

    Private Sub ButtonItem14_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles XlsBtn1.Click, ButtonItem14.Click

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

    Private Sub ComboBoxEx2_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBoxEx2.SelectedIndexChanged

        IntegerInput1.Text = 0
        FpSpread1.ActiveSheet.RowCount = 0
        ComboBoxEx1.Focus()

    End Sub

    Private Sub FrmFIns_Shown(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Shown
        Me.ComboBoxEx2.Focus()
    End Sub

    Private Sub ComboBoxEx1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBoxEx1.SelectedIndexChanged

        FpSpread1.ActiveSheet.RowCount = 0

        If Query_Listview(ListViewEX1, "EXEC SP_COMMON_GETPARTINV '" & Site_id & "','" & ComboBoxEx1.Text & "','" & "C1000" & "',''", True) = True Then
            ListViewEX1.AutoResizeColumn(0, ColumnHeaderAutoResizeStyle.ColumnContent)
        End If

        IntegerInput1.Text = 0
        IntegerInput1.Focus()

    End Sub

    Private Sub IntegerInput1_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles IntegerInput1.TextChanged

        If ComboBoxEx1.Text = "" Then
            System.Console.Beep(3000, 400)
            System.Console.Beep(3000, 400)
            System.Windows.Forms.MessageBox.Show("모델을 입력하세요!!")
            'Me.TextBoxX1.Text = ""
            Exit Sub
        End If

        If ComboBoxEx2.Text = "" Then
            System.Console.Beep(3000, 400)
            System.Console.Beep(3000, 400)
            System.Windows.Forms.MessageBox.Show("입고유형을 선택하세요!!")
            'Me.TextBoxX1.Text = ""
            Exit Sub
        End If




    End Sub


End Class