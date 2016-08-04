Public Class FrmFIns1

    Private Sub FrmFIns_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.GotFocus
        Me.ComboBoxEx2.Focus()
    End Sub

    Private Sub FrmFIns_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        DockContainerItem1.Text = "실행 메뉴"
        DockContainerItem2.Text = "입력 조건"

        'TYPE 
        If Query_Combo(Me.ComboBoxEx2, "SELECT '[' + CODE_ID + '] '+CODE_NAME FROM tbl_CODEmaster WHERE site_id = '" & Site_id & "' and CLASS_ID = 'R0043' AND active = 'Y' ORDER BY CODE_ID") = True Then
            ComboBoxEx2.Text = ComboBoxEx2.Items(0)
        End If
        'ComboBoxEx2.Enabled = False

        'MODEL 
        If Query_Combo(Me.ComboBoxEx1, "SELECT model_no FROM tbl_modelmaster WHERE site_id = '" & Site_id & "' and active = 'Y' ORDER BY model_no") = True Then
        End If


        If Spread_Setting(FpSpread1, "FrmFIns") = True Then
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

        DateTimeInput1.Value = Now

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

        With FpSpread1.ActiveSheet
            If ComboBoxEx2.Text = "[POL] 폴리싱입고" Then
                .RowCount = 1
            Else
                .RowCount = IntegerInput1.Text
            End If
            .Rows(0, .RowCount - 1).ForeColor = Color.OrangeRed

            Dim NO As String = ""
            Dim dt As String = ""

            dt = DateTimeInput1.Text  ' Query_RS ("SELECT CONVERT(VARCHAR(8), GETDATE(),112)")
            NO = Query_RS("SELECT MAX(SUBSTRING(OUT_ESN,2,14)) FROM view_FESNMASTER WHERE out_ESN LIKE '%" & dt & "%'")

            Dim SN As Integer
            If NO = "" Then
                SN = 0
            Else
                SN = CInt(Mid(NO, 9, 6))
            End If

            For I As Integer = 0 To .RowCount - 1
                .SetValue(I, 0, ComboBoxEx1.Text)
                SN = SN + 1
                If ComboBoxEx2.Text = "[BRK] 브로큰입고" Then
                    .SetValue(I, 1, "B" & dt & Microsoft.VisualBasic.Right("00000" & CStr(SN), 6) & "-0")
                ElseIf ComboBoxEx2.Text = "[POL] 폴리싱입고" Then
                    .SetValue(I, 1, "P" & dt & Microsoft.VisualBasic.Right("00000" & CStr(SN), 6) & "-0")
                ElseIf ComboBoxEx2.Text = "[PRD] 정상입고" Then
                    .SetValue(I, 1, "U" & dt & Microsoft.VisualBasic.Right("00000" & CStr(SN), 6) & "-0")
                ElseIf ComboBoxEx2.Text = "[QAR] RMA입고" Then
                    .SetValue(I, 1, "Q" & dt & Microsoft.VisualBasic.Right("00000" & CStr(SN), 6) & "-0")
                ElseIf ComboBoxEx2.Text = "[CHN] 구매입고" Then
                    .SetValue(I, 1, "C" & dt & Microsoft.VisualBasic.Right("00000" & CStr(SN), 6) & "-0")
                End If
            Next

            Spread_AutoCol(FpSpread1)

        End With

    End Sub

    Private Sub SaveBtn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SaveBtn.Click

        Dim PBOX As String
        Dim QRY As String

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


        With FpSpread1.ActiveSheet
            MainFrm.ProgressBarItem1.Maximum = .RowCount

            If .RowCount = 0 Then
                System.Console.Beep(3000, 400)
                System.Console.Beep(3000, 400)
                System.Windows.Forms.MessageBox.Show("입고시킬 수량을 입력하세요!!")
                'Me.TextBoxX1.Text = ""
                Exit Sub
            End If

            Dim dt As String = DateTimeInput1.Text

            For I As Integer = 0 To .RowCount - 1
                If .Rows(I).ForeColor = Color.OrangeRed Then
                    If ComboBoxEx2.Text = "[QAR] RMA입고" Then

                        If I = .RowCount - 1 Then
                            PBOX = Query_RS("SELECT 'R'+ CONVERT(VARCHAR(8),GETDATE(),112) +'-'+ISNULL(RIGHT('000'+CONVERT(VARCHAR(4),CONVERT(INT,RIGHT(MAX(LOT_NO),4))+1),4),'0001') FROM TBL_LOTMASTER WHERE LOT_NO LIKE 'R'+ CONVERT(VARCHAR(8),GETDATE(),112)+'%'")

                            QRY = "EXEC SP_COMMON_LOTCREATE '" & Site_id & "','" & PBOX & "','NT01','" & ComboBoxEx1.Text & "','RMA LOT 생성','" & Emp_No & "','입고대기','PREINSPECTION','U2014-0001','" & dt & "'" & vbNewLine & vbNewLine
                            QRY = QRY & "UPDATE TBL_LOTMASTER SET INIT_QTY = " & IntegerInput1.Text & ", ACT_QTY = " & IntegerInput1.Text & " WHERE LOT_NO = '" & PBOX & "'" & vbNewLine


                            If Insert_Data(QRY) = True Then
                                .Rows(I).ForeColor = Color.Black
                                Print_LOTBarcode(PBOX, ComboBoxEx1.Text, Query_RS("SELECT INIT_QTY FROM TBL_LOTMASTER WHERE LOT_NO = '" & PBOX & "'"), RichTextBox1, "NT01")
                            End If

                        End If


                        If Insert_Data("insert into tbl_fesnmaster_K values (1, '" & Site_id & "','" & dt & .GetValue(I, 1) & "','" & .GetValue(I, 1) & "','" & .GetValue(I, 0) & "','Y',NULL,NULL,NULL,'W0000',1,NULL,NULL,NULL,'Y',GETDATE(),NULL,'" & Mid(ComboBoxEx2.Text, 2, 3) & dt & "-01'" & ", NULL, 'K1000','" & Emp_No & "',GETDATE(),'" & Emp_No & "',GETDATE(),NULL,NULL,NULL,NULL,NULL)") = True Then
                            .Rows(I).ForeColor = Color.Black
                        End If

                    ElseIf ComboBoxEx2.Text = "[POL] 폴리싱입고" Then
                        PBOX = Query_RS("SELECT 'P'+ CONVERT(VARCHAR(8),GETDATE(),112) +'-'+ISNULL(RIGHT('000'+CONVERT(VARCHAR(4),CONVERT(INT,RIGHT(MAX(LOT_NO),4))+1),4),'0001') FROM TBL_LOTMASTER WHERE LOT_NO LIKE 'P'+ CONVERT(VARCHAR(8),GETDATE(),112)+'%'")

                        QRY = "EXEC SP_COMMON_LOTCREATE '" & Site_id & "','" & PBOX & "','NT01','" & ComboBoxEx1.Text & "','생산 LOT 생성','" & Emp_No & "','폴리싱','LCD FUNCTION TEST(1)','P2014-0003','" & dt & "'" & vbNewLine & vbNewLine
                        QRY = QRY & "UPDATE TBL_LOTMASTER SET INIT_QTY = " & IntegerInput1.Text & ", ACT_QTY = " & IntegerInput1.Text & " WHERE LOT_NO = '" & PBOX & "'" & vbNewLine


                        If Insert_Data(QRY) = True Then
                            .Rows(I).ForeColor = Color.Black
                            Print_LOTBarcode(PBOX, ComboBoxEx1.Text, Query_RS("SELECT INIT_QTY FROM TBL_LOTMASTER WHERE LOT_NO = '" & PBOX & "'"), RichTextBox1, "NT01")
                        End If
                        Exit For
                    ElseIf ComboBoxEx2.Text = "[PRD] 정상입고" Then

                        If I = .RowCount - 1 Then
                            PBOX = Query_RS("SELECT 'U'+ CONVERT(VARCHAR(8),GETDATE(),112) +'-'+ISNULL(RIGHT('000'+CONVERT(VARCHAR(4),CONVERT(INT,RIGHT(MAX(LOT_NO),4))+1),4),'0001') FROM TBL_LOTMASTER WHERE LOT_NO LIKE 'U'+ CONVERT(VARCHAR(8),GETDATE(),112)+'%'")

                            QRY = "EXEC SP_COMMON_LOTCREATE '" & Site_id & "','" & PBOX & "','NT01','" & ComboBoxEx1.Text & "','정상 입고 LOT 생성','" & Emp_No & "','입고대기','PREINSPECTION','U2014-0001','" & dt & "'" & vbNewLine & vbNewLine
                            QRY = QRY & "UPDATE TBL_LOTMASTER SET INIT_QTY = " & IntegerInput1.Text & ", ACT_QTY = " & IntegerInput1.Text & " WHERE LOT_NO = '" & PBOX & "'" & vbNewLine


                            If Insert_Data(QRY) = True Then
                                .Rows(I).ForeColor = Color.Black
                                Print_LOTBarcode(PBOX, ComboBoxEx1.Text, Query_RS("SELECT INIT_QTY FROM TBL_LOTMASTER WHERE LOT_NO = '" & PBOX & "'"), RichTextBox1, "NT01")
                            End If

                        End If

                        If Insert_Data("insert into tbl_fesnmaster_K values (1, '" & Site_id & "','" & dt & .GetValue(I, 1) & "','" & .GetValue(I, 1) & "','" & .GetValue(I, 0) & "','N',NULL,NULL,NULL,'W0000',1,NULL,NULL,NULL,'Y',GETDATE(),NULL,'" & Mid(ComboBoxEx2.Text, 2, 3) & dt & "-01'" & ", NULL, 'K5000','" & Emp_No & "',GETDATE(),'" & Emp_No & "',GETDATE(),NULL,NULL,NULL,NULL,NULL)") = True Then
                            .Rows(I).ForeColor = Color.Black
                        End If
                    ElseIf ComboBoxEx2.Text = "[BRK] 브로큰입고" Then
                        If I = .RowCount - 1 Then
                            PBOX = Query_RS("SELECT 'B'+ CONVERT(VARCHAR(8),GETDATE(),112) +'-'+ISNULL(RIGHT('000'+CONVERT(VARCHAR(4),CONVERT(INT,RIGHT(MAX(LOT_NO),4))+1),4),'0001') FROM TBL_LOTMASTER WHERE LOT_NO LIKE 'B'+ CONVERT(VARCHAR(8),GETDATE(),112)+'%'")

                            QRY = "EXEC SP_COMMON_LOTCREATE '" & Site_id & "','" & PBOX & "','NT01','" & ComboBoxEx1.Text & "','BRK 입고 LOT 생성','" & Emp_No & "','입고대기','PREINSPECTION','U2014-0001','" & dt & "'" & vbNewLine & vbNewLine
                            QRY = QRY & "UPDATE TBL_LOTMASTER SET INIT_QTY = " & IntegerInput1.Text & ", ACT_QTY = " & IntegerInput1.Text & " WHERE LOT_NO = '" & PBOX & "'" & vbNewLine


                            If Insert_Data(QRY) = True Then
                                .Rows(I).ForeColor = Color.Black
                                Print_LOTBarcode(PBOX, ComboBoxEx1.Text, Query_RS("SELECT INIT_QTY FROM TBL_LOTMASTER WHERE LOT_NO = '" & PBOX & "'"), RichTextBox1, "NT01")
                            End If

                        End If

                        If Insert_Data("insert into tbl_fesnmaster_K values (1, '" & Site_id & "','" & dt & .GetValue(I, 1) & "','" & .GetValue(I, 1) & "','" & .GetValue(I, 0) & "','N',NULL,NULL,NULL,'W0000',1,NULL,NULL,NULL,'Y',GETDATE(),NULL,'" & Mid(ComboBoxEx2.Text, 2, 3) & dt & "-01'" & ", NULL, 'K5000','" & Emp_No & "',GETDATE(),'" & Emp_No & "',GETDATE(),NULL,NULL,NULL,NULL,NULL)") = True Then
                            .Rows(I).ForeColor = Color.Black
                        End If

                    Else

                        If I = .RowCount - 1 Then
                            PBOX = Query_RS("SELECT 'P'+ CONVERT(VARCHAR(8),GETDATE(),112) +'-'+ISNULL(RIGHT('000'+CONVERT(VARCHAR(4),CONVERT(INT,RIGHT(MAX(LOT_NO),4))+1),4),'0001') FROM TBL_LOTMASTER WHERE LOT_NO LIKE 'P'+ CONVERT(VARCHAR(8),GETDATE(),112)+'%'")

                            QRY = "EXEC SP_COMMON_LOTCREATE '" & Site_id & "','" & PBOX & "','NT01','" & ComboBoxEx1.Text & "','입고 LOT 생성','" & Emp_No & "','입고대기','PREINSPECTION','U2014-0001','" & dt & "'" & vbNewLine & vbNewLine
                            QRY = QRY & "UPDATE TBL_LOTMASTER SET INIT_QTY = " & IntegerInput1.Text & ", ACT_QTY = " & IntegerInput1.Text & " WHERE LOT_NO = '" & PBOX & "'" & vbNewLine


                            If Insert_Data(QRY) = True Then
                                .Rows(I).ForeColor = Color.Black
                                Print_LOTBarcode(PBOX, ComboBoxEx1.Text, Query_RS("SELECT INIT_QTY FROM TBL_LOTMASTER WHERE LOT_NO = '" & PBOX & "'"), RichTextBox1, "NT01")
                            End If

                        End If

                        If Insert_Data("insert into tbl_fesnmaster_K values (1, '" & Site_id & "','" & dt & .GetValue(I, 1) & "','" & .GetValue(I, 1) & "','" & .GetValue(I, 0) & "','N',NULL,NULL,NULL,'W0000',1,NULL,NULL,NULL,'Y',GETDATE(),NULL,'" & Mid(ComboBoxEx2.Text, 2, 3) & dt & "-01'" & ", NULL, 'K5000','" & Emp_No & "',GETDATE(),'" & Emp_No & "',GETDATE(),NULL,NULL,NULL,NULL,NULL)") = True Then
                            .Rows(I).ForeColor = Color.Black
                        End If
                    End If

                End If
                MainFrm.ProgressBarItem1.Value = I

            Next

            System.Windows.Forms.MessageBox.Show("입고 완료되었습니다!!")

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