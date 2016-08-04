Public Class FrmFinalPass

    Private Sub FrmLinePass_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.GotFocus
        Me.TextBoxX4.Focus()
        TextBoxX4.Text = "K9000"

        TextBoxX5.Text = Query_RS("SELECT code_name FROM TBL_codeMASTER WHERE class_id = 'R0001' and code_id = '" & TextBoxX4.Text & "'")
        Me.TextBoxX1.Focus()
    End Sub

    Private Sub FrmLinePass_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        DockContainerItem1.Text = "Form Menu"
        DockContainerItem2.Text = "Input Condition"

        If Spread_Setting(FpSpread1, "FrmLinePass") = True Then
            FpSpread1.ActiveSheet.RowCount = 0
            Spread_AutoCol(FpSpread1)
        End If

        TextBoxX4.Text = "K9000"
        TextBoxX4.Enabled = False
        TextBoxX5.Text = Query_RS("SELECT code_name FROM TBL_codeMASTER WHERE class_id = 'R0001' and code_id = '" & TextBoxX4.Text & "'")
        Me.TextBoxX1.Focus()


        If Query_Combo(ComboBoxEx7, "sELECT '[' + EMP_NO + ']' + EMP_NM FROM TBL_EMPMASTER WHERE dept_cd = (select dept_cd from tbl_empmaster where emp_no = '" & Emp_No & "') and RETIRE_YN = 'N' ORDER BY EMP_NO") = True Then
            ' MessageBox.Show("site_id : " & ComboBoxItem1.ComboBoxEx.Text)
        Else
            MessageBox.Show("error")

        End If
        ComboBoxEx7.Text = MainFrm.ComboBoxItem1.ComboBoxEx.Text

        Formbim_Authority(Me.ButtonItem2, Me.Name, "NEW")
        Formbtn_Authority(Me.NewBtn1, Me.Name, "NEW")
        Formbim_Authority(Me.ButtonItem3, Me.Name, "SAVE")
        Formbtn_Authority(Me.SaveBtn1, Me.Name, "SAVE")
        Formbim_Authority(Me.ButtonItem4, Me.Name, "DELETE")
        Formbtn_Authority(Me.DelBtn1, Me.Name, "DELETE")
        'Formbim_Authority(Me.ButtonItem5, Me.Name, "PRINT")
        Formbtn_Authority(Me.PrtBtn1, Me.Name, "PRINT")
        Formbtn_Authority(Me.XlsBtn1, Me.Name, "EXCEL")
        '        Formbtn_Authority(Me.FindBtn, Me.Name, "FIND")

        ButtonItem2.Enabled = False
        NewBtn1.Enabled = False
        SaveBtn1.Enabled = False
        ButtonItem4.Enabled = False
        DelBtn1.Enabled = False

    End Sub

    Private Sub ComboBox7_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBoxEx7.SelectedIndexChanged
        MainFrm.ComboBoxItem1.ComboBoxEx.Text = ComboBoxEx7.Text
        OP_No = Mid(ComboBoxEx7.Text, 2, 5)
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

    Private Sub TextBoxX1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBoxX1.Click
        Me.TextBoxX1.Text = ""
    End Sub


    Private Sub TextBoxX1_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TextBoxX1.KeyDown

        TextBoxX4.Text = "K9000"
        TextBoxX5.Text = Query_RS("SELECT code_name FROM TBL_codeMASTER WHERE class_id = 'R0001' and code_id = '" & TextBoxX4.Text & "'")
        
        If TextBoxX4.Text = "" Then
            System.Console.Beep(3000, 400)
            System.Console.Beep(3000, 400)
            System.Windows.Forms.MessageBox.Show("다음 공정을 입력하십시오!!")
            'Me.TextBoxX1.Text = ""
            Exit Sub
        End If

        If e.KeyCode <> Keys.Enter Then
            Exit Sub
        End If

        If LOT_VERIFY(Me.TextBoxX1, e) = True Then

            If Check_Valid_LOT(TextBoxX1.Text, Me.Name) = False Then
                'Me.TextBoxX1.Text = ""
                Me.TextBoxX1.Focus()
                Me.TextBoxX1.SelectAll()
                Exit Sub
            End If

            If Insert_Data("EXEC SP_COMMON_LOTSAVE '" & Site_id & "','" & TextBoxX1.Text & "','','','N/A',0,'FINAL QC 합격','" & Emp_No & "','FINAL QC','" & TextBoxX5.Text & "','" & OP_No & "','C1000', 0") = True Then
                'PASS시 개별 플립ID 부여

                Dim CNT As Integer = 0
                CNT = Query_RS("SELECT ACT_QTY FROM TBL_LOTMASTER WHERE LOT_NO = '" & TextBoxX1.Text & "'")
                Dim MODEL As String = Query_RS("SELECT MODEL FROM TBL_LOTMASTER WHERE LOT_NO = '" & TextBoxX1.Text & "'")

                Dim QRY As String = ""

                Dim NO As String = ""
                Dim dt As String = ""

                dt = Query_RS("SELECT CONVERT(VARCHAR(8), GETDATE(),112)")

                For I As Integer = 0 To CNT - 1
                    Dim ESN_RS As New ADODB.Recordset
                    If CInt(Query_RS("SELECT COUNT(OUT_ESN) FROM TBL_FESNMASTER_K WHERE CHK_RCV = 'Y' AND PBOX_NO = '" & TextBoxX1.Text & "' AND IN_ESN IS NULL AND MODEL = '" & MODEL & "' ")) > 0 Then
                        '                        ESN_RS = Query_RS_ALL("SELECT TOP 1 '', OUT_ESN, MODEL, ISNULL((SELECT CODE_NAME FROM TBL_CODEMASTER WHERE CLASS_ID = 'R0001' AND CODE_ID = A.C_PRC),''), ISNULL(T_DEF_CD,'') FROM TBL_FESNMASTER_K A WHERE CHK_RCV = 'Y' AND PBOX_NO = '" & TextBoxX1.Text & "' AND IN_ESN IS NULL AND MODEL = '" & MODEL & "' ORDER BY C_DATE")
                        ESN_RS = Query_RS_ALL("SELECT TOP 1 '', OUT_ESN, MODEL, ISNULL((SELECT CODE_NAME FROM TBL_CODEMASTER WHERE CLASS_ID = 'R0001' AND CODE_ID = A.C_PRC),''), ISNULL(T_DEF_CD,'') FROM TBL_FESNMASTER_K A WHERE CHK_RCV = 'Y' AND IN_ESN IS NULL AND MODEL = '" & MODEL & "' ORDER BY C_DATE")
                    Else
                        '    ESN_RS = Query_RS_ALL("SELECT TOP 1 '', OUT_ESN, MODEL, ISNULL((SELECT CODE_NAME FROM TBL_CODEMASTER WHERE CLASS_ID = 'R0001' AND CODE_ID = A.C_PRC),''), ISNULL(T_DEF_CD,'') FROM TBL_FESNMASTER_K A WHERE CHK_RCV = 'Y' AND PBOX_NO IN (SELECT LOT_NO FROM TBL_LOTMASTER WHERE C_PRC IN ('K9300','K9200')) AND IN_ESN IS NULL AND MODEL = '" & MODEL & "' ORDER BY C_DATE")
                        ESN_RS = Query_RS_ALL("SELECT TOP 1 '', OUT_ESN, MODEL, ISNULL((SELECT CODE_NAME FROM TBL_CODEMASTER WHERE CLASS_ID = 'R0001' AND CODE_ID = A.C_PRC),''), ISNULL(T_DEF_CD,'') FROM TBL_FESNMASTER_K A WHERE CHK_RCV = 'Y' AND IN_ESN IS NULL AND MODEL = '" & MODEL & "' ORDER BY C_DATE")
                    End If

                    NO = Query_RS("SELECT MAX(SUBSTRING(IN_ESN,2,14)) FROM view_FESNMASTER WHERE IN_ESN LIKE '%" & dt & "%'")

                    Dim SN As Integer
                    If NO = "" Then
                        SN = 0
                    Else
                        SN = CInt(Mid(NO, 9, 6))
                    End If

                    SN = SN + 1

                    QRY = "UPDATE TBL_FESNMASTER_K SET " & vbNewLine
                    '                    QRY = QRY & "   IN_ESN = 'K' + '" & Mid(ESN_RS(1).Value, 2, Len(ESN_RS(1).Value) - 1) & "'," & vbNewLine
                    If Mid(TextBoxX1.Text, 1, 1) = "Q" Then
                        QRY = QRY & "   IN_ESN = 'Q" & dt & Microsoft.VisualBasic.Right("00000" & CStr(SN), 6) & "-0" & "'," & vbNewLine
                    Else
                        QRY = QRY & "   IN_ESN = 'K" & dt & Microsoft.VisualBasic.Right("00000" & CStr(SN), 6) & "-0" & "'," & vbNewLine
                    End If
                    QRY = QRY & "   C_PRC = '" & TextBoxX4.Text & "'," & vbNewLine
                    QRY = QRY & "   pbox_no ='" & TextBoxX1.Text & "'," & vbNewLine
                    QRY = QRY & "   U_PERSON = '" & Emp_No & "', U_DATE = GETDATE()" & vbNewLine
                    QRY = QRY & "WHERE OUT_ESN = '" & ESN_RS(1).Value & "'" & vbNewLine
                    QRY = QRY & "" & vbNewLine
                    QRY = QRY & "" & vbNewLine

                    If Insert_Data(QRY) = True Then
                        '개별 바코드 레이블 인쇄
                        '                      Print_NewBarcode(ESN_RS(1).Value, ESN_RS(2).Value, ESN_RS(4).Value, "")
                        Print_ShipBarcode("K" & dt & Microsoft.VisualBasic.Right("00000" & CStr(SN), 6) & "-0", ESN_RS(2).Value, RichTextBox1)
                        System.Threading.Thread.Sleep(700)
                    End If

                    ESN_RS.MoveNext()
                Next

                With FpSpread1.ActiveSheet
                    .AddRows(0, 1)
                    .SetValue(0, 0, MODEL)
                    .SetValue(0, 1, TextBoxX1.Text)
                    .SetValue(0, 2, "FINAL QC")
                    .SetValue(0, 3, Query_RS("select act_qty from tbl_lotmaster where lot_no = '" & TextBoxX1.Text & "'"))
                End With
                Spread_AutoCol(FpSpread1)
            End If

            If Insert_Data("exec SP_COMMON_WKRESULT '" & Site_id & "','FINAL QC','" & OP_No & "','" & FpSpread1.ActiveSheet.Cells(0, 1).Text & "','" & TextBoxX1.Text & "'") = True Then
            End If

            'TextBoxX1.Text = ""
            Me.TextBoxX1.Focus()
            Me.TextBoxX1.SelectAll()
        End If



    End Sub

    Private Sub TextBoxX4_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TextBoxX4.KeyDown
        If e.KeyCode <> Keys.Enter Then
            Exit Sub
        End If

        If TextBoxX4.Text = "0ZZ" Then
            TextBoxX1.Text = ""
            TextBoxX4.Text = ""
            Exit Sub
        Else

            TextBoxX5.Text = Query_RS("SELECT code_name FROM TBL_codeMASTER WHERE class_id = 'R0001' and code_id = '" & TextBoxX4.Text & "'")

            If TextBoxX5.Text = "" Then
                Modal_Error("스캔하신 공정이 존재하지 않습니다.")
                TextBoxX4.Focus()
                TextBoxX4.SelectAll()
                Exit Sub
            End If
        End If

        TextBoxX1.Focus()
        TextBoxX1.SelectAll()
        Exit Sub

    End Sub

    Private Sub FrmWCTrf_Shown(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Shown
        Me.TextBoxX4.Focus()
    End Sub

    Private Sub ComboBoxEx1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Me.TextBoxX1.Focus()
        Me.TextBoxX1.SelectAll()
    End Sub

End Class