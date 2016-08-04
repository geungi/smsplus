Public Class FrmFIns

    Private Sub FrmFIns_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.GotFocus
        Me.ComboBoxEx2.Focus()
    End Sub

    Private Sub FrmFIns_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        DockContainerItem1.Text = "실행 메뉴"
        DockContainerItem2.Text = "입력 조건"

        'TYPE 
        If Query_Combo(Me.ComboBoxEx2, "SELECT '[' + CODE_ID + '] '+CODE_NAME FROM tbl_CODEmaster WHERE site_id = '" & Site_id & "' and CLASS_ID = 'R0001' AND active = 'Y' ORDER BY CODE_ID") = True Then
            ComboBoxEx2.Text = ComboBoxEx2.Items(0)
        End If
        ComboBoxEx2.Enabled = False

        'MODEL 
        If Query_Combo(Me.ComboBoxEx1, "SELECT model_no  FROM tbl_modelmaster WHERE site_id = '" & Site_id & "' and active = 'Y' ORDER BY model_no") = True Then
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
            System.Windows.Forms.MessageBox.Show("창고를 선택하세요!!")
            'Me.TextBoxX1.Text = ""
            Exit Sub
        End If

        With FpSpread1.ActiveSheet
            .RowCount = IntegerInput1.Text
            .Rows(0, .RowCount - 1).ForeColor = Color.OrangeRed

            Dim NO As String = ""
            Dim dt As String = ""

            dt = Query_RS("SELECT CONVERT(VARCHAR(8), DateTimeInput1.Text,112)")
            NO = Query_RS("SELECT MAX(PROD_NO) FROM view_PRODMASTER WHERE PROD_NO LIKE '" & ComboBoxEx1.Text & "%'")

            Dim SN As Integer
            If NO = "" Then
                SN = 0
            Else
                SN = CInt(Mid(NO, 10, 9))
            End If

            For I As Integer = 0 To .RowCount - 1
                SN = SN + 1
                .SetValue(I, 0, ComboBoxEx1.Text & dt & Microsoft.VisualBasic.Right("00000000" & CStr(SN), 9))
                .SetValue(I, 1, ComboBoxEx1.Text)
                .SetValue(I, 2, TextBoxX1.Text)
                .SetValue(I, 3, ComboBoxEx2.Text)
                .SetValue(I, 4, DateTimeInput1.Text)
            Next

            Spread_AutoCol(FpSpread1)

        End With

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



        With FpSpread1.ActiveSheet

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
                    If Insert_Data("insert into tbl_PRODmaster values ('" & Site_id & "','" & .GetValue(I, 0) & "','" & .GetValue(I, 1) & "',NULL, NULL, '" & .GetValue(I, 4) & "',NULL,NULL, '" & Mid(.GetValue(I, 3), 2, 5) & "','" & Emp_No & "', GETDATE(), '" & Emp_No & "', GETDATE(),NULL,NULL,NULL,NULL,NULL)") = True Then
                        .Rows(I).ForeColor = Color.Black
                        Print_SMSBarcode(.GetValue(I, 0), RichTextBox1)
                        '                        Print_UPCBarcode(.GetValue(I, 1), "AMAZON", RichTextBox1)
                        ''Threading.Thread.Sleep(100)
                        ''Kill("C:\TEMP\ESN.PRN")
                    End If
                End If
            Next

        End With

        MessageBox.Show("입고 완료되었습니다.")

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
        TextBoxX1.Text = ""
        TextBoxX1.Text = Query_RS("SELECT MODEL_NAME FROM TBL_MODELMASTER WHERE SITE_ID = '" & Site_id & "' AND MODEL_NO = '" & ComboBoxEx1.Text & "'")
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