Public Class FrmIndPack

    Private Sub FrmWCTrf_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.GotFocus
        Me.ComboBoxEx1.Focus()
    End Sub

    Private Sub FrmWCTrf_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        DockContainerItem1.Text = "Form Menu"
        DockContainerItem2.Text = "Input Condition"

        If Query_Combo(Me.ComboBoxEx1, "SELECT CODE_NAME FROM tbl_CODEmaster WHERE site_id = '" & Site_id & "' and CLASS_ID = 'R0001' AND CODE_ID = 'W1000' AND active = 'Y' ORDER BY CODE_ID") = True Then
        End If
        ComboBoxEx1.Text = ComboBoxEx1.Items(0)
        ComboBoxEx1.Enabled = False

        ComboBoxEx2.Items.Add("AMAZON")
        ComboBoxEx2.Items.Add("GTIN14")
        ComboBoxEx2.Items.Add("GTIN12")
        ComboBoxEx2.Text = ComboBoxEx2.Items(0)


        If Spread_Setting(FpSpread1, Me.Name) = True Then
            FpSpread1.ActiveSheet.RowCount = 0
            Spread_AutoCol(FpSpread1)
        End If

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

    End Sub

    Private Sub ButtonItem4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonItem4.Click, DelBtn1.Click

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
        Dim TRF_RS As New ADODB.Recordset
        Dim QRY As String = ""

        If ComboBoxEx1.Text = "" Then
            System.Console.Beep(3000, 400)
            System.Console.Beep(3000, 400)
            System.Windows.Forms.MessageBox.Show("다음 공정을 선택하세요!!")
            'Me.TextBoxX1.Text = ""
            Exit Sub
        End If

        If e.KeyCode <> Keys.Enter Then
            Exit Sub
        End If

        If PROD_VERIFY(Me.TextBoxX1, e) = True Then

            QRY = "SELECT PROD_NO, MODEL, " & vbNewLine
            QRY = QRY & "(SELECT MODEL_NAME FROM TBL_MODELMASTER WHERE MODEL_NO = A.MODEL)," & vbNewLine
            QRY = QRY & "RCV_DATE, " & vbNewLine
            QRY = QRY & "(SELECT CODE_NAME FROM TBL_CODEMASTER WHERE CLASS_ID = 'R0001' AND CODE_ID = A.C_PRC)" & vbNewLine
            QRY = QRY & "FROM TBL_PRODMASTER A" & vbNewLine
            QRY = QRY & "WHERE PROD_NO = '" & TextBoxX1.Text & "'" & vbNewLine

            TRF_RS = Query_RS_ALL(QRY)

            If TRF_RS Is Nothing Then
                System.Console.Beep(3000, 400)
                System.Console.Beep(3000, 400)
                Modal_Error("존재하지 않는 제품번호 입니다.")
                Exit Sub
            End If

            If TRF_RS.RecordCount > 0 Then

                With FpSpread1.ActiveSheet
                    .AddRows(0, 1)
                    .Rows(0).ForeColor = Color.OrangeRed
                    .SetValue(0, 0, TRF_RS(0).Value)
                    .SetValue(0, 1, TRF_RS(1).Value)
                    .SetValue(0, 2, TRF_RS(2).Value)
                    .SetValue(0, 3, TRF_RS(3).Value)
                    .SetValue(0, 4, TRF_RS(4).Value)
                End With
                Spread_AutoCol(FpSpread1)

            End If
            UcMsgPanel1.LabelX1.Text = CInt(UcMsgPanel1.LabelX1.Text) + 1

            Me.TextBoxX1.Focus()
            Me.TextBoxX1.SelectAll()
        End If


        TRF_RS = Nothing

    End Sub


    Private Sub ComboBoxEx1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBoxEx1.SelectedIndexChanged
        Me.TextBoxX1.Focus()
        Me.TextBoxX1.SelectAll()
    End Sub

    Private Sub SaveBtn_Click(sender As Object, e As EventArgs) Handles SaveBtn.Click

        With FpSpread1.ActiveSheet
            If .RowCount = 0 Then
                Exit Sub
            End If

            For i As Integer = 0 To .RowCount - 1
                If .Rows(i).ForeColor = Color.OrangeRed Then
                    Dim qry As String = ""

                    qry = qry & "update tbl_prodmaster set c_prc = 'W2000', u_date = getdate(), u_person = '" & Emp_No & "' where prod_no = '" & .GetValue(i, 0) & "'" & vbNewLine
                    qry = qry & "insert into tbl_prodmaster_b select * from tbl_prodmaster where prod_no = '" & .GetValue(i, 0) & "'" & vbNewLine
                    qry = qry & "delete from tbl_prodmaster where prod_no = '" & .GetValue(i, 0) & "'" & vbNewLine
                    qry = qry & "" & vbNewLine

                    If Insert_Data(qry) = False Then
                        System.Console.Beep(3000, 400)
                        System.Console.Beep(3000, 400)
                        Modal_Error("저장이 실패하였습니다.")
                        Exit Sub
                    Else
                        .Rows(i).ForeColor = Color.Black
                        Print_UPCBarcode(.GetValue(i, 1), ComboBoxEx2.Text, RichTextBox1)
                    End If
                End If
            Next
            MessageBox.Show("저장 되었습니다.")

            .RowCount = 0
        End With


    End Sub


End Class