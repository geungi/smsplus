Public Class FrmWCTrf

    Private Sub FrmWCTrf_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.GotFocus
        Me.ComboBoxEx2.Focus()
    End Sub

    Private Sub FrmWCTrf_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        DockContainerItem1.Text = "Form Menu"
        DockContainerItem2.Text = "Input Condition"

        If Query_Combo(Me.ComboBoxEx1, "SELECT CODE_NAME FROM tbl_CODEmaster WHERE site_id = '" & Site_id & "' and CLASS_ID = 'R0001' AND active = 'Y' ORDER BY CODE_ID") = True Then
        End If

        ComboBoxEx2.Items.Add("LOT")
        ComboBoxEx2.Text = "LOT"

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

    Private Sub ComboBoxEx2_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBoxEx2.SelectedIndexChanged
        If ComboBoxEx2.Text = "FLIP" Then
            If Query_Combo(Me.ComboBoxEx1, "SELECT CODE_NAME FROM tbl_CODEmaster WHERE site_id = '" & Site_id & "' and CLASS_ID = 'R0002' AND active = 'Y' ORDER BY CODE_ID") = True Then
            End If
        Else
            If Query_Combo(Me.ComboBoxEx1, "SELECT CODE_NAME FROM tbl_CODEmaster WHERE site_id = '" & Site_id & "' and CLASS_ID = 'R0001' AND active = 'Y' ORDER BY CODE_ID") = True Then
            End If
        End If
        Me.ComboBoxEx1.Focus()
    End Sub

    Private Sub TextBoxX1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBoxX1.Click
        Me.TextBoxX1.Text = ""
    End Sub


    Private Sub TextBoxX1_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TextBoxX1.KeyDown
        Dim TRF_RS As New ADODB.Recordset

        If ComboBoxEx1.Text = "" Then
            System.Console.Beep(3000, 400)
            System.Console.Beep(3000, 400)
            System.Windows.Forms.MessageBox.Show("SELECT NEXT WORKCENTER!!")
            'Me.TextBoxX1.Text = ""
            Exit Sub
        End If

        If e.KeyCode <> Keys.Enter Then
            Exit Sub
        End If

        If ComboBoxEx2.Text = "LOT" Then
            If LOT_VERIFY(Me.TextBoxX1, e) = True Then

                TRF_RS = Query_RS_ALL("EXEC SP_FRMWCTRF_LOT'" & Site_id & "','" & TextBoxX1.Text & "','" & Emp_No & "','" & ComboBoxEx1.Text & "','GOOD', '1'")

                If TRF_RS.RecordCount > 0 Then

                    If TRF_RS(0).Value = "NOT EXIST ESN" Then
                        System.Console.Beep(3000, 400)
                        System.Console.Beep(3000, 400)
                        System.Windows.Forms.MessageBox.Show("존재하지 않는 생산LOT 입니다!!")
                        TRF_RS = Nothing

                        'Me.TextBoxX1.Text = ""
                        Me.TextBoxX1.Focus()
                        Me.TextBoxX1.SelectAll()
                        Exit Sub
                    Else
                        With FpSpread1.ActiveSheet
                            .AddRows(0, 1)
                            .SetValue(0, 0, TRF_RS(0).Value)
                            .SetValue(0, 1, TRF_RS(1).Value)
                            .SetValue(0, 2, TRF_RS(2).Value)
                            .SetValue(0, 3, TRF_RS(3).Value)
                        End With
                        Spread_AutoCol(FpSpread1)
                    End If

                End If
                'UcMsgPanel1.LabelX1.Text = CInt(UcMsgPanel1.LabelX1.Text) + 1

                'TextBoxX1.Text = ""
                Me.TextBoxX1.Focus()
                Me.TextBoxX1.SelectAll()
            End If

        End If

        TRF_RS = Nothing

    End Sub


    Private Sub FrmWCTrf_Shown(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Shown
        Me.ComboBoxEx2.Focus()
    End Sub

    Private Sub ComboBoxEx1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBoxEx1.SelectedIndexChanged
        Me.TextBoxX1.Focus()
        Me.TextBoxX1.SelectAll()
    End Sub

End Class