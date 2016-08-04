Public Class FrmPD

    Private Sub FrmPD_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.GotFocus
        Me.TextBoxX1.Focus()
    End Sub

    Private Sub FrmPD_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            DockContainerItem1.Text = "실행 메뉴"
            DockContainerItem2.Text = "입력 조건"
            DockContainerItem3.Text = "BER 현황"

            'Bar3.AutoHide = True

            DateTimeInput1.Value = Now
            DateTimeInput2.Value = Now

            If Query_Combo(Me.ComboBoxEx1, "SELECT CODE_ID FROM tbl_CODEmaster WHERE site_id = '" & Site_id & "' and CLASS_ID = 'R0007' AND active = 'Y' AND CODE_ID <> 'GOOD' ORDER BY DIS_ORDER") = True Then
            End If

            Me.ComboBoxEx1.Text = "BER"

            If Query_Combo(Me.ComboBoxEx2, "SELECT CODE_ID + ':' + code_name FROM TBL_DEFECT WHERE code_id <> 'NT01' and ACTIVE = 'Y' ORDER BY CODE_ID") = True Then
                Me.ComboBoxEx2.Text = "X3"
            End If

            If Spread_Setting(FpSpread1, Me.Name) = True Then
                FpSpread1.ActiveSheet.RowCount = 0
                FpSpread1.ActiveSheet.FrozenColumnCount = 1
                Spread_AutoCol(FpSpread1)
            End If

            If Spread_Setting(FpSpread2, Me.Name) = True Then
                FpSpread2.ActiveSheet.RowCount = 0
                Spread_AutoCol(FpSpread1)
            End If

            Formbim_Authority(Me.ButtonItem1, Me.Name, "NEW")
            Formbtn_Authority(Me.NewBtn1, Me.Name, "NEW")
            Formbim_Authority(Me.ButtonItem2, Me.Name, "SAVE")
            Formbtn_Authority(Me.SaveBtn1, Me.Name, "SAVE")
            Formbim_Authority(Me.ButtonItem3, Me.Name, "DELETE")
            Formbtn_Authority(Me.DelBtn1, Me.Name, "DELETE")
            Formbim_Authority(Me.ButtonItem4, Me.Name, "PRINT")
            Formbtn_Authority(Me.PrtBtn1, Me.Name, "PRINT")
            Formbim_Authority(Me.ButtonItem5, Me.Name, "EXCEL")
            Formbtn_Authority(Me.XlsBtn1, Me.Name, "EXCEL")
            Formbtn_Authority(Me.FindBtn, Me.Name, "FIND")


        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message, "ERROR")
        End Try
    End Sub

    Private Sub TextBoxX1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBoxX1.Click
        Me.TextBoxX1.Text = ""
    End Sub

    Private Sub TextBoxX1_KeyDown1(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TextBoxX1.KeyDown
        Dim pd_rs As New ADODB.Recordset

        Try
            If ComboBoxEx1.Text = "" Or ComboBoxEx2.Text = "" Then
                System.Console.Beep(3000, 400)
                System.Console.Beep(3000, 400)
                Modal_Error(Me.TextBoxX1.Text & vbNewLine & "SELECT INPUT CONDITION!!")
                'Me.TextBoxX1.Text = ""
                Me.TextBoxX1.Focus()
                Me.TextBoxX1.SelectAll()
                Exit Sub
            End If

            If e.KeyCode <> Keys.Enter Then
                Exit Sub
            End If

            If LOT_VERIFY(Me.TextBoxX1, e) = True Then

                If Check_Valid_LOT(TextBoxX1.Text, Me.Name) = False Then
                    Me.TextBoxX1.Text = ""
                    Exit Sub
                End If
            End If

            pd_rs = Nothing

            SaveBtn1.Focus()

        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message, "ERROR")
        End Try

    End Sub

    Private Sub ButtonItem2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonItem2.Click, SaveBtn1.Click
        Dim pd_rs As New ADODB.Recordset

        pd_rs = Query_RS_ALL("SELECT COUNT(LOT_NO), ISNULL(RETURN_DV,'GOOD'), '', '', '', '' AS QAR, '', datediff(day, c_date, getdate()), '', MODEL, '', T_DEF_CD, (select model_dv from tbl_modelmaster where model_no = a.model ), '' FROM TBL_LOTmaster A WHERE SITE_ID = '" & Site_id & "' AND LOT_NO = '" & TextBoxX1.Text & "' GROUP BY LOT_NO, RETURN_DV,  model, c_date,T_DEF_CD")


        If pd_rs Is Nothing Then
            System.Console.Beep(3000, 400)
            System.Console.Beep(3000, 400)
            Modal_Error(TextBoxX1.Text & vbNewLine & "Not exist ESN!!")
            pd_rs = Nothing
            Me.TextBoxX1.Focus()
            Me.TextBoxX1.SelectAll()
            Exit Sub
        End If

        If pd_rs(0).Value = 0 Then
            System.Console.Beep(3000, 400)
            System.Console.Beep(3000, 400)
            Modal_Error(TextBoxX1.Text & vbNewLine & "Not exist ESN!!")
            'Me.TextBoxX1.Text = ""
            pd_rs = Nothing
            Me.TextBoxX1.Focus()
            Me.TextBoxX1.SelectAll()
            Exit Sub
        Else
            If pd_rs(1).Value <> "GOOD" Then
                System.Console.Beep(3000, 400)
                System.Console.Beep(3000, 400)
                Modal_Error(TextBoxX1.Text & vbNewLine & "ALREADY " & pd_rs(1).Value & "!!")
                'Me.TextBoxX1.Text = ""
                pd_rs = Nothing
                Me.TextBoxX1.Focus()
                Me.TextBoxX1.SelectAll()
                Exit Sub
            Else


                If ComboBoxEx1.Text = "BER" Then
                    'If pd_rs(6).Value <> "SPR" Then
                    '    System.Console.Beep(3000, 400)
                    '    System.Console.Beep(3000, 400)
                    '    Modal_Error(TextBoxX1.Text & vbNewLine & "SPRINT MODEL ONLY!!")
                    '    'Me.TextBoxX1.Text = ""
                    '    pd_rs = Nothing
                    '    Me.TextBoxX1.Focus()
                    '    Me.TextBoxX1.SelectAll()
                    '    Exit Sub
                    'End If

                End If

                Setpd()
                Spread_AutoCol(FpSpread1)
            End If
        End If
        'Me.TextBoxX1.Text = ""
        Me.TextBoxX1.Focus()
        Me.TextBoxX1.SelectAll()

        GetSummary()

    End Sub


    Private Sub GetSummary() Handles FindBtn.Click

        If Query_Spread(FpSpread2, "exec SP_FRMPD_GETsummary '" & Site_id & "','" & DateTimeInput1.Text & "','" & DateTimeInput2.Text & "'", 1) = True Then
            Spread_AutoCol(FpSpread2)
        End If

    End Sub

    Private Sub Setpd()
        Try
            With FpSpread1.ActiveSheet
                .AddRows(0, 1)
                .Rows(0).ForeColor = Color.OrangeRed
                .SetValue(0, 0, TextBoxX1.Text)
                .SetValue(0, 1, Query_RS("select model FROM TBL_LOTmaster WHERE SITE_ID = '" & Site_id & "' AND ESN = '" & TextBoxX1.Text & "'"))
                .SetValue(0, 2, Mid(ComboBoxEx2.Text, 1, 4))
                .SetValue(0, 3, "")
                .SetValue(0, 4, "")
                .SetValue(0, 5, Query_RS("select C_date FROM TBL_LOTmaster WHERE SITE_ID = '" & Site_id & "' AND ESN = '" & TextBoxX1.Text & "'"))

                If Insert_Data("EXEC SP_FRMPD_SETPD '" & Site_id & "','" & TextBoxX1.Text & "','" & Mid(ComboBoxEx2.Text, 1, 4) & "','','BER','" & ComboBoxEx1.Text & " : ','" & OP_No & "','포장', '" & ComboBoxEx1.Text & "','1'") = True Then
                    .Rows(0).ForeColor = Color.Black
                    Spread_AutoCol(FpSpread1)
                End If
            End With

            UcMsgPanel1.LabelX1.Text = CInt(UcMsgPanel1.LabelX1.Text) + 1

        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message, "ERROR")
        End Try
    End Sub

    Private Sub ComboBoxEx1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBoxEx1.SelectedIndexChanged


    End Sub

    Private Sub ButtonItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonItem1.Click, NewBtn1.Click

    End Sub


    Private Sub ButtonItem3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonItem3.Click, DelBtn1.Click

    End Sub

    Private Sub ButtonItem4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonItem4.Click, PrtBtn1.Click
        If save_excel = "" Then
            MessageBox.Show("Select Spread for Printing!!")
            Exit Sub
        End If

        If save_excel = "FpSpread1" Then
            If Spread_Print(Me.FpSpread1, "SWAP DETAILS(" & DateTimeInput1.Text & "~" & DateTimeInput2.Text & ")", 0) = False Then
                MsgBox("Fail to Print")
            End If
        ElseIf save_excel = "FpSpread2" Then
            If Spread_Print(Me.FpSpread2, "SWAP Details", 0) = False Then
                MsgBox("Fail to Print")
            End If
        Else
            MessageBox.Show("Select Spread for Printing!!")
        End If
    End Sub

    Private Sub ButtonItem5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonItem5.Click, XlsBtn1.Click

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

    Private Sub ComboBoxEx2_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBoxEx2.SelectedIndexChanged

    End Sub

End Class