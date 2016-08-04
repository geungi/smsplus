Public Class FrmAuthority

    Private Sub FrmAuthority_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        DockContainerItem1.Text = "Form Menu"
        DockContainerItem2.Text = "Input Condition"

        If Query_Combo(Me.ComboBoxEx1, "SELECT '[' + emp_no + '] ' + emp_nm FROM tbl_empmaster WHERE site_id = '" & Site_id & "' and retire_yn = 'N' ORDER BY emp_no") = True Then
        End If

        If Query_Combo(Me.ComboBoxEx2, "SELECT '[' + emp_no + '] ' + emp_nm FROM tbl_empmaster WHERE site_id = '" & Site_id & "' and retire_yn = 'N' ORDER BY emp_no") = True Then
        End If


        If Spread_Setting(FpSpread1, Me.Name) = True Then
            FpSpread1.ActiveSheet.RowCount = 0
            Spread_AutoCol(FpSpread1)
        End If

        Formbim_Authority(Me.NewBtn, Me.Name, "NEW")
        Formbtn_Authority(Me.NewBtn1, Me.Name, "NEW")
        Formbim_Authority(Me.SaveBtn, Me.Name, "SAVE")
        Formbtn_Authority(Me.SaveBtn1, Me.Name, "SAVE")
        Formbim_Authority(Me.DelBtn, Me.Name, "DELETE")
        Formbtn_Authority(Me.DelBtn1, Me.Name, "DELETE")
        Formbim_Authority(Me.PrtBtn, Me.Name, "PRINT")
        Formbtn_Authority(Me.PrtBtn1, Me.Name, "PRINT")
        Formbim_Authority(Me.Excel, Me.Name, "EXCEL")
        Formbtn_Authority(Me.XlsBtn1, Me.Name, "EXCEL")
        Formbim_Authority(Me.FindBtn, Me.Name, "FIND")

    End Sub

    Private Sub FindBtn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBoxEx1.SelectedIndexChanged, FindBtn.Click

        If ComboBoxEx1.Text = "" Then
            System.Console.Beep(3000, 400)
            System.Console.Beep(3000, 400)
            System.Windows.Forms.MessageBox.Show("Select Employee!!")

            Exit Sub

        End If

        If Query_Spread(FpSpread1, "EXEC SP_FRMAUTHORITY_GETFORM '" & Site_id & "','" & Microsoft.VisualBasic.Mid(ComboBoxEx1.Text, 2, 5) & "'", 1) = True Then
            Spread_AutoCol(FpSpread1)
            FpSpread1.ActiveSheet.Columns(3, FpSpread1.ActiveSheet.ColumnCount - 1).Locked = False
        End If

        ComboBoxEx2.Text = ComboBoxEx1.Text

    End Sub


    Private Sub FpSpread1_Change(ByVal sender As Object, ByVal e As FarPoint.Win.Spread.ChangeEventArgs) Handles FpSpread1.Change
        FpSpread1.ActiveSheet.Rows(e.Row).ForeColor = Color.OrangeRed
    End Sub


    Private Sub SaveBtn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SaveBtn.Click, SaveBtn1.Click

        Dim i As Integer
        Me.Cursor = Cursors.WaitCursor


        With FpSpread1.ActiveSheet
            .SetActiveCell(0, 0)
            MainFrm.ProgressBarItem1.Maximum = .RowCount
            For i = 0 To .RowCount - 1
                MainFrm.ProgressBarItem1.Value = i + 1
                If .Rows(i).ForeColor <> Color.OrangeRed Then
                    Continue For
                End If

                'Show, New, Save, Delete, Excel, Print, Find 권한 부여
                If .GetValue(i, 3) <> 0 Then
                    Insert_Data("exec SP_FRMAUTHORITY_SETAUTHO '" & Site_id & "','" & .GetValue(i, 2) & "','', '" & Microsoft.VisualBasic.Mid(ComboBoxEx2.Text, 2, 5) & "','2'")
                Else
                    Insert_Data("exec SP_FRMAUTHORITY_SETAUTHO '" & Site_id & "','" & .GetValue(i, 2) & "','', '" & Microsoft.VisualBasic.Mid(ComboBoxEx2.Text, 2, 5) & "','3'")
                End If

                If .GetValue(i, 4) <> 0 Then
                    Insert_Data("exec SP_FRMAUTHORITY_SETAUTHO '" & Site_id & "','" & .GetValue(i, 2) & "','NEW', '" & Microsoft.VisualBasic.Mid(ComboBoxEx2.Text, 2, 5) & "','1'")
                Else
                    Insert_Data("exec SP_FRMAUTHORITY_SETAUTHO '" & Site_id & "','" & .GetValue(i, 2) & "','NEW', '" & Microsoft.VisualBasic.Mid(ComboBoxEx2.Text, 2, 5) & "','0'")
                End If

                If .GetValue(i, 5) <> 0 Then
                    Insert_Data("exec SP_FRMAUTHORITY_SETAUTHO '" & Site_id & "','" & .GetValue(i, 2) & "','SAVE', '" & Microsoft.VisualBasic.Mid(ComboBoxEx2.Text, 2, 5) & "','1'")
                Else
                    Insert_Data("exec SP_FRMAUTHORITY_SETAUTHO '" & Site_id & "','" & .GetValue(i, 2) & "','SAVE', '" & Microsoft.VisualBasic.Mid(ComboBoxEx2.Text, 2, 5) & "','0'")
                End If

                If .GetValue(i, 6) <> 0 Then
                    Insert_Data("exec SP_FRMAUTHORITY_SETAUTHO '" & Site_id & "','" & .GetValue(i, 2) & "','DELETE', '" & Microsoft.VisualBasic.Mid(ComboBoxEx2.Text, 2, 5) & "','1'")
                Else
                    Insert_Data("exec SP_FRMAUTHORITY_SETAUTHO '" & Site_id & "','" & .GetValue(i, 2) & "','DELETE', '" & Microsoft.VisualBasic.Mid(ComboBoxEx2.Text, 2, 5) & "','0'")
                End If

            

                If .GetValue(i, 7) <> 0 Then
                    Insert_Data("exec SP_FRMAUTHORITY_SETAUTHO '" & Site_id & "','" & .GetValue(i, 2) & "','PRINT', '" & Microsoft.VisualBasic.Mid(ComboBoxEx2.Text, 2, 5) & "','1'")
                Else
                    Insert_Data("exec SP_FRMAUTHORITY_SETAUTHO '" & Site_id & "','" & .GetValue(i, 2) & "','PRINT', '" & Microsoft.VisualBasic.Mid(ComboBoxEx2.Text, 2, 5) & "','0'")
                End If

                If .GetValue(i, 8) <> 0 Then
                    Insert_Data("exec SP_FRMAUTHORITY_SETAUTHO '" & Site_id & "','" & .GetValue(i, 2) & "','EXCEL', '" & Microsoft.VisualBasic.Mid(ComboBoxEx2.Text, 2, 5) & "','1'")
                Else
                    Insert_Data("exec SP_FRMAUTHORITY_SETAUTHO '" & Site_id & "','" & .GetValue(i, 2) & "','EXCEL', '" & Microsoft.VisualBasic.Mid(ComboBoxEx2.Text, 2, 5) & "','0'")
                End If

                If .GetValue(i, 9) <> 0 Then
                    Insert_Data("exec SP_FRMAUTHORITY_SETAUTHO '" & Site_id & "','" & .GetValue(i, 2) & "','FIND', '" & Microsoft.VisualBasic.Mid(ComboBoxEx2.Text, 2, 5) & "','1'")
                Else
                    Insert_Data("exec SP_FRMAUTHORITY_SETAUTHO '" & Site_id & "','" & .GetValue(i, 2) & "','FIND', '" & Microsoft.VisualBasic.Mid(ComboBoxEx2.Text, 2, 5) & "','0'")
                End If

                .Rows(i).ForeColor = Color.Black
            Next
        End With
        MainFrm.ProgressBarItem1.Value = 0

        ComboBoxEx1.Text = ComboBoxEx2.Text


        Me.Cursor = Cursors.Default
    End Sub

    Private Sub ButtonItem4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DelBtn.Click, DelBtn1.Click

    End Sub

    Private Sub ButtonItem5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PrtBtn.Click, PrtBtn1.Click
        If save_excel = "" Then
            MessageBox.Show("Select Spread for Printing!!")
            Exit Sub
        End If

        If save_excel = "FpSpread1" Then
            If Spread_Print(Me.FpSpread1, "Technician Performance", 0) = False Then
                MsgBox("Fail to Print")
            End If
           
        Else
            MessageBox.Show("Select Spread for Printing!!")
        End If
    End Sub

    Private Sub ButtonItem6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Excel.Click, XlsBtn1.Click

        If save_excel = "" Then
            MessageBox.Show("Select Spread for Saving!!")
            Exit Sub
        End If

        If save_excel = "FpSpread1" Then
            File_Save_1(SaveFileDialog1, FpSpread1)
           
        Else
            MessageBox.Show("Select Spread for Saving!!")
        End If

    End Sub


    Private Sub ButtonItem2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonItem2.Click
        Dim i As Integer

        For i = 3 To FpSpread1.ActiveSheet.ColumnCount - 1

            FpSpread1.ActiveSheet.SetValue(FpSpread1.ActiveSheet.ActiveRowIndex, i, True)
            FpSpread1.ActiveSheet.Rows(FpSpread1.ActiveSheet.ActiveRowIndex).ForeColor = Color.OrangeRed

        Next

    End Sub


    Private Sub ButtonItem3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonItem3.Click
        Dim i As Integer

        For i = 0 To FpSpread1.ActiveSheet.RowCount - 1

            FpSpread1.ActiveSheet.SetValue(i, FpSpread1.ActiveSheet.ActiveColumnIndex, True)
            FpSpread1.ActiveSheet.Rows(i).ForeColor = Color.OrangeRed

        Next

    End Sub

    Private Sub ButtonItem4_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonItem4.Click
        Dim i As Integer

        For i = 3 To FpSpread1.ActiveSheet.ColumnCount - 1

            FpSpread1.ActiveSheet.SetValue(FpSpread1.ActiveSheet.ActiveRowIndex, i, False)
            FpSpread1.ActiveSheet.Rows(FpSpread1.ActiveSheet.ActiveRowIndex).ForeColor = Color.OrangeRed

        Next
       
    End Sub

    Private Sub ButtonItem5_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonItem5.Click
        Dim i As Integer

        For i = 0 To FpSpread1.ActiveSheet.RowCount - 1

            FpSpread1.ActiveSheet.SetValue(i, FpSpread1.ActiveSheet.ActiveColumnIndex, False)
            FpSpread1.ActiveSheet.Rows(i).ForeColor = Color.OrangeRed

        Next
    End Sub
End Class