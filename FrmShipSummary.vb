Public Class FrmShipSummary


    Private Sub FrmShipSummary_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        DockContainerItem1.Text = "Form Menu"
        DockContainerItem2.Text = "Retrieve Condition"

        DateTimeInput1.Value = Now
        DateTimeInput2.Value = Now

        If Query_Combo(Me.ComboBoxEx1, "SELECT model_no FROM tbl_modelmaster WHERE site_id = '" & Site_id & "' and active = 'Y' ORDER BY model_no") = True Then
        End If
        Me.ComboBoxEx1.Items.Add("ALL")
        Me.ComboBoxEx1.Text = "ALL"

        If Spread_Setting(FpSpread1, Me.Name) = True Then
            FpSpread1.ActiveSheet.RowCount = 0
            Spread_AutoCol(FpSpread1)
        End If

        If Spread_Setting(FpSpread2, Me.Name) = True Then
            FpSpread2.ActiveSheet.RowCount = 0
            Spread_AutoCol(FpSpread2)
        End If

        If Spread_Setting(FpSpread3, Me.Name) = True Then
            FpSpread3.ActiveSheet.RowCount = 0
            Spread_AutoCol(FpSpread3)
        End If

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
        Formbtn_Authority(Me.FindBtn, Me.Name, "FIND")

    End Sub


    Function Refresh_Result() As Boolean

        With FpSpread1.ActiveSheet
            If Query_Spread(FpSpread1, "EXEC SP_FRMSHIPSUMMARY_GETSUMMARY '" & Site_id & "','" & DateTimeInput1.Text & "','" & DateTimeInput2.Text & "','" & ComboBoxEx1.Text & "'", 1) = True Then
                FpSpread1.ActiveSheet.RowCount = FpSpread1.ActiveSheet.RowCount + 1
                FpSpread1.ActiveSheet.Cells(FpSpread1.ActiveSheet.RowCount - 1, 1, FpSpread1.ActiveSheet.RowCount - 1, FpSpread1.ActiveSheet.ColumnCount - 1).CellType = intcell
                FpSpread1.ActiveSheet.Cells(FpSpread1.ActiveSheet.RowCount - 1, 0, FpSpread1.ActiveSheet.RowCount - 1, FpSpread1.ActiveSheet.ColumnCount - 1).BackColor = Color.Yellow
                FpSpread1.ActiveSheet.SetText(FpSpread1.ActiveSheet.RowCount - 1, 0, "TOTAL")

                SPREAD_ROW_TOTAL(FpSpread1, 1, FpSpread1.ActiveSheet.ColumnCount - 1, 1)
                Spread_AutoCol(FpSpread1)
            End If
        End With

        disp_main1()

    End Function

    Private Sub disp_main1()
        Dim i, j As Integer
        Dim TotRs As ADODB.Recordset
        Dim S1 = Me.FpSpread4.ActiveSheet
        Dim t1, t2, t3, t4, t5 As Integer

        If S1.RowCount > 0 Then
            S1.Rows.Clear()
        End If

        t1 = 0
        t2 = 0
        t3 = 0
        t4 = 0
        t5 = 0

        TotRs = Query_RS_ALL("EXEC SP_FRMMAINDISP '" & Site_id & "',1")

        If TotRs Is Nothing Then
            Exit Sub
        End If

        For i = 0 To TotRs.RecordCount - 1
            S1.RowCount += 1
            For j = 0 To S1.ColumnCount - 1
                S1.SetValue(i, j, TotRs(j).Value)
            Next
            t1 += S1.GetValue(i, 1)
            t2 += S1.GetValue(i, 2)
            t3 += S1.GetValue(i, 3)
            t4 += S1.GetValue(i, 4)
            t5 += S1.GetValue(i, 5)
            TotRs.MoveNext()
        Next

        S1.RowCount += 1
        S1.Rows(S1.RowCount - 1).BackColor = Color.LightPink
        S1.SetValue(S1.RowCount - 1, 0, "TOTAL")
        S1.SetValue(S1.RowCount - 1, 1, t1)
        S1.SetValue(S1.RowCount - 1, 2, t2)
        S1.SetValue(S1.RowCount - 1, 3, t3)
        S1.SetValue(S1.RowCount - 1, 4, t4)
        S1.SetValue(S1.RowCount - 1, 5, t5)
        If t4 = 0 Then
            S1.SetValue(S1.RowCount - 1, 6, 0)
        Else
            S1.SetValue(S1.RowCount - 1, 6, t4 / (t4 + t5))
        End If

        If t1 = 0 Then
            S1.SetValue(S1.RowCount - 1, 7, 0)
        Else
            S1.SetValue(S1.RowCount - 1, 7, (t4 + t5) / t1)
        End If


    End Sub


    Private Sub FindBtn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FindBtn.Click
        FpSpread1.ActiveSheet.RowCount = 0
        FpSpread2.ActiveSheet.RowCount = 0
        FpSpread3.ActiveSheet.RowCount = 0

        Refresh_Result()
    End Sub

    Private Sub FpSpread1_CellDoubleClick(ByVal sender As Object, ByVal e As FarPoint.Win.Spread.CellClickEventArgs) Handles FpSpread1.CellDoubleClick
        FpSpread3.ActiveSheet.RowCount = 0

        If Query_Spread(FpSpread2, "EXEC SP_FRMSHIPSUMMARY_GETDETAILS '" & Site_id & "','" & DateTimeInput1.Text & "','" & DateTimeInput2.Text & "','" & FpSpread1.ActiveSheet.GetValue(e.Row, 0) & "'", 1) = True Then
            Spread_AutoCol(FpSpread2)
            'FpSpread2.ActiveSheet.Protect = False
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
            If Spread_Print(Me.FpSpread1, GroupPanel1.Text, 0) = False Then
                MsgBox("Fail to Print")
            End If
        ElseIf save_excel = "FpSpread2" Then
            If Spread_Print(Me.FpSpread2, GroupPanel2.Text, 0) = False Then
                MsgBox("Fail to Print")
            End If
        ElseIf save_excel = "FpSpread3" Then
            If Spread_Print(Me.FpSpread3, GroupPanel3.Text, 0) = False Then
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
        ElseIf save_excel = "FpSpread3" Then
            File_Save(SaveFileDialog1, FpSpread3)
        Else
            MessageBox.Show("Select Spread for Saving!!")
        End If

    End Sub

    Private Sub FpSpread1_CellClick(ByVal sender As System.Object, ByVal e As FarPoint.Win.Spread.CellClickEventArgs) Handles FpSpread1.CellClick
        save_excel = "FpSpread1"
    End Sub

    Private Sub FpSpread2_CellClick(ByVal sender As System.Object, ByVal e As FarPoint.Win.Spread.CellClickEventArgs) Handles FpSpread2.CellClick
        save_excel = "FpSpread2"
        FpSpread3.ActiveSheet.RowCount = 0

        If Query_Spread(FpSpread3, "EXEC SP_FRMSHIPSUMMARY_BOXDT '" & Site_id & "','" & FpSpread2.ActiveSheet.GetValue(e.Row, 4) & "','" & FpSpread2.ActiveSheet.GetValue(e.Row, 0) & "','" & FpSpread2.ActiveSheet.GetValue(e.Row, 1) & "'", 1) = True Then
            Spread_AutoCol(FpSpread3)
        End If

    End Sub

    Private Sub FpSpread3_CellClick(ByVal sender As Object, ByVal e As FarPoint.Win.Spread.CellClickEventArgs) Handles FpSpread3.CellClick
        save_excel = "FpSpread3"
    End Sub

End Class