Public Class FrmMainDisp

    Private Sub FrmMainDisp_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Me.ContextMenuBar1.SetContextMenuEx(Me.Chart2, ButtonItem1)
        Me.ContextMenuBar1.SetContextMenuEx(Me.FpSpread1, ButtonItem3)
        Me.ContextMenuBar1.SetContextMenuEx(Me.Chart1, ButtonItem5)

        DockContainerItem2.Text = Query_RS("Select datename(year,getdate())") & "년 출하 현황"
        DockContainerItem4.Text = Query_RS("Select datename(month,getdate())") & "월 출하 현황"

        With FpSpread1.ActiveSheet
            .ColumnHeader.Cells(0, 7).RowSpan = 3
        End With

        'With FpSpread3.ActiveSheet
        '    .ColumnHeader.Cells(0, 9).Text = "PACK"
        '    .ColumnHeader.Cells(0, 1).Text = "ASSEMBLY"
        'End With

        disp_main1()
        '        disp_main6()
        disp_main5()
        disp_main2()
        'disp_main3()
        'disp_main4()

    End Sub

    Private Sub disp_main1()
        Dim i, j As Integer
        Dim TotRs As ADODB.Recordset
        Dim S1 = Me.FpSpread1.ActiveSheet
        Dim t1, t2, t3, t4, t5, t6 As Integer

        If S1.RowCount > 0 Then
            S1.Rows.Clear()
        End If

        t1 = 0
        t2 = 0
        t3 = 0
        t4 = 0
        t5 = 0
        t6 = 0

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
            t6 += S1.GetValue(i, 6)
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
        'If t4 = 0 Then
        '    S1.SetValue(S1.RowCount - 1, 6, 0)
        'Else
        S1.SetValue(S1.RowCount - 1, 6, t6)
        'End If
        S1.Columns(0).Width = 100
        S1.Columns(6).CellType = intcell

        If t1 = 0 Then
            S1.SetValue(S1.RowCount - 1, 7, 0)
        Else
            S1.SetValue(S1.RowCount - 1, 7, (t2 + t4) / t1)
            Chart_Query3(Me.Chart3, False, 3, DataVisualization.Charting.SeriesChartType.Column, New String() {"Target", "Shipping"}, 1000, "EXEC SP_FRMMAINDISP'" & Site_id & "',6")

        End If


    End Sub



    Private Sub disp_main5()
        Dim QRY As String

        With FpSpread3.ActiveSheet
            If .RowCount > 0 Then
                .Rows.Clear()
            End If

            QRY = "EXEC SP_FRMMAINDISP '" & Site_id & "',4"
            
            If Query_Spread(FpSpread3, QRY, 1) = True Then
                Spread_AutoCol(FpSpread3)
                .Columns(1, .ColumnCount - 1).Width = 50
                intcell.ShowSeparator = True
                intcell.DecimalPlaces = 0
                .Columns(1, .ColumnCount - 1).CellType = intcell

            End If

            QRY = "EXEC SP_FRMMAINDISP '" & Site_id & "',7"
            

            Chart_Query_main(Me.Chart4, False, 2, DataVisualization.Charting.SeriesChartType.Line, New String() {"TGT_TRG", "TGT_DAS", "TGT_ASM", "TGT_DL", "TGT_QC", "TGT_PKG"}, 1000, QRY)
            'Chart4.Series.Item(
        End With
    End Sub

    Private Sub disp_main2()
        Chart_Query2(Me.Chart1, False, 0, DataVisualization.Charting.SeriesChartType.Column, New String() {"Claimed Qty"}, 1000, "EXEC SP_FRMMAINDISP '" & Site_id & "',2")
    End Sub
    Private Sub disp_main6()

        '        Chart_Query2(Me.Chart3, False, 0, DataVisualization.Charting.SeriesChartType.Column, New String() {"MODEL", "Target", "Shipping"}, 1000, "EXEC SP_FRMMAINDISP '" & Site_id & "',6")
        '        Chart_Spread_All_1(Me.Chart3, False, 2, DataVisualization.Charting.SeriesChartType.Column, Me.FpSpread1, New String() {"Target", "SHIPPING"}, 0, New Integer() {1, 4}, True)
    End Sub

    Private Sub disp_main3()
        Query_Spread(Me.FpSpread2, "EXEC SP_FRMMAINDISP '" & Site_id & "',3", 1)
        If Me.FpSpread2.ActiveSheet.RowCount > 0 Then
            Me.FpSpread2.ActiveSheet.Columns(Me.FpSpread2.ActiveSheet.Columns.Count - 1).BackColor = Color.GreenYellow
            Me.FpSpread2.ActiveSheet.Rows(Me.FpSpread2.ActiveSheet.RowCount - 1).BackColor = Color.GreenYellow
        End If
        Spread_AutoCol(Me.FpSpread2)
    End Sub

    Private Sub disp_main4()
        Chart_Query3(Me.Chart2, False, 0, DataVisualization.Charting.SeriesChartType.Column, New String() {"Receiving", "Shipping", "Claimed", "Invoiced"}, 1000, "EXEC SP_FRMMAINDISP '" & Site_id & "',40")

    End Sub

    Private Sub DotNetBarManager1_DockTabClosing(ByVal sender As Object, ByVal e As DevComponents.DotNetBar.DockTabClosingEventArgs)
        e.Cancel = True
    End Sub

    Private Sub ButtonItem2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonItem2.Click
        disp_main4()
    End Sub

    Private Sub ButtonItem4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonItem4.Click
        disp_main1()
    End Sub

    Private Sub ButtonItem6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonItem6.Click
        disp_main2()
    End Sub

    Private Sub ButtonItem8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonItem8.Click
        disp_main5()
    End Sub

    Private Sub CheckBoxX1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
        'If CheckBoxX1.Checked = True Then
        '    DockContainerItem4.Visible = True
        '    disp_main4()
        'Else
        '    DockContainerItem4.Visible = False
        'End If
    End Sub

    Private Sub CheckBoxX2_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBoxX2.CheckedChanged
        If CheckBoxX2.Checked = True Then
            DockContainerItem2.Visible = True
            disp_main2()
        Else
            DockContainerItem2.Visible = False
        End If
    End Sub
End Class