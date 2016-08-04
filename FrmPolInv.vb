Public Class FrmPolInv

    Private Sub FrmBWIP_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        DockContainerItem1.Text = "실행 메뉴"
        DockContainerItem2.Text = "조회 조건"

        GroupPanel1.Dock = DockStyle.Fill

        FpSpread1.ActiveSheet.RowCount = 0


        With FpSpread1.ActiveSheet

            .RowCount = 0
            .ColumnCount = 5
            .ColumnHeader.Columns(0).Label = "모델"
            .ColumnHeader.Columns(1).Label = "레벨 판정전"
            .ColumnHeader.Columns(2).Label = "레벨 1&2"
            .ColumnHeader.Columns(3).Label = "레벨 3"
            .ColumnHeader.Columns(4).Label = "합계"

            .Columns(0).CellType = textcell

            intcell.DecimalPlaces = 0
            .Columns(1, .ColumnCount - 1).CellType = intcell
            .Columns(.ColumnCount - 1).BackColor = Color.Yellow
            Spread_AutoCol(FpSpread1)
            FpSpread1.ActiveSheet.FrozenColumnCount = 1

        End With

        If Query_Combo(Me.ComboBoxEx1, "SELECT model_no FROM TBL_MODELMASTER WHERE site_id = '" & Site_id & "' and active = 'Y' ORDER BY model_no") = True Then
        End If



        Me.ComboBoxEx1.Items.Add("ALL")
        Me.ComboBoxEx1.Text = "ALL"


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
        Formbim_Authority(Me.FindBtn, Me.Name, "FIND")

    End Sub

    Function Refresh_Result() As Boolean
        Dim qry, tb_esn, TB_MOD As String
        Dim I, J As Integer

        tb_esn = "TBL_fESNMASTER_k"
        TB_MOD = "TBL_MODELMASTER"


        qry = "select B.MODEL_NO," & vbNewLine
        qry = qry & "(SELECT SUM(ACT_QTY) FROM  TBL_LOTMASTER  WHERE SITE_ID = B.SITE_ID AND MODEL = B.MODEL_NO AND C_PRC = 'K3000' AND LOT_NO NOT LIKE 'POL%' )," & vbNewLine
        qry = qry & "(SELECT SUM(ACT_QTY) FROM  TBL_LOTMASTER  WHERE SITE_ID = B.SITE_ID AND MODEL = B.MODEL_NO AND C_PRC = 'K3000' AND LOT_NO LIKE 'POL1%' )," & vbNewLine
        qry = qry & "(SELECT SUM(ACT_QTY) FROM  TBL_LOTMASTER  WHERE SITE_ID = B.SITE_ID AND MODEL = B.MODEL_NO AND C_PRC = 'K3000' AND LOT_NO LIKE 'POL2%' )," & vbNewLine
        qry = qry & "(SELECT SUM(ACT_QTY) FROM  TBL_LOTMASTER  WHERE SITE_ID = B.SITE_ID AND MODEL = B.MODEL_NO AND C_PRC = 'K3000' )" & vbNewLine
        qry = qry & "FROM  " & TB_MOD & "  B "
        qry = qry & "WHERE SITE_ID = '" & Site_id & "'" & vbNewLine
        qry = qry & "and b.active = 'Y'" & vbNewLine

        If ComboBoxEx1.Text <> "ALL" Then
            qry = qry & "  and model_NO = '" & ComboBoxEx1.Text & "'" & vbNewLine
        End If

        qry = qry & "ORDER BY SITE_ID, MODEL_NO"

        If Query_Spread(FpSpread1, qry, 1) = True Then

            With FpSpread1.ActiveSheet

                For I = 0 To .RowCount - 1
                    For J = 1 To .ColumnCount - 1
                        If .GetValue(I, J) = 0 Then
                            .SetText(I, J, "")
                        End If

                    Next
                Next

                .RowCount = .RowCount + 1
                .Cells(.RowCount - 1, 1, .RowCount - 1, .ColumnCount - 1).CellType = intcell
                .Rows(.RowCount - 1).BackColor = Color.Yellow
                .SetText(.RowCount - 1, 0, "TOTAL")

                SPREAD_ROW_TOTAL(FpSpread1, 1, .ColumnCount - 1, 1)

                Spread_AutoCol(FpSpread1)
                .Columns(1, .ColumnCount - 1).Width = 80
            End With
        End If

    End Function

    Private Sub FindBtn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FindBtn.Click
        FpSpread1.ActiveSheet.RowCount = 0

        GroupPanel1.Dock = DockStyle.Fill

        Refresh_Result()
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
            If Spread_Print(Me.FpSpread1, "WIP Inventory(Board) Summary", 0) = False Then
                MsgBox("Fail to Print")
            End If
        ElseIf save_excel = "FpSpread2" Then
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
        Else
            MessageBox.Show("Select Spread for Saving!!")
        End If

    End Sub

    Private Sub FpSpread1_CellClick(ByVal sender As System.Object, ByVal e As FarPoint.Win.Spread.CellClickEventArgs) Handles FpSpread1.CellClick
        save_excel = "FpSpread1"
    End Sub

    Private Sub FpSpread2_CellClick(ByVal sender As System.Object, ByVal e As FarPoint.Win.Spread.CellClickEventArgs)
        save_excel = "FpSpread2"
    End Sub

    Private Sub ButtonItem7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonItem7.Click

    End Sub
End Class