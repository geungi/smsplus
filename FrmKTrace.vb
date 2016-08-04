Public Class FrmKTrace

    Private Sub FrmTrace_ST1_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.GotFocus
        TextBoxX1.Focus()
        TextBoxX1.SelectAll()
    End Sub

    Private Sub ButtonItem10_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonItem10.Click

        Dim r As DialogResult = MessageBox.Show("Are you sure to clear JUDGE?", "Clear JUDGE", MessageBoxButtons.YesNo)

        If r = Windows.Forms.DialogResult.Yes Then
            With FpSpread1.ActiveSheet

                If .RowCount = 0 Then
                    Exit Sub
                End If


                Insert_Data("UPDATE TBL_ESNMASTER SET RETURN_DV = NULL WHERE SITE_ID  = '" & Site_id & "' AND OBID = '" & .GetValue(.ActiveRowIndex, 0) & "'")

                ' TextBoxX1_KeyDown(sender, e)
                MessageBox.Show("Complete to clear")
            End With

        End If

    End Sub

    Private Sub FrmTrace_ST1_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Bar3.AutoHide = True
        DockContainerItem1.Text = "실행 메뉴"
        DockContainerItem2.Text = "조회 조건"
        DockContainerItem3.Text = "PREINSPECTION 불량"
        DockContainerItem4.Text = "제품 이미지"

        Bar4.AutoHide = True

        If Spread_Setting(FpSpread1, Me.Name) = True Then
            FpSpread1.ActiveSheet.RowCount = 0
            Spread_AutoCol(FpSpread1)
            FpSpread1.ActiveSheet.FrozenColumnCount = 4
        End If

        If Spread_Setting(FpSpread3, Me.Name) = True Then
            FpSpread3.ActiveSheet.RowCount = 0
            Spread_AutoCol(FpSpread3)
        End If

        With FpSpread2.ActiveSheet
            .ColumnCount = 6
            .RowCount = 3
            .ColumnHeader.Visible = False
            .RowHeader.Visible = False

            .AlternatingRows(0).BackColor = Color.Beige

            .Columns(0).BackColor = Color.Navy
            .Columns(2).BackColor = Color.Navy
            .Columns(4).BackColor = Color.Navy
            '            .Columns(6).BackColor = Color.Navy

            .Columns(0).ForeColor = Color.White
            .Columns(2).ForeColor = Color.White
            .Columns(4).ForeColor = Color.White
            '           .Columns(6).ForeColor = Color.White

            .Columns(0, 5).HorizontalAlignment = FarPoint.Win.Spread.CellHorizontalAlignment.Center
            .Columns(0, 5).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center

            .Cells(0, 0).Text = "생산LOT"
            .Cells(0, 1).CellType = textcell
            .Cells(1, 0).Text = "모델" '"MODEL"
            .Cells(2, 0).Text = "입고번호" '"ACNT"

            .Cells(0, 2).Text = "불량코드" '"TRIAGE"
            .Cells(1, 2).Text = "입고수량" '"DEFECT"
            .Cells(2, 2).Text = "현재수량" '"REPAIR"

            .Cells(0, 4).Text = "현공정" '"QARJT"
            .Cells(1, 4).Text = "출하번호" '"IN WTY"
            .Cells(2, 4).Text = "인보이스번호" '"AGING"

            .Protect = True
            .Cells(0, 0, 2, 5).Locked = True

            Spread_AutoCol(FpSpread2)
        End With

        FpSpread1.Font = New Font("Verdana", 8, FontStyle.Regular)
        FpSpread2.Font = New Font("Verdana", 8, FontStyle.Regular)

        Me.ComboBoxEx1.Items.AddRange(New String() {"생산LOT", "플립ID"})
        Me.ComboBoxEx1.SelectedIndex = 0

        Formbim_Authority(Me.ButtonItem2, Me.Name, "NEW")
        Formbim_Authority(Me.ButtonItem3, Me.Name, "SAVE")
        Formbim_Authority(Me.ButtonItem4, Me.Name, "DELETE")
        Formbim_Authority(Me.ButtonItem5, Me.Name, "PRINT")
        Formbim_Authority(Me.ButtonItem6, Me.Name, "EXCEL")
        Formbim_Authority(Me.FindBtn, Me.Name, "FIND")

        ButtonItem2.Enabled = False
        ButtonItem3.Enabled = False
        ButtonItem4.Enabled = False

    End Sub

    Private Sub TextBoxX1_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TextBoxX1.KeyDown

        Dim esn_rs As New ADODB.Recordset

        If e.KeyCode <> Keys.Enter Then
            Exit Sub
        End If

        FpSpread1.ActiveSheet.RowCount = 0
        'FpSpread2.ActiveSheet.RowCount = 0
        FpSpread3.ActiveSheet.RowCount = 0

        If LOT_VERIFY(Me.TextBoxX1, e) = True Then

            If Query_Spread(FpSpread1, "SELECT LOT_NO, LOT_NO, MODEL, CASE LEFT(LOT_NO,1) WHEN 'R' THEN 'Y' ELSE 'N' END, ISNULL(C_DATE,''), PSHIP_NO, CSHIP_NO, ISNULL(RETURN_DV,'') FROM TBL_LOTMASTER WHERE LOT_NO = '" & TextBoxX1.Text & "' ORDER BY C_DATE DESC", 1) = True Then
                Spread_AutoCol(FpSpread1)

                If FpSpread1.ActiveSheet.RowCount > 0 Then
                    FpSpread1_CellClick1(1, 0)
                End If
            End If
        End If

    End Sub

    Private Sub FpSpread1_CellClick(ByVal sender As Object, ByVal e As FarPoint.Win.Spread.CellClickEventArgs) Handles FpSpread1.CellClick
        '        Dim i As Integer

        save_excel = "FpSpread1"

        FpSpread1_CellClick1(0, e.Row)


        ' ExpandableSplitter1.SplitPosition = My.Computer.Screen.WorkingArea.Width - (FpSpread4.ActiveSheet.Columns(0).Width + FpSpread4.ActiveSheet.Columns(1).Width + FpSpread4.ActiveSheet.Columns(2).Width + FpSpread4.ActiveSheet.Columns(3).Width + FpSpread4.ActiveSheet.Columns(4).Width + FpSpread4.ActiveSheet.Columns(5).Width + FpSpread4.ActiveSheet.Columns(6).Width + FpSpread4.ActiveSheet.Columns(7).Width) - 90

        TextBoxX1.Focus()
        Me.TextBoxX1.SelectAll()

    End Sub

    Private Sub FpSpread1_CellClick1(ByVal COL As Integer, ByVal ROW As Integer)

        Dim esn_rs As ADODB.Recordset

        FpSpread3.ActiveSheet.RowCount = 0

        save_excel = "FpSpread1"

        If ROW >= 0 Then

            esn_rs = Query_RS_ALL("EXEC SP_FRMKTRACE_GETLOTINFO '" & Site_id & "','" & FpSpread1.ActiveSheet.GetValue(ROW, 1) & "'")

            If esn_rs(0).Value = "NOT EXIST" Then
                System.Console.Beep(3000, 400)
                System.Console.Beep(5000, 400)
                System.Windows.Forms.MessageBox.Show("NO EXIST ESN!!")
                esn_rs = Nothing

                'Me.TextBoxX1.Text = ""
                Me.TextBoxX1.Focus()
                Me.TextBoxX1.SelectAll()
                Exit Sub
            Else
                With FpSpread2.ActiveSheet
                    .Cells(0, 1).Text = esn_rs(0).Value
                    .Cells(1, 1).Text = esn_rs(1).Value
                    .Cells(2, 1).Text = esn_rs(2).Value

                    .Cells(0, 3).Text = esn_rs(3).Value
                    Dim TRG_ARR As String() = GET_TRG(TextBoxX1.Text, "FLIP")
                    .Cells(0, 3).Text = TRG_ARR(0) & ":" & TRG_ARR(1) & ":" & TRG_ARR(2) & ":" & TRG_ARR(3) & ":" & TRG_ARR(4)

                    .Cells(1, 3).Text = esn_rs(4).Value
                    .Cells(2, 3).Text = esn_rs(5).Value
                    .Cells(0, 5).Text = esn_rs(6).Value
                    .Cells(1, 5).Text = esn_rs(7).Value
                    .Cells(2, 5).Text = esn_rs(8).Value
                    Spread_AutoCol(FpSpread2)
                End With
            End If
            esn_rs = Nothing

            If Query_Listview(ListViewEX1, "EXEC SP_FRMKTRACE_GETTRG '" & Site_id & "','" & FpSpread1.ActiveSheet.GetValue(ROW, 0) & "'", True) = True Then
                ListViewEX1.AutoResizeColumn(0, ColumnHeaderAutoResizeStyle.ColumnContent)
            End If

            If Query_Spread(FpSpread3, "SP_FRMKTRACE_GETLOTHISTORY '" & Site_id & "','" & FpSpread1.ActiveSheet.GetValue(ROW, 1) & "'", 1) = True Then
                Spread_AutoCol(FpSpread3)
            End If

        End If



    End Sub
    Private Sub ButtonItem2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub ButtonItem3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

        'If ComboBoxEx1.Text <> "BOARD" Then
        '    MessageBox.Show("No Board")
        '    Exit Sub
        'End If

        With FpSpread1.ActiveSheet
            If .Cells(.ActiveRowIndex, 11).ForeColor = Color.OrangeRed Then
                .ActiveColumnIndex = .ActiveRowIndex - 1

            End If

            .Cells(.ActiveRowIndex, 11).ForeColor = Color.Black
            .Cells(.ActiveRowIndex, 11).Locked = True

        End With


    End Sub

    Private Sub ButtonItem4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub ButtonItem5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        If save_excel = "" Then
            MessageBox.Show("Select Spread for Printing!!")
            Exit Sub
        End If

        If save_excel = "FpSpread1" Then
            If Spread_Print(Me.FpSpread1, DockContainerItem2.Text, 0) = False Then
                MsgBox("Fail to Print")
            End If
        ElseIf save_excel = "FpSpread2" Then
        ElseIf save_excel = "FpSpread3" Then
            If Spread_Print(Me.FpSpread3, DockContainerItem4.Text, 0) = False Then
                MsgBox("Fail to Print")
            End If
        ElseIf save_excel = "FpSpread4" Then
        ElseIf save_excel = "FpSpread5" Then

        Else
            MessageBox.Show("No Printing Object!!")
        End If
    End Sub

    Private Sub ButtonItem6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        If save_excel = "" Then
            MessageBox.Show("Select Spread for saving!!")
            Exit Sub
        End If

        If save_excel = "FpSpread1" Then
            File_Save(SaveFileDialog1, FpSpread1)
        ElseIf save_excel = "FpSpread2" Then
            File_Save(SaveFileDialog1, FpSpread2)
        ElseIf save_excel = "FpSpread3" Then
            File_Save(SaveFileDialog1, FpSpread3)
        ElseIf save_excel = "FpSpread4" Then
        ElseIf save_excel = "FpSpread5" Then

        Else
            MessageBox.Show("Select Spread for Save!!")
        End If
    End Sub

    Private Sub FpSpread2_CellClick(ByVal sender As System.Object, ByVal e As FarPoint.Win.Spread.CellClickEventArgs)
        save_excel = "FpSpread2"

        'PictureBox1.Image = Nothing

        'DISPLAY_IMG(FpSpread2.ActiveSheet.GetValue(e.Row, 0), PictureBox1, 1)

        'TextBoxX1.Focus()
        'Me.TextBoxX1.SelectAll()

    End Sub
    Private Sub FpSpread3_CellClick(ByVal sender As System.Object, ByVal e As FarPoint.Win.Spread.CellClickEventArgs)
        save_excel = "FpSpread3"
        TextBoxX1.Focus()
        Me.TextBoxX1.SelectAll()

    End Sub


    Private Sub ComboBoxEx1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Me.TextBoxX1.Focus()
        Me.TextBoxX1.SelectAll()
    End Sub

    Private Sub FrmTrace_MouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles MyBase.MouseClick
        Me.TextBoxX1.Focus()
        Me.TextBoxX1.SelectAll()
    End Sub

    Private Sub FrmTrace_Shown(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Shown
        Me.TextBoxX1.Focus()
    End Sub



    Private Sub FpSpread2_SelectionChanged(ByVal sender As Object, ByVal e As FarPoint.Win.Spread.SelectionChangedEventArgs)
        Me.TextBoxX1.Focus()
        Me.TextBoxX1.SelectAll()
    End Sub

    Private Sub FpSpread1_SelectionChanged(ByVal sender As Object, ByVal e As FarPoint.Win.Spread.SelectionChangedEventArgs)
        Me.TextBoxX1.Focus()
        Me.TextBoxX1.SelectAll()
    End Sub

    Private Sub FpSpread3_SelectionChanged(ByVal sender As Object, ByVal e As FarPoint.Win.Spread.SelectionChangedEventArgs)
        Me.TextBoxX1.Focus()
        Me.TextBoxX1.SelectAll()
    End Sub

    Private Sub FpSpread5_SelectionChanged(ByVal sender As Object, ByVal e As FarPoint.Win.Spread.SelectionChangedEventArgs)
        Me.TextBoxX1.Focus()
        Me.TextBoxX1.SelectAll()
    End Sub

    Private Sub FpSpread4_SelectionChanged(ByVal sender As Object, ByVal e As FarPoint.Win.Spread.SelectionChangedEventArgs)
        Me.TextBoxX1.Focus()
        Me.TextBoxX1.SelectAll()
    End Sub

    Private Sub ButtonItem7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonItem7.Click

        'Dim r As DialogResult = MessageBox.Show("Are you sure to clear ICC ID?", "Clear ICC", MessageBoxButtons.YesNo)

        'If r = Windows.Forms.DialogResult.Yes Then
        '    With FpSpread4.ActiveSheet

        '        Dim obid As String = ""
        '        obid = Query_RS("select obid from tbl_esnmaster where esn = '" & TextBoxX1.Text & "'")

        '        If obid = "" Then
        '            System.Console.Beep(3000, 400)
        '            System.Console.Beep(3000, 400)
        '            Modal_Error(TextBoxX3.Text & vbNewLine & "NO EXIST MEID IN WIP!!")
        '            Exit Sub
        '        End If

        '        If .GetValue(3, 7) <> "" Then
        '            If Insert_Data("update tbl_icc set esn = null, reserv1 = NULL where esn = '" & .GetValue(0, 1) & "' and iccno = '" & .GetValue(3, 7) & "'") = True Then
        '                If Insert_Data("update tbl_meiddet set reserv1 = null where deviceSerialNumberHex = '" & .GetValue(0, 1) & "' and obid in (select obid from tbl_esnmaster)") = True Then
        '                End If
        '            End If
        '        End If


        '    End With

        '    MessageBox.Show("Complete to clear")
        'End If

    End Sub

    Private Sub TextBoxX1_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs)

    End Sub



    Private Sub PictureBox1_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs)


        'If PictureBox1.Image Is Nothing Then
        '    MessageBox.Show("No Image!")
        '    Exit Sub
        'End If

        'FrmImg.obid = FpSpread2.ActiveSheet.GetValue(FpSpread2.ActiveSheet.ActiveRowIndex, 0)
        'FrmImg.seq = 1
        'FrmImg.esn_obid = FpSpread1.ActiveSheet.GetValue(FpSpread1.ActiveSheet.ActiveRowIndex, 0)
        'FrmImg.LabelX1.Text = FpSpread2.ActiveSheet.GetValue(FpSpread2.ActiveSheet.ActiveRowIndex, 2)

        'FrmImg.StepIndicator1.StepCount = 4
        'FrmImg.StepIndicator1.CurrentStep = 1
        'FrmImg.ShowDialog()



    End Sub

    Private Sub ButtonItem9_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonItem9.Click

        Dim r As DialogResult = MessageBox.Show("Are you sure to delete Image?", "Delete Image", MessageBoxButtons.YesNo)

        If r = Windows.Forms.DialogResult.Yes Then

            With FpSpread2.ActiveSheet

                If Insert_Data("delete from emsweb.emsweb1.dbo.tbl_img where src_obid  = '" & .GetValue(.ActiveRowIndex, 0) & "'") = True Then
                    MessageBox.Show("Complete to Image")
                End If

            End With


        End If


    End Sub


End Class