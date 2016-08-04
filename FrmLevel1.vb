
Public Class FrmLevel1

    Private barwidth As Integer

    Private Sub FrmTechnician_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.GotFocus
        Me.TextBoxX1.Focus()
    End Sub

    Private Sub FrmTechnician_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyUp
        If e.Control = True And e.KeyCode = Keys.T Then

            With FpSpread1.ActiveSheet
                .Cells(0, 3).Locked = False
                .Cells(0, 3).ForeColor = Color.OrangeRed
            End With
            '            ButtonItem3_Click(Me, System.EventArgs.Empty)
        End If
    End Sub

    Private Sub FrmTechnician_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Bar3.AutoHide = True
        Bar5.AutoHide = False
        DockContainerItem1.Text = "실행 메뉴"
        DockContainerItem3.Text = "PREINSPECTION 불량"
        DockContainerItem4.Text = "라인 자재 재고"
        DockContainerItem5.Text = "공정 이력"

        WH_NO = "C1000"

        Me.ComboBoxEx8.Items.Add("GOOD")
        Me.ComboBoxEx8.Items.Add("BER")
        Me.ComboBoxEx8.Items.Add("SCRAP")
        Me.ComboBoxEx8.Text = "GOOD"
        Bar4.AutoHide = True


        If Query_Combo(Me.ComboBoxEx9, "EXEC SP_COMMON_FINDWC '" & Site_id & "','R0001','K2000', 'GOOD' ") = True Then
            ComboBoxEx9.Text = ComboBoxEx9.Items(0)
        End If

        With FpSpread1.ActiveSheet
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
            .Cells(1, 4).Text = "LOT 경과일" '"IN WTY"
            .Cells(2, 4).Text = "" '"AGING"

            .Protect = True
            .Cells(0, 0, 2, 5).Locked = True

            Spread_AutoCol(FpSpread1)

        End With

        If Spread_Setting(FpSpread3, "FrmTechnician") = True Then
            FpSpread3.ActiveSheet.RowCount = 0
            FpSpread3.ActiveSheet.Columns(0, 3).Visible = False
            Spread_AutoCol(FpSpread3)
        End If

        If Spread_Setting(FpSpread4, "FrmTechnician") = True Then
            FpSpread4.ActiveSheet.RowCount = 0
            Spread_AutoCol(FpSpread4)
        End If

        If Spread_Setting(FpSpread5, "FrmTechnician") = True Then
            FpSpread5.ActiveSheet.RowCount = 0
            Spread_AutoCol(FpSpread5)
            FpSpread5.Visible = False
        End If

        If Query_Combo(ComboBoxEx3, "sELECT '[' + EMP_NO + ']' + EMP_NM FROM TBL_EMPMASTER WHERE dept_cd = (select dept_cd from tbl_empmaster where emp_no = '" & Emp_No & "') and RETIRE_YN = 'N' ORDER BY EMP_NO") = True Then
            ' MessageBox.Show("site_id : " & ComboBoxItem1.ComboBoxEx.Text)
        Else
            MessageBox.Show("error")

        End If
        ComboBoxEx3.Text = MainFrm.ComboBoxItem1.ComboBoxEx.Text

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
        '       Formbtn_Authority(Me.FindBtn, Me.Name, "FIND")

        IntegerInput1.Text = 0

        Me.TextBoxX1.TabIndex = 0
        Me.TextBoxX1.TabStop = True
        Me.TextBoxX1.Focus()

        ComboBoxEx6.AutoCompleteMode = Windows.Forms.AutoCompleteMode.SuggestAppend
        ComboBoxEx6.AutoCompleteSource = Windows.Forms.AutoCompleteSource.ListItems


    End Sub

    Private Sub TextBoxX1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBoxX1.Click
        Me.TextBoxX1.Text = ""
    End Sub

    Private Sub TextBoxX1_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TextBoxX1.KeyDown
        Dim esn_rs As New ADODB.Recordset

        If e.KeyCode <> Keys.Enter Then
            Exit Sub
        End If

        barwidth = Bar4.Width

        WH_NO = "C1000"

        FpSpread3.ActiveSheet.RowCount = 0
        FpSpread4.ActiveSheet.RowCount = 0

        If LOT_VERIFY(Me.TextBoxX1, e) = True Then

            If Check_Valid_LOT(TextBoxX1.Text, "FrmLevel1") = False Then
                'Me.TextBoxX1.Text = ""
                Me.TextBoxX1.Focus()
                Me.TextBoxX1.SelectAll()
                Exit Sub
            End If

            esn_rs = Query_RS_ALL("EXEC SP_COMMON_GETLOTINFO '" & Site_id & "','" & TextBoxX1.Text & "'")

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
                With FpSpread1.ActiveSheet
                    .Cells(0, 1).Text = esn_rs(0).Value
                    .Cells(1, 1).Text = esn_rs(1).Value
                    .Cells(2, 1).Text = esn_rs(2).Value

                    .Cells(0, 3).Text = esn_rs(3).Value
                    'Dim TRG_ARR As String() = GET_TRG(TextBoxX1.Text, "BOARD")
                    '.Cells(0, 3).Text = TRG_ARR(0) & ":" & TRG_ARR(1) & ":" & TRG_ARR(2) & ":" & TRG_ARR(3) & ":" & TRG_ARR(4)

                    .Cells(1, 3).Text = esn_rs(4).Value
                    .Cells(2, 3).Text = esn_rs(5).Value
                    .Cells(0, 5).Text = esn_rs(6).Value
                    .Cells(1, 5).Text = esn_rs(7).Value
                    .Cells(2, 5).Text = ""
                    Spread_AutoCol(FpSpread1)
                End With

                'If Query_Listview(ListViewEx2, "EXEC SP_FRMTECHNICIAN_GETTRG '" & Site_id & "','" & TextBoxX1.Text & "'", True) = True Then
                'End If

                'PART 인벤토리 조회
                'If Query_Listview(ListViewEX1, "EXEC SP_COMMON_GETPARTINV '" & Site_id & "','" & FpSpread1.ActiveSheet.Cells(1, 1).Text & "','" & "C1000" & "',''", True) = True Then
                '    ListViewEX1.AutoResizeColumn(0, ColumnHeaderAutoResizeStyle.ColumnContent)
                'End If

                Dim JUDGE As String

                If ComboBoxEx8.Text = "GOOD" Then
                    JUDGE = "PASS"
                Else
                    JUDGE = "BER"
                End If

                If Query_Combo(Me.ComboBoxEx9, "EXEC SP_COMMON_FINDWC '" & Site_id & "','" & FpSpread1.ActiveSheet.GetValue(1, 1) & "','K2000', 'GOOD' ") = True Then
                    ComboBoxEx9.Text = ComboBoxEx9.Items(0)
                End If


                If Query_Spread(FpSpread4, "EXEC SP_COMMON_GETLOTHISTORY '" & Site_id & "','" & FpSpread1.ActiveSheet.Cells(0, 1).Text & "'", 1) = True Then
                    Spread_AutoCol(FpSpread4)
                End If

            End If

            'ButtonItem3_Click(sender, e)

            'Me.TextBoxX1.Text = ""
            Me.ComboBoxEx6.Focus()
            'Me.TextBoxX1.SelectAll()

            esn_rs = Nothing
        End If

        Bar4.Width = barwidth

    End Sub

    Private Sub ComboBoxEx7_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBoxEx7.SelectedIndexChanged

        'If Query_Combo(Me.ComboBoxEx6, "SELECT '['+CODE_ID + '] ' + CODE_NAME FROM tbl_REPAIR WHERE  REP_cd = '" & Microsoft.VisualBasic.Left(ComboBoxEx7.Text, 2) & "' and def_cd = '" & FpSpread1.ActiveSheet.Cells(0, 3).Text & "' ORDER BY DEF_CD") = True Then
        'End If
        With FpSpread3.ActiveSheet
            If .RowCount = 0 Then
                Exit Sub
            End If

            .SetValue(0, 2, Mid(ComboBoxEx7.Text, 2, InStr(ComboBoxEx7.Text, "]") - 2))
            .SetValue(0, 3, Mid(ComboBoxEx7.Text, 6, Len(ComboBoxEx7.Text) - 5))
            .Cells(0, 2, 0, 3).BackColor = Color.White

        End With

    End Sub


    Private Sub ListViewEx1_ItemChecked(ByVal sender As Object, ByVal e As System.Windows.Forms.ItemCheckedEventArgs) Handles ListViewEX1.ItemChecked

        If e.Item.Checked = True Then

            If e.Item.SubItems(3).Text = 0 Then
                System.Console.Beep(3000, 400)
                System.Console.Beep(3000, 400)
                Modal_Error("QTY = 0!!")
                Me.ComboBoxEx6.Focus()
                e.Item.Checked = False
                Exit Sub
            End If

            ComboBoxEx6.Items.Clear()
            ComboBoxEx6.Text = ""

            If FpSpread3.ActiveSheet.RowCount > 0 Then

                'If FpSpread3.ActiveSheet.GetValue(0, 0) = "" Then
                '    System.Console.Beep(3000, 400)
                '    System.Console.Beep(3000, 400)
                '    Modal_Error("Select DEFECT!!")
                '    Me.ComboBoxEx6.Focus()
                '    e.Item.Checked = False
                '    Exit Sub
                'End If

                'If FpSpread3.ActiveSheet.GetValue(0, 2) = "" Then
                '    System.Console.Beep(3000, 400)
                '    System.Console.Beep(3000, 400)
                '    Modal_Error("Select Repair!!")
                '    Me.ComboBoxEx7.Focus()
                '    e.Item.Checked = False
                '    Exit Sub
                'End If

            End If

            FpSpread3.ActiveSheet.AddRows(0, 1)


            With FpSpread1.ActiveSheet
                If .GetValue(0, 3) = "UR" Or .GetValue(0, 3) = "GS" Then
                    System.Console.Beep(3000, 400)
                    System.Console.Beep(3000, 400)
                    System.Windows.Forms.MessageBox.Show("NOT GOOD JUDGE!!")
                    FpSpread3.ActiveSheet.RemoveRows(0, 1) 'AA
                    e.Item.Checked = False
                    Exit Sub
                End If

            End With

            If SPREAD_DUP_CHECK(FpSpread3, e.Item.Text, 4) = False Then
                System.Console.Beep(3000, 400)
                System.Console.Beep(3000, 400)
                System.Windows.Forms.MessageBox.Show("이미 선택된 자재입니다.")
                FpSpread3.ActiveSheet.RemoveRows(0, 1) 'AA
                e.Item.Checked = False
                Exit Sub
            End If

            With FpSpread3.ActiveSheet
                '             .AddRows(0, 1)
                .Rows(0).ForeColor = Color.OrangeRed
                If ComboBoxEx6.Text = "" Then
                    .Cells(0, 0, 0, 1).BackColor = Color.Red
                    '                MessageBox.Show("SELECT DEFECT!!")
                Else
                    .SetValue(0, 0, Microsoft.VisualBasic.Left(ComboBoxEx6.Text, 2))
                    .SetValue(0, 1, Microsoft.VisualBasic.Right(ComboBoxEx6.Text, Len(ComboBoxEx6.Text) - 5))
                End If
                'If ComboBoxEx7.Text = "" Then
                '    MessageBox.Show("SELECT REPAIR!!")
                'Else
                '    .SetValue(0, 2, Mid(ComboBoxEx7.Text, 2, InStr(ComboBoxEx7.Text, "]") - 1))
                '    .SetValue(0, 3, Microsoft.VisualBasic.Right(ComboBoxEx7.Text, Len(ComboBoxEx7.Text) - 5))
                'End If
                .SetValue(0, 4, e.Item.Text)
                .SetValue(0, 5, Query_RS("SELECT PART_NAME FROM TBL_PARTMASTER WHERE PART_NO = '" & e.Item.Text & "'"))
                .SetValue(0, 6, Query_RS("SELECT 'L' + CONVERT(VARCHAR(1),PART_LEVEL) FROM TBL_PARTMASTER WHERE PART_NO = '" & e.Item.Text & "'"))

                If FpSpread1.ActiveSheet.Cells(1, 1).Text = "" Then
                    .SetValue(0, 7, 1)
                Else
                    .SetValue(0, 7, Query_RS("SELECT ISNULL(QTY,0)*" & FpSpread1.ActiveSheet.GetValue(2, 3) & " FROM TBL_BOM WHERE SITE_ID = '" & Site_id & "' AND P_NO = '" & FpSpread1.ActiveSheet.Cells(1, 1).Text & "' AND C_NO = '" & e.Item.Text & "'"))
                End If

                .Cells(0, 7).Locked = False

            End With

            Spread_AutoCol(FpSpread3)
            ComboBoxEx6.Focus()
        Else

            '           If Site_id = "S1000" Then
            Dim i As Integer
            i = SPREAD_DUP_ROW(FpSpread3, e.Item.Text, 4)
            If i > -1 Then
                FpSpread3.ActiveSheet.RemoveRows(i, 1)
                ComboBoxEx6.Items.Clear()
                ComboBoxEx7.Items.Clear()
                ComboBoxEx6.Text = ""
                ComboBoxEx7.Text = ""
            End If
            '        End If
        End If

    End Sub

    Private Sub ButtonItem3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonItem3.Click, SaveBtn1.Click
        Try
            Dim T_FLAG As Boolean
            Dim P_FLAG As Boolean

            T_FLAG = False
            P_FLAG = False

            Dim I As Integer

            If FpSpread1.ActiveSheet.Cells(0, 1).Text = "" Then
                System.Windows.Forms.MessageBox.Show("생산 LOT 정보가 없습니다 !!")
                Exit Sub
            End If

            If FpSpread3.ActiveSheet.RowCount < 1 Then
                'Dim r = MessageBox.Show("입력된 자재가 없습니다!! 그대로 저장하시겠습니까?", "저장 확인", MessageBoxButtons.OKCancel)

                'If r = Windows.Forms.DialogResult.Cancel Then
                '    Exit Sub
                'Else
                If Insert_Data("EXEC SP_COMMON_LOTSAVE '" & Site_id & "','" & FpSpread1.ActiveSheet.Cells(0, 1).Text & "','','','N/A',0,'" & ComboBoxEx8.Text & "','" & WH_NO & "','" & FpSpread1.ActiveSheet.Cells(0, 5).Text & "','" & ComboBoxEx9.Text & "','" & OP_No & "','" & WH_NO & "', " & IntegerInput1.Text) = True Then
                End If

                'End If
            End If

            For I = 0 To FpSpread3.ActiveSheet.RowCount - 1

                If FpSpread3.ActiveSheet.Rows(I).ForeColor = Color.OrangeRed Then

                    If Insert_Data("EXEC SP_COMMON_LOTSAVE '" & Site_id & "','" & FpSpread1.ActiveSheet.Cells(0, 1).Text & "','" & FpSpread3.ActiveSheet.Cells(I, 0).Text & "','" & FpSpread3.ActiveSheet.Cells(I, 2).Text & "','" & FpSpread3.ActiveSheet.Cells(I, 4).Text & "'," & FpSpread3.ActiveSheet.Cells(I, 7).Text & ",'" & ComboBoxEx8.Text & "','" & Emp_No & "','" & FpSpread1.ActiveSheet.GetValue(0, 5) & "','" & ComboBoxEx9.Text & "','" & OP_No & "','" & WH_NO & "',0") = True Then
                        FpSpread3.ActiveSheet.Rows(I).ForeColor = Color.Black
                        P_FLAG = True
                    Else

                    End If
                End If
            Next


            If Insert_Data("exec SP_COMMON_WKRESULT '" & Site_id & "','" & FpSpread1.ActiveSheet.GetValue(0, 5) & "','" & OP_No & "','" & FpSpread1.ActiveSheet.Cells(1, 1).Text & "','" & FpSpread1.ActiveSheet.Cells(0, 1).Text & "'") = True Then
            End If
            ListViewEx2.Items.Clear()
            FpSpread4.ActiveSheet.RowCount = 0

            With FpSpread1.ActiveSheet
                .Cells(0, 1).Text = ""
                .Cells(1, 1).Text = ""
                .Cells(2, 1).Text = ""

                .Cells(2, 1).BackColor = .Cells(2, 3).BackColor

                .Cells(0, 3).Text = ""
                .Cells(1, 3).Text = ""
                .Cells(2, 3).Text = ""

                .Cells(0, 5).Text = ""
                .Cells(0, 5).BackColor = .Cells(0, 3).BackColor

                .Cells(1, 5).Text = ""
                .Cells(1, 5).BackColor = .Cells(1, 3).BackColor
                .Cells(2, 5).Text = ""


                Spread_AutoCol(FpSpread1)
            End With


            FpSpread3.ActiveSheet.RowCount = 0

            'If Query_Listview(ListViewEX1, "EXEC SP_COMMON_GETPARTINV '" & Site_id & "','" & FpSpread1.ActiveSheet.Cells(1, 1).Text & "','" & "C1000" & "',''", True) = True Then
            '    ListViewEX1.AutoResizeColumn(0, ColumnHeaderAutoResizeStyle.ColumnContent)
            'End If
            ComboBoxEx6.Items.Clear()
            ComboBoxEx7.Items.Clear()

            MessageBox.Show("저장되었습니다")


            IntegerInput1.Text = 0

            TextBoxX1.Focus()
            TextBoxX1.SelectAll()
        Catch ex As Exception
            MessageBox.Show("Error : " & ex.Message, "Validation Error")
        End Try

    End Sub

    Private Sub Bar5_AutoHideDisplay(ByVal sender As Object, ByVal e As DevComponents.DotNetBar.AutoHideDisplayEventArgs) Handles Bar5.AutoHideDisplay

        If Query_Spread(FpSpread4, "EXEC SP_FRMTECHNICIAN_GETHISTORY '" & Site_id & "','" & FpSpread1.ActiveSheet.Cells(0, 1).Text & "'", 1) = True Then
            Spread_AutoCol(FpSpread4)
        End If

    End Sub


    Function ListViewEx1_ItemChecked1(ByVal e As ListViewItem) As Boolean

        If e.Checked = True Then
            If ComboBoxEx6.Text = "" Or ComboBoxEx7.Text = "" Or ComboBoxEx8.Text = "" Or ComboBoxEx9.Text = "" Then
                System.Console.Beep(3000, 400)
                System.Console.Beep(3000, 400)
                System.Windows.Forms.MessageBox.Show("Select Repair Information!!")

                Exit Function
            End If



            If SPREAD_DUP_CHECK(FpSpread3, e.Text, 4) = False Then
                System.Console.Beep(3000, 400)
                System.Console.Beep(3000, 400)
                Modal_Error("이미 선택된 자재입니다.")
                Exit Function
            End If

            FpSpread3.ActiveSheet.RowCount = FpSpread3.ActiveSheet.RowCount + 1
            FpSpread3.ActiveSheet.Rows(FpSpread3.ActiveSheet.RowCount - 1).ForeColor = Color.OrangeRed
            FpSpread3.ActiveSheet.SetValue(FpSpread3.ActiveSheet.RowCount - 1, 0, Microsoft.VisualBasic.Left(ComboBoxEx6.Text, 2))
            FpSpread3.ActiveSheet.SetValue(FpSpread3.ActiveSheet.RowCount - 1, 1, Microsoft.VisualBasic.Right(ComboBoxEx6.Text, Len(ComboBoxEx6.Text) - 5))
            FpSpread3.ActiveSheet.SetValue(FpSpread3.ActiveSheet.RowCount - 1, 2, Microsoft.VisualBasic.Left(ComboBoxEx7.Text, 2))
            FpSpread3.ActiveSheet.SetValue(FpSpread3.ActiveSheet.RowCount - 1, 3, Microsoft.VisualBasic.Right(ComboBoxEx7.Text, Len(ComboBoxEx7.Text) - 5))
            FpSpread3.ActiveSheet.SetValue(FpSpread3.ActiveSheet.RowCount - 1, 4, e.Text)
            FpSpread3.ActiveSheet.SetValue(FpSpread3.ActiveSheet.RowCount - 1, 5, Query_RS("SELECT PART_NAME FROM TBL_PARTMASTER WHERE PART_NO = '" & e.Text & "'"))
            FpSpread3.ActiveSheet.SetValue(FpSpread3.ActiveSheet.RowCount - 1, 6, Query_RS("SELECT ISNULL(LOC_CD,'') FROM TBL_BOM WHERE SITE_ID = '" & Site_id & "' AND P_NO = '" & FpSpread1.ActiveSheet.Cells(1, 1).Text & "' AND C_NO = '" & e.Text & "'"))
            FpSpread3.ActiveSheet.SetValue(FpSpread3.ActiveSheet.RowCount - 1, 7, Query_RS("SELECT ISNULL(QTY,0) FROM TBL_BOM WHERE SITE_ID = '" & Site_id & "' AND P_NO = '" & FpSpread1.ActiveSheet.Cells(1, 1).Text & "' AND C_NO = '" & e.Text & "'"))

            Spread_AutoCol(FpSpread3)

            Me.ComboBoxEx6.Focus()

        Else
            If Site_id = "S1000" Then
                Dim i As Integer
                i = SPREAD_DUP_ROW(FpSpread3, e.Text, 4)
                If i > -1 Then
                    FpSpread3.ActiveSheet.RemoveRows(i, 1)
                End If
            End If

        End If

        ListViewEx1_ItemChecked1 = True
    End Function

    Private Sub ListViewEx2_ItemSelectionChanged(ByVal sender As Object, ByVal e As System.Windows.Forms.ListViewItemSelectionChangedEventArgs) Handles ListViewEx2.ItemSelectionChanged

        If Query_Combo(Me.ComboBoxEx7, "SELECT REP_CD + ' : ' + REP_NAME FROM tbl_DRMATCHING WHERE  def_cd = '" & ListViewEx2.Items(e.ItemIndex).Text & "' ORDER BY REP_CD") = True Then
            Me.ComboBoxEx6.Items.Clear()
        End If

    End Sub

    Private Sub ButtonItem7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonItem7.Click
        If Query_Combo(Me.ComboBoxEx7, "SELECT REP_CD + ' : ' + REP_NAME FROM tbl_DRMATCHING GROUP BY REP_CD,REP_NAME ORDER BY REP_CD") = True Then
        End If
    End Sub


    Private Sub ButtonItem2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonItem2.Click, NewBtn1.Click

    End Sub

    Private Sub ButtonItem4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonItem4.Click, DelBtn1.Click

    End Sub

    Private Sub ButtonItem5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonItem5.Click, PrtBtn1.Click
        If save_excel = "" Then
            MessageBox.Show("Select Spread for Printing!!")
            Exit Sub
        End If

        If save_excel = "FpSpread1" Then
            If Spread_Print(Me.FpSpread1, "ESN INFO", 0) = False Then
                MsgBox("Fail to Print")
            End If
        ElseIf save_excel = "FpSpread2" Then
'
        ElseIf save_excel = "FpSpread3" Then
            If Spread_Print(Me.FpSpread3, "Current Repair", 0) = False Then
                MsgBox("Fail to Print")
            End If
        ElseIf save_excel = "FpSpread4" Then
            If Spread_Print(Me.FpSpread4, DockContainerItem5.Text, 0) = False Then
                MsgBox("Fail to Print")
            End If
        Else
            MessageBox.Show("No Printing Object!!")
        End If
    End Sub

    Private Sub ButtonItem6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonItem6.Click, XlsBtn1.Click
        If save_excel = "" Then
            MessageBox.Show("Select Spread for saving!!")
            Exit Sub
        End If

        If save_excel = "FpSpread1" Then
            File_Save(SaveFileDialog1, FpSpread1)
        ElseIf save_excel = "FpSpread2" Then
            'File_Save(SaveFileDialog1, FpSpread2)
        ElseIf save_excel = "FpSpread3" Then
            File_Save(SaveFileDialog1, FpSpread3)
        ElseIf save_excel = "FpSpread4" Then
            File_Save(SaveFileDialog1, FpSpread4)
        Else
            MessageBox.Show("Select Spread for Save!!")
        End If

    End Sub

    Private Sub FpSpread2_CellClick(ByVal sender As System.Object, ByVal e As FarPoint.Win.Spread.CellClickEventArgs)
        save_excel = "FpSpread2"
    End Sub

    Private Sub FpSpread1_CellClick(ByVal sender As System.Object, ByVal e As FarPoint.Win.Spread.CellClickEventArgs) Handles FpSpread1.CellClick
        save_excel = "FpSpread1"
    End Sub

    Private Sub FpSpread3_CellClick(ByVal sender As System.Object, ByVal e As FarPoint.Win.Spread.CellClickEventArgs) Handles FpSpread3.CellClick
        save_excel = "FpSpread3"
    End Sub

    Private Sub FpSpread4_CellClick(ByVal sender As System.Object, ByVal e As FarPoint.Win.Spread.CellClickEventArgs) Handles FpSpread4.CellClick
        save_excel = "FpSpread4"
    End Sub

    Private Sub ButtonItem9_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonItem9.Click
        Dim i As Integer

        If ListViewEX1.Items.Count > 0 Then
            If ListViewEX1.SelectedItems.Item(0).Selected = True Then
                If Query_Spread(FpSpread5, "SP_FRMVIEWREPSET_GETITEM '" & Site_id & "','" & ListViewEX1.SelectedItems.Item(0).Text & "'", 1) = True Then

                    For i = 0 To FpSpread5.ActiveSheet.RowCount - 1
                        FpSpread5.ActiveSheet.SetNote(i, 0, FpSpread5.ActiveSheet.GetValue(i, 1))
                        FpSpread5.ActiveSheet.SetNote(i, 2, FpSpread5.ActiveSheet.GetValue(i, 3))
                    Next

                    Spread_AutoCol(FpSpread5)
                End If

                If FpSpread5.ActiveSheet.RowCount > 0 Then
                    FpSpread5.Visible = True
                Else
                    FpSpread5.Visible = False
                    MessageBox.Show("No RepairSet!!")
                End If
            End If
        End If


    End Sub


    Private Sub FpSpread5_CellDoubleClick(ByVal sender As Object, ByVal e As FarPoint.Win.Spread.CellClickEventArgs) Handles FpSpread5.CellDoubleClick

        ComboBoxEx7.Text = FpSpread5.ActiveSheet.GetValue(e.Row, 2) & " : " & FpSpread5.ActiveSheet.GetValue(e.Row, 3)

        ListViewEX1.SelectedItems.Item(0).Checked = True

        FpSpread5.ActiveSheet.RowCount = 0
        FpSpread5.Visible = False

    End Sub

    Private Sub ButtonItem11_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonItem11.Click
        FpSpread3.ActiveSheet.RowCount = 0

    End Sub


    Private Sub ComboBoxEx6_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles ComboBoxEx6.KeyDown
        If e.KeyValue = Keys.Enter Then
            If ComboBoxEx6.FindString(ComboBoxEx6.Text) >= 0 Then
                ComboBoxEx6.Text = ComboBoxEx6.Items(ComboBoxEx6.FindString(ComboBoxEx6.Text))
            End If
        End If
    End Sub


    Private Sub ComboBoxEx6_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBoxEx6.SelectedIndexChanged

        With FpSpread3.ActiveSheet

            If .RowCount = 0 Then
                Exit Sub
            End If

            .SetValue(0, 0, Mid(ComboBoxEx6.Text, 2, InStr(ComboBoxEx6.Text, "]") - 2))
            .SetValue(0, 1, Mid(ComboBoxEx6.Text, InStr(ComboBoxEx6.Text, "]") + 2), Len(ComboBoxEx6.Text) - InStr(ComboBoxEx6.Text, "]") + 2)
            .Cells(0, 0, 0, 1).BackColor = Color.White

        End With

    End Sub

    Private Sub FpSpread1_CellDoubleClick(ByVal sender As Object, ByVal e As FarPoint.Win.Spread.CellClickEventArgs) Handles FpSpread1.CellDoubleClick
        If e.Row = 0 And e.Column = 1 Then
            If FpSpread1.ActiveSheet.GetValue(e.Row, e.Column) <> "" Then
                'FrmTrace_ST2.TextBoxX1.Text = FpSpread1.ActiveSheet.GetValue(e.Row, e.Column)
                'FrmTrace_ST2.TextBoxX1.Focus()
                'FrmTrace_ST2.ShowDialog()
            End If
        End If
    End Sub

    Private Sub FrmTechnician_Shown(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Shown
        Me.TextBoxX1.Focus()
    End Sub

    Private Sub ComboBox3_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBoxEx3.SelectedIndexChanged
        MainFrm.ComboBoxItem1.ComboBoxEx.Text = ComboBoxEx3.Text
        OP_No = Mid(ComboBoxEx3.Text, 2, 5)
    End Sub

    Private Sub ButtonItem14_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonItem14.Click
        If ListViewEX1.FindItemWithText(FpSpread3.ActiveSheet.Cells(FpSpread3.ActiveSheet.ActiveRowIndex, 4).Text) Is Nothing Then
            FpSpread3.ActiveSheet.RemoveRows(FpSpread3.ActiveSheet.ActiveRowIndex, 1)
        Else
            ListViewEX1.FindItemWithText(FpSpread3.ActiveSheet.Cells(FpSpread3.ActiveSheet.ActiveRowIndex, 4).Text).Checked = False
            FpSpread3.ActiveSheet.RemoveRows(FpSpread3.ActiveSheet.ActiveRowIndex, 1)
        End If


    End Sub


    Private Sub ComboBoxEx8_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBoxEx8.SelectedIndexChanged

        Dim JUDGE As String

        If ComboBoxEx8.Text = "GOOD" Then
            JUDGE = "PASS"
        Else
            JUDGE = "BER"
        End If

        If Query_Combo(Me.ComboBoxEx9, "EXEC SP_COMMON_FINDWC '" & Site_id & "','" & FpSpread1.ActiveSheet.GetValue(1, 1) & "','W4000', '" & JUDGE & "' ") = True Then
            ComboBoxEx9.Text = ComboBoxEx9.Items(0)
        End If

    End Sub
End Class