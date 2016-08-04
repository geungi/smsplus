Public Class FrmQCFail

    Private Sub FrmQCLine_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.GotFocus
         Me.ComboBoxEx6.Focus()

        SCAN_FIRST()
    End Sub

    Private Sub FrmQCLine_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown

        If e.KeyCode = Keys.Enter Then
            SCAN_FIRST()
        End If

    End Sub

    Private Sub ComboBox7_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBoxEx7.SelectedIndexChanged
        MainFrm.ComboBoxItem1.ComboBoxEx.Text = ComboBoxEx7.Text
        OP_No = Mid(ComboBoxEx7.Text, 2, 5)
    End Sub

    Private Sub FrmQCLine_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim fp_rs As New ADODB.Recordset
        Dim i, j As Integer

        DockContainerItem1.Text = "실행 메뉴"
        DockContainerItem2.Text = "입력 조건"
        DockContainerItem3.Text = "불량 코드"
        DockContainerItem4.Text = "스캔 리스트"

        If Query_Combo(Me.ComboBoxEx1, "SELECT model_no FROM tbl_modelmaster WHERE site_id = '" & Site_id & "' and active = 'Y' ORDER BY model_no") = True Then
        End If
        Me.ComboBoxEx1.Items.Add("ALL")
        Me.ComboBoxEx1.Text = "ALL"

        If Query_Combo(Me.ComboBoxEx2, "SELECT CODE_NAME FROM tbl_CODEmaster WHERE site_id = '" & Site_id & "' and CLASS_ID = '10021' AND active = 'Y' ORDER BY DIS_ORDER") = True Then
        End If
        Me.ComboBoxEx2.Text = "N/A"

        If Query_Combo(Me.ComboBoxEx3, "SELECT CODE_ID + ' : ' + CODE_NAME FROM tbl_CODEmaster WHERE site_id = '" & Site_id & "' and CLASS_ID = '10005' AND active = 'Y' ORDER BY DIS_ORDER") = True Then
        End If
        Me.ComboBoxEx3.Text = "N/A : N/A"

        If Query_Combo(Me.ComboBoxEx4, "SELECT CODE_ID FROM tbl_CODEmaster WHERE site_id = '" & Site_id & "' and CLASS_ID = '10006' AND active = 'Y' ORDER BY DIS_ORDER") = True Then
        End If
        Me.ComboBoxEx4.Text = "N/A"

        If Query_Combo2(Me.ComboBoxEx5, "TBL_CODEMASTER", "CODE_NAME", "CODE_ID", "site_id = '" & Site_id & "' and CLASS_ID = 'R0001' AND active = 'Y' ORDER BY CODE_ID", False) = True Then
            Me.ComboBoxEx5.SelectedIndex = 0
        End If

        If Query_Combo(ComboBoxEx7, "sELECT '[' + EMP_NO + ']' + EMP_NM FROM TBL_EMPMASTER WHERE dept_cd = (select dept_cd from tbl_empmaster where emp_no = '" & Emp_No & "') and RETIRE_YN = 'N' ORDER BY EMP_NO") = True Then
            ' MessageBox.Show("site_id : " & ComboBoxItem1.ComboBoxEx.Text)
        Else
            MessageBox.Show("error")

        End If
        ComboBoxEx7.Text = MainFrm.ComboBoxItem1.ComboBoxEx.Text

        ComboBoxEx6.Items.Add("GIB")
        ComboBoxEx6.Items.Add("BER")
        ComboBoxEx6.Text = "GIB"

        fp_rs = Query_RS_ALL("select distinct left(code_id,1) from tbl_defect WHERE ACTIVE = 'Y'  order by left(code_id,1)")


        With FpSpread1.ActiveSheet
            .ColumnCount = fp_rs.RecordCount
            .RowCount = 0


            For i = 0 To fp_rs.RecordCount - 1

                Dim HD As String = ""
                If fp_rs(0).Value = "B" Then
                    HD = "BLU"
                ElseIf fp_rs(0).Value = "L" Then
                    HD = "LCD"
                ElseIf fp_rs(0).Value = "N" Then
                    HD = "NODEF"
                ElseIf fp_rs(0).Value = "T" Then
                    HD = "TOUCH"
                ElseIf fp_rs(0).Value = "F" Then
                    HD = "FPCB"
                ElseIf fp_rs(0).Value = "P" Then
                    HD = "POL"
                ElseIf fp_rs(0).Value = "H" Then
                    HD = "HOUSING"
                ElseIf fp_rs(0).Value = "O" Then
                    HD = "OCA"
                End If


                .ColumnHeader.Columns(i).Label = HD
                fp_rs.MoveNext()
            Next

            fp_rs = Nothing

            For i = 0 To .ColumnCount - 1
                fp_rs = Query_RS_ALL("select code_id, code_name from tbl_defect where site_id = '" & Site_id & "' and code_id like '" & Mid(.ColumnHeader.Columns(i).Label, 1, 1) & "%' order by code_id")

                If fp_rs.RecordCount > .RowCount Then
                    .RowCount = fp_rs.RecordCount
                End If

                For j = 0 To fp_rs.RecordCount - 1
                    Dim btncell As New FarPoint.Win.Spread.CellType.ButtonCellType()

                    .Cells(j, i).CellType = btncell

                    btncell.Text = fp_rs(0).Value & vbNewLine & fp_rs(1).Value

                    .SetValue(j, i, fp_rs(0).Value)
                    fp_rs.MoveNext()
                Next

                fp_rs = Nothing
            Next


            .Columns(0, .ColumnCount - 1).Width = 60
            .Rows(0, .RowCount - 1).Height = 40
            '.Cells(0, 0, .RowCount - 1, .ColumnCount - 1).Locked = True
            .Protect = True
            Spread_AutoCol(FpSpread1)
        End With


        If Spread_Setting(FpSpread4, "FrmTriage") = True Then
            With FpSpread4.ActiveSheet
                .RowCount = 0
                .ColumnCount = .ColumnCount + 1
                .ColumnHeader.Columns(.ColumnCount - 1).Label = "불량수량"
                .Columns(.ColumnCount - 1).CellType = intcell

            End With
            Spread_AutoCol(FpSpread4)
        End If

        If Spread_Setting(FpSpread2, "FrmQCLine") = True Then
            With FpSpread2.ActiveSheet
                .RowCount = 0
            End With
            Spread_AutoCol(FpSpread2)
        End If

        Formbim_Authority(Me.NewBtn, Me.Name, "NEW")
        Formbim_Authority(Me.ButtonItem3, Me.Name, "SAVE")
        Formbim_Authority(Me.ButtonItem4, Me.Name, "DELETE")
        Formbim_Authority(Me.ButtonItem5, Me.Name, "PRINT")
        Formbim_Authority(Me.ButtonItem6, Me.Name, "EXCEL")
        '        Formbtn_Authority(Me.FindBtn, Me.Name, "FIND")

        NewBtn.Enabled = False
        ButtonItem4.Enabled = False
        ButtonItem3.Enabled = False

        TextBoxX2.Enabled = True
        'TextBoxX1.Focus()
        'TextBoxX1.SelectAll()
        TextBoxX2.Focus()
        TextBoxX2.SelectAll()



    End Sub

    Private Sub FpSpread1_ButtonClicked(ByVal sender As Object, ByVal e As FarPoint.Win.Spread.EditorNotifyEventArgs) Handles FpSpread1.ButtonClicked

        Dim i, j As Integer
        Dim b As New FarPoint.Win.Spread.CellType.ButtonCellType()
        Dim fp_rs As ADODB.Recordset

        If FpSpread4.ActiveSheet.RowCount > 4 Then
            Modal_Error("ALREADY 5 DEFECT SELECTED!")
            TextBoxX2.Text = ""
            TextBoxX2.Focus()
            Exit Sub
        End If


        fp_rs = Query_RS_ALL("select distinct left(CODE_ID,1) from tbl_defECT  order by left(CODE_ID,1)")

        With FpSpread1.ActiveSheet
            .ColumnCount = fp_rs.RecordCount
            .RowCount = 0


            For i = 0 To fp_rs.RecordCount - 1
                .ColumnHeader.Columns(i).Label = fp_rs(0).Value
                fp_rs.MoveNext()
            Next

            fp_rs = Nothing

            For i = 0 To .ColumnCount - 1
                fp_rs = Query_RS_ALL("select code_id, code_name from tbl_defect where site_id = '" & Site_id & "' and code_id like '" & .ColumnHeader.Columns(i).Label & "%' order by code_id")

                If fp_rs.RecordCount > .RowCount Then
                    .RowCount = fp_rs.RecordCount
                End If

                For j = 0 To fp_rs.RecordCount - 1
                    Dim btncell As New FarPoint.Win.Spread.CellType.ButtonCellType()

                    .Cells(j, i).CellType = btncell

                    btncell.Text = fp_rs(0).Value & vbNewLine & fp_rs(1).Value

                    .SetValue(j, i, fp_rs(0).Value)
                    fp_rs.MoveNext()
                Next

                fp_rs = Nothing
            Next


            .Columns(0, .ColumnCount - 1).Width = 80
            .Rows(0, .RowCount - 1).Height = 45
            '.Cells(0, 0, .RowCount - 1, .ColumnCount - 1).Locked = True
            .Protect = True
            Spread_AutoCol(FpSpread1)
        End With


        b = FpSpread1.ActiveSheet.Cells(e.Row, e.Column).CellType
        b.BackgroundStyle = FarPoint.Win.BackStyle.Gradient
        b.ButtonColor = System.Drawing.Color.Red
        b.ButtonColor2 = System.Drawing.Color.Red
        b.GradientMode = System.Drawing.Drawing2D.LinearGradientMode.Vertical
        b.UseVisualStyleBackColor = False


        If SPREAD_DUP_CHECK(FpSpread4, b.Text, 0) = True Then
            '            FpSpread4.ActiveSheet.RowCount = 0
            FpSpread4.ActiveSheet.RowCount = FpSpread4.ActiveSheet.RowCount + 1
            FpSpread4.ActiveSheet.SetValue(FpSpread4.ActiveSheet.RowCount - 1, 0, FpSpread1.ActiveSheet.GetValue(e.Row, e.Column))
            FpSpread4.ActiveSheet.SetValue(FpSpread4.ActiveSheet.RowCount - 1, 1, Query_RS("select code_name from tbl_defect where code_id = '" & FpSpread4.ActiveSheet.GetValue(FpSpread4.ActiveSheet.RowCount - 1, 0) & "'"))

            Spread_AutoCol(FpSpread4)
            IntegerInput1.Focus()

            FpSpread1.ActiveSheet.SetActiveCell(e.Row, e.Column)

        End If
    End Sub

    Private Sub FpSpread4_CellClick(ByVal sender As Object, ByVal e As FarPoint.Win.Spread.CellClickEventArgs) Handles FpSpread4.CellClick
        save_excel = "FpSpread4"
        ButtonItem4.Enabled = True
        'ButtonItem4.Visible = True
    End Sub
    Private Sub FpSpread4_CellDoubleClick(ByVal sender As Object, ByVal e As FarPoint.Win.Spread.CellClickEventArgs) Handles FpSpread4.CellDoubleClick

        With FpSpread4.ActiveSheet
            If e.Column = .ColumnCount - 1 Then
                .Cells(e.Row, .ColumnCount - 1).Locked = False
            End If
        End With
    End Sub
    Private Sub FpSpread4_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles FpSpread4.Leave
        ButtonItem4.Enabled = False
    End Sub


    Private Sub TextBoxX3_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TextBoxX3.KeyDown
        If e.KeyCode <> Keys.Enter Then
            Exit Sub
        End If

        If TextBoxX3.Text = "0ZZ" Then
            SCAN_FIRST()
            Exit Sub
        End If

        'If FpSpread4.ActiveSheet.RowCount = 0 Then
        '    System.Console.Beep(3000, 400)
        '    System.Console.Beep(3000, 400)
        '    Modal_Error("SELECT DEFECT!!")
        '    Exit Sub
        'End If

        Dim qc_rs As ADODB.Recordset

        If LOT_VERIFY(Me.TextBoxX3, e) = True Then
            'Validation Process
            If Check_Valid_LOT(TextBoxX3.Text, Me.Name) = False Then
                TextBoxX3.Focus()
                TextBoxX3.SelectAll()
                '            Modal_Error("Wrong Flip")
                Exit Sub
            End If

            '수량체크
            Dim CNT As Integer = 0
            With FpSpread4.ActiveSheet
                For I As Integer = 0 To .RowCount - 1
                    If .GetValue(I, .ColumnCount - 1) = 0 Then
                        Modal_Error("불량수량을 입력하십시오.")
                        Exit Sub
                    Else
                        CNT = CNT + .GetValue(I, .ColumnCount - 1)
                    End If
                Next
            End With

            If CInt(Query_RS("select act_qty from tbl_lotmaster where lot_no = '" & TextBoxX3.Text & "'")) < CNT Then
                System.Console.Beep(3000, 400)
                System.Console.Beep(3000, 400)
                '               System.Windows.Forms.MessageBox.Show("Aleady Received!!")
                Modal_Error(TextBoxX3.Text & vbNewLine & "생산LOT의 현재 수량보다 많습니다!!")
                IntegerInput1.Focus()
                Exit Sub
            End If

            qc_rs = Query_RS_ALL("select model, C_PRC, init_qty, act_qty from tbl_lotmaster where site_id = '" & Site_id & "' and lot_no = '" & TextBoxX3.Text & "'")

            If qc_rs(3).Value = 0 Then
                System.Console.Beep(3000, 400)
                System.Console.Beep(3000, 400)
                '               System.Windows.Forms.MessageBox.Show("Aleady Received!!")
                Modal_Error(TextBoxX3.Text & vbNewLine & "생산수량이 0 인 생산LOT입니다!!")
                'Me.TextBoxX1.Text = ""
                Me.TextBoxX3.Focus()
                TextBoxX3.SelectAll()
                Exit Sub

            End If


            With FpSpread2.ActiveSheet
                For I As Integer = 0 To FpSpread4.ActiveSheet.RowCount - 1
                    If Insert_Data("EXEC SP_COMMON_LOTSAVE '" & Site_id & "','" & TextBoxX3.Text & "','" & FpSpread4.ActiveSheet.GetValue(I, 0) & "','','N/A',0,'2차 검사 불량','" & Emp_No & "','LCD FUNCTION TEST(2)','" & ComboBoxEx6.Text & "','" & OP_No & "','C1000', " & FpSpread4.ActiveSheet.GetValue(I, FpSpread4.ActiveSheet.ColumnCount - 1)) = True Then
                    End If

                    If Insert_Data("exec SP_COMMON_WKRESULT_FAIL '" & Site_id & "','LCD FUNCTION TEST(2)','" & OP_No & "','" & qc_rs(0).Value & "','" & TextBoxX3.Text & "'," & FpSpread4.ActiveSheet.GetValue(I, FpSpread4.ActiveSheet.ColumnCount - 1)) = True Then
                    End If

                    .AddRows(0, 1)
                    .SetValue(0, 0, TextBoxX3.Text)
                    .SetValue(0, 2, FpSpread4.ActiveSheet.GetValue(I, 0))
                    .SetValue(0, 1, qc_rs(0).Value)
                    .SetValue(0, 3, qc_rs(3).Value)
                    .SetValue(0, 4, FpSpread4.ActiveSheet.GetValue(I, FpSpread4.ActiveSheet.ColumnCount - 1))

                Next


                'If Insert_Data("EXEC SP_COMMON_LOTSAVE '" & Site_id & "','" & TextBoxX3.Text & "','" & FpSpread4.ActiveSheet.GetValue(0, 0) & "','','N/A',0,'2차 검사 불량','" & Emp_No & "','LCD FUNCTION TEST(2)','" & TextBoxX5.Text & "','" & OP_No & "','C1000', " & IntegerInput1.Text) = True Then
                'End If

                'If Insert_Data("exec SP_COMMON_WKRESULT_FAIL '" & Site_id & "','LCD FUNCTION TEST(2)','" & OP_No & "','" & qc_rs(0).Value & "','" & TextBoxX3.Text & "'," & IntegerInput1.Text) = True Then
                'End If


                '.AddRows(0, 1)
                '.SetValue(0, 0, TextBoxX3.Text)
                '.SetValue(0, 2, FpSpread4.ActiveSheet.GetValue(0, 0))
                '.SetValue(0, 1, qc_rs(0).Value)
                '.SetValue(0, 3, qc_rs(3).Value)
                '.SetValue(0, 4, IntegerInput1.Text)
                Spread_AutoCol(FpSpread2)

                SCAN_FIRST()
                TextBoxX2.Focus()
                TextBoxX2.SelectAll()

            End With
        End If
    End Sub

    Private Sub TextBoxX4_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TextBoxX4.KeyDown
        If e.KeyCode <> Keys.Enter Then
            Exit Sub
        End If

        If TextBoxX4.Text = "0ZZ" Then
            SCAN_FIRST()
            Exit Sub
        Else

            TextBoxX5.Text = Query_RS("SELECT code_name FROM TBL_codeMASTER WHERE class_id = 'R0001' and code_id = '" & TextBoxX4.Text & "'")

            If TextBoxX5.Text = "" Then
                Modal_Error("스캔하신 공정이 존재하지 않습니다.")
                TextBoxX4.Focus()
                TextBoxX4.SelectAll()
                Exit Sub
            End If
        End If

        IntegerInput1.Focus()
        Exit Sub

    End Sub

    Private Sub TextBoxX2_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TextBoxX2.KeyDown

        If e.KeyCode = Keys.Enter Then
            'defect code : 2 DIGIT

            If Len(Me.TextBoxX2.Text) = 3 Or Len(Me.TextBoxX2.Text) = 4 Then
                If SPREAD_DUP_CHECK(FpSpread4, TextBoxX2.Text, 0) = True Then

                    If TextBoxX2.Text = "0YY" Then
                        If FpSpread4.ActiveSheet.RowCount = 0 Then
                            Modal_Error("불량코드를 선택하십시오!")
                            TextBoxX3.Focus()
                            TextBoxX3.SelectAll()
                        End If

                        TextBoxX2.Text = ""
                        TextBoxX4.Focus()
                        TextBoxX4.SelectAll()
                        Exit Sub
                    ElseIf TextBoxX2.Text = "0ZZ" Then
                        SCAN_FIRST()
                        Exit Sub
                    End If

                    With FpSpread4.ActiveSheet

                        'Defect Code 검사
                        If Query_RS("select count(CODE_ID) from tbl_defECT  WHERE CODE_ID = '" & TextBoxX2.Text & "'") = 0 Then
                            Modal_Error("NO Exact Defect Code")
                            Exit Sub
                        End If

                        Dim DEF As String = Query_RS("select code_name from tbl_defect where code_id = '" & TextBoxX2.Text & "'")

                        If DEF = "" Then
                            Modal_Error("WRONG DEFECT CODE!")
                            TextBoxX2.Text = ""
                            TextBoxX2.Focus()
                            Exit Sub
                        End If

                        If .RowCount > 4 Then
                            Modal_Error("ALREADY 5 DEFECT SELECTED!")
                            TextBoxX2.Text = ""
                            TextBoxX2.Focus()
                            Exit Sub
                        End If

                        .RowCount = .RowCount + 1
                        .SetValue(.RowCount - 1, 0, TextBoxX2.Text)
                        .SetValue(.RowCount - 1, 1, DEF)
                    End With

                    Spread_AutoCol(FpSpread4)
                    TextBoxX2.Text = ""

                Else
                    Modal_Error("ALREADY SELECTED DEFECT CODE!")
                    TextBoxX2.Text = ""
                    TextBoxX2.Focus()
                    Exit Sub
                End If
            End If
        Else
        End If
    End Sub

    Private Sub ButtonItem4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonItem4.Click
        FpSpread4.ActiveSheet.RemoveRows(FpSpread4.ActiveSheet.ActiveRowIndex, 1)
    End Sub

    Private Sub NewBtn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NewBtn.Click

    End Sub

    Private Sub ButtonItem3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonItem3.Click

    End Sub

    Private Sub ButtonItem5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonItem5.Click
        If save_excel = "" Then
            MessageBox.Show("Select Spread for Printing!!")
            Exit Sub
        End If

        If save_excel = "FpSpread1" Then
            If Spread_Print(Me.FpSpread1, "Triage Code", 0) = False Then
                MsgBox("Fail to Print")
            End If
        ElseIf save_excel = "FpSpread4" Then
            If Spread_Print(Me.FpSpread4, "Triage Details", 0) = False Then
                MsgBox("Fail to Print")
            End If
        Else
            MessageBox.Show("Select Spread for Printing!!")
        End If
    End Sub

    Private Sub ButtonItem6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonItem6.Click
        If save_excel = "" Then
            MessageBox.Show("Select Spread for Saving!!")
            Exit Sub
        End If

        If save_excel = "FpSpread1" Then
            File_Save(SaveFileDialog1, FpSpread1)
        ElseIf save_excel = "FpSpread4" Then
            File_Save(SaveFileDialog1, FpSpread4)
        Else
            MessageBox.Show("Select Spread for Saving!!")
        End If
    End Sub

    Private Sub FpSpread1_CellClick(ByVal sender As System.Object, ByVal e As FarPoint.Win.Spread.CellClickEventArgs) Handles FpSpread1.CellClick
        save_excel = "FpSpread1"
    End Sub


    Private Sub ButtonItem8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonItem8.Click
        FpSpread4.ActiveSheet.RowCount = 0
        FpSpread2.ActiveSheet.RowCount = 0
    End Sub

    Private Sub FrmTriage_Shown(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Shown
        TextBoxX2.Focus()
        TextBoxX2.SelectAll()
    End Sub

    Private Sub TextBoxX2_click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBoxX2.Click
        TextBoxX2.Text = ""
        TextBoxX2.SelectAll()
    End Sub

    Private Sub ButtonItem10_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonItem10.Click
        With FpSpread2.ActiveSheet
            Print_NewBarcode(.GetValue(.ActiveRowIndex, 3), .GetValue(.ActiveRowIndex, 1), Query_RS("select ISNULL(TRG01,'') + ':' + ISNULL(TRG02,'') FROM VIEW_LGTRIAGE WHERE OBID = (select OBID from tbl_esnmaster where site_id = 'S1000' and esn = '" & .GetValue(.ActiveRowIndex, 0) & "')"), Query_RS("select convert(varchar(10), c_date, 101) from tbl_fesnmaster where esn = '" & .GetValue(.ActiveRowIndex, 3) & "'"))
            '            Print_NewBarcode(.GetValue(.ActiveRowIndex, 3), .GetValue(.ActiveRowIndex, 1), .GetValue(.ActiveRowIndex, 2), Query_RS("select convert(varchar(10), c_date, 101) from tbl_fesnmaster where esn = '" & .GetValue(.ActiveRowIndex, 3) & "'"))
            '            Print_BBarcode(.GetValue(.ActiveRowIndex, 3), .GetValue(.ActiveRowIndex, 1), .GetValue(.ActiveRowIndex, 1))
        End With
    End Sub


    Function SCAN_FIRST() As Boolean

        TextBoxX2.Text = ""
        TextBoxX3.Text = ""
        TextBoxX4.Text = ""
        TextBoxX5.Text = ""
        '      FpSpread2.ActiveSheet.RowCount = 0
        FpSpread4.ActiveSheet.RowCount = 0

        '        TextBoxX2.Focus()
        ComboBoxEx6.Focus()

    End Function


    Private Sub IntegerInput1_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles IntegerInput1.KeyDown
        If e.KeyCode <> Keys.Enter Then
            Exit Sub
        End If

        If IntegerInput1.Text <= 0 Then
            Modal_Error("수량을 입력하십시오!")
            IntegerInput1.Focus()
            Exit Sub
        Else
            TextBoxX3.Focus()
            TextBoxX3.SelectAll()
        End If

    End Sub
End Class