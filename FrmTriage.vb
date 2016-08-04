Public Class FrmTriage

    Private Sub FrmTriage_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.GotFocus
        Me.TextBoxX2.Focus()
        Me.TextBoxX2.SelectAll()

        SCAN_FIRST()
    End Sub

    Private Sub FrmTriage_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown

        If e.KeyCode = Keys.Enter Then
            SCAN_FIRST()
        End If

    End Sub

    Private Sub FrmTriage_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim fp_rs As New ADODB.Recordset
        Dim i, j As Integer

        DockContainerItem1.Text = "실행 메뉴"
        DockContainerItem2.Text = "입력 조건"
        DockContainerItem3.Text = "불량 코드"
        DockContainerItem4.Text = "스캔 리스트"

        ComboBoxEx5.Visible = True
        ComboBoxEx5.Enabled = True

        DateTimeInput1.Value = Now

        'MODEL 
        If Query_Combo(Me.ComboBoxEx5, "SELECT model_no  FROM tbl_modelmaster WHERE site_id = '" & Site_id & "' and active = 'Y' ORDER BY model_no") = True Then
        End If

        fp_rs = Query_RS_ALL("select distinct left(code_id,1) from tbl_defect  WHERE ACTIVE = 'Y'  order by left(code_id,1)")


        With FpSpread1.ActiveSheet
            .ColumnCount = fp_rs.RecordCount
            .RowCount = 0

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
            .Rows(0, .RowCount - 1).Height = 45
            '.Cells(0, 0, .RowCount - 1, .ColumnCount - 1).Locked = True
            .Protect = True
            Spread_AutoCol(FpSpread1)
        End With


        If Spread_Setting(FpSpread4, "FrmTriage") = True Then
            With FpSpread4.ActiveSheet
                .RowCount = 0
            End With
            Spread_AutoCol(FpSpread4)
        End If

        If Spread_Setting(FpSpread2, "FrmTriage") = True Then
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

        TextBoxX2.Enabled = True
        'TextBoxX1.Focus()
        'TextBoxX1.SelectAll()
        TextBoxX2.Focus()
        TextBoxX2.SelectAll()



    End Sub

    Private Sub FpSpread1_ButtonClicked(ByVal sender As Object, ByVal e As FarPoint.Win.Spread.EditorNotifyEventArgs) Handles FpSpread1.ButtonClicked

        Dim i, j As Integer
        Dim b As New FarPoint.Win.Spread.CellType.ButtonCellType()
        Dim fp_rs As New ADODB.Recordset

        With FpSpread1.ActiveSheet
            '.ColumnCount = fp_rs.RecordCount
            .RowCount = 0

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


        With FpSpread4.ActiveSheet
            If SPREAD_DUP_CHECK(FpSpread4, b.Text, 0) = True Then
                '            FpSpread4.ActiveSheet.RowCount = 0
                .RowCount = .RowCount + 1
                .SetValue(.RowCount - 1, 0, FpSpread1.ActiveSheet.GetValue(e.Row, e.Column))
                .SetValue(.RowCount - 1, 1, Query_RS("select code_name from tbl_defect where code_id = '" & .GetValue(.RowCount - 1, 0) & "'"))
                .SetValue(.RowCount - 1, 2, IntegerInput1.Text)
                .Rows(.RowCount - 1).ForeColor = Color.OrangeRed
                Spread_AutoCol(FpSpread4)

            End If

        End With

        FpSpread1.ActiveSheet.SetActiveCell(e.Row, e.Column)

    End Sub

    Private Sub FpSpread4_CellClick(ByVal sender As Object, ByVal e As FarPoint.Win.Spread.CellClickEventArgs) Handles FpSpread4.CellClick
        save_excel = "FpSpread4"
        ButtonItem4.Enabled = True
        'ButtonItem4.Visible = True
    End Sub

    Private Sub FpSpread4_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles FpSpread4.Leave
        ButtonItem4.Enabled = False
    End Sub



    Private Sub TextBoxX2_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TextBoxX2.KeyDown

        If e.KeyCode = Keys.Enter Then
            'defect code : 2 DIGIT

            If Len(Me.TextBoxX2.Text) = 4 Or Len(Me.TextBoxX2.Text) = 3 Then
                If SPREAD_DUP_CHECK(FpSpread4, TextBoxX2.Text, 0) = True Then

                    If TextBoxX2.Text = "0YY" Then
                        If FpSpread4.ActiveSheet.RowCount = 0 Then
                            Modal_Error("불량코드를 선택하십시오!")
                            TextBoxX2.Focus()
                            TextBoxX2.SelectAll()
                        End If

                        TextBoxX2.Text = ""
                        ComboBoxEx5.Focus()
                        'TextBoxX4.Focus()
                        'TextBoxX4.SelectAll()
                        Exit Sub
                    ElseIf TextBoxX2.Text = "0ZZ" Then
                        SCAN_FIRST()
                        Exit Sub
                    End If

                    With FpSpread4.ActiveSheet

                        'Defect Code 검사
                        If Query_RS("select count(CODE_ID) from tbl_defECT  WHERE CODE_ID = '" & TextBoxX2.Text & "'") = 0 Then
                            Modal_Error("입력하신 불량코드를 확인하세요.")
                            Exit Sub
                        End If



                        Dim DEF As String = Query_RS("select code_name from tbl_defect where code_id = '" & TextBoxX2.Text & "'")

                        If DEF = "" Then
                            Modal_Error("입력하신 불량코드를 확인하세요.!")
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

        Dim QRY As String = ""

        With FpSpread4.ActiveSheet
            If .RowCount = 0 Then
                Modal_Error("불량 내역을 입력하세요!")
                TextBoxX1.Focus()
                TextBoxX1.SelectAll()
                Exit Sub

            End If

            If IntegerInput1.Text = 0 Or IntegerInput1.Text = "" Then
                Modal_Error("불량 수량을 입력하세요!")
                TextBoxX1.Focus()
                TextBoxX1.SelectAll()
            End If

            For I As Integer = 0 To .RowCount - 1
                If .Rows(I).ForeColor = Color.OrangeRed Then

                    QRY = "INSERT INTO TBL_FAILMASTER VALUES ('" & Site_id & "','" & DateTimeInput1.Text & "','" & .GetValue(I, 0) & "','" & .GetValue(I, 1) & "','수입검사'," & vbNewLine
                    QRY = QRY & .GetValue(I, 2) & ", '" & Emp_No & "', GETDATE(), '" & Emp_No & "', GETDATE(), NULL,NULL,NULL,NULL,NULL)" & vbNewLine

                    If Insert_Data(QRY) = True Then
                        .Rows(I).ForeColor = Color.Black

                        FpSpread2.ActiveSheet.AddRows(0, 1)
                        FpSpread2.ActiveSheet.SetValue(0, 0, ComboBoxEx5.Text)
                        FpSpread2.ActiveSheet.SetValue(0, 1, .GetValue(I, 0))
                        FpSpread2.ActiveSheet.SetValue(0, 2, .GetValue(I, 2))
                        Spread_AutoCol(FpSpread2)
                    End If
                End If
            Next

        End With

        FpSpread4.ActiveSheet.RowCount = 0

        MessageBox.Show("저장이 완료되었습니다.")
        'SCAN_FIRST()
        TextBoxX2.Focus()

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

    Private Sub TextBoxX1_click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        TextBoxX1.Text = ""
        TextBoxX1.SelectAll()
    End Sub

    Private Sub TextBoxX2_click(ByVal sender As System.Object, ByVal e As System.EventArgs)
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

    Private Sub ComboBoxEx5_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBoxEx5.SelectedIndexChanged

        FpSpread2.ActiveSheet.RowCount = 0
        FpSpread4.ActiveSheet.RowCount = 0

        TextBoxX1.Text = ""
        TextBoxX1.Text = Query_RS("SELECT MODEL_NAME FROM TBL_MODELMASTER WHERE SITE_ID = '" & Site_id & "' AND MODEL_NO = '" & ComboBoxEx5.Text & "'")

    End Sub

    Function SCAN_FIRST() As Boolean


        TextBoxX2.Text = ""

        FpSpread2.ActiveSheet.RowCount = 0
        FpSpread4.ActiveSheet.RowCount = 0

        TextBoxX2.Focus()

    End Function


End Class