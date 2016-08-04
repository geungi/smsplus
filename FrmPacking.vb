Imports System.Xml
Imports System.Xml.XPath
Imports System.Xml.Xsl

Public Class FrmPacking

    Private Sub FrmPacking_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        DockContainerItem1.Text = "실행 메뉴"
        DockContainerItem2.Text = "입력 조건"
        DockContainerItem3.Text = "현재 포장 현황"
        DockContainerItem4.Text = "금일 출하된 포장 현황"

        If Spread_Setting(FpSpread1, "FrmPacking") = True Then
            With FpSpread1.ActiveSheet
                .RowCount = 0

                .ColumnCount = .ColumnCount + 1
                .ColumnHeader.Columns(.ColumnCount - 1).Label = "거래처"
            End With



            Spread_AutoCol(FpSpread1)
        End If

        If Spread_Setting(FpSpread2, "FrmPacking") = True Then
            With FpSpread2.ActiveSheet
                .RowCount = 0

                .ColumnCount = .ColumnCount + 1
                .ColumnHeader.Columns(.ColumnCount - 1).Label = "거래처"
            End With


            Spread_AutoCol(FpSpread2)
        End If

        If Spread_Setting(FpSpread3, "FrmPacking") = True Then
            FpSpread3.ActiveSheet.RowCount = 0
            Spread_AutoCol(FpSpread3)
        End If

        If Query_Combo(Me.ComboBoxEx1, "SELECT CODE_ID FROM tbl_CODEmaster WHERE site_id = '" & Site_id & "' and CLASS_ID = 'R0007' AND active = 'Y' ORDER BY DIS_ORDER") = True Then
        End If
        Me.ComboBoxEx1.Text = "GOOD"

        If Query_Combo(Me.ComboBoxEx4, "SELECT CODE_ID FROM tbl_CODEmaster WHERE site_id = '" & Site_id & "' and CLASS_ID = 'R0007' AND active = 'Y' ORDER BY DIS_ORDER") = True Then
        End If
        Me.ComboBoxEx4.Text = "GOOD"


        If Query_Combo(Me.ComboBoxEx2, "SELECT distinct inbox_NO FROM tbl_Fesnmaster_K WHERE site_id = '" & Site_id & "' and inbox_NO is not null ORDER BY inbox_NO") = True Then
        End If

        If Query_Combo(Me.ComboBoxEx3, "SELECT model_no FROM tbl_modelmaster WHERE site_id = '" & Site_id & "' and active = 'Y' ORDER BY model_no") = True Then
        End If
        Me.ComboBoxEx3.Items.Add("ALL")
        Me.ComboBoxEx3.Text = "ALL"

        If Query_Combo(Me.ComboBoxEx5, "SELECT AC_NM FROM tbl_ACINFO WHERE site_id = '" & Site_id & "' ORDER BY AC_NM") = True Then
        End If
        Me.ComboBoxEx5.Items.Add("ALL")
        Me.ComboBoxEx5.Text = "ALL"




        Formbim_Authority(Me.NewBtn, Me.Name, "NEW")
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

        Refresh_Current(Me.ComboBoxEx3.Text)

    End Sub

    Private Sub TextBoxX1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBoxX1.Click
        Me.TextBoxX1.Text = ""
    End Sub

    Private Sub TextBoxX1_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TextBoxX1.KeyDown
        Dim esn_rs As New ADODB.Recordset
        Dim QTY As Integer

        LabelX4.Text = ""

        If ComboBoxEx2.Text = "" Then
            System.Console.Beep(3000, 400)
            System.Console.Beep(3000, 400)
            Modal_Error("SELECT BOXNO!!")
            esn_rs = Nothing

            Me.TextBoxX1.Text = ""
            Exit Sub
        End If

        If ComboBoxEx3.Text = "" Then
            System.Console.Beep(3000, 400)
            System.Console.Beep(3000, 400)
            Modal_Error("SELECT MODEL!!")
            esn_rs = Nothing

            Me.TextBoxX1.Text = ""
            Exit Sub
        End If

        If ComboBoxEx5.Text = "" Or ComboBoxEx5.Text = "ALL" Then
            System.Console.Beep(3000, 400)
            System.Console.Beep(3000, 400)
            Modal_Error("거래처를 선택하세요!!")
            esn_rs = Nothing

            Me.TextBoxX1.Text = ""
            Exit Sub
        End If


        If e.KeyCode <> Keys.Enter Then
            Exit Sub
        End If


        If FESN_VERIFY(Me.TextBoxX1, e) = True Then

            'Dim O_ESN As String = Query_RS("SELECT MODEL_NO FROM TBL_MODELMASTER WHERE RESERV2 = '" & TextBoxX1.Text & "'")

            Dim O_ESN As String = Query_RS("SELECT MODEL_NO FROM TBL_MODELMASTER WHERE RESERV1 = '" & TextBoxX1.Text & "'")
            If Check_Valid_FEsn(TextBoxX1.Text, Me.Name) = False Then
                'Me.TextBoxX1.Text = ""
                Me.TextBoxX1.Focus()
                Me.TextBoxX1.SelectAll()
                Exit Sub
            End If

            Dim LOT_QTY As Integer = Query_RS("SELECT ISNULL((SELECT ISNULL(ACT_QTY,0) FROM TBL_LOTMASTER WHERE LOT_NO = '" & ComboBoxEx2.Text & "'),0)")
            Dim LOT_TOT_QTY As Integer = CInt(Query_RS("SELECT ISNULL((SELECT isnull(SUM(act_QTY),0) FROM TBL_LOTMASTER WHERE MODEL = '" & O_ESN & "' and return_dv = '" & ComboBoxEx1.Text & "' AND LOT_NO <> '" & ComboBoxEx2.Text & "'),0)"))

            If LOT_QTY > 0 Then
                'If Query_RS("SELECT MODEL FROM TBL_LOTMASTER WHERE LOT_NO = '" & ComboBoxEx2.Text & "'") <> O_ESN Then
                '    System.Console.Beep(3000, 400)
                '    System.Console.Beep(3000, 400)
                '    Modal_Error(TextBoxX1.Text & vbNewLine & "선택한 모델과 동일하지 않은 모델입니다!!")

                '    'Me.TextBoxX1.Text = ""
                '    Me.TextBoxX1.Focus()
                '    Me.TextBoxX1.SelectAll()
                '    Exit Sub

                'End If

            End If

            QTY = Query_RS("SELECT sum(QTY) FROM TBL_PARTINV WHERE WH_CD in ('W1000', 'C2000') AND PART_NO = '" & O_ESN & "'")

            If QTY <= 0 Then
                System.Console.Beep(3000, 400)
                System.Console.Beep(3000, 400)
                Modal_Error(TextBoxX1.Text & vbNewLine & "재고가 존재하지 않는 모델입니다!!")
                
                'Me.TextBoxX1.Text = ""
                Me.TextBoxX1.Focus()
                Me.TextBoxX1.SelectAll()
                Exit Sub
            Else
                If QTY < (LOT_TOT_QTY + LOT_QTY + FpSpread3.ActiveSheet.RowCount) Then
                    System.Console.Beep(3000, 400)
                    System.Console.Beep(3000, 400)
                    Modal_Error(TextBoxX1.Text & vbNewLine & "자재창고의 재고를 초과하였습니다!!")
                    esn_rs = Nothing

                    'Me.TextBoxX1.Text = ""
                    Me.TextBoxX1.Focus()
                    Me.TextBoxX1.SelectAll()
                    Exit Sub

                Else

                    Dim CNT As Integer
                    CNT = (CInt(FpSpread1.ActiveSheet.GetValue(SPREAD_DUP_ROW(FpSpread1, ComboBoxEx2.Text, 3), 4)) + FpSpread3.ActiveSheet.RowCount)

                    If ComboBoxEx1.Text = "GOOD" Then
                        If CNT < CInt(Query_RS("SELECT CODE_NAME FROM TBL_CODEMASTER WHERE SITE_ID = '" & Site_id & "' AND CLASS_ID = 'R0005' AND CODE_ID  = '" & Microsoft.VisualBasic.Left(ComboBoxEx2.Text, 1) & "'")) Then
                        Else
                            System.Console.Beep(3000, 400)
                            System.Console.Beep(3000, 400)
                            Modal_Error("박스 적재 수량 초과!!")
                            esn_rs = Nothing

                            'Me.TextBoxX1.Text = ""
                            Me.TextBoxX1.Focus()
                            Me.TextBoxX1.SelectAll()
                            Exit Sub
                        End If


                    End If

                    FpSpread3.ActiveSheet.AddRows(0, 1)
                    FpSpread3.ActiveSheet.Rows(0).ForeColor = Color.OrangeRed
                    FpSpread3.ActiveSheet.SetValue(0, 0, O_ESN) 'OBID
                    FpSpread3.ActiveSheet.SetValue(0, 1, TextBoxX1.Text) 'ESN
                    FpSpread3.ActiveSheet.SetValue(0, 2, O_ESN) 'MODEL
                    FpSpread3.ActiveSheet.SetValue(0, 3, ComboBoxEx2.Text) 'BOX
                    If FpSpread3.ActiveSheet.RowCount = 1 Then
                        FpSpread3.ActiveSheet.SetValue(0, 4, CInt(FpSpread1.ActiveSheet.GetValue(SPREAD_DUP_ROW(FpSpread1, ComboBoxEx2.Text, 3), 4)) + 1) 'POSITION
                        'FpSpread1.ActiveSheet.SetValue(SPREAD_DUP_ROW(FpSpread1, ComboBoxEx2.Text, 3), 0, esn_rs(2).Value)
                        '                      FpSpread1.ActiveSheet.SetValue(SPREAD_DUP_ROW(FpSpread1, ComboBoxEx2.Text, 3), 4, esn_rs(2).Value)
                    Else
                        FpSpread3.ActiveSheet.SetValue(0, 4, CInt(FpSpread3.ActiveSheet.GetValue(1, 4)) + 1) 'POSITION
                    End If
                    FpSpread3.ActiveSheet.SetValue(0, 5, "")
                    FpSpread3.ActiveSheet.SetValue(0, 6, "")
                    FpSpread3.ActiveSheet.SetValue(0, 7, "")

                    FpSpread3.ActiveSheet.SetActiveCell(FpSpread3.ActiveSheet.RowCount - 1, 0)
                    Me.FpSpread3.ShowActiveCell(FarPoint.Win.Spread.VerticalPosition.Bottom, FarPoint.Win.Spread.HorizontalPosition.Left)
                    '                    Print_Barcode_Coffin(TextBoxX1.Text, esn_rs(3).Value, esn_rs(19).Value)

                    Spread_AutoCol(FpSpread3)
                    Spread_AutoCol(FpSpread1)

                    Me.TextBoxX1.Focus()
                    Me.TextBoxX1.SelectAll()

                End If

            End If
        End If
        esn_rs = Nothing

    End Sub


    Private Sub ButtonItem3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonItem3.Click, SaveBtn1.Click
        Dim O_CNT As Integer = 0
        Dim N_CNT As Integer = 0

        Dim RS As New ADODB.Recordset
        'If Form_Authority(Me.Name, ButtonItem3) = False Then
        '    MessageBox.Show("No Authority")
        'End If


        If ComboBoxEx2.Text = "" Then
            System.Console.Beep(3000, 400)
            System.Console.Beep(3000, 400)
            Modal_Error("SELECT BOXNO!!")


            Me.TextBoxX1.Text = ""
            Exit Sub
        End If

        If ComboBoxEx3.Text = "" Then
            System.Console.Beep(3000, 400)
            System.Console.Beep(3000, 400)
            Modal_Error("SELECT MODEL!!")


            Me.TextBoxX1.Text = ""
            Exit Sub
        End If


        With FpSpread3.ActiveSheet
            If .RowCount = 0 Then
                Modal_Error("저장할 내용이 존재하지 않습니다.")
                Exit Sub
            End If


            RS = Query_RS_ALL("SELECT ISNULL(ACT_QTY,0), ISNULL(MODEL,'') FROM TBL_LOTMASTER WHERE LOT_NO = '" & ComboBoxEx2.Text & "'")

            If RS Is Nothing Then
                O_CNT = 0
            Else
                O_CNT = RS(0).Value
            End If

            N_CNT = .RowCount

            If N_CNT > 0 Then
                For I As Integer = 0 To .RowCount - 1
                    'Dim O_ESN As String = Query_RS("SELECT MODEL_NO FROM TBL_MODELMASTER WHERE RESERV1 = '" & TextBoxX1.Text & "'")

                    If Insert_Data("exec SP_FRMPACKING_LOTSAVE '" & Site_id & "','" & ComboBoxEx2.Text & "','" & .GetValue(I, 2) & "','1','포장 " & ComboBoxEx2.Text & "'  ,'" & Emp_No & "','W1000','C2000','" & OP_No & "','" & ComboBoxEx1.Text & "','" & ComboBoxEx5.Text & "'") = True Then
                    End If
                Next
            Else
                Modal_Error("이미 저장된 내용입니다.")
                Exit Sub
            End If

            RS = Nothing
        End With

        FpSpread3.ActiveSheet.RowCount = 0

        ComboBoxEx2_SelectedIndexChanged(sender, e)
        ComboBoxEx2.Text = ""

        MessageBox.Show("저장되었습니다")

        Refresh_Current(ComboBoxEx3.Text)

    End Sub

    Private Sub NewBtn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NewBtn.Click, NewBtn1.Click
        'If Form_Authority(Me.Name, NewBtn) = False Then
        '    MessageBox.Show("No Authority")
        'End If

        ComboBoxEx2.Text = ""
        If ComboBoxEx1.Text = "GOOD" Then
            Dim prefix As String = ""
            If ComboBoxEx5.Text = "윌리스" Then
                prefix = "W"
            ElseIf ComboBoxEx5.Text = "SCOTTII MARKETING" Then
                prefix = "M"
            ElseIf ComboBoxEx5.Text = "SCOTTII SAMPLE" Then
                prefix = "S"
            ElseIf ComboBoxEx5.Text = "SCOTTII USA" Then
                prefix = "U"
            Else
                prefix = "K"
            End If

            ComboBoxEx2.Text = Query_RS("SELECT '" & prefix & "'+ CONVERT(VARCHAR(8),GETDATE(),112) +'-'+ISNULL(RIGHT('00'+CONVERT(VARCHAR(3),CONVERT(INT,RIGHT(MAX(right(LOT_NO,3)),3))+1),3),'001') FROM VIEW_LOTMASTER WHERE LOT_NO LIKE '" & prefix & "'+ CONVERT(VARCHAR(8),GETDATE(),112) + '%'") 'LIKE 'K'+ CONVERT(VARCHAR(8),GETDATE(),112)+'%'")
            '            ComboBoxEx2.Text = Query_RS("SELECT 'K'+ CONVERT(VARCHAR(8),GETDATE(),112) +'-'+ISNULL(RIGHT('00'+CONVERT(VARCHAR(4),CONVERT(INT,RIGHT(MAX(INBOX_NO),4))+1),4),'001') FROM TBL_FESNMASTER_K WHERE INBOX_NO LIKE 'K'+ CONVERT(VARCHAR(8),GETDATE(),112)+'%'")
        Else
            ComboBoxEx2.Text = Query_RS("SELECT 'B'+ CONVERT(VARCHAR(8),GETDATE(),112) +'-'+ISNULL(RIGHT('00'+CONVERT(VARCHAR(3),CONVERT(INT,RIGHT(MAX(LOT_NO),3))+1),3),'001') FROM VIEW_LOTMASTER WHERE LOT_NO LIKE 'B'+ CONVERT(VARCHAR(8),GETDATE(),112) + '%'")
        End If
        ComboBoxEx2_Leave(sender, e)

        Spread_AutoCol(FpSpread1)

    End Sub

    Private Sub ComboBoxEx2_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles ComboBoxEx2.KeyDown

        If e.KeyData = Keys.Enter Then
            TextBoxX1.Focus()
            Try
                If CInt(Mid(ComboBoxEx2.Text, 2, 1)) >= 0 And CInt(Mid(ComboBoxEx2.Text, 2, 1)) <= 9 Then
                End If

                If Query_RS("select count(inboxlotid) from tbl_shipdtl where substring(ship_no,3,8)  = convert(varchar(8), getdate(), 112) and inboxlotid = '" & ComboBoxEx2.Text & "'") > 0 Then
                    ComboBoxEx2.Text = ""
                    MessageBox.Show("Already Used BOXNO at today!!")
                    Exit Sub
                End If

            Catch ex As Exception
                ComboBoxEx2.Text = ""
                MessageBox.Show("Not Valid BOXno")
                Exit Sub
            End Try

        End If

    End Sub

    Private Sub ComboBoxEx2_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBoxEx2.Leave

        ComboBoxEx2.Text = Microsoft.VisualBasic.UCase(ComboBoxEx2.Text)

        Try
            If Len(ComboBoxEx2.Text) > 0 Then
                If CInt(Mid(ComboBoxEx2.Text, 2, 1)) >= 0 And CInt(Mid(ComboBoxEx2.Text, 2, 1)) <= 9 Then
                End If

            End If

            'If Query_RS("select count(inboxlotid) from tbl_shipdtl where substring(ship_no,3,8)  = convert(varchar(8), getdate(), 112) and inboxlotid = '" & ComboBoxEx2.Text & "'") > 0 Then
            '    ComboBoxEx2.Text = ""
            '    MessageBox.Show("Already Used BOXNO at today!!")
            '    Exit Sub
            'End If


        Catch ex As Exception
            ComboBoxEx2.Text = ""
            ComboBoxEx2.Focus()
            '            MessageBox.Show("Not Valid BOXno")
            Exit Sub
        End Try

        If SPREAD_DUP_CHECK(FpSpread1, ComboBoxEx2.Text, 3) = True Then

            'Refresh_Current(Me.ComboBoxEx3.Text)

            If Len(ComboBoxEx2.Text) > 4 Then
                If Query_RS("SELECT COUNT(INBOX_NO) FROM TBL_FESNMASTER_K WHERE INBOX_NO = '" & ComboBoxEx2.Text & "'") > 0 Then
                Else
                    If SPREAD_DUP_CHECK(FpSpread1, ComboBoxEx2.Text, 2) = True Then
                        FpSpread1.ActiveSheet.AddRows(0, 1)
                        If ComboBoxEx3.Text <> "ALL" Then
                            FpSpread1.ActiveSheet.SetValue(0, 0, ComboBoxEx3.Text)
                        End If
                        FpSpread1.ActiveSheet.SetValue(0, 1, ComboBoxEx1.Text)
                        FpSpread1.ActiveSheet.SetValue(0, 3, ComboBoxEx2.Text)
                        FpSpread1.ActiveSheet.SetValue(0, 4, 0)
                    End If
                End If
            End If

        Else
            Exit Sub
        End If




    End Sub


    Private Sub ComboBoxEx2_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBoxEx2.SelectedIndexChanged
        ComboBoxEx2.Text = Microsoft.VisualBasic.UCase(ComboBoxEx2.Text)

        'Refresh_Current(Me.ComboBoxEx3.Text)
    End Sub

    Private Sub ComboBoxEx2_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBoxEx2.TextChanged
        If Len(ComboBoxEx2.Text) > 4 Then
            If SPREAD_DUP_CHECK(FpSpread1, ComboBoxEx2.Text, 2) = False Then
            End If

            FpSpread3.ActiveSheet.RowCount = 0

            If Query_RS("select count(inbox_NO) from tbl_FESNMASTER_K where substring(Cship_no,3,8)  = convert(varchar(8), getdate(), 112) and inbox_NO = '" & ComboBoxEx2.Text & "'") > 0 Then
                ComboBoxEx2.Text = ""
                MessageBox.Show("Already Used BOXNO at today!!")
                Exit Sub
            End If
            'Refresh_Current(Me.ComboBoxEx3.Text)
        End If
    End Sub

    Function Refresh_Current(ByVal model As String) As Boolean

        FpSpread3.ActiveSheet.RowCount = 0
        FpSpread1.AllowUserFormulas = True

        ComboBoxEx2.Text = ""

        If Query_Spread(FpSpread1, "exec SP_FRMPACKING_GETCURBOX_K '" & Site_id & "', '" & model & "','" & ComboBoxEx4.Text & "','" & ComboBoxEx5.Text & "'", 1) = True Then
            FpSpread1.ActiveSheet.RowCount = FpSpread1.ActiveSheet.RowCount + 1
            FpSpread1.ActiveSheet.Rows(FpSpread1.ActiveSheet.RowCount - 1).BackColor = Color.Yellow
            'FpSpread1.ActiveSheet.Cells(FpSpread1.ActiveSheet.RowCount - 1, 0).ColumnSpan = 3
            FpSpread1.ActiveSheet.Cells(FpSpread1.ActiveSheet.RowCount - 1, 0).Text = "TOTAL"

            FpSpread1.ActiveSheet.Cells(FpSpread1.ActiveSheet.RowCount - 1, 4).Formula = "SUM(E1:E" & FpSpread1.ActiveSheet.RowCount - 1 & ")"
            Spread_AutoCol(FpSpread1)
            Spread_AutoCol(FpSpread1)

        End If

        If Query_Spread(FpSpread2, "exec SP_FRMPACKING_GETDAILYTOTAL '" & Site_id & "', '" & model & "'", 1) = True Then

            FpSpread2.ActiveSheet.RowCount = FpSpread2.ActiveSheet.RowCount + 1
            FpSpread2.ActiveSheet.Rows(FpSpread2.ActiveSheet.RowCount - 1).BackColor = Color.Yellow
            'FpSpread1.ActiveSheet.Cells(FpSpread1.ActiveSheet.RowCount - 1, 0).ColumnSpan = 3
            FpSpread2.ActiveSheet.Cells(FpSpread2.ActiveSheet.RowCount - 1, 0).Text = "TOTAL"

            FpSpread2.ActiveSheet.Cells(FpSpread2.ActiveSheet.RowCount - 1, 3).Formula = "SUM(D1:D" & FpSpread2.ActiveSheet.RowCount - 1 & ")"
            Spread_AutoCol(FpSpread2)
        End If

    End Function

    Private Sub ComboBoxEx3_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBoxEx3.SelectedIndexChanged, ButtonItem11.Click, ButtonItem13.Click, ButtonX1.Click
        Refresh_Current(Me.ComboBoxEx3.Text)
    End Sub

    Private Sub ComboBoxEx1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBoxEx1.SelectedIndexChanged
        FpSpread3.ActiveSheet.RowCount = 0
        ComboBoxEx2.Text = ""
    End Sub

    Private Sub ButtonItem4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonItem4.Click, DelBtn1.Click

    End Sub

    Private Sub ButtonItem5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonItem5.Click, PrtBtn1.Click
        If save_excel = "" Then
            MessageBox.Show("Select Spread for Printing!!")
            Exit Sub
        End If

        If save_excel = "FpSpread1" Then
            If Spread_Print(Me.FpSpread1, DockContainerItem3.Text, 0) = False Then
                MsgBox("Fail to Print")
            End If
        ElseIf save_excel = "FpSpread2" Then
            If Spread_Print(Me.FpSpread2, DockContainerItem4.Text, 0) = False Then
                MsgBox("Fail to Print")
            End If
        ElseIf save_excel = "FpSpread3" Then
            If Spread_Print(Me.FpSpread3, "Packing Details", 0) = False Then
                MsgBox("Fail to Print")
            End If
            'ElseIf save_excel = "FpSpread4" Then
            '    If Spread_Print(Me.FpSpread4, DockContainerItem5.Text, 0) = False Then
            '        MsgBox("Fail to Print")
            '    End If
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
            'ElseIf save_excel = "FpSpread4" Then
            '    File_Save(SaveFileDialog1, FpSpread4)
        Else
            MessageBox.Show("Select Spread for Saving!!")
        End If
    End Sub

    Private Sub FpSpread1_CellClick(ByVal sender As Object, ByVal e As FarPoint.Win.Spread.CellClickEventArgs) Handles FpSpread1.CellClick
        FpSpread3.ActiveSheet.RowCount = 0
        ComboBoxEx2.Text = FpSpread1.ActiveSheet.GetValue(e.Row, 3)
        save_excel = "FpSpread1"
        'ComboBoxEx3.Text = FpSpread1.ActiveSheet.GetValue(e.Row, 0)
    End Sub

    Private Sub FpSpread1_CellDoubleClick(ByVal sender As Object, ByVal e As FarPoint.Win.Spread.CellClickEventArgs) Handles FpSpread1.CellDoubleClick
        FpSpread3.ActiveSheet.RowCount = 0
        save_excel = "FpSpread1"

        Dim QRY As String

        QRY = "select OBID, IN_esn , MODEL, ISNULL(INBOX_NO,''), IN_POS,  OUT_ESN, PBOX_NO, ISNULL(RETURN_DV,'GOOD') FROM tbl_fesnmaster_k A" & vbNewLine
        QRY = QRY & "WHERE SITE_ID = '" & Site_id & "'" & vbNewLine
        QRY = QRY & "  AND INBOX_NO = '" & FpSpread1.ActiveSheet.GetValue(e.Row, 3) & "'" & vbNewLine
        QRY = QRY & "ORDER BY CONVERT(INT,IN_POS) DESC" & vbNewLine

        If Query_Spread(FpSpread3, QRY, 1) = True Then
            Spread_AutoCol(FpSpread3)
        End If
    End Sub

    Private Sub FpSpread2_CellClick(ByVal sender As System.Object, ByVal e As FarPoint.Win.Spread.CellClickEventArgs) Handles FpSpread2.CellClick
        save_excel = "FpSpread2"
    End Sub
    Private Sub FpSpread3_CellClick(ByVal sender As System.Object, ByVal e As FarPoint.Win.Spread.CellClickEventArgs) Handles FpSpread3.CellClick
        save_excel = "FpSpread3"
    End Sub
    Private Sub FpSpread4_CellClick(ByVal sender As System.Object, ByVal e As FarPoint.Win.Spread.CellClickEventArgs)
        save_excel = "FpSpread4"
    End Sub

    Private Sub ButtonItem7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonItem7.Click
        'With FpSpread1.ActiveSheet
        '    If .RowCount > 0 Then
        '        Dim r As DialogResult = MessageBox.Show("Are you sure to cancel box " & .GetValue(.ActiveRowIndex, 3) & "?", "Cancel Packing BOX", MessageBoxButtons.YesNo)
        '        If r = Windows.Forms.DialogResult.Yes Then
        '            Insert_Data("EXEC SP_FRMPACKING_CANCELPACKING '" & Site_id & "','" & .GetValue(.ActiveRowIndex, 3) & "','" & .GetValue(.ActiveRowIndex, 0) & "','" & Emp_No & "'")
        '            Refresh_Current(ComboBoxEx3.Text)
        '        End If

        '    End If
        'End With
    End Sub

    Private Sub FrmPacking_Shown(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Shown
        Me.ComboBoxEx2.Focus()
    End Sub


    Private Sub ButtonItem9_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonItem9.Click
        'Dim i As Integer

        'With FpSpread3.ActiveSheet

        '    If .RowCount = 0 Then
        '        Exit Sub
        '    End If

        '    Dim r As DialogResult = MessageBox.Show("Are you sure to delete esn " & .GetValue(.ActiveRowIndex, 1) & " in this box?", "Delete ESN", MessageBoxButtons.YesNo)
        '    If r = Windows.Forms.DialogResult.Yes Then
        '        For i = 0 To .ActiveRowIndex
        '            .SetValue(i, 2, .GetValue(i, 2) - 1)
        '        Next

        '        If .Rows(.ActiveRowIndex).ForeColor = Color.OrangeRed Then
        '            .RemoveRows(.ActiveRowIndex, 1)
        '        Else
        '            Insert_Data("Update tbl_esnmaster set pos = pos-1 where site_id = '" & Site_id & "' and inboxid = '" & ComboBoxEx2.Text & "' and pos > " & .GetValue(.ActiveRowIndex, 2))
        '            Insert_Data("Update tbl_esnmaster set inboxid = null, pos = null where site_id = '" & Site_id & "' and esn = '" & .GetValue(.ActiveRowIndex, 1) & "'")
        '            .RemoveRows(.ActiveRowIndex, 1)
        '        End If

        '        If .RowCount > 0 Then
        '            FpSpread1.ActiveSheet.SetValue(SPREAD_DUP_ROW(FpSpread1, ComboBoxEx2.Text, 3), 4, .GetValue(0, 2))
        '        Else
        '            FpSpread1.ActiveSheet.RemoveRows(SPREAD_DUP_ROW(FpSpread1, ComboBoxEx2.Text, 3), 1)
        '        End If
        '        'Insert_Data("EXEC SP_FRMPACKING_CANCELPACKING '" & Site_id & "','" & .GetValue(.ActiveRowIndex, 3) & "','" & .GetValue(.ActiveRowIndex, 0) & "','" & Emp_No & "'")
        '        'Refresh_Current(ComboBoxEx3.Text)
        '    End If
        'End With
    End Sub


    Private Sub ButtonItem10_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonItem10.Click

        'If FpSpread1.ActiveSheet.RowCount = 0 Then
        '    Exit Sub
        'End If

        'With FpSpread5.ActiveSheet
        '    .RowCount = 0
        '    .ColumnCount = 2
        '    .ColumnHeader.Cells(0, 0).Text = "INESN"
        '    .ColumnHeader.Cells(0, 1).Text = "INBOXLOTID"
        '    .Columns(0, 1).CellType = textcell

        '    If Query_Spread(FpSpread5, "select esn, inboxid from tbl_esnmaster where isnull(inboxid, '') <> '' AND RETURN_DV = 'GOOD' order by model, inboxid", 1) = True Then
        '        SaveFileDialog1.FileName = "DM Check"
        '        File_Save(SaveFileDialog1, FpSpread5)

        '    End If

        'End With


    End Sub

    Private Sub ButtonItem14_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonItem14.Click

        'If FpSpread1.ActiveSheet.RowCount = 0 Then
        '    Exit Sub
        'End If

        'With FpSpread5.ActiveSheet
        '    .RowCount = 0
        '    .ColumnCount = 2
        '    .ColumnHeader.Cells(0, 0).Text = "INESN"
        '    .ColumnHeader.Cells(0, 1).Text = "INBOXLOTID"
        '    .Columns(0, 1).CellType = textcell

        '    If Query_Spread(FpSpread5, "select esn,inboxid from tbl_esnmaster where inboxid = '" & FpSpread1.ActiveSheet.GetValue(FpSpread1.ActiveSheet.ActiveRowIndex, 3) & "' AND RETURN_DV = 'GOOD' order by model, inboxid", 1) = True Then
        '        SaveFileDialog1.FileName = "DM Check"
        '        File_Save(SaveFileDialog1, FpSpread5)

        '    End If

        'End With

    End Sub

    Private Sub ButtonItem15_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonItem15.Click

        If MessageBox.Show("박스 라벨을 출력하시겠습니까?", "박스 출력", MessageBoxButtons.YesNo) <> Windows.Forms.DialogResult.Yes Then
            Exit Sub
        End If

        With FpSpread1.ActiveSheet
            If .RowCount = 0 Then
                Exit Sub
            End If

            Dim QRY As String = "SELECT MODEL, (SELECT RESERV2 FROM TBL_MODELMASTER WHERE MODEL_NO = A.MODEL), " & vbNewLine
            QRY = QRY & "   act_QTY" & vbNewLine
            QRY = QRY & "FROM TBL_LOTMASTER A" & vbNewLine
            QRY = QRY & "WHERE LOT_NO = '" & .GetValue(.ActiveRowIndex, 3) & "'" & vbNewLine
            '            QRY = QRY & "GROUP BY MODEL, INBOX_NO" & vbNewLine

            Dim CB_RS As ADODB.Recordset = Query_RS_ALL(QRY)

            If CB_RS Is Nothing Then
                MessageBox.Show("No Data!!")
                Exit Sub
            Else
                Print_BOXBarcode(CB_RS(1).Value, CB_RS(0).Value, CB_RS(2).Value, RichTextBox1)
                '     Print_Barcode_Carton(CB_RS(1).Value, CB_RS(2).Value, CB_RS(3).Value, CB_RS(4).Value, CB_RS(5).Value)
            End If


        End With

    End Sub

    Private Sub ButtonItem16_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonItem16.Click

        If MessageBox.Show("Are you sure to print Carton box barcode?", "Carton box", MessageBoxButtons.YesNo) <> Windows.Forms.DialogResult.Yes Then
            Exit Sub
        End If

        With FpSpread2.ActiveSheet
            If .RowCount = 0 Then
                Exit Sub
            End If

            Dim QRY As String = "SELECT MODEL, INBOXID, " & vbNewLine
            QRY = QRY & "   (SELECT RESERV4 FROM TBL_MODELMASTER WHERE MODEL_NO = A.MODEL), " & vbNewLine
            QRY = QRY & "   (SELECT PRL FROM TBL_MODELMASTER WHERE MODEL_NO = A.MODEL) , (SELECT RESERV5 FROM TBL_MODELMASTER WHERE MODEL_NO = A.MODEL), " & vbNewLine
            QRY = QRY & "   COUNT(MODEL)" & vbNewLine
            QRY = QRY & "FROM TBL_ESNMASTER_B a" & vbNewLine
            QRY = QRY & "WHERE INBOXID = '" & .GetValue(.ActiveRowIndex, 2) & "'" & vbNewLine
            QRY = QRY & "GROUP BY MODEL, INBOXID" & vbNewLine

            Dim CB_RS As ADODB.Recordset = Query_RS_ALL(QRY)

            If CB_RS Is Nothing Then
                MessageBox.Show("No Data!!")
                Exit Sub
            Else
                Print_Barcode_Carton_ship(CB_RS(1).Value, CB_RS(2).Value, CB_RS(3).Value, CB_RS(4).Value, CB_RS(5).Value)
            End If


        End With


    End Sub
End Class