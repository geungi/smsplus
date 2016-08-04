Public Class FrmChangeLot

    Private Sub FrmWCTrf_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.GotFocus
        Me.ComboBoxEx2.Focus()
    End Sub

    Private Sub FrmWCTrf_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        DockContainerItem1.Text = "실행 메뉴"
        DockContainerItem2.Text = "조회 조건"

        If Query_Combo(Me.ComboBoxEx1, "SELECT CODE_NAME FROM tbl_CODEmaster WHERE site_id = '" & Site_id & "' and CLASS_ID = 'R0001' AND active = 'Y' ORDER BY CODE_ID") = True Then
        End If


        ComboBoxEx3.Items.Add("LOT 병합")
        'ComboBoxEx3.Items.Add("LOT 분리")
        ComboBoxEx3.Items.Add("LOT 모델 변경")

        ComboBoxEx4.Visible = False

        If Query_Combo(Me.ComboBoxEx2, "SELECT model_no FROM tbl_modelmaster WHERE site_id = '" & Site_id & "' and active = 'Y' ORDER BY model_no") = True Then
        End If

        Formbim_Authority(Me.ButtonItem2, Me.Name, "NEW")
        Formbim_Authority(Me.ButtonItem3, Me.Name, "SAVE")
        Formbim_Authority(Me.ButtonItem4, Me.Name, "DELETE")
        Formbim_Authority(Me.SaveBtn, Me.Name, "SAVE")
        Formbim_Authority(Me.FindBtn, Me.Name, "FIND")

        ButtonItem2.Enabled = False
        ButtonItem4.Enabled = False

    End Sub

    Private Sub ButtonItem4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonItem4.Click

    End Sub

    Private Sub ButtonItem13_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonItem13.Click
        If save_excel = "" Then
            MessageBox.Show("Select Spread for Printing!!")
            Exit Sub
        End If

        If save_excel = "FpSpread1" Then
            'If Spread_Print(Me.FpSpread1, "NO SHIP(Board) Summary", 0) = False Then
            '    MsgBox("Fail to Print")
            'End If

        Else
            MessageBox.Show("Select Spread for Printing!!")
        End If
    End Sub

    Private Sub ButtonItem14_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonItem14.Click

        If save_excel = "" Then
            MessageBox.Show("Select Spread for Saving!!")
            Exit Sub
        End If

        If save_excel = "FpSpread1" Then
            'File_Save(SaveFileDialog1, FpSpread1)

        Else
            MessageBox.Show("Select Spread for Saving!!")
        End If

    End Sub

    Private Sub ComboBoxEx2_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBoxEx2.SelectedIndexChanged
        ListViewEX1.Items.Clear()
        Me.ComboBoxEx1.Focus()
    End Sub

    Private Sub ComboBoxEx1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBoxEx1.SelectedIndexChanged
        ListViewEX1.Items.Clear()
        Me.ComboBoxEx1.Focus()
    End Sub

    Private Sub FrmWCTrf_Shown(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Shown
        Me.ComboBoxEx2.Focus()
    End Sub

    Private Sub FindBtn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FindBtn.Click
        If ComboBoxEx1.Text = "" Then
            MessageBox.Show("공정을 선택하십시오.")
            Exit Sub
        End If

        If ComboBoxEx2.Text = "" Then
            MessageBox.Show("모델을 선택하십시오.")
            Exit Sub
        End If

        Dim qry As String

        qry = "select LOT_NO, model, " & vbNewLine
        qry = qry & "   (select code_name from TBL_CODEMASTER WHERE SITE_ID = A.SITE_ID AND CLASS_ID = 'R0001' AND CODE_id = a.c_prc), INIT_QTY, ACT_QTY, ISNULL(t_def_cd,''), ISNULL((SELECT CODE_NAME FROM TBL_DEFECT WHERE CODE_ID = A.T_DEF_CD),'')" & vbNewLine
        qry = qry & "FROM TBL_LOTMASTER A" & vbNewLine
        QRY = QRY & "WHERE SITE_ID = '" & Site_id & "'" & vbNewLine

        qry = qry & "  AND MODEL = '" & ComboBoxEx2.Text & "'" & vbNewLine
        qry = qry & "  AND C_PRC = (SELECT CODE_ID FROM TBL_CODEMASTER WHERE SITE_ID = A.SITE_ID AND CLASS_ID = 'R0001' AND CODE_NAME = '" & ComboBoxEx1.Text & "')" & vbNewLine
        qry = qry & "ORDER BY LOT_NO" & vbNewLine

        If Query_Listview(ListViewEX1, qry, True) = True Then

        End If

    End Sub

    Private Sub SaveBtn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SaveBtn.Click

        If ComboBoxEx3.Text = "" Then
            MessageBox.Show("작업내용을 선택하세요.")
            Exit Sub
        End If

        Dim PBOX As String
        Dim QRY As String
        If ComboBoxEx1.Text = "GIB" Then
            PBOX = Query_RS("SELECT 'G'+ CONVERT(VARCHAR(8),GETDATE(),112) +'-'+ISNULL(RIGHT('000'+CONVERT(VARCHAR(4),CONVERT(INT,RIGHT(MAX(LOT_NO),4))+1),4),'0001') FROM TBL_LOTMASTER WHERE LOT_NO LIKE 'G'+ CONVERT(VARCHAR(8),GETDATE(),112)+'%'")
        ElseIf ComboBoxEx1.Text = "BER" Then
            PBOX = Query_RS("SELECT 'B'+ CONVERT(VARCHAR(8),GETDATE(),112) +'-'+ISNULL(RIGHT('000'+CONVERT(VARCHAR(4),CONVERT(INT,RIGHT(MAX(LOT_NO),4))+1),4),'0001') FROM TBL_LOTMASTER WHERE LOT_NO LIKE 'B'+ CONVERT(VARCHAR(8),GETDATE(),112)+'%'")
        Else
            PBOX = Query_RS("SELECT 'P'+ CONVERT(VARCHAR(8),GETDATE(),112) +'-'+ISNULL(RIGHT('000'+CONVERT(VARCHAR(4),CONVERT(INT,RIGHT(MAX(LOT_NO),4))+1),4),'0001') FROM TBL_LOTMASTER WHERE LOT_NO LIKE 'P'+ CONVERT(VARCHAR(8),GETDATE(),112)+'%'")
        End If

        If ComboBoxEx3.Text = "LOT 병합" Then
            If CHECK_SAME_MODEL() = False Then
                Exit Sub
            End If

            '생산LOT 생성 및 TRIAGE 입력
            With ListViewEX1.CheckedItems

                QRY = ""
                QRY = QRY & "EXEC SP_COMMON_LOTCREATE '" & Site_id & "','" & PBOX & "','" & ListViewEX1.CheckedItems.Item(0).SubItems(5).Text & "','" & ComboBoxEx2.Text & "','생산 LOT 생성','" & Emp_No & "','" & ComboBoxEx1.Text & "','" & ComboBoxEx1.Text & "','',''" & vbNewLine

                If Insert_Data(QRY) = True Then
                    Dim QRY1 As String = ""

                    For I As Integer = 0 To .Count - 1
                        If Insert_Data("EXEC SP_COMMON_LOTSAVE_UNION '" & Site_id & "','" & PBOX & "','','','N/A',0,'LOT병합_" & .Item(I).Text & "' ,'" & Emp_No & "','" & ComboBoxEx1.Text & "','" & ComboBoxEx1.Text & "','" & OP_No & "','', -" & .Item(I).SubItems(4).Text) = True Then
                            If Insert_Data("EXEC SP_COMMON_LOTSAVE_UNION '" & Site_id & "','" & .Item(I).Text & "','','','N/A',0,'LOT병합_" & PBOX & "' ,'" & Emp_No & "','" & ComboBoxEx1.Text & "','출하완료','" & OP_No & "','', " & .Item(I).SubItems(4).Text) = True Then
                            End If

                            QRY1 = "UPDATE TBL_FESNMASTER_K" & vbNewLine
                            QRY1 = QRY1 & "SET PBOX_NO = '" & PBOX & "', U_PERSON = '" & Emp_No & "', U_DATE = GETDATE() " & vbNewLine
                            QRY1 = QRY1 & "WHERE OUT_ESN IN (SELECT TOP " & .Item(I).SubItems(4).Text & " OUT_ESN FROM TBL_FESNMASTER_k WHERE PBOX_NO = '" & .Item(I).Text & "' AND INBOX_NO IS NULL AND IN_ESN IS NULL )" & vbNewLine
                            QRY1 = QRY1 & "" & vbNewLine
                        End If

                        If Insert_Data(QRY1) = False Then
                            Modal_Error("병합 오류")
                            Exit Sub
                        End If
                    Next
                End If

            End With
            Print_LOTBarcode(PBOX, ComboBoxEx2.Text, Query_RS("SELECT ACT_QTY FROM TBL_LOTMASTER WHERE LOT_NO = '" & PBOX & "'"), RichTextBox1, "NT01")

        ElseIf ComboBoxEx3.Text = "LOT 모델 변경" Then

            With ListViewEX1.CheckedItems

                QRY = ""

                For I As Integer = 0 To .Count - 1

                    QRY = QRY & "UPDATE TBL_LOTMASTER SET MODEL = '" & ComboBoxEx4.Text & "'" & vbNewLine
                    QRY = QRY & "WHERE LOT_NO = '" & .Item(I).Text & "'" & vbNewLine
                    QRY = QRY & "UPDATE TBL_FESNMASTER_K SET MODEL = '" & ComboBoxEx4.Text & "'" & vbNewLine
                    QRY = QRY & "WHERE PBOX_NO = '" & .Item(I).Text & "'" & vbNewLine
                    QRY = QRY & "" & vbNewLine

                    If Insert_Data(QRY) = True Then
                        Print_LOTBarcode(.Item(I).Text, ComboBoxEx4.Text, Query_RS("SELECT ACT_QTY FROM TBL_LOTMASTER WHERE LOT_NO = '" & .Item(I).Text & "'"), RichTextBox1, "NT01")
                    End If

                Next


            End With


        End If


        MessageBox.Show("작업이 완료되었습니다.")


        FindBtn_Click(sender, e)

        'QRY = "INSERT INTO TBL_LOTMASTER VALUES ('" & Site_id & "','" & PBOX & "',"
        'QRY = QRY & "0,0,'" & FpSpread2.ActiveSheet.GetValue(0, 1) & "','"
        'QRY = QRY & "', "  'T_DEF
        'QRY = QRY & "NULL,NULL,'" & TextBoxX4.Text & "',"
        'QRY = QRY & "'" & Emp_No & "', GETDATE(), '" & Emp_No & "', GETDATE(), "
        'QRY = QRY & "NULL,NULL,NULL,NULL,NULL)"





    End Sub

    Function CHECK_SAME_MODEL() As Boolean

        With ListViewEX1.CheckedItems


            If .Count < 2 Then
                Modal_Error("병합할 LOT를 선택하세요.")
                CHECK_SAME_MODEL = False
                Exit Function
            End If

            For I As Integer = 1 To .Count - 1
                If .Item(I).SubItems(1).Text <> .Item(I - 1).SubItems(1).Text Then
                    Modal_Error("선택한 LOT의 모델이 틀립니다.")
                    CHECK_SAME_MODEL = False
                    Exit Function
                End If

                If Query_RS("SELECT ISNULL(PSHIP_NO,'') FROM TBL_LOTMASTER WHERE LOT_NO = '" & .Item(0).Text & "'") <> "" Then
                    Modal_Error("선택한 LOT는 이미 출하된 LOT 입니다.")
                    Exit Function
                    CHECK_SAME_MODEL = False
                End If

                If Query_RS("SELECT COUNT(IN_ESN) FROM TBL_FESNMASTER_K WHERE PBOX_NO = '" & .Item(0).Text & "' AND IN_ESN IS NOT NULL") > 0 Then
                    Modal_Error("선택한 LOT는 이미 출하검사 완료된 LOT 입니다.")
                    CHECK_SAME_MODEL = False
                    Exit Function
                End If


            Next

        End With

        CHECK_SAME_MODEL = True

    End Function


    Private Sub ComboBoxEx3_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBoxEx3.SelectedIndexChanged

        If ComboBoxEx3.Text = "LOT 모델 변경" Then
            ComboBoxEx4.Visible = True
            If Query_Combo(Me.ComboBoxEx4, "SELECT model_no FROM tbl_modelmaster WHERE site_id = '" & Site_id & "' AND RESERV2 = (SELECT RESERV2 FROM TBL_MODELMASTER WHERE MODEL_NO = '" & ComboBoxEx2.Text & "') and active = 'Y' ORDER BY model_no") = True Then
            End If
        Else
            ComboBoxEx4.Visible = False
        End If

    End Sub
End Class