Public Class FrmWorkFlow

    Private Sub FrmWorkFlow_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        DockContainerItem1.Text = "Form Menu"
        DockContainerItem2.Text = "Retrieve Condition"

        If Query_Combo(ComboBoxEx1, "SELECT model_no FROM tbl_modelmaster WHERE site_id = '" & Site_id & "' ORDER BY model_no") = True Then
            ComboBoxEx1.Items.Add("ALL")
            ComboBoxEx1.Text = "ALL"
        End If

        Query_Combo(Me.ModelCb, "select cus_nm from tbl_customer where site_id = '" & Site_id & "' and cus_type = '1차'  ORDER BY cus_nm")
        Me.ModelCb.Items.Add("ALL")
        Me.ModelCb.Text = "ALL"

        Me.ComboBoxEx2.Text = "ALL"
        Query_Combo(Me.ComboBoxEx2, "select code_name from tbl_codemaster where site_id = '" & Site_id & "' and class_id = 'R0036' order by dis_order")
        Me.ComboBoxEx2.Items.Add("ALL")


        If Spread_Setting(FpSpread1, Me.Name) = True Then
            FpSpread1.ActiveSheet.RowCount = 0
            Spread_AutoCol(FpSpread1)
            FpSpread1.ActiveSheet.FrozenColumnCount = 3
            FpSpread1.ActiveSheet.Protect = False
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
        Formbim_Authority(Me.FindBtn, Me.Name, "FIND")

    End Sub

    Private Sub ButtonItem2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonItem2.Click, NewBtn1.Click


        With FpSpread1.ActiveSheet
            .RowCount = .RowCount + 1
            .Rows(.RowCount - 1).ForeColor = Color.OrangeRed

            .SetValue(.RowCount - 1, 0, Site_id)
            .SetValue(.RowCount - 1, 8, Emp_No)
            .SetValue(.RowCount - 1, 9, Now)
            .SetValue(.RowCount - 1, 10, Emp_No)
            .SetValue(.RowCount - 1, 11, Now)
            .SetActiveCell(.RowCount - 1, 0)
            FpSpread1.ShowActiveCell(FarPoint.Win.Spread.VerticalPosition.Bottom, FarPoint.Win.Spread.HorizontalPosition.Left)

        End With

    End Sub

    Private Sub ButtonItem3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonItem3.Click, SaveBtn1.Click
        Dim I As Integer

        With FpSpread1.ActiveSheet

            Dim QRY As String = ""

            For I = 0 To .RowCount - 1
                If .Rows(I).ForeColor = Color.OrangeRed Then
                    Insert_Data("EXEC SP_FRMWORKFLOW_SETPRICE '" & Site_id & "','" & .GetValue(I, 0) & "','" & .GetValue(I, 1) & "','" & .GetValue(I, 2) & "','" & .GetValue(I, 3) & "'," & .GetValue(I, 5) & "," & .GetValue(I, 6) & ",'" & .GetValue(I, 7) & "','" & Emp_No & "'")
                    .Rows(I).ForeColor = Color.Black
                End If
            Next
        End With

    End Sub

    Private Sub FindBtn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FindBtn.Click

        Dim QRY As String = ""

        QRY = QRY & "SELECT SITE_ID, " & vbNewLine
        QRY = QRY & "       (SELECT CUS_NM FROM TBL_CUSTOMER WHERE CUS_NO = A.CUS_NO1)," & vbNewLine
        QRY = QRY & "       (SELECT CODE_NAME FROM TBL_CODEMASTER WHERE CLASS_ID = 'R0036' AND CODE_ID = A.CUS_NO2)," & vbNewLine
        QRY = QRY & "       MODEL_NO, (SELECT MODEL_NAME FROM TBL_MODELMASTER WHERE MODEL_NO = A.MODEL_NO)," & vbNewLine
        QRY = QRY & "       UPRICE, TAX, RESERV1, " & vbNewLine
        QRY = QRY & "       (SELECT EMP_NM FROM TBL_EMPMASTER WHERE EMP_NO = A.C_PERSON), C_DATE," & vbNewLine
        QRY = QRY & "       (SELECT EMP_NM FROM TBL_EMPMASTER WHERE EMP_NO = A.U_PERSON), U_DATE" & vbNewLine
        QRY = QRY & "FROM TBL_PRICE A" & vbNewLine
        QRY = QRY & "WHERE SITE_ID = '" & Site_id & "'" & vbNewLine

        If ComboBoxEx1.Text <> "ALL" Then
            QRY = QRY & "AND MODEL_NO = '" & ComboBoxEx1.Text & "'" & vbNewLine
        End If

        If ComboBoxEx2.Text <> "ALL" Then
            QRY = QRY & "AND CUS_NO2 = (SELECT CODE_ID FROM TBL_CODEMASTER WHERE CLASS_ID = 'R0036' AND CODE_NAME = '" & ComboBoxEx2.Text & "')" & vbNewLine
        End If

        If ModelCb.Text <> "ALL" Then
            QRY = QRY & "AND CUS_NO1 = (SELECT CUS_NO FROM TBL_CUSTOMER WHERE CUS_NO = A.CUS_NM = '" & ModelCb.Text & "')" & vbNewLine
        End If

        QRY = QRY & "ORDER BY MODEL_NO, CUS_NO1, CUS_NO2" & vbNewLine
        QRY = QRY & "" & vbNewLine
        QRY = QRY & "" & vbNewLine

        If Query_Spread(FpSpread1, QRY, 1) = True Then
            Spread_AutoCol(FpSpread1)
        End If

    End Sub

    Private Sub ComboBoxEx1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBoxEx1.SelectedIndexChanged
        FpSpread1.ActiveSheet.RowCount = 0
    End Sub

    Private Sub ButtonItem4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonItem4.Click, DelBtn1.Click

        Try
            Dim QRY As String = ""


            With FpSpread1.ActiveSheet
                If .RowCount = 0 Then
                    Exit Sub
                End If



                Dim r As DialogResult = MessageBox.Show("Are you sure To delete this row?", "Delete", MessageBoxButtons.YesNo)
                If r = Windows.Forms.DialogResult.Yes Then
                    QRY = QRY & "DELETE FROM TBL_PRICE" & vbNewLine
                    QRY = QRY & "WHERE MODEL_NO = '" & .GetValue(.ActiveRowIndex, 3) & "'" & vbNewLine
                    QRY = QRY & "  AND CUS_NO1 = (SELECT CUS_NO FROM TBL_CUSTOMER WHERE CUS_NO = A.CUS_NM = '" & .GetValue(.ActiveRowIndex, 1) & "')" & vbNewLine
                    QRY = QRY & "  AND CUS_NO2 = (SELECT CODE_ID FROM TBL_CODEMASTER WHERE CLASS_ID = 'R0036' AND CODE_NAME = '" & .GetValue(.ActiveRowIndex, 2) & "')" & vbNewLine
                    QRY = QRY & "  AND UPRICE = '" & .GetValue(.ActiveRowIndex, 5) & "'" & vbNewLine
                    QRY = QRY & "" & vbNewLine

                    If Insert_Data(QRY) = True Then
                        .RemoveRows(.ActiveRowIndex, 1)
                    End If

                    MessageBox.Show("삭제되었습니다", "Message")
                End If

            End With
        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message, "ERROR")
        End Try

    End Sub

    Private Sub ButtonItem5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonItem5.Click, PrtBtn1.Click
        If save_excel = "" Then
            MessageBox.Show("Select Spread for Printing!!")
            Exit Sub
        End If

        If save_excel = "FpSpread1" Then
            If Spread_Print(Me.FpSpread1, "NO SHIP(Board) Summary", 0) = False Then
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

        Else
            MessageBox.Show("Select Spread for Saving!!")
        End If

    End Sub

    Private Sub FpSpread1_CellClick(ByVal sender As System.Object, ByVal e As FarPoint.Win.Spread.CellClickEventArgs) Handles FpSpread1.CellClick
        save_excel = "FpSpread1"
    End Sub

    Private Sub FpSpread1_Change(ByVal sender As Object, ByVal e As FarPoint.Win.Spread.ChangeEventArgs) Handles FpSpread1.Change
        Spread_Change(FpSpread1, e.Row)
    End Sub

End Class
