Public Class FrmSupplier

    Private Sub FrmSupplier_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        DockContainerItem1.Text = "실행 메뉴"
        DockContainerItem2.Text = "조회 조건"

        If Spread_Setting(FpSpread1, Me.Name) = True Then
            FpSpread1.ActiveSheet.RowCount = 0
            Spread_AutoCol(FpSpread1)
            FpSpread1.ActiveSheet.FrozenColumnCount = 1
            FpSpread1.ActiveSheet.Protect = False
        End If

        Formbim_Authority(Me.NewBtn, Me.Name, "NEW")
        Formbtn_Authority(Me.NewBtn1, Me.Name, "NEW")
        Formbim_Authority(Me.SaveBtn, Me.Name, "SAVE")
        Formbtn_Authority(Me.SaveBtn1, Me.Name, "SAVE")
        Formbim_Authority(Me.DelBtn, Me.Name, "DELETE")
        Formbtn_Authority(Me.DelBtn1, Me.Name, "DELETE")
        Formbim_Authority(Me.PrtBtn, Me.Name, "PRINT")
        Formbtn_Authority(Me.PrtBtn1, Me.Name, "PRINT")
        Formbim_Authority(Me.Excel, Me.Name, "EXCEL")
        Formbtn_Authority(Me.XlsBtn1, Me.Name, "EXCEL")
        Formbim_Authority(Me.ButtonItem1, Me.Name, "FIND")


    End Sub

    Private Sub FindBtn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonItem1.Click
        Try

            Dim qry As String = ""
            qry = qry & "select SUP_NO, SUP_NM, ceo_nm, REG_NO,  country_nm , state_nm, city_nm, ADRESS1 , ADRESS2, ZIP_CODE ,TELNO1, address_nm , TELNO2, FAXNO," & vbNewLine
            qry = qry & "			(select emp_nm from tbl_empmaster where emp_no = a.c_person) , c_date," & vbNewLine
            qry = qry & "			(select emp_nm from tbl_empmaster where emp_no = a.u_person) , c_date" & vbNewLine
            qry = qry & "from TBL_SUPMST  a" & vbNewLine
            qry = qry & "where site_id = '" & Site_id & "'" & vbNewLine

            qry = qry & "and sup_nm like '%" & PartNoTxt.Text & "%'" & vbNewLine
            qry = qry & "order by sup_nm" & vbNewLine
            qry = qry & "" & vbNewLine
            qry = qry & "" & vbNewLine

            If Query_Spread(FpSpread1, qry, 1) = True Then
                Spread_AutoCol(FpSpread1)
                With FpSpread1.ActiveSheet
                    .Cells(0, 0, .RowCount - 1, 0).Locked = True
                    '                    .Cells(0, 3, .RowCount - 1, 3).Locked = True

                End With

            End If

        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message, "ERROR")
        End Try

    End Sub

    Private Sub FpSpread1_CellDoubleClick(ByVal sender As Object, ByVal e As FarPoint.Win.Spread.CellClickEventArgs) Handles FpSpread1.CellDoubleClick
        Try

            If e.Column = 0 Then
                FpSpread1.ActiveSheet.Protect = True
                MessageBox.Show("거래처번호는 수정할 수 없습니다!!")
                Exit Sub

            Else
                FpSpread1.ActiveSheet.Protect = False
            End If

        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message, "ERROR")
        End Try
    End Sub

    Private Sub FpSpread1_Change(ByVal sender As Object, ByVal e As FarPoint.Win.Spread.ChangeEventArgs) Handles FpSpread1.Change
        Try


            Me.FpSpread1.ActiveSheet.Cells(e.Row, 16).Text = Emp_No
            Me.FpSpread1.ActiveSheet.Cells(e.Row, 17).Text = Now

            FpSpread1.ActiveSheet.Rows(e.Row).ForeColor = Color.OrangeRed
            FpSpread1.ActiveSheet.Protect = True

        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message, "ERROR")
        End Try

    End Sub

    Private Sub FpSpread1_LeaveCell(ByVal sender As Object, ByVal e As FarPoint.Win.Spread.LeaveCellEventArgs) Handles FpSpread1.LeaveCell
        Try
            FpSpread1.ActiveSheet.Protect = True
        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message, "ERROR")
        End Try
    End Sub


    Private Sub SaveBtn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SaveBtn1.Click, SaveBtn.Click
        Try
            Dim i As Integer
            Dim qry As String = ""

            With FpSpread1.ActiveSheet
                For i = 0 To .RowCount - 1
                    If .Rows(i).ForeColor = Color.OrangeRed Then
                        qry = ""
                        qry = qry & "if exists (select sup_no from tbl_supmst where sup_no = '" & .GetValue(i, 0) & "')" & vbNewLine
                        qry = qry & "begin" & vbNewLine
                        qry = qry & "   update tbl_supmsT " & vbNewLine
                        qry = qry & "   set sup_nm = '" & .GetValue(i, 1) & "'" & vbNewLine
                        qry = qry & ",      ceo_nm = '" & .GetValue(i, 2) & "'" & vbNewLine
                        qry = qry & ",      REG_NO = '" & .GetValue(i, 3) & "'" & vbNewLine
                        qry = qry & ",      country_nm = '" & .GetValue(i, 4) & "'" & vbNewLine
                        qry = qry & ",      state_nm = '" & .GetValue(i, 5) & "'" & vbNewLine
                        qry = qry & ",      city_nm = '" & .GetValue(i, 6) & "'" & vbNewLine
                        qry = qry & ",      ADRESS1 = '" & .GetValue(i, 7) & "'" & vbNewLine
                        qry = qry & ",      ADRESS2 = '" & .GetValue(i, 8) & "'" & vbNewLine
                        qry = qry & ",      ZIP_CODE = '" & .GetValue(i, 9) & "'" & vbNewLine
                        qry = qry & ",      TELNO1 = '" & .GetValue(i, 10) & "'" & vbNewLine
                        qry = qry & ",      address_nm = '" & .GetValue(i, 11) & "'" & vbNewLine
                        qry = qry & ",      TELNO2 = '" & .GetValue(i, 12) & "'" & vbNewLine
                        qry = qry & ",      FAXNO = '" & .GetValue(i, 13) & "'" & vbNewLine
                        qry = qry & ",      U_PERSON = '" & Emp_No & "'" & vbNewLine
                        qry = qry & ",      U_DATE = GETDATE()" & vbNewLine           
                        qry = qry & "where sup_no = '" & .GetValue(i, 0) & "'" & vbNewLine
                        qry = qry & "end" & vbNewLine
                        qry = qry & "ELSE" & vbNewLine
                        qry = qry & "BEGIN" & vbNewLine
                        Dim SUP_NO As String = Query_RS("SELECT 'S'+ CAST(DATEPART(YEAR, GETDATE()) AS VARCHAR(4)) + '-' + RIGHT('000'+CAST((ISNULL(CAST(RIGHT(MAX(SUP_NO),4) AS INT),  0) +1) AS VARCHAR(4)),4) FROM TBL_SUPMST WHERE SUP_NO LIKE 'S'+ CAST(DATEPART(YEAR, GETDATE()) AS VARCHAR(4)) + '%'")

                        qry = qry & "   INSERT INTO TBL_SUPMST VALUES ('" & Site_id & "','" & SUP_NO & "',"


                        qry = qry & "'" & .GetValue(i, 1) & "',"
                        qry = qry & "'" & .GetValue(i, 2) & "',"
                        qry = qry & "'" & .GetValue(i, 3) & "',"
                        qry = qry & "'" & .GetValue(i, 4) & "',"
                        qry = qry & "'" & .GetValue(i, 6) & "',"
                        qry = qry & "'" & .GetValue(i, 5) & "',"
                        qry = qry & "'" & .GetValue(i, 11) & "',"
                        qry = qry & "'" & .GetValue(i, 7) & "',"
                        qry = qry & "'" & .GetValue(i, 8) & "',"
                        qry = qry & "'" & .GetValue(i, 9) & "',"
                        qry = qry & "'" & .GetValue(i, 10) & "',"
                        qry = qry & "'" & .GetValue(i, 12) & "',"
                        qry = qry & "'" & .GetValue(i, 13) & "',"

                        qry = qry & "'" & Emp_No & "',GETDATE(), '" & Emp_No & "', GETDATE(), NULL,NULL,NULL,NULL,NULL)" & vbNewLine

                        qry = qry & "END" & vbNewLine
                        

                        If Insert_Data(qry) = True Then
                            .Rows(i).ForeColor = Color.Black
                        End If
                        Me.FpSpread1.ActiveSheet.Rows(i).ForeColor = Color.Black
                    End If
                Next
            End With

        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message, "ERROR")
        End Try
    End Sub


    Private Sub NewBtn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NewBtn1.Click, NewBtn.Click
        Try

            With Me.FpSpread1.ActiveSheet
                .RowCount += 1


                .Cells(.RowCount - 1, 14).Text = Emp_No
                .Cells(.RowCount - 1, 15).Text = Now
                .Cells(.RowCount - 1, 16).Text = Emp_No
                .Cells(.RowCount - 1, 17).Text = Now
                .Cells(.RowCount - 1, 0, .RowCount - 1, .ColumnCount - 1).Locked = False
                .Rows(.RowCount - 1).ForeColor = Color.OrangeRed

            End With

            FpSpread1.ActiveSheet.SetActiveCell(FpSpread1.ActiveSheet.RowCount - 1, 0)
            FpSpread1.ShowActiveCell(FarPoint.Win.Spread.VerticalPosition.Bottom, FarPoint.Win.Spread.HorizontalPosition.Left)


        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message, "ERROR")
        End Try

    End Sub



    Private Sub DelBtn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DelBtn1.Click, DelBtn.Click
        Try
            Dim r As DialogResult = MessageBox.Show("선택된 거래처를 삭제하시겠습니까?", "거래처 삭제", MessageBoxButtons.YesNo)
            If r = Windows.Forms.DialogResult.Yes Then
                With FpSpread1.ActiveSheet
                    If .ActiveRowIndex >= 0 Then
                        If Insert_Data("DELETE FROM TBL_SUPMST WHERE SITE_ID = '" & Site_id & "' AND SUP_NO = '" & .GetValue(.ActiveRowIndex, 0) & "'") = True Then
                        End If
                    End If
                End With
                MessageBox.Show("삭제되었습니다.", "Message")
                FindBtn_Click(Me.FpSpread1, System.EventArgs.Empty)
            End If
        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message, "ERROR")
        End Try
    End Sub

    Private Sub PrtBtn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PrtBtn1.Click, PrtBtn.Click
        If Spread_Print(Me.FpSpread1, "Repair Set", 1) = False Then
            MsgBox("Fail to Print")
        End If

    End Sub


    Private Sub Excel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles XlsBtn1.Click, Excel.Click
        File_Save(SaveFileDialog1, FpSpread1)
    End Sub




End Class