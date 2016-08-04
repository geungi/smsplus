Public Class FrmDefect

    Private RecCnt As Integer = 0
    Private chkFp As String

    Private Sub FrmDefect_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown, ClassCb.KeyDown
        If e.KeyValue = Keys.Enter Then
            FindBtn_Click(sender, e)
        End If
    End Sub

    Private Sub FrmDefect_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Me.DockContainerItem1.Text = "실행 메뉴"
        Me.DockContainerItem3.Text = "조회 조건"

        Bar1.Width = 180
        DockContainerItem1.Width = 180

        ModelCpDisp()

        If Spread_Setting(FpSpread1, Me.Name) = True Then
            FpSpread1.ActiveSheet.RowCount = 0
            Spread_AutoCol(FpSpread1) '스프레드의 컬럼사이즈를 셀의 텍스트에 맞게 자동 설정
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
        Formbim_Authority(Me.FindBtn, Me.Name, "FIND")

    End Sub
    Private Sub ModelCpDisp()
        Try
            If Query_Combo(Me.ClassCb, "select code_id from tbl_defect where site_id = '" & Site_id & "' order by code_id ") = True Then
                ClassCb.Items.Add("ALL")
            End If

            Me.ClassCb.Text = "ALL"

        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message, "ERROR")
        End Try
    End Sub
    Private Sub FindBtn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FindBtn.Click
        Try
            Dim Act_YN As String = ""
            Dim CodeVal As String = ""
            Dim ClassVal As String = ""
            Dim qry As String = ""

            qry = "SELECT SITE_id, code_id, code_NAME, priority, ACTIVE, " & vbNewLine
            qry = qry & "     ISNULL((SELECT EMP_NM FROM TBL_EMPMASTER WHERE SITE_ID = A.SITE_ID AND EMP_NO = A.C_PERSON),'ADMIN'), C_DATE," & vbNewLine
            qry = qry & "     ISNULL((SELECT EMP_NM FROM TBL_EMPMASTER WHERE SITE_ID = A.SITE_ID AND EMP_NO = A.U_PERSON),'ADMIN'), U_DATE" & vbNewLine
            qry = qry & "FROM TBL_defect A " & vbNewLine
            qry = qry & "WHERE SITE_ID = '" & Site_id & "'" & vbNewLine

            If ClassCb.Text <> "ALL" Then
                qry = qry & "  AND code_id = '" & ClassCb.Text & "'" & vbNewLine
            End If

            If ActCb.Text <> "ALL" Then
                qry = qry & "  AND ACTIVE = '" & ActCb.Text & "'" & vbNewLine
            End If

            qry = qry & "ORDER BY code_id" & vbNewLine

            If Query_Spread(FpSpread1, qry, 1) = True Then
                Spread_AutoCol(FpSpread1) '스프레드에 데이터를 출력후, 데이터의 사이즈에 맞게 컬럼사이즈 재조정
            End If

            MessageBox.Show("조회되었습니다.", "Message")

        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message, "ERROR")
        End Try
    End Sub

    Private Sub FpSpread1_CellClick(ByVal sender As Object, ByVal e As FarPoint.Win.Spread.CellClickEventArgs) Handles FpSpread1.CellClick
        save_excel = "FpSpread1"
    End Sub

    Private Sub FpSpread1_CellDoubleClick(ByVal sender As Object, ByVal e As FarPoint.Win.Spread.CellClickEventArgs) Handles FpSpread1.CellDoubleClick
        Try

            Select Case e.Column
                Case 2
                    Me.FpSpread1.ActiveSheet.Cells(e.Row, e.Column).Locked = False
                    Spread_Celltype(Me.FpSpread1, Me.Name, e.Row, e.Column)
                Case 3
                    Me.FpSpread1.ActiveSheet.Cells(e.Row, e.Column).Locked = False
                    'Spread_Celltype(Me.FpSpread1, Me.Name, e.Row, e.Column)
                Case 4
                    Me.FpSpread1.ActiveSheet.Cells(e.Row, e.Column).Locked = False
                    Chg_ComboCell(FpSpread1, e.Row, 6, New String() {"Y", "N"})
            End Select

        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message, "ERROR")
        End Try
    End Sub

    Private Sub FpSpread1_Change(ByVal sender As Object, ByVal e As FarPoint.Win.Spread.ChangeEventArgs) Handles FpSpread1.Change
        Try

            Spread_Change(FpSpread1, e.Row)
            Me.FpSpread1.ActiveSheet.Cells(e.Row, 7).Text = Emp_No
            Me.FpSpread1.ActiveSheet.Cells(e.Row, 8).Text = Now

        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message, "ERROR")
        End Try
    End Sub

    Private Sub SaveBtn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SaveBtn1.Click, SaveBtn.Click, bSave.Click
        Dim i As Integer = 0
        Dim j As Integer = 0
        Dim rowCnt As Integer = FpSpread1.ActiveSheet.RowCount

        If rowCnt <= 0 Then
            Exit Sub
        End If

        '셀에 내용변경후, 포커스가 다른 셀로 이동해야 셀에 값이 변경되기 때문에, 셀의 포커스 변경
        Me.FpSpread1.ActiveSheet.SetActiveCell(0, 0)

        Try
            For i = 0 To rowCnt - 1

                With FpSpread1.ActiveSheet
                    If .Rows(i).ForeColor = Color.OrangeRed Then
                        If Insert_Data("EXEC SP_FrmDefect_SAVE '" _
                                        & Site_id & "', '" _
                                        & .Cells(i, 1).Text & "', '" _
                                        & .Cells(i, 2).Text & "', " _
                                        & .Cells(i, 3).Text & ", '" _
                                        & .Cells(i, 4).Text & "', '" _
                                        & Emp_No & "'") = True Then
                            .Rows(i).ForeColor = Color.Black
                        End If
                    End If
                End With
            Next
            ModelCpDisp()

            Spread_AutoCol(Me.FpSpread1)

            ' MessageBox.Show("저장이 완료되었습니다.", "Message")
            FindBtn_Click(Me.FpSpread1, System.EventArgs.Empty)

        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message, "ERROR")
        End Try
    End Sub

    Private Sub NewBtn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NewBtn1.Click, NewBtn.Click, bNew.Click
        Try

            With Me.FpSpread1.ActiveSheet
                .RowCount += 1

                Spread_Celltype(Me.FpSpread1, Me.Name, .RowCount - 1, 1)
                Spread_Celltype(Me.FpSpread1, Me.Name, .RowCount - 1, 2)
                Spread_Celltype(Me.FpSpread1, Me.Name, .RowCount - 1, 3)
                Spread_Celltype(Me.FpSpread1, Me.Name, .RowCount - 1, 4)

                .Cells(.RowCount - 1, 0).Value = Site_id

                Chg_ComboCell(FpSpread1, .RowCount - 1, 4, New String() {"Y", "N"})
                .Cells(.RowCount - 1, 4).Text = "Y"
                .Cells(.RowCount - 1, 5).Text = Emp_No
                .Cells(.RowCount - 1, 6).Text = Now
                .Cells(.RowCount - 1, 7).Text = Emp_No
                .Cells(.RowCount - 1, 8).Text = Now

                Dim i As Integer
                For i = 1 To .ColumnCount - 5
                    .Cells(.RowCount - 1, i).Locked = False
                Next

                .SetActiveCell(.RowCount - 1, 0)
                FpSpread1.ShowActiveCell(FarPoint.Win.Spread.VerticalPosition.Bottom, FarPoint.Win.Spread.HorizontalPosition.Left) 'ActiveCell로 스크롤 자동이동

            End With

        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message, "ERROR")
        End Try

    End Sub



    Private Sub DelBtn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DelBtn1.Click, DelBtn.Click, bDel.Click
        Dim RowCnt = Me.FpSpread1.ActiveSheet.RowCount
        Dim answer As DialogResult

        Dim cnt_index As Integer = Me.FpSpread1.ActiveSheet.ActiveRowIndex
        Dim site_id As String = ""
        Dim class_id As String = ""

        Try
            site_id = FpSpread1.ActiveSheet.GetText(cnt_index, 0)
            class_id = FpSpread1.ActiveSheet.GetText(cnt_index, 1)

            If (site_id <> "") And (class_id <> "") Then

                answer = MessageBox.Show("선택된 행을 삭제하시겠습니까?", "Selected Rows Delete", MessageBoxButtons.YesNo)

                If answer = Windows.Forms.DialogResult.Yes Then
                    Insert_Data("delete from TBL_defect where site_ID = '" & site_id & "' and code_id = '" & class_id & "'")
                    MessageBox.Show("삭제되었습니다.")
                    FpSpread1.ActiveSheet.RemoveRows(cnt_index, 1)
                Else
                    Return
                End If

            Else
                FpSpread1.ActiveSheet.RemoveRows(cnt_index, 1)
            End If

        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message, "ERROR")
        End Try

        Spread_AutoCol(FpSpread1)

    End Sub

    Private Sub PrtBtn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PrtBtn1.Click, PrtBtn.Click, bPrint.Click
        If Spread_Print(Me.FpSpread1, "Code Master", 1) = False Then
            MsgBox("Fail to Print")
        End If
    End Sub

    Private Sub Excel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles XlsBtn1.Click, Excel.Click, bExcel.Click
        File_Save(SaveFileDialog1, FpSpread1)
    End Sub




End Class