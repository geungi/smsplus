Public Class FrmBuying

    Private RecCnt As Integer = 0
    Private chkFp As String

    Private Sub FrmSTCodeMst_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        If e.KeyValue = Keys.Enter Then
            FindBtn_Click(sender, e)
        End If
    End Sub

    Private Sub FrmSTCodeMst_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Me.DockContainerItem1.Text = "실행 메뉴"
        Me.DockContainerItem4.Text = "조회 조건"

        DateTimeInput1.Value = Now
        DateTimeInput2.Value = Now

        If Query_Combo(Me.ComboBoxEx1, "SELECT sup_nm FROM TBL_supmst WHERE site_id = '" & Site_id & "' and sup_no like 'P%' ORDER BY SUP_no") = True Then
        End If
        Me.ComboBoxEx1.Items.Add("ALL")
        Me.ComboBoxEx1.Text = "ALL"

        If Query_Combo(Me.ComboBoxEx2, "SELECT CODE_NAME FROM tbl_CODEmaster WHERE site_id = '" & Site_id & "' and CLASS_ID = 'R0009' AND active = 'Y' ORDER BY DIS_ORDER") = True Then
        End If
        Me.ComboBoxEx2.Items.Add("ALL")
        Me.ComboBoxEx2.Text = "ALL"

        If Query_Combo(Me.ComboBoxEx3, "SELECT model_no FROM TBL_MODELMASTER WHERE site_id = '" & Site_id & "' and active = 'Y' ORDER BY model_no") = True Then
        End If
        Me.ComboBoxEx3.Items.Add("ALL")
        Me.ComboBoxEx3.Text = "ALL"

        Bar1.Width = 180
        DockContainerItem1.Width = 180

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

    Private Sub FindBtn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FindBtn.Click
        Try
            Dim Act_YN As String = ""
            Dim CodeVal As String = ""
            Dim ClassVal As String = ""
            Dim qry As String = ""

            qry = "SELECT BUY_NO, (SELECT SUP_NM FROM TBL_SUPMST WHERE SUP_NO = A.SUP_NO), BUY_DT, (SELECT CODE_NAME FROM TBL_CODEMASTER WHERE CLASS_ID = 'R0009' AND CODE_ID = A.BUY_TYPE), ISNULL(MODEL,''),  " & vbNewLine
            qry = qry & "   ISNULL(TOT_QTY,0), ISNULL(GOOD_QTY,0), ISNULL(BAD_QTY,0), ISNULL(price,0), ISNULL(amount,0), PAY_DT, ISNULL(PAY_AMT,0), ISNULL(REMARK,'')," & vbNewLine
            qry = qry & "     ISNULL((SELECT EMP_NM FROM TBL_EMPMASTER WHERE SITE_ID = A.SITE_ID AND EMP_NO = A.C_PERSON),'ADMIN'), C_DATE," & vbNewLine
            qry = qry & "     ISNULL((SELECT EMP_NM FROM TBL_EMPMASTER WHERE SITE_ID = A.SITE_ID AND EMP_NO = A.U_PERSON),'ADMIN'), U_DATE" & vbNewLine
            qry = qry & "FROM TBL_BUYING A " & vbNewLine
            qry = qry & "WHERE SITE_ID = '" & Site_id & "'" & vbNewLine
            qry = qry & "   AND BUY_DT BETWEEN '" & DateTimeInput1.Text & "' AND '" & DateTimeInput2.Text & "'" & vbNewLine
            If ComboBoxEx1.Text <> "ALL" Then
                qry = qry & "  AND SUP_NO = (SELECT SUP_NO FROM TBL_SUPMST WHERE SUP_NM = '" & ComboBoxEx1.Text & "')" & vbNewLine
            End If

            If ComboBoxEx2.Text <> "ALL" Then
                qry = qry & "  AND BUY_TYPE = (SELECT CODE_ID FROM TBL_CODEMASTER WHERE CLASS_ID = 'R0009' AND CODE_NAME = '" & ComboBoxEx2.Text & "')" & vbNewLine
            End If

            If ComboBoxEx3.Text <> "ALL" Then
                qry = qry & "  AND MODEL = '" & ComboBoxEx3.Text & "'" & vbNewLine
            End If

            qry = qry & "ORDER BY BUY_NO, BUY_DT" & vbNewLine

            If Query_Spread(FpSpread1, qry, 1) = True Then
                Spread_AutoCol(FpSpread1) '스프레드에 데이터를 출력후, 데이터의 사이즈에 맞게 컬럼사이즈 재조정
            End If

            MessageBox.Show("조회가 완료되었습니다.", "Message")

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
                Case 4
                    Me.FpSpread1.ActiveSheet.Cells(e.Row, e.Column).Locked = False
                    Spread_Celltype(Me.FpSpread1, Me.Name, e.Row, e.Column)
                Case 5
                    Me.FpSpread1.ActiveSheet.Cells(e.Row, e.Column).Locked = False
                    Spread_Celltype(Me.FpSpread1, Me.Name, e.Row, e.Column)
                Case 6
                    Me.FpSpread1.ActiveSheet.Cells(e.Row, e.Column).Locked = False
                Case 7
                    Me.FpSpread1.ActiveSheet.Cells(e.Row, e.Column).Locked = False
                    Spread_Celltype(Me.FpSpread1, Me.Name, e.Row, e.Column)
                Case 8
                    Me.FpSpread1.ActiveSheet.Cells(e.Row, e.Column).Locked = False
                    Spread_Celltype(Me.FpSpread1, Me.Name, e.Row, e.Column)
                Case 9
                    Me.FpSpread1.ActiveSheet.Cells(e.Row, e.Column).Locked = False
                    Spread_Celltype(Me.FpSpread1, Me.Name, e.Row, e.Column)
                Case 10
                    Me.FpSpread1.ActiveSheet.Cells(e.Row, e.Column).Locked = False
                    Me.FpSpread1.ActiveSheet.Cells(e.Row, e.Column).CellType = datecell
                Case 11
                    Me.FpSpread1.ActiveSheet.Cells(e.Row, e.Column).Locked = False
                    Me.FpSpread1.ActiveSheet.Cells(e.Row, e.Column).CellType = curcell
                Case 12
                    Me.FpSpread1.ActiveSheet.Cells(e.Row, e.Column).Locked = False
                    Spread_Celltype(Me.FpSpread1, Me.Name, e.Row, e.Column)
            End Select

            With FpSpread1.ActiveSheet
                Spread_Change(FpSpread1, e.Row)
                Spread_AutoCol(FpSpread1)
            End With

        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message, "ERROR")
        End Try
    End Sub

    Private Sub FpSpread1_Change(ByVal sender As Object, ByVal e As FarPoint.Win.Spread.ChangeEventArgs) Handles FpSpread1.Change
        Try

            Spread_Change(FpSpread1, e.Row)
            Me.FpSpread1.ActiveSheet.Cells(e.Row, 15).Text = Emp_No
            Me.FpSpread1.ActiveSheet.Cells(e.Row, 16).Text = Now

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
                        If Insert_Data("EXEC SP_FrmBUYING_SAVE '" _
                                        & Site_id & "', '" _
                                        & .Cells(i, 0).Text & "', '" _
                                        & .Cells(i, 1).Text & "', '" _
                                        & .Cells(i, 2).Text & "', '" _
                                        & .Cells(i, 3).Text & "', '" _
                                        & .Cells(i, 4).Value & "', " _
                                        & .Cells(i, 5).Value & ", " _
                                        & .Cells(i, 6).Value & ", " _
                                        & .Cells(i, 7).Value & ", " _
                                        & .Cells(i, 8).Value & ", " _
                                        & .Cells(i, 9).Value & ", '" _
                                        & .Cells(i, 10).Value & "', " _
                                        & .Cells(i, 11).Value & ", '" _
                                        & .Cells(i, 12).Value & "', '" _
                                        & Emp_No & "'") = True Then
                            .Rows(i).ForeColor = Color.Black
                        End If
                    End If
                End With
            Next

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
                .Rows(.RowCount - 1).ForeColor = Color.OrangeRed

                Spread_Celltype(Me.FpSpread1, Me.Name, .RowCount - 1, 1)
                .Cells(.RowCount - 1, 2).CellType = datecell
                .Cells(.RowCount - 1, 10).CellType = datecell

                .Cells(.RowCount - 1, 5, .RowCount - 1, 7).CellType = intcell
                .Cells(.RowCount - 1, 8, .RowCount - 1, 9).CellType = curcell
                .Cells(.RowCount - 1, 11).CellType = curcell

                .Cells(.RowCount - 1, 0).Value = ""

                Chg_ComboCell(FpSpread1, .RowCount - 1, 1, Query_Cell_Code2("select sup_nm from tbl_supmst where sup_no like 'P%' ORDER BY SUP_NO"))
                'Chg_ComboCell(FpSpread1, .RowCount - 1, 2, Query_Cell_Code2("select GETDATE()"))
                Chg_ComboCell(FpSpread1, .RowCount - 1, 3, Query_Cell_Code1("CODE_NAME", "R0009"))

                .Cells(.RowCount - 1, 5).Text = 0
                .Cells(.RowCount - 1, 6).Text = 0
                .Cells(.RowCount - 1, 7).Text = 0
                .Cells(.RowCount - 1, 8).Text = 0
                .Cells(.RowCount - 1, 9).Text = 0
                .Cells(.RowCount - 1, 11).Text = 0


                .Cells(.RowCount - 1, 13).Text = Emp_No
                .Cells(.RowCount - 1, 14).Text = Now
                .Cells(.RowCount - 1, 15).Text = Emp_No
                .Cells(.RowCount - 1, 16).Text = Now

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

        Try

            answer = MessageBox.Show("선택된 행을 삭제하시겠습니까?", "Selected Rows Delete", MessageBoxButtons.YesNo)

            If answer = Windows.Forms.DialogResult.Yes Then
                Insert_Data("delete from TBL_BUYING where site_ID = '" & Site_id & "' and BUY_NO = '" & FpSpread1.ActiveSheet.GetText(cnt_index, 0) & "'")
                MessageBox.Show("삭제되었습니다.")
                FpSpread1.ActiveSheet.RemoveRows(cnt_index, 1)
            Else
                Return
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