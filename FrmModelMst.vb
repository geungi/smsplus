Public Class FrmModelMst

    Protected FpSpread As FarPoint.Win.Spread.FpSpread
    Private RecCnt As Integer

    Private Sub FrmModelMst_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Me.DockContainerItem1.Text = "실행 메뉴"
        Me.DockContainerItem3.Text = "조회 조건"

        Me.ModelCb.Text = "ALL"
        ModelCpDisp() 'Modelcb에 item 추가
        Me.ModelCb.Items.Add("ALL")

        If Spread_Setting(FpSpread1, "FrmModelMst") = True Then
            Spread_AutoCol(FpSpread1)
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

            Dim qry As String

            qry = "select site_id, model_no, model_name, model_desc, model_dv, color, op_cd, sw_ver, " & vbNewLine
            qry = qry & "   ISNULL((SELECT '['+CODE_ID+'] '+CODE_NAME FROM TBL_CODEMASTER WHERE SITE_ID = A.SITE_ID AND CLASS_ID = 'R3001' AND CODE_ID = A.VENDOR_CD),'')," & vbNewLine
            qry = qry & "   ISNULL((SELECT '['+CODE_ID+'] '+CODE_NAME FROM TBL_CODEMASTER WHERE SITE_ID = A.SITE_ID AND CLASS_ID = 'R3000' AND CODE_ID = A.CUST_CD),'')," & vbNewLine
            qry = qry & "   rep_lev, active, RESERV1, RESERV2, RESERV3, RESERV4, RESERV5," & vbNewLine
            qry = qry & "isnull(phoneType,'')," & vbNewLine
            qry = qry & "isnull(phoneOwnership,'')," & vbNewLine
            qry = qry & "isnull(transactionType,'')," & vbNewLine
            qry = qry & "isnull(poOrder,'')," & vbNewLine
            qry = qry & "isnull(locationDestination,'')," & vbNewLine
            qry = qry & "isnull(uedfRevisionNumber,'')," & vbNewLine
            qry = qry & "isnull(edfSerialType,'')," & vbNewLine
            qry = qry & "isnull(equipType,'')," & vbNewLine
            qry = qry & "isnull(manufId,'')," & vbNewLine
            qry = qry & "isnull(manufName,'')," & vbNewLine
            qry = qry & "isnull(prl,'')," & vbNewLine
            qry = qry & "isnull(cardSku,'')," & vbNewLine
            qry = qry & "isnull(reserv6,'')," & vbNewLine
            qry = qry & "isnull(reserv7,'')," & vbNewLine
            qry = qry & "isnull(reserv8,'')," & vbNewLine
            qry = qry & "isnull(reserv9,'')," & vbNewLine
            qry = qry & "     ISNULL((SELECT EMP_NM FROM TBL_EMPMASTER WHERE SITE_ID = A.SITE_ID AND EMP_NO = A.C_PERSON),'ADMIN'), C_DATE," & vbNewLine
            qry = qry & "     ISNULL((SELECT EMP_NM FROM TBL_EMPMASTER WHERE SITE_ID = A.SITE_ID AND EMP_NO = A.U_PERSON),'ADMIN'), U_DATE" & vbNewLine
            qry = qry & "FROM TBL_MODELMASTER A " & vbNewLine
            qry = qry & "WHERE SITE_ID = '" & Site_id & "'" & vbNewLine
            If ModelCb.Text <> "ALL" Then
                qry = qry & "  AND MODEL_NO = '" & ModelCb.Text & "'" & vbNewLine
            End If

            If ActCb.Text <> "ALL" Then
                qry = qry & "  AND ACTIVE = '" & Mid(ActCb.Text, 1, 1) & "'" & vbNewLine
            End If
            qry = qry & "ORDER BY MODEL_NO" & vbNewLine

            If Query_Spread(FpSpread1, qry, 1) = True Then
                Spread_AutoCol(FpSpread1) '스프레드에 데이터를 출력후, 데이터의 사이즈에 맞게 컬럼사이즈 재조정
            End If

        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message, "ERROR")
        End Try

    End Sub

    Private Sub FpSpread1_CellDoubleClick(ByVal sender As Object, ByVal e As FarPoint.Win.Spread.CellClickEventArgs) Handles FpSpread1.CellDoubleClick
        Try

            If e.Column > 0 And e.Column < 33 Then
                Me.FpSpread1.ActiveSheet.Cells(e.Row, e.Column).Locked = False
            End If

            Select Case e.Column
                Case 1
                    Spread_Celltype(Me.FpSpread1, Me.Name, e.Row, e.Column)
                Case 2
                    Spread_Celltype(Me.FpSpread1, Me.Name, e.Row, e.Column)
                Case 3
                    Spread_Celltype(Me.FpSpread1, Me.Name, e.Row, e.Column)
                Case 4
                    Chg_ComboCell(FpSpread1, Me.FpSpread1.ActiveSheet.ActiveRowIndex, 4, Query_Cell_Code1("code_id", "R0036")) '코드마스터에서 모델구분을 셀에 콤보로 표시
                Case 5
                    FpSpread1.ActiveSheet.Cells(e.Row, e.Column).Locked = False
                Case 6
                    Chg_ComboCell(FpSpread1, Me.FpSpread1.ActiveSheet.ActiveRowIndex, 6, Query_Cell_Code1("code_id", "10000")) '코드마스터에서 통신사관련코드를 셀에 콤보로 표시
                Case 7
                    Chg_ComboCell(FpSpread1, Me.FpSpread1.ActiveSheet.ActiveRowIndex, 7, Query_Cell_Code1("code_name", "10000")) '코드마스터에서 통신사관련코드를 셀에 콤보로 표시
                Case 8
                    Chg_ComboCell(FpSpread1, Me.FpSpread1.ActiveSheet.ActiveRowIndex, 8, Query_Cell_Code1("'['+code_id+'] '+code_name", "R3001")) '코드마스터에서 통신사관련코드를 셀에 콤보로 표시
                Case 9
                    Chg_ComboCell(FpSpread1, Me.FpSpread1.ActiveSheet.ActiveRowIndex, 9, Query_Cell_Code1("'['+code_id+'] '+code_name", "R3000")) '코드마스터에서 제조사관련코드를 셀에 콤보로 표시
                Case 10
                    Chg_ComboCell(FpSpread1, Me.FpSpread1.ActiveSheet.ActiveRowIndex, 10, Query_Cell_Code2("select distinct level_id from tbl_modellevel where site_id ='" & Site_id & "'")) 'level_id를 셀에 콤보로 표시
                Case 11
                    Chg_ComboCell(FpSpread1, Me.FpSpread1.ActiveSheet.ActiveRowIndex, 11, New String() {"Y", "N"})
            End Select

        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message, "ERROR")
        End Try
    End Sub

    Private Sub FpSpread1_Change(ByVal sender As Object, ByVal e As FarPoint.Win.Spread.ChangeEventArgs) Handles FpSpread1.Change
        Try
            Spread_Change(FpSpread1, e.Row)
            Me.FpSpread1.ActiveSheet.Cells(e.Row, 35).Text = Emp_No
            Me.FpSpread1.ActiveSheet.Cells(e.Row, 36).Text = Now
            
            Spread_AutoCol(FpSpread1)

        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message, "ERROR")
        End Try

    End Sub

    Private Sub SaveBtn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SaveBtn1.Click, SaveBtn.Click, bSave.Click
        Try
            If FpSpread1.ActiveSheet.RowCount < 1 Then
                Exit Sub
            End If

            Me.FpSpread1.ActiveSheet.SetActiveCell(0, 0)
            With FpSpread1.ActiveSheet
                For i As Integer = 0 To .RowCount - 1
                    With FpSpread1.ActiveSheet
                        If .Rows(i).ForeColor = Color.OrangeRed Then
                            If Insert_Data("EXEC SP_FRMMODELMST_SAVE '" _
                                            & Site_id & "', '" _
                                            & .Cells(i, 1).Text & "', '" _
                                            & .Cells(i, 2).Text & "', '" _
                                            & .Cells(i, 3).Text & "', '" _
                                            & .Cells(i, 4).Text & "', '" _
                                            & .Cells(i, 5).Text & "', '" _
                                            & .Cells(i, 6).Text & "', '" _
                                            & .Cells(i, 7).Text & "', '" _
                                            & Mid(.Cells(i, 8).Text, 2, 5) & "', '" _
                                            & Mid(.Cells(i, 9).Text, 2, 5) & "', '" _
                                            & .Cells(i, 10).Text & "', '" _
                                            & .Cells(i, 11).Text & "', '" _
                                            & .Cells(i, 12).Text & "', '" _
                                            & .Cells(i, 13).Text & "', '" _
                                            & .Cells(i, 14).Text & "', '" _
                                            & .Cells(i, 15).Text & "', '" _
                                            & .Cells(i, 16).Text & "', '" _
                                            & .Cells(i, 17).Text & "', '" _
                                            & .Cells(i, 18).Text & "', '" _
                                            & .Cells(i, 19).Text & "', '" _
                                            & .Cells(i, 20).Text & "', '" _
                                            & .Cells(i, 21).Text & "', '" _
                                            & .Cells(i, 22).Text & "', '" _
                                            & .Cells(i, 23).Text & "', '" _
                                            & .Cells(i, 24).Text & "', '" _
                                            & .Cells(i, 25).Text & "', '" _
                                            & .Cells(i, 26).Text & "', '" _
                                            & .Cells(i, 27).Text & "', '" _
                                            & .Cells(i, 28).Text & "', '" _
                                            & .Cells(i, 29).Text & "', '" _
                                            & .Cells(i, 30).Text & "', '" _
                                            & .Cells(i, 31).Text & "', '" _
                                            & .Cells(i, 32).Text & "', '" _
                                            & Emp_No & "'") = True Then
                                .Rows(i).ForeColor = Color.Black
                            End If
                        End If
                    End With
                Next
            End With

            ModelCpDisp()
            Spread_AutoCol(Me.FpSpread1)

            ' MessageBox.Show("저장이 완료되었습니다.", "Message")
            MessageBox.Show("저장되었습니다", "Message")
            FindBtn_Click(Me.FpSpread1, System.EventArgs.Empty)

        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message, "ERROR")
        End Try
    End Sub


    Private Sub NewBtn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NewBtn1.Click, NewBtn.Click, bNew.Click
        Try

            With Me.FpSpread1.ActiveSheet
                .RowCount += 1

                '셀타입을 text셀로, tbl_spread테이블에 있는 해당 스프레드의 컬럼의 속성(maxlength,대소문자구분)을 셀속성으로 설정
                Spread_Celltype(Me.FpSpread1, Me.Name, .RowCount - 1, 1)
                Spread_Celltype(Me.FpSpread1, Me.Name, .RowCount - 1, 2)
                Spread_Celltype(Me.FpSpread1, Me.Name, .RowCount - 1, 3)
                Spread_Celltype(Me.FpSpread1, Me.Name, .RowCount - 1, 5)
                Spread_Celltype(Me.FpSpread1, Me.Name, .RowCount - 1, 9)


                .Cells(.RowCount - 1, 0).Text = Site_id
                Chg_ComboCell(FpSpread1, .RowCount - 1, 4, Query_Cell_Code1("code_id", "R0036"))
                Chg_ComboCell(FpSpread1, .RowCount - 1, 6, Query_Cell_Code1("code_id", "10000"))
                Chg_ComboCell(FpSpread1, .RowCount - 1, 7, Query_Cell_Code1("code_name", "10000"))
                Chg_ComboCell(FpSpread1, Me.FpSpread1.ActiveSheet.ActiveRowIndex, 8, Query_Cell_Code1("'['+code_id+'] '+code_name", "R3001")) '코드마스터에서 고객관련코드를 셀에 콤보로 표시
                Chg_ComboCell(FpSpread1, Me.FpSpread1.ActiveSheet.ActiveRowIndex, 9, Query_Cell_Code1("'['+code_id+'] '+code_name", "R3000")) '코드마스터에서 제조사관련코드를 셀에 콤보로 표시
                Chg_ComboCell(FpSpread1, .RowCount - 1, 10, Query_Cell_Code2("select distinct level_id from tbl_modellevel where site_id ='" & Site_id & "'"))
                Chg_ComboCell(FpSpread1, .RowCount - 1, 11, New String() {"Y", "N"})
                .Cells(.RowCount - 1, 11).Text = "Y"
                .Cells(.RowCount - 1, 33).Text = Emp_No
                .Cells(.RowCount - 1, 34).Text = Now
                .Cells(.RowCount - 1, 35).Text = Emp_No
                .Cells(.RowCount - 1, 36).Text = Now
                Dim i As Integer
                For i = 1 To .ColumnCount - 5
                    .Cells(.RowCount - 1, i).Locked = False
                Next
            End With

            FpSpread1.ActiveSheet.SetActiveCell(FpSpread1.ActiveSheet.RowCount - 1, 0)
            FpSpread1.ShowActiveCell(FarPoint.Win.Spread.VerticalPosition.Bottom, FarPoint.Win.Spread.HorizontalPosition.Left) 'ShowActivecell은 지정된 위치로 스크롤을 자동 이동시킴

        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message, "ERROR")
        End Try
    End Sub

    Private Sub ModelCpDisp()
        Try
            If Query_Combo(Me.ModelCb, "SELECT model_no FROM tbl_modelmaster WHERE site_id = '" & Site_id & "' ORDER BY model_no") = True Then
            End If

        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message, "ERROR")
        End Try
    End Sub

    Private Sub DelBtn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DelBtn1.Click, DelBtn.Click, bDel.Click
        Try
            Dim RowCnt = Me.FpSpread1.ActiveSheet.RowCount
            Dim r As DialogResult = MessageBox.Show("Selected row delete now?", "Selected Rows Delete", MessageBoxButtons.YesNo)
            If r = Windows.Forms.DialogResult.Yes Then

                Insert_Data("delete from TBL_MODELMASTER where site_ID = '" & Site_id & "' and MODEL_NO = '" & FpSpread1.ActiveSheet.GetValue(FpSpread1.ActiveSheet.ActiveRowIndex, 1) & "'")
                MessageBox.Show("삭제되었습니다.")
                FpSpread1.ActiveSheet.RemoveRows(FpSpread1.ActiveSheet.ActiveRowIndex, 1)

            End If
        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message, "ERROR")
        End Try
    End Sub

    Private Sub PrtBtn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PrtBtn1.Click, PrtBtn.Click, bPrint.Click

        If Spread_Print(Me.FpSpread1, "Model Master", 1) = False Then
            MsgBox("Fail to Print")
        End If

    End Sub

    Private Sub Excel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles XlsBtn1.Click, Excel.Click, bExcel.Click
        File_Save(SaveFileDialog1, FpSpread1)
    End Sub

End Class