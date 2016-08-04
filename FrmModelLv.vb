Public Class FrmModelLv


    Private RecCnt As Integer
    Private ChkNew As Boolean

    Private Sub FrmModelLV_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Me.DockContainerItem1.Text = "실행 메뉴"
        Me.DockContainerItem3.Text = "조회 조건"

        Me.ModelLvCb.Text = "ALL"
        ModelCpDisp()
        Me.ModelLvCb.Items.Add("ALL")


        If Spread_Setting(FpSpread1, "FrmModelLv") = True Then
            Spread_AutoCol(FpSpread1)
        End If

        If Query_Combo(Me.ClassCb, "select MODEL_NO from tbl_MODELMASTER where site_id = '" & Site_id & "' AND ACTIVE = 'Y' order by MODEL_NO") = True Then
            ClassCb.Items.Add("ALL")
        End If

        Me.ClassCb.Text = "ALL"

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
        Formbtn_Authority(Me.FindBtn, Me.Name, "FIND")

    End Sub


    Private Sub FindBtn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FindBtn.Click
        Try

            Dim qry As String = ""

            qry = "select a.SITE_ID, model, SERVICE_TYPE, SERVICE_NAME, RETURN_STATE, PRICE," & vbNewLine
            qry = qry & "   (SELECT EMP_NM FROM TBL_EMPMASTER WHERE EMP_NO = A.C_PERSON), C_DATE," & vbNewLine
            qry = qry & "   (SELECT EMP_NM FROM TBL_EMPMASTER WHERE EMP_NO = A.U_PERSON), U_DATE" & vbNewLine
            qry = qry & "from tbl_modelLEVEL A" & vbNewLine
            qry = qry & "WHERE SITE_ID = '" & Site_id & "'" & vbNewLine

            If ClassCb.Text <> "ALL" Then
                qry = qry & "  and model = '" & ClassCb.Text & "'" & vbNewLine

            End If

            If ModelLvCb.Text <> "ALL" Then
                qry = qry & "  AND SERVICE_TYPE = '" & ModelLvCb.Text & "'" & vbNewLine
            End If

            qry = qry & "ORDER BY MODEL" & vbNewLine

            If Query_Spread(FpSpread1, qry, 1) = True Then
                Spread_AutoCol(FpSpread1) '스프레드에 데이터를 출력후, 데이터의 사이즈에 맞게 컬럼사이즈 재조정
            End If

        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message, "ERROR")
        End Try
    End Sub

    Private Sub FpSpread1_CellDoubleClick(ByVal sender As Object, ByVal e As FarPoint.Win.Spread.CellClickEventArgs) Handles FpSpread1.CellDoubleClick
        Try
            Select Case e.Column
                Case 1
                    Me.FpSpread1.ActiveSheet.Cells(e.Row, e.Column).Locked = False
                    '셀타입을 text셀로, tbl_spread테이블에 있는 해당 스프레드의 컬럼의 속성(maxlength,대소문자구분)을 셀속성으로 설정
                    Spread_Celltype(Me.FpSpread1, Me.Name, e.Row, e.Column)
                Case 2
                    Me.FpSpread1.ActiveSheet.Cells(e.Row, e.Column).Locked = False
                    Chg_ComboCell(FpSpread1, Me.FpSpread1.ActiveSheet.ActiveRowIndex, 2, Query_Cell_Code1("code_id", "10002")) 'Service type 콤보처리
                Case 3
                    Me.FpSpread1.ActiveSheet.Cells(e.Row, e.Column).Locked = False
                    '기존에 들어있는 데이터와 코드마스터의 데이터가 약간 다름 
                    Chg_ComboCell(FpSpread1, Me.FpSpread1.ActiveSheet.ActiveRowIndex, 3, Query_Cell_Code2("select distinct code_name from tbl_codemaster where site_id ='" & Site_id & "' and class_name <> 'RM' and class_id in ('10002','R0007')"))
                Case 4
                    Me.FpSpread1.ActiveSheet.Cells(e.Row, e.Column).Locked = False
                    'RETURN STATE 콤보처리
                    Chg_ComboCell(FpSpread1, Me.FpSpread1.ActiveSheet.ActiveRowIndex, 4, Query_Cell_Code1("code_id", "R0007"))
                Case 5
                    Me.FpSpread1.ActiveSheet.Cells(e.Row, e.Column).Locked = False
            End Select
        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message, "ERROR")
        End Try
    End Sub

    Private Sub FpSpread1_Change(ByVal sender As Object, ByVal e As FarPoint.Win.Spread.ChangeEventArgs) Handles FpSpread1.Change
        Try
            Spread_Change(FpSpread1, e.Row)
            If e.Row < RecCnt Then
                Me.FpSpread1.ActiveSheet.Cells(e.Row, 8).Text = Emp_No
                Me.FpSpread1.ActiveSheet.Cells(e.Row, 9).Text = Now
            End If
        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message, "ERROR")
        End Try
    End Sub


    Private Sub SaveBtn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SaveBtn1.Click, SaveBtn.Click, bSave.Click
        Try
       
            If FpSpread1.ActiveSheet.RowCount < 1 Then
                Exit Sub
            End If

            With FpSpread1.ActiveSheet
                For I As Integer = 0 To .RowCount - 1
                    If .Rows(I).ForeColor = Color.OrangeRed Then
                        Dim QRY As String = ""

                        QRY = "IF EXISTS (SELECT MODEL FROM TBL_MODELLEVEL WHERE MODEL = '" & .GetValue(I, 1) & "' AND SERVICE_TYPE = '" & .GetValue(I, 2) & "')" & vbNewLine
                        QRY = QRY & "BEGIN" & vbNewLine
                        QRY = QRY & "   UPDATE TBL_MODELLEVEL SET SERVICE_NAME = '" & .GetValue(I, 3) & "', RETURN_STATE = '" & .GetValue(I, 4) & "', PRICE = " & .GetValue(I, 5) & " WHERE MODEL = '" & .GetValue(I, 1) & "' AND SERVICE_TYPE = '" & .GetValue(I, 2) & "'" & vbNewLine
                        QRY = QRY & "END" & vbNewLine
                        QRY = QRY & "ELSE" & vbNewLine
                        QRY = QRY & "BEGIN" & vbNewLine
                        QRY = QRY & "   INSERT INTO TBL_MODELLEVEL VALUES ('" & Site_id & "','" & .GetValue(I, 1) & "','" & .GetValue(I, 2) & "','" & .GetValue(I, 3) & "','" & .GetValue(I, 4) & "'," & .GetValue(I, 5) & ",'" & Emp_No & "', GETDATE(),'" & Emp_No & "', GETDATE())" & vbNewLine
                        QRY = QRY & "END" & vbNewLine

                        If Insert_Data(QRY) = True Then
                            .Rows(I).ForeColor = Color.Black
                        End If
                    End If
                Next

            End With

            Spread_AutoCol(Me.FpSpread1)

            MessageBox.Show("저장되었습니다", "Message")

            FindBtn_Click(Me.FpSpread1, System.EventArgs.Empty)

        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message, "ERROR")
        End Try

    End Sub


    Private Sub NewBtn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NewBtn1.Click, NewBtn.Click, bNew.Click
        Try

            If ClassCb.Text = "ALL" Then
                System.Console.Beep(3000, 400)
                System.Console.Beep(3000, 400)
                Modal_Error("SELECT MODEL!!")
                Exit Sub
            End If


            With Me.FpSpread1.ActiveSheet
                .RowCount += 1

                Spread_Celltype(Me.FpSpread1, Me.Name, .RowCount - 1, 1)

                .Cells(.RowCount - 1, 0).Text = Site_id
                If ClassCb.Text <> "ALL" Then
                    .Cells(.RowCount - 1, 1).Text = ClassCb.Text
                End If

                'If ModelLvCb.Text <> "ALL" Then
                '    .Cells(.RowCount - 1, 2).Text = ModelLvCb.Text
                'End If
                Chg_ComboCell(FpSpread1, .RowCount - 1, 2, Query_Cell_Code1("code_id", "10002"))
                Chg_ComboCell(FpSpread1, .RowCount - 1, 3, Query_Cell_Code2("select distinct code_name from tbl_codemaster where site_id ='" & Site_id & "' and class_name <> 'RM' and class_id in ('10002','R0007')"))
                Chg_ComboCell(FpSpread1, .RowCount - 1, 4, Query_Cell_Code1("code_id", "R0007"))
                .Cells(.RowCount - 1, 6).Text = Emp_No
                .Cells(.RowCount - 1, 7).Text = Now
                .Cells(.RowCount - 1, 8).Text = Emp_No
                .Cells(.RowCount - 1, 9).Text = Now
                Dim i As Integer
                For i = 1 To .ColumnCount - 5
                    .Cells(.RowCount - 1, i).Locked = False
                Next
            End With

            FpSpread1.ActiveSheet.SetActiveCell(FpSpread1.ActiveSheet.RowCount - 1, 0)
            FpSpread1.ShowActiveCell(FarPoint.Win.Spread.VerticalPosition.Bottom, FarPoint.Win.Spread.HorizontalPosition.Left) 'ActiveCell로 자동 스크롤 이동

        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message, "ERROR")
        End Try
    End Sub


    Private Sub ModelCpDisp()
        Try
            If Query_Combo(Me.ModelLvCb, "SELECT distinct SERVICE_TYPE FROM tbl_modellevel WHERE site_id = '" & Site_id & "' ORDER BY SERVICE_TYPE") = True Then
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
                With FpSpread1.ActiveSheet
                    If Insert_Data("DELETE FROM TBL_MODELLEVEL WHERE MODEL = '" & .GetValue(.ActiveRowIndex, 1) & "' AND SERVICE_TYPE = '" & .GetValue(.ActiveRowIndex, 2) & "'") = True Then
                        .RemoveRows(.ActiveRowIndex, 1)
                    End If

                End With

                MessageBox.Show("삭제되었습니다", "Message")

            End If
        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message, "ERROR")
        End Try
    End Sub

    Private Sub PrtBtn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PrtBtn1.Click, PrtBtn.Click, bPrint.Click

        If Spread_Print(Me.FpSpread1, "Model Level", 1) = False Then
            MsgBox("Fail to Print")
        End If

    End Sub


    Private Sub Excel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles XlsBtn1.Click, Excel.Click, bExcel.Click
        File_Save(SaveFileDialog1, FpSpread1)
    End Sub


End Class