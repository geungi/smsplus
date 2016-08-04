Public Class FrmPartMst

    Private Sub FrmPartMst_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Me.DockContainerItem1.Text = "실행 메뉴"
        Me.DockContainerItem3.Text = "조회 조건"
        Me.PartTxt.Text = ""

        If Spread_Setting(FpSpread1, "FrmPartMst") = True Then
            FpSpread1.ActiveSheet.FrozenColumnCount = 3
            Spread_AutoCol(FpSpread1)
        End If

        'Me.RecCb.Visible = True

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
            Dim Sub_YN As String = ""
            Dim Rec_YN As String = ""
            Dim RecSp_YN As String = ""


            Dim qry As String

            qry = "SELECT site_id,part_no,part_name,part_spec," & vbNewLine
            qry = qry & "     ISNULL((SELECT CODE_NAME FROM TBL_CODEMASTER WHERE SITE_ID = A.SITE_ID AND CLASS_ID = 'R0038' AND CODE_ID = A.PART_DV),'')," & vbNewLine
            qry = qry & "     ISNULL((SELECT CODE_NAME FROM TBL_CODEMASTER WHERE SITE_ID = A.SITE_ID AND CLASS_ID = 'R4002' AND CODE_ID = A.ASSY_DV),'')," & vbNewLine
            qry = qry & "     color,part_level,price,substitute_yn,oringinal_no,recycle_yn,recycle_spec,repair_cd,lg_partno,active," & vbNewLine
            qry = qry & "     ISNULL((SELECT EMP_NM FROM TBL_EMPMASTER WHERE SITE_ID = A.SITE_ID AND EMP_NO = A.C_PERSON),'ADMIN'), C_DATE," & vbNewLine
            qry = qry & "     ISNULL((SELECT EMP_NM FROM TBL_EMPMASTER WHERE SITE_ID = A.SITE_ID AND EMP_NO = A.U_PERSON),'ADMIN'), U_DATE, USGAVG" & vbNewLine
            qry = qry & "FROM TBL_PARTMASTER A" & vbNewLine
            qry = qry & "WHERE SITE_ID = '" & Site_id & "'" & vbNewLine

            If PartTxt.Text <> "" Then
                qry = qry & "  AND PART_NO LIKE '%" & PartTxt.Text & "%'" & vbNewLine
            End If

            If Me.ActCb.Text <> "ALL" Then
                qry = qry & "  AND ACTIVE = '" & Mid(ActCb.Text, 1, 1) & "'" & vbNewLine
            End If

            If Me.SubCb.Text <> "ALL" Then
                qry = qry & "  AND substitute_yn = '" & Mid(SubCb.Text, 1, 1) & "'" & vbNewLine
            End If

            If Me.RecCb.Text <> "ALL" Then
                qry = qry & "  AND recycle_yn = '" & Mid(RecCb.Text, 1, 1) & "'" & vbNewLine
            End If

            If Me.RecSpCb.Text <> "ALL" Then
                qry = qry & "  AND recycle_spec = '" & RecSpCb.Text & "'" & vbNewLine
            End If

            qry = qry & "ORDER BY PART_NO" & vbNewLine

            If Query_Spread(FpSpread1, qry, 1) = True Then
                Spread_AutoCol(FpSpread1)
            End If

            Spread_AutoCol(FpSpread1) '스프레드에 데이터를 출력후, 데이터의 사이즈에 맞게 컬럼사이즈 재조정

        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message, "ERROR")
        End Try
    End Sub

    Private Sub FpSpread1_CellDoubleClick(ByVal sender As Object, ByVal e As FarPoint.Win.Spread.CellClickEventArgs) Handles FpSpread1.CellDoubleClick
        Try

            If (e.Column > 0 And e.Column < 15) Or e.Column = 20 Then
                Me.FpSpread1.ActiveSheet.Cells(e.Row, e.Column).Locked = False
            End If

            Select Case e.Column
                Case 1
                    '셀타입을 text셀로, tbl_spread테이블에 있는 해당 스프레드의 컬럼의 속성(maxlength,대소문자구분)을 셀속성으로 설정
                    Spread_Celltype(Me.FpSpread1, Me.Name, e.Row, e.Column)
                Case 2
                    Spread_Celltype(Me.FpSpread1, Me.Name, e.Row, e.Column)
                Case 3
                    Spread_Celltype(Me.FpSpread1, Me.Name, e.Row, e.Column)
                Case 4
                    Chg_ComboCell(FpSpread1, Me.FpSpread1.ActiveSheet.ActiveRowIndex, 4, Query_Cell_Code1("code_NAME", "R0038"))
                Case 5
                    Chg_ComboCell(Me.FpSpread1, e.Row, 5, Query_Cell_Code1("code_NAME", "R4002"))
                Case 6
                    Spread_Celltype(Me.FpSpread1, Me.Name, e.Row, e.Column)
                Case 9
                    Chg_ComboCell(FpSpread1, Me.FpSpread1.ActiveSheet.ActiveRowIndex, 9, New String() {"Y", "N"})
                Case 12
                    Chg_ComboCell(FpSpread1, Me.FpSpread1.ActiveSheet.ActiveRowIndex, 12, New String() {"NEW", "SITE"})
                Case 11
                    Chg_ComboCell(FpSpread1, Me.FpSpread1.ActiveSheet.ActiveRowIndex, 11, New String() {"Y", "N"})
                Case 13
                    Chg_ComboCell(FpSpread1, Me.FpSpread1.ActiveSheet.ActiveRowIndex, 13, Query_Cell_Code1("code_NAME", "R4001"))
                Case 15
                    Chg_ComboCell(FpSpread1, Me.FpSpread1.ActiveSheet.ActiveRowIndex, 15, New String() {"Y", "N"})
                    'Case 20
                    '    Spread_Celltype(Me.FpSpread1, Me.Name, e.Row, e.Column)
            End Select
        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message, "ERROR")
        End Try
    End Sub

    Private Sub FpSpread1_Change(ByVal sender As Object, ByVal e As FarPoint.Win.Spread.ChangeEventArgs) Handles FpSpread1.Change
        Try
            Spread_Change(FpSpread1, e.Row)
            Me.FpSpread1.ActiveSheet.Cells(e.Row, 18).Text = Emp_No
            Me.FpSpread1.ActiveSheet.Cells(e.Row, 19).Text = Now
            Dim S1 = Me.FpSpread1.ActiveSheet
            If e.Column = 12 Then
                Select Case S1.GetText(e.Row, 12)   'LG 리사이클 파트인 경우, LG PART NO 입력하도록 처리
                    Case "NEW"
                        Me.FpSpread1.ActiveSheet.SetText(e.Row, 14, Me.FpSpread1.ActiveSheet.GetValue(e.Row, 1))
                    Case "SITE"
                        Me.FpSpread1.ActiveSheet.SetText(e.Row, 14, Me.FpSpread1.ActiveSheet.GetValue(e.Row, 1))
                End Select
            End If
            If e.Column = 1 Then
                Me.FpSpread1.ActiveSheet.SetValue(e.Row, 14, Me.FpSpread1.ActiveSheet.GetValue(e.Row, 1))
            End If
        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message, "ERROR")
        End Try

    End Sub

    Private Sub SaveBtn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SaveBtn1.Click, SaveBtn.Click, bSave.Click
        Dim part As String = ""

        Try
            Dim QRY As String

            If FpSpread1.ActiveSheet.RowCount < 1 Then
                Exit Sub
            End If

            FpSpread1.ActiveSheet.SetActiveCell(0, 0)

            If FpSpread1.ActiveSheet.ActiveRowIndex = (Me.FpSpread1.ActiveSheet.RowCount - 1) Then
                Me.FpSpread1.ActiveSheet.SetActiveCell(FpSpread1.ActiveSheet.ActiveRowIndex - 1, FpSpread1.ActiveSheet.ActiveColumnIndex)
            Else
                Me.FpSpread1.ActiveSheet.SetActiveCell(FpSpread1.ActiveSheet.ActiveRowIndex + 1, FpSpread1.ActiveSheet.ActiveColumnIndex)
            End If

            Dim i As Integer

            With Me.FpSpread1.ActiveSheet

                For i = 0 To .RowCount - 1
                    If .Rows(i).ForeColor = Color.OrangeRed Then

                        QRY = "EXEC SP_FRMPARTMST_SAVE '" & Site_id & "'," & vbNewLine
                        QRY = QRY & "'" & .GetValue(i, 1) & "'," & vbNewLine
                        QRY = QRY & "'" & .GetValue(i, 2) & "'," & vbNewLine
                        QRY = QRY & "'" & .GetValue(i, 3) & "'," & vbNewLine
                        QRY = QRY & "'" & .GetValue(i, 4) & "'," & vbNewLine
                        QRY = QRY & "'" & .GetValue(i, 5) & "'," & vbNewLine
                        QRY = QRY & "'" & .GetValue(i, 6) & "'," & vbNewLine
                        QRY = QRY & "" & .GetValue(i, 7) & "," & vbNewLine
                        QRY = QRY & "" & .GetValue(i, 8) & "," & vbNewLine
                        QRY = QRY & "'" & .GetValue(i, 9) & "'," & vbNewLine
                        QRY = QRY & "'" & .GetValue(i, 10) & "'," & vbNewLine
                        QRY = QRY & "'" & .GetValue(i, 11) & "'," & vbNewLine
                        QRY = QRY & "'" & .GetValue(i, 12) & "'," & vbNewLine
                        QRY = QRY & "'" & .GetValue(i, 13) & "'," & vbNewLine
                        QRY = QRY & "'" & .GetValue(i, 14) & "'," & vbNewLine
                        QRY = QRY & "'" & .GetValue(i, 15) & "'," & vbNewLine
                        QRY = QRY & "" & .GetValue(i, 20) & "," & vbNewLine
                        QRY = QRY & "'" & Emp_No & "'" & vbNewLine

                        If Insert_Data(QRY) = True Then
                            .Rows(i).ForeColor = Color.Black
                        End If
                    End If
                Next
            End With

            Spread_AutoCol(Me.FpSpread1)

            MessageBox.Show("저장되었습니다", "Message")

        Catch ex As Exception
            MessageBox.Show("Error: (" & part & ")" & ex.Message, "ERROR")
        End Try

    End Sub


    Private Sub NewBtn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NewBtn1.Click, NewBtn.Click, bNew.Click
        Try

            With Me.FpSpread1.ActiveSheet
                .RowCount += 1

                Spread_Celltype(Me.FpSpread1, Me.Name, .RowCount - 1, 1)
                Spread_Celltype(Me.FpSpread1, Me.Name, .RowCount - 1, 2)
                Spread_Celltype(Me.FpSpread1, Me.Name, .RowCount - 1, 3)

                Spread_Celltype(Me.FpSpread1, Me.Name, .RowCount - 1, 5)
                Spread_Celltype(Me.FpSpread1, Me.Name, .RowCount - 1, 6)

                .Cells(.RowCount - 1, 0).Text = Site_id
                If PartTxt.Text <> "" Then
                    .Cells(.RowCount - 1, 1).Text = PartTxt.Text
                End If
                Chg_ComboCell(FpSpread1, .RowCount - 1, 4, Query_Cell_Code1("code_NAME", "R0038"))
                Chg_ComboCell(FpSpread1, .RowCount - 1, 9, New String() {"Y", "N"})
                Chg_ComboCell(FpSpread1, .RowCount - 1, 11, New String() {"Y", "N"})
                Chg_ComboCell(FpSpread1, .RowCount - 1, 12, New String() {"NEW", "SITE"})
                Chg_ComboCell(FpSpread1, .RowCount - 1, 13, Query_Cell_Code1("code_NAME", "R4001"))
                Chg_ComboCell(FpSpread1, .RowCount - 1, 15, New String() {"Y", "N"})
                .Cells(.RowCount - 1, 9).Text = "N"
                .Cells(.RowCount - 1, 12).Text = "NEW"
                .Cells(.RowCount - 1, 11).Text = "N"
                .Cells(.RowCount - 1, 15).Text = "Y"
                .Cells(.RowCount - 1, 16).Text = Emp_No
                .Cells(.RowCount - 1, 17).Text = Now
                .Cells(.RowCount - 1, 18).Text = Emp_No
                .Cells(.RowCount - 1, 19).Text = Now
                Dim i As Integer
                For i = 1 To .ColumnCount - 5
                    .Cells(.RowCount - 1, i).Locked = False
                Next
            End With

            Me.FpSpread1.ActiveSheet.SetActiveCell(Me.FpSpread1.ActiveSheet.RowCount - 1, 0)
            Me.FpSpread1.ShowActiveCell(FarPoint.Win.Spread.VerticalPosition.Bottom, FarPoint.Win.Spread.HorizontalPosition.Left) 'Activecell 위치로 자동 스크롤 이동

        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message, "ERROR")
        End Try
    End Sub


    Private Sub DelBtn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DelBtn1.Click, DelBtn.Click, bDel.Click
        Try
            Dim r As DialogResult = MessageBox.Show("Selected row delete now?", "Selected Rows Delete", MessageBoxButtons.YesNo)
            With FpSpread1.ActiveSheet
                If r = Windows.Forms.DialogResult.Yes Then
                    '기존 파트에 대한 삭제시, tbl_bpartinv(기초재고)에 재고수량이 존재하는지 체크하여, 있는경우 삭제 할 수 없음
                    '파트에 대한 삭제는 로직 재정의 필요함
                    If CInt(Query_RS("select isnull(sum(qty),0) from tbl_bpartinv where site_id = '" & Site_id & "' and part_no = '" & .GetValue(.ActiveRowIndex, 1) & "'")) > 0 Then
                        MessageBox.Show("Part's Qty is not Zero in Basic Part Inventory'", "Validation Error")
                        Exit Sub
                    Else
                        '삭제조건을 만족하면, tbl_bpartinv(기초재고)도 삭제
                        Insert_Data("delete tbl_bpartinv where site_id = '" & Site_id & "' and part_no = '" & .GetValue(.ActiveRowIndex, 1) & "'")
                    End If
                    .RemoveRows(.ActiveRowIndex, 1)
                End If

            End With

            MessageBox.Show("삭제되었습니다", "Message")

        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message, "ERROR")
        End Try
    End Sub

    Private Sub PrtBtn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PrtBtn1.Click, PrtBtn.Click, bPrint.Click
        If Spread_Print(Me.FpSpread1, "Part Master", 1) = False Then
            MsgBox("Fail to Print")
        End If

    End Sub

    Private Sub Excel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles XlsBtn1.Click, Excel.Click, bExcel.Click
        File_Save(SaveFileDialog1, FpSpread1)
    End Sub

    'Private Sub UcXlsUpload1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UcXlsUpload1.Load
    '    UcXup_SP = Me.FpSpread1
    '    UcXup_FrmNm = Me.Name
    'End Sub

    Private Sub CheckBoxX1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBoxX1.CheckedChanged
        'If Me.CheckBoxX1.Checked = True Then

        '    Me.UcXlsUpload1.Visible = True
        'Else
        '    Me.UcXlsUpload1.Visible = False
        'End If
    End Sub



End Class