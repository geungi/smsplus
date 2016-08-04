Public Class FrmBOM

    Private RecCnt As Integer
    Const CtrlMask As Byte = 8

    Private Sub FrmBOM_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Me.DockContainerItem1.Text = "실행 메뉴"
        Me.DockContainerItem10.Text = "조회 조건"

        ModelCpDisp()

        If Spread_Setting(FpSpread1, "FrmBOM") = True Then
            Spread_AutoCol(FpSpread1)
        End If

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
            If Me.ModelCb.SelectedItem = "" Then
                ModelCb.ForeColor = Color.Red
                ModelCb.Text = "Select Model"
            Else

                Me.LabelX2.Text = Me.ModelCb.Text

                Dim Act_YN As String = ""


                Select Case Me.ActCb.Text
                    Case "YES"
                        Act_YN = "Y"
                    Case "NO"
                        Act_YN = "N"
                    Case "ALL"
                        Act_YN = ""
                End Select

                Query_Spread(Me.FpSpread1, "EXEC SP_FRMBOM_LIST '" & Site_id & "','" & Me.ModelCb.Text & "'", 1)
                Spread_AutoCol(Me.FpSpread1)
            End If

        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message, "ERROR")
        End Try
    End Sub

    Private Sub PartList_ItemDrag(ByVal sender As System.Object, ByVal e As System.Windows.Forms.ItemDragEventArgs) Handles PartList.ItemDrag
        If e.Button = Windows.Forms.MouseButtons.Left Then
            'invoke the drag and drop operation
            DoDragDrop(e.Item, DragDropEffects.Move Or DragDropEffects.Copy)
        End If
    End Sub

    Private Sub PartList_DragEnter(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles PartList.DragEnter, FpSpread1.DragEnter
        If (e.Data.GetDataPresent("System.Windows.Forms.ListViewItem")) Then
            ' If the Ctrl key was pressed during the drag operation then perform
            ' a Copy. If not, perform a Move.
            If (e.KeyState And CtrlMask) = CtrlMask Then
                e.Effect = DragDropEffects.Copy
            Else
                e.Effect = DragDropEffects.Move
            End If
        Else
            e.Effect = DragDropEffects.None
        End If

    End Sub

    Private Sub PartList_DragDrop(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles PartList.DragDrop, FpSpread1.DragDrop

        If e.Data.GetDataPresent("System.Windows.Forms.ListViewItem", False) Then
            FpSpread_Fill(Me.PartList)
        End If
    End Sub

    Private Sub PartList_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles PartList.DoubleClick

        If SPREAD_DUP_CHECK(FpSpread1, PartList.SelectedItems(0).Text, 2) = False Then
            MessageBox.Show("Already exists.")
            Exit Sub
        End If
        FpSpread_Fill(Me.PartList)
    End Sub

    Private Sub FpSpread_Fill(ByVal ListVw As ListView)
        'BOM의 신규파트 추가
        Try
            Dim i, j As Integer

            For i = 0 To ListVw.SelectedItems.Count - 1

                With Me.FpSpread1.ActiveSheet
                    .RowCount += 1
                    .Cells(.RowCount - 1, 0).Text = Site_id
                    .Cells(.RowCount - 1, 1).Text = LabelX2.Text
                    .Cells(.RowCount - 1, 2).Text = ListVw.SelectedItems(i).SubItems(0).Text
                    .Cells(.RowCount - 1, 8).Text = 0
                    Chg_ComboCell(FpSpread1, .RowCount - 1, 5, New String() {"Y", "N"})
                    Chg_ComboCell(FpSpread1, .RowCount - 1, 6, New String() {"Y", "N"})
                    Chg_ComboCell(FpSpread1, .RowCount - 1, 7, New String() {"Y", "N"})
                    Chg_ComboCell(FpSpread1, .RowCount - 1, 9, New String() {"Y", "N"})
                    Select Case ListVw.SelectedItems(i).SubItems(4).Text
                        Case "LEVEL 1&2" '코스메틱
                            .Cells(.RowCount - 1, 5).Text = "Y"
                            .Cells(.RowCount - 1, 7).Text = "N"
                            .Cells(.RowCount - 1, 12).Text = "N"
                        Case "Technician" '텍
                            .Cells(.RowCount - 1, 5).Text = "N"
                            .Cells(.RowCount - 1, 7).Text = "N"
                            .Cells(.RowCount - 1, 12).Text = "N"
                        Case "FOC" 'FOC
                            .Cells(.RowCount - 1, 5).Text = "Y"
                            .Cells(.RowCount - 1, 7).Text = "Y"
                            .Cells(.RowCount - 1, 12).Text = "N"
                        Case "ETC" '시설재
                            .Cells(.RowCount - 1, 5).Text = "N"
                            .Cells(.RowCount - 1, 7).Text = "N"
                            .Cells(.RowCount - 1, 12).Text = "Y"
                    End Select
                    .Cells(.RowCount - 1, 6).Text = "N"
                    .Cells(.RowCount - 1, 9).Text = "N"
                    .Cells(.RowCount - 1, 10).Text = Query_RS("select isnull(repair_cd,'') from tbl_partmaster where site_id = '" & Site_id & "' and part_no = '" & ListVw.SelectedItems(i).SubItems(0).Text & "'")
                    .Cells(.RowCount - 1, 11).Text = Query_RS("select top def_cd from tbl_repset where site_id = '" & Site_id & "' and part_no = '" & ListVw.SelectedItems(i).SubItems(0).Text & "' and isnull(rep_cd,'') = '" & .Cells(.RowCount - 1, 10).Text & "'")
                    .Cells(.RowCount - 1, 13).Text = "Y"
                    .Cells(.RowCount - 1, 14).Text = Emp_No
                    .Cells(.RowCount - 1, 15).Text = Now
                    .Cells(.RowCount - 1, 16).Text = Emp_No
                    .Cells(.RowCount - 1, 17).Text = Now
                    .Cells(.RowCount - 1, 18).Text = ListVw.SelectedItems(i).SubItems(1).Text
                    .Cells(.RowCount - 1, 19).Text = ListVw.SelectedItems(i).SubItems(2).Text
                    .Cells(.RowCount - 1, 20).Text = ListVw.SelectedItems(i).SubItems(3).Text
                    .Rows(.RowCount - 1).ForeColor = Color.OrangeRed
                    For j = 1 To .ColumnCount - 8
                        .Cells(.RowCount - 1, j).Locked = False
                    Next
                End With

            Next

            Spread_AutoCol(Me.FpSpread1)

            For i = ListVw.SelectedItems.Count - 1 To 0 Step -1 'BOM에 추가된 파트는 리스트에서 삭제
                ListVw.Items.Remove(ListVw.SelectedItems.Item(i))
            Next

            FpSpread1.ActiveSheet.SetActiveCell(FpSpread1.ActiveSheet.RowCount - 1, 0)
            FpSpread1.ShowActiveCell(FarPoint.Win.Spread.VerticalPosition.Bottom, FarPoint.Win.Spread.HorizontalPosition.Left) 'Activecell위치로 스크롤 자동이동

        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message, "ERROR")
        End Try
    End Sub

    Private Sub FpSpread1_CellDoubleClick(ByVal sender As Object, ByVal e As FarPoint.Win.Spread.CellClickEventArgs) Handles FpSpread1.CellDoubleClick
        Try

            Me.FpSpread1.ActiveSheet.Cells(e.Row, e.Column).Locked = False


            Select Case e.Column
                Case 5
                    Chg_ComboCell(FpSpread1, e.Row, 5, New String() {"Y", "N"})
                Case 6
                    Chg_ComboCell(FpSpread1, e.Row, 6, New String() {"Y", "N"})
                Case 7
                    Chg_ComboCell(FpSpread1, e.Row, 7, New String() {"Y", "N"})
                Case 9
                    Chg_ComboCell(FpSpread1, e.Row, 9, New String() {"Y", "N"})
                Case 10
                    Spread_TxtType(FpSpread1, e.Row, 10, True, 0)
                Case 11
                    Spread_TxtType(FpSpread1, e.Row, 11, True, 0)
                Case 12
                    Chg_ComboCell(FpSpread1, e.Row, 12, New String() {"Y", "N"})
                Case 13
                    Chg_ComboCell(FpSpread1, e.Row, 13, New String() {"Y", "N"})
            End Select
        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message, "ERROR")
        End Try
    End Sub


    Private Sub FpSpread1_Change(ByVal sender As Object, ByVal e As FarPoint.Win.Spread.ChangeEventArgs) Handles FpSpread1.Change
        Try
            Spread_Change(FpSpread1, e.Row)

            With FpSpread1.ActiveSheet

                If e.Row < RecCnt Then
                    .Cells(e.Row, 16).Text = Emp_No
                    .Cells(e.Row, 17).Text = Now
                End If

            End With

        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message, "ERROR")
        End Try

    End Sub


    Private Sub FpSpread1_LeaveCell(ByVal sender As Object, ByVal e As FarPoint.Win.Spread.LeaveCellEventArgs) Handles FpSpread1.LeaveCell
        Try
         
            With FpSpread1.ActiveSheet

                If e.Column >= 9 And e.Column <= 11 Then

                    'If .GetText(e.Row, 9) = "Y" Then  'Basic 파트의 경우, 리페어셋 등록은 필수
                    '    If .GetText(e.Row, 10) = "" Then
                    '        MessageBox.Show("REPAIR CODE is Empty !!", "Alert")
                    '        Spread_TxtType(FpSpread1, e.Row, 10, True, 0)
                    '        .Cells(e.Row, 10).Locked = False
                    '        e.NewRow = e.Row
                    '        e.NewColumn = 10

                    '        Exit Sub
                    '    End If
                    '    If .GetText(e.Row, 11) = "" Then
                    '        MessageBox.Show("DEFECT CODE is Empty !!", "Alert")
                    '        Spread_TxtType(FpSpread1, e.Row, 11, True, 0)
                    '        .Cells(e.Row, 11).Locked = False
                    '        e.NewRow = e.Row
                    '        e.NewColumn = 11

                    '        Exit Sub
                    '    End If

                    'End If
                End If

            End With

        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message, "ERROR")
        End Try
    End Sub

    Private Sub SaveBtn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SaveBtn1.Click, SaveBtn.Click, bSave.Click
        Dim i As Integer
        Try
            With Me.FpSpread1.ActiveSheet

                If FpSpread1.ActiveSheet.RowCount < 1 Then
                    Exit Sub
                End If

                FpSpread1.ActiveSheet.SetActiveCell(0, 0)
               

                If Save_Vallidate(Me.FpSpread1, "BOM", New String() {1, 2, 3}) = False Then '셀들의 Validation 체크
                    Exit Sub
                End If

                For i = 0 To .RowCount - 1
                    If .Rows(i).ForeColor = Color.OrangeRed Then
                        If Query_RS("SELECT COUNT(P_NO) FROM TBL_BOM WHERE P_NO = '" & ModelCb.Text & "' And C_NO  = '" & FpSpread1.ActiveSheet.GetValue(i, 2) & "'") = 0 Then
                            Insert_Data( _
                                "insert tbl_bom (site_id,p_no,c_no,qty,loc_cd,cosmetic_yn,autopo_yn,foc_yn,foc_qty,basic_yn,etc_yn,active,c_person,c_date,u_person,u_date) " & _
                                " values ( " & _
                                "'" & .Cells(i, 0).Text & "'," & _
                                "'" & .Cells(i, 1).Text & "'," & _
                                "'" & .Cells(i, 2).Text & "'," & _
                                CInt(.Cells(i, 3).Text) & "," & _
                                "''," & _
                                "'" & .Cells(i, 5).Text & "'," & _
                                "'" & .Cells(i, 6).Text & "'," & _
                                "'" & .Cells(i, 7).Text & "'," & _
                                CInt(.Cells(i, 8).Text) & "," & _
                                "'" & .Cells(i, 9).Text & "'," & _
                                "'" & .Cells(i, 12).Text & "'," & _
                                "'" & .Cells(i, 13).Text & "'," & _
                                "'" & Emp_No & "', getdate(), " & _
                                "'" & Emp_No & "', getdate() )")
                        Else
                            Insert_Data( _
                            "UPDATE tbl_bom SET " & _
                            "site_id = '" & .Cells(i, 0).Text & "', " & _
                            "p_no = '" & .Cells(i, 1).Text & "', " & _
                            "c_no = '" & .Cells(i, 2).Text & "', " & _
                            "qty = " & .Cells(i, 3).Value & ", " & _
                            "loc_cd = '', " & _
                            "cosmetic_yn = '" & .Cells(i, 5).Text & "', " & _
                            "autopo_yn = '" & .Cells(i, 6).Text & "', " & _
                            "foc_yn = '" & .Cells(i, 7).Text & "', " & _
                            "foc_qty = " & .Cells(i, 8).Value & ", " & _
                            "basic_yn = '" & .Cells(i, 9).Text & "', " & _
                            "etc_yn = '" & .Cells(i, 12).Text & "', " & _
                            "active = '" & .Cells(i, 13).Text & "', " & _
                            "c_person = '" & Emp_No & "', " & _
                            "c_date = getdate(), " & _
                            "u_person = '" & Emp_No & "', " & _
                            "u_date = getdate() " & _
                            "WHERE  (site_id = '" & .Cells(i, 0).Text & "') AND (p_no = '" & .Cells(i, 1).Text & "') AND (c_no = '" & .Cells(i, 2).Text & "') " _
                            )
                        End If
                        'Insert_Data("EXEC SP_FRMBOM_REPSET '" & .Cells(i, 0).Text & "', '" & .Cells(i, 2).Text & "', '" & .Cells(i, 10).Text & "', '" & .Cells(i, 11).Text & "', '" & .Cells(i, 21).Text & "', '" & .Cells(i, 22).Text & "', '" & Emp_No & "' ")
                        .SetText(i, 21, .GetText(i, 10))
                        .SetText(i, 22, .GetText(i, 11))
                        .Rows(i).ForeColor = Color.Black
                    End If
                Next

                RecCnt = .RowCount

                For i = 0 To .RowCount - 1
                    .Rows(i).ForeColor = Color.Black
                Next


                'If Insert_Data("EXEC SP_FRMBOM_SETFAMILY '" & Site_id & "','" & .Cells(0, 1).Text & "'") = True Then
                'Else
                '    Modal_Error("CHECK FAMILY MODEL BOM!!")
                '    Exit Sub
                'End If

            End With

            'Insert_Data("exec SP_FRMPARTMASTER_DELDUMMYREPSET '" & Site_id & "'")


            MessageBox.Show("저장되었습니다", "Message")

        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message, "ERROR")
        End Try

    End Sub


    Private Sub ModelCpDisp()
        Try
            If Query_Combo(Me.ModelCb, "SELECT model_no FROM tbl_modelmaster WHERE site_id = '" & Site_id & "' and active='Y' ORDER BY model_no") = True Then
                If Site_id = "E1000" Then
                    Me.ModelCb.Items.Add("EQUIPMENT")
                End If
            End If
        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message, "ERROR")
        End Try
    End Sub

    Private Sub DelBtn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DelBtn1.Click, DelBtn.Click, bDel.Click
        Try
            Dim RowCnt = Me.FpSpread1.ActiveSheet.RowCount
            Dim i As Integer = Me.FpSpread1.ActiveSheet.ActiveRowIndex
            Dim part_dv As String
            Dim r As DialogResult = MessageBox.Show("Selected rows delete now?", "Selected Rows Delete", MessageBoxButtons.YesNo)
            If r = Windows.Forms.DialogResult.Yes Then
                part_dv = Query_RS("select part_dv from tbl_partmaster where part_no = '" & Me.FpSpread1.ActiveSheet.GetValue(FpSpread1.ActiveSheet.ActiveRowIndex, 2) & "'")
                'BOM에서 삭제되는 파트를 다시 파트리스트에 추가하여 선택할수 있도록 함
                Me.PartList.Items.Add(New ListViewItem(New String() {Me.FpSpread1.ActiveSheet.Cells(i, 2).Text, Me.FpSpread1.ActiveSheet.Cells(i, 18).Text, Me.FpSpread1.ActiveSheet.Cells(i, 19).Text, Me.FpSpread1.ActiveSheet.Cells(i, 20).Text, part_dv}))
                If Me.FpSpread1.ActiveSheet.ActiveRowIndex < RecCnt Then '기존데이터인 경우, tbl_bom에서 삭제
                    Insert_Data("delete tbl_bom where site_id='" & Site_id & "' and p_no ='" & FpSpread1.ActiveSheet.GetValue(FpSpread1.ActiveSheet.ActiveRowIndex, 1) & "' and c_no ='" & FpSpread1.ActiveSheet.GetValue(FpSpread1.ActiveSheet.ActiveRowIndex, 2) & "'")
                    Me.FpSpread1.ActiveSheet.RemoveRows(Me.FpSpread1.ActiveSheet.ActiveRowIndex, 1)
                    RecCnt = RecCnt - 1  '기존데이터가 삭제되면, 추가되는 파트의 인덱스 시작번호가 달라지므로 -1
                Else
                    Me.FpSpread1.ActiveSheet.RemoveRows(Me.FpSpread1.ActiveSheet.ActiveRowIndex, 1)
                End If

                MessageBox.Show("삭제되었습니다", "Message")
            End If
        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message, "ERROR")
        End Try
    End Sub

    Private Sub PrtBtn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PrtBtn1.Click, PrtBtn.Click, bPrint.Click
        If Spread_Print(Me.FpSpread1, "BOM", 1) = False Then
            MsgBox("Fail to Print")
        End If

    End Sub


    Private Sub Excel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles XlsBtn1.Click, Excel.Click, bExcel.Click
        File_Save(SaveFileDialog1, FpSpread1)
    End Sub


    'Private Sub UcXlsUpload1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UcXlsUpload1.Load
    '    UcXup_SP = Me.FpSpread1
    '    UcXup_FrmNm = Me.Name
    '    UcXup_ListVw = Me.PartList
    'End Sub

    Private Sub CheckBoxX1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBoxX1.CheckedChanged
        If Me.LabelX2.Text = "MODEL NAME" Or Me.LabelX2.Text = "" Then
            MessageBox.Show("Selected Model is not!!! ", "Validation error")
            Exit Sub
        End If
        If Me.CheckBoxX1.Checked = True Then
            Me.Cursor = Cursors.WaitCursor

            With FpSpread1.ActiveSheet
                .RowCount = 0
                If File_Open3(OpenFileDialog1, FpSpread1, TextBoxX1, Me.Name) = True Then
                    .Rows(0, .RowCount - 1).ForeColor = Color.OrangeRed
                End If
                '                .Cells(0, 0, .RowCount - 1, .ColumnCount - 1).ForeColor = Color.OrangeRed
            End With
            MessageBox.Show("COMPLETE TO OPEN")
            Me.Cursor = Cursors.Default
            '            Me.UcXlsUpload1.Visible = True
        Else
            '  Me.UcXlsUpload1.Visible = False
        End If
    End Sub


    Private Sub TextBoxX1_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TextBoxX1.KeyUp
        Try
            If e.KeyCode <> Keys.Enter Then
                Exit Sub
            End If

            Dim QRY As String

            QRY = "SELECT PART_NO, PART_NAME, oringinal_no, PRICE, ISNULL((SELECT CODE_NAME FROM TBL_CODEMASTER WHERE SITE_ID = A.SITE_ID AND CLASS_ID = 'R0038' AND CODE_ID = A.PART_DV),'')" & vbNewLine
            QRY = QRY & "FROM TBL_PARTMASTER A" & vbNewLine
            QRY = QRY & "WHERE SITE_ID = '" & Site_id & "'" & vbNewLine
            QRY = QRY & "  AND PART_NO LIKE '%" & TextBoxX1.Text & "%'" & vbNewLine
            QRY = QRY & "  and active = 'Y'" & vbNewLine
            QRY = QRY & "ORDER BY PART_NO" & vbNewLine

            If Query_Listview(PartList, QRY, True) = True Then
            End If

            If TextBoxX1.Text <> "" Then
                Dim item1 As ListViewItem = PartList.FindItemWithText(TextBoxX1.Text) '파트검색
                If (item1 IsNot Nothing) Then
                    'MsgBox(item1.ToString)
                    item1.EnsureVisible()  '해당위치로 스크롤 이동
                    item1.Selected = True
                    PartList.Focus()
                End If
            End If

        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message, "ERROR")
        End Try
    End Sub

    Private Sub ModelCb_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ModelCb.SelectedIndexChanged
        FpSpread1.ActiveSheet.RowCount = 0
    End Sub
End Class