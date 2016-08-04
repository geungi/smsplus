Public Class FrmPurchaseOrder
    Const CtrlMask As Byte = 8
    Private NowPoNo As String = ""
    Private SelPoNo As String = ""
    Private TotRec As Integer = 0
    Private TotRec2 As Integer = 0
    Private PoRec As Integer = 0


    Private Sub FrmPurchaseOrder_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Me.DockContainerItem1.Text = "실행 메뉴"
        Me.DockContainerItem6.Text = "조회 조건"
        Me.DockContainerItem2.Text = "품목 리스트"
        Me.DockContainerItem7.Text = "구매발주서"
        Bar5.Visible = False

        Condi_Disp() '콤보박스의 조건데이터 출력

        Me.Bar13.Visible = False
        '
        If Spread_Setting(FpSpread1, Me.Name) = True Then
            Spread_AutoCol(FpSpread1)
        End If
        If Spread_Setting(FpSpread2, Me.Name) = True Then
            Spread_AutoCol(FpSpread2)
        End If
        If Spread_Setting(FpSpread3, Me.Name) = True Then
            Spread_AutoCol(FpSpread3)
        End If
        Me.ContextMenuBar1.SetContextMenuEx(Me.FpSpread2, CtxSp)
        Me.ContextMenuBar1.SetContextMenuEx(Me.FpSpread3, CtxSp2)
        Me.ContextMenuBar1.SetContextMenuEx(Me.FpSpread1, CtxSp3)

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

        If Emp_No = "11111" Then
            ButtonX2.Visible = True
        Else
            ButtonX2.Visible = False
        End If


    End Sub

    Private Sub Condi_Disp() 'CONTROL PANEL 및 파트리스트뷰에 있는 콤보박스에 데이터 출력
        Me.POStDate.Value = Now
        Me.POEdDate.Value = Now

        'Part List의 model
        Query_Combo(Me.PtMdCb, "select model_no from tbl_modelmaster where site_id = '" & Site_id & "' and active='Y' ORDER BY model_no")
        Me.PtMdCb.Items.Add("ALL")
        'STATUS
        Me.StatCb.Text = "ALL"
        Query_Combo(Me.StatCb, "select code_name from tbl_codemaster where site_id = '" & Site_id & "' and class_id = 'R0004' order by dis_order")
        Me.StatCb.Items.Add("ALL")
        'SUPPLIER
        Me.SupCb.Text = "ALL"
        Query_Combo(Me.SupCb, "select SUP_NM from tbl_SUPMST where site_id = '" & Site_id & "' ORDER BY SUP_NM")
        Me.SupCb.Items.Add("ALL")
        'MODEL
        Me.ModelCb.Text = "ALL"
        Query_Combo(Me.ModelCb, "select model_no from tbl_modelmaster where site_id = '" & Site_id & "' and active='Y' ORDER BY model_no")
        Me.ModelCb.Items.Add("ALL")
    End Sub

    Private Sub PtMdCb_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PtMdCb.SelectedIndexChanged
        Try
            Dim i As Integer
            Query_Listview(Me.PartList, "exec SP_FRMPURCHASEORDER_PARTLIST '" & Site_id & "','" & Me.PtMdCb.Text & "'", True)

            For i = 0 To Me.PartList.Items.Count - 1
                If Len(Me.PartList.Items(i).Text) > 11 Then '리사이클 파트의 경우, 색상표시
                    Me.PartList.Items(i).ForeColor = Color.Red
                End If
            Next

            If Me.FpSpread1.ActiveSheet.RowCount = 1 Then
                Me.FpSpread1.ActiveSheet.SetText(0, 7, Me.PtMdCb.Text)
            End If

            If Me.FpSpread2.ActiveSheet.RowCount > 0 Then

                For i = 0 To Me.FpSpread2.ActiveSheet.RowCount - 1
                    'Po detail에 등록되어있는 파트는 리스트뷰에서 제거하여 중복으로 등록되는 것을 방지
                    Remove_Listview(Me.PartList, Me.FpSpread2.ActiveSheet.Cells(i, 1).Text)
                Next
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

    Private Sub FpSpread2_Change(ByVal sender As Object, ByVal e As FarPoint.Win.Spread.ChangeEventArgs) Handles FpSpread2.Change
        Try
            Spread_Change(sender, e.Row)
        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message, "ERROR")
        End Try
    End Sub

    Private Sub FpSpread1_CellClick(ByVal sender As Object, ByVal e As FarPoint.Win.Spread.CellClickEventArgs) Handles FpSpread1.CellClick
        Try
            If FpSpread1.ActiveSheet.RowCount > 0 Then
                Dim AA As String = FpSpread1.ActiveSheet.GetValue(e.Row, 0)

                If SelPoNo <> FpSpread1.ActiveSheet.GetValue(e.Row, 0) Then
                    Me.DockContainerItem5.Text = "P/O DETAIL[" & Me.FpSpread1.ActiveSheet.GetValue(e.Row, 0) & "]"
                    disp_list(e.Row)
                End If
            End If
        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message, "ERROR")
        End Try
    End Sub


    Private Sub FpSpread2_LeaveCell(ByVal sender As Object, ByVal e As FarPoint.Win.Spread.LeaveCellEventArgs) Handles FpSpread2.LeaveCell
        Try
       
            Select Case e.Column
                Case 4
                    CalcPart()
                Case 5
                    If Me.FpSpread2.ActiveSheet.Cells(e.Row, e.Column).Text = "" Then

                        Exit Sub
                    End If
                    CalcPart()
            End Select
        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message, "ERROR")
        End Try
    End Sub


    Private Sub FpSpread1_Change(ByVal sender As Object, ByVal e As FarPoint.Win.Spread.ChangeEventArgs) Handles FpSpread1.Change
        Try
            Spread_Change(sender, e.Row)

            Spread_AutoCol(FpSpread1)
            'Select Case sender.ActiveSheet.ActiveColumnIndex
            '    Case 1
            '        ChgNewPoNo(Me.FpSpread1.ActiveSheet.ActiveRowIndex)
            'End Select

            'Select Case e.Column
            '    Case 2
            '        If FpSpread1.ActiveSheet.RowCount > 0 Then
            '            Dim r = MessageBox.Show("선택된 P/O의 공급처를 " & Me.FpSpread1.ActiveSheet.GetText(e.Row, e.Column) & "로 변경하시겠습니까?", "공급처 변경", MessageBoxButtons.OKCancel)
            '            If r = Windows.Forms.DialogResult.OK Then
            '                If CInt(Query_RS("select count(*) tbl_podetail from po_no = '" & Me.FpSpread1.ActiveSheet.GetValue(e.Row, 0) & "' and status = 'CLOSED'")) > 0 Then
            '                    MessageBox.Show("Already receiving!!!", "Error")
            '                    Exit Sub
            '                End If
            '                Insert_Data("update tbl_poheader set sup_cd = '" & Me.FpSpread1.ActiveSheet.GetText(e.Row, e.Column) & "' where site_id = '" & Site_id & "' and po_no = '" & Me.FpSpread1.ActiveSheet.GetValue(e.Row, 0) & "'")
            '                MessageBox.Show("Complete to modify", "Message")
            '            End If
            '        End If
            'End Select

        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message, "ERROR")
        End Try
    End Sub

    Private Sub FpSpread3_CellClick(ByVal sender As Object, ByVal e As FarPoint.Win.Spread.CellClickEventArgs) Handles FpSpread3.CellClick
        Dim S2 = Me.FpSpread2.ActiveSheet
        Dim S3 = Me.FpSpread3.ActiveSheet

        Dim J As String() = SPREAD_SEARCH(Me.FpSpread2, 0, S3.Cells(e.Row, 1).Text, 0, 1, S2.RowCount - 1, 1, False)
        If J(0) > -1 And J(1) > -1 Then
            S2.SetActiveCell(J(0), 5)
            Me.FpSpread2.ShowActiveCell(J(0), 5)
        End If
    End Sub

    Private Sub FpSpread3_Change(ByVal sender As Object, ByVal e As FarPoint.Win.Spread.ChangeEventArgs) Handles FpSpread3.Change
        Try
            Dim S2 = Me.FpSpread2.ActiveSheet
            Dim S3 = Me.FpSpread3.ActiveSheet
            Dim i As Integer
            Dim totamt As Decimal

            Spread_Change(sender, e.Row)

            Select Case e.Column
                Case 3
                    If S3.Cells(e.Row, e.Column).Text <> "" Then
                        S3.Cells(e.Row, 6).Text = CDec(S3.Cells(e.Row, 2).Text) * CInt(S3.Cells(e.Row, 3).Text)
                        For i = 0 To S2.RowCount - 1
                            If S2.Cells(i, 1).Text = S3.Cells(e.Row, 1).Text Then
                                If (CInt(S2.Cells(i, 6).Text) + CInt(S3.Cells(e.Row, 3).Text)) > CInt(S2.Cells(i, 5).Text) Then
                                    MessageBox.Show("[P/O DETAILS(" & i + 1 & ")]R_QTY > O_QTY !!", "validation Error")  'R_QTY의 합계가 O_Qty보다 클수 없다.
                                    S3.SetActiveCell(e.Row, e.Column)
                                    Me.FpSpread3.UndoManager.Undo(1)

                                    If S3.Cells(e.Row, 3).Text <> "" Then
                                        S3.Cells(e.Row, 6).Text = CDec(S3.Cells(e.Row, 2).Text) * CInt((S3.Cells(e.Row, 3).Text))
                                    Else
                                        S3.Cells(e.Row, 6).Text = ""
                                    End If

                                    Exit Sub

                                End If
                            End If
                        Next
                        For i = 0 To S3.RowCount - 2
                            If S3.Cells(i, 6).Text <> "" Then
                                totamt += CDec(S3.Cells(i, 6).Text)
                            End If
                        Next
                        S3.Cells(S3.RowCount - 1, 6).Text = totamt
                    Else
                        'MessageBox.Show("[" & S3.ColumnHeader.Columns(3).Label & "] isn't Empty!!!", "validation Error")
                        'S3.SetActiveCell(e.Row, e.Column)
                    End If
                Case 2
                    If S3.Cells(e.Row, 3).Text <> "" Then
                        S3.Cells(e.Row, 6).Text = CDec(S3.Cells(e.Row, 2).Text) * CInt(S3.Cells(e.Row, 3).Text)
                        For i = 0 To S3.RowCount - 2
                            If S3.Cells(i, 6).Text <> "" Then
                                totamt += CDec(S3.Cells(i, 6).Text)
                            End If
                        Next
                        S3.Cells(S3.RowCount - 1, 6).Text = totamt
                    Else
                        'MessageBox.Show("[" & S3.ColumnHeader.Columns(3).Label & "] isn't Empty!!!", "validation Error")
                        'S3.SetActiveCell(e.Row, e.Column)
                    End If
            End Select
        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message, "ERROR")
        End Try
    End Sub

    Private Sub FpSpread3_LeaveCell(ByVal sender As Object, ByVal e As FarPoint.Win.Spread.LeaveCellEventArgs) Handles FpSpread3.LeaveCell
        Try
            Dim S3 = Me.FpSpread3.ActiveSheet
            If S3.ActiveRowIndex < S3.RowCount - 1 Then
                Select Case e.Column
                    Case 3
                        If S3.Cells(e.Row, e.Column).Text = "" Then
                            S3.Cells(e.Row, e.Column).Locked = False
                        End If
                    Case 4
                        If S3.Cells(e.Row, e.Column).Text = "" Then
                            S3.Cells(e.Row, e.Column).Locked = False

                        End If
                End Select
            End If
        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message, "ERROR")
        End Try
    End Sub

    Private Sub PartList_DragEnter(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles PartList.DragEnter, FpSpread2.DragEnter
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

    Private Sub PartList_DragDrop(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles PartList.DragDrop, FpSpread2.DragDrop

        If e.Data.GetDataPresent("System.Windows.Forms.ListViewItem", False) Then
            FpSpread_Fill()
        End If
    End Sub

    Private Sub PartList_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles PartList.DoubleClick
        FpSpread_Fill()
    End Sub

    Private Sub NewBtn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NewBtn1.Click, NewBtn.Click, bNew.Click
        Try
            Dim S = Me.FpSpread1.ActiveSheet
            Dim i As Integer

            S.Rows.Clear()
            Me.FpSpread2.ActiveSheet.Rows.Remove(0, Me.FpSpread2.ActiveSheet.RowCount)
            Me.FpSpread3.ActiveSheet.Rows.Remove(0, Me.FpSpread3.ActiveSheet.RowCount)
            NowPoNo = ""
            If NowPoNo = "" Then


                If Me.PtMdCb.Text = "" Then
                    MessageBox.Show("Select PartList's Model !!!", "Validation Error")
                    Exit Sub
                End If

                Create_PO()

                If S.RowCount > 0 Then
                    If S.Cells(S.RowCount - 1, 0).Text = NowPoNo Then
                        MessageBox.Show("[" & NowPoNo & "]'is existed!", "validation Error")
                        Exit Sub
                    End If
                End If

                S.Rows.Clear()
                Me.DockContainerItem5.Text = "P/O DETAIL"

                S.RowCount += 1
                S.Cells(S.RowCount - 1, 0).Text = NowPoNo
                S.Cells(S.RowCount - 1, 1).Text = Now
                Chg_ComboCell(Me.FpSpread1, S.RowCount - 1, 2, Query_Cell_Code2("select SUP_NM from TBL_SUPMST where site_id = '" & Site_id & "'"))
                Chg_ComboCell(Me.FpSpread1, S.RowCount - 1, 3, Query_Cell_Code1("code_name", "R0004"))
                Chg_ComboCell(Me.FpSpread1, S.RowCount - 1, 7, Query_Cell_Code2("select model_no from tbl_modelmaster where site_id = '" & Site_id & "' and active='Y'"))
                Chg_ComboCell(Me.FpSpread1, S.RowCount - 1, 8, Query_Cell_Code1("code_name", "R4001"))
                S.Cells(S.RowCount - 1, 8).Text = "KRW"
                S.Cells(S.RowCount - 1, 2).Text = Me.SupCb.Items(0).ToString
                S.Cells(S.RowCount - 1, 3).Text = Me.StatCb.Items(0).ToString
                If Me.PtMdCb.Text <> "ALL" Then
                    S.Cells(S.RowCount - 1, 7).Text = Me.PtMdCb.Text
                End If
                S.Cells(S.RowCount - 1, 1).Locked = False
                'S.Cells(S.RowCount - 1, 2).Locked = False
                'S.Cells(S.RowCount - 1, 3).Locked = False
                S.Cells(S.RowCount - 1, 6).Locked = False
                S.Cells(S.RowCount - 1, 7).Locked = False
                S.Cells(S.RowCount - 1, 8).Locked = False
                S.Rows(S.RowCount - 1).ForeColor = Color.OrangeRed
                S.SetActiveCell(S.RowCount - 1, 0)
                SelPoNo = NowPoNo
                TotRec = 0
                TotRec2 = 0
                Spread_AutoCol(Me.FpSpread1)
                Me.FpSpread2.ActiveSheet.Rows.Remove(0, Me.FpSpread2.ActiveSheet.RowCount)
                Me.FpSpread3.ActiveSheet.Rows.Remove(0, Me.FpSpread3.ActiveSheet.RowCount)
                'Query_Listview(Me.PartList, "exec SP_FRMPURCHASEORDER_PARTLIST '" & Site_id & "','" & Me.PtMdCb.Text & "'", True)

                For i = 0 To Me.PartList.Items.Count - 1
                    If Len(Me.PartList.Items(i).Text) > 11 Then
                        Me.PartList.Items(i).ForeColor = Color.Red
                    End If
                Next

                RightBtn()

            Else
                MessageBox.Show("[" & NowPoNo & "]'is existed!", "validation Error")
            End If
        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message, "ERROR")
        End Try
    End Sub

    Private Sub FindBtn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FindBtn.Click
        Try
            Dim QueryPo As String = ""
            Dim Condi As String = ""
            Dim i As Integer

            Bar5.Visible = False

            If Me.StatCb.Text <> "ALL" Then
                Condi = "and a.status_dv = '" & Me.StatCb.Text & "' "
            End If
            If Me.SupCb.Text <> "ALL" Then
                Condi = Condi & "and a.sup_cd = '" & Me.SupCb.Text & "' "
            End If
            If Me.ModelCb.Text <> "ALL" Then
                Condi = Condi & "and a.model = '" & Me.ModelCb.Text & "' "
            End If
            'PO HEADER 출력
            QueryPo = "select	a.po_no, max(a.po_date),(SELECT SUP_NM FROM TBL_SUPMST WHERE SUP_NO = MAX(a.sup_cd)),max(a.status_dv),count(b.part_no),sum(b.o_qty*o_price),max(a.remark),max(a.model), ISNULL((SELECT CODE_NAME FROM TBL_CODEMASTER WHERE CLASS_ID = 'R4001' AND CODE_ID = A.CUR_UNIT),'') " & _
                    "from tbl_poheader a, tbl_podetail b " & _
                    "where	a.site_id = '" & Site_id & "' and a.po_date between '" & Me.POStDate.Text & "' and '" & Me.POEdDate.Text & " 23:59:59' and a.po_no = b.po_no " & Condi & _
                    "group by a.po_no, A.CUR_UNIT " & _
                    "order by max(a.po_date) desc, a.po_no desc"

            FpSpread1.ActiveSheet.Rows.Clear()
            Me.DockContainerItem5.Text = "P/O DETAIL"
            FpSpread2.ActiveSheet.Rows.Clear()
            TotRec = 0
            FpSpread3.ActiveSheet.Rows.Clear()
            TotRec2 = 0

            NowPoNo = ""
            SelPoNo = ""

            ButtonItem1.Text = "Close &Rceiving"
            dispRcv()

            If Query_Spread(Me.FpSpread1, QueryPo, 1) = True Then
                For i = 0 To Me.FpSpread1.ActiveSheet.RowCount - 1
                    If Me.FpSpread1.ActiveSheet.GetValue(i, 3) <> "OPENED" And Me.FpSpread1.ActiveSheet.GetValue(i, 3) <> "APPROVED" Then
                        Me.FpSpread1.ActiveSheet.Rows(i).BackColor = Color.YellowGreen
                    End If
                    Me.FpSpread1.ActiveSheet.Rows(i).Locked = True
                Next
            End If
            Spread_AutoCol(Me.FpSpread1)
            PoRec = Me.FpSpread1.ActiveSheet.RowCount
            RightBtn()
        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message, "ERROR")
        End Try

    End Sub
    Private Sub FpSpread_Fill() '파트리스트에서 선택한 파트를 해당 스프레드시트에 넣고, 넣은 파트는 리스트에서 삭제

        Try
            Dim S2 = Me.FpSpread2.ActiveSheet
            Dim S1 = Me.FpSpread1.ActiveSheet

            If TotRec > 0 Then
                Exit Sub
            End If

            If S1.RowCount = 0 Then
                Modal_Error("발주 내역을 생성하십시오")
                Exit Sub
            End If

            If S1.Cells(S1.ActiveRowIndex, 3).Text = Me.StatCb.Items(1).ToString Then
                Modal_Error("발주상태가 CLOSED 입니다")
                Exit Sub
            End If

            Dim i, j As Integer
            Dim k As Integer = Me.PartList.SelectedItems.Count
            j = 0
            For i = 0 To Me.PartList.SelectedItems.Count - 1

                S2.RowCount += 1
                S2.Cells(S2.RowCount - 1, 0).Text = S1.Cells(S1.ActiveRowIndex, 0).Text
                S2.Cells(S2.RowCount - 1, 1).Text = Me.PartList.SelectedItems(i).SubItems(0).Text
                S2.Cells(S2.RowCount - 1, 2).Text = Me.PartList.SelectedItems(i).SubItems(1).Text
                S2.Cells(S2.RowCount - 1, 3).Text = Me.PartList.SelectedItems(i).SubItems(2).Text
                S2.Cells(S2.RowCount - 1, 4).Text = Me.PartList.SelectedItems(i).SubItems(4).Text
                S2.Cells(S2.RowCount - 1, 6).Text = 0
                S2.Cells(S2.RowCount - 1, 7).Text = Me.StatCb.Items(0).ToString
                S2.Cells(S2.RowCount - 1, 8).Text = Query_RS("select ASSY_DV from tbl_PARTMASTER where site_id = '" & Site_id & "' and part_no = '" & S2.Cells(S2.RowCount - 1, 1).Text & "'")
                S2.Cells(S2.RowCount - 1, 9).Text = Query_RS("select case count(p_no) when 0 then '' when 1 then max(p_no) else 'COMM' end from tbl_bom where site_id = '" & Site_id & "' and c_no = '" & S2.Cells(S2.RowCount - 1, 1).Text & "'")
                S2.Cells(S2.RowCount - 1, 10).Text = Me.PartList.SelectedItems(i).SubItems(5).Text
                S2.Cells(S2.RowCount - 1, 5).Locked = False
                S2.Cells(S2.RowCount - 1, 7).Locked = True
                S2.Rows(S2.RowCount - 1).ForeColor = Color.OrangeRed

            Next
            CalcPart()
            For i = Me.PartList.SelectedItems.Count - 1 To 0 Step -1
                Me.PartList.Items.Remove(Me.PartList.SelectedItems.Item(i))
            Next
            Spread_AutoCol(Me.FpSpread2)
        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message, "ERROR")
        End Try
    End Sub
    Private Sub FpSpread1_CellDoubleClick(ByVal sender As Object, ByVal e As FarPoint.Win.Spread.CellClickEventArgs) Handles FpSpread1.CellDoubleClick
        Try
            If Me.FpSpread1.ActiveSheet.RowCount < 1 Then
                Exit Sub
            End If

            If FpSpread1.ActiveSheet.Cells(e.Row, 3).Text = Me.StatCb.Items(0).ToString Then
                Select Case e.Column
                    Case 2
                        Me.FpSpread1.ActiveSheet.Cells(e.Row, e.Column).Locked = False
                        Chg_ComboCell(sender, e.Row, e.Column, Query_Cell_Code2("select SUP_NM from TBL_SUPMST where site_id = '" & Site_id & "'"))
                    Case 3
                        'Me.FpSpread1.ActiveSheet.Cells(e.Row, e.Column).Locked = False
                        'Chg_ComboCell(sender, e.Row, e.Column, Query_Cell_Code1("code_name", "R0004"))
                    Case 6
                        Me.FpSpread1.ActiveSheet.Cells(e.Row, e.Column).Locked = False
                    Case 7
                        Me.FpSpread1.ActiveSheet.Cells(e.Row, e.Column).Locked = False
                        Chg_ComboCell(sender, e.Row, e.Column, Query_Cell_Code2("select model_no from tbl_modelmaster where site_id = '" & Site_id & "' and active='Y'"))
                End Select

                Spread_AutoCol(FpSpread1)
            End If
        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message, "ERROR")
        End Try
    End Sub
    Private Sub FpSpread2_CellDoubleClick(ByVal sender As Object, ByVal e As FarPoint.Win.Spread.CellClickEventArgs) Handles FpSpread2.CellDoubleClick
        '파트 리시빙
        Try
            Dim S3 = Me.FpSpread3.ActiveSheet
            Dim S2 = Me.FpSpread2.ActiveSheet
            Dim totamt As Decimal = 0D
            Dim i As Integer = 0
            Dim rowidx As Integer
            If FpSpread2.ActiveSheet.RowCount > 0 And e.Row >= TotRec Then
                Select Case e.Column
                    Case 5
                        If Me.FpSpread2.ActiveSheet.Cells(e.Row, 6).Text = 0 Then
                            Me.FpSpread2.ActiveSheet.Cells(e.Row, e.Column).Locked = False
                        End If
                        'Case 7
                        ' If Me.FpSpread2.ActiveSheet.Cells(e.Row, 7).Text = Me.StatCb.Items(0).ToString Then
                        'Chg_ComboCell(sender, e.Row, e.Column, Query_Cell_Code1("code_name", "R0004"))
                        ' Me.FpSpread2.ActiveSheet.Cells(e.Row, e.Column).Locked = False
                        'End If
                End Select
            Else
                If Query_RS("select STATUS_DV from tbl_poheader where po_no='" & S2.Cells(e.Row, 0).Text & "'") = "OPENED" Then
                    Modal_Error("미승인 상태입니다.")
                    Exit Sub
                End If


                If CInt(Query_RS("select count(*) from tbl_poheader where po_no='" & S2.Cells(e.Row, 0).Text & "'")) > 0 Then
                    If Me.Bar3.Visible = False Then
                        Exit Sub
                    End If

                    If S2.Cells(e.Row, 5).Text = "" Then
                        MessageBox.Show("[" & S2.ColumnHeader.Columns(5).Label & "] 을 입력하십시오!!!", "validation Error")

                        Exit Sub
                    End If

                    If e.Row >= TotRec Then 'P/O의 Part가 저장되어 있을 경우에만 허용
                        MessageBox.Show("선택한 품목이 저장되지 않았습니다!!!", "validation Error")
                    End If

                    If S2.Cells(e.Row, 7).Text <> Me.StatCb.Items(0).ToString Then ' STATUS가 OPENED인 경우에만 허용
                        MessageBox.Show("주문상태가 " & S2.Cells(e.Row, 7).Text & "입니다.", "validation Error")
                        Exit Sub
                    End If

                    If Me.ButtonItem1.Text = "Open &Rceiving" Then
                        dispRcv()
                    End If

                    Dim J As String() = SPREAD_SEARCH(Me.FpSpread3, 0, S2.Cells(e.Row, 1).Text, TotRec2, 1, S3.RowCount - 1, 1, False)

                    If J(0) >= TotRec2 And J(1) >= 0 Then
                        MessageBox.Show("이미 입고 완료 되었습니다!!", "validation Error")
                        Exit Sub
                    End If


                    If S3.RowCount < 1 Then
                        S3.RowCount += 1
                    End If
                    rowidx = S3.RowCount


                    If (rowidx - 1) Mod 2 = 0 Then
                        S3.Rows(rowidx - 1).BackColor = Color.Beige
                    Else
                        S3.Rows(rowidx - 1).BackColor = Color.White
                    End If

  
                    S3.Cells(rowidx - 1, 0).Text = S2.Cells(e.Row, 0).Text
                    S3.Cells(rowidx - 1, 1).Text = S2.Cells(e.Row, 1).Text
                    S3.Cells(rowidx - 1, 2).Text = S2.Cells(e.Row, 4).Text
                    If Site_id = "S1000" Then '스텔라의 경우, r_qty에 o_qty - in_qty 한 수량을 일괄적으로 표시
                        S3.Cells(rowidx - 1, 3).Text = S2.GetValue(e.Row, 5) - S2.GetValue(e.Row, 6)
                        S3.Cells(rowidx - 1, 6).Text = CDec(S3.Cells(rowidx - 1, 2).Text) * CInt(S3.Cells(rowidx - 1, 3).Text)
                    Else
                        S3.Cells(rowidx - 1, 3).Text = ""
                        S3.Cells(rowidx - 1, 6).Value = 0
                    End If
                   
                    S3.Cells(rowidx - 1, 7).Text = Query_RS("select lg_partno from tbl_partmaster where site_id = '" & Site_id & "' and part_no = '" & S3.Cells(rowidx - 1, 1).Text & "'")
                    S3.Cells(rowidx - 1, 8).Text = Now
                    S3.Cells(rowidx - 1, 9).Text = Query_RS("select emp_nm from tbl_empmaster where site_id = '" & Site_id & "' and emp_no = '" & Emp_No & "'")
                    S3.Cells(rowidx - 1, 2).Locked = False
                    S3.Cells(rowidx - 1, 3).Locked = False
                    S3.Cells(rowidx - 1, 4).Locked = False
                    S3.Cells(rowidx - 1, 5).Locked = False

                    Spread_AutoCol(Me.FpSpread3)
                    Spread_Change(Me.FpSpread3, S3.RowCount - 1)
                    Spread_Change(Me.FpSpread2, S2.ActiveRowIndex)
                    disp_total()
                Else
                    'MessageBox.Show("P/O is not Saved!!!", "validation Error")
                End If
            End If
        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message, "ERROR")
        End Try
    End Sub
    Private Sub FpSpread1_LeaveCell(ByVal sender As Object, ByVal e As FarPoint.Win.Spread.LeaveCellEventArgs) Handles FpSpread1.LeaveCell
        Try
            Select Case e.Column
                Case 1
                    ChgNewPoNo(e.Row)
            End Select
            
        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message, "ERROR")
        End Try
    End Sub

    Private Sub SaveBtn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SaveBtn1.Click, SaveBtn.Click, bSave.Click
        Try
            Dim i, j, totstat As Integer
            Dim ClsDate As String = "NULL"
            Dim S = Me.FpSpread2.ActiveSheet
            Dim S3 = Me.FpSpread3.ActiveSheet
            Dim P = Me.FpSpread1.ActiveSheet
            Dim msg As String = ""

            'PO가 저장되어 있는 경우만 receving가능
            If CInt(Query_RS("select count(*) from tbl_poheader where po_no='" & P.Cells(P.ActiveRowIndex, 0).Text & "'")) > 0 Then
                If S3.RowCount - 1 > Me.TotRec2 Then 'part receiving에 따라, P/O details에 R_QTY가 변하기 때문에 저장시 제일 우선적으로 체크
                    S3.Rows.Remove(S3.RowCount - 1, 1)
                    'Part Receiving 저장

                    S3.SetActiveCell(0, 0) '마지막 변경셀에서 다른 셀로 포커스 이동하지 않아도 저장되도록 자동 포커스 이동

                    If Save_Vallidate(Me.FpSpread3, "RECEIVING DETAILS", New String() {3, 4}) = True Then

                        For i = TotRec2 To S3.RowCount - 1
                            If S3.Rows(i).ForeColor = Color.OrangeRed Then
                                For j = 0 To S.RowCount - 1
                                    If S.Cells(j, 1).Text = S3.Cells(i, 1).Text Then
                                        If (CInt(S.Cells(j, 6).Text) + CInt(S3.Cells(i, 3).Text)) > CInt(S.Cells(j, 5).Text) Then
                                            msg += "[" & j + 1 & " th] R_QTY > O_QTY !!" & Chr(13) 'R_QTY의 합계가 O_Qty보다 클수 없다.
                                        End If
                                    End If
                                Next
                            End If
                        Next

                        If msg <> "" Then
                            MessageBox.Show(msg, "validation Error")
                            disp_total()
                            Exit Sub
                        End If

                        For i = TotRec2 To S3.RowCount - 1
                            If S3.Rows(i).ForeColor = Color.OrangeRed Then
                                If Insert_Data("EXEC SP_FRMPURCHASEORDER_INSRCV '" & Site_id & "','" & S3.Cells(i, 0).Text & "','" & S3.Cells(i, 4).Text & "','" & S3.Cells(i, 1).Text & "'," & S3.Cells(i, 3).Value & "," & S3.Cells(i, 2).Value & ",'" & S3.Cells(i, 5).Text & "','" & Emp_No & "','검사대기창고','" & P.Cells(P.ActiveRowIndex, 2).Text & "'") = False Then
                                    MessageBox.Show("[P/O RECEIVING] Failed to Save")
                                    disp_total()
                                    Exit Sub
                                End If

                                If P.Cells(P.ActiveRowIndex, 2).Text = "LGEAI" Then
                                    If CInt(Query_RS("select count(*) from tbl_partrcv where site_id = '" & Site_id & "' and po_no = '" & S3.Cells(i, 0).Text & "' and  part_no = '" & S3.Cells(i, 1).Text & "'")) < 1 Then
                                        MessageBox.Show("[P/O RECEIVING] " & S3.Cells(i, 1).Text & " Failed to Receiving", S3.Cells(i, 0).Text)
                                    End If
                                End If
                            End If
                        Next
                    Else
                        disp_total()
                        Exit Sub
                    End If
                    disp_total()
                End If

            End If
            '
            If S.ActiveRowIndex = (S.RowCount - 1) Then   '마지막 변경셀에서 다른 셀로 포커스 이동하지 않아도 저장되도록 자동 포커스 이동
                S.SetActiveCell(S.ActiveRowIndex - 1, S.ActiveColumnIndex)
            Else
                S.SetActiveCell(S.ActiveRowIndex + 1, S.ActiveColumnIndex)
            End If

            S.SetActiveCell(0, 0)

            If S.RowCount > 0 Then

                Dim S_CD As String = Query_RS("SELECT sup_no FROM TBL_SUPMST WHERE SUP_Nm = '" & P.Cells(P.ActiveRowIndex, 2).Text & "'")

                If S_CD = "" Then
                    Modal_Error("공급처를 입력하세요.")
                    Exit Sub
                End If


                If Save_Vallidate(Me.FpSpread2, "P/O DETAILS", New String() {5, 6}) = False Then
                    Exit Sub
                End If
                totstat = 0 'status가 closed인 part의 갯수 체크

                For i = 0 To S.RowCount - 1
                    If S3.RowCount - 1 > 0 Then
                        For j = TotRec2 To S3.RowCount - 2
                            If S.Cells(i, 1).Text = S3.Cells(j, 1).Text Then
                                S.Cells(i, 6).Text = CInt(S.Cells(i, 6).Text) + CInt(S3.Cells(j, 3).Text)
                                If CInt(S.Cells(i, 6).Text) = CInt(S.Cells(i, 5).Text) Then
                                    S.Cells(i, 7).Text = Me.StatCb.Items(1).ToString  'R_QTY의 합계가 O_Qty와 같으면 파트의 RECEVING이 완료되었으므로, STATUS = CLOSED
                                Else
                                    S.Cells(i, 7).Text = Me.StatCb.Items(0).ToString 'R_QTY의 합계가 O_Qty보다 작으면 파트의 RECEVING이 미완료이므로, STATUS = OPENED
                                End If
                            End If
                        Next
                    End If
                    If S.Cells(i, 7).Text = "CLOSED" Then
                        ClsDate = "getdate()"
                        totstat += 1
                    Else
                        ClsDate = "NULL"
                    End If
                    If S.Rows(i).ForeColor = Color.OrangeRed Then
                        If i < TotRec Then
                            If Insert_Data("update tbl_podetail set o_qty = " & S.Cells(i, 5).Value & ",in_qty = " & S.Cells(i, 6).Value & ",closing_date = " & ClsDate & " ,status='" & S.Cells(i, 7).Text & "',u_person='" & Emp_No & "',u_date = getdate() where site_id = '" & Site_id & "' and po_no = '" & S.Cells(i, 0).Text & "' and part_no = '" & S.Cells(i, 1).Text & "'") = False Then
                                MessageBox.Show("[P/O DETAILS] Failed to Save")
                                Exit Sub
                            End If
                        Else
                            If Insert_Data("insert tbl_podetail(site_id,po_no,part_no,o_qty,in_qty,o_price,closing_date,status,c_person,c_date,u_person,u_date) values ('" & Site_id & "','" & S.Cells(i, 0).Text & "','" & S.Cells(i, 1).Text & "'," & S.Cells(i, 5).Value & "," & S.Cells(i, 6).Value & "," & S.Cells(i, 4).Value & ",Null,'" & S.Cells(i, 7).Text & "','" & Emp_No & "',getdate(),'" & Emp_No & "',getdate())") = False Then
                                MessageBox.Show("[P/O DETAILS] Failed to Save")
                                Exit Sub

                            End If
                        End If
                    End If
                Next

                Dim opencnt As Integer
                opencnt = CInt(Query_RS("select count(status) from tbl_podetail where po_no ='" & S.Cells(S.ActiveRowIndex, 0).Text & "' and status = 'OPENED'"))

                If opencnt = 0 And totstat > 0 Then
                    P.Cells(P.ActiveRowIndex, 3).Text = Me.StatCb.Items(1).ToString
                    P.Rows(P.ActiveRowIndex).ForeColor = Color.OrangeRed
                End If

                'POHEADER INSERT


                If P.Rows(P.ActiveRowIndex).ForeColor = Color.OrangeRed Then
                    If Query_RS("select po_no from tbl_poheader where site_id = '" & Site_id & "' and po_no = '" & S.Cells(S.ActiveRowIndex, 0).Text & "'") = S.Cells(S.ActiveRowIndex, 0).Text Then
                        If Insert_Data("update tbl_poheader set sup_cd = '" & S_CD & "',status_dv = '" & P.Cells(P.ActiveRowIndex, 3).Text & "', model = '" & P.Cells(P.ActiveRowIndex, 7).Text & "', remark = '" & P.Cells(P.ActiveRowIndex, 6).Text & "', CUR_UNIT = '" & P.Cells(P.ActiveRowIndex, 8).Text & "' ,u_person='" & Emp_No & "',u_date = getdate() where site_id = '" & Site_id & "' and po_no = '" & P.Cells(P.ActiveRowIndex, 0).Text & "'") = False Then
                            MessageBox.Show("[P/O HEADER] Failed to Save")
                            disp_total()
                            Exit Sub
                        End If
                    Else
                        If Insert_Data("insert tbl_poheader(site_id,po_no,po_date,sup_cd,status_dv,model,remark,CUR_UNIT,c_person,c_date,u_person,u_date) values ('" & Site_id & "','" & P.Cells(P.ActiveRowIndex, 0).Text & "', getdate() ,'" & S_CD & "','" & P.Cells(P.ActiveRowIndex, 3).Text & "','" & P.Cells(P.ActiveRowIndex, 7).Text & "','" & P.Cells(P.ActiveRowIndex, 6).Text & "','" & P.Cells(P.ActiveRowIndex, 8).Text & "','" & Emp_No & "',getdate(),'" & Emp_No & "',getdate())") = False Then
                            MessageBox.Show("[P/O HEADER] Failed to Save")
                            disp_total()
                            Exit Sub
                        End If
                    End If

                    NowPoNo = ""
                End If

            End If

            'If Send_Pmail("1", S.Cells(S.ActiveRowIndex, 0).Text) = True Then

            'End If

            TotRec2 = 0
            TotRec = 0
            FindBtn_Click(Me.FpSpread1, EventArgs.Empty)
            MessageBox.Show("저장되었습니다", "Message")

        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message, "ERROR")
        End Try
    End Sub

    Private Sub DelBtn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DelBtn1.Click, DelBtn.Click, bDel.Click
        Try
            Dim Rowidx As String = Me.FpSpread2.ActiveSheet.ActiveRowIndex

            If CInt(Me.FpSpread2.ActiveSheet.Cells(Rowidx, 6).Text) > 0 Then
                MessageBox.Show("Aleady Received!!", "validation Error")
                Exit Sub
            End If

            If Rowidx < TotRec Then
                      MessageBox.Show("Selected PART is saved!!" & Chr(13) & Chr(10) & "Status select CANCLED ", "Validation Error")
                Exit Sub
            End If

            Dim r As DialogResult = MessageBox.Show("Selected rows delete now?", "Selected Rows Delete", MessageBoxButtons.YesNo)

            If r = Windows.Forms.DialogResult.Yes Then

                Dim J As String() = SPREAD_SEARCH(Me.FpSpread3, 0, Me.FpSpread2.ActiveSheet.Cells(Rowidx, 1).Text, 0, 1, Me.FpSpread3.ActiveSheet.RowCount - 1, 1, False)

                If J(0) >= 0 And J(1) >= 0 Then
                    Me.FpSpread3.ActiveSheet.RemoveRows(J(0), 1)
                    Me.FpSpread3.ActiveSheet.RemoveRows(TotRec2 - 1, 1)
                    disp_total()
                    Exit Sub
                End If

                Query_Listview2(Me.PartList, "EXEC SP_FRMPURCHASEORDER_PARTDEL '" & Site_id & "','" & Me.PtMdCb.Text & "','" & Me.FpSpread2.ActiveSheet.Cells(Rowidx, 1).Text & "'", False)


                Me.FpSpread2.ActiveSheet.RemoveRows(Rowidx, 1)
                CalcPart()

                MessageBox.Show("삭제되었습니다", "Message")
            End If

        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message, "ERROR")
        End Try
    End Sub


    Private Sub CalcPart() 'PO HEADER의 CNT, AMOUNT 값 계산
        Try
            Dim i As Integer
            Dim PoCnt As Integer = 0
            Dim PoAmt As Decimal = 0

            If Me.FpSpread2.ActiveSheet.RowCount > 0 Then
                For i = 0 To Me.FpSpread2.ActiveSheet.RowCount - 1
                    If Me.FpSpread2.ActiveSheet.Cells(i, 5).Text <> "" Then
                        PoCnt += CInt(Me.FpSpread2.ActiveSheet.Cells(i, 5).Text)
                        PoAmt += CInt(Me.FpSpread2.ActiveSheet.Cells(i, 5).Text) * CDec(Me.FpSpread2.ActiveSheet.Cells(i, 4).Text)
                    End If
                Next
            End If
            Me.FpSpread1.ActiveSheet.Cells(Me.FpSpread1.ActiveSheet.ActiveRowIndex, 4).Text = Me.FpSpread2.ActiveSheet.RowCount
            Me.FpSpread1.ActiveSheet.Cells(Me.FpSpread1.ActiveSheet.ActiveRowIndex, 5).Text = PoAmt
        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message, "ERROR")
        End Try
    End Sub


    Private Sub ChgNewPoNo(ByVal rowidx As Integer) '새 P/O no 생성
        Try
            Dim i As Integer
            Dim ChkDmodify As Boolean = False
            If NowPoNo = FpSpread1.ActiveSheet.Cells(rowidx, 0).Text Then '새롭게 만든 PO에 대해서만 PO NO 수정이 가능하게 함
                If Me.FpSpread2.ActiveSheet.RowCount > 0 Then
                    If Me.FpSpread2.ActiveSheet.Cells(Me.FpSpread2.ActiveSheet.ActiveRowIndex, 0).Text = NowPoNo Then
                        If MessageBox.Show("[" & NowPoNo & "]'s DETAILS Exist" & Chr(13) & Chr(10) & "DETAILS's modify P/O NO?", "P/O NO MODIFY ", MessageBoxButtons.YesNo) = Windows.Forms.DialogResult.Yes Then
                            ChkDmodify = True
                        Else
                            FpSpread1.UndoManager.Undo()
                            Exit Sub
                        End If
                    End If
                End If

                If Me.PtMdCb.Text = "" Then
                    MessageBox.Show("Select PartList's Model !!!", "Validation Error")
                    Exit Sub
                End If

                Create_PO()
                SelPoNo = SelPoNo
                FpSpread1.ActiveSheet.Cells(rowidx, 0).Text = NowPoNo

                If ChkDmodify = True Then
                    For i = 0 To Me.FpSpread2.ActiveSheet.RowCount - 1
                        Me.FpSpread2.ActiveSheet.Cells(i, 0).Text = NowPoNo
                    Next
                End If

            End If
        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message, "ERROR")
        End Try
    End Sub

    Private Sub PrtBtn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PrtBtn1.Click, PrtBtn.Click, bPrint.Click

        If save_excel = "" Then
            MessageBox.Show("Select Spread for Printing!!")
            Exit Sub
        End If

        If save_excel = "FpSpread1" Then
            If Spread_Print(Me.FpSpread1, "Receiving Summary", 0) = False Then
                MsgBox("Fail to Print")
            End If
        ElseIf save_excel = "FpSpread2" Then
            If Me.FpSpread2.ActiveSheet.RowCount > 0 Then
                If Spread_Print(Me.FpSpread2, "P/O DETAILS[" & Me.FpSpread2.ActiveSheet.Cells(Me.FpSpread2.ActiveSheet.ActiveRowIndex, 0).Text & "]", 1) = False Then
                    MsgBox("Fail to Print")
                End If
            End If
        ElseIf save_excel = "FpSpread3" Then
            If Spread_Print(Me.FpSpread3, "Receiving details", 0) = False Then
                MsgBox("Fail to Print")
            End If
        ElseIf save_excel = "FpSpread5" Then
            If Spread_Print3(Me.FpSpread5, 0) = False Then
                MsgBox("Fail to Print")
            End If
        Else
            MessageBox.Show("Select Spread for Printing!!")
        End If

    End Sub

    Private Sub Excel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles XlsBtn1.Click, Excel.Click, bExcel.Click

        If save_excel = "" Then
            MessageBox.Show("Select Spread for Saving!!")
            Exit Sub
        End If

        If save_excel = "FpSpread1" Then
            'File_Save(SaveFileDialog1, FpSpread1)
        ElseIf save_excel = "FpSpread1" Then
            File_Save(SaveFileDialog1, FpSpread1)
        ElseIf save_excel = "FpSpread2" Then
            File_Save(SaveFileDialog1, FpSpread2)
        ElseIf save_excel = "FpSpread3" Then
            File_Save(SaveFileDialog1, FpSpread2)
        ElseIf save_excel = "FpSpread5" Then
            File_Save(SaveFileDialog1, FpSpread5)
        Else
            MessageBox.Show("Select Spread for Saving!!")
        End If

    End Sub


    Private Sub disp_list(ByVal rowidx As String)
        Try
            Dim i As Integer
            Dim S1 = Me.FpSpread1.ActiveSheet
            Dim S2 = Me.FpSpread2.ActiveSheet
            Dim S3 = Me.FpSpread3.ActiveSheet

            SelPoNo = S1.GetValue(rowidx, 0)

            Me.FpSpread2.ActiveSheet.Rows.Remove(0, Me.FpSpread2.ActiveSheet.RowCount)
            Me.FpSpread3.ActiveSheet.Rows.Remove(0, Me.FpSpread3.ActiveSheet.RowCount)
            TotRec = S2.RowCount
            TotRec2 = S3.RowCount

            Query_Listview(Me.PartList, "exec SP_FRMPURCHASEORDER_PARTLIST '" & Site_id & "','" & Me.PtMdCb.Text & "'", True)


            If Query_Spread(Me.FpSpread2, "exec SP_FRMPURCHASEORDER_PODETAIL '" & Site_id & "','" & S1.Cells(rowidx, 0).Text & "'", 1) = True Then

                S2.SortRows(1, True, False)
                For i = 0 To S2.RowCount - 1
                    If S2.GetValue(i, 7) <> "OPENED" Then
                        S2.Rows(i).BackColor = Color.YellowGreen
                    End If
                    S2.Rows(i).Locked = True
                    If Me.PartList.Items.Count > 0 Then
                        Remove_Listview(Me.PartList, S2.Cells(i, 1).Text)
                    End If
                Next
                Spread_AutoCol(Me.FpSpread2)
                TotRec = S2.RowCount
            End If

            If Query_Spread(Me.FpSpread3, "SELECT po_no,PART_NO,PRICE,R_QTY,S_NO,REMARK, PRICE * R_QTY, (select lg_partno from tbl_partmaster where site_id = '" & Site_id & "' and part_no = a.part_no) , c_date ,(select emp_nm from tbl_empmaster where site_id = '" & Site_id & "' and emp_no = a.c_person) FROM TBL_PARTRCV a WHERE SITE_ID = '" & Site_id & "' AND PO_NO = '" & Me.FpSpread1.ActiveSheet.Cells(rowidx, 0).Text & "' ORDER BY PART_NO", 1) = True Then

                For i = 0 To S3.RowCount - 1
                    If Query_RS("select status from tbl_podetail where site_id = '" & Site_id & "' and po_no = '" & Me.FpSpread1.ActiveSheet.Cells(rowidx, 0).Text & "'") <> "OPENED" Then
                        S3.Rows(i).BackColor = Color.YellowGreen
                    End If
                    S3.Rows(i).Locked = True
                Next
                S3.SortRows(1, True, False)
                Spread_AutoCol(Me.FpSpread3)
                TotRec2 = S3.RowCount
                disp_total()
            End If
        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message, "ERROR")
        End Try
    End Sub

    Private Sub disp_total()
        Try
            Dim S3 = Me.FpSpread3.ActiveSheet
            Dim i As Integer
            Dim totamt As Decimal = 0
            S3.RowCount += 1

            For i = 0 To S3.RowCount - 2
                If S3.Cells(i, 6).Text <> "" Then
                    totamt += CDec(S3.Cells(i, 6).Text)
                End If
            Next

            S3.Rows(S3.RowCount - 1).BackColor = Color.Aquamarine
            S3.Cells(S3.RowCount - 1, 1).Text = "Total"
            S3.Cells(S3.RowCount - 1, 3).Text = S3.RowCount - 1
            S3.Cells(S3.RowCount - 1, 6).Text = totamt
        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message, "ERROR")
        End Try
    End Sub

    Private Sub bClear_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bClear.Click
        Try
            If Me.FpSpread2.ActiveSheet.RowCount > 0 Then
                Dim i As Integer
                If TotRec > 0 Then
                    MessageBox.Show("Selected PART is saved!!" & Chr(13) & Chr(10) & "Status select CANCLED ", "Validation Error")
                    Exit Sub
                End If
                Dim r = MessageBox.Show("Selected P/O is Clear?", "P/O Clear", MessageBoxButtons.OKCancel)
                If r = Windows.Forms.DialogResult.OK Then
                    'If NowPoNo = Me.FpSpread1.ActiveSheet.Cells(Me.FpSpread1.ActiveSheet.ActiveRowIndex, 0).Text Then
                    'Me.FpSpread1.ActiveSheet.Rows.Clear()
                    'End If
                    Me.FpSpread1.ActiveSheet.Rows.Remove(0, Me.FpSpread1.ActiveSheet.RowCount)
                    Me.FpSpread2.ActiveSheet.Rows.Remove(0, Me.FpSpread2.ActiveSheet.RowCount)
                    Me.FpSpread3.ActiveSheet.Rows.Remove(0, Me.FpSpread3.ActiveSheet.RowCount)
                    Query_Listview(Me.PartList, "exec SP_FRMPURCHASEORDER_PARTLIST '" & Site_id & "','" & Me.PtMdCb.Text & "'", True)
                    For i = 0 To Me.PartList.Items.Count - 1
                        If Len(Me.PartList.Items(i).Text) > 11 Then
                            Me.PartList.Items(i).ForeColor = Color.Red
                        End If
                    Next
                    NowPoNo = ""
                End If
            End If
        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message, "ERROR")
        End Try
    End Sub

    Private Sub cDel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cDel.Click
        Try
            If Me.FpSpread3.ActiveSheet.RowCount - 1 > 0 Then
                If Me.FpSpread3.ActiveSheet.ActiveRowIndex >= TotRec2 Then
                    Dim r = MessageBox.Show("Selected rows delete now?", "Selected Rows Delete", MessageBoxButtons.YesNo)
                    If r = Windows.Forms.DialogResult.Yes Then
                        Me.FpSpread3.ActiveSheet.RemoveRows(Me.FpSpread3.ActiveSheet.ActiveRowIndex, 1)
                        Me.FpSpread3.ActiveSheet.RemoveRows(Me.FpSpread3.ActiveSheet.RowCount - 1, 1)
                        disp_total()
                    End If
                Else
                    MessageBox.Show("Selected row must not be delete", "Valdition Error")
                End If
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error")
        End Try
    End Sub

    Private Sub cClear_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cClear.Click
        Try
            If Me.FpSpread3.ActiveSheet.RowCount > 0 Then
                Dim r = MessageBox.Show("Receiving'Data is Clear?", "Receiving Clear", MessageBoxButtons.OKCancel)
                If r = Windows.Forms.DialogResult.OK Then
                    Me.FpSpread3.ActiveSheet.RemoveRows(TotRec2, Me.FpSpread3.ActiveSheet.RowCount - TotRec2)
                    disp_total()
                End If
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error")
        End Try
    End Sub

    Private Sub dAllDown_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dAllDown.Click
        Try
            If Me.FpSpread1.ActiveSheet.RowCount > 0 Then
                PO2Excel_new(True)
            Else
                MessageBox.Show("Must be Search P/O", "Vaildation Error")
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error")
        End Try
    End Sub

    Private Sub dSelDown_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dSelDown.Click
        Try
            If Me.FpSpread1.ActiveSheet.RowCount > 0 Then
                PO2Excel_new(False)
            Else
                MessageBox.Show("Must be Search P/O", "Vaildation Error")
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error")
        End Try
    End Sub

    Private Function PO2Excel(ByVal All_YN As Boolean) As Boolean
        Try
            Dim f = Me.SaveFileDialog1
            Dim S = Me.FpSpread1.ActiveSheet
            Dim i, j As Integer
            Dim fspl As String()
            Dim filepath As String
            f.InitialDirectory = My.Computer.FileSystem.SpecialDirectories.MyDocuments
            f.Filter = "Microsoft Office Excel File (*.xls)|*.xls*|All Files(*.*)|*.*"
            f.FileName = S.Cells(S.ActiveRowIndex, 0).Text & ".xls"
            f.AddExtension = True
            If f.ShowDialog() = Windows.Forms.DialogResult.OK Then
                If All_YN = False Then
                    If S.Cells(S.ActiveRowIndex, 3).Text <> "OPENED" Then
                        MessageBox.Show("Status is not OPENED", "Validation error")
                        Exit Function
                    End If
                    DB2XLS(Me.FpSpread4, f.FileName, S.Cells(S.ActiveRowIndex, 0).Text, False, "*" & S.Cells(S.ActiveRowIndex, 0).Text & ":GMSP UPLOAD PO.", "select part_no + ':' + cast(o_qty as varchar(10)) from tbl_podetail where status='OPENED' and po_no ='" & S.Cells(S.ActiveRowIndex, 0).Text & "'")
                Else
                    For i = 0 To S.RowCount - 1
                        If S.Cells(i, 3).Text = "OPENED" Then
                            filepath = ""
                            fspl = Split(f.FileName, "\")
                            For j = 0 To fspl.Length - 2
                                filepath += fspl(j).ToString & "\"
                            Next
                            DB2XLS(Me.FpSpread4, filepath & S.Cells(i, 0).Text & ".xls", S.Cells(i, 0).Text, False, "*" & S.Cells(i, 0).Text & ":GMSP UPLOAD PO.", "select part_no + ':' + cast(o_qty as varchar(10)) from tbl_podetail where status='OPENED' and po_no ='" & S.Cells(i, 0).Text & "'")
                        End If
                    Next
                End If
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error")
        End Try
    End Function

    Private Function PO2Excel_new(ByVal All_YN As Boolean) As Boolean
        Try
            Dim f = Me.SaveFileDialog1
            Dim S = Me.FpSpread1.ActiveSheet
            Dim i, j As Integer
            Dim fspl As String()
            Dim filepath As String
            f.InitialDirectory = My.Computer.FileSystem.SpecialDirectories.MyDocuments
            f.Filter = "Microsoft Office Excel File (*.xls)|*.xls*|All Files(*.*)|*.*"
            f.FileName = S.Cells(S.ActiveRowIndex, 0).Text & ".xls"
            f.AddExtension = True
            If f.ShowDialog() = Windows.Forms.DialogResult.OK Then
                If All_YN = False Then
                    If S.Cells(S.ActiveRowIndex, 3).Text <> "OPENED" Then
                        MessageBox.Show("Status is not OPENED", "Validation error")
                        Exit Function
                    End If
                    DB2XLS_NEW(Me.FpSpread4, f.FileName, S.Cells(S.ActiveRowIndex, 0).Text, False, "select b.lg_partno, cast(a.o_qty as varchar(10)) from tbl_podetail a, tbl_partmaster b where a.site_id = '" & Site_id & "' and a.site_id = b.site_id and a.status='OPENED' and a.po_no ='" & S.Cells(S.ActiveRowIndex, 0).Text & "' and a.part_no = b.part_no")
                Else
                    For i = 0 To S.RowCount - 1
                        If S.Cells(i, 3).Text = "OPENED" Then
                            filepath = ""
                            fspl = Split(f.FileName, "\")
                            For j = 0 To fspl.Length - 2
                                filepath += fspl(j).ToString & "\"
                            Next
                            DB2XLS_NEW(Me.FpSpread4, filepath & S.Cells(i, 0).Text & ".xls", S.Cells(i, 0).Text, False, "select b.lg_partno, cast(a.o_qty as varchar(10)) from tbl_podetail a, tbl_partmaster b where a.site_id = '" & Site_id & "' and a.site_id = b.site_id and a.status='OPENED' and a.po_no ='" & S.Cells(i, 0).Text & "' and a.part_no = b.part_no")
                        End If
                    Next
                End If
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error")
        End Try
    End Function

    Private Sub bsClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bsClose.Click
        Try
            Dim S2 = Me.FpSpread2.ActiveSheet
            If S2.RowCount < 1 Then
                MessageBox.Show("Nothing row !! ", "Validation Error")
                Exit Sub
            End If
            If S2.GetValue(S2.ActiveRowIndex, 7) = Me.StatCb.Items(0).ToString Then
                Dim r = MessageBox.Show("Are You Modify to [" & S2.GetValue(S2.ActiveRowIndex, 7) & "]'s Status CLOSED?, ", "Status Modify", MessageBoxButtons.OKCancel)
                If r = Windows.Forms.DialogResult.OK Then
                    S2.SetValue(S2.ActiveRowIndex, 7, Me.StatCb.Items(1).ToString)
                    S2.Rows(S2.ActiveRowIndex).ForeColor = Color.OrangeRed
                End If
            Else
                MessageBox.Show("Status is " & S2.GetValue(S2.ActiveRowIndex, 7), "Validation Error")
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error")
        End Try
    End Sub

    Private Sub bsCancled_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bsCancled.Click
        Try
            Dim S2 = Me.FpSpread2.ActiveSheet
            If S2.RowCount < 1 Then
                MessageBox.Show("Nothing row !! ", "Validation Error")
                Exit Sub
            End If
            If S2.GetValue(S2.ActiveRowIndex, 7) = Me.StatCb.Items(0).ToString Then
                If S2.GetValue(S2.ActiveRowIndex, 6) > 0 Then
                    MessageBox.Show("Selected Part is received Part !!", "Validation Error")
                    Exit Sub
                End If
                Dim r = MessageBox.Show("Are you Modify to [" & S2.GetValue(S2.ActiveRowIndex, 7) & "]'s Status CANCLED?, ", "Status Modify", MessageBoxButtons.OKCancel)
                If r = Windows.Forms.DialogResult.OK Then
                    S2.SetValue(S2.ActiveRowIndex, 7, Me.StatCb.Items(2).ToString)
                    S2.Rows(S2.ActiveRowIndex).ForeColor = Color.OrangeRed
                End If
            Else
                MessageBox.Show("Status is " & S2.GetValue(S2.ActiveRowIndex, 7), "Validation Error")
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error")
        End Try
    End Sub

    Private Sub dsClosed_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dsClosed.Click
        Try
            Dim S1 = Me.FpSpread1.ActiveSheet
            Dim S2 = Me.FpSpread2.ActiveSheet
            Dim i As Integer
            If S2.RowCount < 1 Then
                MessageBox.Show("Nothing row !! ", "Validation Error")
                Exit Sub
            End If
            If S1.GetValue(S1.ActiveRowIndex, 3) = Me.StatCb.Items(0).ToString Then

                Dim r = MessageBox.Show("Are you modify to [" & S1.GetValue(S1.ActiveRowIndex, 0) & "]'s Status CLOSED ?, OK is to save status ", "Status Modify", MessageBoxButtons.OKCancel)
                If r = Windows.Forms.DialogResult.OK Then
                    S1.SetValue(S1.ActiveRowIndex, 3, Me.StatCb.Items(1).ToString)
                    For i = 0 To S2.RowCount - 1
                        If S2.GetValue(i, 7) = Me.StatCb.Items(0).ToString Then
                            S2.SetValue(i, 7, Me.StatCb.Items(1).ToString)
                            S2.Rows(i).ForeColor = Color.OrangeRed
                        End If
                    Next
                    S1.Rows(S1.ActiveRowIndex).ForeColor = Color.OrangeRed
                    SaveBtn_Click(Me, System.EventArgs.Empty)
                End If
            Else
                MessageBox.Show("Status is " & S1.GetValue(S1.ActiveRowIndex, 3), "Validation Error")
                Exit Sub
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error")
        End Try
    End Sub

    Private Sub dsCancled_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dsCancled.Click
        Try
            Dim S1 = Me.FpSpread1.ActiveSheet
            Dim S2 = Me.FpSpread2.ActiveSheet
            Dim i As Integer
            Dim TotRqty As Integer = 0

            If S2.RowCount < 1 Then
                MessageBox.Show("Nothing row !! ", "Validation Error")
                Exit Sub
            End If

            If S1.GetValue(S1.ActiveRowIndex, 3) = Me.StatCb.Items(0).ToString Then
                For i = 1 To S2.RowCount - 1
                    TotRqty += CInt(S2.GetValue(i, 6))
                Next
                If TotRqty > 0 Then
                    MessageBox.Show("Selected P/O is received !!", "Validation Error")
                    Exit Sub
                End If
                If S2.GetValue(S2.ActiveRowIndex, 6) > 0 Then
                    MessageBox.Show("Selected Part is received Part !!", "Validation Error")
                    Exit Sub
                End If
                Dim r = MessageBox.Show("Are you modify to [" & S1.GetValue(S1.ActiveRowIndex, 0) & "]'s Status CANCLED ?, OK is to save status ", "Status Modify", MessageBoxButtons.OKCancel)
                If r = Windows.Forms.DialogResult.OK Then
                    For i = 0 To S2.RowCount - 1
                        If S2.GetValue(i, 7) = Me.StatCb.Items(0).ToString Then
                            S2.SetValue(i, 7, Me.StatCb.Items(2).ToString)
                            S2.Rows(i).ForeColor = Color.OrangeRed
                        End If
                    Next
                    S1.SetValue(S1.ActiveRowIndex, 3, Me.StatCb.Items(2).ToString)
                    S1.Rows(S1.ActiveRowIndex).ForeColor = Color.OrangeRed
                    SaveBtn_Click(Me, System.EventArgs.Empty)
                End If
            Else
                MessageBox.Show("Status is " & S1.GetValue(S1.ActiveRowIndex, 3), "Validation Error")
                Exit Sub
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error")
        End Try
    End Sub

    Private Sub FrmGuiMgr_DockTabClosing(ByVal sender As Object, ByVal e As DevComponents.DotNetBar.DockTabClosingEventArgs) Handles FrmGuiMgr.DockTabClosing
        e.Cancel = True
    End Sub


    Private Sub ButtonItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonItem1.Click
        dispRcv()
    End Sub

    Private Sub dispRcv()

        Dim S3 = Me.FpSpread3.ActiveSheet

        If ButtonItem1.Text = "Open &Rceiving" Then
            Bar2.AutoHide = True
            Bar12.Visible = False
            Bar13.Visible = True
            ButtonItem1.Text = "Close &Rceiving"
        Else
            Bar2.AutoHide = False
            Bar12.Visible = True
            Bar13.Visible = False
            S3.Rows.Clear()
            TotRec2 = 0
            ButtonItem1.Text = "Open &Rceiving"
        End If
    End Sub

    Private Sub Create_PO()
        Try
            Dim S1 = Me.FpSpread1.ActiveSheet
            Dim podate As String
            Dim mdno As String
            If S1.RowCount < 1 Then
                podate = Now.Date
                mdno = Me.PtMdCb.Text
            Else
                podate = S1.Cells(S1.ActiveRowIndex, 1).Text
                mdno = S1.Cells(S1.ActiveRowIndex, 7).Text
            End If

            NowPoNo = Query_RS("exec SP_FrmPurchaseOrder_ETSpono '" & Site_id & "','" & podate & "','" & mdno & "','N'")
        Catch ex As Exception
            MessageBox.Show("Err: " & ex.Message, "Error")
        End Try
    End Sub

    Public Sub resrch()
        FindBtn_Click(Me, System.EventArgs.Empty)
    End Sub

    Public Sub RightBtn()
        If PoRec = 0 Then
            Me.dAllDown.Visible = False
            Me.dSelDown.Visible = False
            Me.dStatus.Visible = False
            Me.ButtonItem2.Visible = False
            Me.ButtonItem3.Visible = False
        Else
            Me.dAllDown.Visible = True
            Me.dSelDown.Visible = True
            Me.dStatus.Visible = True
            Me.ButtonItem2.Visible = True
            Me.ButtonItem3.Visible = True
        End If
    End Sub


    Private Sub ButtonItem3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonItem3.Click
        Dim pono As String = Me.FpSpread1.ActiveSheet.GetValue(Me.FpSpread1.ActiveSheet.ActiveRowIndex, 0)
        If Me.FpSpread2.ActiveSheet.RowCount > 0 Then
            If Me.FpSpread1.ActiveSheet.GetValue(Me.FpSpread1.ActiveSheet.ActiveRowIndex, 3) <> "OPENED" Then
                MessageBox.Show("Status is not OPENED!!", pono)
                Exit Sub
            End If
            Dim r = MessageBox.Show("Selected P/O is exist Partlist" & Chr(13) & Chr(10) & "Part list delete too", "Delete is Now", MessageBoxButtons.OKCancel)

            If r = Windows.Forms.DialogResult.OK Then
                If CInt(Query_RS("select count(po_no) from tbl_partrcv where site_id = '" & Site_id & "' and po_no = '" & pono & "'")) > 0 Then
                    MessageBox.Show("Alerady PO No is Receiving!!", pono)
                    Exit Sub
                Else
                    Insert_Data("exec SP_FRMPURCHASEORDER_PODEL '" & Site_id & "','" & pono & "'")
                    MessageBox.Show("삭제되었습니다 !!", pono)
                    FindBtn_Click(Me.FpSpread1, System.EventArgs.Empty)
                End If

            Else
                Exit Sub
            End If
        End If
        ' NowPoNo = ""
        ' Me.FpSpread1.ActiveSheet.Rows.Clear()
    End Sub

    Private Sub ModelCb_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ModelCb.SelectedIndexChanged
        Me.PtMdCb.Text = Me.ModelCb.Text
        PtMdCb_SelectedIndexChanged(Me.PtMdCb, System.EventArgs.Empty)
    End Sub

    Private Sub ButtonItem4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonItem4.Click
        Dim S3 = Me.FpSpread3.ActiveSheet
        Dim i As Integer

        For i = TotRec2 + 1 To S3.RowCount - 2
            S3.Cells(i, 4).Text = S3.Cells(TotRec2, 4).Text
        Next
    End Sub

    Private Sub FpSpread1_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles FpSpread1.KeyDown
        If e.Control = True And e.KeyCode = Keys.C Then
            If Me.FpSpread1.ActiveSheet.RowCount > 0 Then
                If Me.FpSpread1.ActiveSheet.ActiveColumnIndex = 0 Then
                    Clipboard.SetText(Me.FpSpread1.ActiveSheet.GetText(Me.FpSpread1.ActiveSheet.ActiveRowIndex, 0))
                    'MessageBox.Show("Copy " & Me.FpSpread1.ActiveSheet.GetValue(Me.FpSpread1.ActiveSheet.ActiveRowIndex, 0), "COPY MESSAGE")
                End If
            End If
        End If
    End Sub

    Private Sub TextBoxX1_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TextBoxX1.KeyUp
        Try
            If e.KeyCode <> Keys.Enter Then
                Exit Sub
            End If

            If TextBoxX1.Text <> "" Then
                Dim item1 As ListViewItem = PartList.FindItemWithText(TextBoxX1.Text)
                If (item1 IsNot Nothing) Then
                    'MsgBox(item1.ToString)
                    item1.EnsureVisible()
                    item1.Selected = True
                    PartList.Focus()
                End If
            End If

        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message, "ERROR")
        End Try
    End Sub


    Private Sub ButtonItem2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonItem2.Click
        Try
            Dim r = MessageBox.Show("One More Checking P/O please!!!", "Alert Modify P/O", MessageBoxButtons.YesNo)
            If r = Windows.Forms.DialogResult.No Then
                FrmModifyPO.Text = "MODIFY P/O's PART QTY"
                FrmModifyPO.partno.Text = ""
                FrmModifyPO.partname.Text = ""
                FrmModifyPO.partqty.Text = ""
                FrmModifyPO.TecNm.Text = ""
                FrmModifyPO.TecQty.Text = ""
                FrmModifyPO.OrderQty.Text = ""
                FrmModifyPO.RcvQty.Text = ""
                FrmModifyPO.LabelX5.Text = ""
                FrmModifyPO.LabelX9.Text = ""
                FrmModifyPO.ComboBoxEx1.Text = ""
                FrmModifyPO.ComboBoxEx2.Text = ""
                FrmModifyPO.LabelX1.Text = Me.FpSpread2.ActiveSheet.GetValue(Me.FpSpread2.ActiveSheet.ActiveRowIndex, 1)
                FrmModifyPO.LabelX4.Text = Me.FpSpread1.ActiveSheet.GetValue(Me.FpSpread1.ActiveSheet.ActiveRowIndex, 0)
                FrmModifyPO.TextBoxX1.Text = Me.FpSpread1.ActiveSheet.Cells(Me.FpSpread1.ActiveSheet.ActiveRowIndex, 7).Text
                FrmModifyPO.Owner = Me
                FrmModifyPO.ShowDialog()
            End If
        Catch ex As Exception
            MessageBox.Show("Err: " & ex.Message, "Error")
        End Try
    End Sub

    Private Sub ButtonItem5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonItem5.Click
        Try
            Dim r = MessageBox.Show("One More Checking P/O please!!!", "Alert Modify P/O", MessageBoxButtons.YesNo)
            If r = Windows.Forms.DialogResult.No Then
                FrmModifyPO.Text = "INSERT P/O's PART QTY"
                FrmModifyPO.partno.Text = ""
                FrmModifyPO.partname.Text = ""
                FrmModifyPO.partqty.Text = ""
                FrmModifyPO.TecNm.Text = ""
                FrmModifyPO.TecQty.Text = ""
                FrmModifyPO.OrderQty.Text = ""
                FrmModifyPO.RcvQty.Text = ""
                FrmModifyPO.LabelX5.Text = ""
                FrmModifyPO.LabelX9.Text = ""
                FrmModifyPO.TextBoxX1.Text = ""
                FrmModifyPO.LabelX1.Text = ""
                FrmModifyPO.LabelX4.Text = Me.FpSpread1.ActiveSheet.GetValue(Me.FpSpread1.ActiveSheet.ActiveRowIndex, 0)
                FrmModifyPO.ComboBoxEx1.Text = ""
                FrmModifyPO.ComboBoxEx2.Text = ""
                FrmModifyPO.ShowDialog()
            End If
        Catch ex As Exception
            MessageBox.Show("Err: " & ex.Message, "Error")
        End Try
    End Sub

    Private Sub Send_Grpmail(ByVal Wipamt As Decimal, ByVal MaxBal As Decimal) '메일 발송
        Try
            Dim kk, msrvnm, mid, mpw, fmail, mtitle As String
            Dim tmail As String()
            Dim S1 = Me.FpSpread1.ActiveSheet
            Dim TmailRs As ADODB.Recordset
            Dim i As Integer

            msrvnm = Query_RS("select code_name from tbl_codemaster where site_id = '" & Site_id & "' and class_id = 'R1001' and code_id = 'HOST'")
            mid = Query_RS("select code_name from tbl_codemaster where site_id = '" & Site_id & "' and class_id = 'R1001' and code_id = 'LOGID'")
            mpw = Query_RS("select code_name from tbl_codemaster where site_id = '" & Site_id & "' and class_id = 'R1001' and code_id = 'LOGPW'")
            fmail = Query_RS("select code_name from tbl_codemaster where site_id = '" & Site_id & "' and class_id = 'R1001' and code_id = 'FMAIL'")
            '메일 수신자 추가는 tbl_codemaster의 R1002 클래스에 추가하면 됨
            TmailRs = Query_RS_ALL("select code_name from tbl_codemaster where site_id = '" & Site_id & "' and class_id = 'R1002'")

            If TmailRs Is Nothing Then
                Exit Sub
            End If

            ReDim tmail(TmailRs.RecordCount - 1)

            For i = 0 To TmailRs.RecordCount - 1
                tmail(i) = TmailRs(0).Value
                TmailRs.MoveNext()
            Next

            '메일 본문내용 수정시 아래 부분을 고치면 됨
            mtitle = "Emergency Notice of AR Balance "
            kk = "- " & Format(Now(), "MM/dd/yy") & " ,  $" & FormatNumber(Wipamt - MaxBal, 2, , , TriState.True) & " Exeed the available ETS AR Balance. <p>"
            kk = kk & "- All part Ordering Process Suspended now.<p>"
            kk = kk & "- Must be reduce the AR balance under $" & FormatNumber(MaxBal, 2, , , TriState.True) & ". <p>"
            kk = kk & "&nbsp;&nbsp;&nbsp;&nbsp;From IcellPhone Plus administrator"

            FSendMail(msrvnm, mid, mpw, fmail, tmail, mtitle, kk)

        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message, "ERROR")
        End Try
    End Sub

    Private Sub FpSpread2_CellClick(ByVal sender As System.Object, ByVal e As FarPoint.Win.Spread.CellClickEventArgs) Handles FpSpread2.CellClick
        save_excel = "FpSpread2"
    End Sub

    Private Sub FpSpread5_CellClick(ByVal sender As System.Object, ByVal e As FarPoint.Win.Spread.CellClickEventArgs) Handles FpSpread5.CellClick
        save_excel = "FpSpread5"
    End Sub



    Private Sub ButtonX1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonX1.Click
        With FpSpread1.ActiveSheet
            If .RowCount = 0 Then
                Modal_Error("발주 내역을 조회하세요.")
                Exit Sub
            End If
        End With

        With FpSpread2.ActiveSheet
            If .RowCount = 0 Then
                Modal_Error("상세 발주 내역을 조회하세요.")
                Exit Sub
            End If
        End With

        Bar5.Visible = True

        curcell.CurrencySymbol = "\"


        With FpSpread5.ActiveSheet
            .Cells(4, 4, 7, 11).Value = ""
            .Cells(12, 2, 21, 25).Value = ""
            .Cells(12, 12, 21, 12).Value = ""
            .Cells(12, 16, 21, 16).Value = ""
            .Cells(12, 21, 21, 21).Value = ""


            '.Cells(12, 16, 21, 21).Value = 0

            .SetValue(4, 4, FpSpread1.ActiveSheet.GetValue(FpSpread1.ActiveSheet.ActiveRowIndex, 0))
            .SetValue(5, 4, "")
            .SetValue(6, 4, "")
            .SetValue(7, 4, "")


            .SetValue(12, 21, "")
            .SetValue(13, 21, "")
            .SetValue(14, 21, "")
            .SetValue(15, 21, "")
            .SetValue(16, 21, "")
            .SetValue(17, 21, "")
            .SetValue(18, 21, "")
            .SetValue(19, 21, "")
            .SetValue(20, 21, "")
            .SetValue(21, 21, "")

            .SetValue(43, 16, "")


            For I As Integer = 0 To FpSpread2.ActiveSheet.RowCount - 1
                .SetValue(12 + I, 2, FpSpread2.ActiveSheet.GetValue(I, 2) & "(" & FpSpread2.ActiveSheet.GetValue(I, 1) & ")")
                .SetValue(12 + I, 7, Query_RS("SELECT PART_SPEC FROM TBL_PARTMASTER WHERE PART_NO = '" & FpSpread2.ActiveSheet.GetValue(I, 1) & "'"))
                .SetValue(12 + I, 12, FpSpread2.ActiveSheet.GetValue(I, 5))
                .SetValue(12 + I, 14, Query_RS("SELECT ASSY_DV FROM TBL_PARTMASTER WHERE PART_NO = '" & FpSpread2.ActiveSheet.GetValue(I, 1) & "'"))
                .SetValue(12 + I, 16, FpSpread2.ActiveSheet.GetValue(I, 4))
                .Cells(12 + I, 21).CellType = curcell
                .SetValue(12 + I, 21, CInt(.GetValue(12 + I, 12)) * CInt(.GetValue(12 + I, 16)))
            Next
            .Cells(43, 16).CellType = curcell
            .SetFormula(43, 16, "SUM(V13:V42)")
            curcell.ShowSeparator = True


            .SetValue(46, 1, "발주일자 : 20" & Mid(FpSpread1.ActiveSheet.GetValue(FpSpread1.ActiveSheet.ActiveRowIndex, 0), 5, 2) & "년 " & Mid(FpSpread1.ActiveSheet.GetValue(FpSpread1.ActiveSheet.ActiveRowIndex, 0), 7, 2) & "월 " & Mid(FpSpread1.ActiveSheet.GetValue(FpSpread1.ActiveSheet.ActiveRowIndex, 0), 9, 2) & "일")

        End With


    End Sub

    Private Sub ButtonX2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonX2.Click

        With FpSpread1.ActiveSheet

            If .RowCount = 0 Then
                Modal_Error("승인할 발주건이 없습니다.")
                Exit Sub
            End If


            If .GetValue(.ActiveRowIndex, 3) <> "OPENED" Then
                Modal_Error("이미 승인 되었습니다.")
                Exit Sub
            End If



            If Insert_Data("UPDATE TBL_POHEADER SET STATUS_DV = 'APPROVED' WHERE PO_NO = '" & .GetValue(.ActiveRowIndex, 0) & "'") = True Then
                .SetValue(.ActiveRowIndex, 3, "APPROVED")

                If Send_Pmail("2", .Cells(.ActiveRowIndex, 0).Text) = True Then

                End If
                MessageBox.Show("PO가 승인되었습니다.")

            End If


        End With


    End Sub


End Class