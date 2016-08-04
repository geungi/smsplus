Imports FarPoint.Win.Spread

Public Class FrmSalesPO
    Const CtrlMask As Byte = 8
    Private NowPoNo As String = ""
    Private SelPoNo As String = ""
    Private TotRec As Integer = 0
    Private TotRec2 As Integer = 0
    Private PoRec As Integer = 0


    Private Sub FrmSalesPO_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Me.DockContainerItem1.Text = "실행 메뉴"
        Me.DockContainerItem6.Text = "조회 조건"
        Me.DockContainerItem2.Text = "품목 리스트"

        Condi_Disp() '콤보박스의 조건데이터 출력

        If Spread_Setting(FpSpread1, Me.Name) = True Then
            Spread_AutoCol(FpSpread1)
        End If

        If Spread_Setting(FpSpread2, Me.Name) = True Then
            Spread_AutoCol(FpSpread2)
        End If

        Me.ContextMenuBar1.SetContextMenuEx(Me.FpSpread2, CtxSp)
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

    End Sub

    Private Sub Condi_Disp() 'CONTROL PANEL 및 파트리스트뷰에 있는 콤보박스에 데이터 출력
        Me.POStDate.Value = Now
        Me.POEdDate.Value = Now

        '수주처 1차
        Query_Combo(Me.ModelCb, "select cus_nm from tbl_customer where site_id = '" & Site_id & "' and cus_type = '1차'  ORDER BY cus_nm")
        Me.ModelCb.Items.Add("ALL")
        Me.ModelCb.Text = "ALL"

        '수주처 2차
        Me.SupCb.Items.Add("ALL")
        Me.SupCb.Text = "ALL"

        '수주처 3차
        Me.ComboBoxEx1.Items.Add("ALL")
        Me.ComboBoxEx1.Text = "ALL"

        'Part List의 model
        Query_Combo(Me.PtMdCb, "SELECT DISTINCT SW_VER from tbl_modelmaster where site_id = '" & Site_id & "' and active='Y' ORDER BY SW_VER")
        Me.PtMdCb.Items.Add("ALL")

        'STATUS
        Me.StatCb.Text = "ALL"
        Query_Combo(Me.StatCb, "select code_name from tbl_codemaster where site_id = '" & Site_id & "' and class_id = 'R0004' order by dis_order")
        Me.StatCb.Items.Add("ALL")

    End Sub

    Private Sub ModelCb_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ModelCb.SelectedIndexChanged

        Me.SupCb.Items.Clear()
        Me.ComboBoxEx1.Items.Clear()

        Query_Combo(Me.SupCb, "select cus_nm from tbl_customer where site_id = '" & Site_id & "' and p_cus_no in (select cus_no from tbl_customer where cus_nm = '" & ModelCb.Text & "') and cus_type = '매출그룹'  ORDER BY cus_nm")
        Me.SupCb.Items.Add("ALL")
        Me.SupCb.Text = "ALL"

    End Sub
    Private Sub SupCb_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SupCb.SelectedIndexChanged

        Me.ComboBoxEx1.Items.Clear()
        Query_Combo(Me.ComboBoxEx1, "select cus_nm from tbl_customer where site_id = '" & Site_id & "' and p_cus_no in (select cus_no from tbl_customer where cus_nm = '" & SupCb.Text & "') and cus_type = '매장'  ORDER BY cus_nm")
        Me.ComboBoxEx1.Items.Add("ALL")
        Me.ComboBoxEx1.Text = "ALL"

    End Sub

    Private Sub PtMdCb_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PtMdCb.SelectedIndexChanged
        Try
            Dim i As Integer
            Query_Listview(Me.PartList, "EXEC SP_FRMSALESPO_PARTLIST '" & Site_id & "','" & Me.PtMdCb.Text & "'", True)

            'If Me.FpSpread1.ActiveSheet.RowCount = 1 Then
            '    Me.FpSpread1.ActiveSheet.SetText(0, 7, Me.PtMdCb.Text)
            'End If

            'If Me.FpSpread2.ActiveSheet.RowCount > 0 Then
            '    For i = 0 To Me.FpSpread2.ActiveSheet.RowCount - 1
            '        'Po detail에 등록되어있는 파트는 리스트뷰에서 제거하여 중복으로 등록되는 것을 방지
            '        Remove_Listview(Me.PartList, Me.FpSpread2.ActiveSheet.Cells(i, 1).Text)
            '    Next
            'End If

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
        'Try
        '    Spread_Change(sender, e.Row)

        Try
            With FpSpread2.ActiveSheet
                If .RowCount < 1 Then
                    Exit Sub
                End If

                Select Case e.Column
                    Case 3
                        .SetValue(e.Row, 7, .Cells(e.Row, 3).Text * .Cells(e.Row, 5).Text)
                    Case 5
                        .SetValue(e.Row, 7, .Cells(e.Row, 3).Text * .Cells(e.Row, 5).Text)
                End Select

                CalcPart()
                Spread_AutoCol(FpSpread2)
                Spread_AutoCol(FpSpread1)

            End With


        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message, "ERROR")
        End Try
    End Sub

    Private Sub FpSpread1_CellClick(ByVal sender As Object, ByVal e As FarPoint.Win.Spread.CellClickEventArgs) Handles FpSpread1.CellClick
        Try

            If FpSpread1.ActiveSheet.RowCount > 0 Then
                Dim AA As String = FpSpread1.ActiveSheet.GetValue(e.Row, 0)

                ' If SelPoNo <> FpSpread1.ActiveSheet.GetValue(e.Row, 0) Then
                Me.DockContainerItem5.Text = "수주 상세[" & Me.FpSpread1.ActiveSheet.GetValue(e.Row, 0) & "]"
                    disp_list(e.Row)
                'End If
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
            NowPoNo = ""
            If NowPoNo = "" Then

                If Me.ModelCb.Text = "ALL" Then
                    MessageBox.Show("수주처를 선택하십시오!!!", "Validation Error")
                    Exit Sub
                End If

                Create_PO()

                If S.RowCount > 0 Then
                    If S.Cells(S.RowCount - 1, 0).Text = NowPoNo Then
                        MessageBox.Show("[" & NowPoNo & "]는 이미 존재하는 수주번호입니다!", "validation Error")
                        Exit Sub
                    End If
                End If

                S.Rows.Clear()

                S.RowCount += 1
                S.Cells(S.RowCount - 1, 0).Text = NowPoNo
                S.Cells(S.RowCount - 1, 1).Text = Now
                S.Cells(S.RowCount - 1, 2).Text = ModelCb.Text                                          '고객
                S.Cells(S.RowCount - 1, 3).Text = ""                                                    '고객발주번호
                Chg_ComboCell(Me.FpSpread1, S.RowCount - 1, 4, Query_Cell_Code1("code_name", "R0004"))  '수주상태
                S.Cells(S.RowCount - 1, 5).Text = ""                                                    '납기일자
                S.Cells(S.RowCount - 1, 6).Text = 0                                                     '수주자재건수
                S.Cells(S.RowCount - 1, 7).Text = ""                                                    '결제조건
                S.Cells(S.RowCount - 1, 8).Text = 0                                                     '수주금액
                S.Cells(S.RowCount - 1, 9).Text = ""                                                    '발주담당
                S.Cells(S.RowCount - 1, 10).Text = ""                                                    '물류담당
                S.Cells(S.RowCount - 1, 11).Text = ""                                                    '비고

                S.Cells(S.RowCount - 1, 1).Locked = False
                S.Cells(S.RowCount - 1, 3).Locked = False
                S.Cells(S.RowCount - 1, 5).Locked = False
                S.Cells(S.RowCount - 1, 7).Locked = False
                S.Cells(S.RowCount - 1, 9).Locked = False
                S.Cells(S.RowCount - 1, 10).Locked = False
                S.Cells(S.RowCount - 1, 11).Locked = False
                S.Rows(S.RowCount - 1).ForeColor = Color.OrangeRed
                S.SetActiveCell(S.RowCount - 1, 0)

                'SelPoNo = NowPoNo
                'TotRec = 0
                'TotRec2 = 0
                'Spread_AutoCol(Me.FpSpread1)
                'Me.FpSpread2.ActiveSheet.Rows.Remove(0, Me.FpSpread2.ActiveSheet.RowCount)

                'For i = 0 To Me.PartList.Items.Count - 1
                '    If Len(Me.PartList.Items(i).Text) > 11 Then
                '        Me.PartList.Items(i).ForeColor = Color.Red
                '    End If
                'Next

                'RightBtn()

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


            'If Me.StatCb.Text <> "ALL" Then
            '    Condi = "and a.status_dv = '" & Me.StatCb.Text & "' "
            'End If

            If Me.ModelCb.Text <> "ALL" Then
                Condi = Condi & "and a.CUST_NO = '" & Query_RS("SELECT CUS_NO FROM TBL_CUSTOMER WHERE CUS_NM = '" & ModelCb.Text & "'") & "'"
            End If

            'If Me.SupCb.Text <> "ALL" Then
            '    Condi = Condi & "and a.sup_cd = '" & Me.SupCb.Text & "' "
            'End If

            'SALESMASTER 출력
            QueryPo = "select a.sales_no, a.sales_date,(SELECT cus_NM FROM TBL_CUSTOMER WHERE CUS_NO = A.CUST_NO), A.CUST_PO_NO, a.status_dv, A.DELIVERY_DATE, (SELECT count(part_no) FROM TBL_SALESDET WHERE SALES_NO = A.SALES_NO), " &
                      " a.PAY_COND, (Select sum(o_qty*o_price) FROM TBL_SALESDET WHERE SALES_NO = A.SALES_NO),A.PO_PERSON, LOG_PERSON, a.remark " &
                    "from TBL_SALESMASTER a " &
                    "where	a.site_id = '" & Site_id & "' and a.SALES_date between '" & Me.POStDate.Text & "' and '" & Me.POEdDate.Text & " 23:59:59' " & Condi &
                    "order by a.SALES_date desc, a.SALES_no desc"

            FpSpread1.ActiveSheet.Rows.Clear()
            Me.DockContainerItem5.Text = "P/O DETAIL"
            FpSpread2.ActiveSheet.Rows.Clear()
            TotRec = 0
            TotRec2 = 0

            NowPoNo = ""
            SelPoNo = ""

            ButtonItem1.Text = "Close &Rceiving"

            If Query_Spread(Me.FpSpread1, QueryPo, 1) = True Then
                For i = 0 To Me.FpSpread1.ActiveSheet.RowCount - 1
                    If Me.FpSpread1.ActiveSheet.GetValue(i, 4) <> "OPENED" And Me.FpSpread1.ActiveSheet.GetValue(i, 4) <> "APPROVED" Then
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
                '               S2.Cells(S2.RowCount - 1, 3).Text = Me.PartList.SelectedItems(i).SubItems(4).Text
                Chg_ComboCell(FpSpread2, S2.RowCount - 1, 3, Query_Cell_Code2("select UPRICE from tbl_PRICE where site_id = '" & Site_id & "' and cus_no1 = (select cus_no from tbl_customer where cus_nm = '" & ModelCb.Text & "') AND /*cus_no2 = (select cus_no from tbl_customer where cus_nm = '" & SupCb.Text & "') and*/  MODEL_NO = '" & S2.GetValue(S2.RowCount - 1, 1) & "'  ORDER BY UPRICE"))  '수주상태
                Chg_ComboCell(Me.FpSpread2, S2.RowCount - 1, 4, Query_Cell_Code1("code_name", "R0037")) ' Online or Offline

                S2.Cells(S2.RowCount - 1, 5).Text = 0
                S2.Cells(S2.RowCount - 1, 6).Text = 0
                S2.Cells(S2.RowCount - 1, 7).Text = 0

                Chg_ComboCell(FpSpread2, S2.RowCount - 1, 8, Query_Cell_Code2("select cus_nm from tbl_customer where site_id = '" & Site_id & "' and p_cus_no in (select cus_no from tbl_customer where cus_nm in ('" & ModelCb.Text & "','" & SupCb.Text & "')) and cus_type in ('매장','물류센터')  ORDER BY cus_nm"))  '수주상태
                Chg_ComboCell(Me.FpSpread2, S2.RowCount - 1, 9, Query_Cell_Code1("code_name", "R0036")) ' Online or Offline
                Chg_ComboCell(FpSpread2, S2.RowCount - 1, 10, Query_Cell_Code2("select cus_nm from tbl_customer where site_id = '" & Site_id & "' and p_cus_no in (select cus_no from tbl_customer where cus_nm in ('" & ModelCb.Text & "','" & SupCb.Text & "')) and cus_type in ('매출그룹')  ORDER BY cus_nm"))  '수주상태

                Chg_ComboCell(Me.FpSpread2, S2.RowCount - 1, 11, Query_Cell_Code1("code_name", "R0004")) ' 출고상태
                S2.Cells(S2.RowCount - 1, 12).Text = ""
                S2.Cells(S2.RowCount - 1, 13).Text = Emp_No
                S2.Cells(S2.RowCount - 1, 14).Text = Now
                S2.Cells(S2.RowCount - 1, 15).Text = Emp_No
                S2.Cells(S2.RowCount - 1, 16).Text = Now

                S2.Cells(S2.RowCount - 1, 3, S2.RowCount - 1, 5).Locked = False

                S2.Cells(S2.RowCount - 1, 6, S2.RowCount - 1, 12).Locked = False
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
            'If Me.FpSpread1.ActiveSheet.RowCount < 1 Then
            '    Exit Sub
            'End If

            'If FpSpread1.ActiveSheet.Cells(e.Row, 3).Text = Me.StatCb.Items(0).ToString Then
            '    Select Case e.Column
            '        Case 1
            '            Me.FpSpread1.ActiveSheet.Cells(e.Row, e.Column).Locked = False
            '        Case 3
            '            Me.FpSpread1.ActiveSheet.Cells(e.Row, e.Column).Locked = False
            '        Case 4
            '            Me.FpSpread1.ActiveSheet.Cells(e.Row, e.Column).Locked = False
            '        Case 5
            '            Me.FpSpread1.ActiveSheet.Cells(e.Row, e.Column).Locked = False
            '            '                        Chg_ComboCell(sender, e.Row, e.Column, Query_Cell_Code2("select model_no from tbl_modelmaster where site_id = '" & Site_id & "' and active='Y'"))
            '    End Select

            '    Spread_AutoCol(FpSpread1)
            'End If
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

            Dim ClsDate As String = "NULL"
            Dim S = Me.FpSpread2.ActiveSheet
            Dim P = Me.FpSpread1.ActiveSheet
            Dim msg As String = ""

            If S.ActiveRowIndex = (S.RowCount - 1) Then   '마지막 변경셀에서 다른 셀로 포커스 이동하지 않아도 저장되도록 자동 포커스 이동
                S.SetActiveCell(S.ActiveRowIndex - 1, S.ActiveColumnIndex)
            Else
                S.SetActiveCell(S.ActiveRowIndex + 1, S.ActiveColumnIndex)
            End If

            S.SetActiveCell(0, 0)

            If S.RowCount > 0 Then

                Dim S_CD As String = Query_RS("SELECT CUS_no FROM TBL_CUSTOMER WHERE CUS_Nm = '" & P.Cells(P.ActiveRowIndex, 2).Text & "'")

                If S_CD = "" Then
                    Modal_Error("수주처를 입력하세요.")
                    Exit Sub
                End If

                'SALESMASTER INSERT
                If P.Rows(P.ActiveRowIndex).ForeColor = Color.OrangeRed Then
                    If Query_RS("select SALES_no from TBL_SALESMASTER where site_id = '" & Site_id & "' and SALES_no = '" & S.Cells(S.ActiveRowIndex, 0).Text & "'") = S.Cells(S.ActiveRowIndex, 0).Text Then
                        If Insert_Data("update tbl_SLAESMASTER set status_dv = '" & P.Cells(P.ActiveRowIndex, 3).Text & "',  remark = '" & P.Cells(P.ActiveRowIndex, 6).Text & "',u_person='" & Emp_No & "',u_date = getdate() where site_id = '" & Site_id & "' and SALES_no = '" & P.Cells(P.ActiveRowIndex, 0).Text & "'") = False Then
                            MessageBox.Show("[수주개요] 저장 실패")

                            Exit Sub
                        End If
                    Else
                        Dim QRY As String = "insert TBL_SALESMASTER values ('" & Site_id & "',"

                        QRY = QRY & "'" & P.GetValue(P.ActiveRowIndex, 0) & "'," '수주번호
                        QRY = QRY & "'" & P.GetValue(P.ActiveRowIndex, 1) & "'," '수주일자
                        QRY = QRY & "'" & Query_RS("SELECT CUS_NO FROM TBL_CUSTOMER WHERE CUS_NM = '" & P.GetValue(P.ActiveRowIndex, 2) & "'") & "'," '공급처
                        QRY = QRY & "'" & P.GetValue(P.ActiveRowIndex, 3) & "'," '고객발주번호
                        QRY = QRY & "'OPENED'," '수주상태
                        QRY = QRY & "'" & P.GetValue(P.ActiveRowIndex, 5) & "'," '납기일자
                        '수주자재건수
                        QRY = QRY & "'" & P.GetValue(P.ActiveRowIndex, 7) & "'," '결제조건
                        '                        QRY = QRY & "" & P.GetValue(P.ActiveRowIndex, 8) & "," '수주금액
                        QRY = QRY & "'" & P.GetValue(P.ActiveRowIndex, 9) & "'," '발주담당
                        QRY = QRY & "'" & P.GetValue(P.ActiveRowIndex, 10) & "'," '물류담당
                        QRY = QRY & "'" & P.GetValue(P.ActiveRowIndex, 11) & "'," '비고
                        QRY = QRY & "'" & Emp_No & "',GETDATE(),'" & Emp_No & "',GETDATE(),NULL,NULL,NULL,NULL,NULL)"

                        If Insert_Data(QRY) = False Then
                            MessageBox.Show("[수주개요] 저장 실패")
                            Exit Sub
                        End If

                    End If
                    NowPoNo = ""
                End If

                'SALESDET INSERT 
                With FpSpread2.ActiveSheet
                    For I As Integer = 0 To .RowCount - 1
                        If .Rows(I).ForeColor = Color.OrangeRed Then

                            Dim QRY As String = "insert TBL_SALESDET values ('" & Site_id & "',"

                            QRY = QRY & "'" & .GetValue(I, 0) & "',"    '수주번호
                            QRY = QRY & "'" & .GetValue(I, 1) & "',"    '품목번호
                            QRY = QRY & "" & .GetValue(I, 3) & ","      '단가
                            QRY = QRY & "'" & .GetValue(I, 4) & "',"      '통화단위
                            QRY = QRY & "" & .GetValue(I, 5) & ",0,"      '수주수량, 출고 수량
                            QRY = QRY & "NULL,"                         'CLOSING DATE
                            QRY = QRY & "'" & Query_RS("SELECT CUS_NO FROM TBL_CUSTOMER WHERE CUS_NM = '" & .GetValue(I, 8) & "'") & "'," '물류센터
                            QRY = QRY & "'" & Query_RS("SELECT CODE_ID FROM TBL_CODEMASTER WHERE CLASS_ID = 'R0036' AND CODE_NAME = '" & .GetValue(I, 9) & "'") & "'," '포장타입
                            QRY = QRY & "'" & Query_RS("SELECT CUS_NO FROM TBL_CUSTOMER WHERE CUS_NM = '" & .GetValue(I, 10) & "'") & "'," '집계그룹
                            QRY = QRY & "'OPENED'," '출고상태
                            QRY = QRY & "'" & .GetValue(I, 12) & "',"    '비고
                            '                            QRY = QRY & "'" & Query_RS("SELECT CODE_ID FROM TBL_CODEMASTER WHERE CLASS_ID = 'R0037' AND CODE_NAME = '" & .GetValue(.ActiveRowIndex, 13) & "'") & "'," '화폐단위
                            QRY = QRY & "'" & Emp_No & "',GETDATE(),'" & Emp_No & "',GETDATE(),NULL,NULL,NULL,NULL,NULL)"

                            If Insert_Data(QRY) = False Then
                                MessageBox.Show("[수주상세] 저장 실패")
                                Exit Sub
                            End If




                        End If

                    Next


                End With


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

    Private Sub CalcPart() 'PO HEADER의 CNT, AMOUNT 값 계산
        Try
            Dim i As Integer
            Dim PoCnt As Integer = 0
            Dim PoAmt As Decimal = 0

            If Me.FpSpread2.ActiveSheet.RowCount > 0 Then
                For i = 0 To Me.FpSpread2.ActiveSheet.RowCount - 1
                    If Me.FpSpread2.ActiveSheet.Cells(i, 7).Text <> "" Then
                        'PoCnt += CInt(Me.FpSpread2.ActiveSheet.Cells(i, 5).Text)
                        PoAmt += CDec(Me.FpSpread2.ActiveSheet.Cells(i, 7).Text)
                    End If
                Next
            End If
            Me.FpSpread1.ActiveSheet.Cells(Me.FpSpread1.ActiveSheet.ActiveRowIndex, 6).Text = Me.FpSpread2.ActiveSheet.RowCount
            Me.FpSpread1.ActiveSheet.Cells(Me.FpSpread1.ActiveSheet.ActiveRowIndex, 8).Text = PoAmt
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
        ElseIf save_excel = "FpSpread5" Then
            'If Spread_Print3(Me.FpSpread5, 0) = False Then
            '    MsgBox("Fail to Print")
            'End If
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
            '    File_Save(SaveFileDialog1, FpSpread5)
        Else
            MessageBox.Show("Select Spread for Saving!!")
        End If

    End Sub


    Private Sub disp_list(ByVal rowidx As String)
        Try
            Dim i As Integer
            Dim S1 = Me.FpSpread1.ActiveSheet
            Dim S2 = Me.FpSpread2.ActiveSheet

            SelPoNo = S1.GetValue(rowidx, 0)

            Me.FpSpread2.ActiveSheet.Rows.Remove(0, Me.FpSpread2.ActiveSheet.RowCount)
            TotRec = S2.RowCount

            '            Query_Listview(Me.PartList, "exec SP_FRMPURCHASEORDER_PARTLIST '" & Site_id & "','" & Me.PtMdCb.Text & "'", True)

            Dim QRY As String = "Select "
            QRY = QRY & "SALES_no," & vbNewLine
            QRY = QRY & "part_no," & vbNewLine
            QRY = QRY & "(select model_name from tbl_modelmaster where model_no = a.part_no)," & vbNewLine
            QRY = QRY & "o_price," & vbNewLine
            QRY = QRY & "Currency," & vbNewLine
            QRY = QRY & "o_qty," & vbNewLine
            QRY = QRY & "in_qty," & vbNewLine
            QRY = QRY & "o_price*o_qty," & vbNewLine

            QRY = QRY & "(select cus_nm from tbl_customer where cus_no = a.l_cust_no)," & vbNewLine
            QRY = QRY & "(select code_name from tbl_codemaster where class_id = 'R0036' and code_id = a.PACKING_TYPE)," & vbNewLine
            QRY = QRY & "(select cus_nm from tbl_customer where cus_no = a.sales_GROUP)," & vbNewLine
            QRY = QRY & "status," & vbNewLine
            QRY = QRY & "REMARK," & vbNewLine
            QRY = QRY & "(select emp_nm from tbl_empmaster where emp_no = a.c_person)," & vbNewLine
            QRY = QRY & "c_date," & vbNewLine
            QRY = QRY & "(select emp_nm from tbl_empmaster where emp_no = a.u_person)," & vbNewLine
            QRY = QRY & "u_date" & vbNewLine
            QRY = QRY & "From TBL_SALESDET a" & vbNewLine
            QRY = QRY & "where sales_no = '" & FpSpread1.ActiveSheet.GetValue(FpSpread1.ActiveSheet.ActiveRowIndex, 0) & "'" & vbNewLine
            QRY = QRY & "" & vbNewLine
            QRY = QRY & "" & vbNewLine
            QRY = QRY & "" & vbNewLine


            If Query_Spread(Me.FpSpread2, QRY, 1) = True Then

                S2.SortRows(1, True, False)
                For i = 0 To S2.RowCount - 1
                    If S2.GetValue(i, 11) <> "OPENED" Then
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
                'mdno = S1.Cells(S1.ActiveRowIndex, 7).Text
            End If

            NowPoNo = Query_RS("exec SP_FRMSALESORDER_SALESNO '" & Site_id & "','" & podate & "'")

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

    Private Sub FpSpread5_CellClick(ByVal sender As System.Object, ByVal e As FarPoint.Win.Spread.CellClickEventArgs)
        save_excel = "FpSpread5"
    End Sub



    Private Sub ButtonX1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
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

        'Bar5.Visible = True

        curcell.CurrencySymbol = "\"


        'With FpSpread5.ActiveSheet
        '    .Cells(4, 4, 7, 11).Value = ""
        '    .Cells(12, 2, 21, 25).Value = ""
        '    .Cells(12, 12, 21, 12).Value = ""
        '    .Cells(12, 16, 21, 16).Value = ""
        '    .Cells(12, 21, 21, 21).Value = ""


        '    '.Cells(12, 16, 21, 21).Value = 0

        '    .SetValue(4, 4, FpSpread1.ActiveSheet.GetValue(FpSpread1.ActiveSheet.ActiveRowIndex, 0))
        '    .SetValue(5, 4, "")
        '    .SetValue(6, 4, "")
        '    .SetValue(7, 4, "")


        '    .SetValue(12, 21, "")
        '    .SetValue(13, 21, "")
        '    .SetValue(14, 21, "")
        '    .SetValue(15, 21, "")
        '    .SetValue(16, 21, "")
        '    .SetValue(17, 21, "")
        '    .SetValue(18, 21, "")
        '    .SetValue(19, 21, "")
        '    .SetValue(20, 21, "")
        '    .SetValue(21, 21, "")

        '    .SetValue(43, 16, "")


        '    For I As Integer = 0 To FpSpread2.ActiveSheet.RowCount - 1
        '        .SetValue(12 + I, 2, FpSpread2.ActiveSheet.GetValue(I, 2) & "(" & FpSpread2.ActiveSheet.GetValue(I, 1) & ")")
        '        .SetValue(12 + I, 7, Query_RS("SELECT PART_SPEC FROM TBL_PARTMASTER WHERE PART_NO = '" & FpSpread2.ActiveSheet.GetValue(I, 1) & "'"))
        '        .SetValue(12 + I, 12, FpSpread2.ActiveSheet.GetValue(I, 5))
        '        .SetValue(12 + I, 14, Query_RS("SELECT ASSY_DV FROM TBL_PARTMASTER WHERE PART_NO = '" & FpSpread2.ActiveSheet.GetValue(I, 1) & "'"))
        '        .SetValue(12 + I, 16, FpSpread2.ActiveSheet.GetValue(I, 4))
        '        .Cells(12 + I, 21).CellType = curcell
        '        .SetValue(12 + I, 21, CInt(.GetValue(12 + I, 12)) * CInt(.GetValue(12 + I, 16)))
        '    Next
        '    .Cells(43, 16).CellType = curcell
        '    .SetFormula(43, 16, "SUM(V13:V42)")
        '    curcell.ShowSeparator = True


        '    .SetValue(46, 1, "발주일자 : 20" & Mid(FpSpread1.ActiveSheet.GetValue(FpSpread1.ActiveSheet.ActiveRowIndex, 0), 5, 2) & "년 " & Mid(FpSpread1.ActiveSheet.GetValue(FpSpread1.ActiveSheet.ActiveRowIndex, 0), 7, 2) & "월 " & Mid(FpSpread1.ActiveSheet.GetValue(FpSpread1.ActiveSheet.ActiveRowIndex, 0), 9, 2) & "일")

        'End With


    End Sub

    Private Sub ButtonX2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

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

    Private Sub FpSpread2_TextChanged(sender As Object, e As CellClickEventArgs) Handles FpSpread2.TextChanged

        Try
            With FpSpread2.ActiveSheet
                If .RowCount < 1 Then
                    Exit Sub
                End If

                Select Case e.Column
                    Case 3
                        .SetValue(e.Row, 6, .Cells(e.Row, 3).Text * .Cells(e.Row, 4).Text)
                    Case 4
                        .SetValue(e.Row, 6, .Cells(e.Row, 3).Text * .Cells(e.Row, 4).Text)
                End Select

                Spread_AutoCol(FpSpread2)

            End With

        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message, "ERROR")
        End Try

    End Sub

End Class