Imports FarPoint.Win.Spread


Public Class FrmBoxInput
    Const CtrlMask As Byte = 8
    Private NowPoNo As String = ""
    Private SelPoNo As String = ""
    Private TotRec As Integer = 0
    Private TotRec2 As Integer = 0
    Private PoRec As Integer = 0
    Private chkModel As Boolean = False
    Private cusNM As String = ""


    Private Sub FrmBoxInput_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Me.DockContainerItem1.Text = "실행 메뉴"
        Me.DockContainerItem6.Text = "조회 조건"
        Me.DockContainerItem2.Text = "발주 목록"

        Condi_Disp() '콤보박스의 조건데이터 출력

        If Spread_Setting(FpSpread1, Me.Name) = True Then
            Spread_AutoCol(FpSpread1)
        End If

        If Spread_Setting(FpSpread2, Me.Name) = True Then
            Spread_AutoCol(FpSpread2)
        End If

        'Me.ContextMenuBar1.SetContextMenuEx(Me.FpSpread2, CtxSp)
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


        Me.NewBtn.Visible = False

        Query_Listview(Me.SalesPoList, "EXEC SP_FrmBoxInput_LIST'" & Site_id & "','ALL','" & Me.POStDate.Text & "','" & Me.POEdDate.Text & "','" & cusNM & "'", True)

    End Sub

    Private Sub Condi_Disp() 'CONTROL PANEL 및 파트리스트뷰에 있는 콤보박스에 데이터 출력


        Me.POStDate.Text = Now.Date.ToString()
        Me.POEdDate.Text = Now.Date.ToString()

        '수주처 1차
        Query_Combo(Me.CusCb1, "select cus_nm from tbl_customer where site_id = '" & Site_id & "' and cus_type = '1차'  ORDER BY cus_nm")
        Me.ModelCb.Items.Add("ALL")
        Me.ModelCb.Text = "ALL"

        '수주처 2차
        Me.CusCb2.Items.Add("ALL")
        Me.CusCb2.Text = "ALL"

        '수주처 3차
        Me.CusCb3.Items.Add("ALL")
        Me.CusCb3.Text = "ALL"

        '모델명
        Query_Combo2(Me.ModelCb, "tbl_modelmaster", "model_name", "model_no", "site_id = '" & Site_id & "' and active = 'Y'  ORDER BY model_no", True)




    End Sub

    Private Sub CusCb1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CusCb1.SelectedIndexChanged

        Me.CusCb2.Items.Clear()
        Me.CusCb3.Items.Clear()

        If Me.CusCb1.Text = "ALL" Then
            cusNM = ""
        Else
            cusNM = Me.CusCb1.Text
        End If

        Query_Combo(Me.CusCb2, "select cus_nm from tbl_customer where site_id = '" & Site_id & "' and p_cus_no in (select cus_no from tbl_customer where cus_nm = '" & CusCb1.Text & "') and cus_type = '매출그룹'  ORDER BY cus_nm")
        Me.CusCb2.Items.Add("ALL")
        Me.CusCb2.Text = "ALL"

    End Sub
    Private Sub CusCb2_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CusCb2.SelectedIndexChanged

        Me.CusCb3.Items.Clear()
        Query_Combo(Me.CusCb3, "select cus_nm from tbl_customer where site_id = '" & Site_id & "' and p_cus_no in (select cus_no from tbl_customer where cus_nm = '" & CusCb2.Text & "') and cus_type = '매장'  ORDER BY cus_nm")
        Me.CusCb3.Items.Add("ALL")
        Me.CusCb3.Text = "ALL"
    End Sub


    Private Sub PartList_ItemDrag(ByVal sender As System.Object, ByVal e As System.Windows.Forms.ItemDragEventArgs) Handles SalesPoList.ItemDrag
        If e.Button = Windows.Forms.MouseButtons.Left Then
            'invoke the drag and drop operation
            DoDragDrop(e.Item, DragDropEffects.Move Or DragDropEffects.Copy)
        End If
    End Sub

    Private Sub FpSpread2_Change(ByVal sender As Object, ByVal e As FarPoint.Win.Spread.ChangeEventArgs) Handles FpSpread2.Change
        'Try
        '    Spread_Change(sender, e.Row)

        Try
            Dim TotQty As Integer = 0
            Dim uPrice As Double = 0.0000
            Dim P = Me.FpSpread1.ActiveSheet
            With FpSpread2.ActiveSheet
                If .RowCount < 1 Then
                    Exit Sub
                End If

                Select Case e.Column
                    Case 3
                        For i = 0 To .RowCount - 1
                            TotQty += CInt(.Cells(i, 3).Text)
                            uPrice += CDbl(.Cells(i, 4).Text) * CInt(.Cells(i, 3).Text)
                        Next
                        P.Cells(P.ActiveRowIndex, 6).Text = TotQty
                        P.Cells(P.ActiveRowIndex, 7).Text = uPrice

                End Select
                If (.Rows(e.Row).ForeColor = Color.Black) Then
                    .Cells(e.Row, 5).Text = "U" '업데이트
                End If
                .Rows(e.Row).ForeColor = Color.OrangeRed
            End With


        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message, "ERROR")
        End Try
    End Sub

    Private Sub FpSpread1_CellClick(ByVal sender As Object, ByVal e As FarPoint.Win.Spread.CellClickEventArgs) Handles FpSpread1.CellClick
        Try
            save_excel = "FpSpread1"
            If e.Column = 0 Then
                If FpSpread1.ActiveSheet.RowCount > 0 Then
                    Dim AA As String = FpSpread1.ActiveSheet.GetValue(e.Row, 0)

                    ' If SelPoNo <> FpSpread1.ActiveSheet.GetValue(e.Row, 0) Then
                    Me.DockContainerItem5.Text = "포장 상세[" & Me.FpSpread1.ActiveSheet.GetValue(e.Row, 0) & "]"
                    disp_list(e.Row)
                    'End If
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


    Private Sub PartList_DragEnter(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles SalesPoList.DragEnter, FpSpread2.DragEnter
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





    Private Sub FindBtn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FindBtn.Click
        Try
            Dim QueryPo As String = ""
            Dim Condi As String = ""
            Dim CustNo1_Condi As String = ""
            Dim CustNo2_Condi As String = ""
            Dim CustNo3_Condi As String = ""
            Dim i As Integer

            ' CustCondi = " And ("

            Query_Listview(Me.SalesPoList, "EXEC SP_FRMBOXINPUT_LIST '" & Site_id & "','ALL','" & Me.POStDate.Text & "','" & Me.POEdDate.Text & "','" & cusNM & "'", True)
            Me.FpSpread1.ActiveSheet.RowCount = 0
            Me.FpSpread2.ActiveSheet.RowCount = 0
            Me.SalesPoDetail.Items.Clear()
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

            Dim S = Me.FpSpread2.ActiveSheet
            Dim P = Me.FpSpread1.ActiveSheet
            Dim msg As String = ""
            Dim totqty As Integer = 0

            If S.ActiveRowIndex = (S.RowCount - 1) Then   '마지막 변경셀에서 다른 셀로 포커스 이동하지 않아도 저장되도록 자동 포커스 이동
                S.SetActiveCell(S.ActiveRowIndex - 1, S.ActiveColumnIndex)
            Else
                S.SetActiveCell(S.ActiveRowIndex + 1, S.ActiveColumnIndex)
            End If

            S.SetActiveCell(0, 0)

            If S.RowCount = 0 Then

                MessageBox.Show("포장할 박스LOT를 선택하세요.!!!", "Validation Error")

            Else

                Dim inQty As Integer = 0

                For i = 0 To S.RowCount - 1
                    inQty = CInt(S.Cells(i, 2).Text) - CInt(S.Cells(i, 5).Text)
                    'prodmaster_b테이블에 공정이 W2000 인 것이 적재수량만큼 재고가 있는지 체크함
                    Dim prodQty As Integer = CInt(Query_RS("select count(*) from tbl_prodmaster_b where site_id = '" & Site_id & "' and model = '" & S.Cells(i, 0).Text & "' and c_prc = 'W2000'"))

                    If prodQty < inQty Then
                        MessageBox.Show("[" & S.Cells(i, 0).Text & "]개별포장공정에 있는 수량이 적재할 수량보다 적습니다.!!!", "Validation Error")
                        Exit Sub
                    End If
                Next

                For i = 0 To S.RowCount - 1
                    If Insert_Data("update tbl_packingdet set in_qty = " & S.Cells(i, 2).Text & ",u_person='" & Emp_No & "',u_date = getdate() where site_id = '" & Site_id & "' and LOT_NO = '" & S.Cells(i, 4).Text & "' and model = '" & S.Cells(i, 0).Text & "'") = False Then
                        MessageBox.Show("[변경 실패]")
                    Else

                        totqty += CInt(S.Cells(i, 2).Text)
                        inQty = CInt(S.Cells(i, 2).Text) - CInt(S.Cells(i, 5).Text)


                        Dim Qry As String = "update tbl_prodmaster_b set lot_no = '" & S.Cells(i, 4).Text & "',c_prc = 'W3000',u_person='" & Emp_No & "',u_date = getdate() where site_id = '" & Site_id & "' "
                        Qry += " And prod_no in (select top " & inQty & " prod_no from tbl_prodmaster_b where site_id = '" & Site_id & "' and model = '" & S.Cells(i, 0).Text & "' and c_prc = 'W2000' order by prod_no) "
                        Insert_Data(Qry)
                        S.Cells(i, 5).Text = S.Cells(i, 2).Text
                    End If
                Next


                NowPoNo = ""



            End If

            'MessageBox.Show("저장되었습니다", "Message")
            Insert_Data("UPDATE TBL_LOTMASTER SET C_PRC = 'w3000' WHERE SITE_ID = '" & Site_id & "' AND LOT_NO = '" & P.Cells(P.ActiveRowIndex, 0).Text & "'")
            PackingBarcode_Print()  '패킹바코드 출력
            Query_Listview(Me.SalesPoList, "EXEC SP_FrmBoxInput_LIST'" & Site_id & "','ALL','" & Me.POStDate.Text & "','" & Me.POEdDate.Text & "','" & cusNM & "'", True)
            Query_Listview(Me.SalesPoDetail, "EXEC SP_FRMBOXINPUTDET_LIST '" & Site_id & "','" & P.Cells(0, 0).Text & "','ALL'", True)

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
            MessageBox.Show("인쇄할 시트를 선택하세요.!!")
            Exit Sub
        End If

        If save_excel = "FpSpread1" Then
            If Spread_Print(Me.FpSpread1, "BOX LOT_NO 목록", 0) = False Then
                MsgBox("Fail to Print")
            End If
        ElseIf save_excel = "FpSpread2" Then
            If Me.FpSpread2.ActiveSheet.RowCount > 0 Then
                If Spread_Print(Me.FpSpread2, "BOX 적재할 제품목록[" & Me.FpSpread2.ActiveSheet.Cells(Me.FpSpread2.ActiveSheet.ActiveRowIndex, 0).Text & "]", 1) = False Then
                    MsgBox("Fail to Print")
                End If
            End If
        Else
            MessageBox.Show("인쇄할 시트를 선택하세요.!!")
        End If

    End Sub

    Private Sub Excel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles XlsBtn1.Click, Excel.Click, bExcel.Click

        If save_excel = "" Then
            MessageBox.Show("액셀전환할 시트를 선택하세요.")
            Exit Sub
        End If

        If save_excel = "FpSpread1" Then
            File_Save(SaveFileDialog1, FpSpread1)
        ElseIf save_excel = "FpSpread2" Then
            File_Save(SaveFileDialog1, FpSpread2)
        Else
            MessageBox.Show("액셀전환할 시트를 선택하세요.")
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


            Dim QRY As String = "Select LOT_NO,MODEL,(SELECT MODEL_NAME FROM TBL_MODELMASTER WHERE MODEL_NO = B.MODEL) AS MODELNM, O_QTY, price,'S'  "
            QRY += " from tbl_packingdet B where site_id = '" & Site_id & "' and lot_no = '" & S1.Cells(rowidx, 0).Text & "'"


            If Query_Spread(Me.FpSpread2, QRY, 1) = True Then
                Spread_AutoCol(Me.FpSpread2)
            End If

            If S2.RowCount > 0 Then
                For i = 0 To S2.RowCount - 1
                    S2.Cells(i, 3).Locked = False
                Next
            End If


        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message, "ERROR")
        End Try
    End Sub

    Private Sub bClear_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Try
            Dim S2 = Me.FpSpread2.ActiveSheet
            Dim S1 = Me.FpSpread1.ActiveSheet
            Dim IN_yn As Boolean = False
            Dim delQty As Integer = 0
            Dim delPrice As Double = 0.00

            If S2.RowCount = 0 Then
                MessageBox.Show("삭제할 제품이 없습니다. ", "Validation Error")
                Exit Sub
            End If

            Dim r = MessageBox.Show(S2.Cells(S2.ActiveRowIndex, 0).Text & "의 제품을 모두 삭제하시겠습니까?" & vbCrLf & "이미 적재된 제품은 삭제되지 않습니다.", "전체삭제", MessageBoxButtons.OKCancel)
            If r = Windows.Forms.DialogResult.OK Then
                For I = 0 To S2.RowCount - 1
                    If CInt(Query_RS("SELECT isnull(IN_QTY,0) FROM TBL_PACKINGDET WHERE site_id = '" & Site_id & "' and lot_no = '" & S2.Cells(I, 0).Text & "' and model = '" & S2.Cells(I, 1).Text & "'")) = 0 Then
                        If Insert_Data("delete tbl_packingdet where site_id = '" & Site_id & "' and lot_no = '" & S2.Cells(I, 0).Text & "' and model = '" & S2.Cells(I, 1).Text & "'") = False Then
                            MessageBox.Show(S2.Cells(I, 1).Text & " 삭제실패")
                        Else
                            delQty += CInt(S2.Cells(I, 3).Text)
                            delPrice += CDbl(S2.Cells(I, 4).Text()) * CInt(S2.Cells(I, 3).Text)
                        End If
                    End If

                Next

                S1.Cells(S1.ActiveRowIndex, 6).Text = CInt(S1.Cells(S1.ActiveRowIndex, 6).Text) - delQty
                S1.Cells(S1.ActiveRowIndex, 7).Text = CDbl(S1.Cells(S1.ActiveRowIndex, 7).Text) - delPrice
                S1.Rows(S1.ActiveRowIndex).ForeColor = Color.OrangeRed

                disp_list(S1.ActiveRowIndex)
            End If


        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message, "ERROR")
        End Try
    End Sub



    Private Sub dAllDown_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
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

    Private Sub dSelDown_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
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

    Private Sub bsClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Try
            Dim S2 = Me.FpSpread2.ActiveSheet
            If S2.RowCount < 1 Then
                MessageBox.Show("Nothing row !! ", "Validation Error")
                Exit Sub
            End If
            If S2.GetValue(S2.ActiveRowIndex, 7) = Me.ModelCb.Items(0).ToString Then
                Dim r = MessageBox.Show("Are You Modify to [" & S2.GetValue(S2.ActiveRowIndex, 7) & "]'s Status CLOSED?, ", "Status Modify", MessageBoxButtons.OKCancel)
                If r = Windows.Forms.DialogResult.OK Then
                    S2.SetValue(S2.ActiveRowIndex, 7, Me.ModelCb.Items(1).ToString)
                    S2.Rows(S2.ActiveRowIndex).ForeColor = Color.OrangeRed
                End If
            Else
                MessageBox.Show("Status is " & S2.GetValue(S2.ActiveRowIndex, 7), "Validation Error")
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error")
        End Try
    End Sub

    Private Sub bsCancled_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Try
            Dim S2 = Me.FpSpread2.ActiveSheet
            If S2.RowCount < 1 Then
                MessageBox.Show("Nothing row !! ", "Validation Error")
                Exit Sub
            End If
            If S2.GetValue(S2.ActiveRowIndex, 7) = Me.ModelCb.Items(0).ToString Then
                If S2.GetValue(S2.ActiveRowIndex, 6) > 0 Then
                    MessageBox.Show("Selected Part is received Part !!", "Validation Error")
                    Exit Sub
                End If
                Dim r = MessageBox.Show("Are you Modify to [" & S2.GetValue(S2.ActiveRowIndex, 7) & "]'s Status CANCLED?, ", "Status Modify", MessageBoxButtons.OKCancel)
                If r = Windows.Forms.DialogResult.OK Then
                    S2.SetValue(S2.ActiveRowIndex, 7, Me.ModelCb.Items(2).ToString)
                    S2.Rows(S2.ActiveRowIndex).ForeColor = Color.OrangeRed
                End If
            Else
                MessageBox.Show("Status is " & S2.GetValue(S2.ActiveRowIndex, 7), "Validation Error")
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error")
        End Try
    End Sub

    Private Sub dsClosed_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Try
            Dim S1 = Me.FpSpread1.ActiveSheet
            Dim S2 = Me.FpSpread2.ActiveSheet
            Dim i As Integer
            If S2.RowCount < 1 Then
                MessageBox.Show("Nothing row !! ", "Validation Error")
                Exit Sub
            End If
            If S1.GetValue(S1.ActiveRowIndex, 3) = Me.ModelCb.Items(0).ToString Then

                Dim r = MessageBox.Show("Are you modify to [" & S1.GetValue(S1.ActiveRowIndex, 0) & "]'s Status CLOSED ?, OK is to save status ", "Status Modify", MessageBoxButtons.OKCancel)
                If r = Windows.Forms.DialogResult.OK Then
                    S1.SetValue(S1.ActiveRowIndex, 3, Me.ModelCb.Items(1).ToString)
                    For i = 0 To S2.RowCount - 1
                        If S2.GetValue(i, 7) = Me.ModelCb.Items(0).ToString Then
                            S2.SetValue(i, 7, Me.ModelCb.Items(1).ToString)
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

    Private Sub dsCancled_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Try
            Dim S1 = Me.FpSpread1.ActiveSheet
            Dim S2 = Me.FpSpread2.ActiveSheet
            Dim i As Integer
            Dim TotRqty As Integer = 0

            If S2.RowCount < 1 Then
                MessageBox.Show("Nothing row !! ", "Validation Error")
                Exit Sub
            End If

            If S1.GetValue(S1.ActiveRowIndex, 3) = Me.ModelCb.Items(0).ToString Then
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
                        If S2.GetValue(i, 7) = Me.ModelCb.Items(0).ToString Then
                            S2.SetValue(i, 7, Me.ModelCb.Items(2).ToString)
                            S2.Rows(i).ForeColor = Color.OrangeRed
                        End If
                    Next
                    S1.SetValue(S1.ActiveRowIndex, 3, Me.ModelCb.Items(2).ToString)
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


    Private Sub ButtonItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub



    Private Sub Create_PO()
        Try
            Dim S1 = Me.FpSpread1.ActiveSheet
            Dim podate As String

            podate = Now.Date


            NowPoNo = Query_RS("exec SP_FrmBoxInput_LOTNO '" & Site_id & "','" & podate & "'")

        Catch ex As Exception
            MessageBox.Show("Err: " & ex.Message, "Error")
        End Try
    End Sub

    Public Sub resrch()
        FindBtn_Click(Me, System.EventArgs.Empty)
    End Sub

    Public Sub RightBtn()

    End Sub


    Private Sub ButtonItem3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
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



    Private Sub FpSpread2_CellClick(ByVal sender As System.Object, ByVal e As FarPoint.Win.Spread.CellClickEventArgs) Handles FpSpread2.CellClick
        save_excel = "FpSpread2"
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

    Private Sub CtxSp2_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub SalesPoList_SelectedIndexChanged(sender As Object, e As EventArgs) Handles SalesPoList.Click
        Dim s = Me.FpSpread1.ActiveSheet
        Dim s2 = Me.FpSpread2.ActiveSheet
        Dim rs As New ADODB.Recordset
        Dim Packtype As String = ""

        s.RowCount = 0
        s2.RowCount = 0
        Query_Listview(Me.SalesPoDetail, "EXEC SP_FRMBOXINPUTDET_LIST '" & Site_id & "','" & SalesPoList.SelectedItems.Item(0).Text & "','ALL'", True)
        Dim Qry As String = "Select LOT_NO,MODEL,(SELECT MODEL_NAME FROM TBL_MODELMASTER WHERE MODEL_NO = B.MODEL) AS MODELNM, "
        Qry += " C_NO, (SELECT CUS_NM FROM TBL_CUSTOMER WHERE CUS_NO = B.C_NO) as CUSNM ,isnull(L_CUST_NO,''), (SELECT CUS_NM FROM TBL_CUSTOMER WHERE CUS_NO =isnull(B.L_CUST_NO,'')) as LCUSTNM , "
        Qry += "S_PONO,O_QTY,0,  BOX_SIZE, (SELECT CODE_NAME FROM TBL_CODEMASTER WHERE SITE_ID='" & Site_id & "' AND CLASS_ID='r0036' AND CODE_ID =B.PACKING_TYPE) "
        Qry += " from tbl_lotmaster B where site_id = '" & Site_id & "' and lot_no = '" & SalesPoList.SelectedItems.Item(0).Text & "'"
        rs = Query_RS_ALL(Qry)
        If rs IsNot Nothing Then
            Do Until rs.EOF
                s.RowCount += 1
                s.Cells(s.RowCount - 1, 0).Text = rs(0).Value    'Lot No
                s.Cells(s.RowCount - 1, 1).Text = rs(1).Value   '모델코드
                s.Cells(s.RowCount - 1, 2).Text = rs(2).Value   '모델명
                s.Cells(s.RowCount - 1, 3).Text = rs(3).Value    '고객코드
                s.Cells(s.RowCount - 1, 4).Text = rs(4).Value     '고객명
                s.Cells(s.RowCount - 1, 5).Text = rs(5).Value    '매장코드
                s.Cells(s.RowCount - 1, 6).Text = rs(6).Value  '매장명
                s.Cells(s.RowCount - 1, 7).Text = rs(7).Value    '수주번호
                s.Cells(s.RowCount - 1, 8).Text = rs(8).Value    '요청수량
                s.Cells(s.RowCount - 1, 9).Text = SalesPoList.SelectedItems.Item(0).SubItems(2).Text   '적재수량 
                s.Cells(s.RowCount - 1, 10).Text = rs(10).Value    '박스크기
                s.Cells(s.RowCount - 1, 11).Text = rs(11).Value    '포장유형 
                's.Rows(s.RowCount - 1).ForeColor = Color.OrangeRed
                s.SetActiveCell(s.RowCount - 1, 0)
                rs.MoveNext()
            Loop
            s.SetActiveCell(s.RowCount - 1, 0)

            rs = Query_RS_ALL("select LOT_NO,MODEL,(SELECT MODEL_NAME FROM TBL_MODELMASTER WHERE MODEL_NO = B.MODEL) AS MODELNM,O_QTY,isnull(IN_QTY,0) from tbl_packingdet B where LOT_NO = '" & SalesPoList.SelectedItems.Item(0).Text & "'")
            If rs IsNot Nothing Then
                Do Until rs.EOF
                    s2.RowCount += 1
                    s2.Rows(s2.RowCount - 1).Height = 50
                    s2.Columns(0).Width = 100
                    s2.Columns(1).Width = 100
                    s2.Columns(2).Width = 50
                    s2.Columns(3).Width = 50
                    s2.Cells(s2.RowCount - 1, 0).Text = rs(1).Value     '모델코드
                    s2.Cells(s2.RowCount - 1, 1).Text = rs(3).Value    '요청수량
                    s2.Cells(s2.RowCount - 1, 2).Text = rs(4).Value   '적재수량
                    s2.Cells(s2.RowCount - 1, 3).Text = rs(2).Value   '모델명
                    s2.Cells(s2.RowCount - 1, 4).Text = rs(0).Value 'LOTNO 
                    s2.Cells(s2.RowCount - 1, 5).Text = rs(4).Value   '기적재수량
                    s2.Cells(s2.RowCount - 1, 2).ForeColor = Color.Red
                    rs.MoveNext()
                Loop
                's2.SetActiveCell(s.RowCount - 1, 0)
            End If

        End If

        Me.FpSpread2.Font = New System.Drawing.Font("Verdana", 18.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Spread_AutoCol(Me.FpSpread2)


    End Sub

    Private Sub TextBoxX1_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TextBoxX1.KeyUp
        Try
            If e.KeyCode <> Keys.Enter Then
                Exit Sub
            End If

            If TextBoxX1.Text <> "" Then
                Dim item1 As ListViewItem = SalesPoList.FindItemWithText(TextBoxX1.Text)
                If (item1 IsNot Nothing) Then
                    'MsgBox(item1.ToString)
                    item1.EnsureVisible()
                    item1.Selected = True
                    SalesPoList.Focus()
                End If
            End If

        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message, "ERROR")

        End Try
    End Sub


    Private Sub PackingBarcode_Print()

        Dim barcode_type = Query_RS("select reserv5 from tbl_customer where site_id='" & Site_id & "' and cus_no = '" & SalesPoList.SelectedItems(0).SubItems(3).Text & "'")
        Dim Barcodeval = ""

        Select Case barcode_type
            Case "barcode1"
                '한국 바코드 -----
                Barcodeval = Query_RS("select gtin12 from tbl_barcode where site_id='" & Site_id & "' and model = '" & SalesPoList.SelectedItems(0).SubItems(7).Text & "'")
                Print_BOx_Barcode1(SalesPoList.SelectedItems(0).SubItems(0).Text, SalesPoList.SelectedItems(0).SubItems(4).Text, SalesPoList.SelectedItems(0).SubItems(6).Text, SalesPoList.SelectedItems(0).SubItems(2).Text, rt)
                Barcodeval = ""
            Case "barcode2"
                For i = 0 To SalesPoDetail.Items.Count - 1
                    '미국 바코드 ----
                    Barcodeval = Query_RS("select gtin12 from tbl_barcode where site_id='" & Site_id & "' and model = '" & SalesPoDetail.Items(i).SubItems(1).Text & "'")
                    Print_BOx_Barcode2(SalesPoDetail.Items(i).SubItems(0).Text, Barcodeval, SalesPoDetail.Items(i).SubItems(1).Text, SalesPoDetail.Items(i).SubItems(2).Text, SalesPoDetail.Items(i).SubItems(4).Text, i + 1, rt)
                Next
            Case Else
                MessageBox.Show("[" & SalesPoList.SelectedItems(0).SubItems(3).Text & "]인쇄할 바코드타입이 지정되지 않았습니다.", "Validation Error")
                Exit Sub
        End Select








    End Sub

    Private Sub ButtonItem6_Click(sender As Object, e As EventArgs) Handles ButtonItem6.Click
        Query_Listview(Me.SalesPoList, "EXEC SP_FRMBOXINPUT_LIST '" & Site_id & "','ALL'", True)
    End Sub


    Private Sub bDel_Click(sender As Object, e As EventArgs)
        Dim S2 = Me.FpSpread2.ActiveSheet
        Dim S1 = Me.FpSpread1.ActiveSheet

        If S2.RowCount = 0 Then
            MessageBox.Show("삭제할 제품이 없습니다..", "Validation Error")
            Exit Sub
        End If
        If S2.Cells(S2.ActiveRowIndex, 5).Text = "U" Or S2.Cells(S2.ActiveRowIndex, 5).Text = "S" Then
            Dim inQty As Integer = CInt(Query_RS("SELECT isnull(IN_QTY,0) as inqty FROM TBL_PACKINGDET WHERE site_id = '" & Site_id & "' and lot_no = '" & S2.Cells(S2.ActiveRowIndex, 0).Text & "' and model = '" & S2.Cells(S2.ActiveRowIndex, 1).Text & "'"))

            If inQty > 0 Then
                MessageBox.Show("박스내 제품이 " & inQty & "개 적재되었으므로 삭제할수 없습니다. .", "Validation Error")
                Exit Sub
            End If

            If Insert_Data("delete tbl_packingdet where site_id = '" & Site_id & "' and lot_no = '" & S2.Cells(S2.ActiveRowIndex, 0).Text & "' and model = '" & S2.Cells(S2.ActiveRowIndex, 1).Text & "'") = False Then
                MessageBox.Show("제품삭제 실패")
                Exit Sub
            End If
        End If
        S1.Cells(S1.ActiveRowIndex, 6).Text = CInt(S1.Cells(S1.ActiveRowIndex, 6).Text) - CInt(S2.Cells(S2.ActiveRowIndex, 3).Text)
        S1.Cells(S1.ActiveRowIndex, 7).Text = CDbl(S1.Cells(S1.ActiveRowIndex, 7).Text) - (CInt(S2.Cells(S2.ActiveRowIndex, 3).Text) * CDbl(S2.Cells(S2.ActiveRowIndex, 4).Text))

        S2.Rows(S2.ActiveRowIndex).Remove()

        S1.Rows(S1.ActiveRowIndex).ForeColor = Color.OrangeRed
    End Sub

    Private Sub TextBoxX2_KeyDown(sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TextBoxX2.KeyDown
        Dim S1 = Me.FpSpread1.ActiveSheet
        Dim S2 = Me.FpSpread2.ActiveSheet
        Dim i As Integer
        Dim flag As Boolean = False
        Dim rs As ADODB.Recordset
        Try

            If e.KeyCode <> Keys.Enter Then
                Exit Sub
            End If

            If TextBoxX2.Text <> "" Then
                '스캔하는 고객바코드를 이용히여 모델을 찾는다.

                rs = Query_RS_ALL("select model, (select part_name from tbl_partmaster where part_no = b.model) as modelnm from tbl_barcode b where site_id = '" & Site_id & "' and (GTIN14 = '" & TextBoxX2.Text & "' or GTIN12 = '" & TextBoxX2.Text & "' or AMAZON = '" & TextBoxX2.Text & "')")

                If rs Is Nothing Then
                    MessageBox.Show("등록되지 않은 고객바코드입니다.. 다시 확인하세요!!!", "Validation Error")
                    TextBoxX2.Text = ""
                    Exit Sub
                End If

                Dim modelNo As String = rs(0).Value
                Dim modelNM As String = rs(1).Value

                For i = 0 To Me.SalesPoDetail.Items.Count - 1
                    If Me.SalesPoDetail.Items(i).SubItems(1).Text = modelNo Then
                        flag = True
                        Exit For
                    End If
                Next

                If flag = False Then
                    MessageBox.Show("포장대상 제품이 아닙니다. 다시 확인하세요!!!", "Validation Error")
                    TextBoxX2.Text = ""
                    Exit Sub
                End If


                Dim inQty As Integer = 0

                For i = 0 To S2.RowCount - 1
                    inQty = CInt(S2.Cells(i, 1).Text) - CInt(S2.Cells(i, 2).Text)
                    'prodmaster_b테이블에 공정이 W2000 인 것이 적재수량만큼 재고가 있는지 체크함
                    Dim prodQty As Integer = CInt(Query_RS("select count(*) from tbl_prodmaster_b where site_id = '" & Site_id & "' and model = '" & S2.Cells(i, 0).Text & "' and c_prc = 'W2000'"))


                    If prodQty < inQty Then
                        MessageBox.Show("[" & S2.Cells(i, 0).Text & "]개별포장공정에 있는 수량이 포장할 수량보다 적습니다.!!!", "Validation Error")
                        TextBoxX2.Text = ""
                        Exit Sub
                    End If
                Next


                Dim cusNm As String = Query_RS("select cus_nm from tbl_customer where site_id = '" & Site_id & "' and cus_no = '" & S1.Cells(S1.ActiveRowIndex, 3).Text & "'")

                CusBarCode.Text = TextBoxX2.Text
                CusName.Text = cusNm
                ModelName.Text = modelNM

                For i = 0 To S2.RowCount - 1
                    If S2.Cells(i, 0).Text = modelNo Then
                        If (CInt(S2.Cells(i, 2).Text) + 1) > CInt(S2.Cells(i, 1).Text) Then
                            Modal_Error("이미 적재완료되었습니다.. 다시 확인하세요!!!")
                            Exit For
                        End If
                        S2.Cells(i, 2).Text = CInt(S2.Cells(i, 2).Text) + 1
                        S1.Cells(0, 9).Text = CInt(S1.Cells(0, 9).Text) + 1
                        'S2.Rows(i).ForeColor = Color.OrangeRed
                        Exit For
                    End If

                Next


                TextBoxX2.Text = ""



            End If
        Catch ex As Exception
            TextBoxX2.Text = ""
            MessageBox.Show("Error: " & ex.Message, "ERROR")
        End Try
    End Sub

    Private Sub ButtonItem1_Click_1(sender As Object, e As EventArgs) Handles ButtonItem1.Click
        Dim S1 = FpSpread1.ActiveSheet
        PackingBarcode_Print()
    End Sub

    Private Sub DelBtn_Click(sender As Object, e As EventArgs) Handles DelBtn.Click

        Dim S1 = Me.FpSpread1.ActiveSheet
        Dim S2 = Me.FpSpread2.ActiveSheet

        Dim r = MessageBox.Show("박스포장내역을 초기화 하시겠습니까?" & vbCrLf & "적재제품수량을 0으로 초기화 합니다.", "박스초기화", MessageBoxButtons.OKCancel)

        If r = Windows.Forms.DialogResult.OK Then

            For i = 0 To S2.RowCount - 1
                S2.Cells(i, 2).Text = 0
                S2.Cells(i, 5).Text = 0
            Next
            S1.Cells(0, 9).Text = 0

            Insert_Data("update tbl_packingdet set in_qty = 0,u_person='" & Emp_No & "',u_date = getdate() where site_id = '" & Site_id & "' and LOT_NO = '" & S1.Cells(0, 0).Text & "'")
            Insert_Data("update tbl_prodmaster set c_prc = 'W2000',u_person='" & Emp_No & "',u_date = getdate() where site_id = '" & Site_id & "' and LOT_NO = '" & S1.Cells(0, 0).Text & "'")
            Query_Listview(Me.SalesPoList, "EXEC SP_FrmBoxInput_LIST'" & Site_id & "','ALL'", True)
            Query_Listview(Me.SalesPoDetail, "EXEC SP_FRMBOXINPUTDET_LIST '" & Site_id & "','" & S1.Cells(0, 0).Text & "','ALL'", True)
        End If
    End Sub

End Class