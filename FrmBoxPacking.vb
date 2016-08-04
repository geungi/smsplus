Imports FarPoint.Win.Spread


Public Class FrmBoxPacking
    Const CtrlMask As Byte = 8
    Private NowPoNo As String = ""
    Private SelPoNo As String = ""
    Private TotRec As Integer = 0
    Private TotRec2 As Integer = 0
    Private PoRec As Integer = 0
    Private chkModel As Boolean = False
    Private cusNM As String = ""


    Private Sub FrmBoxPacking_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

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

        Me.ContextMenuBar1.SetContextMenuEx(Me.FpSpread2, CtxSp)
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

        Me.DelBtn.Visible = False
        Query_Listview(Me.SalesPoList, "EXEC SP_FRMSALESPOMST_LIST '" & Site_id & "','ALL','" & Me.POStDate.Text & "','" & Me.POEdDate.Text & "','" & cusNM & "'", True)

    End Sub

    Private Sub Condi_Disp() 'CONTROL PANEL 및 파트리스트뷰에 있는 콤보박스에 데이터 출력


        Me.POStDate.Text = "1900-01-01"
        Me.POEdDate.Text = Now.Date.ToString()

        '수주처 1차
        Query_Combo(Me.CusCb1, "select cus_nm from tbl_customer where site_id = '" & Site_id & "' and cus_type = '1차'  ORDER BY cus_nm")
        Me.CusCb1.Items.Add("ALL")
        Me.CusCb1.Text = "ALL"

        If Me.CusCb1.Text = "ALL" Then
            cusNM = ""
        Else
            cusNM = Me.CusCb1.Text
        End If

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



    Private Sub NewBtn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NewBtn1.Click, NewBtn.Click, bNew.Click
        Try
            Dim S = Me.FpSpread1.ActiveSheet
            Dim S2 = Me.FpSpread2.ActiveSheet
            Dim i As Integer
            Dim Qry As String
            Dim Boxsize As String = ""
            Dim PkQty As Integer = 0
            Dim Auto_Lot As Boolean = False
            Dim PackType As String = ""

            NowPoNo = ""

            If NowPoNo = "" Then

                If Me.SalesPoList.SelectedItems.Count < 1 Or Me.SalesPoDetail.SelectedItems.Count < 1 Then
                    MessageBox.Show("수주번호 선택 후, 포장할 품목을 선택하십시오!!!", "Validation Error")
                    Exit Sub
                End If

                'If chkModel = True Then
                'MessageBox.Show("매장이 같은 경우에만 동일박스포장이 가능합니다.!!!", "Validation Error")
                'Exit Sub
                'End If

                Dim rslt As Integer = MessageBox.Show("大박스로 하시겠습니까? " & vbNewLine & "아니오를 선택하시면 小박스입니다.", "포장박스 선택", MessageBoxButtons.YesNo)

                If rslt = DialogResult.Yes Then
                    Boxsize = "대"
                Else
                    Boxsize = "소"
                End If

                PackType = Query_RS("SELECT CODE_ID FROM TBL_CODEMASTER WHERE SITE_ID='" & Site_id & "' AND CLASS_ID='R0036' AND CODE_NAME = '" & Me.SalesPoDetail.SelectedItems(0).SubItems(6).Text & "'")

                Qry = "SELECT P_QTY FROM TBL_BOXCNT WHERE site_id = '" & Site_id & "' and MODEL = '" & Me.SalesPoDetail.SelectedItems(0).SubItems(3).Text & "' and BOX_SIZE = '" & Boxsize & "' and PACKING_TYPE =  '" & PackType & "' "
                PkQty = CInt(Query_RS(Qry))

                If CInt(Me.SalesPoDetail.SelectedItems(0).SubItems(5).Text) > PkQty Then
                    Dim rslt2 As Integer = MessageBox.Show("제품수량이 포장가능수량보다 많습니다." & vbNewLine & "자동으로 박스를 추가하시겠습니까?.", "박스 자동추가", MessageBoxButtons.YesNo)
                    If rslt = DialogResult.Yes Then
                        Auto_Lot = True
                    Else
                        Auto_Lot = False
                    End If
                End If

                Dim totQty = CInt(Me.SalesPoDetail.SelectedItems(0).SubItems(5).Text)

                If Auto_Lot = False Then '자동박스구성이 아닌경우

                    Create_PO()

                    S2.RowCount += 1
                    S2.Cells(S2.RowCount - 1, 0).Text = NowPoNo    'Lot No
                    S2.Cells(S2.RowCount - 1, 1).Text = Me.SalesPoDetail.SelectedItems(0).SubItems(3).Text    '모델코드
                    S2.Cells(S2.RowCount - 1, 2).Text = Me.SalesPoDetail.SelectedItems(0).SubItems(4).Text    '모델명
                    S2.Cells(S2.RowCount - 1, 3).Text = Me.SalesPoDetail.SelectedItems(0).SubItems(5).Text    '제품수량
                    S2.Cells(S2.RowCount - 1, 4).Text = Me.SalesPoDetail.SelectedItems(0).SubItems(8).Text    '단가
                    S2.Cells(S2.RowCount - 1, 5).Text = "S"  '저장모드   추가 I / 변경 U
                    S2.Cells(S2.RowCount - 1, 3).Locked = False
                    S2.Rows(S2.RowCount - 1).ForeColor = Color.Black
                    S2.SetActiveCell(S2.RowCount - 1, 0)



                    '포장할 품목정보를 저장함 - 우선 저장하고 수량 변경시 업데이트 
                    Qry = "insert tbl_packingdet(site_id,LOT_NO,MODEL,O_QTY,price,c_person,c_date,u_person,u_date) "
                    Qry += "values('" & Site_id & "','" & S2.Cells(S2.RowCount - 1, 0).Text & "','" & S2.Cells(S2.RowCount - 1, 1).Text & "', "
                    Qry += S2.Cells(S2.RowCount - 1, 3).Text & "," & S2.Cells(S2.RowCount - 1, 4).Text & ", '" & Emp_No & "', getdate() ,'" & Emp_No & "', getdate() "
                    Qry += ")"

                    Insert_Data(Qry)

                    S.RowCount += 1
                    S.Cells(S.RowCount - 1, 0).Text = NowPoNo    'Lot No
                    S.Cells(S.RowCount - 1, 1).Text = Me.SalesPoDetail.SelectedItems(0).SubItems(0).Text    '수주번호
                    S.Cells(S.RowCount - 1, 2).Text = Me.SalesPoDetail.SelectedItems(0).SubItems(3).Text   '모델코드
                    S.Cells(S.RowCount - 1, 3).Text = Me.SalesPoDetail.SelectedItems(0).SubItems(4).Text    '모델명
                    S.Cells(S.RowCount - 1, 4).Text = Boxsize     '박스크기
                    S.Cells(S.RowCount - 1, 5).Text = Me.SalesPoDetail.SelectedItems(0).SubItems(6).Text    '포장유형
                    S.Cells(S.RowCount - 1, 6).Text = totQty  '포장수량
                    S.Cells(S.RowCount - 1, 7).Text = CDbl(Me.SalesPoDetail.SelectedItems(0).SubItems(8).Text) * totQty    '총금액
                    S.Cells(S.RowCount - 1, 8).Text = Me.SalesPoList.SelectedItems(0).SubItems(2).Text    '고객코드
                    S.Cells(S.RowCount - 1, 9).Text = Me.SalesPoList.SelectedItems(0).SubItems(3).Text    '고객사 
                    S.Cells(S.RowCount - 1, 10).Text = Me.SalesPoDetail.SelectedItems(0).SubItems(1).Text    '매장코드
                    S.Cells(S.RowCount - 1, 11).Text = Me.SalesPoDetail.SelectedItems(0).SubItems(2).Text    '매장명 
                    S.Cells(S.RowCount - 1, 12).Text = Me.SalesPoDetail.SelectedItems(0).SubItems(7).Text    '판매그룹 
                    'S.Rows(S.RowCount - 1).ForeColor = Color.OrangeRed
                    S.SetActiveCell(S.RowCount - 1, 0)

                    PackType = Query_RS("SELECT CODE_ID FROM TBL_CODEMASTER WHERE SITE_ID='" & Site_id & "' AND CLASS_ID='R0036' AND CODE_NAME = '" & S.Cells(S.RowCount - 1, 5).Text & "'")


                    '포장할 박스정보를 저장함 - PO번호의 시쿼스를 유지하기 위함 
                    Qry = "insert tbl_lotmaster(site_id,LOT_NO,MODEL,O_QTY,P_QTY,S_PONO,C_NO,L_CUST_NO,SALES_GROUP,PACKING_TYPE,BOX_SIZE,U_PRICE,c_person,c_date,u_person,u_date,status) "
                    Qry += "values('" & Site_id & "','" & S.Cells(S.RowCount - 1, 0).Text & "','" & S.Cells(S.RowCount - 1, 2).Text & "', "
                    Qry += S.Cells(S.RowCount - 1, 6).Text & ", 0 , '" & S.Cells(S.RowCount - 1, 1).Text & "' ,'" & S.Cells(S.RowCount - 1, 8).Text & "','" & S.Cells(S.RowCount - 1, 10).Text & "','" & S.Cells(S.RowCount - 1, 12).Text & "', "
                    Qry += "'" & PackType & "', '" & S.Cells(S.RowCount - 1, 4).Text & "', '" & S.Cells(S.RowCount - 1, 7).Text & "', '" & Emp_No & "', getdate() ,'" & Emp_No & "', getdate(),'OPENED'"
                    Qry += ")"

                    Insert_Data(Qry)

                Else
                    Dim boxCnt As Integer
                    Dim mQty As Integer = totQty
                    Dim inQty As Integer = 0
                    boxCnt = Math.Ceiling(totQty / PkQty)
                    For i = 0 To boxCnt - 1
                        If mQty >= PkQty Then
                            inQty = PkQty
                        Else
                            inQty = mQty
                        End If

                        Create_PO()

                        S2.RowCount += 1
                        S2.Cells(S2.RowCount - 1, 0).Text = NowPoNo    'Lot No
                        S2.Cells(S2.RowCount - 1, 1).Text = Me.SalesPoDetail.SelectedItems(0).SubItems(3).Text    '모델코드
                        S2.Cells(S2.RowCount - 1, 2).Text = Me.SalesPoDetail.SelectedItems(0).SubItems(4).Text    '모델명
                        S2.Cells(S2.RowCount - 1, 3).Text = inQty    '제품수량
                        S2.Cells(S2.RowCount - 1, 4).Text = Me.SalesPoDetail.SelectedItems(0).SubItems(8).Text    '단가
                        S2.Cells(S2.RowCount - 1, 5).Text = "S"  '저장모드   추가 I / 변경 U
                        S2.Cells(S2.RowCount - 1, 3).Locked = False
                        S2.Rows(S2.RowCount - 1).ForeColor = Color.Black
                        S2.SetActiveCell(S2.RowCount - 1, 0)



                        '포장할 품목정보를 저장함 - 우선 저장하고 수량 변경시 업데이트 
                        Qry = "insert tbl_packingdet(site_id,LOT_NO,MODEL,O_QTY,price,c_person,c_date,u_person,u_date) "
                        Qry += "values('" & Site_id & "','" & S2.Cells(S2.RowCount - 1, 0).Text & "','" & S2.Cells(S2.RowCount - 1, 1).Text & "', "
                        Qry += S2.Cells(S2.RowCount - 1, 3).Text & "," & S2.Cells(S2.RowCount - 1, 4).Text & ", '" & Emp_No & "', getdate() ,'" & Emp_No & "', getdate() "
                        Qry += ")"

                        Insert_Data(Qry)

                        S.RowCount += 1
                        S.Cells(S.RowCount - 1, 0).Text = NowPoNo    'Lot No
                        S.Cells(S.RowCount - 1, 1).Text = Me.SalesPoDetail.SelectedItems(0).SubItems(0).Text    '수주번호
                        S.Cells(S.RowCount - 1, 2).Text = Me.SalesPoDetail.SelectedItems(0).SubItems(3).Text   '모델코드
                        S.Cells(S.RowCount - 1, 3).Text = Me.SalesPoDetail.SelectedItems(0).SubItems(4).Text    '모델명
                        S.Cells(S.RowCount - 1, 4).Text = Boxsize     '박스크기
                        S.Cells(S.RowCount - 1, 5).Text = Me.SalesPoDetail.SelectedItems(0).SubItems(6).Text    '포장유형
                        S.Cells(S.RowCount - 1, 6).Text = inQty  '포장수량
                        S.Cells(S.RowCount - 1, 7).Text = CInt(Me.SalesPoDetail.SelectedItems(0).SubItems(8).Text) * inQty    '총금액
                        S.Cells(S.RowCount - 1, 8).Text = Me.SalesPoList.SelectedItems(0).SubItems(2).Text    '고객코드
                        S.Cells(S.RowCount - 1, 9).Text = Me.SalesPoList.SelectedItems(0).SubItems(3).Text    '고객사 
                        S.Cells(S.RowCount - 1, 10).Text = Me.SalesPoDetail.SelectedItems(0).SubItems(1).Text    '매장코드
                        S.Cells(S.RowCount - 1, 11).Text = Me.SalesPoDetail.SelectedItems(0).SubItems(2).Text    '매장명 
                        S.Cells(S.RowCount - 1, 12).Text = Me.SalesPoDetail.SelectedItems(0).SubItems(7).Text    '판매그룹 
                        'S.Rows(S.RowCount - 1).ForeColor = Color.OrangeRed
                        S.SetActiveCell(S.RowCount - 1, 0)

                        PackType = Query_RS("SELECT CODE_ID FROM TBL_CODEMASTER WHERE SITE_ID='" & Site_id & "' AND CLASS_ID='R0036' AND CODE_NAME = '" & S.Cells(S.RowCount - 1, 5).Text & "'")

                        '포장할 박스정보를 저장함 - PO번호의 시쿼스를 유지하기 위함 
                        Qry = "insert tbl_lotmaster(site_id,LOT_NO,MODEL,O_QTY,P_QTY,S_PONO,C_NO,L_CUST_NO,SALES_GROUP,PACKING_TYPE,BOX_SIZE,U_PRICE,c_person,c_date,u_person,u_date,status) "
                        Qry += "values('" & Site_id & "','" & S.Cells(S.RowCount - 1, 0).Text & "','" & S.Cells(S.RowCount - 1, 2).Text & "', "
                        Qry += S.Cells(S.RowCount - 1, 6).Text & ", 0 , '" & S.Cells(S.RowCount - 1, 1).Text & "' ,'" & S.Cells(S.RowCount - 1, 8).Text & "','" & S.Cells(S.RowCount - 1, 10).Text & "','" & S.Cells(S.RowCount - 1, 12).Text & "', "
                        Qry += "'" & PackType & "', '" & S.Cells(S.RowCount - 1, 4).Text & "', '" & S.Cells(S.RowCount - 1, 7).Text & "', '" & Emp_No & "', getdate() ,'" & Emp_No & "', getdate(),'OPENED'"
                        Qry += ")"

                        Insert_Data(Qry)
                        mQty = mQty - PkQty
                        If i < boxCnt - 1 Then
                            S2.RowCount = 0
                        End If

                    Next

                End If

                Spread_AutoCol(FpSpread1)
                Spread_AutoCol(FpSpread2)

            End If
        Catch ex As Exception
            MessageBox.Show("Error " & ex.Message, "ERROR")
        End Try
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
            Query_Listview(Me.SalesPoList, "EXEC SP_FRMSALESPOMST_LIST '" & Site_id & "','ALL','" & Me.POStDate.Text & "','" & Me.POEdDate.Text & "','" & cusNM & "'", True)
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


            Dim S2 = Me.FpSpread2.ActiveSheet
            Dim P = Me.FpSpread1.ActiveSheet
            Dim msg As String = ""
            Dim Qry As String = ""

            'If S.ActiveRowIndex = (S.RowCount - 1) Then   '마지막 변경셀에서 다른 셀로 포커스 이동하지 않아도 저장되도록 자동 포커스 이동
            'S.SetActiveCell(S.ActiveRowIndex - 1, S.ActiveColumnIndex)
            ' Else
            'S.SetActiveCell(S.ActiveRowIndex + 1, S.ActiveColumnIndex)
            'End If

            'S.SetActiveCell(0, 0)

            If S2.RowCount = 0 Then

                MessageBox.Show("포장할  제품이 등록되지 않았습니다..!!!", "Validation Error")
                Exit Sub

            Else

                For i = 0 To S2.RowCount - 1
                    If S2.Rows(i).ForeColor = Color.OrangeRed Then  '추가 또는 변경 건만 업데이트
                        If S2.Cells(i, 5).Text = "I" Then '신규추가
                            Qry = "insert tbl_packingdet(site_id,LOT_NO,MODEL,O_QTY,price,c_person,c_date,u_person,u_date) "
                            Qry += "values('" & Site_id & "','" & S2.Cells(i, 0).Text & "','" & S2.Cells(i, 1).Text & "', "
                            Qry += S2.Cells(i, 3).Text & "," & S2.Cells(i, 4).Text & ", '" & Emp_No & "', getdate() ,'" & Emp_No & "', getdate() "
                            Qry += ")"
                        Else
                            Qry = "update tbl_packingdet set o_qty = " & S2.Cells(i, 3).Text & ",price = '" & S2.Cells(i, 4).Text & "',u_person='" & Emp_No & "',u_date = getdate() where site_id = '" & Site_id & "' and LOT_NO = '" & S2.Cells(i, 0).Text & "' and model = '" & S2.Cells(i, 1).Text & "'"
                        End If
                        If Insert_Data(Qry) = False Then
                            MessageBox.Show("제품등록(변경) 실패")
                            Exit Sub
                        Else
                            S2.Cells(i, 5).Text = "S"
                            S2.Rows(i).ForeColor = Color.Black
                        End If
                    End If
                Next

                If Insert_Data("update tbl_LOTMASTER Set o_qty = " & P.Cells(P.ActiveRowIndex, 6).Text & ", u_price = " & P.Cells(P.ActiveRowIndex, 7).Text & ", u_person='" & Emp_No & "',u_date = getdate() where site_id = '" & Site_id & "' and LOT_NO = '" & P.Cells(P.ActiveRowIndex, 0).Text & "'") = False Then
                    MessageBox.Show("박스구성 실패")
                    Exit Sub
                Else
                    P.Rows(P.ActiveRowIndex).ForeColor = Color.Black
                End If



                NowPoNo = ""



            End If

            ' FindBtn_Click(Me.FpSpread1, EventArgs.Empty)
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

    Private Sub bClear_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bClear.Click
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


            NowPoNo = Query_RS("exec SP_FRMBOXPACKING_LOTNO '" & Site_id & "','" & podate & "'")

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
        s.RowCount = 0
        s2.RowCount = 0
        Query_Listview(Me.SalesPoDetail, "EXEC SP_FRMSALESPODET_LIST '" & Site_id & "','" & SalesPoList.SelectedItems.Item(0).Text & "','ALL'", True)
        Dim Qry As String = "Select LOT_NO,S_PONO,MODEL,(SELECT MODEL_NAME FROM TBL_MODELMASTER WHERE MODEL_NO = B.MODEL) AS MODELNM, "
        Qry += " BOX_SIZE, (SELECT CODE_NAME FROM TBL_CODEMASTER WHERE SITE_ID='" & Site_id & "' AND CLASS_ID='r0036' AND CODE_ID =B.PACKING_TYPE), O_QTY, isnull(U_PRICE,0.00), C_NO, (SELECT CUS_NM FROM TBL_CUSTOMER WHERE CUS_NO = B.C_NO) as CUSNM , "
        Qry += "isnull(L_CUST_NO,''), (SELECT CUS_NM FROM TBL_CUSTOMER WHERE CUS_NO =isnull(B.L_CUST_NO,'')) as LCUSTNM , SALES_GROUP "
        Qry += " from tbl_lotmaster B where site_id = '" & Site_id & "' and S_PONO = '" & SalesPoList.SelectedItems.Item(0).Text & "'"
        rs = Query_RS_ALL(Qry)
        If rs IsNot Nothing Then
            Do Until rs.EOF
                s.RowCount += 1
                s.Cells(s.RowCount - 1, 0).Text = rs(0).Value    'Lot No
                s.Cells(s.RowCount - 1, 1).Text = rs(1).Value   '수주번호
                s.Cells(s.RowCount - 1, 2).Text = rs(2).Value   '모델코드
                s.Cells(s.RowCount - 1, 3).Text = rs(3).Value    '모델명
                s.Cells(s.RowCount - 1, 4).Text = rs(4).Value     '박스크기
                s.Cells(s.RowCount - 1, 5).Text = rs(5).Value    '포장유형
                s.Cells(s.RowCount - 1, 6).Text = rs(6).Value  '포장수량
                s.Cells(s.RowCount - 1, 7).Text = rs(7).Value    '총금액
                s.Cells(s.RowCount - 1, 8).Text = rs(8).Value    '고객코드
                s.Cells(s.RowCount - 1, 9).Text = rs(9).Value    '고객사 
                s.Cells(s.RowCount - 1, 10).Text = rs(10).Value    '매장코드
                s.Cells(s.RowCount - 1, 11).Text = rs(11).Value    '매장명 
                s.Cells(s.RowCount - 1, 12).Text = rs(12).Value    '판매그룹 
                's.Rows(s.RowCount - 1).ForeColor = Color.OrangeRed
                s.SetActiveCell(s.RowCount - 1, 0)
                rs.MoveNext()
            Loop
            s.SetActiveCell(s.RowCount - 1, 0)

            Qry = "Select LOT_NO,MODEL,(SELECT MODEL_NAME FROM TBL_MODELMASTER WHERE MODEL_NO = B.MODEL) AS MODELNM, O_QTY, price  "
            Qry += " from tbl_packingdet B where site_id = '" & Site_id & "' and lot_no = '" & s.Cells(s.ActiveRowIndex, 0).Text & "'"
            rs = Query_RS_ALL(Qry)
            If rs IsNot Nothing Then
                Do Until rs.EOF
                    s2.RowCount += 1
                    s2.Cells(s2.RowCount - 1, 0).Text = rs(0).Value    'Lot No
                    s2.Cells(s2.RowCount - 1, 1).Text = rs(1).Value    '모델코드
                    s2.Cells(s2.RowCount - 1, 2).Text = rs(2).Value   '모델명
                    s2.Cells(s2.RowCount - 1, 3).Text = rs(3).Value    '제품수량
                    s2.Cells(s2.RowCount - 1, 4).Text = rs(4).Value    '단가
                    s2.Cells(s2.RowCount - 1, 5).Text = "S"  '저장모드   추가 I / 변경 U
                    s2.Cells(s2.RowCount - 1, 3).Locked = False
                    's2.Rows(s2.RowCount - 1).ForeColor = Color.OrangeRed
                    s2.SetActiveCell(s2.RowCount - 1, 0)
                    rs.MoveNext()
                    Spread_AutoCol(Me.FpSpread2)
                Loop
            End If

        End If

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

    Private Sub SerialCode_TextChanged2(sender As Object, e As EventArgs)

        Dim S1 = Me.FpSpread1.ActiveSheet
        Dim S2 = Me.FpSpread2.ActiveSheet
        Dim i As Integer
        Dim flag As Boolean = False
        Try
            If TextBoxX1.Text <> "" Then
                '스캔하는 제품이 포장대상이 아닌지 체크하기 위해 모델명을 확인하기 위한 조회이나, 혼입인 경우에는 체크가 어려운 부분이 있음
                Dim modelNo As String = Query_RS("select model from tbl_prodmaster where site_id = '" & Site_id & "' and prod_no = '" & TextBoxX1.Text & "'")

                For i = 0 To Me.SalesPoDetail.SelectedItems.Count - 1
                    If Me.SalesPoDetail.SelectedItems(i).SubItems(1).Text = modelNo Then
                        flag = True
                        Exit For
                    End If
                Next

                If flag = False Then
                    MessageBox.Show("포장대상 제품이 아닙니다. 다시 확인하세요!!!", "Validation Error")
                    Exit Sub
                End If


                S2.RowCount += 1
                S2.Cells(S2.RowCount - 1, 0).Text = TextBoxX1.Text
                S2.Cells(S2.RowCount - 1, 1).Text = modelNo
                S2.Cells(S2.RowCount - 1, 2).Text = S1.Cells(S1.ActiveRowIndex, 0).Text
                S2.Cells(S2.RowCount - 1, 3).Text = S1.Cells(S1.ActiveRowIndex, 4).Text
                S2.Rows(S2.RowCount - 1).ForeColor = Color.OrangeRed
                CustomBarcode_Print(S2.Cells(S2.RowCount - 1, 3).Text) '고객바코드 출력


                TextBoxX1.Text = ""



            End If
        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message, "ERROR")
        End Try
    End Sub

    Private Sub CustomBarcode_Print(c_no As String)

        MessageBox.Show("고객바코드 출력:" & c_no)

    End Sub

    Private Sub PackingBarcode_Print(lot_no As String)

        MessageBox.Show("패킹바코드 출력:" & lot_no)

    End Sub

    Private Sub ButtonItem6_Click(sender As Object, e As EventArgs) Handles ButtonItem6.Click
        Query_Listview(Me.SalesPoList, "EXEC SP_FRMSALESPOMST_LIST '" & Site_id & "','ALL','" & Me.POStDate.Text & "','" & Me.POEdDate.Text & "','" & cusNM & "'", True)
    End Sub

    Private Sub SalesPoDetail_DoubleClick(sender As Object, e As EventArgs) Handles SalesPoDetail.DoubleClick
        Dim S2 = Me.FpSpread2.ActiveSheet
        Dim S1 = Me.FpSpread1.ActiveSheet

        If S2.RowCount > 0 Then
            If SalesPoDetail.SelectedItems(0).SubItems(1).Text <> S1.Cells(S1.ActiveRowIndex, 10).Text Then
                MessageBox.Show("배송매장이 동일한 제품만 추가할수 있습니다.", "Validation Error")
                Exit Sub
            End If
            For i = 0 To S2.RowCount - 1
                If SalesPoDetail.SelectedItems(0).SubItems(3).Text = S2.Cells(i, 1).Text Then
                    MessageBox.Show("이미 추가된 제품입니다.", "Validation Error")
                    Exit Sub
                End If
            Next
        End If

        If S1.RowCount > 0 And S2.RowCount = 0 Then
            S1.Cells(S1.ActiveRowIndex, 2).Text = Me.SalesPoDetail.SelectedItems(0).SubItems(3).Text
            S1.Cells(S1.ActiveRowIndex, 3).Text = Me.SalesPoDetail.SelectedItems(0).SubItems(4).Text
            S1.Cells(S1.ActiveRowIndex, 6).Text = "0"
            S1.Cells(S1.ActiveRowIndex, 7).Text = "0.00"
        End If

        S2.RowCount += 1
            S2.Cells(S2.RowCount - 1, 0).Text = S1.Cells(S1.ActiveRowIndex, 0).Text     'Lot No
            S2.Cells(S2.RowCount - 1, 1).Text = Me.SalesPoDetail.SelectedItems(0).SubItems(3).Text    '모델코드
            S2.Cells(S2.RowCount - 1, 2).Text = Me.SalesPoDetail.SelectedItems(0).SubItems(4).Text    '모델명
            S2.Cells(S2.RowCount - 1, 3).Text = Me.SalesPoDetail.SelectedItems(0).SubItems(5).Text    '제품수량
            S2.Cells(S2.RowCount - 1, 4).Text = Me.SalesPoDetail.SelectedItems(0).SubItems(8).Text    '단가
            S2.Cells(S2.RowCount - 1, 5).Text = "I"  '저장모드   추가 I / 변경 U
            S2.Cells(S2.RowCount - 1, 3).Locked = False
        S2.Rows(S2.RowCount - 1).ForeColor = Color.OrangeRed
        S1.Cells(S1.ActiveRowIndex, 6).Text = CInt(S1.Cells(S1.ActiveRowIndex, 6).Text) + CInt(S2.Cells(S2.RowCount - 1, 3).Text)
        S1.Cells(S1.ActiveRowIndex, 7).Text = CDbl(S1.Cells(S1.ActiveRowIndex, 7).Text) + (CInt(S2.Cells(S2.RowCount - 1, 3).Text) * CDbl(S2.Cells(S2.RowCount - 1, 4).Text))
        S2.SetActiveCell(S2.RowCount - 1, 0)

    End Sub


    Private Sub bDel_Click(sender As Object, e As EventArgs) Handles bDel.Click
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


End Class