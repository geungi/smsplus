Public Class FrmPOF

    Private Sub CheckedListBox1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles CheckedListBox1.LostFocus
        TextBoxDropDown1.Text = MODEL_SELECTED(CheckedListBox1)
    End Sub

    Private Sub FrmResult_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        DockContainerItem1.Text = "실행 메뉴"
        DockContainerItem2.Text = "조회 조건"

        DateTimeInput1.Value = Now
        DateTimeInput2.Value = Now

        curcell.Separator = ","
        curcell.ShowSeparator = True
        curcell.ShowCurrencySymbol = True
        curcell.CurrencySymbol = "$"

        'If Query_Combo(Me.ComboBoxEx1, "SELECT model_no FROM tbl_modelmaster WHERE site_id = '" & Site_id & "' and active = 'Y' ORDER BY model_no") = True Then
        'End If
        'Me.ComboBoxEx1.Items.Add("ALL")
        'Me.ComboBoxEx1.Text = "ALL"
        If Query_CheckList(Me.CheckedListBox1, "SELECT model_no FROM tbl_modelmaster WHERE site_id = '" & Site_id & "' and active = 'Y' ORDER BY model_no") = True Then
            Me.CheckedListBox1.Items.Add("ALL")
            Me.CheckedListBox1.Text = "ALL"
        End If

        If Spread_Setting(FpSpread1, Me.Name) = True Then
            FpSpread1.ActiveSheet.RowCount = 0
            FpSpread1.ActiveSheet.Columns(12).CellType = CHKcell
            Spread_AutoCol(FpSpread1)
        End If

        Formbim_Authority(Me.ButtonItem2, Me.Name, "NEW")
        Formbtn_Authority(Me.NewBtn1, Me.Name, "NEW")
        Formbim_Authority(Me.ButtonItem3, Me.Name, "SAVE")
        Formbtn_Authority(Me.SaveBtn1, Me.Name, "SAVE")
        Formbim_Authority(Me.ButtonItem4, Me.Name, "DELETE")
        Formbtn_Authority(Me.DelBtn1, Me.Name, "DELETE")
        Formbim_Authority(Me.ButtonItem5, Me.Name, "PRINT")
        Formbtn_Authority(Me.PrtBtn1, Me.Name, "PRINT")
        Formbim_Authority(Me.ButtonItem6, Me.Name, "EXCEL")
        Formbtn_Authority(Me.XlsBtn1, Me.Name, "EXCEL")
        Formbim_Authority(Me.FindBtn, Me.Name, "FIND")

        ButtonItem2.Enabled = False
        ButtonItem3.Enabled = False
        ButtonItem4.Enabled = False


        'If Query_RS("select isnull(insa_yn,'N') from tbl_empmaster where emp_no = '" & Emp_No & "'") = "N" Then
        '    FpSpread1.ActiveSheet.Columns(4).Visible = False
        '    FpSpread1.ActiveSheet.Columns(10).Visible = False
        'End If

    End Sub


    Private Sub FindBtn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FindBtn.Click

        Dim Qry As String
        Dim m_qry As String = ""
        m_qry = MODEL_SELECTED(CheckedListBox1)

        FpSpread1.ActiveSheet.RowCount = 0

        Qry = "set nocount on" & vbNewLine & vbNewLine

        'Qry = Qry & "select a.p_no, a.c_no,  b.part_name , a.qty, B.usgavg, b.price, " & vbNewLine
        'Qry = Qry & "			isnull((select sum(qty) from tbl_pgoal where p_date between '" & DateTimeInput1.Text & "' and '" & DateTimeInput2.Text & "' and model = a.p_no and div = 'P2000'),0) as planq," & vbNewLine
        'Qry = Qry & "			isnull((select sum(qty) from tbl_pgoal where p_date between '" & DateTimeInput1.Text & "' and '" & DateTimeInput2.Text & "' and model = a.p_no and div = 'P2000'),0)*a.qty + isnull((select sum(qty) from tbl_pgoal where p_date between '" & DateTimeInput1.Text & "' and '" & DateTimeInput2.Text & "' and model = a.p_no and div = 'P2000'),0)*B.usgavg as planq1," & vbNewLine
        'Qry = Qry & "			isnull((select sum(qty) from tbl_partinv where part_no = a.c_no and wh_cd in ('E1000','W1000','A1000')),0) as pinv," & vbNewLine

        'Qry = Qry & "			isnull((select sum(o_qty - in_qty)  from tbl_podetail where part_no = a.c_no and status = 'OPENED'),0) as norcv, 0,0" & vbNewLine
        'Qry = Qry & "from tbl_bom a, tbl_partmaster b" & vbNewLine
        'Qry = Qry & "where a.active = 'Y'" & vbNewLine
        ''If ComboBoxEx1.Text <> "ALL" Then
        ''    Qry = Qry & "AND p_no  = '" & ComboBoxEx1.Text & "'" & vbNewLine
        ''End If
        'If m_qry <> "" Then
        '    Qry = Qry & "AND p_no IN (" & m_qry & ")" & vbNewLine
        'End If
        'Qry = Qry & "and a.c_no = b.part_no " & vbNewLine
        'Qry = Qry & "order by p_no, c_no " & vbNewLine

        Qry = Qry & "select a.p_no, c_no,  (SELECT TOP 1 part_name FROM TBL_PARTMASTER WHERE part_no = b.part_no), " & vbNewLine
        Qry = Qry & "			(SELECT TOP 1 QTY FROM TBL_BOM WHERE P_NO = A.P_NO AND c_no = a.c_no), " & vbNewLine
        Qry = Qry & "			(SELECT TOP 1 usgavg FROM TBL_PARTMASTER WHERE part_no = b.part_no), " & vbNewLine
        Qry = Qry & "			(SELECT TOP 1 price FROM TBL_PARTMASTER WHERE part_no = b.part_no)," & vbNewLine
        Qry = Qry & "			isnull((select sum(qty) from tbl_pgoal where p_date between '" & DateTimeInput1.Text & "' and '" & DateTimeInput2.Text & "' and model = a.p_no and div = 'P2000'),0) as planq," & vbNewLine
        Qry = Qry & "			isnull((select sum(qty) from tbl_pgoal where p_date between '" & DateTimeInput1.Text & "' and '" & DateTimeInput2.Text & "' and model = a.p_no and div = 'P2000'),0)*(SELECT TOP 1 QTY FROM TBL_BOM WHERE P_NO = A.P_NO AND c_no = a.c_no) + " & vbNewLine
        Qry = Qry & "			isnull((select sum(qty) from tbl_pgoal where p_date between '" & DateTimeInput1.Text & "' and '" & DateTimeInput2.Text & "' and model = a.p_no and div = 'P2000'),0)*(SELECT TOP 1 usgavg FROM TBL_PARTMASTER WHERE part_no = b.part_no) as planq1," & vbNewLine
        Qry = Qry & "			isnull((select sum(qty) from tbl_partinv where part_no = a.c_no and wh_cd in ('E1000','W1000','A1000')),0) as pinv," & vbNewLine
        Qry = Qry & "			isnull((select sum(o_qty - in_qty)  from tbl_podetail where part_no = a.c_no and status = 'OPENED'),0) as norcv, 0,0,0" & vbNewLine
        Qry = Qry & "from tbl_bom a, tbl_partmaster b" & vbNewLine
        Qry = Qry & "where a.active = 'Y'" & vbNewLine

        If m_qry <> "" Then
            Qry = Qry & "AND p_no IN (" & m_qry & ")" & vbNewLine
        End If
        Qry = Qry & "and  c_no = part_no " & vbNewLine
        Qry = Qry & "GROUP BY P_NO, c_no, part_no " & vbNewLine
        Qry = Qry & "order by p_no, c_no" & vbNewLine



        If Query_Spread(FpSpread1, Qry, 1) = True Then

            With FpSpread1.ActiveSheet

                For I As Integer = 0 To .RowCount - 1
                    '                    .SetValue(I, 10, .GetValue(I, 7) - .GetValue(I, 8) - .GetValue(I, 9))
                    .SetValue(I, 10, .GetValue(I, 7)) '- .GetValue(I, 9))
                    If .GetValue(I, 10) < 0 Then
                        .SetValue(I, 10, 0)
                    End If
                    .SetValue(I, 11, .GetText(I, 10) * .GetText(I, 5))
                Next

                .Columns(12).Locked = False

            End With

            '        If Query_Spread(FpSpread1, Qry & " '" & Site_id & "','" & ComboBoxEx1.Text & "','" & DateTimeInput1.Text & "','" & DateTimeInput2.Text & "'", 1) = True Then
            'With FpSpread1.ActiveSheet
            '    .RowCount = .RowCount + 1
            '    .Cells(.RowCount - 1, 1, .RowCount - 1, .ColumnCount - 1).CellType = intcell
            '    .Cells(.RowCount - 1, .ColumnCount - 1).CellType = deccell
            '    .Rows(.RowCount - 1).BackColor = Color.Yellow
            '    .SetText(.RowCount - 1, 0, "TOTAL")

            '    .Cells(.RowCount - 1, 7, .RowCount - 1, 7).CellType = curcell
            '    '.Cells(.RowCount - 1, 9, .RowCount - 1, 9).CellType = curcell
            '    .Cells(.RowCount - 1, 10, .RowCount - 1, 10).CellType = curcell


            '    SPREAD_ROW_TOTAL(FpSpread1, 1, .ColumnCount - 1, 1)
            '    .Cells(.RowCount - 1, 7, .RowCount - 1, 7).Value = 0
            '    .Cells(.RowCount - 1, 10, .RowCount - 1, 10).Value = 0
            '    SPREAD_ROW_TOTAL_DEC(FpSpread1, 7, 7, 1)
            '    'SPREAD_ROW_TOTAL_DEC(FpSpread1, 9, 9, 1)
            '    SPREAD_ROW_TOTAL_DEC(FpSpread1, 10, 10, 1)

            'End With

            'curcell.CurrencySymbol = "$"
            'curcell.DecimalPlaces = 2
            'curcell.ShowCurrencySymbol = True
            'curcell.ShowSeparator = True

            Spread_AutoCol(FpSpread1)
            'FpSpread1.ActiveSheet.Columns(1, 10).Width = 80



        End If

        SaveBtn1.Enabled = True
        ButtonItem3.Enabled = True


        'End If

    End Sub


    Private Sub ButtonItem2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonItem2.Click

    End Sub

    Private Sub ButtonItem3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonItem3.Click


        If MessageBox.Show("자재 발주를 저장합니다. 계속 진행하시겠습니까?", "자재 발주", MessageBoxButtons.YesNo) <> Windows.Forms.DialogResult.Yes Then
            Exit Sub
        End If

        With FpSpread1.ActiveSheet
            Dim podate As String = DateTimeInput1.Text
            Dim NowPoNo As String = ""
            Dim mdno As String = ""
            Dim S_CD As String = ""

            For i As Integer = 0 To .RowCount - 1
                If .GetValue(i, 10) > 0 Then
                    If .GetValue(i, 12) = True Then
                        If Mid(.GetValue(i, 1), 1, 1) = "F" Then
                            mdno = .GetValue(i, 0)
                            NowPoNo = Query_RS("exec SP_FrmPurchaseOrder_ETSpono '" & Site_id & "','" & podate & "','" & mdno & "','N'")
                            S_CD = "P2014-0004"

                            If Insert_Data("insert tbl_poheader(site_id,po_no,po_date,sup_cd,status_dv,model,remark,CUR_UNIT,c_person,c_date,u_person,u_date) values ('" & Site_id & "','" & NowPoNo & "', '" & podate & "' ,'" & S_CD & "','OPENED','" & mdno & "','자동발주','KRW','" & Emp_No & "',getdate(),'" & Emp_No & "',getdate())") = False Then
                                MessageBox.Show("[P/O HEADER] Failed to Save")
                                Exit Sub
                            End If
                        ElseIf Mid(.GetValue(i, 1), 1, 1) = "P" Then
                            mdno = .GetValue(i, 0)
                            NowPoNo = Query_RS("exec SP_FrmPurchaseOrder_ETSpono '" & Site_id & "','" & podate & "','" & mdno & "','N'")
                            S_CD = "P2014-0006"

                            If Insert_Data("insert tbl_poheader(site_id,po_no,po_date,sup_cd,status_dv,model,remark,CUR_UNIT,c_person,c_date,u_person,u_date) values ('" & Site_id & "','" & NowPoNo & "', '" & podate & "' ,'" & S_CD & "','OPENED','" & mdno & "','자동발주','KRW','" & Emp_No & "',getdate(),'" & Emp_No & "',getdate())") = False Then
                                MessageBox.Show("[P/O HEADER] Failed to Save")
                                Exit Sub
                            End If
                        ElseIf Mid(.GetValue(i, 1), 1, 1) = "O" Then
                            mdno = .GetValue(i, 0)
                            NowPoNo = Query_RS("exec SP_FrmPurchaseOrder_ETSpono '" & Site_id & "','" & podate & "','" & mdno & "','N'")
                            S_CD = "P2014-0005"

                            If Insert_Data("insert tbl_poheader(site_id,po_no,po_date,sup_cd,status_dv,model,remark,CUR_UNIT,c_person,c_date,u_person,u_date) values ('" & Site_id & "','" & NowPoNo & "', '" & podate & "' ,'" & S_CD & "','OPENED','" & mdno & "','자동발주','KRW','" & Emp_No & "',getdate(),'" & Emp_No & "',getdate())") = False Then
                                MessageBox.Show("[P/O HEADER] Failed to Save")
                                Exit Sub
                            End If
                        Else
                            If mdno <> .GetValue(i, 0) Or S_CD <> "S2014-0001" Then
                                mdno = .GetValue(i, 0)
                                NowPoNo = Query_RS("exec SP_FrmPurchaseOrder_ETSpono '" & Site_id & "','" & podate & "','" & mdno & "','N'")
                                S_CD = "S2014-0001"

                                If Insert_Data("insert tbl_poheader(site_id,po_no,po_date,sup_cd,status_dv,model,remark,CUR_UNIT,c_person,c_date,u_person,u_date) values ('" & Site_id & "','" & NowPoNo & "','" & podate & "' ,'" & S_CD & "','OPENED','" & mdno & "','자동발주','KRW','" & Emp_No & "',getdate(),'" & Emp_No & "',getdate())") = False Then
                                    MessageBox.Show("[P/O HEADER] Failed to Save")
                                    Exit Sub
                                End If
                            End If
                        End If



                        If Insert_Data("insert tbl_podetail(site_id,po_no,part_no,o_qty,in_qty,o_price,closing_date,status,c_person,c_date,u_person,u_date) values ('" & Site_id & "','" & NowPoNo & "','" & .GetValue(i, 1) & "'," & .GetValue(i, 10) & ",0," & .GetValue(i, 5) & ",Null,'OPENED','" & Emp_No & "',getdate(),'" & Emp_No & "',getdate())") = False Then
                            MessageBox.Show("[P/O DETAILS] Failed to Save")
                            Exit Sub

                        End If


                    End If


                End If
            Next
        End With

        MessageBox.Show("자재 발주가 완료되었습니다. 자재 주문/입고에서 확인하십시오.")


    End Sub

    Private Sub ButtonItem4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonItem4.Click

    End Sub

    Private Sub ButtonItem5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonItem5.Click, PrtBtn1.Click
        If save_excel = "" Then
            MessageBox.Show("Select Spread for Printing!!")
            Exit Sub
        End If

        If save_excel = "FpSpread1" Then
            If Spread_Print(Me.FpSpread1, "NO SHIP(Board) Summary", 0) = False Then
                MsgBox("Fail to Print")
            End If
            'ElseIf save_excel = "FpSpread2" Then
            '    If Spread_Print(Me.FpSpread2, "NO SHIP(Board) Details", 0) = False Then
            '        MsgBox("Fail to Print")
            '    End If
        Else
            MessageBox.Show("Select Spread for Printing!!")
        End If
    End Sub

    Private Sub ButtonItem6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonItem6.Click, XlsBtn1.Click

        If save_excel = "" Then
            MessageBox.Show("Select Spread for Saving!!")
            Exit Sub
        End If

        If save_excel = "FpSpread1" Then
            File_Save(SaveFileDialog1, FpSpread1)
            'ElseIf save_excel = "FpSpread2" Then
            '    File_Save(SaveFileDialog1, FpSpread2)
        Else
            MessageBox.Show("Select Spread for Saving!!")
        End If

    End Sub

    Private Sub FpSpread1_CellClick(ByVal sender As System.Object, ByVal e As FarPoint.Win.Spread.CellClickEventArgs) Handles FpSpread1.CellClick
        save_excel = "FpSpread1"

        '  disp_chart(e.Row)

    End Sub


    Private Sub FpSpread1_CellDoubleClick(ByVal sender As Object, ByVal e As FarPoint.Win.Spread.CellClickEventArgs) Handles FpSpread1.CellDoubleClick

        With FpSpread1.ActiveSheet

            If .RowCount = 0 Then
                Exit Sub
            End If

            Select Case e.Column
                Case 3
                    .Cells(e.Row, e.Column).Locked = False
                Case 4
                    .Cells(e.Row, e.Column).Locked = False
                Case 5
                    .Cells(e.Row, e.Column).Locked = False
                Case 6
                    .Cells(e.Row, e.Column).Locked = False
            End Select

        End With

    End Sub

    Private Sub FpSpread1_Change(ByVal sender As Object, ByVal e As FarPoint.Win.Spread.ChangeEventArgs) Handles FpSpread1.Change

        With FpSpread1.ActiveSheet
            .SetActiveCell(e.Row, 0)
            .Rows(e.Row).ForeColor = Color.OrangeRed

            .Cells(e.Row, 7).Value = .Cells(e.Row, 3).Value * .Cells(e.Row, 4).Value * .Cells(e.Row, 6).Value + .Cells(e.Row, 3).Value * .Cells(e.Row, 6).Value
            .Cells(e.Row, 10).Value = .Cells(e.Row, 7).Value '+ .Cells(e.Row, 9).Value
            .Cells(e.Row, 11).Value = (.Cells(e.Row, 10).Value) * .Cells(e.Row, 5).Value

            If .Cells(e.Row, 10).Value < 0 Then
                .Cells(e.Row, 10).Value = 0
                .Cells(e.Row, 11).Value = 0
            End If

        End With

    End Sub

    Private Sub ButtonItem7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonItem7.Click
        With FpSpread1.ActiveSheet
            For i As Integer = 0 To .RowCount - 1
                .Cells(i, 12).Value = 1
            Next
        End With
    End Sub

    Private Sub ButtonItem8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonItem8.Click
        With FpSpread1.ActiveSheet
            For i As Integer = 0 To .RowCount - 1
                .Cells(i, 12).Value = 0
            Next
        End With
    End Sub
End Class
