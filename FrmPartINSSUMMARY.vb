Public Class FrmPartINSSUMMARY
    Const CtrlMask As Byte = 8
    Private NowPoNo As String = ""
    Private TempSno As String = ""
    Private TotRec As Integer = 0
    Private TotRec2 As Integer = 0
    Private FwhCd As String = ""
    Private TwhCd As String = ""
    Private EmpNmTXT As String = ""
    Private PartAuth As Boolean = False
    Private CosWH As String = ""

    Private s As String
    Private cr As FarPoint.Win.Spread.Model.CellRange

    Private Sub CheckedListBox1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles CheckedListBox1.LostFocus
        TextBoxDropDown1.Text = MODEL_SELECTED(CheckedListBox1)
    End Sub

    Private Sub FrmPartTransfer_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        'If Partauth_yn = "Y" Then
        '    PartAuth = True
        'End If

        DockContainerItem2.Text = "조회 조건"

        If Spread_Setting(FpSpread2, "FrmPartIns") = True Then
            FpSpread2.ActiveSheet.FrozenColumnCount = 3
            Spread_AutoCol(FpSpread2)
            FpSpread2.ActiveSheet.OperationMode = FarPoint.Win.Spread.OperationMode.Normal
        End If

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

        Dim qry As String = ""
        Dim m_qry As String = ""
        m_qry = MODEL_SELECTED(CheckedListBox1)

        FpSpread2.ActiveSheet.RowCount = 0

        qry = qry & "select po_no, part_no, (select part_name from tbl_partmaster where part_no =  a.part_no), o_qty, " & vbNewLine
        qry = qry & "			(select sum(r_qty) from tbl_partrcv where po_no = a.po_no and part_no = a.part_no), " & vbNewLine
        qry = qry & "			isnull((select sum(good_qty) from tbl_partins where po_no = a.po_no and part_no = a.part_no), 0),  " & vbNewLine
        qry = qry & "			isnull((select sum(bad_qty) from tbl_partins where po_no = a.po_no and part_no = a.part_no),0)," & vbNewLine
        qry = qry & "			(select sum(r_qty) from tbl_partrcv where po_no = a.po_no and part_no = a.part_no)-  " & vbNewLine
        qry = qry & "			isnull((select sum(good_qty) from tbl_partins where po_no = a.po_no and part_no = a.part_no), 0)-  " & vbNewLine
        qry = qry & "			isnull((select sum(bad_qty) from tbl_partins where po_no = a.po_no and part_no = a.part_no),0)" & vbNewLine
        qry = qry & "from tbl_podetail a" & vbNewLine
        qry = qry & "where '20'+substring(po_no,5,6) between convert(varchar(8), convert(datetime, '" & DateTimeInput1.Text & "'), 112) and convert(varchar(8), convert(datetime, '" & DateTimeInput2.Text & "'), 112)" & vbNewLine

        'If ComboBoxEx1.Text <> "ALL" Then
        '    qry = qry & "and part_no in (select c_no from tbl_bom where p_no = '" & ComboBoxEx1.Text & "' and active = 'Y')" & vbNewLine
        'End If
        If m_qry <> "" Then
            qry = qry & "and part_no in (select c_no from tbl_bom where p_no IN (" & m_qry & ") and active = 'Y')" & vbNewLine
        End If

        If PartNoTxt.Text <> "" Then
            qry = qry & "and part_no like '" & PartNoTxt.Text & "%'" & vbNewLine
        End If
        '        qry = qry & "  AND O_QTY > ISNULL((SELECT GOOD_QTY + BAD_QTY FROM TBL_PARTINS WHERE po_no = a.po_no and part_no = a.part_no),0)" & vbNewLine
        qry = qry & "order by po_no" & vbNewLine

        If Query_Spread(FpSpread2, qry, 1) = True Then
            Spread_AutoCol(FpSpread2)
        End If


    End Sub

    Private Sub SaveBtn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SaveBtn1.Click, SaveBtn.Click, bSave.Click

    End Sub

    Private Sub DelBtn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DelBtn1.Click, DelBtn.Click, bDel.Click


    End Sub

    Private Sub PrtBtn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PrtBtn1.Click, PrtBtn.Click, bPrint.Click

    End Sub

    Private Sub Excel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles XlsBtn1.Click, Excel.Click, bExcel.Click
        Me.FpSpread2.ActiveSheet.Columns(0).Visible = True
        If File_Save(SaveFileDialog1, FpSpread2) = True Then
            Me.FpSpread2.ActiveSheet.Columns(0).Visible = False
        End If
    End Sub
    Private Sub PrtBtn2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cPrint.Click
        Dim S2 = Me.FpSpread2.ActiveSheet
        If S2.RowCount > 0 Then
            If Spread_Print(Me.FpSpread2, "", 1) = False Then
                MsgBox("Fail to Print")
            End If
        End If
    End Sub

    Private Sub Excel2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cExcel.Click
        Me.FpSpread2.ActiveSheet.Columns(0).Visible = True
        If File_Save(SaveFileDialog1, FpSpread2) = True Then
            Me.FpSpread2.ActiveSheet.Columns(0).Visible = False
        End If
    End Sub

    Private Sub disp_part()
        Try
            Query_Spread(Me.FpSpread2, "", 1)

            Spread_AutoCol(FpSpread2)
        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message, "ERROR")
        End Try
    End Sub





    Private Sub FpSpread2_CellClick(ByVal sender As System.Object, ByVal e As FarPoint.Win.Spread.CellClickEventArgs) Handles FpSpread2.CellClick
        save_excel = "FpSpread2"
    End Sub


    Private Sub ButtonItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonItem1.Click
        Dim S3 = Me.FpSpread2.ActiveSheet

        If S3.RowCount > 0 Then
            S3.Rows.Clear()
        End If
    End Sub

    Private Sub FpSpread2_MouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles FpSpread2.MouseUp
        Try
            cr = FpSpread2.ActiveSheet.GetSelection(0)

            s = FpSpread2.ActiveSheet.GetClip(cr.Row, 0, cr.RowCount, 1)
        Catch ex As Exception
            'MessageBox.Show("You didn't make a selection!!")
            Return
        End Try
    End Sub



    Private Sub PartNoTxt_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles PartNoTxt.KeyDown

        TextBoxX1.Text = ""
        TextBoxX2.Text = ""


        If e.KeyValue <> Keys.Enter Then
            Exit Sub
        End If

        TextBoxX1.Text = Query_RS("select part_name from tbl_partmaster where part_no = '" & PartNoTxt.Text & "'")
        TextBoxX2.Text = Query_RS("select part_spec from tbl_partmaster where part_no = '" & PartNoTxt.Text & "'")

        FindBtn_Click(sender, e)

    End Sub


End Class