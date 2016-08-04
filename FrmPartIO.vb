Public Class FrmPartIO
    Const CtrlMask As Byte = 8
    Private NowPoNo As String = ""
    Private TempSno As String = ""
    Private TotRec As Integer = 0
    Private TotRec2 As Integer = 0
    Private FwhCd As String = ""
    Private TwhCd As String = ""
    Private EmpNmTXT As String = ""
    Private PartAuth As Boolean = True

    Private Sub FrmPartIO_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Condi_Disp() '콤보박스의 조건데이터 출력
        Me.DockContainerItem7.Text = "DETAIL CONDITOINS"

        If Spread_Setting(FpSpread1, "FrmPartIO") = True Then
            Spread_AutoCol(FpSpread1)
        End If

        Me.ContextMenuBar1.SetContextMenuEx(Me.FpSpread1, CtxSp)

        Me.LabelItem3.Text = ""
        Me.LabelItem4.Text = ""

        Formbtn_Authority(Me.NewBtn1, Me.Name, "NEW")
        Formbtn_Authority(Me.SaveBtn1, Me.Name, "SAVE")
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

        'From WH

        Query_WHCombo2(Me.FromCb, Site_id, True)

        'To WH

        Query_WHCombo2(Me.ToCb, Site_id, True)

        Dim i As Integer
        Dim ary1 As String()
        Dim ary2 As String()

        ReDim ary1(23)
        ReDim ary2(59)

        For i = 0 To 23
            ary1(i) = Microsoft.VisualBasic.Right(("0" + CStr(i)), 2)
        Next
        Me.ComboBoxEx1.Items.AddRange(ary1)
        Me.ComboBoxEx3.Items.AddRange(ary1)

        For i = 0 To 59
            ary2(i) = Microsoft.VisualBasic.Right(("0" + CStr(i)), 2)
        Next
        Me.ComboBoxEx2.Items.AddRange(ary2)
        Me.ComboBoxEx4.Items.AddRange(ary2)

        Me.ComboBoxEx1.Text = "00"
        Me.ComboBoxEx2.Text = "00"
        Me.ComboBoxEx3.Text = "23"
        Me.ComboBoxEx4.Text = "59"
    End Sub

    Private Sub FindBtn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FindBtn.Click
        disp_io()
    End Sub

    Private Sub PrtBtn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PrtBtn1.Click, PrtBtn.Click, bPrint.Click
        Dim S1 = Me.FpSpread1.ActiveSheet
        If S1.RowCount > 0 Then
            If Spread_Print(Me.FpSpread1, "Part I/O History", 1) = False Then
                MsgBox("Fail to Print")
            End If
        End If
    End Sub

    Private Sub Excel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles XlsBtn1.Click, Excel.Click, bExcel.Click
        Me.FpSpread1.ActiveSheet.Columns(0).Visible = True
        If File_Save(SaveFileDialog1, FpSpread1) = True Then
            Me.FpSpread1.ActiveSheet.Columns(0).Visible = False
        End If
    End Sub

    Private Sub disp_io()
        Try
            Dim QueryPo As String = ""
            Dim totamt As Decimal = 0D

            'PART I/O 출력
            '@SITE VARCHAR(10),
            '@FROM_WAREHOUSE VARCHAR(25),
            '@TO_WAREHOUSE VARCHAR(25),
            '@ST_DT DATETIME,
            '@ED_DT DATETIME,
            '@PART_NO	VARCHAR(50)
            'QueryPo = "EXEC SP_PART_IOLIST2 '" & Site_id & "', '" & Me.FromCb.SelectedValue.ToString & "', '" & Me.ToCb.SelectedValue.ToString & "', '" & Me.POStDate.Text & "', ' " & Me.ComboBoxEx1.Text & ":" & Me.ComboBoxEx2.Text & ":00','" & Me.POEdDate.Text & "',' " & Me.ComboBoxEx3.Text & ":" & Me.ComboBoxEx4.Text & ":59','" & Me.PartNoTxt.Text & "', '" & Me.SrcNoTxt.Text & "'"

            QueryPo = "SELECT C_DATE, PART_NO," & vbNewLine
            QueryPo = QueryPo & "	  isnull((SELECT '['+ CODE_ID +'] '+CODE_NAME FROM TBL_CODEMASTER WHERE SITE_ID='" & Site_id & "' AND CLASS_ID = '10024' AND CODE_ID = A.F_WH),(select sup_nm from tbl_supmst where sup_no = a.f_wh))," & vbNewLine
            QueryPo = QueryPo & "	  isnull((SELECT '['+ CODE_ID +'] '+CODE_NAME FROM TBL_CODEMASTER WHERE SITE_ID='" & Site_id & "' AND CLASS_ID = '10024' AND CODE_ID = A.T_WH),isnull((select sup_nm from tbl_supmst where sup_no = a.T_wh),'[S2014-0000] 자체처리'))," & vbNewLine
            QueryPo = QueryPo & "	   QTY, SRC_NO," & vbNewLine
            QueryPo = QueryPo & "	   (SELECT MAX(S_NO) FROM TBL_PARTRCV WHERE SITE_ID = A.SITE_ID AND PO_NO = A.SRC_NO AND PART_NO = A.PART_NO)," & vbNewLine
            QueryPo = QueryPo & "	   REMARK,(SELECT EMP_NM FROM TBL_EMPMASTER WHERE SITE_ID = A.SITE_ID AND EMP_NO = A.C_PERSON) " & vbNewLine
            QueryPo = QueryPo & "FROM TBL_PARTIO A" & vbNewLine
            QueryPo = QueryPo & "WHERE SITE_ID = '" & Site_id & "'" & vbNewLine
            If FromCb.Text <> "ALL" Then
                QueryPo = QueryPo & "	 AND F_WH LIKE '%" & Mid(FromCb.Text, 2, InStr(FromCb.Text, "]") - 2) & "%'" & vbNewLine
            End If
            If ToCb.Text <> "ALL" Then

                '                Dim AA As Integer = InStr(ToCb.Text, "]") - 1

                QueryPo = QueryPo & "	 AND T_WH LIKE '%" & Mid(ToCb.Text, 2, InStr(ToCb.Text, "]") - 2) & "%'" & vbNewLine
            End If
            QueryPo = QueryPo & "	 AND C_DATE BETWEEN '" & POStDate.Text & " " & ComboBoxEx1.Text & ":" & ComboBoxEx2.Text & ":00' AND '" & POEdDate.Text & " " & ComboBoxEx3.Text & ":" & ComboBoxEx4.Text & ":00'" & vbNewLine
            QueryPo = QueryPo & "	 AND PART_NO LIKE '%" & PartNoTxt.Text & "%'" & vbNewLine
            QueryPo = QueryPo & "	 AND SRC_NO LIKE '%" & SrcNoTxt.Text & "%'" & vbNewLine
            QueryPo = QueryPo & "ORDER BY C_DATE-- DESC" & vbNewLine

            If Me.FpSpread1.ActiveSheet.RowCount > 9999 Then
                Me.FpSpread1.ActiveSheet.RowHeader.Columns(0).Width = 50
            End If

            If Query_Spread(Me.FpSpread1, QueryPo, 1) = True Then
                'For i = 0 To Me.FpSpread1.ActiveSheet.RowCount - 1
                '    totamt += Me.FpSpread1.ActiveSheet.Cells(i, 6).Value
                '    Me.FpSpread1.ActiveSheet.Rows(i).Locked = True
                'Next
            End If

            'If Me.FromCb.Text <> "[ALL]FROM WH" And Me.ToCb.Text <> "[ALL]TO WH" Then
            '    Me.LabelItem3.Text = "Total : "
            '    Me.LabelItem4.Text = "$" & Format(totamt, "###,###,###,###,##0.00")
            'Else
            '    Me.LabelItem3.Text = ""
            '    Me.LabelItem4.Text = ""
            'End If

            Spread_AutoCol(Me.FpSpread1)

        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message, "ERROR")
        End Try
    End Sub


End Class