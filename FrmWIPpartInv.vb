Public Class FrmWIPpartInv
    Private Sub FrmWIPpartInv_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Me.DockContainerItem2.Text = "조회 조건"
        Me.DockContainerItem1.Text = "실행 메뉴"
        Me.DockContainerItem7.Text = "수불 현황"

        curcell.DecimalPlaces = 0
        'curcell.CurrencySymbol = "\"

        Condi_Disp() '콤보박스의 조건데이터 출력
        If Spread_Setting(FpSpread1, Me.Name) = True Then
            disp_hr()
            Spread_AutoCol(FpSpread1)
            Me.FpSpread1.ActiveSheet.FrozenColumnCount = 1
        End If

        With FpSpread2.ActiveSheet
            .ColumnCount = 8
            .RowCount = 0

            .ColumnHeader.Columns(0).Label = "모델"
            .ColumnHeader.Columns(1).Label = "품목번호"
            .ColumnHeader.Columns(2).Label = "품목명"
            .ColumnHeader.Columns(3).Label = "수불일자"
            .ColumnHeader.Columns(4).Label = "입고수량"
            .ColumnHeader.Columns(5).Label = "RMA수량"
            .ColumnHeader.Columns(6).Label = "출고수량"
            .ColumnHeader.Columns(7).Label = "불량수량"

            .Columns(4, 7).CellType = intcell

        End With


        Me.ContextMenuBar1.SetContextMenuEx(Me.FpSpread1, CtxSp)

        chkPHYCIAL_Click(sender, e)

        curcell.DecimalPlaces = 0
        'curcell.CurrencySymbol = "\"

        Formbim_Authority(Me.SaveBtn, Me.Name, "SAVE")
        Formbtn_Authority(Me.SaveBtn1, Me.Name, "SAVE")

        Formbim_Authority(Me.PrtBtn, Me.Name, "PRINT")
        Formbtn_Authority(Me.PrtBtn1, Me.Name, "PRINT")
        Formbim_Authority(Me.Excel, Me.Name, "EXCEL")
        Formbtn_Authority(Me.XlsBtn1, Me.Name, "EXCEL")
        Formbim_Authority(Me.FindBtn, Me.Name, "FIND")
    End Sub

    Private Sub Condi_Disp() 'CONTROL PANEL 및 파트리스트뷰에 있는 콤보박스에 데이터 출력

        Me.ModelCb.Text = "ALL"
        Query_Combo(Me.ModelCb, "select distinct p_no from tbl_bom where site_id = '" & Site_id & "'")
        Me.ModelCb.Items.Add("ALL")


        Me.ComboBoxEx1.Text = "ALL"
        Me.ComboBoxEx1.Items.AddRange(New String() {"ALL", "COSMETIC", "TECH", "FOC", "ETC"})

        Me.DateTimeInput1.Value = Now.Date

    End Sub

    Private Sub FindBtn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FindBtn.Click
        disp_wippartinv()
    End Sub

    Private Sub PrtBtn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PrtBtn1.Click, PrtBtn.Click, bPrint.Click
        Dim S1 = Me.FpSpread1.ActiveSheet
        If S1.RowCount > 0 Then
            If Spread_Print(Me.FpSpread1, "WIP Part Inventory Summary", 1) = False Then
                MsgBox("Fail to Print")
            End If
        End If
    End Sub

    Private Sub Excel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles XlsBtn1.Click, Excel.Click, bExcel.Click
        Me.FpSpread1.ActiveSheet.Columns(0).Visible = True
        If File_Save(SaveFileDialog1, FpSpread1) = True Then
            'Me.FpSpread1.ActiveSheet.Columns(0).Visible = False
        End If
    End Sub


    Private Sub disp_wippartinv()
        Try
            Dim QueryPo As String = ""
            Dim i, j As Integer
            Dim totAmt As Decimal = 0D
            Dim totAmt2 As Decimal = 0D
            Dim partdv As String = ""

            modelb.Text = ""
            '  "ALL", "LEVEL 1&2", "TECH", "FOC", "ETC"
            Select Case Me.ComboBoxEx1.Text  '파트 구분
                Case "ALL"
                    partdv = ""
                Case "COSMETIC"
                    partdv = "C"
                Case "TECH"
                    partdv = "T"
                Case "FOC"
                    partdv = "F"
                Case "ETC"
                    partdv = "E"
            End Select

            If Me.CheckBoxX1.Checked = False Then  '파트 실사하여 업로드한 최종 데이터 표시
                QueryPo = "EXEC SP_FrmWIPpartInv_LIST2 '" & Site_id & "', '" & Me.ModelCb.Text & "','" & Me.PartNoTxt.Text & "','" & Format(Me.DateTimeInput1.Value, "yyyyMMdd") & "','" & partdv & "'"

            Else 'Real 체크시 실시간 파트 수량 표시
                'QueryPo = "EXEC SP_FrmWIPpartInv_LIST1 '" & Site_id & "', '" & Me.ModelCb.Text & "','" & Me.PartNoTxt.Text & "','" & partdv & "','" & CheckBoxX2.Checked & "'"

                'Dim aa As String = Format(DateTimeInput1.Value, "yyyyMMdd")

                If Format(DateTimeInput1.Value, "yyyyMMdd") = Query_RS("select convert(varchar(8), getdate(),112)") Then
                    QueryPo = "SELECT 	C_NO, PART_NAME, PRICE, MODEL, EXAM_WH, PART_WH, PARTWAIT_WH, PARTLOSS_WH, LINELOSS_WH, LINEUSE_WH, LINEWAIT_WH,  (EXAM_WH + PART_WH + PARTWAIT_WH + PARTLOSS_WH+LINEWAIT_WH +LINEUSE_WH  + LINELOSS_WH)," & vbNewLine
                    QueryPo = QueryPo & "	(EXAM_WH + PART_WH + PARTWAIT_WH + PARTLOSS_WH+LINEWAIT_WH +LINEUSE_WH  + LINELOSS_WH /*LOSSWH*/)*PRICE," & vbNewLine
                    QueryPo = QueryPo & "0, 0,0, 0, 0, 0,		0" & vbNewLine
                    QueryPo = QueryPo & "FROM" & vbNewLine
                    QueryPo = QueryPo & "(" & vbNewLine
                    QueryPo = QueryPo & "	SELECT  B.PART_NO AS C_NO, B.PART_NAME AS PART_NAME," & vbNewLine
                    QueryPo = QueryPo & "CASE (SELECT COUNT(P_NO) FROM TBL_BOM WHERE SITE_ID = B.SITE_ID AND C_NO = B.PART_NO) WHEN 0 THEN '' WHEN 1 THEN (SELECT P_NO FROM TBL_BOM WHERE SITE_ID = B.SITE_ID AND C_NO = B.PART_NO) ELSE (SELECT MAX(P_NO) FROM TBL_BOM WHERE SITE_ID = B.SITE_ID AND C_NO = B.PART_NO)  END AS MODEL," & vbNewLine
                    QueryPo = QueryPo & "ISNULL((SELECT SUM(QTY) FROM TBL_PARTINV WHERE SITE_ID = B.SITE_ID AND PART_NO = B.PART_NO AND WH_CD ='E1000'),0) AS EXAM_WH," & vbNewLine
                    QueryPo = QueryPo & "ISNULL((SELECT SUM(QTY) FROM TBL_PARTINV WHERE SITE_ID = B.SITE_ID AND PART_NO = B.PART_NO AND WH_CD ='W1000'),0) AS PART_WH," & vbNewLine
                    QueryPo = QueryPo & "ISNULL((SELECT SUM(QTY) FROM TBL_PARTINV WHERE SITE_ID = B.SITE_ID AND PART_NO = B.PART_NO AND WH_CD ='W2000'),0) AS PARTwait_WH," & vbNewLine
                    QueryPo = QueryPo & "ISNULL((SELECT SUM(QTY) FROM TBL_PARTINV WHERE SITE_ID = B.SITE_ID AND PART_NO = B.PART_NO AND WH_CD ='B1000'),0) AS PARTLOSS_WH," & vbNewLine
                    QueryPo = QueryPo & "ISNULL((SELECT SUM(QTY) FROM TBL_PARTINV WHERE SITE_ID = B.SITE_ID AND PART_NO = B.PART_NO AND WH_CD ='C1000'),0) AS LINEWAIT_WH," & vbNewLine
                    QueryPo = QueryPo & "ISNULL((SELECT SUM(QTY) FROM TBL_PARTINV WHERE SITE_ID = B.SITE_ID AND PART_NO = B.PART_NO AND WH_CD ='C2000'),0) AS LINEUSE_WH," & vbNewLine
                    QueryPo = QueryPo & "ISNULL((SELECT SUM(QTY) FROM TBL_PARTINV WHERE SITE_ID = B.SITE_ID AND PART_NO = B.PART_NO AND WH_CD ='B2000'),0) AS LINELOSS_WH," & vbNewLine
                    QueryPo = QueryPo & "0 AS PPARTWH," & vbNewLine
                    QueryPo = QueryPo & "0 AS PTECHWH," & vbNewLine
                    QueryPo = QueryPo & "0 AS PCOSWH," & vbNewLine
                    QueryPo = QueryPo & "B.PRICE" & vbNewLine

                    If ModelCb.Text = "ALL" Then
                        QueryPo = QueryPo & "FROM TBL_PARTMASTER B" & vbNewLine
                        QueryPo = QueryPo & "WHERE B.SITE_ID = '" & Site_id & "'" & vbNewLine
                    Else
                        QueryPo = QueryPo & "FROM TBL_BOM A, TBL_PARTMASTER B" & vbNewLine
                        QueryPo = QueryPo & "WHERE B.SITE_ID = '" & Site_id & "'" & vbNewLine

                        QueryPo = QueryPo & "AND A.P_NO = '" & ModelCb.Text & "'" & vbNewLine
                        QueryPo = QueryPo & "AND A.C_NO = B.PART_NO" & vbNewLine
                    End If


                    QueryPo = QueryPo & "	 AND B.PART_NO LIKE '%' + '" & PartNoTxt.Text & "' + '%'" & vbNewLine
                    ' QueryPo = QueryPo & "	 AND RIGHT(B.PART_NO,1) <> 'R'" & vbNewLine
                    QueryPo = QueryPo & "GROUP BY B.SITE_ID, B.PART_NO, B.PART_NAME, B.PRICE" & vbNewLine
                    QueryPo = QueryPo & ")  C" & vbNewLine

                    If CheckBoxX2.Checked = True Then
                        QueryPo = QueryPo & "WHERE (EXAM_WH + PART_WH + PARTWAIT_WH + PARTLOSS_WH+LINEWAIT_WH +LINEUSE_WH  + LINELOSS_WH /*LOSSWH*/) > 0" & vbNewLine
                    End If


                    QueryPo = QueryPo & "ORDER BY C_NO" & vbNewLine

                Else
                    QueryPo = "SELECT C_NO, PART_NAME, PRICE, MODEL, EXAM_WH, PART_WH,PARTWAIT_WH, PARTLOSS_WH, LINELOSS_WH , LINEUSE_WH, LINEWAIT_WH, (EXAM_WH + PART_WH + PARTWAIT_WH + PARTLOSS_WH+LINEWAIT_WH +LINEUSE_WH  + LINELOSS_WH /* LOSSWH*/)," & vbNewLine
                    QueryPo = QueryPo & "	(EXAM_WH + PART_WH + PARTWAIT_WH + PARTLOSS_WH+LINEWAIT_WH +LINEUSE_WH  + LINELOSS_WH /* LOSSWH*/)*PRICE," & vbNewLine
                    QueryPo = QueryPo & "0, 0,0, 0, 0, 0,		0" & vbNewLine
                    QueryPo = QueryPo & "FROM" & vbNewLine
                    QueryPo = QueryPo & "(" & vbNewLine
                    QueryPo = QueryPo & "	SELECT  B.PART_NO AS C_NO, B.PART_NAME AS PART_NAME," & vbNewLine
                    QueryPo = QueryPo & "CASE (SELECT COUNT(P_NO) FROM TBL_BOM WHERE SITE_ID = B.SITE_ID AND C_NO = B.PART_NO) WHEN 0 THEN '' WHEN 1 THEN (SELECT P_NO FROM TBL_BOM WHERE SITE_ID = B.SITE_ID AND C_NO = B.PART_NO) ELSE (SELECT MAX(P_NO) FROM TBL_BOM WHERE SITE_ID = B.SITE_ID AND C_NO = B.PART_NO)  END AS MODEL," & vbNewLine
                    QueryPo = QueryPo & "ISNULL((SELECT SUM(QTY) FROM TBL_PARTINV WHERE SITE_ID = B.SITE_ID AND PART_NO = B.PART_NO AND WH_CD ='E1000'),0) AS EXAM_WH," & vbNewLine
                    QueryPo = QueryPo & "ISNULL((SELECT SUM(QTY) FROM TBL_PARTINV WHERE SITE_ID = B.SITE_ID AND PART_NO = B.PART_NO AND WH_CD ='W1000'),0) AS PART_WH," & vbNewLine
                    QueryPo = QueryPo & "ISNULL((SELECT SUM(QTY) FROM TBL_PARTINV WHERE SITE_ID = B.SITE_ID AND PART_NO = B.PART_NO AND WH_CD ='W2000'),0) AS PARTWAIT_WH," & vbNewLine
                    QueryPo = QueryPo & "ISNULL((SELECT SUM(QTY) FROM TBL_PARTINV WHERE SITE_ID = B.SITE_ID AND PART_NO = B.PART_NO AND WH_CD ='B1000'),0) AS PARTLOSS_WH," & vbNewLine
                    QueryPo = QueryPo & "ISNULL((SELECT SUM(QTY) FROM TBL_PARTINV WHERE SITE_ID = B.SITE_ID AND PART_NO = B.PART_NO AND WH_CD ='C1000'),0) AS LINEWAIT_WH," & vbNewLine
                    QueryPo = QueryPo & "ISNULL((SELECT SUM(QTY) FROM TBL_PARTINV WHERE SITE_ID = B.SITE_ID AND PART_NO = B.PART_NO AND WH_CD ='C2000'),0) AS LINEUSE_WH," & vbNewLine
                    QueryPo = QueryPo & "ISNULL((SELECT SUM(QTY) FROM TBL_PARTINV WHERE SITE_ID = B.SITE_ID AND PART_NO = B.PART_NO AND WH_CD ='B2000'),0) AS LINELOSS_WH," & vbNewLine
                    QueryPo = QueryPo & "0 AS PPARTWH," & vbNewLine
                    QueryPo = QueryPo & "0 AS PTECHWH," & vbNewLine
                    QueryPo = QueryPo & "0 AS PCOSWH," & vbNewLine
                    QueryPo = QueryPo & "B.PRICE" & vbNewLine

                    If ModelCb.Text = "ALL" Then
                        QueryPo = QueryPo & "FROM TBL_PARTMASTER B" & vbNewLine
                        QueryPo = QueryPo & "WHERE B.SITE_ID = '" & Site_id & "'" & vbNewLine
                    Else
                        QueryPo = QueryPo & "FROM TBL_BOM A, TBL_PARTMASTER B" & vbNewLine
                        QueryPo = QueryPo & "WHERE B.SITE_ID = '" & Site_id & "'" & vbNewLine

                        QueryPo = QueryPo & "AND A.P_NO = '" & ModelCb.Text & "'" & vbNewLine
                        QueryPo = QueryPo & "AND A.C_NO = B.PART_NO" & vbNewLine
                    End If

                    QueryPo = QueryPo & "	 AND B.PART_NO LIKE '%' + '" & PartNoTxt.Text & "' + '%'" & vbNewLine
                    'QueryPo = QueryPo & "	 AND RIGHT(B.PART_NO,1) <> 'R'" & vbNewLine
                    QueryPo = QueryPo & "GROUP BY B.SITE_ID, B.PART_NO, B.PART_NAME, B.PRICE" & vbNewLine
                    QueryPo = QueryPo & ")  C" & vbNewLine

                    If CheckBoxX2.Checked = True Then
                        QueryPo = QueryPo & "WHERE (EXAM_WH + PART_WH + PARTWAIT_WH + PARTLOSS_WH+LINEWAIT_WH +LINEUSE_WH  + LINELOSS_WH /*LOSSWH*/) > 0" & vbNewLine
                    End If

                    QueryPo = QueryPo & "ORDER BY C_NO" & vbNewLine
                End If
            End If

            If Query_Spread(Me.FpSpread1, QueryPo, 1) = True Then


                For i = 0 To Me.FpSpread1.ActiveSheet.RowCount - 1
                    totAmt += Me.FpSpread1.ActiveSheet.GetValue(i, 12)
                    totAmt2 += Me.FpSpread1.ActiveSheet.GetValue(i, 17)
                    If Me.CheckBoxX1.Checked = True Then
                        For j = 0 To 11
                            Me.FpSpread1.ActiveSheet.Cells(i, j).Locked = True
                        Next
                    End If
                Next

                With FpSpread1.ActiveSheet
                    .RowCount = .RowCount + 1
                    .Cells(.RowCount - 1, 4, .RowCount - 1, .ColumnCount - 1).CellType = intcell
                    .Cells(.RowCount - 1, 12).CellType = curcell
                    .Cells(.RowCount - 1, 18).CellType = curcell
                    .Rows(.RowCount - 1).BackColor = Color.Yellow
                    .SetText(.RowCount - 1, 0, "TOTAL")

                    SPREAD_ROW_TOTAL(FpSpread1, 4, .ColumnCount - 1, 1)

                End With

            End If
            Me.LabelItem4.Text = Format(totAmt, "" & "###,###,###,##0") '북재고 총 수량
            Me.LabelItem6.Text = Format(totAmt2, "" & "###,###,###,##0.00") '실물재고 총 수량
            Spread_AutoCol(Me.FpSpread1)
            MessageBox.Show("Complete to Load", "Message")
        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message, "ERROR")
        End Try
    End Sub


    Private Sub disp_techwh(ByVal partno As String)

        '수량 클릭시 입출고 현황 조회
        Try
            'Dim i As Integer
            'Me.Bar4.AutoHide = False

            Dim qry As String = ""

            With FpSpread2.ActiveSheet
                qry = qry & "select distinct case when (select count(p_no) from tbl_bom where c_no = a.part_no) > 1 then '다중모델' else (select p_no from tbl_bom where c_no = a.part_no) end, a.part_no , a.part_name, replace(convert(varchar(10), c.c_date, 102),'.','-') as io_dt, " & vbNewLine
                qry = qry & "		ISNULL((select SUM(qty) from tbl_partio where ((f_wh = 'S2014-0001' and t_wh = 'W1000') /*OR (f_wh = 'S2014-0001' and t_wh = 'E1000')*/  OR (f_wh = 'E1000' and t_wh = 'W1000')) " & vbNewLine
                qry = qry & "		     and part_no = a.part_no and part_no not like '%R' and convert(varchar(10), c_date, 102) = convert(varchar(10), c.c_date, 102)),0) AS RCV," & vbNewLine
                qry = qry & "		ISNULL((select SUM(qty) from tbl_partio where ((f_wh = 'S2014-0001' and t_wh = 'W1000') OR (f_wh = 'S2014-0001' and t_wh = 'E1000'))  " & vbNewLine
                qry = qry & "		     and part_no = LEFT(a.part_no,11)+'R' and convert(varchar(10), c_date, 102) = convert(varchar(10), c.c_date, 102)),0) AS RMA,  " & vbNewLine
                qry = qry & "		ISNULL((select SUM(qty) from tbl_partio where T_wh = 'U2014-0001' " & vbNewLine
                qry = qry & "		     and part_no = a.part_no /*and part_no not like '%R'*/ and convert(varchar(10), c_date, 102) = convert(varchar(10), c.c_date, 102)),0) AS SHIPOUT,   " & vbNewLine
                qry = qry & "		ISNULL((select SUM(qty) from tbl_partio where F_WH = 'C1000' AND T_wh = 'B1000' " & vbNewLine
                qry = qry & "		     and part_no = a.part_no /*and part_no not like '%R'*/ and convert(varchar(10), c_date, 102) = convert(varchar(10), c.c_date, 102)),0) AS FAIL   " & vbNewLine
                qry = qry & "from tbl_partmaster a, tbl_partio c" & vbNewLine
                qry = qry & "where a.part_no IN ('" & partno & "')" & vbNewLine
                qry = qry & "and (a.part_no = c.part_no )" & vbNewLine
                qry = qry & "and c.c_date >= '2015-01-01'" & vbNewLine
                qry = qry & "and c.t_wh in ('E1000','B1000','W1000', 'U2014-0001')" & vbNewLine
                'qry = qry & "UNION " & vbNewLine
                'qry = qry & "select distinct case when (select count(p_no) from tbl_bom where c_no = a.part_no) > 1 then '다중모델' else (select p_no from tbl_bom where c_no = a.part_no) end, a.part_no , a.part_name, replace(convert(varchar(10), c.c_date, 102),'.','-') as io_dt, " & vbNewLine
                'qry = qry & "		0 AS RCV," & vbNewLine
                'qry = qry & "		ISNULL((select SUM(qty) from tbl_partio where ((f_wh = 'S2014-0001' and t_wh = 'W1000') OR (f_wh = 'S2014-0001' and t_wh = 'E1000'))  " & vbNewLine
                'qry = qry & "		     and part_no = LEFT(a.part_no,11)+'R' and convert(varchar(10), c_date, 102) = convert(varchar(10), c.c_date, 102)),0) AS RMA,  " & vbNewLine
                'qry = qry & "		0 AS SHIPOUT,   " & vbNewLine
                'qry = qry & "		0 AS FAIL   " & vbNewLine
                'qry = qry & "from tbl_partmaster a, tbl_partio c" & vbNewLine
                'qry = qry & "where a.part_no IN ('" & Mid(partno, 1, 11) & "R')" & vbNewLine
                'qry = qry & "and (a.part_no = LEFT(c.part_no,11)+'R' )" & vbNewLine
                'qry = qry & "and c.c_date >= '2015-01-01'" & vbNewLine
                'qry = qry & "and c.t_wh in ('E1000','B1000','W1000', 'U2014-0001')" & vbNewLine

                qry = qry & "order by io_dt " & vbNewLine

                If Query_Spread(Me.FpSpread2, qry, 1) = True Then

                    .RowCount = .RowCount + 1
                    .SetValue(.RowCount - 1, 0, "TOTAL")
                    .Rows(.RowCount - 1).BackColor = Color.Yellow
                    .Cells(.RowCount - 1, 4, .RowCount - 1, .ColumnCount - 1).CellType = intcell

                    SPREAD_ROW_TOTAL_LTD(FpSpread2, 4, .ColumnCount - 1, .RowCount - 1)

                    .Cells(.RowCount - 1, 3, .RowCount - 1, 7).CellType = intcell
                    .Cells(.RowCount - 1, 3).Value = .Cells(.RowCount - 1, 4).Value + .Cells(.RowCount - 1, 5).Value - .Cells(.RowCount - 1, 6).Value ' + .Cells(.RowCount - 1, 7).Value

                    Spread_AutoCol(Me.FpSpread2)
                End If
            End With

        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message, "ERROR")
        End Try
    End Sub


    Private Sub ButtonX1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        'Try
        '    If Me.FpSpread1.ActiveSheet.RowCount < 1 Or ModelCb.Text <> "ALL" Or ComboBoxEx1.Text <> "ALL" Then
        '        MessageBox.Show("Please MODEL is ALL and PART DV is ALL and Find Click !! ", "Alert")
        '        Exit Sub
        '    End If
        '    If File_Open3(OpenFileDialog1, Me.FpSpread3, Me.TextBoxX1, "FRMWIPPARTINV") = True Then
        '        MessageBox.Show("Complete to Upload", "PHYCIAL PART INVENTORY")


        '    End If
        'Catch ex As Exception
        '    MessageBox.Show("Error: " & ex.Message, "ERROR")
        'End Try
    End Sub

    Private Sub SaveBtn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SaveBtn.Click
        Try
            Dim i As Integer
            Dim S1 = Me.FpSpread1.ActiveSheet

            If S1.RowCount < 1 Then
                Exit Sub
            End If

            If Format(DateTimeInput1.Value, "yyyyMMdd") <> Query_RS("select convert(varchar(8), getdate(),112)") Then
                Modal_Error("당일 재고만 변경이 가능합니다.")
                Exit Sub
            End If


            With S1

                For i = 0 To .RowCount - 1
                    If .Rows(i).ForeColor = Color.OrangeRed Then
                        Dim QRY As String = ""
                        QRY = "IF EXISTS (SELECT PART_NO FROM TBL_PARTINV WHERE PART_NO = '" & .GetValue(i, 0) & "' AND WH_CD = 'C1000')" & vbNewLine
                        QRY = QRY & "BEGIN" & vbNewLine
                        QRY = QRY & "   UPDATE TBL_PARTINV SET QTY = " & .GetValue(i, 7) & " WHERE PART_NO = '" & .GetValue(i, 0) & "' AND WH_CD = 'C1000'" & vbNewLine
                        QRY = QRY & "END" & vbNewLine
                        QRY = QRY & "ELSE" & vbNewLine
                        QRY = QRY & "BEGIN" & vbNewLine
                        QRY = QRY & "	INSERT INTO TBL_PARTINV VALUES ('" & Site_id & "','C1000','" & .GetValue(i, 0) & "'," & .GetValue(i, 7) & ",'" & Emp_No & "', GETDATE(),'" & Emp_No & "', GETDATE())" & vbNewLine
                        QRY = QRY & "END" & vbNewLine

                        QRY = QRY & "IF EXISTS (SELECT PART_NO FROM TBL_PARTINV WHERE PART_NO = '" & .GetValue(i, 0) & "' AND WH_CD = 'B1000')" & vbNewLine
                        QRY = QRY & "BEGIN" & vbNewLine
                        QRY = QRY & "   UPDATE TBL_PARTINV SET QTY = " & .GetValue(i, 8) & " WHERE PART_NO = '" & .GetValue(i, 0) & "' AND WH_CD = 'B1000'" & vbNewLine
                        QRY = QRY & "END" & vbNewLine
                        QRY = QRY & "ELSE" & vbNewLine
                        QRY = QRY & "BEGIN" & vbNewLine
                        QRY = QRY & "	INSERT INTO TBL_PARTINV VALUES ('" & Site_id & "','B1000','" & .GetValue(i, 0) & "'," & .GetValue(i, 7) & ",'" & Emp_No & "', GETDATE(),'" & Emp_No & "', GETDATE())" & vbNewLine
                        QRY = QRY & "END" & vbNewLine

                        If Insert_Data(QRY) = True Then
                            .Rows(i).ForeColor = Color.Black
                        End If
                    End If
                Next

            End With

        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message, "ERROR")
        End Try
    End Sub

    Private Sub ButtonItem2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonItem2.Click
        Dim S1 = Me.FpSpread1.ActiveSheet
        modelb.Text = ""
        S1.Rows.Clear()
    End Sub

    Private Sub FpSpread1_CellDoubleClick(ByVal sender As Object, ByVal e As FarPoint.Win.Spread.CellClickEventArgs) Handles FpSpread1.CellDoubleClick

        With FpSpread1.ActiveSheet

            disp_techwh(.GetValue(.ActiveRowIndex, 0))

            If e.Column = 7 Then
                .Cells(e.Row, e.Column).Locked = False
            ElseIf e.Column = 8 Then
                .Cells(e.Row, e.Column).Locked = False
            End If

        End With

    End Sub

    Private Sub FpSpread1_Change(ByVal sender As Object, ByVal e As FarPoint.Win.Spread.ChangeEventArgs) Handles FpSpread1.Change
        Try
            'Me.FpSpread1.ActiveSheet.Cells(e.Row, 16).Formula = "SUM(M" & e.Row + 1 & ":P" & e.Row + 1 & ")"
            'Me.FpSpread1.ActiveSheet.SetValue(e.Row, 17, Me.FpSpread1.ActiveSheet.GetValue(e.Row, 16) * Me.FpSpread1.ActiveSheet.GetValue(e.Row, 2))
            'Me.FpSpread1.ActiveSheet.SetValue(e.Row, 18, Me.FpSpread1.ActiveSheet.GetValue(e.Row, 10) - Me.FpSpread1.ActiveSheet.GetValue(e.Row, 16))
            Spread_Change(Me.FpSpread1, e.Row)
        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message, "ERROR")
        End Try
    End Sub

    Private Sub disp_hr()
        With Me.FpSpread1.ActiveSheet
            .ColumnHeader.Rows(0).Add()
            .ColumnHeader.Cells(0, 0).BackColor = Color.AliceBlue
            .ColumnHeader.Cells(0, 0).ColumnSpan = 4
            .ColumnHeader.Cells(0, 0).Text = "품목 개요"
            .ColumnHeader.Cells(0, 3).BackColor = Color.Green
            .ColumnHeader.Cells(0, 4).ColumnSpan = 8
            .ColumnHeader.Cells(0, 4).Text = "시스템"
            .ColumnHeader.Cells(0, 13).BackColor = Color.Yellow
            .ColumnHeader.Cells(0, 13).ColumnSpan = 8
            .ColumnHeader.Cells(0, 13).Text = "재고조사"

            .Columns(12).Locked = False
            .Columns(13).Locked = False
            .Columns(14).Locked = False
        End With
    End Sub


    Private Sub chkPHYCIAL_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkPHYCIAL.Click
        Dim S1 = Me.FpSpread1.ActiveSheet
        Dim i As Integer
        For i = 13 To 16
            If chkPHYCIAL.Checked = False Then
                S1.Columns(i).Visible = False
            Else
                S1.Columns(i).Visible = True
            End If
        Next
        If chkPHYCIAL.Checked = False Then
            S1.Columns(17).Visible = False
            S1.Columns(18).Visible = False
            S1.Columns(19).Visible = False
        Else
            S1.Columns(17).Visible = True
            S1.Columns(18).Visible = True
            S1.Columns(19).Visible = True
        End If
    End Sub

    Private Sub chkBOOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkBOOK.Click
        Dim S1 = Me.FpSpread1.ActiveSheet
        Dim i As Integer
        For i = 5 To 9
            If chkBOOK.Checked = False Then
                S1.Columns(i).Visible = False
            Else
                S1.Columns(i).Visible = True
            End If
        Next
        If chkBOOK.Checked = False Then
            S1.Columns(10).Visible = False
            S1.Columns(11).Visible = False
        Else
            S1.Columns(10).Visible = True
            S1.Columns(11).Visible = True
        End If
    End Sub

    Private Sub ButtonItem3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonItem3.Click
        'Show Book(ALL),Show Phycial(ALL)
        Dim S1 = Me.FpSpread1.ActiveSheet
        Dim i As Integer
        For i = 5 To 11
            If ButtonItem3.Text = "Hide Book(ALL)" Then
                S1.Columns(i).Visible = False
            Else
                S1.Columns(i).Visible = True
            End If
        Next
        If ButtonItem3.Text = "Hide Book(ALL)" Then
            ButtonItem3.Text = "Show Book(ALL)"
            chkBOOK.Checked = False
            chkBOOK.Visible = False
        Else
            ButtonItem3.Text = "Hide Book(ALL)"
            chkBOOK.Checked = True
            chkBOOK.Visible = True
            Bar3.Refresh()
        End If
    End Sub

    Private Sub ButtonItem4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonItem4.Click
        'Show Book(ALL),Show Phycial(ALL)
        Dim S1 = Me.FpSpread1.ActiveSheet
        Dim i As Integer
        For i = 12 To 18
            If ButtonItem4.Text = "Hide Phycial(ALL)" Then
                S1.Columns(i).Visible = False
            Else
                S1.Columns(i).Visible = True
            End If
        Next
        If ButtonItem4.Text = "Hide Phycial(ALL)" Then
            ButtonItem4.Text = "Show Phycial(ALL)"
            chkPHYCIAL.Checked = False
            chkPHYCIAL.Visible = False
        Else
            ButtonItem4.Text = "Hide Phycial(ALL)"
            chkPHYCIAL.Checked = True
            chkPHYCIAL.Visible = True
            Bar3.Refresh()
        End If
    End Sub

    'Private Sub ButtonX2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
    '    '파트실사 후 업로드 한 실물 수량을 저장
    '    Try
    '        Dim S1 = Me.FpSpread1.ActiveSheet
    '        Dim i As Integer
    '        If S3.RowCount > 0 Then
    '            MainFrm.ProgressBarItem1.Maximum = S1.RowCount
    '            For i = 0 To S3.RowCount - 1
    '                Insert_Data("EXEC SP_REALPHYINS '" & Site_id & "', '" & S3.GetValue(i, 0) & "'," & Emp2Null(S3.GetValue(i, 1)) & "," & Emp2Null(S3.GetValue(i, 2)) & "," & Emp2Null(S3.GetValue(i, 3)) & ",'" & Emp_No & "'")
    '                MainFrm.ProgressBarItem1.Value = i
    '            Next
    '        End If
    '        S3.RowCount = 0
    '        MessageBox.Show("Complete to Modify", "Sucess Modify")
    '        disp_wippartinv()
    '    Catch ex As Exception
    '        MessageBox.Show("Error: " & ex.Message, "ERROR")
    '    End Try
    'End Sub

    Private Sub FpSpread2_CellClick(ByVal sender As System.Object, ByVal e As FarPoint.Win.Spread.CellClickEventArgs) Handles FpSpread2.CellClick

    End Sub


End Class