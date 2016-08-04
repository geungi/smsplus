Public Class FrmWIPpartInv1
    Private Sub FrmWIPpartInv_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Me.DockContainerItem2.Text = "조회 조건"
        Me.DockContainerItem1.Text = "실행 메뉴"

        curcell.DecimalPlaces = 0
        'curcell.CurrencySymbol = "\"

        Condi_Disp() '콤보박스의 조건데이터 출력
        If Spread_Setting(FpSpread1, Me.Name) = True Then
            disp_hr()
            Spread_AutoCol(FpSpread1)
            Me.FpSpread1.ActiveSheet.FrozenColumnCount = 1
        End If

        Me.ContextMenuBar1.SetContextMenuEx(Me.FpSpread1, CtxSp)


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
        Me.DateTimeInput2.Value = Now.Date

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

            QueryPo = "SELECT DISTINCT A.INV_DT, A.PART_NO AS C_NO, B.PART_NAME, B.price, (SELECT MIN(P_NO) FROM TBL_BOM WHERE C_NO = B.part_no), " & vbNewLine
            QueryPo = QueryPo & "	ISNULL((SELECT QTY FROM TBL_PARTINV_d WHERE SITE_ID = B.SITE_ID AND PART_NO = B.PART_NO AND WH_CD ='E1000' and inv_dt = A.inv_dt ),0) AS EXAM_WH," & vbNewLine
            QueryPo = QueryPo & "	ISNULL((SELECT QTY FROM TBL_PARTINV_d WHERE SITE_ID = B.SITE_ID AND PART_NO = B.PART_NO AND WH_CD ='W1000' and inv_dt = A.inv_dt ),0) AS PART_WH," & vbNewLine
            ' QueryPo = QueryPo & "	ISNULL((SELECT QTY FROM TBL_PARTINV_d WHERE SITE_ID = B.SITE_ID AND PART_NO = B.PART_NO AND WH_CD ='A1000' and inv_dt = A.inv_dt ),0) AS WAIT_WH," & vbNewLine
            QueryPo = QueryPo & "	ISNULL((SELECT QTY FROM TBL_PARTINV_d WHERE SITE_ID = B.SITE_ID AND PART_NO = B.PART_NO AND WH_CD ='C1000' and inv_dt = A.inv_dt ),0) AS LINE_WH," & vbNewLine
            QueryPo = QueryPo & "	ISNULL((SELECT QTY FROM TBL_PARTINV_d WHERE SITE_ID = B.SITE_ID AND PART_NO = B.PART_NO AND WH_CD ='B1000' and inv_dt = A.inv_dt ),0) AS BAD_WH," & vbNewLine
            QueryPo = QueryPo & "	ISNULL((SELECT SUM(QTY) FROM TBL_PARTINV_d WHERE SITE_ID = B.SITE_ID AND PART_NO = B.PART_NO and inv_dt = A.inv_dt ),0) AS TOTAL," & vbNewLine
            QueryPo = QueryPo & "	ISNULL((SELECT SUM(QTY) FROM TBL_PARTINV_d WHERE SITE_ID = B.SITE_ID AND PART_NO = B.PART_NO and inv_dt = A.inv_dt ),0) * B.PRICE AS AMOUNT" & vbNewLine


            If ModelCb.Text = "ALL" Then
                QueryPo = QueryPo & "FROM TBL_PARTINV_D A, tbl_partmaster B" & vbNewLine
                QueryPo = QueryPo & "WHERE B.SITE_ID = '" & Site_id & "'" & vbNewLine
                QueryPo = QueryPo & "AND A.INV_DT BETWEEN '" & DateTimeInput1.Text & "' AND '" & DateTimeInput2.Text & "'" & vbNewLine
                QueryPo = QueryPo & "AND A.part_no = B.part_no " & vbNewLine
            Else
                QueryPo = QueryPo & "FROM TBL_PARTINV_D A, tbl_partmaster B, TBL_BOM C" & vbNewLine
                QueryPo = QueryPo & "WHERE B.SITE_ID = '" & Site_id & "'" & vbNewLine
                QueryPo = QueryPo & "AND A.INV_DT BETWEEN '" & DateTimeInput1.Text & "' AND '" & DateTimeInput2.Text & "'" & vbNewLine
                QueryPo = QueryPo & "AND A.part_no = B.part_no " & vbNewLine

                QueryPo = QueryPo & "AND C.P_NO = '" & ModelCb.Text & "'" & vbNewLine
                QueryPo = QueryPo & "AND C.C_NO = B.PART_NO" & vbNewLine
            End If

            QueryPo = QueryPo & "	 AND B.PART_NO LIKE '%' + '" & PartNoTxt.Text & "' + '%'" & vbNewLine

            If CheckBoxX2.Checked = True Then
                QueryPo = QueryPo & "AND ISNULL((SELECT SUM(QTY) FROM TBL_PARTINV_d WHERE SITE_ID = B.SITE_ID AND PART_NO = B.PART_NO and inv_dt = A.inv_dt ),0) > 0" & vbNewLine
            End If
            QueryPo = QueryPo & "ORDER BY A.PART_NO, A.INV_DT " & vbNewLine

            If Query_Spread(Me.FpSpread1, QueryPo, 1) = True Then
                For i = 0 To Me.FpSpread1.ActiveSheet.RowCount - 1
                    totAmt += Me.FpSpread1.ActiveSheet.GetValue(i, 9)
                    totAmt2 += Me.FpSpread1.ActiveSheet.GetValue(i, 10)
                    If Me.CheckBoxX1.Checked = True Then
                        For j = 0 To 10
                            Me.FpSpread1.ActiveSheet.Cells(i, j).Locked = True
                        Next
                    End If
                Next
            End If
            Me.LabelItem4.Text = Format(totAmt2, "" & "###,###,###,##0") '북재고 총 수량
            Me.LabelItem6.Text = Format(totAmt2, "" & "###,###,###,##0.00") '실물재고 총 수량
            Spread_AutoCol(Me.FpSpread1)

            MessageBox.Show("Complete to Load", "Message")

        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message, "ERROR")
        End Try
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
            .ColumnHeader.Cells(0, 4).ColumnSpan = 9
            .ColumnHeader.Cells(0, 4).Text = "시스템"




        End With
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



End Class