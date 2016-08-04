Public Class FrmInvoice


    Private PkRow As New Integer
    Private p_cnt, s_cnt As New Integer


    Private Sub FrmInvoice_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        DockContainerItem1.Text = "실행메뉴"
        DockContainerItem2.Text = "조회 조건"
        DockContainerItem3.Text = "출하현황"

        If Spread_Setting(FpSpread1, Me.Name) = True Then
            With FpSpread1.ActiveSheet
                .RowCount = 0
            End With
            Spread_AutoCol(FpSpread1)
        End If

        If Spread_Setting(FpSpread2, Me.Name) = True Then
            With FpSpread2.ActiveSheet
                .RowCount = 0
            End With

            Spread_AutoCol(FpSpread2)
        End If


        If Spread_Setting(FpSpread3, Me.Name) = True Then
            With FpSpread3.ActiveSheet
                .RowCount = 0
            End With
            Spread_AutoCol(FpSpread3)
        End If


        DateTimeInput1.Value = Now
        DateTimeInput2.Value = Now
        DateTimeInput3.Value = Now

        'If Query_Combo(Me.ComboBoxEx1, "SELECT model_no FROM tbl_modelmaster WHERE site_id = '" & Site_id & "' and active = 'Y' ORDER BY model_no") = True Then
        'End If
        'Me.ComboBoxEx1.Items.Add("ALL")
        'Me.ComboBoxEx1.Text = "ALL"


        If Query_CheckList(Me.CheckedListBox1, "SELECT model_no FROM tbl_modelmaster WHERE site_id = '" & Site_id & "' and active = 'Y' ORDER BY model_no") = True Then
            Me.CheckedListBox1.Items.Add("ALL")
            Me.CheckedListBox1.Text = "ALL"
        End If



        If Query_Combo(Me.ComboBoxEx2, "SELECT CODE_ID FROM tbl_CODEmaster WHERE site_id = '" & Site_id & "' and CLASS_ID = 'R0007' AND active = 'Y' ORDER BY DIS_ORDER") = True Then
        End If
        Me.ComboBoxEx2.Items.Add("ALL")
        Me.ComboBoxEx2.Text = "ALL"

        Formbim_Authority(Me.NewBtn, Me.Name, "NEW")
        Formbtn_Authority(Me.NewBtn1, Me.Name, "NEW")
        Formbim_Authority(Me.SaveBtn, Me.Name, "SAVE")
        Formbtn_Authority(Me.SaveBtn1, Me.Name, "SAVE")
        Formbim_Authority(Me.DelBtn, Me.Name, "DELETE")
        Formbtn_Authority(Me.DelBtn1, Me.Name, "DELETE")
        Formbim_Authority(Me.PrtBtn, Me.Name, "PRINT")
        Formbtn_Authority(Me.PrtBtn1, Me.Name, "PRINT")
        Formbim_Authority(Me.XlsBtn, Me.Name, "EXCEL")
        Formbtn_Authority(Me.XlsBtn1, Me.Name, "EXCEL")
        Formbim_Authority(Me.FindBtn, Me.Name, "FIND")

        DockContainerItem3.Selected = True


        If Query_RS("select isnull(insa_yn,'N') from tbl_empmaster where emp_no = '" & Emp_No & "'") = "N" Then
            FpSpread1.ActiveSheet.Columns(5).Visible = False
            FpSpread2.ActiveSheet.Columns(3).Visible = False
            FpSpread3.ActiveSheet.Columns(2).Visible = False
            FpSpread3.ActiveSheet.Columns(4).Visible = False
            FpSpread3.ActiveSheet.Columns(6).Visible = False
        End If

    End Sub

    'ship_no 클릭
    Private Sub ButtonItem7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FpSpread2.CellDoubleClick

        FpSpread1.ActiveSheet.RowCount = 0
        FpSpread1.ActiveSheet.FrozenColumnCount = 3

        If FpSpread2.ActiveSheet.RowCount = 0 Then
            MessageBox.Show("출하번호를 선택하세요.")
            Exit Sub
        End If

        'If FpSpread2.ActiveSheet.GetValue(FpSpread2.ActiveSheet.ActiveRowIndex, FpSpread2.ActiveSheet.ColumnCount - 1) = False Then
        '    FpSpread2.ActiveSheet.SetValue(FpSpread2.ActiveSheet.ActiveRowIndex, FpSpread2.ActiveSheet.ColumnCount - 1, True)
        'End If

        'invoice
        With FpSpread1.ActiveSheet
            .RowCount = 0
            FpSpread1.AllowUserFormulas = True
            '.DataAutoCellTypes = True

            'For R_CNT As Integer = 0 To FpSpread2.ActiveSheet.RowCount - 1


            Dim QRY As String = ""

            QRY = QRY & "SELECT SHIP_DATE,SHIP_NO,  '생산출하',CASE WHEN (SELECT COUNT(MODEL_NO) FROM TBL_MODELMASTER WHERE MODEL_NO = A.MODEL) > 0 THEN A.MODEL ELSE (SELECT PART_NO + '('+PART_NAME +')' FROM TBL_PARTMASTER WHERE PART_NO = A.MODEL) END AS MODEL, SUM(QTY) , SUM(QTY*CHARGE) , RETURN_DV ," & vbNewLine
            QRY = QRY & "			ISNULL((SELECT SHIP_NO FROM TBL_INVSUMMARY WHERE SHIP_NO = A.SHIP_NO),''), " & vbNewLine
            QRY = QRY & "			ISNULL((SELECT INV_NO FROM TBL_INVSUMMARY WHERE SHIP_NO = A.SHIP_NO),''), " & vbNewLine
            QRY = QRY & "			ISNULL((SELECT INV_DATE FROM TBL_INVSUMMARY WHERE SHIP_NO = A.SHIP_NO),'') " & vbNewLine
            QRY = QRY & "FROM TBL_SHIPSUMMARY A" & vbNewLine
            QRY = QRY & "WHERE SHIP_DATE BETWEEN CONVERT(VARCHAR(8), CONVERT(DATETIME,'" & DateTimeInput1.Text & "'),112) AND CONVERT(VARCHAR(8), CONVERT(DATETIME,'" & DateTimeInput2.Text & "'),112)" & vbNewLine
            If FpSpread2.ActiveSheet.GetValue(FpSpread2.ActiveSheet.ActiveRowIndex, 0) <> "TOTAL" Then
                QRY = QRY & "AND SHIP_NO = '" & FpSpread2.ActiveSheet.GetValue(FpSpread2.ActiveSheet.ActiveRowIndex, 1) & "'" & vbNewLine
            End If
            'If ComboBoxEx1.Text <> "ALL" Then
            '    QRY = QRY & "AND MODEL = '" & ComboBoxEx1.Text & "'" & vbNewLine
            'End If
            If ComboBoxEx2.Text <> "ALL" Then
                QRY = QRY & "AND RETURN_DV = '" & ComboBoxEx2.Text & "'" & vbNewLine
            End If
            QRY = QRY & "GROUP BY SHIP_NO, SHIP_DATE , MODEL, RETURN_DV " & vbNewLine
            QRY = QRY & "UNION" & vbNewLine
            QRY = QRY & "SELECT SHIP_DATE, SHIP_NO,  '상품출하',CASE WHEN (SELECT COUNT(MODEL_NO) FROM TBL_MODELMASTER WHERE MODEL_NO = B.MODEL) > 0 THEN B.MODEL ELSE (SELECT PART_NO + '('+PART_NAME +')' FROM TBL_PARTMASTER WHERE PART_NO = B.MODEL) END AS MODEL, SUM(QTY) , SUM(QTY*CHARGE) , RETURN_DV ,'','',''" & vbNewLine
            QRY = QRY & "FROM TBL_SHIPSUMMARY_GOODS B" & vbNewLine
            QRY = QRY & "WHERE SHIP_DATE BETWEEN CONVERT(VARCHAR(8), CONVERT(DATETIME,'" & DateTimeInput1.Text & "'),112) AND CONVERT(VARCHAR(8), CONVERT(DATETIME,'" & DateTimeInput2.Text & "'),112)" & vbNewLine
            If FpSpread2.ActiveSheet.GetValue(FpSpread2.ActiveSheet.ActiveRowIndex, 0) <> "TOTAL" Then
                QRY = QRY & "AND SHIP_NO = '" & FpSpread2.ActiveSheet.GetValue(FpSpread2.ActiveSheet.ActiveRowIndex, 1) & "'" & vbNewLine
            End If
            'If ComboBoxEx1.Text <> "ALL" Then
            '    QRY = QRY & "AND MODEL = '" & ComboBoxEx1.Text & "'" & vbNewLine
            'End If
            If ComboBoxEx2.Text <> "ALL" Then
                QRY = QRY & "AND RETURN_DV = '" & ComboBoxEx2.Text & "'" & vbNewLine
            End If
            QRY = QRY & "GROUP BY SHIP_NO, SHIP_DATE , MODEL, RETURN_DV " & vbNewLine
            QRY = QRY & "ORDER BY SHIP_NO, SHIP_DATE , MODEL, RETURN_DV " & vbNewLine
            QRY = QRY & "" & vbNewLine
            QRY = QRY & "" & vbNewLine
            QRY = QRY & "" & vbNewLine
            QRY = QRY & "" & vbNewLine

            If Query_Spread(FpSpread1, QRY, 1) = True Then
                .RowCount = .RowCount + 1
                .SetValue(.RowCount - 1, 0, "TOTAL")
                .Rows(.RowCount - 1).BackColor = Color.Yellow
                .Cells(.RowCount - 1, 5).CellType = deccell
                .Cells(.RowCount - 1, 4).CellType = intcell

                .Cells(.RowCount - 1, 5).Formula = "SUM(F1:F" & .RowCount - 1 & ")"
                .Cells(.RowCount - 1, 4).Formula = "SUM(E1:E" & .RowCount - 1 & ")"


                .ColumnFooterVisible = True
                .ColumnFooter.Cells(0, 0).Text = "출하합계"
                .ColumnFooter.Cells(0, 1).CellType = intcell
                .ColumnFooter.Cells(0, 1).Text = .GetValue(.RowCount - 1, 4)
                .ColumnFooter.Cells(0, 2).CellType = deccell
                .ColumnFooter.Cells(0, 2).Text = .GetValue(.RowCount - 1, 5)

                .ColumnFooter.Cells(0, 3).Text = "생산출하"
                .ColumnFooter.Cells(0, 4).CellType = intcell

                QRY = "SELECT SUM(QTY) " & vbNewLine
                QRY = QRY & "FROM TBL_SHIPSUMMARY" & vbNewLine
                QRY = QRY & "WHERE SHIP_DATE BETWEEN CONVERT(VARCHAR(8), CONVERT(DATETIME,'" & DateTimeInput1.Text & "'),112) AND CONVERT(VARCHAR(8), CONVERT(DATETIME,'" & DateTimeInput2.Text & "'),112)" & vbNewLine
                If FpSpread2.ActiveSheet.GetValue(FpSpread2.ActiveSheet.ActiveRowIndex, 0) <> "TOTAL" Then
                    QRY = QRY & "AND SHIP_NO = '" & FpSpread2.ActiveSheet.GetValue(FpSpread2.ActiveSheet.ActiveRowIndex, 1) & "'" & vbNewLine
                End If
                'If ComboBoxEx1.Text <> "ALL" Then
                '    QRY = QRY & "AND MODEL = '" & ComboBoxEx1.Text & "'" & vbNewLine
                'End If
                If ComboBoxEx2.Text <> "ALL" Then
                    QRY = QRY & "AND RETURN_DV = '" & ComboBoxEx2.Text & "'" & vbNewLine
                End If

                .ColumnFooter.Cells(0, 4).Text = Query_RS(QRY)

                .ColumnFooter.Cells(0, 5).CellType = deccell
                QRY = "SELECT SUM(QTY*CHARGE) " & vbNewLine
                QRY = QRY & "FROM TBL_SHIPSUMMARY" & vbNewLine
                QRY = QRY & "WHERE SHIP_DATE BETWEEN CONVERT(VARCHAR(8), CONVERT(DATETIME,'" & DateTimeInput1.Text & "'),112) AND CONVERT(VARCHAR(8), CONVERT(DATETIME,'" & DateTimeInput2.Text & "'),112)" & vbNewLine
                If FpSpread2.ActiveSheet.GetValue(FpSpread2.ActiveSheet.ActiveRowIndex, 0) <> "TOTAL" Then
                    QRY = QRY & "AND SHIP_NO = '" & FpSpread2.ActiveSheet.GetValue(FpSpread2.ActiveSheet.ActiveRowIndex, 1) & "'" & vbNewLine
                End If
                'If ComboBoxEx1.Text <> "ALL" Then
                '    QRY = QRY & "AND MODEL = '" & ComboBoxEx1.Text & "'" & vbNewLine
                'End If
                If ComboBoxEx2.Text <> "ALL" Then
                    QRY = QRY & "AND RETURN_DV = '" & ComboBoxEx2.Text & "'" & vbNewLine
                End If

                .ColumnFooter.Cells(0, 5).Text = Query_RS(QRY)

                .ColumnFooter.Cells(0, 6).Text = "상품출하"
                .ColumnFooter.Cells(0, 7).CellType = intcell


                QRY = "SELECT SUM(QTY) " & vbNewLine
                QRY = QRY & "FROM TBL_SHIPSUMMARY_GOODS" & vbNewLine
                QRY = QRY & "WHERE SHIP_DATE BETWEEN CONVERT(VARCHAR(8), CONVERT(DATETIME,'" & DateTimeInput1.Text & "'),112) AND CONVERT(VARCHAR(8), CONVERT(DATETIME,'" & DateTimeInput2.Text & "'),112)" & vbNewLine
                If FpSpread2.ActiveSheet.GetValue(FpSpread2.ActiveSheet.ActiveRowIndex, 0) <> "TOTAL" Then
                    QRY = QRY & "AND SHIP_NO = '" & FpSpread2.ActiveSheet.GetValue(FpSpread2.ActiveSheet.ActiveRowIndex, 1) & "'" & vbNewLine
                End If
                'If ComboBoxEx1.Text <> "ALL" Then
                '    QRY = QRY & "AND MODEL = '" & ComboBoxEx1.Text & "'" & vbNewLine
                'End If
                If ComboBoxEx2.Text <> "ALL" Then
                    QRY = QRY & "AND RETURN_DV = '" & ComboBoxEx2.Text & "'" & vbNewLine
                End If

                .ColumnFooter.Cells(0, 7).Text = Query_RS(QRY)
                .ColumnFooter.Cells(0, 8).CellType = deccell

                QRY = "SELECT SUM(QTY*CHARGE) " & vbNewLine
                QRY = QRY & "FROM TBL_SHIPSUMMARY_GOODS" & vbNewLine
                QRY = QRY & "WHERE SHIP_DATE BETWEEN CONVERT(VARCHAR(8), CONVERT(DATETIME,'" & DateTimeInput1.Text & "'),112) AND CONVERT(VARCHAR(8), CONVERT(DATETIME,'" & DateTimeInput2.Text & "'),112)" & vbNewLine
                If FpSpread2.ActiveSheet.GetValue(FpSpread2.ActiveSheet.ActiveRowIndex, 0) <> "TOTAL" Then
                    QRY = QRY & "AND SHIP_NO = '" & FpSpread2.ActiveSheet.GetValue(FpSpread2.ActiveSheet.ActiveRowIndex, 1) & "'" & vbNewLine
                End If
                'If ComboBoxEx1.Text <> "ALL" Then
                '    QRY = QRY & "AND MODEL = '" & ComboBoxEx1.Text & "'" & vbNewLine
                'End If
                If ComboBoxEx2.Text <> "ALL" Then
                    QRY = QRY & "AND RETURN_DV = '" & ComboBoxEx2.Text & "'" & vbNewLine
                End If

                .ColumnFooter.Cells(0, 8).Text = Query_RS(QRY)


            End If

            Spread_AutoCol(FpSpread1)
            .Columns(6).Width = 60

        End With

        ''part summary
        'With FpSpread3.ActiveSheet
        '    .RowCount = 0

        '    For R_CNT As Integer = 0 To FpSpread2.ActiveSheet.RowCount - 1
        '        If FpSpread2.ActiveSheet.GetValue(R_CNT, FpSpread2.ActiveSheet.ColumnCount - 1) = True Then
        '            Dim QRY As String = "select b.svc_type, a.part_no, (SELECT PART_NAME FROM TBL_PARTMASTER WHERE PART_NO = a.part_no), sum(a.qty) as cnt," & vbNewLine
        '            QRY = QRY & "		round(ISNULL((SELECT PRICE FROM TBL_PARTMASTER WHERE PART_NO = a.PART_NO),0),2)," & vbNewLine
        '            QRY = QRY & "		case right(a.part_no,1)" & vbNewLine
        '            QRY = QRY & "		when 'K' then 0" & vbNewLine
        '            QRY = QRY & "		when 'C' then 0" & vbNewLine
        '            QRY = QRY & "		else round(ISNULL((SELECT PRICE FROM TBL_PARTMASTER WHERE PART_NO = a.PART_NO),0)*0.05,2)" & vbNewLine
        '            QRY = QRY & "		end as price, 0,0" & vbNewLine
        '            QRY = QRY & "from tbl_repairmaster_b a, tbl_esnmaster_b b" & vbNewLine
        '            QRY = QRY & "where b.ship_no = '" & FpSpread2.ActiveSheet.GetValue(R_CNT, 1) & "'" & vbNewLine
        '            QRY = QRY & "  and b.return_dv = 'GOOD'" & vbNewLine
        '            QRY = QRY & "  and a.src_no = b.obid" & vbNewLine
        '            QRY = QRY & "  and a.qty > 0" & vbNewLine
        '            QRY = QRY & "group by b.svc_type, a.part_no" & vbNewLine
        '            QRY = QRY & "order by b.svc_type, a.part_no, cnt desc" & vbNewLine
        '            QRY = QRY & "" & vbNewLine

        '            If Query_Spread_LTD_ROW(FpSpread3, QRY, 0, .ColumnCount - 1, .RowCount) = True Then
        '                'If Query_Spread(FpSpread3, QRY, 1) = True Then
        '                For I As Integer = 0 To .RowCount - 1
        '                    .Cells(I, 6).Value = CDec(.Cells(I, 4).Value) + CDec(.Cells(I, 5).Value)
        '                    .Cells(I, 7).Value = CDec(.Cells(I, 3).Value) * CDec(.Cells(I, 6).Value)
        '                Next
        '                Spread_AutoCol(FpSpread3)
        '            End If

        '        End If
        '    Next



        'End With



    End Sub

    Private Sub FindBtn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FindBtn.Click
        Dim i, i_qty As New Integer
        Dim m_qry As String = ""
        m_qry = MODEL_SELECTED(CheckedListBox1)

        FpSpread1.ActiveSheet.RowCount = 0
        FpSpread3.ActiveSheet.RowCount = 0
        With FpSpread2.ActiveSheet

            .RowCount = 0
            
            Dim QRY As String = ""

            QRY = QRY & "SELECT SHIP_DATE, SHIP_NO, SUM(QTY) , SUM(QTY*CHARGE) , RETURN_DV ," & vbNewLine
            QRY = QRY & "			ISNULL((SELECT SHIP_NO FROM TBL_INVSUMMARY WHERE SHIP_NO = A.SHIP_NO),''), " & vbNewLine
            QRY = QRY & "			ISNULL((SELECT INV_NO FROM TBL_INVSUMMARY WHERE SHIP_NO = A.SHIP_NO),''), " & vbNewLine
            QRY = QRY & "			ISNULL((SELECT INV_DATE FROM TBL_INVSUMMARY WHERE SHIP_NO = A.SHIP_NO),'') " & vbNewLine

            QRY = QRY & "FROM VIEW_SHIPSUMMARY A" & vbNewLine
            QRY = QRY & "WHERE SHIP_DATE BETWEEN CONVERT(VARCHAR(8), CONVERT(DATETIME,'" & DateTimeInput1.Text & "'),112) AND CONVERT(VARCHAR(8), CONVERT(DATETIME,'" & DateTimeInput2.Text & "'),112)" & vbNewLine


            If m_qry <> "" Then
                QRY = QRY & "AND MODEL IN (" & m_qry & ")" & vbNewLine
            End If

            'If ComboBoxEx1.Text <> "ALL" Then
            '    QRY = QRY & "AND MODEL = '" & ComboBoxEx1.Text & "'" & vbNewLine
            'End If

            If ComboBoxEx2.Text <> "ALL" Then
                QRY = QRY & "AND RETURN_DV = '" & ComboBoxEx2.Text & "'" & vbNewLine
            End If


            QRY = QRY & "GROUP BY SHIP_NO, SHIP_DATE , RETURN_DV " & vbNewLine
            QRY = QRY & "ORDER BY SHIP_NO, SHIP_DATE , RETURN_DV " & vbNewLine
            QRY = QRY & "" & vbNewLine

            If Query_Spread(FpSpread2, QRY, 1) = True Then
                .RowCount = FpSpread2.ActiveSheet.RowCount + 1
                .Rows(FpSpread2.ActiveSheet.RowCount - 1).BackColor = Color.Yellow
                .SetValue(FpSpread2.ActiveSheet.RowCount - 1, 0, "TOTAL")

                FpSpread2.AllowUserFormulas = True

                .Cells(.RowCount - 1, 2).Formula = "SUM(C1:C" & .RowCount - 1 & ")"
                .Cells(.RowCount - 1, 3).Formula = "SUM(D1:D" & .RowCount - 1 & ")"


                .ColumnFooterVisible = True
                .ColumnFooter.Cells(0, 0).Text = "출하합계"
                .ColumnFooter.Cells(0, 2).CellType = intcell
                .ColumnFooter.Cells(0, 2).Text = .GetValue(.RowCount - 1, 2)
                .ColumnFooter.Cells(0, 3).CellType = deccell
                .ColumnFooter.Cells(0, 3).Text = .GetValue(.RowCount - 1, 3)

                '.ColumnCount = .ColumnCount + 1
                '.ColumnHeader.Columns(.ColumnCount - 1).Label = "CHK"
                '.Columns(.ColumnCount - 1).CellType = CHKcell
                '.ColumnHeader.Columns(.ColumnCount - 1).HorizontalAlignment = FarPoint.Win.Spread.CellHorizontalAlignment.Center
                Spread_AutoCol(FpSpread2)
            End If

            DockContainerItem3.Selected = True
        End With


        With FpSpread3.ActiveSheet
            .RowCount = 0

            Dim QRY As String = ""

            QRY = QRY & "SELECT CASE WHEN (SELECT COUNT(MODEL_NO) FROM TBL_MODELMASTER WHERE MODEL_NO = A.MODEL) > 0 THEN A.MODEL ELSE (SELECT PART_NO + '('+PART_NAME +')' FROM TBL_PARTMASTER WHERE PART_NO = A.MODEL) END," & vbNewLine
            QRY = QRY & "		 ISNULL((SELECT SUM(QTY) FROM TBL_SHIPSUMMARY WHERE MODEL = A.MODEL AND RETURN_DV = A.RETURN_DV AND SHIP_DATE BETWEEN CONVERT(VARCHAR(8), CONVERT(DATETIME,'" & DateTimeInput1.Text & "'),112) AND CONVERT(VARCHAR(8), CONVERT(DATETIME,'" & DateTimeInput2.Text & "'),112)  ),0)," & vbNewLine
            QRY = QRY & "		 ISNULL((SELECT SUM(QTY*CHARGE) FROM TBL_SHIPSUMMARY WHERE MODEL = A.MODEL AND RETURN_DV = A.RETURN_DV  AND SHIP_DATE BETWEEN CONVERT(VARCHAR(8), CONVERT(DATETIME,'" & DateTimeInput1.Text & "'),112) AND CONVERT(VARCHAR(8), CONVERT(DATETIME,'" & DateTimeInput2.Text & "'),112)  ),0)," & vbNewLine
            QRY = QRY & "		 ISNULL((SELECT SUM(QTY) FROM TBL_SHIPSUMMARY_GOODS WHERE MODEL = A.MODEL AND RETURN_DV = A.RETURN_DV  AND SHIP_DATE BETWEEN CONVERT(VARCHAR(8), CONVERT(DATETIME,'" & DateTimeInput1.Text & "'),112) AND CONVERT(VARCHAR(8), CONVERT(DATETIME,'" & DateTimeInput2.Text & "'),112)  ),0)," & vbNewLine
            QRY = QRY & "		 ISNULL((SELECT SUM(QTY*CHARGE) FROM TBL_SHIPSUMMARY_GOODS WHERE MODEL = A.MODEL AND RETURN_DV = A.RETURN_DV  AND SHIP_DATE BETWEEN CONVERT(VARCHAR(8), CONVERT(DATETIME,'" & DateTimeInput1.Text & "'),112) AND CONVERT(VARCHAR(8), CONVERT(DATETIME,'" & DateTimeInput2.Text & "'),112)  ),0)," & vbNewLine
            QRY = QRY & "		 ISNULL((SELECT SUM(QTY) FROM VIEW_SHIPSUMMARY WHERE MODEL = A.MODEL AND RETURN_DV = A.RETURN_DV  AND SHIP_DATE BETWEEN CONVERT(VARCHAR(8), CONVERT(DATETIME,'" & DateTimeInput1.Text & "'),112) AND CONVERT(VARCHAR(8), CONVERT(DATETIME,'" & DateTimeInput2.Text & "'),112)  ),0)," & vbNewLine
            QRY = QRY & "		 ISNULL((SELECT SUM(QTY*CHARGE) FROM VIEW_SHIPSUMMARY WHERE MODEL = A.MODEL AND RETURN_DV = A.RETURN_DV AND SHIP_DATE BETWEEN CONVERT(VARCHAR(8), CONVERT(DATETIME,'" & DateTimeInput1.Text & "'),112) AND CONVERT(VARCHAR(8), CONVERT(DATETIME,'" & DateTimeInput2.Text & "'),112)  ),0)" & vbNewLine
            QRY = QRY & "FROM VIEW_SHIPSUMMARY   A" & vbNewLine
            QRY = QRY & "WHERE SHIP_NO IS NOT NULL" & vbNewLine
            'If ComboBoxEx1.Text <> "ALL" Then
            '    QRY = QRY & "AND MODEL = '" & ComboBoxEx1.Text & "'" & vbNewLine
            'End If

            If m_qry <> "" Then
                QRY = QRY & "AND MODEL IN (" & m_qry & ")" & vbNewLine
            End If

            If ComboBoxEx2.Text <> "ALL" Then
                QRY = QRY & "AND RETURN_DV = '" & ComboBoxEx2.Text & "'" & vbNewLine
            End If
            QRY = QRY & "GROUP BY MODEL , RETURN_DV " & vbNewLine
            QRY = QRY & "ORDER BY MODEL , RETURN_DV " & vbNewLine
            QRY = QRY & "" & vbNewLine

            If Query_Spread(FpSpread3, QRY, 1) = True Then
                .RowCount = .RowCount + 1
                .Rows(.RowCount - 1).BackColor = Color.Yellow
                .SetValue(.RowCount - 1, 0, "TOTAL")

                FpSpread3.AllowUserFormulas = True

                .Cells(.RowCount - 1, 1).Formula = "SUM(B1:B" & .RowCount - 1 & ")"
                .Cells(.RowCount - 1, 2).Formula = "SUM(C1:C" & .RowCount - 1 & ")"
                .Cells(.RowCount - 1, 3).Formula = "SUM(D1:D" & .RowCount - 1 & ")"
                .Cells(.RowCount - 1, 4).Formula = "SUM(E1:E" & .RowCount - 1 & ")"
                .Cells(.RowCount - 1, 5).Formula = "SUM(F1:F" & .RowCount - 1 & ")"
                .Cells(.RowCount - 1, 6).Formula = "SUM(G1:G" & .RowCount - 1 & ")"


                .ColumnFooterVisible = True
                .ColumnFooter.Cells(0, 0).Text = "출하합계"
                .ColumnFooter.Cells(0, 1).CellType = intcell
                .ColumnFooter.Cells(0, 1).Text = .GetValue(.RowCount - 1, 1)
                .ColumnFooter.Cells(0, 2).CellType = deccell
                .ColumnFooter.Cells(0, 2).Text = .GetValue(.RowCount - 1, 2)
                .ColumnFooter.Cells(0, 3).CellType = intcell
                .ColumnFooter.Cells(0, 3).Text = .GetValue(.RowCount - 1, 3)
                .ColumnFooter.Cells(0, 4).CellType = deccell
                .ColumnFooter.Cells(0, 4).Text = .GetValue(.RowCount - 1, 4)
                .ColumnFooter.Cells(0, 5).CellType = intcell
                .ColumnFooter.Cells(0, 5).Text = .GetValue(.RowCount - 1, 5)
                .ColumnFooter.Cells(0, 6).CellType = deccell
                .ColumnFooter.Cells(0, 6).Text = .GetValue(.RowCount - 1, 6)

                Spread_AutoCol(FpSpread3)
            End If

        End With



    End Sub

    Private Sub NewBtn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NewBtn.Click, NewBtn1.Click
    End Sub

    Private Sub SaveBtn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SaveBtn.Click, SaveBtn1.Click
        If MessageBox.Show("Are you sure to save Invoice Date?", "Save Invoice Date", MessageBoxButtons.YesNo) <> Windows.Forms.DialogResult.Yes Then
            Exit Sub
        End If

        With FpSpread2.ActiveSheet
            For i As Integer = 0 To .RowCount - 1
                If .GetValue(i, .ColumnCount - 1) = True Then
                    If Insert_Data("update tbl_esnmaster_b set startdate = '" & DateTimeInput3.Text & "' where ship_no = '" & .GetValue(i, 1) & "'") = True Then
                    Else
                        MessageBox.Show("Error on saving. Check Data!!")
                        Exit Sub
                    End If
                End If
            Next

        End With
        MessageBox.Show("저장되었습니다!")
        FindBtn_Click(sender, e)

    End Sub

    Private Sub XlsBtn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles XlsBtn.Click, XlsBtn1.Click

        If save_excel = "" Then
            MessageBox.Show("Select Spread for saving!!")
            Exit Sub
        End If

        If save_excel = "FpSpread1" Then
            If MessageBox.Show("Are you sure to make invoice ?", "Make Invoice", MessageBoxButtons.YesNo) = Windows.Forms.DialogResult.No Then
                Exit Sub
            End If

            Dim f As SaveFileDialog = SaveFileDialog1

            f.DefaultExt = "XLS"
            f.InitialDirectory = My.Computer.FileSystem.SpecialDirectories.MyDocuments
            f.FileName = Mid(FpSpread2.ActiveSheet.GetValue(FpSpread2.ActiveSheet.ActiveRowIndex, 1), 3, Len(FpSpread2.ActiveSheet.GetValue(FpSpread2.ActiveSheet.ActiveRowIndex, 1)) - 2) '& ".csv"
            f.Filter = "Microsoft Office Excel (*.xls)|*.xls*|All Files(*.*)|*.*"

            f.FileName = Mid(FpSpread2.ActiveSheet.GetValue(FpSpread2.ActiveSheet.ActiveRowIndex, 1), 3, Len(FpSpread2.ActiveSheet.GetValue(FpSpread2.ActiveSheet.ActiveRowIndex, 1)) - 2) '& ".csv"

            If FpSpread2.ActiveSheet.GetValue(FpSpread2.ActiveSheet.ActiveRowIndex, 5) = "INWTY" Then
                SaveFileDialog1.FileName = Mid(FpSpread2.ActiveSheet.GetValue(FpSpread2.ActiveSheet.ActiveRowIndex, 1), 12, 2) & "0" & Mid(FpSpread2.ActiveSheet.GetValue(FpSpread2.ActiveSheet.ActiveRowIndex, 1), 14, 5) '& ".csv"
            Else
                SaveFileDialog1.FileName = Mid(FpSpread2.ActiveSheet.GetValue(FpSpread2.ActiveSheet.ActiveRowIndex, 1), 6, 1) & Mid(FpSpread2.ActiveSheet.GetValue(FpSpread2.ActiveSheet.ActiveRowIndex, 1), 12, 6) '& ".csv"
            End If


            If f.ShowDialog() = Windows.Forms.DialogResult.OK Then
                With FpSpread1.ActiveSheet
                    .Protect = False
                    If FpSpread1.SaveExcel(f.FileName, FarPoint.Excel.ExcelSaveFlags.SaveCustomColumnHeaders) = True Then
                        f.FileName = ""
                        SaveFileDialog1.FileName = ""
                    End If
                End With
            End If
        ElseIf save_excel = "FpSpread2" Then
            SaveFileDialog1.FileName = ""
            File_Save(SaveFileDialog1, FpSpread2)
        ElseIf save_excel = "FpSpread3" Then
            SaveFileDialog1.FileName = ""
            File_Save(SaveFileDialog1, FpSpread3)
        Else
            MessageBox.Show("Select Spread for Save!!")
        End If

    End Sub

    Private Sub PrtBtn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PrtBtn.Click, PrtBtn1.Click

        Dim i As New Integer

        If save_excel = "" Then
            MessageBox.Show("Select Spread for Printing!!")
            Exit Sub
        End If

        If save_excel = "FpSpread1" Then
            If Spread_Print(Me.FpSpread1, "Repair Upload", 0) = False Then
                MsgBox("Fail to Print")
            End If
        ElseIf save_excel = "FpSpread2" Then
            If Spread_Print(Me.FpSpread2, "Shipping Summary", 0) = False Then
                MsgBox("Fail to Print")
            End If
        ElseIf save_excel = "FpSpread3" Then
            If Spread_Print(Me.FpSpread3, "Shipping Summary", 0) = False Then
                MsgBox("Fail to Print")
            End If
        Else
            MessageBox.Show("No Printing Object!!")
        End If

    End Sub

    Private Sub FpSpread2_CellClick(ByVal sender As System.Object, ByVal e As FarPoint.Win.Spread.CellClickEventArgs) Handles FpSpread2.CellClick
        save_excel = "FpSpread2"
    End Sub

    Private Sub FpSpread1_CellClick(ByVal sender As System.Object, ByVal e As FarPoint.Win.Spread.CellClickEventArgs) Handles FpSpread1.CellClick
        save_excel = "FpSpread1"
    End Sub

    Private Sub FpSpread3_CellClick(ByVal sender As System.Object, ByVal e As FarPoint.Win.Spread.CellClickEventArgs) Handles FpSpread3.CellClick
        save_excel = "FpSpread3"
    End Sub

    Private Sub ButtonItem7_Click(ByVal sender As System.Object, ByVal e As FarPoint.Win.Spread.CellClickEventArgs) Handles FpSpread2.CellDoubleClick

    End Sub

    Private Sub TextBoxDropDown1_LEAVE(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBoxDropDown1.Leave

    End Sub


    Private Sub CheckedListBox1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles CheckedListBox1.LostFocus
        TextBoxDropDown1.Text = MODEL_SELECTED(CheckedListBox1)
    End Sub
End Class