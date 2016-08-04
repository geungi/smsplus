Public Class FrmGoal

    Private Sub CheckedListBox1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles CheckedListBox1.LostFocus
        TextBoxDropDown1.Text = MODEL_SELECTED(CheckedListBox1)
    End Sub


    Private Sub FrmGoal_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Dim G_intcell As New FarPoint.Win.Spread.CellType.NumberCellType()

        G_intcell.DecimalPlaces = 0
        G_intcell.Separator = ","
        G_intcell.ShowSeparator = True

        If Spread_Setting(FpSpread1, Me.Name) = True Then
            FpSpread1.ActiveSheet.RowCount = 0
            Spread_AutoCol(FpSpread1)
        End If

        If Spread_Setting(FpSpread2, Me.Name) = True Then
            FpSpread2.ActiveSheet.RowCount = 0
            Spread_AutoCol(FpSpread2)
        End If

        'If Query_Combo(Me.ComboBoxEx2, "SELECT model_no FROM tbl_modelmaster WHERE site_id = '" & Site_id & "' and active='Y' ORDER BY model_no") = True Then
        '    ComboBoxEx2.Items.Add("ALL")
        '    ComboBoxEx2.Text = "ALL"
        'End If
        If Query_CheckList(Me.CheckedListBox1, "SELECT model_no FROM tbl_modelmaster WHERE site_id = '" & Site_id & "' and active = 'Y' ORDER BY model_no") = True Then
            Me.CheckedListBox1.Items.Add("ALL")
            Me.CheckedListBox1.Text = "ALL"
        End If

        If Query_Combo(Me.ComboBoxEx1, "SELECT CODE_NAME FROM tbl_CODEmaster WHERE site_id = '" & Site_id & "' AND CLASS_ID = 'R0008' and active='Y' ORDER BY CODE_ID") = True Then
            ComboBoxEx1.Items.Add("ALL")
            ComboBoxEx1.Text = "ALL"
        End If

        DockContainerItem2.Text = "실행 메뉴"
        DockContainerItem1.Text = "입력 조건"
        DockContainerItem3.Text = "USA 생산 목표"

        POStDate.Value = Now
        DateTimeInput1.Value = Now


        HD_SET()

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

    End Sub


    Private Sub ButtonItem2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonItem2.Click, NewBtn1.Click


        Dim m_qry As String = ""
        m_qry = MODEL_SELECTED(CheckedListBox1)

        With FpSpread1.ActiveSheet
            .ClearRange(0, 0, .RowCount, .ColumnCount, False)
            .RowCount = 1
            .Cells(0, 0, .RowCount - 1, .ColumnCount - 1).Locked = False
            .RowCount = 0

            Dim QRY As String = ""
            QRY = QRY & "SELECT " & vbNewLine

            If ComboBoxEx1.Text = "ALL" Then
                QRY = QRY & "B.CODE_NAME," & vbNewLine
            Else
                QRY = QRY & "'" & ComboBoxEx1.Text & "', " & vbNewLine
            End If


            QRY = QRY & "MODEL_NO, " & vbNewLine

            If .ColumnCount > 3 Then
                For I As Integer = 0 To .ColumnCount - 4
                    QRY = QRY & "ISNULL((SELECT QTY FROM TBL_PGOAL WHERE P_DATE = '" & .ColumnHeader.Columns(I + 2).Label & "' AND MODEL = A.MODEL_NO AND DIV = B.CODE_ID ),0)," & vbNewLine
                Next
            End If
            QRY = QRY & "ISNULL((SELECT SUM(QTY) FROM TBL_PGOAL WHERE P_DATE BETWEEN '" & .ColumnHeader.Columns(2).Label & "' AND '" & .ColumnHeader.Columns(.ColumnCount - 2).Label & "' AND MODEL = A.MODEL_NO AND DIV = B.CODE_ID ),0)" & vbNewLine

            QRY = QRY & "FROM TBL_MODELMASTER A, TBL_CODEMASTER B" & vbNewLine
            QRY = QRY & "WHERE A.ACTIVE = 'Y'" & vbNewLine

            'If ComboBoxEx2.Text <> "ALL" Then
            '    QRY = QRY & "  AND A.MODEL_NO = '" & ComboBoxEx2.Text & "'" & vbNewLine
            'End If
            If m_qry <> "" Then
                QRY = QRY & "AND A.MODEL_NO IN (" & m_qry & ")" & vbNewLine
            End If

            QRY = QRY & "  AND B.CLASS_ID = 'R0008'" & vbNewLine

            If ComboBoxEx1.Text <> "ALL" Then
                QRY = QRY & "  AND B.CODE_NAME = '" & ComboBoxEx1.Text & "'" & vbNewLine
            End If

            QRY = QRY & "ORDER BY A.MODEL_NO, B.CODE_NAME" & vbNewLine


            QRY = QRY & "" & vbNewLine


            If Query_Spread(FpSpread1, QRY, 1) = True Then
                .Cells(0, 2, .RowCount - 1, .ColumnCount - 1).ForeColor = Color.OrangeRed


                .RowCount = .RowCount + 1
                .Rows(.RowCount - 1).BackColor = Color.Yellow
                .Rows(.RowCount - 1).ForeColor = Color.Black
                .Rows(.RowCount - 1).Locked = True
                .SetValue(.RowCount - 1, 0, "TOTAL")

                SPREAD_ROW_TOTAL(FpSpread1, 2, .ColumnCount - 1, 1)
                Spread_AutoCol(FpSpread1)
            End If
        End With



    End Sub

    Private Sub FpSpread1_Change(ByVal sender As Object, ByVal e As FarPoint.Win.Spread.ChangeEventArgs) Handles FpSpread1.Change

        save_excel = "FpSpread1"

        'If e.Column = 9 Then
        '    With Me.FpSpread1.ActiveSheet
        '        If .ActiveRowIndex = (.RowCount - 1) Then
        '            .SetActiveCell(.ActiveRowIndex - 1, .ActiveColumnIndex)

        '            .SetValue(.RowCount - 1, 10, CInt(.GetValue(.RowCount - 1, 9)) * 0.5)
        '            .SetValue(.RowCount - 1, 11, CInt(.GetValue(.RowCount - 1, 9)) * 0.3)
        '            .SetValue(.RowCount - 1, 12, CInt(.GetValue(.RowCount - 1, 9)) * 0.2)
        '            .Rows(.RowCount - 1).ForeColor = Color.OrangeRed
        '        Else
        '            .SetActiveCell(.ActiveRowIndex + 1, .ActiveColumnIndex)

        '            .SetValue(.ActiveRowIndex - 1, 10, CInt(.GetValue(.ActiveRowIndex - 1, 9)) * 0.5)
        '            .SetValue(.ActiveRowIndex - 1, 11, CInt(.GetValue(.ActiveRowIndex - 1, 9)) * 0.3)
        '            .SetValue(.ActiveRowIndex - 1, 12, CInt(.GetValue(.ActiveRowIndex - 1, 9)) * 0.2)
        '            .Rows(.ActiveRowIndex - 1).ForeColor = Color.OrangeRed

        '        End If

        '        '.RemoveRows(.RowCount - 1, 1)
        '        '.RowCount = .RowCount + 1
        '        '.Rows(.RowCount - 1).BackColor = Color.Yellow
        '        '.Rows(.RowCount - 1).Locked = True
        '        '.SetValue(.RowCount - 1, 0, "TOTAL")


        '        If .Rows(.RowCount - 2).BackColor = Color.Yellow Then
        '            .Cells(.RowCount - 2, 9, .RowCount - 2, 12).Value = 0
        '            SPREAD_ROW_TOTAL_LTD(FpSpread1, 9, 12, .RowCount - 2)
        '        Else
        '            .Cells(.RowCount - 1, 9, .RowCount - 1, 12).Value = 0
        '            SPREAD_ROW_TOTAL(FpSpread1, 9, 12, 1)

        '        End If
        '    End With
        With FpSpread1.ActiveSheet
            .Cells(e.Row, e.Column).ForeColor = Color.OrangeRed
        End With


        'End If

    End Sub


    Private Sub ButtonItem3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonItem3.Click, SaveBtn1.Click
        Dim i, J As Integer

        With FpSpread1.ActiveSheet
            For i = 0 To .RowCount - 1
                For J = 2 To .ColumnCount - 2
                    If .Cells(i, J).ForeColor = Color.OrangeRed Then

                        Dim QRY As String = ""
                        QRY = QRY & "IF EXISTS (SELECT MODEL FROM TBL_PGOAL WHERE P_DATE = '" & .ColumnHeader.Columns(J).Label & "' AND MODEL = '" & .GetValue(i, 1) & "' AND DIV = (SELECT CODE_ID FROM TBL_CODEMASTER WHERE CLASS_ID = 'R0008' AND CODE_NAME = '" & .GetValue(i, 0) & "'))" & vbNewLine
                        QRY = QRY & "BEGIN" & vbNewLine
                        QRY = QRY & "   UPDATE TBL_PGOAL SET QTY = " & .GetValue(i, J) & ", U_PERSON = '" & Emp_No & "', U_DATE = GETDATE() WHERE P_DATE = '" & .ColumnHeader.Columns(J).Label & "' AND MODEL = '" & .GetValue(i, 1) & "' AND DIV = (SELECT CODE_ID FROM TBL_CODEMASTER WHERE CLASS_ID = 'R0008' AND CODE_NAME = '" & .GetValue(i, 0) & "')" & vbNewLine
                        QRY = QRY & "END" & vbNewLine
                        QRY = QRY & "ELSE" & vbNewLine
                        QRY = QRY & "BEGIN" & vbNewLine
                        QRY = QRY & "   INSERT INTO TBL_PGOAL VALUES ('" & Site_id & "','" & .ColumnHeader.Columns(J).Label & "','" & .GetValue(i, 1) & "','" & Query_RS("SELECT CODE_ID FROM TBL_CODEMASTER WHERE CLASS_ID = 'R0008' AND CODE_NAME = '" & .GetValue(i, 0) & "'") & "'," & .GetValue(i, J) & ",'" & Emp_No & "', GETDATE(), '" & Emp_No & "',GETDATE(), NULL,NULL)" & vbNewLine
                        QRY = QRY & "END" & vbNewLine

                        If Insert_Data(QRY) = True Then
                            .Cells(i, J).ForeColor = Color.Black
                        Else
                            Modal_Error("저장 중 오류 발생.")
                            Exit Sub
                        End If
                    End If
                Next
            Next
        End With

        MessageBox.Show("Complete to Save!!")
        FindBtn_Click(sender, e)

    End Sub

    Function HD_SET() As Boolean

        Dim I As Integer = Query_RS("SELECT DATEDIFF(DAY, '" & POStDate.Text & "','" & DateTimeInput1.Text & "')")

        With FpSpread1.ActiveSheet
            .RowCount = 0
            .ColumnCount = 2

            For I = 0 To I
                .ColumnCount = .ColumnCount + 1
                .ColumnHeader.Columns(.ColumnCount - 1).Label = Query_RS("select REPLACE(CONVERT(VARCHAR(10), DATEADD(DAY," & I & ",'" & POStDate.Text & "'), 102),'.','-')")
                .Columns(.ColumnCount - 1).CellType = intcell
            Next
            .ColumnCount = .ColumnCount + 1
            .ColumnHeader.Columns(.ColumnCount - 1).Label = "합계"
            .Columns(.ColumnCount - 1).CellType = intcell
        End With

        Spread_AutoCol(FpSpread1)


        I = Query_RS("SELECT DATEDIFF(week, '" & POStDate.Text & "','" & DateTimeInput1.Text & "')")
        Dim mon As String = Query_RS("select convert(varchar(10), ((Convert(datetime, '" & POStDate.Text & "')) - DatePart(dw,  '" & POStDate.Text & "') + 2), 101)")

        With FpSpread2.ActiveSheet
            .RowCount = 0
            .ColumnCount = 2

            For I = 0 To I
                .ColumnCount = .ColumnCount + 1
                .ColumnHeader.Columns(.ColumnCount - 1).Label = Query_RS("select REPLACE(convert(varchar(10), ((Convert(datetime, '" & POStDate.Text & "')) - DatePart(dw,  '" & POStDate.Text & "') + 2)+(" & I & "*7) , 102) ,'.','-')")
                .Columns(.ColumnCount - 1).CellType = intcell
            Next
            .ColumnCount = .ColumnCount + 1
            .ColumnHeader.Columns(.ColumnCount - 1).Label = "합계"
            .Columns(.ColumnCount - 1).CellType = intcell
        End With

        Spread_AutoCol(FpSpread2)




    End Function

    Private Sub FindBtn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FindBtn.Click
        Dim m_qry As String = ""
        m_qry = MODEL_SELECTED(CheckedListBox1)

        With FpSpread1.ActiveSheet
            HD_SET()

            .ClearRange(0, 0, .RowCount, .ColumnCount, False)
            .RowCount = 1
            .Cells(0, 0, .RowCount - 1, .ColumnCount - 1).Locked = False
            .RowCount = 0

            Dim QRY As String = ""
            QRY = QRY & "SELECT  distinct (SELECT CODE_NAME FROM tbl_codemaster WHERE CLASS_ID = 'R0008' AND CODE_ID = A.DIV) AS DV,model," & vbNewLine


            If .ColumnCount > 3 Then
                For I As Integer = 0 To .ColumnCount - 4
                    QRY = QRY & "ISNULL((SELECT QTY FROM TBL_PGOAL WHERE P_DATE = '" & .ColumnHeader.Columns(I + 2).Label & "' AND MODEL = A.MODEL AND DIV = a.div ),0)," & vbNewLine
                Next
            End If
            QRY = QRY & "ISNULL((SELECT SUM(QTY) FROM TBL_PGOAL WHERE P_DATE BETWEEN '" & .ColumnHeader.Columns(2).Label & "' AND '" & .ColumnHeader.Columns(.ColumnCount - 2).Label & "' AND MODEL = A.MODEL AND DIV = a.div ),0)" & vbNewLine

            QRY = QRY & "FROM TBL_PGOAL a" & vbNewLine
            QRY = QRY & "where P_DATE BETWEEN '" & .ColumnHeader.Columns(2).Label & "' AND '" & .ColumnHeader.Columns(.ColumnCount - 2).Label & "'" & vbNewLine

            'If ComboBoxEx2.Text <> "ALL" Then
            '    QRY = QRY & "  AND A.MODEl = '" & ComboBoxEx2.Text & "'" & vbNewLine
            'End If

            If m_qry <> "" Then
                QRY = QRY & "AND A.MODEL IN (" & m_qry & ")" & vbNewLine
            End If

            If ComboBoxEx1.Text <> "ALL" Then
                QRY = QRY & "  AND A.DIV = (SELECT CODE_ID FROM tbl_codemaster WHERE CLASS_ID = 'R0008' AND CODE_NAME = '" & ComboBoxEx1.Text & "')" & vbNewLine
            End If

            QRY = QRY & "GROUP BY A.MODEL, A.DIV,A.P_DATE" & vbNewLine
            QRY = QRY & "ORDER BY A.MODEL, DV" & vbNewLine


            QRY = QRY & "" & vbNewLine


            If Query_Spread(FpSpread1, QRY, 1) = True Then

                .RowCount = .RowCount + 1
                .Rows(.RowCount - 1).BackColor = Color.Yellow
                .Rows(.RowCount - 1).ForeColor = Color.Black
                .Rows(.RowCount - 1).Locked = True
                .SetValue(.RowCount - 1, 0, "TOTAL")

                SPREAD_ROW_TOTAL(FpSpread1, 2, .ColumnCount - 1, 1)
                Spread_AutoCol(FpSpread1)
            End If
        End With

        With FpSpread2.ActiveSheet
            .ClearRange(0, 0, .RowCount, .ColumnCount, False)
            .RowCount = 1
            .Cells(0, 0, .RowCount - 1, .ColumnCount - 1).Locked = False
            .RowCount = 0

            Dim QRY As String = ""
            QRY = QRY & "SELECT  distinct '' AS DV,model," & vbNewLine
            If .ColumnCount > 3 Then
                For I As Integer = 0 To .ColumnCount - 4
                    QRY = QRY & "ISNULL((SELECT tgt_QTY FROM TBL_PGOAL WHERE P_DATE = '" & .ColumnHeader.Columns(I + 2).Label & "' AND MODEL = A.MODEL ),0)," & vbNewLine
                Next
            End If
            QRY = QRY & "0" & vbNewLine
            '            QRY = QRY & "ISNULL((SELECT SUM(tgt_QTY) FROM TBL_PGOAL WHERE P_DATE BETWEEN '" & .ColumnHeader.Columns(2).Label & "' AND '" & .ColumnHeader.Columns(.ColumnCount - 2).Label & "' AND MODEL = A.MODEL ),0)" & vbNewLine

            QRY = QRY & "FROM TBL_PGOAL a" & vbNewLine
            QRY = QRY & "where P_DATE BETWEEN '" & .ColumnHeader.Columns(2).Label & "' AND '" & .ColumnHeader.Columns(.ColumnCount - 2).Label & "'" & vbNewLine
            QRY = QRY & "  AND tgt_qty > 0" & vbNewLine

            If ComboBoxEx2.Text <> "ALL" Then
                QRY = QRY & "  AND A.MODEl = '" & ComboBoxEx2.Text & "'" & vbNewLine
            End If
            QRY = QRY & "GROUP BY A.MODEL, A.P_DATE" & vbNewLine
            QRY = QRY & "ORDER BY A.MODEL, DV" & vbNewLine


            QRY = QRY & "" & vbNewLine

            If Query_Spread_USA(FpSpread2, QRY, 1) = True Then
                .Cells(0, 0, .RowCount - 1, 0).Text = "주간목표"

                .RowCount = .RowCount + 1
                .Rows(.RowCount - 1).BackColor = Color.Yellow
                .Rows(.RowCount - 1).ForeColor = Color.Black
                .Rows(.RowCount - 1).Locked = True
                .SetValue(.RowCount - 1, 0, "TOTAL")

                SPREAD_ROW_TOTAL(FpSpread2, 2, .ColumnCount - 1, 1)
                SPREAD_COL_TOTAL(FpSpread2, 2, .ColumnCount - 2, 1)

                Spread_AutoCol(FpSpread2)
            End If

        End With




    End Sub

    Private Sub ButtonItem4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonItem4.Click, DelBtn1.Click

    End Sub

    Private Sub ButtonItem5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonItem5.Click, PrtBtn1.Click
        If save_excel = "" Then
            MessageBox.Show("Select Spread for Printing!!")
            Exit Sub
        End If

        If save_excel = "FpSpread1" Then
            If Spread_Print(Me.FpSpread1, "Production Planning (" & FpSpread1.ActiveSheet.ColumnHeader.Cells(0, 7).Text & ")", 0) = False Then
                MsgBox("Fail to Print")
            End If
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
        Else
            MessageBox.Show("Select Spread for Saving!!")
        End If

    End Sub

    Private Sub FpSpread1_CellClick(ByVal sender As System.Object, ByVal e As FarPoint.Win.Spread.CellClickEventArgs) Handles FpSpread1.CellClick
        save_excel = "FpSpread1"
    End Sub




End Class