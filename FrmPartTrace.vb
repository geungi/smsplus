Public Class FrmPartTrace
    Const CtrlMask As Byte = 8
    Private NowPoNo As String = ""
    Private TempSno As String = ""
    Private TotRec As Integer = 0
    Private TotRec2 As Integer = 0
    Private FwhCd As String = ""
    Private TwhCd As String = ""
    Private EmpNmTXT As String = ""
    Private PartAuth As Boolean = True


    Private Sub FrmPartTrace_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Condi_Disp() '콤보박스의 조건데이터 출력

        Me.ContextMenuBar1.SetContextMenuEx(Me.FpSpread1, CtxSp)

        With FpSpread1.ActiveSheet
            .Columns(2).CellType = datecell
        End With

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
    End Sub

    Private Sub FindBtn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FindBtn.Click
        disp_io()
    End Sub

    Private Sub PrtBtn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PrtBtn1.Click, PrtBtn.Click, bPrint.Click
        Dim S1 = Me.FpSpread1.ActiveSheet
        If S1.RowCount > 0 Then
            If Spread_Print(Me.FpSpread1, "Part I/O Trace", 1) = False Then
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
            Dim i, j, k, l As Integer
            Dim trsRs As ADODB.Recordset
            Dim S1 = Me.FpSpread1.ActiveSheet
            Dim oldPartNo As String


            S1.RowCount = 0

            If S1.RowCount > 0 Then
                S1.Rows.Remove(0, TotRec)
            End If
            '@SITE_ID	VARCHAR(10),
            '@WH_CD	VARCHAR(5),
            '@ST_DATE	DATETIME,
            '@ED_DATE	DATETIME,
            '@PART_NO	VARCHAR(15)

            QueryPo = "EXEC SP_FRMPARTTRACE '" & Site_id & "', 'E1000', '" & Me.POStDate.Text & "','" & Me.POEdDate.Text & "','" & Me.PartNoTxt.Text & "'"

            trsRs = Query_RS_ALL(QueryPo)

            If trsRs Is Nothing Then
                MessageBox.Show("조회할 데이터가 없습니다 !!", "FIND")
                Exit Sub
            End If
            oldPartNo = ""
            k = 0
            l = 0
            TotRec = trsRs.RecordCount
            MainFrm.ProgressBarItem1.Maximum = trsRs.RecordCount
            For i = 0 To trsRs.RecordCount - 1
                S1.RowCount += 1
                S1.Rows(S1.RowCount - 1).Locked = True
                For j = 0 To S1.ColumnCount - 1
                    S1.SetValue(i, j, trsRs(j).Value)
                Next
                Spread_NumType(FpSpread1, i, 3)
                Spread_NumType(FpSpread1, i, 5)
                '               Spread_NumType(FpSpread1, i, 7)

                If S1.GetValue(i, 0) <> oldPartNo Then
                    If oldPartNo <> "" Then
                        S1.AddSpanCell(l, 0, k + 1, 1)
                        S1.AddSpanCell(l, 1, k + 1, 1)
                        k = 0
                    End If
                    '                  tot = S1.GetValue(i, 7) + S1.GetValue(i, 3) - S1.GetValue(i, 5)
                    '                 S1.SetValue(i, 7, tot)
                    oldPartNo = S1.GetValue(i, 0)
                    l = i
                Else
                    '                tot = tot + S1.GetValue(i, 3) - S1.GetValue(i, 5)
                    '               S1.SetValue(i, 7, tot)
                    k += 1
                End If
                If i = trsRs.RecordCount - 1 Then
                    S1.AddSpanCell(l, 0, k + 1, 1)
                    S1.AddSpanCell(l, 1, k + 1, 1)
                    '             tot = tot + S1.GetValue(i, 3) - S1.GetValue(i, 5)
                    '              S1.SetValue(i, 7, tot)
                End If

                With FpSpread1.ActiveSheet
                    If .GetValue(i, 3) = 0 Then
                        .SetText(i, 3, "")
                        .SetText(i, 4, "")
                    End If

                    If .GetValue(i, 5) = 0 Then
                        .SetText(i, 5, "")
                        .SetText(i, 6, "")
                    End If

                End With

                MainFrm.ProgressBarItem1.Value = i
                trsRs.MoveNext()
            Next


            With FpSpread1.ActiveSheet
                If .RowCount > 0 Then
                    .RowCount = .RowCount + 1
                    .Rows(.RowCount - 1).BackColor = Color.Yellow
                    .Cells(.RowCount - 1, 0).Text = "TOTAL"
                    .Cells(.RowCount - 1, 1).Text = PartNoTxt.Text & "(" & POStDate.Text & "~" & POEdDate.Text & ")"

                    .Cells(.RowCount - 1, 3).CellType = intcell
                    .Cells(.RowCount - 1, 3).Formula = "SUM(D1:D" & .RowCount - 1 & ")"

                    .Cells(.RowCount - 1, 5).CellType = intcell
                    .Cells(.RowCount - 1, 5).Formula = "SUM(F1:F" & .RowCount - 1 & ")"
                End If
            End With

            Spread_AutoCol(FpSpread1)

        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message, "ERROR")
        End Try
    End Sub


End Class