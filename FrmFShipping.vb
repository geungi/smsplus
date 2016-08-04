Public Class FrmFShipping

    Private PkRow As New Integer
    Private p_cnt, s_cnt As New Integer

    Private Sub FrmfShipping_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        DockContainerItem1.Text = "Form Menu"
        DockContainerItem2.Text = "Retrieve Condition"
        DockContainerItem3.Text = "Shipping Summary"
        DockContainerItem4.Text = "Waiting BOX Summary"
        DockContainerItem5.Text = "Packing List"
        
        ButtonItem12.Visible = False

        If Emp_No = "10001" Or Emp_No = "M0007" Then
            ButtonItem13.Visible = True
            CheckBoxX1.Visible = True
        End If

        Bar5.AutoHide = True
        Bar6.Visible = False

        If Spread_Setting(FpSpread1, "FrmFShipping") = True Then
            FpSpread1.ActiveSheet.RowCount = 0
            Spread_AutoCol(FpSpread1)
        End If

        If Spread_Setting(FpSpread2, "FrmShipping") = True Then
            FpSpread2.ActiveSheet.RowCount = 0
            Spread_AutoCol(FpSpread2)
        End If

        DateTimeInput1.Value = Now
        DateTimeInput2.Value = Now

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

        SaveBtn.Enabled = False
        SaveBtn1.Enabled = False

        DockContainerItem3.Selected = True

    End Sub

    Private Sub ListViewEx1_ItemChecked(ByVal sender As Object, ByVal e As System.Windows.Forms.ItemCheckedEventArgs) Handles ListViewEx1.ItemChecked
        Dim I As New Integer

        If e.Item.Checked = True Then
            If e.Item.Text = "TOTAL" Then
                For I = 0 To ListViewEx1.Items.Count - 2
                    ListViewEx1.Items(I).Checked = True
                Next
            Else
                If Query_Spread_LTD(FpSpread1, "EXEC SP_FRMFSHIPPING_GETSHIPPINGINFO '" & Site_id & "','" & e.Item.SubItems(3).Text & "','1'", 0, 7) = True Then
                    FpSpread1.ActiveSheet.Rows(0, FpSpread1.ActiveSheet.RowCount - 1).ForeColor = Color.OrangeRed
                    Spread_AutoCol(FpSpread1)
                End If
            End If
        Else
            If e.Item.Text = "TOTAL" Then
                For I = 0 To ListViewEx1.Items.Count - 2
                    ListViewEx1.Items(I).Checked = False
                Next
            Else
                For I = 0 To FpSpread1.ActiveSheet.RowCount - 1
                    If e.Item.SubItems(3).Text = FpSpread1.ActiveSheet.GetValue(I, 3) Then
                        FpSpread1.ActiveSheet.RemoveRows(I, e.Item.SubItems(4).Text)
                        Exit For
                    End If
                Next
            End If
        End If

    End Sub


    Private Sub ButtonItem7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonItem7.Click, FpSpread2.CellDoubleClick

        FpSpread1.ActiveSheet.RowCount = 0

        FpSpread1.ActiveSheet.FrozenColumnCount = 3

        If FpSpread2.ActiveSheet.RowCount = 0 Then
            MessageBox.Show("No Selected shipping no")
            Exit Sub
        End If

        If FpSpread2.ActiveSheet.Rows(FpSpread2.ActiveSheet.ActiveRowIndex).ForeColor = Color.OrangeRed Then
            MessageBox.Show("No Shipped shipping no")
        Else
            MainFrm.ProgressBarItem1.ProgressType = DevComponents.DotNetBar.eProgressItemType.Marquee

            If Query_Spread(FpSpread1, "exec SP_FRMFSHIPPING_GETSHIPPEDDTL '" & Site_id & "','" & FpSpread2.ActiveSheet.GetValue(FpSpread2.ActiveSheet.ActiveRowIndex, 1) & "','1'", 1) = True Then
                '              Spread_AutoCol(FpSpread1)
                LabelX4.Text = FpSpread2.ActiveSheet.GetValue(FpSpread2.ActiveSheet.ActiveRowIndex, 1)

            End If
            MainFrm.ProgressBarItem1.ProgressType = DevComponents.DotNetBar.eProgressItemType.Standard
            MessageBox.Show("조회 완료되었습니다.")

        End If
    End Sub

    Private Sub FindBtn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FindBtn.Click
        Dim i, i_qty As New Integer

        LabelX4.Text = ""
        FpSpread1.ActiveSheet.RowCount = 0
        FpSpread2.ActiveSheet.RowCount = 0

        For i = 0 To ListViewEx1.Items.Count - 1
            ListViewEx1.Items(i).Checked = False
        Next

        ListViewEx1.CheckBoxes = False

        Bar5.AutoHide = True

        If Query_Spread(FpSpread2, "EXEC SP_FRMFSHIPPING_GETSHIPPEDSUMMARY '" & Site_id & "','" & DateTimeInput1.Text & "','" & DateTimeInput2.Text & "'", 1) = True Then
            FpSpread2.ActiveSheet.RowCount = FpSpread2.ActiveSheet.RowCount + 1
            FpSpread2.ActiveSheet.Rows(FpSpread2.ActiveSheet.RowCount - 1).BackColor = Color.Yellow
            FpSpread2.ActiveSheet.SetValue(FpSpread2.ActiveSheet.RowCount - 1, 0, "TOTAL")

            FpSpread2.AllowUserFormulas = True

            FpSpread2.ActiveSheet.Cells(FpSpread2.ActiveSheet.RowCount - 1, 2).Formula = "SUM(C1:C" & FpSpread2.ActiveSheet.RowCount - 1 & ")"

            Spread_AutoCol(FpSpread2)
        End If


        If Query_Listview(ListViewEx1, "exec SP_FRMFshipping_GETCURBOX '" & Site_id & "', 'all'", True) = True Then

            For i = 0 To ListViewEx1.Items.Count - 1
                i_qty = i_qty + CInt(ListViewEx1.Items(i).SubItems(4).Text)
            Next

            ListViewEx1.Items.Add("TOTAL")
            ListViewEx1.Items(ListViewEx1.Items.Count - 1).SubItems.Add("")
            ListViewEx1.Items(ListViewEx1.Items.Count - 1).SubItems.Add("")
            ListViewEx1.Items(ListViewEx1.Items.Count - 1).SubItems.Add("")
            ListViewEx1.Items(ListViewEx1.Items.Count - 1).SubItems.Add(i_qty)
            ListViewEx1.Items(ListViewEx1.Items.Count - 1).BackColor = Color.Yellow
            ListViewEx1.Items(ListViewEx1.Items.Count - 1).SubItems(1).BackColor = Color.Yellow
            ListViewEx1.Items(ListViewEx1.Items.Count - 1).SubItems(2).BackColor = Color.Yellow
            ListViewEx1.Items(ListViewEx1.Items.Count - 1).SubItems(3).BackColor = Color.Yellow
            ListViewEx1.Items(ListViewEx1.Items.Count - 1).SubItems(4).BackColor = Color.Yellow


            ListViewEx1.AutoResizeColumn(0, ColumnHeaderAutoResizeStyle.ColumnContent)
            'Bar4.AutoHide = False
        End If


        SaveBtn.Enabled = False
        SaveBtn1.Enabled = False

        DockContainerItem3.Selected = True

    End Sub

    Private Sub ButtonItem8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonItem8.Click
        Try


            Dim I As New Integer
            Dim PK_RS As New ADODB.Recordset

            If FpSpread2.ActiveSheet.Rows(FpSpread2.ActiveSheet.ActiveRowIndex).ForeColor = Color.OrangeRed Then
                MessageBox.Show("No Shipped shipping no")
            Else
                MainFrm.ProgressBarItem1.ProgressType = DevComponents.DotNetBar.eProgressItemType.Marquee

                Dim S2 = Me.FpSpread2.ActiveSheet
                Dim S3 = Me.FpSpread3_Sheet1
                Dim S4 = Me.FpSpread3_Sheet2
                Dim S4_1 = Me.FpSpread5.ActiveSheet
                Dim S5 = Me.FpSpread3_Sheet3
                Dim S5_1 = Me.FpSpread6.ActiveSheet
                Dim pmd1 As String = ""
                Dim pmd2 As String = ""
                Dim pmd3 As String = ""
                Dim j, k, l1, l2, l3, m As Integer

                S3.Visible = False
                S4.Visible = False
                S5.Visible = False

                k = 0
                l1 = 0
                l2 = 0
                l3 = 0

                'S3.Cells(1, 11).Text = CDate(Mid(S2.Cells(S2.ActiveRowIndex, 0).Text, 1, 4) & "-" & Mid(S2.Cells(S2.ActiveRowIndex, 0).Text, 5, 2) & "-" & Mid(S2.Cells(S2.ActiveRowIndex, 0).Text, 7, 2))

                'Dim aaa As New System.Globalization.GregorianCalendar
                'aaa.CalendarType = Globalization.GregorianCalendarTypes.USEnglish

                ''                Dim bbb As New FarPoint.Win.Spread.CellType.DateTimeFormat

                'datecell.Calendar = aaa  ' "Gregorian(USEnglish)"
                '                datecell.SetCalendarFormat("Gregorian(USEnglish)")
                'S3.Cells(1, 11).CellType = datecell
                'datecell.DateTimeFormat = FarPoint.Win.Spread.CellType.DateTimeFormat.UserDefined
                'S3.Cells(1, 11).Value = Now.Date
                S3.Cells(1, 11).Value = Query_RS("select datename(month, '" & Now.Date & "') + ' ' + convert(varchar(2),datepart(dd, '" & Now.Date & "')) + ', ' + convert(varchar(4),datepart(yyyy, '" & Now.Date & "'))")


                S4.Cells(1, 11).Text = S3.Cells(1, 11).Text
                S5.Cells(1, 11).Text = S3.Cells(1, 11).Text
                S4_1.Cells(1, 11).Text = S3.Cells(1, 11).Text
                S5_1.Cells(1, 11).Text = S3.Cells(1, 11).Text
                S3.Cells(3, 11).Text = "Packing Slip No. : " & S2.Cells(S2.ActiveRowIndex, 1).Text
                S3.Cells(8, 12).Text = "EXODUS Wireless, Corp." 'Site_nm & " Corp."
                S4.Cells(8, 12).Text = "EXODUS Wireless, Corp." 'Site_nm & " Corp."
                S5.Cells(8, 12).Text = "EXODUS Wireless, Corp." 'Site_nm & " Corp."
                S4_1.Cells(8, 12).Text = "EXODUS Wireless, Corp." 'Site_nm & " Corp."
                S5_1.Cells(8, 12).Text = "EXODUS Wireless, Corp." 'Site_nm & " Corp."
                If Site_nm = "Stellar" Then
                    S3.Cells(43, 0).Text = "EXODUS Wireless, Corp."
                ElseIf Site_nm = "ETS" Then
                    S3.Cells(43, 0).Text = "EXODUS WIRELESS, INC"
                Else
                    S3.Cells(43, 0).Text = Site_nm & " CORPORATION"
                End If
                S4.Cells(43, 0).Text = S3.Cells(43, 0).Text
                S4_1.Cells(43, 0).Text = S3.Cells(43, 0).Text
                S5.Cells(43, 0).Text = S3.Cells(43, 0).Text
                S5_1.Cells(43, 0).Text = S3.Cells(43, 0).Text

                PK_RS = Query_RS_ALL("EXEC SP_FRMFSHIPPING_PACKINGLSIT '" & Site_id & "','" & S2.GetValue(S2.ActiveRowIndex, 1) & "'")

                PkRow = PK_RS.RecordCount

                S3.Visible = True

                If PkRow > 45 Then
                    MessageBox.Show("Box's total > 45", "ALERT")
                    Exit Sub
                End If

                If PkRow > 15 Then
                    S3.Cells(3, 11).Text = "Packing Slip No. : " & S2.Cells(S2.ActiveRowIndex, 1).Text & "-1"
                    S4.Cells(3, 11).Text = "Packing Slip No. : " & S2.Cells(S2.ActiveRowIndex, 1).Text & "-2"
                    S4_1.Cells(3, 11).Text = "Packing Slip No. : " & S2.Cells(S2.ActiveRowIndex, 1).Text & "-2"
                    S4.Visible = True

                End If
                If PkRow > 30 Then
                    S5.Cells(3, 11).Text = "Packing Slip No. : " & S2.Cells(S2.ActiveRowIndex, 1).Text & "-3"
                    S5_1.Cells(3, 11).Text = "Packing Slip No. : " & S2.Cells(S2.ActiveRowIndex, 1).Text & "-3"
                    S5.Visible = True

                End If

                If PK_RS.RecordCount > 0 Then

                    For I = 12 To 26
                        For j = 1 To S3.ColumnCount - 1
                            S3.Cells(I, j).Text = ""
                            S4.Cells(I, j).Text = ""
                            S5.Cells(I, j).Text = ""
                            S4_1.Cells(I, j).Text = ""
                            S5_1.Cells(I, j).Text = ""
                        Next
                    Next

                    For I = 30 To 39
                        S3.Cells(I, 1).Text = ""
                        S3.Cells(I, 3).Text = ""
                        S3.Cells(I, 6).Text = ""
                        S3.Cells(I, 8).Text = ""
                        S4.Cells(I, 1).Text = ""
                        S4.Cells(I, 3).Text = ""
                        S4.Cells(I, 6).Text = ""
                        S4.Cells(I, 8).Text = ""
                        S5.Cells(I, 1).Text = ""
                        S5.Cells(I, 3).Text = ""
                        S5.Cells(I, 6).Text = ""
                        S5.Cells(I, 8).Text = ""
                        S4_1.Cells(I, 1).Text = ""
                        S4_1.Cells(I, 3).Text = ""
                        S4_1.Cells(I, 6).Text = ""
                        S4_1.Cells(I, 8).Text = ""
                        S5_1.Cells(I, 1).Text = ""
                        S5_1.Cells(I, 3).Text = ""
                        S5_1.Cells(I, 6).Text = ""
                        S5_1.Cells(I, 8).Text = ""

                    Next

                    For I = 12 To PK_RS.RecordCount + 11
                        If I < 27 Then
                            S3.Cells(I, 1).Text = PK_RS(0).Value
                            S3.Cells(I, 3).Text = PK_RS(1).Value
                            S3.Cells(I, 6).Text = PK_RS(2).Value
                            S3.Cells(I, 9).Text = PK_RS(3).Value
                            S3.Cells(I, 11).Text = PK_RS(4).Value
                            S3.Cells(I, 12).Text = PK_RS(5).Value

                            If pmd1 <> PK_RS(2).Value Then
                                If l1 < 10 Then
                                    S3.Cells(30 + l1, 1).Text = PK_RS(2).Value
                                Else
                                    S3.Cells(20 + l1, 6).Text = PK_RS(2).Value
                                End If
                                pmd1 = PK_RS(2).Value
                                l1 += 1
                            End If
                        ElseIf I > 26 And I < 42 Then
                            S4.Cells(I - 15, 1).Text = PK_RS(0).Value
                            S4.Cells(I - 15, 3).Text = PK_RS(1).Value
                            S4.Cells(I - 15, 6).Text = PK_RS(2).Value
                            S4.Cells(I - 15, 9).Text = PK_RS(3).Value
                            S4.Cells(I - 15, 11).Text = PK_RS(4).Value
                            S4.Cells(I - 15, 12).Text = PK_RS(5).Value

                            If pmd2 <> PK_RS(2).Value Then
                                If l2 < 10 Then
                                    S4.Cells(30 + l2, 1).Text = PK_RS(2).Value
                                    S4_1.Cells(30 + l2, 1).Text = PK_RS(2).Value
                                Else
                                    S4.Cells(20 + l2, 6).Text = PK_RS(2).Value
                                    S4_1.Cells(20 + l2, 6).Text = PK_RS(2).Value
                                End If
                                pmd2 = PK_RS(2).Value
                                l2 += 1
                            End If

                            S4_1.Cells(I - 15, 1).Text = PK_RS(0).Value
                            S4_1.Cells(I - 15, 3).Text = PK_RS(1).Value
                            S4_1.Cells(I - 15, 6).Text = PK_RS(2).Value
                            S4_1.Cells(I - 15, 9).Text = PK_RS(3).Value
                            S4_1.Cells(I - 15, 11).Text = PK_RS(4).Value
                            S4_1.Cells(I - 15, 12).Text = PK_RS(5).Value
                        Else
                            S5.Cells(I - 30, 1).Text = PK_RS(0).Value
                            S5.Cells(I - 30, 3).Text = PK_RS(1).Value
                            S5.Cells(I - 30, 6).Text = PK_RS(2).Value
                            S5.Cells(I - 30, 9).Text = PK_RS(3).Value
                            S5.Cells(I - 30, 11).Text = PK_RS(4).Value
                            S5.Cells(I - 30, 12).Text = PK_RS(5).Value
                            If pmd3 <> PK_RS(2).Value Then
                                If l3 < 10 Then
                                    S5.Cells(30 + l3, 1).Text = PK_RS(2).Value
                                    S5_1.Cells(30 + l3, 1).Text = PK_RS(2).Value
                                Else
                                    S5_1.Cells(20 + l3, 6).Text = PK_RS(2).Value
                                    S5.Cells(20 + l3, 6).Text = PK_RS(2).Value
                                End If
                                pmd3 = PK_RS(2).Value
                                l3 += 1
                            End If

                            S5_1.Cells(I - 30, 1).Text = PK_RS(0).Value
                            S5_1.Cells(I - 30, 3).Text = PK_RS(1).Value
                            S5_1.Cells(I - 30, 6).Text = PK_RS(2).Value
                            S5_1.Cells(I - 30, 9).Text = PK_RS(3).Value
                            S5_1.Cells(I - 30, 11).Text = PK_RS(4).Value
                            S5_1.Cells(I - 30, 12).Text = PK_RS(5).Value
                        End If


                        PK_RS.MoveNext()
                    Next

                End If


                If l1 > 0 Then
                    For m = 0 To l1
                        If m < 10 Then
                            For j = 12 To 26
                                If S3.Cells(j, 6).Text <> "" And S3.Cells(j, 6).Text = S3.Cells(30 + m, 1).Text Then
                                    S3.Cells(30 + m, 3).Value += S3.Cells(j, 9).Value
                                End If
                            Next
                        Else
                            For j = 12 To 26
                                If S3.Cells(j, 6).Text <> "" And S3.Cells(j, 6).Text = S3.Cells(20 + m, 6).Text Then
                                    S3.Cells(20 + m, 8).Value += S3.Cells(j, 9).Value
                                End If
                            Next
                        End If
                    Next
                End If

                If l2 > 0 Then
                    For m = 0 To l2
                        If m < 10 Then
                            For j = 12 To 26
                                If S4.Cells(j, 6).Text <> "" And S4.Cells(j, 6).Text = S4.Cells(30 + m, 1).Text Then
                                    S4.Cells(30 + m, 3).Value += S4.Cells(j, 9).Value
                                End If
                                If S4_1.Cells(j, 6).Text <> "" And S4_1.Cells(j, 6).Text = S4_1.Cells(30 + m, 1).Text Then
                                    S4_1.Cells(30 + m, 3).Value += S4_1.Cells(j, 9).Value
                                End If

                            Next
                        Else
                            For j = 12 To 26
                                If S4.Cells(j, 6).Text <> "" And S4.Cells(j, 6).Text = S4.Cells(20 + m, 6).Text Then
                                    S4.Cells(20 + m, 8).Value += S4.Cells(j, 9).Value
                                End If

                                If S4_1.Cells(j, 6).Text <> "" And S4_1.Cells(j, 6).Text = S4_1.Cells(20 + m, 6).Text Then
                                    S4_1.Cells(20 + m, 8).Value += S4_1.Cells(j, 9).Value
                                End If
                            Next
                        End If
                    Next
                End If

                If l3 > 0 Then
                    For m = 0 To l3
                        If m < 10 Then
                            For j = 12 To 26
                                If S5.Cells(j, 6).Text <> "" And S5.Cells(j, 6).Text = S5.Cells(30 + m, 1).Text Then
                                    S5.Cells(30 + m, 3).Value += S5.Cells(j, 9).Value
                                End If
                                If S5_1.Cells(j, 6).Text <> "" And S5_1.Cells(j, 6).Text = S5_1.Cells(30 + m, 1).Text Then
                                    S5_1.Cells(30 + m, 3).Value += S5_1.Cells(j, 9).Value
                                End If
                            Next
                        Else
                            For j = 12 To 26
                                If S5.Cells(j, 6).Text <> "" And S5.Cells(j, 6).Text = S5.Cells(20 + m, 6).Text Then
                                    S5.Cells(20 + m, 8).Value += S5.Cells(j, 9).Value
                                End If

                                If S5_1.Cells(j, 6).Text <> "" And S5_1.Cells(j, 6).Text = S5_1.Cells(20 + m, 6).Text Then
                                    S5_1.Cells(20 + m, 8).Value += S5_1.Cells(j, 9).Value
                                End If
                            Next
                        End If
                    Next
                End If

                PK_RS = Nothing



            End If
            MainFrm.ProgressBarItem1.ProgressType = DevComponents.DotNetBar.eProgressItemType.Standard
            Bar5.AutoHide = False
            '            Bar5.Left = FpSpread1.Left
            DockContainerItem5.Width = 800

        Catch ex As Exception
            MessageBox.Show("Error : " & ex.Message, "Error")
        End Try
    End Sub

    Private Sub NewBtn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NewBtn.Click, NewBtn1.Click
        Dim i, i_qty As New Integer
        LabelX4.Text = ""

        FpSpread1.ActiveSheet.RowCount = 0
        FpSpread2.ActiveSheet.RowCount = 0

        'FpSpread3.ActiveSheet.ColumnCount = 4
        'FpSpread3.ActiveSheet.ClearRange(0, 0, FpSpread3.ActiveSheet.RowCount, 4, False)
        'FpSpread3.ActiveSheet.RowCount = 0


        FpSpread2.ActiveSheet.RowCount = 1
        FpSpread2.ActiveSheet.Rows(0).ForeColor = Color.OrangeRed
        ListViewEx1.CheckBoxes = True

        ListViewEx1.Items.Clear()

        Bar6.Visible = True

        FpSpread2.ActiveSheet.SetValue(0, 0, Query_RS("select convert(varchar(8), getdate(), 112)"))
        FpSpread2.ActiveSheet.SetValue(0, 1, Query_RS("EXEC SP_FRMFSHIPPING_GETNEWSHIPNO '" & Site_id & "'"))

        If Query_Listview(ListViewEx1, "exec SP_FRMFSHIPPING_GETCURBOX '" & Site_id & "', 'all'", True) = True Then

            For i = 0 To ListViewEx1.Items.Count - 1
                i_qty = i_qty + CInt(ListViewEx1.Items(i).SubItems(4).Text)
            Next

            ListViewEx1.Items.Add("TOTAL")
            ListViewEx1.Items(ListViewEx1.Items.Count - 1).SubItems.Add("")
            ListViewEx1.Items(ListViewEx1.Items.Count - 1).SubItems.Add("")
            ListViewEx1.Items(ListViewEx1.Items.Count - 1).SubItems.Add("")
            ListViewEx1.Items(ListViewEx1.Items.Count - 1).SubItems.Add(i_qty)
            ListViewEx1.Items(ListViewEx1.Items.Count - 1).BackColor = Color.Yellow
            ListViewEx1.Items(ListViewEx1.Items.Count - 1).SubItems(1).BackColor = Color.Yellow
            ListViewEx1.Items(ListViewEx1.Items.Count - 1).SubItems(2).BackColor = Color.Yellow
            ListViewEx1.Items(ListViewEx1.Items.Count - 1).SubItems(3).BackColor = Color.Yellow
            ListViewEx1.Items(ListViewEx1.Items.Count - 1).SubItems(4).BackColor = Color.Yellow


            ListViewEx1.AutoResizeColumn(0, ColumnHeaderAutoResizeStyle.ColumnContent)
            'Bar4.AutoHide = False
        End If
        DockContainerItem4.Selected = True


        SaveBtn.Enabled = True
        SaveBtn1.Enabled = True

        TextBoxX1.Focus()

    End Sub

    Private Sub SaveBtn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SaveBtn.Click, SaveBtn1.Click
        Dim i As New Integer

        Me.Cursor = Cursors.WaitCursor

        If MessageBox.Show("Are you save this shipping as " & FpSpread2.ActiveSheet.GetValue(0, 1) & "?", "Save Shipping", MessageBoxButtons.YesNo) = Windows.Forms.DialogResult.Yes Then
            SaveBtn.Enabled = False
            MainFrm.ProgressBarItem1.Maximum = FpSpread1.ActiveSheet.RowCount

            ' 기사용한 Shipping No 인지 Check
            If Query_RS("select count(Pship_no) from view_Fesnmaster where Pship_no = '" & FpSpread2.ActiveSheet.GetValue(0, 1) & "'") > 0 Then
                System.Console.Beep(3000, 400)
                System.Console.Beep(3000, 400)
                Modal_Error("Already Used Shipping no " & vbNewLine & FpSpread2.ActiveSheet.GetValue(0, 1) & "!!")

                FindBtn_Click(sender, e)
                Exit Sub
            End If

            ' Wty Upload 중인지 Check
            'If Query_RS("select count(inv_no) from tbl_esnmaster_b where inv_no is not null") > 0 Then
            '    System.Console.Beep(3000, 400)
            '    System.Console.Beep(3000, 400)
            '    Modal_Error("On Wty Uploading!! Please Save after Wty Upload")

            '    FindBtn_Click(sender, e)
            '    Exit Sub
            'End If


            For i = 0 To FpSpread1.ActiveSheet.RowCount - 1

                MainFrm.ProgressBarItem1.Value = i + 1

                If FpSpread1.ActiveSheet.Rows(i).ForeColor = Color.OrangeRed Then
                    If Insert_Data("EXEC SP_FRMFSHIPPING_SAVE_1 '" & Site_id & "','" & FpSpread1.ActiveSheet.GetValue(i, 1) & "','" & Emp_No & "','W1700','" & FpSpread2.ActiveSheet.GetValue(0, 1) & "'") = True Then
                        FpSpread1.ActiveSheet.Rows(i).ForeColor = Color.Black
                    End If
                End If
            Next

            If Insert_Data("EXEC SP_FRMFSHIPPING_TRANSFER '" & Site_id & "','" & FpSpread2.ActiveSheet.GetValue(0, 1) & "','S0001', '" & Emp_No & "'") = False Then
                MessageBox.Show("SHIPPING TRANSFER ERROR")
            End If

            FpSpread1.ActiveSheet.RowCount = 0
            MessageBox.Show("저장되었습니다")
            MainFrm.ProgressBarItem1.Value = 0
            'Bar4.AutoHide = True

            If Query_Listview(ListViewEx1, "exec SP_FRMFSHIPPING_GETCURBOX '" & Site_id & "', 'all'", True) = True Then
                ListViewEx1.AutoResizeColumn(0, ColumnHeaderAutoResizeStyle.ColumnContent)
                'Bar4.AutoHide = False
            End If

            If Query_Spread(FpSpread2, "EXEC SP_FRMFSHIPPING_GETSHIPPEDSUMMARY '" & Site_id & "','" & DateTimeInput1.Text & "','" & DateTimeInput2.Text & "'", 1) = True Then
                FpSpread2.ActiveSheet.RowCount = FpSpread2.ActiveSheet.RowCount + 1
                FpSpread2.ActiveSheet.Rows(FpSpread2.ActiveSheet.RowCount - 1).BackColor = Color.Yellow
                FpSpread2.ActiveSheet.SetValue(FpSpread2.ActiveSheet.RowCount - 1, 0, "TOTAL")

                FpSpread2.AllowUserFormulas = True

                FpSpread2.ActiveSheet.Cells(FpSpread2.ActiveSheet.RowCount - 1, 2).Formula = "SUM(C1:C" & FpSpread2.ActiveSheet.RowCount - 1 & ")"

                Spread_AutoCol(FpSpread2)
            End If

            ListViewEx1.CheckBoxes = False
        Else

        End If

        SaveBtn.Enabled = False
        SaveBtn1.Enabled = False



        MessageBox.Show("저장되었습니다")

        Bar6.Visible = False
        Me.Cursor = Cursors.Default
        SaveBtn.Enabled = True
        DockContainerItem3.Selected = True



    End Sub

    Private Sub XlsBtn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles XlsBtn.Click, XlsBtn1.Click

        If save_excel = "" Then
            MessageBox.Show("Select Spread for saving!!")
            Exit Sub
        End If


        If save_excel = "FpSpread1" Then
            SaveFileDialog1.FileName = LabelX4.Text
            File_Save(SaveFileDialog1, FpSpread1)
            SaveFileDialog1.FileName = ""

        ElseIf save_excel = "FpSpread2" Then
            File_Save(SaveFileDialog1, FpSpread2)
        ElseIf save_excel = "FpSpread3" Then
            File_Save2(SaveFileDialog1, FpSpread3, PkRow)
        Else
            MessageBox.Show("Select Spread for Save!!")
        End If

    End Sub

    Private Sub PrtBtn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PrtBtn.Click, PrtBtn1.Click, ButtonItem6.Click
        '      Try
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

            If PkRow < 16 Then
                s_cnt = 1
            ElseIf PkRow > 15 And PkRow < 31 Then
                s_cnt = 2
            ElseIf PkRow > 30 Then
                s_cnt = 3

            End If

            sheet_print2()

        Else
            MessageBox.Show("No Printing Object!!")
        End If

        'Catch ex As Exception
        '    MessageBox.Show("Error : " & ex.Message, "Error")
        'End Try
    End Sub

    Private Sub sheet_print2()
        Try

            Dim pi As New FarPoint.Win.Spread.PrintInfo()
            Dim pm As New FarPoint.Win.Spread.PrintMargin

            pm.Right = 0
            pm.Top = 40
            pm.Bottom = 20
            pm.Left = 30

            pi.ShowBorder = False
            pi.ShowColor = False
            pi.ShowGrid = False
            pi.ShowShadows = False
            pi.UseMax = False
            pi.Margin = pm

            pi.Orientation = FarPoint.Win.Spread.PrintOrientation.Auto

            pi.Centering = FarPoint.Win.Spread.Centering.Horizontal

            pi.ColStart = 0
            pi.RowStart = 0
            pi.ColEnd = 14
            pi.RowEnd = 44
            pi.PrintType = FarPoint.Win.Spread.PrintType.PageRange
            pi.PageStart = 1
            pi.PageEnd = 2
            pi.ShowPrintDialog = False


            pi.UseSmartPrint = False
            pi.PrintType = FarPoint.Win.Spread.PrintType.PageRange

            With FpSpread3
                .Sheets(0).PrintInfo = pi
                .PrintSheet(0)
            End With

            With FpSpread5
                If .ActiveSheet.GetValue(12, 1) <> "" Then
                    .Sheets(0).PrintInfo = pi
                    .PrintSheet(0)
                End If
            End With

            With FpSpread6
                If .ActiveSheet.GetValue(12, 1) <> "" Then
                    .Sheets(0).PrintInfo = pi
                    .PrintSheet(0)
                End If
            End With

            pi = Nothing
            pm = Nothing

        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message, "ERROR")
        End Try
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

    Private Sub TextBoxX1_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TextBoxX1.KeyDown


        If e.KeyValue = Keys.Enter Then

            If ListViewEx1.FindItemWithText(TextBoxX1.Text) Is Nothing Then
                Modal_Error(TextBoxX1.Text & vbNewLine & "Wrong BOXID !!")
            Else
                ListViewEx1.FindItemWithText(TextBoxX1.Text).Checked = True
            End If
            TextBoxX1.Text = ""
        End If

    End Sub


    Private Sub ButtonItem3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonItem3.Click
        If Me.ButtonItem3.Text = "Go to Next page" Then
            Me.FpSpread3.ActiveSheet = Me.FpSpread3_Sheet2
            Me.ButtonItem3.Text = "Go to First page"
        Else
            Me.FpSpread3.ActiveSheet = Me.FpSpread3_Sheet1
            Me.ButtonItem3.Text = "Go to Next page"
        End If

    End Sub

    'Private Sub ButtonItem6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonItem6.Click
    '    Try
    '        Paking_Print(Me.FpSpread3, PkRow, 0)

    '    Catch ex As Exception
    '        MessageBox.Show("Error : " & ex.Message, "Error")
    '    End Try
    'End Sub


    Private Sub FpSpread3_CellDoubleClick(ByVal sender As Object, ByVal e As FarPoint.Win.Spread.CellClickEventArgs) Handles FpSpread3.CellDoubleClick

        Me.FpSpread3.ActiveSheet.Cells(1, 11).Locked = False

    End Sub

    Private Sub FpSpread3_Change(ByVal sender As Object, ByVal e As FarPoint.Win.Spread.ChangeEventArgs) Handles FpSpread3.Change
        If Me.FpSpread3.ActiveSheetIndex = 0 Then
            Me.FpSpread3.ActiveSheet.Cells(1, 11).Locked = False
            Me.FpSpread3.Sheets(1).Cells(1, 11).Text = Me.FpSpread3.Sheets(0).Cells(1, 11).Text
        End If
        If Me.FpSpread3.ActiveSheetIndex = 1 Then
            Me.FpSpread3.ActiveSheet.Cells(1, 11).Locked = False
            Me.FpSpread3.Sheets(0).Cells(1, 11).Text = Me.FpSpread3.Sheets(1).Cells(1, 11).Text
        End If
    End Sub

    Private Sub ButtonItem11_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonItem11.Click
        Try


            Dim S1 = Me.FpSpread1.ActiveSheet
            Dim s2 = Me.FpSpread2.ActiveSheet
            Dim S4 = Me.FpSpread4.ActiveSheet
            Dim i As Integer
            If S1.RowCount < 1 Then
                Exit Sub
            End If

            If S4.RowCount > 0 Then
                S4.Rows.Clear()
            End If

            For i = 0 To S1.RowCount - 1
                S4.RowCount += 1
                S4.SetValue(i, 0, S1.GetValue(i, 2))
            Next

            SaveFileDialog1.FileName = s2.GetValue(s2.ActiveRowIndex, 1) & "-DMCHK"
            File_Save(SaveFileDialog1, FpSpread4)

        Catch ex As Exception
            MessageBox.Show("Error : " & ex.Message, "Error")
        End Try
    End Sub

    'Private Sub FpSpread3_PrintMessageBox(ByVal sender As Object, ByVal e As FarPoint.Win.Spread.PrintMessageBoxEventArgs) Handles FpSpread3.PrintMessageBox

    '    If p_cnt > 2 Then
    '        Exit Sub
    '    End If

    '    If e.BeginPrinting = True And s_cnt > 1 And p_cnt < 3 Then

    '        If PkRow > 15 And PkRow < 31 Then
    '            sheet_print(Me.FpSpread3, 1, 0)
    '            p_cnt += 1
    '            Exit Sub
    '        Else
    '            sheet_print(Me.FpSpread3, 2, 0)
    '            Exit Sub
    '        End If

    '        'If PkRow > 30 Then
    '        '    sheet_print(Me.FpSpread3, 2, 0)
    '        '    p_cnt += 1
    '        'End If
    '    Else
    '        Exit Sub
    '    End If


    'End Sub


    Private Sub ButtonItem12_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonItem12.Click

        'If MessageBox.Show("Are you sure to fix Error ?", "Fix Error", MessageBoxButtons.YesNo) = Windows.Forms.DialogResult.Yes Then
        '    MessageBox.Show(Query_RS("EXEC SP_FRMSHIPPING_SHIPERROR"))
        '    FindBtn_Click(sender, e)
        '    CheckBoxX1.Checked = False
        'End If

    End Sub

    Private Sub CheckBoxX1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBoxX1.CheckedChanged
        If CheckBoxX1.Checked = True Then
            ButtonItem12.Visible = True
        Else
            ButtonItem12.Visible = False

        End If
    End Sub

    Private Sub ButtonItem14_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonItem14.Click

        With ListViewEx1
            If .Items.Count > 0 Then
                Dim r As DialogResult = MessageBox.Show("Are you sure to cancel box " & .SelectedItems(0).SubItems(3).Text & "?", "Cancel Packing BOX", MessageBoxButtons.YesNo)
                If r = Windows.Forms.DialogResult.Yes Then
                    '                    Insert_Data("EXEC SP_FRMPACKING_CANCELPACKING_PQC '" & Site_id & "','" & .SelectedItems(0).SubItems(3).Text & "','" & .SelectedItems(0).Text & "','" & Emp_No & "'")
                    Insert_Data("EXEC SP_FRMFPACKING_CANCELPACKING '" & Site_id & "','" & .SelectedItems(0).SubItems(3).Text & "','" & .SelectedItems(0).Text & "','" & Emp_No & "'")                    '                    Refresh_Current(ComboBoxEx3.Text)
                End If

            End If
        End With


    End Sub




End Class