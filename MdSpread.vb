Module MdSpread
    Public textcell As New FarPoint.Win.Spread.CellType.TextCellType()
    Public datecell As New FarPoint.Win.Spread.CellType.DateTimeCellType()
    Public combocell As New FarPoint.Win.Spread.CellType.ComboBoxCellType()
    Public intcell As New FarPoint.Win.Spread.CellType.NumberCellType()
    Public deccell As New FarPoint.Win.Spread.CellType.NumberCellType()
    Dim btncell As New FarPoint.Win.Spread.CellType.ButtonCellType()
    Public CHKcell As New FarPoint.Win.Spread.CellType.CheckBoxCellType
    Public percell As New FarPoint.Win.Spread.CellType.PercentCellType()
    Public curcell As New FarPoint.Win.Spread.CellType.CurrencyCellType()
    Public gencell As New FarPoint.Win.Spread.CellType.GeneralCellType
    Public imgcell As New FarPoint.Win.Spread.CellType.ImageCellType
    Public timecell As New FarPoint.Win.Spread.CellType.DateTimeCellType()

    Public sheetcnt As Integer = 0


    Function Spread_Setting(ByVal s As FarPoint.Win.Spread.FpSpread, ByVal form_name As String) As Boolean
        Dim spread_rs As New ADODB.Recordset
        Dim i As Integer
        Dim textcell1 As New FarPoint.Win.Spread.CellType.TextCellType()

        intcell.DecimalPlaces = 0
        intcell.Separator = ","
        intcell.ShowSeparator = True
        intcell.MaximumValue = 99999999999
        '        intcell.MinimumValue = -99999999999

        deccell.DecimalPlaces = 2
        deccell.Separator = ","
        deccell.ShowSeparator = True
        deccell.MaximumValue = 99999999999
        '       deccell.MinimumValue = 0

        curcell.DecimalPlaces = 0
        curcell.Separator = ","
        curcell.ShowSeparator = True
        curcell.ShowCurrencySymbol = True
        curcell.CurrencySymbol = ""
        curcell.MaximumValue = 99999999999
        '      curcell.MinimumValue = 0

        datecell.DropDownButton = False
        datecell.ButtonAlign = FarPoint.Win.ButtonAlign.Right

        spread_rs = Query_RS_ALL("Select column_name, isnull(celltype,''),hide_yn, isnull(cellallign,'LEFT'), isnull(sort_yn,'N'), isnull(column_size,0), isnull(case_type, '') from tbl_spread where form_name ='" & form_name & "' and spread_id = '" & s.Name & "' order by column_no")

        s.Font = New Font("Verdana", 8)

        '       textcell.CharacterCasing = CharacterCasing.Upper

        If spread_rs Is Nothing Then
            Exit Function
        End If

        If spread_rs.RecordCount > 0 Then

            With s.ActiveSheet

                Dim c As New FarPoint.Win.Spread.CellType.ColumnHeaderRenderer
                c.WordWrap = False
                .ColumnHeader.DefaultStyle.Renderer = c

                .ColumnCount = spread_rs.RecordCount
                '               .RowCount = 0

                For i = 0 To spread_rs.RecordCount - 1
                    .ColumnHeader.Columns(i).Label = spread_rs(0).Value
                    If spread_rs(1).Value = "INT" Then
                        .Columns(i).CellType = intcell
                    ElseIf spread_rs(1).Value = "DEC" Then
                        .Columns(i).CellType = deccell
                    ElseIf spread_rs(1).Value = "COMBO" Then
                        .Columns(i).CellType = combocell
                    ElseIf spread_rs(1).Value = "DATE" Then
                        .Columns(i).CellType = datecell
                    ElseIf spread_rs(1).Value = "CHECKBOX" Then
                        .Columns(i).CellType = CHKcell
                    ElseIf spread_rs(1).Value = "PERCENT" Then
                        .Columns(i).CellType = percell
                    ElseIf spread_rs(1).Value = "CURRENCY" Then
                        .Columns(i).CellType = curcell
                    Else

                        If spread_rs(5).Value = 0 And spread_rs(6).Value = "" Then
                            .Columns(i).CellType = textcell
                        ElseIf spread_rs(5).Value > 0 And spread_rs(6).Value = "" Then
                            'textcell1.MaxLength = spread_rs(5).Value
                            .Columns(i).CellType = textcell1
                        ElseIf spread_rs(5).Value > 0 And spread_rs(6).Value = "UPPER" Then
                            'textcell1.MaxLength = spread_rs(5).Value
                            textcell1.CharacterCasing = CharacterCasing.Upper
                            .Columns(i).CellType = textcell1
                        Else
                            'If spread_rs(6).Value = "UPPER" Then
                            '    textcell.CharacterCasing = CharacterCasing.Upper
                            'End If
                            .Columns(i).CellType = textcell
                        End If


                    End If

                    If spread_rs(2).Value = "Y" Then
                        .Columns(i).Visible = False
                    End If

                    .Columns(i).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
                    If spread_rs(3).Value = "LEFT" Then
                        .Columns(i).HorizontalAlignment = FarPoint.Win.Spread.CellHorizontalAlignment.General
                    ElseIf spread_rs(3).Value = "RIGHT" Then
                        .Columns(i).HorizontalAlignment = FarPoint.Win.Spread.CellHorizontalAlignment.Right
                    ElseIf spread_rs(3).Value = "CENTER" Then
                        .Columns(i).HorizontalAlignment = FarPoint.Win.Spread.CellHorizontalAlignment.Center
                    End If

                    If spread_rs(4).Value = "Y" Then
                        ' .Columns(i).SortIndicator = FarPoint.Win.Spread.Model.SortIndicator.Ascending
                        .Columns(i).AllowAutoSort = True
                    End If

                    spread_rs.MoveNext()
                Next

                .OperationMode = FarPoint.Win.Spread.OperationMode.RowMode
                '.AlternatingRows(0).BackColor = Color.Beige
                '.Protect = True
                .Columns(0, spread_rs.RecordCount - 1).Locked = True
                .SheetName = Replace(form_name, "Frm", "")
            End With

        End If

        Spread_Setting = True

        spread_rs = Nothing

    End Function

    Function Spread_Celltype(ByVal S As FarPoint.Win.Spread.FpSpread, ByVal frmNm As String, ByVal rowidx As Integer, ByVal colidx As Integer) As Boolean
        Dim txtcell As New FarPoint.Win.Spread.CellType.TextCellType()

        txtcell.MaxLength = CInt(Query_RS("select ISNULL(column_size,99999) from tbl_spread where form_name = '" & frmNm & "' and spread_id = '" & S.Name & "' and column_no = " & colidx))

        If Query_RS("select case_type from tbl_spread where form_name = '" & frmNm & "' and spread_id = '" & S.Name & "' and column_no = " & colidx) = "UPPER" Then
            txtcell.CharacterCasing = CharacterCasing.Upper
        End If

        S.ActiveSheet.Cells(rowidx, colidx).CellType = txtcell

    End Function

    Function Spread_CurrencyType(ByVal S As FarPoint.Win.Spread.FpSpread, ByVal rowidx As Integer, ByVal colidx As Integer) As Boolean
        Dim CurrencyType As New FarPoint.Win.Spread.CellType.CurrencyCellType()
        CurrencyType.CurrencySymbol = "$"
        CurrencyType.DecimalPlaces = 2
        CurrencyType.DecimalSeparator = "."
        CurrencyType.Separator = ","
        CurrencyType.ShowSeparator = True

        S.ActiveSheet.Cells(rowidx, colidx).CellType = CurrencyType
    End Function

    Function Spread_NumType(ByVal S As FarPoint.Win.Spread.FpSpread, ByVal rowidx As Integer, ByVal colidx As Integer) As Boolean
        Dim NumType As New FarPoint.Win.Spread.CellType.NumberCellType
        NumType.DecimalPlaces = 0
        NumType.Separator = ","
        NumType.ShowSeparator = True
        S.ActiveSheet.Cells(rowidx, colidx).CellType = NumType
    End Function

    Function Spread_PercentType(ByVal S As FarPoint.Win.Spread.FpSpread, ByVal rowidx As Integer, ByVal colidx As Integer) As Boolean

        Dim PercentType As New FarPoint.Win.Spread.CellType.PercentCellType
        PercentType.DecimalPlaces = 2
        PercentType.DecimalSeparator = "."
        PercentType.Separator = ","
        PercentType.ShowSeparator = True
        S.ActiveSheet.Cells(rowidx, colidx).CellType = PercentType
    End Function

    Function Spread_TxtType(ByVal S As FarPoint.Win.Spread.FpSpread, ByVal rowidx As Integer, ByVal colidx As Integer, ByVal Upper_Yn As Boolean, ByVal Maxlength As Integer) As Boolean
        Dim TxtType As New FarPoint.Win.Spread.CellType.TextCellType
        If Maxlength > 0 Then
            TxtType.MaxLength = Maxlength
        End If
        If Upper_Yn = True Then
            TxtType.CharacterCasing = CharacterCasing.Upper
        End If
        S.ActiveSheet.Cells(rowidx, colidx).CellType = TxtType
    End Function

    Function Spread_border(ByVal S As FarPoint.Win.Spread.FpSpread, ByVal rowidx As Integer, ByVal colidx As Integer, ByVal first_yn As Boolean) As Boolean
        Dim CellBorder As FarPoint.Win.ComplexBorder
        If first_yn = True Then
            CellBorder = New FarPoint.Win.ComplexBorder(New FarPoint.Win.ComplexBorderSide(FarPoint.Win.ComplexBorderSideStyle.ThinLine), New FarPoint.Win.ComplexBorderSide(FarPoint.Win.ComplexBorderSideStyle.None), New FarPoint.Win.ComplexBorderSide(FarPoint.Win.ComplexBorderSideStyle.ThinLine), New FarPoint.Win.ComplexBorderSide(FarPoint.Win.ComplexBorderSideStyle.ThinLine))
        Else
            CellBorder = New FarPoint.Win.ComplexBorder(New FarPoint.Win.ComplexBorderSide(FarPoint.Win.ComplexBorderSideStyle.None), New FarPoint.Win.ComplexBorderSide(FarPoint.Win.ComplexBorderSideStyle.None), New FarPoint.Win.ComplexBorderSide(FarPoint.Win.ComplexBorderSideStyle.ThinLine), New FarPoint.Win.ComplexBorderSide(FarPoint.Win.ComplexBorderSideStyle.ThinLine))
        End If

        S.ActiveSheet.Cells(rowidx, colidx).Border = CellBorder

    End Function

    Function cell_border(ByVal S As FarPoint.Win.Spread.FpSpread, ByVal rowidx As Integer, ByVal colidx As Integer, ByVal left_bd As FarPoint.Win.ComplexBorderSideStyle, ByVal right_bd As FarPoint.Win.ComplexBorderSideStyle, ByVal top_bd As FarPoint.Win.ComplexBorderSideStyle, ByVal bottom_bd As FarPoint.Win.ComplexBorderSideStyle) As Boolean
        'complexborder의 테두리
        Dim CellBorder As FarPoint.Win.ComplexBorder

        CellBorder = New FarPoint.Win.ComplexBorder(New FarPoint.Win.ComplexBorderSide(left_bd), New FarPoint.Win.ComplexBorderSide(top_bd), New FarPoint.Win.ComplexBorderSide(right_bd), New FarPoint.Win.ComplexBorderSide(bottom_bd))
        S.ActiveSheet.Cells(rowidx, colidx).Border = CellBorder


    End Function

    Function Spread_CellAlign(ByVal S As FarPoint.Win.Spread.FpSpread, ByVal rowidx As Integer, ByVal colidx As Integer, ByVal align As String) As Boolean
        'align : Horizontal left, right, center
        S.ActiveSheet.Cells(rowidx, colidx).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Select LCase(align)
            Case "left"
                S.ActiveSheet.Cells(rowidx, colidx).HorizontalAlignment = FarPoint.Win.Spread.CellHorizontalAlignment.Left
            Case "center"
                S.ActiveSheet.Cells(rowidx, colidx).HorizontalAlignment = FarPoint.Win.Spread.CellHorizontalAlignment.Center
            Case "right"
                S.ActiveSheet.Cells(rowidx, colidx).HorizontalAlignment = FarPoint.Win.Spread.CellHorizontalAlignment.Right
        End Select
    End Function
    Function Spread_AutoCol(ByVal S As FarPoint.Win.Spread.FpSpread) As Boolean
        Dim col As FarPoint.Win.Spread.Column
        Dim size As Single
        Dim i As Integer
        Dim lencol As Integer
        Dim colnm As String

        S.ActiveSheet.DataAutoSizeColumns = False
        For i = 0 To S.ActiveSheet.ColumnCount - 1
            '            S.Sheets(0).GetPreferredColumnWidth(i, False, False)
            col = S.ActiveSheet.Columns(i)
            If col.Visible = False Then
                Continue For
            End If
            '            FpSpread1.ActiveSheet.Cells(0, 0).Text = "This text will be used to determine the width of the column"
            size = col.GetPreferredWidth() + 4
            colnm = S.ActiveSheet.Columns(i).Label

            If size > Len(S.ActiveSheet.ColumnHeader.Columns(i).Label) * 10 Then
                lencol = Len(S.ActiveSheet.ColumnHeader.Columns(i).Label)
                col.Width = size
            Else
                col.Width = Len(S.ActiveSheet.ColumnHeader.Columns(i).Label) * 10
            End If
        Next

    End Function


    Function Spread_Print(ByVal S As FarPoint.Win.Spread.FpSpread, ByVal hd As String, ByVal Ppage As Integer) As Boolean
        Spread_Print = False

        Dim pi As New FarPoint.Win.Spread.PrintInfo
        Dim pi1 As New FarPoint.Win.Spread.PrintInfo
        Dim pm As New FarPoint.Win.Spread.PrintMargin

        pm.Right = 0
        pm.Top = 30
        pm.Bottom = 20
        pm.Left = 10

        'Set Printing options for spreadsheet
        pi.ShowColumnHeader = FarPoint.Win.Spread.PrintHeader.Hide
        pi.ShowRowHeader = FarPoint.Win.Spread.PrintHeader.Hide

        pi.ShowBorder = False
        pi.ShowColor = False
        pi.ShowGrid = True
        pi.ShowShadows = False
        pi.UseMax = False
        pi.Margin = pm

        Select Case Ppage
            Case 0
                pi.Orientation = FarPoint.Win.Spread.PrintOrientation.Auto
            Case 1
                pi.Orientation = FarPoint.Win.Spread.PrintOrientation.Landscape
            Case 2
                pi.Orientation = FarPoint.Win.Spread.PrintOrientation.Portrait
        End Select

        pi.Centering = FarPoint.Win.Spread.Centering.Horizontal

        pi.ColStart = 0
        pi.RowStart = 0
        pi.ColEnd = S.ActiveSheet.ColumnCount - 1
        pi.RowEnd = S.ActiveSheet.RowCount - 1
        pi.PrintType = FarPoint.Win.Spread.PrintType.CellRange

       
        pi.ZoomFactor = S.ZoomFactor
        pi.Preview = True
        pi.ShowPrintDialog = True


        '       pi.Header = "/c/fz""16""/fb1" & hd & "/n/n/n/n/fz"
        pi.Footer = "/l/fz""9""/fb1" & Site_nm & "" & "/r/fb1" & Now

        S.ActiveSheet.PrintInfo = pi
        S.PrintSheet(0)

        Spread_Print = True
    End Function

    Function Spread_Print2(ByVal S As FarPoint.Win.Spread.FpSpread, ByVal hd As String, ByVal Ppage As Integer) As Boolean
        Spread_Print2 = False

        Dim pi As New FarPoint.Win.Spread.PrintInfo
        Dim pi1 As New FarPoint.Win.Spread.PrintInfo
        Dim pm As New FarPoint.Win.Spread.PrintMargin

        pm.Right = 0
        pm.Top = 40
        pm.Bottom = 20
        pm.Left = 30

        'Set Printing options for spreadsheet
        pi.ShowColumnHeader = FarPoint.Win.Spread.PrintHeader.Show
        pi.ShowRowHeader = FarPoint.Win.Spread.PrintHeader.Show

        pi.ShowBorder = True
        pi.ShowColor = False
        pi.ShowGrid = True
        pi.ShowShadows = False
        pi.UseMax = False
        pi.Margin = pm

        Select Case Ppage
            Case 0
                pi.Orientation = FarPoint.Win.Spread.PrintOrientation.Auto
            Case 1
                pi.Orientation = FarPoint.Win.Spread.PrintOrientation.Landscape
            Case 2
                pi.Orientation = FarPoint.Win.Spread.PrintOrientation.Portrait
        End Select

        pi.Centering = FarPoint.Win.Spread.Centering.Horizontal

        pi.ColStart = 0
        pi.RowStart = 0
        pi.ColEnd = S.ActiveSheet.ColumnCount - 1
        pi.RowEnd = S.ActiveSheet.RowCount - 1
        pi.PrintType = FarPoint.Win.Spread.PrintType.CellRange
        pi.Preview = True
        pi.ShowPrintDialog = True
        pi.ZoomFactor = 90%

        pi.Header = "/c/fz""16""/fb1" & hd & "/n/n/n/n/fz"
        pi.Footer = "/l/fz""9""/fb1" & Site_nm & ", Inc." & "/r/fb1" & Now

        S.ActiveSheet.PrintInfo = pi
        S.PrintSheet(0)

        Spread_Print2 = True
    End Function

    Function Spread_Print3(ByVal S As FarPoint.Win.Spread.FpSpread, ByVal Ppage As Integer) As Boolean
        Spread_Print3 = False

        Dim pi As New FarPoint.Win.Spread.PrintInfo
        Dim pi1 As New FarPoint.Win.Spread.PrintInfo
        Dim pm As New FarPoint.Win.Spread.PrintMargin

        pm.Right = 0
        pm.Top = 60
        pm.Bottom = 20
        pm.Left = 30

        'Set Printing options for spreadsheet
        pi.ShowColumnHeader = FarPoint.Win.Spread.PrintHeader.Hide
        pi.ShowRowHeader = FarPoint.Win.Spread.PrintHeader.Hide
        'pi.Footer = "/l/fz""9""/fb1" & Site_nm & " CORPATION "
        pi.ShowBorder = False
        pi.ShowColor = False
        pi.ShowGrid = False
        pi.ShowShadows = False
        pi.UseMax = False
        pi.Margin = pm

        Select Case Ppage
            Case 0
                pi.Orientation = FarPoint.Win.Spread.PrintOrientation.Auto
            Case 1
                pi.Orientation = FarPoint.Win.Spread.PrintOrientation.Landscape
            Case 2
                pi.Orientation = FarPoint.Win.Spread.PrintOrientation.Portrait
        End Select

        pi.Centering = FarPoint.Win.Spread.Centering.Horizontal

        pi.ColStart = 0
        pi.RowStart = 0
        pi.ColEnd = S.ActiveSheet.ColumnCount - 1
        pi.RowEnd = S.ActiveSheet.RowCount - 1
        pi.PrintType = FarPoint.Win.Spread.PrintType.CellRange
        pi.ShowPrintDialog = True


        S.ActiveSheet.PrintInfo = pi
        S.PrintSheet(0)

        Spread_Print3 = True
    End Function


    Function Paking_Print(ByVal S As FarPoint.Win.Spread.FpSpread, ByVal Pkrow As Integer, ByVal Ppage As Integer) As Boolean
        Try

            Paking_Print = False


            sheet_print(S, 0, Ppage)



            If Pkrow > 15 Then
                sheetcnt = 1
            Else
                sheetcnt = 0
            End If

            Paking_Print = True


        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message, "ERROR")
        End Try

    End Function

    Function sheet_print(ByVal S As FarPoint.Win.Spread.FpSpread, ByVal i As Integer, ByVal Ppage As Integer) As Boolean
        '      Try

        sheet_print = False

        Dim pi As New FarPoint.Win.Spread.PrintInfo
        Dim pm As New FarPoint.Win.Spread.PrintMargin

        pm.Right = 0
        pm.Top = 40
        pm.Bottom = 20
        pm.Left = 30

        'Set Printing options for spreadsheet
        pi.ShowColumnHeader = FarPoint.Win.Spread.PrintHeader.Hide
        pi.ShowRowHeader = FarPoint.Win.Spread.PrintHeader.Hide

        pi.ShowBorder = False
        pi.ShowColor = False
        pi.ShowGrid = False
        pi.ShowShadows = False
        pi.UseMax = False
        pi.Margin = pm

        Select Case Ppage
            Case 0
                pi.Orientation = FarPoint.Win.Spread.PrintOrientation.Auto
            Case 1
                pi.Orientation = FarPoint.Win.Spread.PrintOrientation.Landscape
            Case 2
                pi.Orientation = FarPoint.Win.Spread.PrintOrientation.Portrait
        End Select

        pi.Centering = FarPoint.Win.Spread.Centering.Horizontal

        pi.ColStart = 0
        pi.RowStart = 0
        'pi.ColEnd = 14
        'pi.RowEnd = 44
        pi.PrintType = FarPoint.Win.Spread.PrintType.CellRange

        pi.ZoomFactor = S.ZoomFactor
        pi.Preview = True
        pi.ShowPrintDialog = True

        'pi.Header = "/c/fz""16""/fb1" & hd & "/n/n/n/n/fz"
        'pi.Footer = "/l/fz""9""/fb1UDIATECH, Inc." & "/r/fb1" & Now

        S.Sheets(i).PrintInfo = pi
        S.PrintSheet(i)

        sheet_print = True

        pi = Nothing
        pm = Nothing

        'Catch ex As Exception
        '    MessageBox.Show("Error: " & ex.Message, "ERROR")
        'End Try
    End Function

    Function sheet_print1(ByVal S As FarPoint.Win.Spread.FpSpread, ByVal i As Integer, ByVal j As Integer, ByVal Ppage As Integer) As Boolean
        Try

            sheet_print1 = False

            Dim pi As New FarPoint.Win.Spread.PrintInfo
            Dim pm As New FarPoint.Win.Spread.PrintMargin

            pm.Right = 0
            pm.Top = 40
            pm.Bottom = 20
            pm.Left = 30

            'Set Printing options for spreadsheet
            'pi.ShowColumnHeader = FarPoint.Win.Spread.PrintHeader.Show
            'pi.ShowRowHeader = FarPoint.Win.Spread.PrintHeader.Show

            pi.ShowBorder = False
            pi.ShowColor = False
            pi.ShowGrid = False
            pi.ShowShadows = False
            pi.UseMax = False
            pi.Margin = pm

            Select Case Ppage
                Case 0
                    pi.Orientation = FarPoint.Win.Spread.PrintOrientation.Auto
                Case 1
                    pi.Orientation = FarPoint.Win.Spread.PrintOrientation.Landscape
                Case 2
                    pi.Orientation = FarPoint.Win.Spread.PrintOrientation.Portrait
            End Select

            pi.Centering = FarPoint.Win.Spread.Centering.Horizontal

            pi.ColStart = 0
            pi.RowStart = 0
            pi.ColEnd = 14
            pi.RowEnd = 44
            pi.PrintType = FarPoint.Win.Spread.PrintType.CellRange
            pi.ShowPrintDialog = True



            S.Sheets(0).PrintInfo = pi

            'For k = 0 To j
            S.PrintSheet(0)
            '            S.PrintSheet(S.Sheets(1))
            'Next

            sheet_print1 = True

            pi = Nothing
            pm = Nothing

        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message, "ERROR")
        End Try
    End Function

    Function sheet_print2(ByVal S As FarPoint.Win.Spread.FpSpread) As Boolean
        Try

            sheet_print2 = False

            Dim pi As New FarPoint.Win.Spread.PrintInfo
            Dim pm As New FarPoint.Win.Spread.PrintMargin

            'pm.Right = 0
            'pm.Top = 40
            'pm.Bottom = 20
            'pm.Left = 30

            'Set Printing options for spreadsheet
            pi.ShowColumnHeader = FarPoint.Win.Spread.PrintHeader.Hide
            pi.ShowRowHeader = FarPoint.Win.Spread.PrintHeader.Hide

            pi.ShowBorder = False
            pi.ShowColor = False
            pi.ShowGrid = False
            pi.ShowShadows = False
            pi.UseMax = False
            pi.Margin = pm

            'Select Case Ppage
            '    Case 0
            'pi.Orientation = FarPoint.Win.Spread.PrintOrientation.Auto
            '    Case 1
            pi.Orientation = FarPoint.Win.Spread.PrintOrientation.Landscape
            '    Case 2
            'pi.Orientation = FarPoint.Win.Spread.PrintOrientation.Portrait
            'End Select

            pi.Centering = FarPoint.Win.Spread.Centering.Both

            'pi.ColStart = 0
            'pi.RowStart = 0
            'pi.ColEnd = 14
            'pi.RowEnd = 44
            pi.PrintType = FarPoint.Win.Spread.PrintType.All
            'pi.PageStart = 1
            'pi.PageEnd = 2
            pi.ShowPrintDialog = True

            Dim i, j As Integer

            j = S.Sheets.Count

            For i = 0 To j - 1
                S.Sheets(i).PrintInfo = pi
                S.PrintSheet(i)
                System.Threading.Thread.Sleep(2000)
            Next

            sheet_print2 = True

            pi = Nothing
            pm = Nothing

        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message, "ERROR")
        End Try
    End Function

    Function sheet_print3(ByVal S As FarPoint.Win.Spread.FpSpread) As Boolean
        'TEST REPORT만 출력, 다이얼로그 띄우지 않음.
        Try

            sheet_print3 = False

            Dim pi As New FarPoint.Win.Spread.PrintInfo
            Dim pm As New FarPoint.Win.Spread.PrintMargin
            Dim AA As FarPoint.Win.Spread.SheetView

            'Set Printing options for spreadsheet
            pi.ShowColumnHeader = FarPoint.Win.Spread.PrintHeader.Hide
            pi.ShowRowHeader = FarPoint.Win.Spread.PrintHeader.Hide
            pm.Left = 40
            pm.Top = 50

            pi.ShowBorder = False
            pi.ShowColor = False
            pi.ShowGrid = False
            pi.ShowShadows = False
            pi.UseMax = False
            pi.Margin = pm

            pi.Orientation = FarPoint.Win.Spread.PrintOrientation.Landscape
            pi.Centering = FarPoint.Win.Spread.Centering.Both
            pi.PrintType = FarPoint.Win.Spread.PrintType.CurrentPage

            'pi.PrintType
            'pi.PageStart = S.Sheets.Count - 1
            'pi.PageEnd = S.Sheets.Count - 1
            pi.ShowPrintDialog = False



            AA = S.Sheets("Test Report")

            S.Sheets(S.Sheets.IndexOf(AA)).PrintInfo = pi

            'For k = 0 To j
            S.PrintSheet(S.Sheets.IndexOf(AA))
            '            S.PrintSheet(S.Sheets(1))
            'Next

            sheet_print3 = True

            pi = Nothing
            pm = Nothing

        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message, "ERROR")
        End Try
    End Function

    Function sheet_print4(ByVal S As FarPoint.Win.Spread.FpSpread, ByVal sheet As String) As Boolean
        'TEST REPORT만 출력, 다이얼로그 띄우지 않음.
        Try

            sheet_print4 = False

            Dim pi As New FarPoint.Win.Spread.PrintInfo
            Dim pm As New FarPoint.Win.Spread.PrintMargin
            Dim AA As FarPoint.Win.Spread.SheetView

            'Set Printing options for spreadsheet
            pi.ShowColumnHeader = FarPoint.Win.Spread.PrintHeader.Hide
            pi.ShowRowHeader = FarPoint.Win.Spread.PrintHeader.Hide
            pm.Left = 0
            pm.Right = 0
            pm.Top = 0

            pi.ShowBorder = False
            pi.ShowColor = False
            pi.ShowGrid = False
            pi.ShowShadows = False
            pi.UseMax = True
            pi.Margin = pm

            pi.UseSmartPrint = True
            pi.SmartPrintPagesWide = 1

            pi.Orientation = FarPoint.Win.Spread.PrintOrientation.Landscape
            pi.Centering = FarPoint.Win.Spread.Centering.Horizontal
            pi.PrintType = FarPoint.Win.Spread.PrintType.CellRange

            pi.ColStart = 0
            pi.ColEnd = S.ActiveSheet.ColumnCount - 1
            pi.RowStart = 0
            pi.RowEnd = S.ActiveSheet.RowCount - 1

            '            pi.Preview = True

            pi.ShowPrintDialog = False

            AA = S.Sheets.Find(sheet)

            pi.ColStart = 0
            pi.ColEnd = AA.ColumnCount - 1
            pi.RowStart = 0
            pi.RowEnd = AA.RowCount - 1

            'Dim i As Integer

            'For i = 0 To AA.ColumnCount - 1
            '    MessageBox.Show(i & ": " & AA.Columns(i).Width)
            'Next

            System.Threading.Thread.Sleep(1000)
            AA.PrintInfo = pi

            S.PrintSheet(AA)
            AA.Dispose()

            sheet_print4 = True

            pi = Nothing
            pm = Nothing

        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message, "ERROR")
        End Try
    End Function

    Function Spread_Change(ByVal s As FarPoint.Win.Spread.FpSpread, ByVal row As Integer) As Boolean

        s.ActiveSheet.Rows(row).ForeColor = Color.OrangeRed

    End Function

    Function Query_Spread_USA(ByVal S As FarPoint.Win.Spread.FpSpread, ByVal Q_TEXT As String, ByVal FLAG As Integer) As Boolean
        Try

            Dim conn As New ADODB.Connection
            Dim cmd As New ADODB.Command
            Dim rs As New ADODB.Recordset
            Dim i, j As Integer

            S.ActiveSheet.Rows.Clear()

            conn.CommandTimeout = 600
            conn.ConnectionTimeout = 600

            conn.Open(db_conn(2))

            cmd.ActiveConnection = conn
            cmd.CommandTimeout = 600


            cmd.CommandType = ADODB.CommandTypeEnum.adCmdText
            cmd.CommandText = Q_TEXT

            rs = New ADODB.Recordset
            rs.ActiveConnection = conn
            rs.CursorLocation = ADODB.CursorLocationEnum.adUseClient
            rs.CursorType = ADODB.CursorTypeEnum.adOpenForwardOnly
            rs.LockType = ADODB.LockTypeEnum.adLockReadOnly

            rs.Open(cmd.CommandText)

            If rs.RecordCount > 0 Then
                MainFrm.ProgressBarItem1.Maximum = rs.RecordCount

                Query_Spread_USA = True
                If FLAG = 1 Then
                    S.ActiveSheet.RowCount = rs.RecordCount

                    For i = 0 To rs.RecordCount - 1
                        For j = 0 To S.ActiveSheet.ColumnCount - 1
                            If rs(j).Value IsNot System.DBNull.Value Then

                                If S.ActiveSheet.Columns(j).CellType Is Nothing Then
                                    S.ActiveSheet.Columns(j).CellType = textcell
                                End If

                                If S.ActiveSheet.Columns(j).CellType.ToString = "DateTimeCellType" Then
                                    If Trim(rs(j).Value) = "1900-01-01" Then
                                    Else

                                        S.ActiveSheet.SetValue(i, j, Trim(rs(j).Value))
                                        S.ActiveSheet.SetText(i, j, Trim(rs(j).Value))
                                    End If
                                Else
                                    S.ActiveSheet.SetValue(i, j, Trim(rs(j).Value))
                                    S.ActiveSheet.SetText(i, j, Trim(rs(j).Value))

                                End If
                            End If
                        Next
                        rs.MoveNext()
                        MainFrm.ProgressBarItem1.Value = i + 1
                    Next
                Else
                    S.ActiveSheet.RowCount = S.ActiveSheet.RowCount + 1
                    For j = 0 To S.ActiveSheet.ColumnCount - 1
                        If rs(j).Value IsNot System.DBNull.Value Then
                            S.ActiveSheet.SetValue(S.ActiveSheet.RowCount - 1, j, Trim(rs(j).Value))
                        End If
                    Next
                End If

            Else
                S.ActiveSheet.RowCount = 0
                Query_Spread_USA = False
            End If

            '        rs.Close()
            rs = Nothing
            MainFrm.ProgressBarItem1.Maximum = 0
            MainFrm.ProgressBarItem1.Value = 0
            conn.Close()


        Catch ex As Exception

            MessageBox.Show("Error: " & ex.Message, "ERROR")
        End Try

    End Function



    Function Query_Spread(ByVal S As FarPoint.Win.Spread.FpSpread, ByVal Q_TEXT As String, ByVal FLAG As Integer) As Boolean
        Try

            Dim conn As New ADODB.Connection
            Dim cmd As New ADODB.Command
            Dim rs As New ADODB.Recordset
            Dim i, j As Integer

            S.ActiveSheet.Rows.Clear()

            conn.CommandTimeout = 600
            conn.ConnectionTimeout = 600

            conn.Open(db_conn(1))

            cmd.ActiveConnection = conn
            cmd.CommandTimeout = 600


            cmd.CommandType = ADODB.CommandTypeEnum.adCmdText
            cmd.CommandText = Q_TEXT

            rs = New ADODB.Recordset
            rs.ActiveConnection = conn
            rs.CursorLocation = ADODB.CursorLocationEnum.adUseClient
            rs.CursorType = ADODB.CursorTypeEnum.adOpenForwardOnly
            rs.LockType = ADODB.LockTypeEnum.adLockReadOnly

            rs.Open(cmd.CommandText)

            If rs.RecordCount > 0 Then
                MainFrm.ProgressBarItem1.Maximum = rs.RecordCount

                Query_Spread = True
                If FLAG = 1 Then
                    S.ActiveSheet.RowCount = rs.RecordCount

                    For i = 0 To rs.RecordCount - 1
                        For j = 0 To S.ActiveSheet.ColumnCount - 1
                            If rs(j).Value IsNot System.DBNull.Value Then

                                If S.ActiveSheet.Columns(j).CellType Is Nothing Then
                                    S.ActiveSheet.Columns(j).CellType = textcell
                                End If

                                If S.ActiveSheet.Columns(j).CellType.ToString = "DateTimeCellType" Then
                                    If Trim(rs(j).Value) = "1900-01-01" Then
                                    Else

                                        S.ActiveSheet.SetValue(i, j, Trim(rs(j).Value))
                                        S.ActiveSheet.SetText(i, j, Trim(rs(j).Value))
                                    End If
                                Else
                                    S.ActiveSheet.SetValue(i, j, Trim(rs(j).Value))
                                    S.ActiveSheet.SetText(i, j, Trim(rs(j).Value))

                                End If
                            End If
                        Next
                        rs.MoveNext()
                        MainFrm.ProgressBarItem1.Value = i + 1
                    Next
                Else
                    S.ActiveSheet.RowCount = S.ActiveSheet.RowCount + 1
                    For j = 0 To S.ActiveSheet.ColumnCount - 1
                        If rs(j).Value IsNot System.DBNull.Value Then
                            S.ActiveSheet.SetValue(S.ActiveSheet.RowCount - 1, j, Trim(rs(j).Value))
                        End If
                    Next
                End If

            Else
                S.ActiveSheet.RowCount = 0
                Query_Spread = False
            End If

            '        rs.Close()
            rs = Nothing
            MainFrm.ProgressBarItem1.Maximum = 0
            MainFrm.ProgressBarItem1.Value = 0
            conn.Close()


        Catch ex As Exception

            MessageBox.Show("Error: " & ex.Message, "ERROR")
        End Try

    End Function

    Function Query_Spread_LTD(ByVal S As FarPoint.Win.Spread.FpSpread, ByVal Q_TEXT As String, ByVal F_COL As Integer, ByVal T_COL As Integer) As Boolean
        Dim conn As New ADODB.Connection
        Dim cmd As New ADODB.Command
        Dim rs As New ADODB.Recordset
        Dim i, j As Integer

        conn.Open(db_conn(1))
        cmd.ActiveConnection = conn

        cmd.CommandType = ADODB.CommandTypeEnum.adCmdText
        cmd.CommandText = Q_TEXT

        rs = New ADODB.Recordset
        rs.ActiveConnection = conn
        rs.CursorLocation = ADODB.CursorLocationEnum.adUseClient
        rs.CursorType = ADODB.CursorTypeEnum.adOpenForwardOnly
        rs.LockType = ADODB.LockTypeEnum.adLockReadOnly

        rs.Open(cmd.CommandText)

        If rs.RecordCount > 0 Then
            Query_Spread_LTD = True

            S.ActiveSheet.RowCount = S.ActiveSheet.RowCount + rs.RecordCount
            MainFrm.ProgressBarItem1.Maximum = rs.RecordCount

            For i = 0 To rs.RecordCount - 1
                For j = F_COL To T_COL
                    If rs(j).Value IsNot System.DBNull.Value Then
                        S.ActiveSheet.SetValue(i + S.ActiveSheet.RowCount - rs.RecordCount, j, Trim(rs(j).Value))
                        '     S.ActiveSheet.SetText(i + S.ActiveSheet.RowCount - rs.RecordCount, j, Trim(rs(j).Value))
                    End If
                Next
                rs.MoveNext()
                MainFrm.ProgressBarItem1.Value = i
            Next
        Else
            'S.ActiveSheet.RowCount = 0
            Query_Spread_LTD = False
        End If
        MainFrm.ProgressBarItem1.Maximum = 0

        rs = Nothing

    End Function

    Function Query_Spread_LTD_ROW(ByVal S As FarPoint.Win.Spread.FpSpread, ByVal Q_TEXT As String, ByVal F_COL As Integer, ByVal T_COL As Integer, ByVal S_ROW As Integer) As Boolean
        Dim conn As New ADODB.Connection
        Dim cmd As New ADODB.Command
        Dim rs As New ADODB.Recordset
        Dim i, j As Integer

        conn.Open(db_conn(1))
        cmd.ActiveConnection = conn

        cmd.CommandType = ADODB.CommandTypeEnum.adCmdText
        cmd.CommandText = Q_TEXT

        rs = New ADODB.Recordset
        rs.ActiveConnection = conn
        rs.CursorLocation = ADODB.CursorLocationEnum.adUseClient
        rs.CursorType = ADODB.CursorTypeEnum.adOpenForwardOnly
        rs.LockType = ADODB.LockTypeEnum.adLockReadOnly

        rs.Open(cmd.CommandText)

        If rs.RecordCount > 0 Then
            Query_Spread_LTD_ROW = True

            S.ActiveSheet.RowCount = S.ActiveSheet.RowCount + rs.RecordCount
            MainFrm.ProgressBarItem1.Maximum = rs.RecordCount

            For i = S_ROW To rs.RecordCount + S_ROW - 1
                For j = F_COL To T_COL
                    If rs(j).Value IsNot System.DBNull.Value Then
                        S.ActiveSheet.SetValue(i, j, Trim(rs(j).Value))
                    End If
                Next
                rs.MoveNext()
                MainFrm.ProgressBarItem1.Value = i
            Next
        Else
            'S.ActiveSheet.RowCount = 0
            Query_Spread_LTD_ROW = False
        End If
        MainFrm.ProgressBarItem1.Maximum = 0

        rs = Nothing

    End Function


    Function Add_Dataset(ByVal DsNm As DataSet, ByVal TableName As String, ByVal S As FarPoint.Win.Spread.FpSpread, ByVal rowidx As Integer) As Boolean
        Dim TbNewRow As DataRow = DsNm.Tables(TableName).NewRow
        Dim ColCnt As Integer
        For ColCnt = 0 To S.ActiveSheet.ColumnCount - 1
            TbNewRow(ColCnt) = S.ActiveSheet.Cells(rowidx, ColCnt).Text
        Next
        DsNm.Tables(TableName).Rows.Add(TbNewRow)
    End Function


    Function Chg_ComboCell(ByVal S As FarPoint.Win.Spread.FpSpread, ByVal Rowidx As Integer, ByVal colidx As Integer, ByVal itemary As String()) As Boolean
        Dim CbCellType = New FarPoint.Win.Spread.CellType.ComboBoxCellType
        Dim Dataary = itemary
        CbCellType.Items = Dataary
        S.ActiveSheet.Cells(Rowidx, colidx).CellType = CbCellType
    End Function

    Function Chg_ComboCell2(ByVal S As FarPoint.Win.Spread.FpSpread, ByVal Rowidx As Integer, ByVal colidx As Integer, ByVal itemary As Array) As Boolean
        Dim CbCellType = New FarPoint.Win.Spread.CellType.ComboBoxCellType
        Dim item0(0), item1(0) As String
        Dim i As Integer

        ReDim item0(itemary.GetLength(1) - 1)
        ReDim item1(itemary.GetLength(1) - 1)
        For i = 0 To itemary.GetLength(1) - 1
            item0(i) = itemary(0, i)
            item1(i) = itemary(1, i)
        Next
        CbCellType.EditorValue = FarPoint.Win.Spread.CellType.EditorValue.ItemData
        CbCellType.Items = item0
        CbCellType.ItemData = item1
        S.ActiveSheet.Cells(Rowidx, colidx).CellType = CbCellType
    End Function

    Function SPREAD_DUP_CHECK(ByVal S As FarPoint.Win.Spread.FpSpread, ByVal NAME As String, ByVal POS As Integer) As Boolean
        Dim I As Integer
        Dim NAME2 As String

        For I = 0 To S.ActiveSheet.RowCount - 1
            NAME2 = S.ActiveSheet.GetText(I, POS)
            If NAME2 = NAME Then
                SPREAD_DUP_CHECK = False
                Exit Function
            End If
        Next

        SPREAD_DUP_CHECK = True

    End Function

    Function SPREAD_DUP_CHECK_PART(ByVal S As FarPoint.Win.Spread.FpSpread, ByVal NAME As String, ByVal POS As Integer) As Boolean
        Dim I As Integer
        Dim NAME2 As String

        For I = 0 To S.ActiveSheet.RowCount - 1
            NAME2 = Mid(S.ActiveSheet.GetText(I, POS), 1, 11)
            If NAME2 = NAME Then
                SPREAD_DUP_CHECK_PART = False
                Exit Function
            End If
        Next

        SPREAD_DUP_CHECK_PART = True

    End Function

    Function SPREAD_DUP_ROW(ByVal S As FarPoint.Win.Spread.FpSpread, ByVal NAME As String, ByVal POS As Integer) As Integer
        Dim I As Integer
        Dim NAME2 As String

        For I = 0 To S.ActiveSheet.RowCount - 1
            NAME2 = S.ActiveSheet.GetText(I, POS)
            If NAME2 = NAME Then
                SPREAD_DUP_ROW = I
                Exit Function
            End If
        Next

        SPREAD_DUP_ROW = -1

    End Function

    Function SPREAD_DUP_ROW_1(ByVal S As FarPoint.Win.Spread.FpSpread, ByVal SIZE As Integer, ByVal NAME As String, ByVal POS As Integer) As Integer
        Dim I As Integer
        Dim NAME2 As String

        For I = 0 To S.ActiveSheet.RowCount - 1
            NAME2 = Microsoft.VisualBasic.Left(S.ActiveSheet.GetText(I, POS), SIZE)
            If NAME2 = NAME Then
                SPREAD_DUP_ROW_1 = I
                Exit Function
            End If
        Next

        SPREAD_DUP_ROW_1 = -1

    End Function

    Function SPREAD_DUP_CNT(ByVal S As FarPoint.Win.Spread.FpSpread, ByVal NAME As String, ByVal POS As Integer) As Integer
        Dim I As Integer

        SPREAD_DUP_CNT = 0

        For I = 0 To S.ActiveSheet.RowCount - 1
            If S.ActiveSheet.GetText(I, POS) = NAME Then
                SPREAD_DUP_CNT = SPREAD_DUP_CNT + 1

            End If
        Next

    End Function

    Function Basic_SValidate(ByVal DsNm As DataSet, ByVal TbNm As String, ByVal S As FarPoint.Win.Spread.FpSpread, ByVal stcnt As Integer) As Boolean
        Try
            'Save시에 새로 추가한 행들에 컬럼들이 NULL을 허용하지 않는 경우
            Basic_SValidate = False
            Dim TableNm = DsNm.Tables(TbNm)
            Dim TbCol, InsRow, endCnt As Integer
            Dim MsgTxt As String = ""
            endCnt = S.ActiveSheet.RowCount - 1
            For InsRow = stcnt To endCnt
                For TbCol = 0 To TableNm.Columns.Count - 1
                    If TableNm.Columns(TbCol).AllowDBNull = False And Len(S.ActiveSheet.Cells(InsRow, TbCol).Text) = 0 Then
                        S.ActiveSheet.Cells(InsRow, TbCol).Text = ""
                        MsgTxt = MsgTxt & InsRow + 1 & "TH OF [" & S.ActiveSheet.ColumnHeader.Columns(TbCol).Label & "] IS EMPTY!!" & Chr(13)
                    End If
                    If (TableNm.Columns(TbCol).DataType.Name.ToString = "Decimal" Or TableNm.Columns(TbCol).DataType.Name.ToString = "Int32") And Len(S.ActiveSheet.Cells(InsRow, TbCol).Text) = 0 Then
                        S.ActiveSheet.Cells(InsRow, TbCol).Text = 0
                    End If
                Next
            Next
            If MsgTxt <> "" Then
                MessageBox.Show(MsgTxt, "Validation Error")
            Else
                Basic_SValidate = True
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Function

    Function Cell_Vallidate(ByVal DsNm As DataSet, ByVal TbNm As String, ByVal S As FarPoint.Win.Spread.FpSpread, ByVal rowidx As Integer, ByVal colidx As Integer) As Boolean
        Try
            'Save시에 새로 추가한 행들에 컬럼들이 NULL을 허용하지 않는 경우
            Cell_Vallidate = False
            Dim TableNm = DsNm.Tables(TbNm)

            'If Len(S.ActiveSheet.Cells(rowidx, colidx).Text) > TableNm.Columns(colidx).MaxLength Then
            '    MessageBox.Show("[" & S.ActiveSheet.ColumnHeader.Columns(colidx).Label & "] Size is " & TableNm.Columns(colidx).MaxLength, "Validation Error")
            '    S.ActiveSheet.Cells(rowidx, colidx).Text = ""
            '    Exit Function
            'End If
            If TableNm.Columns(colidx).AllowDBNull = False And Len(S.ActiveSheet.Cells(rowidx, colidx).Text) = 0 Then
                MessageBox.Show("[" & S.ActiveSheet.ColumnHeader.Columns(colidx).Label & "] is empty", "Validation Error")
                S.ActiveSheet.Cells(rowidx, colidx).Text = ""

            Else
                Cell_Vallidate = True
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Function

    Function Save_Vallidate(ByVal S As FarPoint.Win.Spread.FpSpread, ByVal Title As String, ByVal colary As String()) As Boolean
        Try
            'Save시에 새로 추가한 행들에 컬럼들이 NULL을 허용하지 않는 경우
            Save_Vallidate = False
            Dim edrow = S.ActiveSheet.RowCount - 1
            Dim i, j As Integer
            Dim errmsg = ""
            For i = 0 To edrow
                For j = 0 To colary.Length - 1
                    If S.ActiveSheet.Cells(i, CInt(colary(j))).Text = "" Then
                        errmsg += i + 1 & "th [" & S.ActiveSheet.ColumnHeader.Columns(CInt(colary(j))).Label & "] is Empty!!" & Chr(13) & Chr(10)
                    End If
                Next
            Next
            If errmsg <> "" Then
                MessageBox.Show(errmsg, Title & "'s Validation Error")
            Else
                Save_Vallidate = True
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Function


    Function Spread_Setting_ByCode(ByVal s As FarPoint.Win.Spread.FpSpread, ByVal SP_NAME As String, ByVal CLASS_ID As String, ByVal TYPE As String) As Boolean
        Dim spread_rs As New ADODB.Recordset
        Dim i As Integer
        Try
            intcell.DecimalPlaces = 0
            intcell.Separator = ","
            intcell.ShowSeparator = True
            deccell.DecimalPlaces = 2
            deccell.Separator = ","
            deccell.ShowSeparator = True


            spread_rs = Query_RS_ALL("EXEC " & SP_NAME & " '" & Site_id & "','" & CLASS_ID & "'")

            If spread_rs Is Nothing Then
                s.ActiveSheet.RowCount = 0
                s.ActiveSheet.ColumnCount = 0
                Exit Function
            End If

            If spread_rs.RecordCount > 0 Then
                With s.ActiveSheet

                    Dim c As New FarPoint.Win.Spread.CellType.ColumnHeaderRenderer
                    c.WordWrap = False
                    .ColumnHeader.DefaultStyle.Renderer = c

                    .ColumnCount = spread_rs.RecordCount
                    '               .RowCount = 0

                    For i = 0 To spread_rs.RecordCount - 1
                        .ColumnHeader.Columns(i).Label = spread_rs(0).Value

                        If TYPE = "INT" Then
                            .Columns(i).CellType = intcell
                        ElseIf TYPE = "DEC" Then
                            .Columns(i).CellType = deccell
                        ElseIf TYPE = "COMBO" Then
                            .Columns(i).CellType = combocell
                        ElseIf TYPE = "DATE" Then
                            .Columns(i).CellType = datecell
                        ElseIf TYPE = "CHECKBOX" Then
                            .Columns(i).CellType = CHKcell
                        Else
                            .Columns(i).CellType = textcell
                        End If

                        spread_rs.MoveNext()
                    Next

                    .OperationMode = FarPoint.Win.Spread.OperationMode.RowMode
                    .AlternatingRows(0).BackColor = Color.Beige
                    .Protect = True
                    .Columns(0, spread_rs.RecordCount - 1).Locked = True

                End With

            End If

            Spread_Setting_ByCode = True

            spread_rs = Nothing
        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message, "ERROR")
        End Try

    End Function


    Function Spread_Setting_ByQuery(ByVal s As FarPoint.Win.Spread.FpSpread, ByVal SP_NAME As String, ByVal Type As String) As Boolean
        Dim spread_rs As New ADODB.Recordset
        Dim i As Integer
        Try
            intcell.DecimalPlaces = 0
            intcell.Separator = ","
            intcell.ShowSeparator = True
            deccell.DecimalPlaces = 2
            deccell.Separator = ","
            deccell.ShowSeparator = True

            'If SVR = "localhost" Then
            '    SVR = "."
            'End If
            spread_rs = Query_RS_ALL("EXEC " & SP_NAME)

            If spread_rs Is Nothing Then
                s.ActiveSheet.RowCount = 0
                s.ActiveSheet.ColumnCount = 0
                Exit Function
            End If

            If spread_rs.RecordCount > 0 Then
                With s.ActiveSheet

                    Dim c As New FarPoint.Win.Spread.CellType.ColumnHeaderRenderer
                    c.WordWrap = False
                    .ColumnHeader.DefaultStyle.Renderer = c

                    .ColumnCount = spread_rs.RecordCount
                    '               .RowCount = 0

                    For i = 0 To spread_rs.RecordCount - 1
                        .ColumnHeader.Columns(i).Label = spread_rs(0).Value

                        If Type = "INT" Then
                            .Columns(i).CellType = intcell
                        ElseIf Type = "DEC" Then
                            .Columns(i).CellType = deccell
                        ElseIf Type = "COMBO" Then
                            .Columns(i).CellType = combocell
                        ElseIf Type = "DATE" Then
                            .Columns(i).CellType = datecell
                        ElseIf Type = "CHECKBOX" Then
                            .Columns(i).CellType = CHKcell
                        Else
                            .Columns(i).CellType = textcell
                        End If

                        spread_rs.MoveNext()
                    Next

                    .OperationMode = FarPoint.Win.Spread.OperationMode.RowMode
                    .AlternatingRows(0).BackColor = Color.Beige
                    .Protect = True
                    .Columns(0, spread_rs.RecordCount - 1).Locked = True

                End With

            End If

            Spread_Setting_ByQuery = True

            spread_rs = Nothing
        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message, "ERROR")
        End Try

    End Function


    Function Generate_Spread_Query(ByVal s As FarPoint.Win.Spread.FpSpread, ByVal q1 As String, ByVal q2 As String, ByVal pos As Integer) As String
        Dim i As Integer

        Generate_Spread_Query = ""

        For i = pos To s.ActiveSheet.ColumnCount - pos - 1

            Generate_Spread_Query = Generate_Spread_Query & q1 & s.ActiveSheet.ColumnHeader.Columns(i).Label & q2 & vbNewLine

        Next
    End Function

    Function Generate_Spread_Query_1(ByVal s As FarPoint.Win.Spread.FpSpread, ByVal q1 As String, ByVal q2 As String, ByVal pos As Integer) As String
        Dim i As Integer

        Generate_Spread_Query_1 = ""

        For i = pos To s.ActiveSheet.ColumnCount - 2

            Generate_Spread_Query_1 = Generate_Spread_Query_1 & q1 & s.ActiveSheet.ColumnHeader.Columns(i).Label & q2 & vbNewLine

        Next
    End Function

    Function SPREAD_COL_TOTAL(ByVal S As FarPoint.Win.Spread.FpSpread, ByVal col1 As Integer, ByVal col2 As Integer, ByVal FLAG As Integer) As Boolean
        Dim i, J As Integer

        For J = col1 To col2
            For i = 0 To S.ActiveSheet.RowCount - 1
                S.ActiveSheet.SetValue(i, S.ActiveSheet.ColumnCount - 1, CInt(S.ActiveSheet.GetValue(i, S.ActiveSheet.ColumnCount - 1)) + CInt(S.ActiveSheet.GetValue(i, J)))
            Next
        Next

    End Function


    Function SPREAD_ROW_TOTAL(ByVal S As FarPoint.Win.Spread.FpSpread, ByVal col1 As Integer, ByVal col2 As Integer, ByVal FLAG As Integer) As Boolean
        Dim i, J As Integer

        For J = col1 To col2
            For i = 0 To S.ActiveSheet.RowCount - 2
                S.ActiveSheet.SetValue(S.ActiveSheet.RowCount - 1, J, CInt(S.ActiveSheet.GetValue(S.ActiveSheet.RowCount - 1, J)) + CInt(S.ActiveSheet.GetValue(i, J)))
            Next
        Next

    End Function

    Function SPREAD_ROW_TOTAL_DEC(ByVal S As FarPoint.Win.Spread.FpSpread, ByVal col1 As Integer, ByVal col2 As Integer, ByVal FLAG As Integer) As Boolean
        Dim i, J As Integer

        For J = col1 To col2
            For i = 0 To S.ActiveSheet.RowCount - 2
                S.ActiveSheet.SetValue(S.ActiveSheet.RowCount - 1, J, CDec(S.ActiveSheet.GetValue(S.ActiveSheet.RowCount - 1, J)) + CDec(S.ActiveSheet.GetValue(i, J)))
            Next
        Next

    End Function
    Function SPREAD_ROW_TOTAL_LTD(ByVal S As FarPoint.Win.Spread.FpSpread, ByVal col1 As Integer, ByVal col2 As Integer, ByVal row As Integer) As Boolean
        Dim i, J As Integer

        For J = col1 To col2
            For i = 0 To row - 1
                S.ActiveSheet.SetValue(row, J, CInt(S.ActiveSheet.GetValue(row, J)) + CInt(S.ActiveSheet.GetValue(i, J)))
            Next
        Next

    End Function

    'Function SPREAD_SEARCH(ByVal S As FarPoint.Win.Spread.FpSpread, ByVal Sheetidx As Integer, ByVal Search_TXT As String, ByVal StartRow As Integer, ByVal StartCol As Integer, ByVal EndRow As Integer, ByVal EndCol As Integer, ByVal useWildcards As Boolean) As String()
    '    Try
    '        Dim SchFp As FarPoint.Win.Spread.SearchFoundFlags
    '        Dim Fx, Fy As Integer
    '        SchFp = S.Search(0, Search_TXT, False, False, False, useWildcards, True, False, False, True, StartRow, StartCol, EndRow, EndCol, Fx, Fy)
    '        'S.Sheets(0).SetActiveCell(Fx, Fy)
    '        SPREAD_SEARCH = New String() {Fx, Fy}
    '    Catch ex As Exception
    '        MessageBox.Show(ex.Message)
    '    End Try
    'End Function

    Function SPREAD_SEARCH(ByVal S As FarPoint.Win.Spread.FpSpread, ByVal Sheetidx As Integer, ByVal Search_TXT As String, ByVal StartRow As Integer, ByVal StartCol As Integer, ByVal EndRow As Integer, ByVal EndCol As Integer, ByVal useWildcards As Boolean) As String()
        '스프레드에서 텍스트 검색
        Try
            Dim i, Fx, Fy As Integer
            Fx = -1
            Fy = -1
            'SchFp = S.Search(0, Search_TXT, False, False, False, useWildcards, True, False, False, True, StartRow, StartCol, EndRow, EndCol, Fx, Fy)
            If S.Sheets(Sheetidx).RowCount > 0 Then
                For i = 0 To S.ActiveSheet.RowCount - 1
                    If S.ActiveSheet.Cells(i, StartCol).Text = Search_TXT Then
                        Fx = i
                        Fy = StartCol
                        Exit For
                    End If
                Next
            End If
            'S.Sheets(0).SetActiveCell(Fx, Fy)
            SPREAD_SEARCH = New String() {Fx, Fy}
        Catch ex As Exception
            MessageBox.Show(ex.Message)
            SPREAD_SEARCH = New String() {}
        End Try
    End Function

    Function Cal_Qty(ByVal S As FarPoint.Win.Spread.FpSpread, ByVal Grpidx As Integer, ByVal GrpNm As String, ByVal Calidx As Integer) As Integer
        Try
            Dim i, totQty As Integer
            totQty = 0
            For i = 0 To S.ActiveSheet.RowCount - 1
                If S.ActiveSheet.Rows(i).ForeColor = Color.OrangeRed Then
                    If S.ActiveSheet.Cells(i, Grpidx).Text = GrpNm And S.ActiveSheet.Cells(i, Calidx).Text <> "" Then
                        totQty += CInt(S.ActiveSheet.Cells(i, Calidx).Text)
                    End If
                End If
            Next
            Cal_Qty = totQty
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Function

    Function SPREAD_LOADING(ByVal maxidx As Integer, ByVal minidx As Integer, ByVal stepidx As Integer) As Boolean
        MainFrm.ProgressBarItem1.Maximum = maxidx
        MainFrm.ProgressBarItem1.Minimum = minidx
        'MainFrm.ProgressBarItem1.Step = stepidx
        MainFrm.ProgressBarItem1.Value += 1
    End Function

    Function SPREAD_XLSIns(ByVal file_name As String, ByVal s As FarPoint.Win.Spread.FpSpread, ByVal frmname As String) As Boolean
        Dim conn As New ADODB.Connection
        Dim cmd As New ADODB.Command
        Dim rs, rst As New ADODB.Recordset
        Dim i, j, k, colcnt, stcnt, totcol As Integer
        Dim colary(50)
        Dim sheetnm As String

        SPREAD_XLSIns = False

        Try

            conn.CommandTimeout = 600
            conn.ConnectionTimeout = 600
            conn.Open(xls_conn(file_name))

            cmd.ActiveConnection = conn
            cmd.CommandTimeout = 600
            cmd.CommandType = ADODB.CommandTypeEnum.adCmdText
            cmd.CommandText = "select * from [" & Replace(frmname, "Frm", "") & "$]"

            rs = New ADODB.Recordset
            rs.ActiveConnection = conn
            rs.CursorLocation = ADODB.CursorLocationEnum.adUseClient
            rs.CursorType = ADODB.CursorTypeEnum.adOpenForwardOnly
            rs.LockType = ADODB.LockTypeEnum.adLockReadOnly

            rst = conn.OpenSchema(ADODB.SchemaEnum.adSchemaColumns)

            If rst.Fields(2).Value IsNot Nothing Then
                sheetnm = Replace(rst.Fields(2).Value, "$", "")
                If sheetnm <> Replace(frmname, "Frm", "") Then
                    MessageBox.Show("Excel's Sheetname is incorrect!!!", "Validation Error") '액셀시트명과 호출한 폼명이 다르면 처리안함
                    MsgBox(sheetnm & "," & Replace(frmname, "Frm", ""))
                    Exit Function
                End If
            Else
                MessageBox.Show("Inserted Sheet is nothing!!!", "Validation Error")
                Exit Function
            End If

            ReDim colary(s.ActiveSheet.ColumnCount)

            '스프레드시트에 Visible된 컬럼의 컬럼번호와 총 갯수를 가져온다.
            '액셀의 컬럼수는 스프레드시트에 Visible된 컬럼갯수와 같거나 작아야 한다. hidden 컬럼은 save시 자동 입력된다.
            totcol = 0
            For i = 0 To s.ActiveSheet.ColumnCount - 1
                ' If s.ActiveSheet.Columns(i).Visible = True Then
                colary(totcol) = i
                totcol += 1
                ' End If
            Next

            If rst.RecordCount = s.ActiveSheet.ColumnCount Then
                colcnt = s.ActiveSheet.ColumnCount
            ElseIf rst.RecordCount < s.ActiveSheet.ColumnCount Then
                colcnt = rst.RecordCount
            Else
                MessageBox.Show("Check Excel's Sheet_column count  !!!", "Validation Error")
                Exit Function
            End If
            'If s.ActiveSheet.Columns(0).Visible = False Then
            'stcnt = 1
            'Else
            stcnt = 0

            'End If
            For i = 0 To colcnt - 1 'create, udate 정보는 row에 추가될때 자동으로 입력하기 때문에, 비교에서 삭제
                If rst.Fields("COLUMN_NAME").Value <> s.ActiveSheet.ColumnHeader.Columns(colary(i)).Label Then
                    MessageBox.Show("[" & colary(i) & "] SPREAD'HEAD [" & s.ActiveSheet.ColumnHeader.Columns(colary(i)).Label & "] <> EXCEL'HEAD [" & rst.Fields("COLUMN_NAME").Value & "]", "Validation Error")
                    Exit Function
                End If
                rst.MoveNext()
            Next

            rs.Open(cmd.CommandText)
            'Using scope As New TransactionScope()
            Try

                If rs.RecordCount > 0 Then
                    j = 1
                    Do Until rs.EOF
                        With s.ActiveSheet
                            .RowCount += 1
                            'If stcnt = 1 Then
                            '.Cells(.RowCount - 1, 0).Text = Site_id
                            'End If

                            For i = 0 To colcnt - 1
                                If IsDBNull(rs(i).Value) = True Then
                                    .Cells(.RowCount - 1, colary(i)).Text = ""
                                Else
                                    .Cells(.RowCount - 1, colary(i)).Text = rs(i).Value
                                End If

                                .Cells(.RowCount - 1, colary(i)).Locked = False
                            Next

                            If s.ActiveSheet.ColumnHeader.Columns(0).Label = "SITE ID" And s.ActiveSheet.ColumnHeader.Columns(0).Visible = False Then
                                .Cells(.RowCount - 1, 0).Text = Site_id
                            End If

                            For k = totcol - 4 To totcol - 1
                                Select Case .ColumnHeader.Columns(k).Label
                                    Case "CREATOR"
                                        .Cells(.RowCount - 1, k).Text = Emp_No
                                    Case "CREATE DATE"
                                        .Cells(.RowCount - 1, k).Text = Now
                                    Case "UPDATOR"
                                        .Cells(.RowCount - 1, k).Text = Emp_No
                                    Case "UPDATE DATE"
                                        .Cells(.RowCount - 1, k).Text = Now
                                End Select
                            Next
                        End With
                        Spread_Change(s, s.ActiveSheet.RowCount - 1)
                        j += 1
                        rs.MoveNext()
                    Loop
                    Spread_AutoCol(s)
                End If

                SPREAD_XLSIns = True
            Catch ex As Exception
                s.ActiveSheet.Rows.Remove(s.ActiveSheet.RowCount - j, j)
                MessageBox.Show("Transaction Error: [Cell row :" & j & " , col :" & i + 1 & "]" & ex.Message)
            End Try
            'scope.Complete()
            'End Using
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try

    End Function


    Function Set_diagonal(ByVal s As FarPoint.Win.Spread.FpSpread, ByVal sheet_name As String, ByVal row As Integer, ByVal col As Integer) As Boolean
        Try

            Dim ls As New FarPoint.Win.Spread.DrawingSpace.LineShape
            Dim ls1 As New FarPoint.Win.Spread.DrawingSpace.LineShape

            With s.Sheets(sheet_name)
                ls.Width = .Columns(col).Width 'endPt.X - startPt.X
                ls.Height = .Rows(row).Height 'endPt.Y - startPt.Y - 5

                ls1.Width = .Columns(col).Width 'endPt.X - startPt.X
                ls1.Height = .Rows(row).Height 'endPt.Y - startPt.Y - 5
                ls1.FlipHorizontal = True

                .AddShape(ls, row, col)
                .AddShape(ls1, row, col)

            End With

        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Function
    Function Set_diagonal_Range2(ByVal s As FarPoint.Win.Spread.FpSpread, ByVal sheet_name As String, ByVal row As Integer, ByVal col As Integer, ByVal row1 As Integer, ByVal col1 As Integer) As Boolean
        Dim I, J As Integer

        For I = row To row1
            For J = col To col1
                Set_diagonal(s, sheet_name, I, J)
            Next
        Next
    End Function


    Function Set_diagonal_Range(ByVal s As FarPoint.Win.Spread.FpSpread, ByVal sheet_name As String, ByVal row As Integer, ByVal col As Integer, ByVal row1 As Integer, ByVal col1 As Integer) As Boolean
        Try

            Dim i As Integer
            Dim ls As New FarPoint.Win.Spread.DrawingSpace.LineShape
            Dim ls1 As New FarPoint.Win.Spread.DrawingSpace.LineShape
            ls.Height = 0
            ls.Width = 0
            ls1.Height = 0
            ls1.Width = 0

            With s.Sheets(sheet_name)
                For i = row To row1
                    ls.Height = ls.Height + .Rows(i).Height 'endPt.Y - startPt.Y - 5
                    ls1.Height = ls1.Height + .Rows(i).Height 'endPt.Y - startPt.Y - 5
                Next

                For i = col To col1
                    ls.Width = ls.Width + .Columns(i).Width
                    ls1.Width = ls1.Width + .Columns(i).Width
                Next
                ls1.FlipHorizontal = True

                ls.Height = ls.Height - 25
                ls.Width = ls.Width - 100
                ls1.Height = ls1.Height - 25
                ls1.Width = ls1.Width - 100

                .AddShape(ls, row, col)
                .AddShape(ls1, row, col)

            End With

        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Function

    Function SHEET_PAGESET(ByVal s As FarPoint.Win.Spread.FpSpread, ByVal SHEETNAME As String, ByVal col As Integer, ByVal row As Integer) As Boolean
        s.Sheets.Find(SHEETNAME).RowCount = row
        s.Sheets.Find(SHEETNAME).ColumnCount = col
    End Function

End Module
