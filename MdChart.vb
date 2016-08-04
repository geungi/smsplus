Module MdChart
    Function Chart_Query(ByVal ChartObj As Windows.Forms.DataVisualization.Charting.Chart, ByVal Chart3D_yn As Boolean, ByVal BarTypeIdx As Integer, ByVal Chart_Type As System.Windows.Forms.DataVisualization.Charting.SeriesChartType, ByVal SeriesItem As String(), ByVal Qtext As String) As Boolean
        ' DrawingStyle : default = 0 , Wedge = 1, Cylinder = 2, Emboss = 3, LightToDark = 4
        ' Qtext의 1st 컬럼은 X축, 2nd 이후 컬럼은 Y축에 표시됨
        ' SeriesItem은 Series의 Item 명을 의미함
        Try

            ChartObj.ChartAreas.Clear()
            ChartObj.Legends.Clear()
            ChartObj.Series.Clear()


            Dim Crs As ADODB.Recordset
            Dim i, j As Integer
            Dim Scnt As Integer = SeriesItem.Length
            Dim Lnd As System.Windows.Forms.DataVisualization.Charting.Legend = New System.Windows.Forms.DataVisualization.Charting.Legend
            Dim CArea As System.Windows.Forms.DataVisualization.Charting.ChartArea = New System.Windows.Forms.DataVisualization.Charting.ChartArea
            Dim Yarys(0, 0) As Integer
            Dim Xary(0) As String
            Dim Yary(0) As Integer
            Dim BarType As String = ""
            Dim FontInfo = New System.Drawing.Font("Verdana", 7.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))

            Lnd.Name = "Lnd"
            Lnd.BackColor = Color.Transparent
            Lnd.Font = FontInfo
            ChartObj.Legends.Add(Lnd)


            ChartObj.BackColor = Color.Transparent
            CArea.Name = "CArea"
            CArea.BackColor = Color.Transparent
            CArea.Area3DStyle.Enable3D = Chart3D_yn
            ChartObj.ChartAreas.Add(CArea)

            CArea.AxisX.IsLabelAutoFit = False
            CArea.AxisY.IsLabelAutoFit = False
            CArea.AxisX2.IsLabelAutoFit = False
            CArea.AxisY2.IsLabelAutoFit = False
            CArea.AxisX.LabelStyle.Font = FontInfo
            CArea.AxisY.LabelStyle.Font = FontInfo
            CArea.AxisX.LabelStyle.Interval = 1

            '         CArea.AxisY.LabelStyle.Interval = 1

            Crs = Query_RS_ALL(Qtext)

            If SeriesItem.Length <> Crs.Fields.Count - 1 Then
                MessageBox.Show("Series's Count must be Query's Column Count", "Validation Error")
                Exit Function
            End If

            If Crs.RecordCount < 1 Then
                MessageBox.Show("Record count 0 ", "Validation Error")
                Exit Function
            End If

            ReDim Xary(Crs.RecordCount - 1)
            ReDim Yary(Crs.RecordCount - 1)
            ReDim Yarys(Scnt - 1, Crs.RecordCount - 1)

            Select Case BarTypeIdx
                Case 0
                    BarType = ""
                Case 1
                    BarType = "DrawingStyle = Wedge"
                Case 2
                    BarType = "DrawingStyle = Cylinder"
                Case 3
                    BarType = "DrawingStyle = Emboss"
                Case 4
                    BarType = "DrawingStyle = LightToDark"
            End Select


            For i = 0 To Crs.RecordCount - 1
                Xary(i) = Crs.Fields(0).Value
                For j = 1 To Scnt
                    Yarys(j - 1, i) = Crs.Fields(j).Value
                Next
                Crs.MoveNext()
            Next

            For i = 0 To Scnt - 1
                For j = 0 To Yary.Length - 1
                    Yary(j) = Yarys(i, j)
                Next
                ChartObj.Series.Add(SeriesItem(i))
                ChartObj.Series.Item(SeriesItem(i)).ChartArea = CArea.Name
                ChartObj.Series.Item(SeriesItem(i)).CustomProperties = BarType
                ChartObj.Series.Item(SeriesItem(i)).ChartType = Chart_Type
                ChartObj.Series.Item(SeriesItem(i)).Font = FontInfo
                ChartObj.Series.Item(SeriesItem(i)).BorderWidth = 2
                'ChartObj.Series.Item(SeriesItem(i)).IsValueShownAsLabel = True
                ChartObj.Series.Item(SeriesItem(i)).Points.DataBindXY(Xary, Yary)
            Next

        Catch ex As Exception
            MessageBox.Show("Error:" & ex.Message, "Error")
        End Try

    End Function

    Function Chart_Query_2(ByVal ChartObj As Windows.Forms.DataVisualization.Charting.Chart, ByVal Chart3D_yn As Boolean, ByVal BarTypeIdx As Integer, ByVal Chart_Type As System.Windows.Forms.DataVisualization.Charting.SeriesChartType, ByVal SeriesItem As String(), ByVal Qtext As String) As Boolean
        ' DrawingStyle : default = 0 , Wedge = 1, Cylinder = 2, Emboss = 3, LightToDark = 4
        ' Qtext의 1st 컬럼은 X축, 2nd 이후 컬럼은 Y축에 표시됨
        ' SeriesItem은 Series의 Item 명을 의미함
        Try

            ChartObj.ChartAreas.Clear()
            ChartObj.Legends.Clear()
            ChartObj.Series.Clear()


            Dim Crs As ADODB.Recordset
            Dim i, j As Integer
            Dim Scnt As Integer = SeriesItem.Length
            Dim Lnd As System.Windows.Forms.DataVisualization.Charting.Legend = New System.Windows.Forms.DataVisualization.Charting.Legend
            Dim CArea As System.Windows.Forms.DataVisualization.Charting.ChartArea = New System.Windows.Forms.DataVisualization.Charting.ChartArea
            Dim Yarys(0, 0) As Integer
            Dim Xary(0) As String
            Dim Yary(0) As Integer
            Dim BarType As String = ""
            Dim FontInfo = New System.Drawing.Font("Verdana", 7.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))

            Lnd.Name = "Lnd"
            Lnd.BackColor = Color.Transparent
            Lnd.Font = FontInfo
            ChartObj.Legends.Add(Lnd)


            ChartObj.BackColor = Color.Transparent
            CArea.Name = "CArea"
            CArea.BackColor = Color.Transparent
            CArea.Area3DStyle.Enable3D = Chart3D_yn
            ChartObj.ChartAreas.Add(CArea)

            CArea.AxisX.IsLabelAutoFit = False
            CArea.AxisY.IsLabelAutoFit = False
            CArea.AxisX2.IsLabelAutoFit = False
            CArea.AxisY2.IsLabelAutoFit = False
            CArea.AxisX.LabelStyle.Font = FontInfo
            CArea.AxisY.LabelStyle.Font = FontInfo
            CArea.AxisX.LabelStyle.Interval = 1

            '         CArea.AxisY.LabelStyle.Interval = 1

            Crs = Query_RS_ALL(Qtext)

            If SeriesItem.Length <> Crs.Fields.Count - 1 Then
                MessageBox.Show("Series's Count must be Query's Column Count", "Validation Error")
                Exit Function
            End If

            If Crs.RecordCount < 1 Then
                MessageBox.Show("Record count 0 ", "Validation Error")
                Exit Function
            End If

            ReDim Xary(Crs.RecordCount - 1)
            ReDim Yary(Crs.RecordCount - 1)
            ReDim Yarys(Scnt - 1, Crs.RecordCount - 1)

            Select Case BarTypeIdx
                Case 0
                    BarType = ""
                Case 1
                    BarType = "DrawingStyle = Wedge"
                Case 2
                    BarType = "DrawingStyle = Cylinder"
                Case 3
                    BarType = "DrawingStyle = Emboss"
                Case 4
                    BarType = "DrawingStyle = LightToDark"
            End Select


            For i = 0 To Crs.RecordCount - 1
                Xary(i) = Crs.Fields(0).Value
                For j = 1 To Scnt
                    Yarys(j - 1, i) = Crs.Fields(j).Value
                Next
                Crs.MoveNext()
            Next

            For i = 0 To Scnt - 1
                For j = 0 To Yary.Length - 1
                    Yary(j) = Yarys(i, j)
                Next
                ChartObj.Series.Add(SeriesItem(i))
                ChartObj.Series.Item(SeriesItem(i)).ChartArea = CArea.Name
                ChartObj.Series.Item(SeriesItem(i)).CustomProperties = BarType
                ChartObj.Series.Item(SeriesItem(i)).ChartType = Chart_Type
                ChartObj.Series.Item(SeriesItem(i)).Font = FontInfo
                ChartObj.Series.Item(SeriesItem(i)).BorderWidth = 2
                ChartObj.Series.Item(SeriesItem(i)).IsValueShownAsLabel = True
                ChartObj.Series.Item(SeriesItem(i)).Points.DataBindXY(Xary, Yary)
            Next

        Catch ex As Exception
            MessageBox.Show("Error:" & ex.Message, "Error")
        End Try

    End Function

    Function Chart_Query2(ByVal ChartObj As Windows.Forms.DataVisualization.Charting.Chart, ByVal Chart3D_yn As Boolean, ByVal BarTypeIdx As Integer, ByVal Chart_Type As System.Windows.Forms.DataVisualization.Charting.SeriesChartType, ByVal SeriesItem As String(), ByVal Interval_init As Integer, ByVal Qtext As String) As Boolean
        ' DrawingStyle : default = 0 , Wedge = 1, Cylinder = 2, Emboss = 3, LightToDark = 4
        ' Qtext의 1st 컬럼은 X축, 2nd 이후 컬럼은 Y축에 표시됨
        ' SeriesItem은 Series의 Item 명을 의미함
        Try

            ChartObj.ChartAreas.Clear()
            ChartObj.Legends.Clear()
            ChartObj.Series.Clear()


            Dim Crs As ADODB.Recordset
            Dim i As Integer
            Dim Scnt As Integer = SeriesItem.Length
            Dim Lnd As System.Windows.Forms.DataVisualization.Charting.Legend = New System.Windows.Forms.DataVisualization.Charting.Legend
            Dim CArea As System.Windows.Forms.DataVisualization.Charting.ChartArea = New System.Windows.Forms.DataVisualization.Charting.ChartArea
            Dim Yarys(0, 0) As Integer
            Dim Xary(0) As String
            Dim Yary(0) As Integer
            Dim BarType As String = ""
            Dim FontInfo = New System.Drawing.Font("Verdana", 7.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))

            Lnd.Name = "Lnd"
            Lnd.BackColor = Color.Transparent
            Lnd.Font = FontInfo
            ChartObj.Legends.Add(Lnd)


            ChartObj.BackColor = Color.Transparent
            CArea.Name = "CArea"
            CArea.BackColor = Color.Transparent
            CArea.Area3DStyle.Enable3D = Chart3D_yn
            ChartObj.ChartAreas.Add(CArea)

            CArea.AxisX.IsLabelAutoFit = True
            CArea.AxisY.IsLabelAutoFit = False
            CArea.AxisX2.IsLabelAutoFit = True
            CArea.AxisY2.IsLabelAutoFit = False
            'CArea.AxisX.LabelStyle.Font = FontInfo
            CArea.AxisY.LabelStyle.Font = FontInfo
            CArea.AxisX.LabelStyle.Interval = 1
            'CArea.AxisY.LabelStyle.Interval = Interval_init
            CArea.AxisX.MinorGrid.Enabled = True
            CArea.AxisX.MinorGrid.Interval = 0.5
            CArea.AxisX.MinorGrid.LineColor = Color.DarkGray
            CArea.AxisY.MinorGrid.Enabled = True
            CArea.AxisY.MinorGrid.Interval = CArea.AxisY.LabelStyle.Interval '
            CArea.AxisY.MinorGrid.LineColor = Color.DarkGray
            '       CArea.AxisX.LabelStyle.IsStaggered = True

            'CArea.AxisY.LabelStyle.Interval = 1

            Crs = Query_RS_ALL(Qtext)

            If Crs.RecordCount < 1 Then
                MessageBox.Show("Record count 0 ", "Validation Error")
                Exit Function
            End If

            ReDim Xary(Crs.Fields.Count - 1)
            ReDim Yary(Crs.Fields.Count - 1)

            Select Case BarTypeIdx
                Case 0
                    BarType = ""
                Case 1
                    BarType = "DrawingStyle = Wedge"
                Case 2
                    BarType = "DrawingStyle = Cylinder"
                Case 3
                    BarType = "DrawingStyle = Emboss"
                Case 4
                    BarType = "DrawingStyle = LightToDark"
            End Select

            For i = 0 To Crs.Fields.Count - 1
                Xary(i) = Crs.Fields(i).Name
                Yary(i) = Crs.Fields(i).Value
            Next

            ChartObj.Series.Add(SeriesItem(0))
            ChartObj.Series.Item(SeriesItem(0)).ChartArea = CArea.Name
            ChartObj.Series.Item(SeriesItem(0)).CustomProperties = BarType
            ChartObj.Series.Item(SeriesItem(0)).ChartType = Chart_Type
            ChartObj.Series.Item(SeriesItem(0)).Font = FontInfo
            ChartObj.Series.Item(SeriesItem(0)).Points.DataBindXY(Xary, Yary)
            ChartObj.Legends.Item(0).Enabled = False


        Catch ex As Exception
            MessageBox.Show("Error:" & ex.Message, "Error")
        End Try

    End Function

    Function Chart_Query3(ByVal ChartObj As Windows.Forms.DataVisualization.Charting.Chart, ByVal Chart3D_yn As Boolean, ByVal BarTypeIdx As Integer, ByVal Chart_Type As System.Windows.Forms.DataVisualization.Charting.SeriesChartType, ByVal SeriesItem As String(), ByVal Interval_init As Integer, ByVal Qtext As String) As Boolean
        ' DrawingStyle : default = 0 , Wedge = 1, Cylinder = 2, Emboss = 3, LightToDark = 4
        ' Qtext의 1st 컬럼은 X축, 2nd 이후 컬럼은 Y축에 표시됨
        ' SeriesItem은 Series의 Item 명을 의미함
        Try

            ChartObj.ChartAreas.Clear()
            ChartObj.Legends.Clear()
            ChartObj.Series.Clear()


            Dim Crs As ADODB.Recordset
            Dim i, j As Integer
            Dim Scnt As Integer = SeriesItem.Length
            Dim Lnd As System.Windows.Forms.DataVisualization.Charting.Legend = New System.Windows.Forms.DataVisualization.Charting.Legend
            Dim CArea As System.Windows.Forms.DataVisualization.Charting.ChartArea = New System.Windows.Forms.DataVisualization.Charting.ChartArea
            Dim Yarys(0, 0) As Integer
            Dim Xary(0) As String
            Dim Yary(0) As Integer
            Dim BarType As String = ""
            Dim FontInfo = New System.Drawing.Font("Verdana", 7.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))

            Lnd.Name = "Lnd"
            Lnd.BackColor = Color.Transparent
            Lnd.Font = FontInfo
            ChartObj.Legends.Add(Lnd)


            ChartObj.BackColor = Color.Transparent
            CArea.Name = "CArea"
            CArea.BackColor = Color.Transparent
            CArea.Area3DStyle.Enable3D = Chart3D_yn
            ChartObj.ChartAreas.Add(CArea)

            CArea.AxisX.IsLabelAutoFit = True
            CArea.AxisY.IsLabelAutoFit = False
            CArea.AxisX2.IsLabelAutoFit = True
            CArea.AxisY2.IsLabelAutoFit = False
            'CArea.AxisX.LabelStyle.Font = FontInfo
            CArea.AxisY.LabelStyle.Font = FontInfo
            CArea.AxisX.LabelStyle.Interval = 1
            'CArea.AxisY.LabelStyle.Interval = Interval_init
            CArea.AxisX.MinorGrid.Enabled = True
            CArea.AxisX.MinorGrid.Interval = 0.5
            CArea.AxisX.MinorGrid.LineColor = Color.DarkGray
            CArea.AxisY.MinorGrid.Enabled = True
            CArea.AxisY.MinorGrid.Interval = CArea.AxisY.LabelStyle.Interval / 2
            CArea.AxisY.MinorGrid.LineColor = Color.DarkGray
            '        CArea.AxisX.LabelStyle.IsStaggered = True



            Crs = Query_RS_ALL(Qtext)

            If SeriesItem.Length <> Crs.Fields.Count - 1 Then
                MessageBox.Show("Series's Count must be Query's Column Count", "Validation Error")
                Exit Function
            End If

            If Crs.RecordCount < 1 Then
                MessageBox.Show("Record count 0 ", "Validation Error")
                Exit Function
            End If

            ReDim Xary(Crs.RecordCount - 1)
            ReDim Yary(Crs.RecordCount - 1)
            ReDim Yarys(Scnt - 1, Crs.RecordCount - 1)

            Select Case BarTypeIdx
                Case 0
                    BarType = ""
                Case 1
                    BarType = "DrawingStyle = Wedge"
                Case 2
                    BarType = "DrawingStyle = Cylinder"
                Case 3
                    BarType = "DrawingStyle = Emboss"
                Case 4
                    BarType = "DrawingStyle = LightToDark"
            End Select


            For i = 0 To Crs.RecordCount - 1
                Xary(i) = Crs.Fields(0).Value
                For j = 1 To Scnt
                    Yarys(j - 1, i) = Crs.Fields(j).Value
                Next
                Crs.MoveNext()
            Next

            ChartObj.Legends(0).Docking = DataVisualization.Charting.Docking.Top
            ChartObj.Legends(0).Alignment = StringAlignment.Center
            ChartObj.Legends(0).BorderDashStyle = DataVisualization.Charting.ChartDashStyle.Solid

            For i = 0 To Scnt - 1
                For j = 0 To Yary.Length - 1
                    Yary(j) = Yarys(i, j)
                Next

                ChartObj.Series.Add(SeriesItem(i))
                ChartObj.Series.Item(SeriesItem(i)).ChartArea = CArea.Name
                ChartObj.Series.Item(SeriesItem(i)).CustomProperties = BarType
                ChartObj.Series.Item(SeriesItem(i)).ChartType = Chart_Type
                'ChartObj.Series.Item(SeriesItem(i)).Font = FontInfo
                ChartObj.Series.Item(SeriesItem(i)).Points.DataBindXY(Xary, Yary)
                ChartObj.Series.Item(SeriesItem(i)).XValueType = DataVisualization.Charting.ChartValueType.String


            Next



        Catch ex As Exception
            MessageBox.Show("Error:" & ex.Message, "Error")
        End Try

    End Function

    Function Chart_Query4(ByVal ChartObj As Windows.Forms.DataVisualization.Charting.Chart, ByVal Chart3D_yn As Boolean, ByVal BarTypeIdx As Integer, ByVal Chart_Type As System.Windows.Forms.DataVisualization.Charting.SeriesChartType, ByVal SeriesItem As String(), ByVal Qtext As String) As Boolean
        ' DrawingStyle : default = 0 , Wedge = 1, Cylinder = 2, Emboss = 3, LightToDark = 4
        ' Qtext의 1st 컬럼은 X축, 2nd 이후 컬럼은 Y축에 표시됨
        ' SeriesItem은 Series의 Item 명을 의미함
        Try

            ChartObj.ChartAreas.Clear()
            ChartObj.Legends.Clear()
            ChartObj.Series.Clear()


            Dim Crs As ADODB.Recordset
            Dim i, j As Integer
            Dim Scnt As Integer = SeriesItem.Length
            Dim Lnd As System.Windows.Forms.DataVisualization.Charting.Legend = New System.Windows.Forms.DataVisualization.Charting.Legend
            Dim CArea As System.Windows.Forms.DataVisualization.Charting.ChartArea = New System.Windows.Forms.DataVisualization.Charting.ChartArea
            Dim Yarys(0, 0) As Decimal
            Dim Xary(0) As String
            Dim Yary(0) As Decimal
            Dim BarType As String = ""
            Dim FontInfo = New System.Drawing.Font("Verdana", 7.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))

            Lnd.Name = "Lnd"
            Lnd.BackColor = Color.Transparent
            Lnd.Font = FontInfo
            ChartObj.Legends.Add(Lnd)


            ChartObj.BackColor = Color.Transparent
            CArea.Name = "CArea"
            CArea.BackColor = Color.Transparent
            CArea.Area3DStyle.Enable3D = Chart3D_yn
            ChartObj.ChartAreas.Add(CArea)

            CArea.AxisX.IsLabelAutoFit = True
            CArea.AxisY.IsLabelAutoFit = False
            CArea.AxisX2.IsLabelAutoFit = True
            CArea.AxisY2.IsLabelAutoFit = False
            CArea.AxisX.LabelStyle.Font = FontInfo
            CArea.AxisY.LabelStyle.Font = FontInfo
            'CArea.AxisX.LabelStyle.Interval = 1
            'CArea.AxisY.LabelStyle.Interval = 2
            'CArea.AxisX.MinorGrid.Enabled = True
            'CArea.AxisX.MinorGrid.Interval = 1
            'CArea.AxisX.MinorGrid.LineColor = Color.DarkGray
            'CArea.AxisY.MinorGrid.Enabled = True
            'CArea.AxisY.MinorGrid.Interval = CArea.AxisY.LabelStyle.Interval / 2
            'CArea.AxisY.MinorGrid.LineColor = Color.DarkGray
            'CArea.AxisY.LabelStyle.Interval = 1
            '         CArea.AxisX.LabelStyle.IsStaggered = True

            Crs = Query_RS_ALL(Qtext)

            If SeriesItem.Length <> Crs.Fields.Count - 1 Then
                MessageBox.Show("Series's Count must be Query's Column Count", "Validation Error")
                Exit Function
            End If

            If Crs.RecordCount < 1 Then
                MessageBox.Show("Record count 0 ", "Validation Error")
                Exit Function
            End If

            ReDim Xary(Crs.RecordCount - 1)
            ReDim Yary(Crs.RecordCount - 1)
            ReDim Yarys(Scnt - 1, Crs.RecordCount - 1)

            Select Case BarTypeIdx
                Case 0
                    BarType = ""
                Case 1
                    BarType = "DrawingStyle = Wedge"
                Case 2
                    BarType = "DrawingStyle = Cylinder"
                Case 3
                    BarType = "DrawingStyle = Emboss"
                Case 4
                    BarType = "DrawingStyle = LightToDark"
            End Select


            For i = 0 To Crs.RecordCount - 1
                Xary(i) = Crs.Fields(0).Value
                For j = 1 To Scnt
                    Yarys(j - 1, i) = Crs.Fields(j).Value
                Next
                Crs.MoveNext()
            Next

            For i = 0 To Scnt - 1
                For j = 0 To Yary.Length - 1
                    Yary(j) = Yarys(i, j)
                Next
                ChartObj.Series.Add(SeriesItem(i))
                ChartObj.Series.Item(SeriesItem(i)).ChartArea = CArea.Name
                ChartObj.Series.Item(SeriesItem(i)).CustomProperties = BarType
                ChartObj.Series.Item(SeriesItem(i)).ChartType = Chart_Type
                ChartObj.Series.Item(SeriesItem(i)).Font = FontInfo
                ChartObj.Series.Item(SeriesItem(i)).Points.DataBindXY(Xary, Yary)
            Next

        Catch ex As Exception
            MessageBox.Show("Error:" & ex.Message, "Error")
        End Try

    End Function

    Function Chart_Query5(ByVal ChartObj As Windows.Forms.DataVisualization.Charting.Chart, ByVal Chart3D_yn As Boolean, ByVal BarTypeIdx As Integer, ByVal Chart_Type As System.Windows.Forms.DataVisualization.Charting.SeriesChartType, ByVal SeriesItem As String(), ByVal title As String, ByVal xlabel_rad As Boolean, ByVal Qtext As String) As Boolean
        ' DrawingStyle : default = 0 , Wedge = 1, Cylinder = 2, Emboss = 3, LightToDark = 4
        ' Qtext의 1st 컬럼은 X축, 2nd 이후 컬럼은 Y축에 표시됨
        ' SeriesItem은 Series의 Item 명을 의미함
        Try

            ChartObj.ChartAreas.Clear()
            ChartObj.Legends.Clear()
            ChartObj.Series.Clear()
            ChartObj.Titles.Clear()


            Dim Crs As ADODB.Recordset
            Dim i, j As Integer
            Dim Scnt As Integer = SeriesItem.Length
            Dim Lnd As System.Windows.Forms.DataVisualization.Charting.Legend = New System.Windows.Forms.DataVisualization.Charting.Legend
            Dim CArea As System.Windows.Forms.DataVisualization.Charting.ChartArea = New System.Windows.Forms.DataVisualization.Charting.ChartArea
            Dim Yarys(0, 0) As Integer
            Dim Xary(0) As String
            Dim Yary(0) As Integer
            Dim BarType As String = ""
            Dim FontInfo = New System.Drawing.Font("Verdana", 7.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))

            Lnd.Name = "Lnd"
            Lnd.BackColor = Color.Transparent
            Lnd.Font = FontInfo
            ChartObj.Legends.Add(Lnd)

            ChartObj.Titles.Add(title)
            ChartObj.Titles(0).Font = New System.Drawing.Font("Verdana", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
            ChartObj.BackColor = Color.Transparent
            CArea.Name = "CArea"
            CArea.BackColor = Color.Transparent
            CArea.Area3DStyle.Enable3D = Chart3D_yn
            ChartObj.ChartAreas.Add(CArea)
            If xlabel_rad = True Then
                CArea.AxisX.LabelStyle.Angle = 30
                CArea.AxisX.LabelStyle.IsStaggered = True
            End If
            CArea.AxisX.IsLabelAutoFit = False
            CArea.AxisY.IsLabelAutoFit = False
            CArea.AxisX2.IsLabelAutoFit = False
            CArea.AxisY2.IsLabelAutoFit = False
            CArea.AxisX.LabelStyle.Font = FontInfo
            CArea.AxisY.LabelStyle.Font = FontInfo
            CArea.AxisX.LabelStyle.Interval = 1
            '      CArea.AxisX.LabelStyle.IsStaggered = True
            '         CArea.AxisY.LabelStyle.Interval = 1

            Crs = Query_RS_ALL(Qtext)

            If SeriesItem.Length <> Crs.Fields.Count - 1 Then
                MessageBox.Show("Series's Count must be Query's Column Count", "Validation Error")
                Exit Function
            End If

            If Crs.RecordCount < 1 Then
                MessageBox.Show("Record count 0 ", "Validation Error")
                Exit Function
            End If

            ReDim Xary(Crs.RecordCount - 1)
            ReDim Yary(Crs.RecordCount - 1)
            ReDim Yarys(Scnt - 1, Crs.RecordCount - 1)

            Select Case BarTypeIdx
                Case 0
                    BarType = ""
                Case 1
                    BarType = "DrawingStyle = Wedge"
                Case 2
                    BarType = "DrawingStyle = Cylinder"
                Case 3
                    BarType = "DrawingStyle = Emboss"
                Case 4
                    BarType = "DrawingStyle = LightToDark"
            End Select


            For i = 0 To Crs.RecordCount - 1
                Xary(i) = Crs.Fields(0).Value
                For j = 1 To Scnt
                    Yarys(j - 1, i) = Crs.Fields(j).Value
                Next
                Crs.MoveNext()
            Next

            For i = 0 To Scnt - 1
                For j = 0 To Yary.Length - 1
                    Yary(j) = Yarys(i, j)
                Next
                ChartObj.Series.Add(SeriesItem(i))
                ChartObj.Series.Item(SeriesItem(i)).ChartArea = CArea.Name
                ChartObj.Series.Item(SeriesItem(i)).CustomProperties = BarType
                ChartObj.Series.Item(SeriesItem(i)).ChartType = Chart_Type
                ChartObj.Series.Item(SeriesItem(i)).Font = FontInfo
                ChartObj.Series.Item(SeriesItem(i)).BorderWidth = 2
                ChartObj.Series.Item(SeriesItem(i)).MarkerStyle = DataVisualization.Charting.MarkerStyle.Diamond
                ChartObj.Series.Item(SeriesItem(i)).MarkerSize = 7
                ChartObj.Series.Item(SeriesItem(i)).Points.DataBindXY(Xary, Yary)
            Next

        Catch ex As Exception
            MessageBox.Show("Error:" & ex.Message, "Error")
        End Try

    End Function

    Function Chart_Spread_All(ByVal ChartObj As Windows.Forms.DataVisualization.Charting.Chart, ByVal Chart3D_yn As Boolean, ByVal BarTypeIdx As Integer, ByVal Chart_Type As System.Windows.Forms.DataVisualization.Charting.SeriesChartType, ByVal S As FarPoint.Win.Spread.FpSpread, ByVal Seriesitem As String(), ByVal XColidx As Integer, ByVal YColidx As Integer(), ByVal LastRowHidden As Boolean) As Boolean
        ' DrawingStyle : default = 0 , Wedge = 1, Cylinder = 2, Emboss = 3, LightToDark = 4
        ' SeriesItem은 Series의 Item 명을 의미함, 만일 스프레드의 컬럼명과 다른 이름으로 표시하는 경우, new String() {'item1','item2'} 식으로 나열하면 됨
        Try

            ChartObj.ChartAreas.Clear()
            ChartObj.Legends.Clear()
            ChartObj.Series.Clear()

            Dim i, j, Rcnt As Integer
            Dim SeriesNm As String = ""
            Dim Lnd As System.Windows.Forms.DataVisualization.Charting.Legend = New System.Windows.Forms.DataVisualization.Charting.Legend
            Dim CArea As System.Windows.Forms.DataVisualization.Charting.ChartArea = New System.Windows.Forms.DataVisualization.Charting.ChartArea
            Dim Yarys(0, 0) As Integer
            Dim Xary(0) As String
            Dim Yary(0) As Integer
            Dim BarType As String = ""
            Dim FontInfo = New System.Drawing.Font("verdana", 7.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))

            Lnd.Name = "Lnd"
            Lnd.BackColor = Color.Transparent
            Lnd.Font = FontInfo
            ChartObj.Legends.Add(Lnd)

            ChartObj.BackColor = Color.Transparent
            CArea.Name = "CArea"
            CArea.BackColor = Color.Transparent
            CArea.Area3DStyle.Enable3D = Chart3D_yn
            ChartObj.ChartAreas.Add(CArea)

            CArea.AxisX.IsLabelAutoFit = True
            CArea.AxisY.IsLabelAutoFit = False
            CArea.AxisX2.IsLabelAutoFit = True
            CArea.AxisY2.IsLabelAutoFit = True
            CArea.AxisX.LabelStyle.Font = FontInfo
            CArea.AxisY.LabelStyle.Font = FontInfo
            CArea.AxisX.LabelStyle.Interval = 1
            '          CArea.AxisY.LabelStyle.Interval = 1
            'CArea.AxisX.LabelStyle.Angle = 30
            '          CArea.AxisX.LabelStyle.IsStaggered = True
            'CArea.AxisY.LabelStyle.Interval = 2
            CArea.AxisX.MinorGrid.Enabled = True
            CArea.AxisX.MinorGrid.Interval = 1
            CArea.AxisX.MinorGrid.LineColor = Color.DarkGray
            'CArea.AxisY.MinorGrid.Enabled = True
            'CArea.AxisY.MinorGrid.Interval = CArea.AxisY.LabelStyle.Interval / 2
            'CArea.AxisY.MinorGrid.LineColor = Color.DarkGray
            'CArea.AxisY.LabelStyle.Interval = 1




            If S.ActiveSheet.RowCount < 1 Then
                Exit Function
            End If

            Select Case BarTypeIdx
                Case 0
                    BarType = ""
                Case 1
                    BarType = "DrawingStyle = Wedge"
                Case 2
                    BarType = "DrawingStyle = Cylinder"
                Case 3
                    BarType = "DrawingStyle = Emboss"
                Case 4
                    BarType = "DrawingStyle = LightToDark"
            End Select

            If LastRowHidden = True Then
                Rcnt = S.ActiveSheet.RowCount - 2
            Else
                Rcnt = S.ActiveSheet.RowCount - 1
            End If

            ReDim Xary(Rcnt)
            ReDim Yary(Rcnt)



            For i = 0 To YColidx.Length - 1
                For j = 0 To Rcnt
                    Xary(j) = S.ActiveSheet.GetValue(j, XColidx)
                    Yary(j) = S.ActiveSheet.GetValue(j, YColidx(i))
                Next
                If Seriesitem.Length > 0 Then
                    SeriesNm = Seriesitem(i)
                Else
                    SeriesNm = S.ActiveSheet.ColumnHeader.Columns(YColidx(i)).Label
                End If
                ChartObj.Series.Add(SeriesNm)
                ChartObj.Series.Item(SeriesNm).ChartArea = CArea.Name
                ChartObj.Series.Item(SeriesNm).CustomProperties = BarType
                ChartObj.Series.Item(SeriesNm).ChartType = Chart_Type
                ChartObj.Series.Item(SeriesNm).Font = FontInfo
                ChartObj.Series.Item(SeriesNm).Points.DataBindXY(Xary, Yary)
            Next

        Catch ex As Exception
            MessageBox.Show("Error:" & ex.Message, "Error")
        End Try

    End Function



    Function Chart_Spread_All_1(ByVal ChartObj As Windows.Forms.DataVisualization.Charting.Chart, ByVal Chart3D_yn As Boolean, ByVal BarTypeIdx As Integer, ByVal Chart_Type As System.Windows.Forms.DataVisualization.Charting.SeriesChartType, ByVal S As FarPoint.Win.Spread.FpSpread, ByVal Seriesitem As String(), ByVal XColidx As Integer, ByVal YColidx As Integer(), ByVal LastRowHidden As Boolean) As Boolean
        ' DrawingStyle : default = 0 , Wedge = 1, Cylinder = 2, Emboss = 3, LightToDark = 4
        ' SeriesItem은 Series의 Item 명을 의미함, 만일 스프레드의 컬럼명과 다른 이름으로 표시하는 경우, new String() {'item1','item2'} 식으로 나열하면 됨
        Try

            ChartObj.ChartAreas.Clear()
            ChartObj.Legends.Clear()
            ChartObj.Series.Clear()

            Dim i, j, Rcnt As Integer
            Dim SeriesNm As String = ""
            Dim Lnd As System.Windows.Forms.DataVisualization.Charting.Legend = New System.Windows.Forms.DataVisualization.Charting.Legend
            Dim CArea As System.Windows.Forms.DataVisualization.Charting.ChartArea = New System.Windows.Forms.DataVisualization.Charting.ChartArea
            Dim Yarys(0, 0) As Integer
            Dim Xary(0) As String
            Dim Yary(0) As Integer
            Dim BarType As String = ""
            Dim FontInfo = New System.Drawing.Font("arial", 6.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))

            Lnd.Name = "Lnd"
            Lnd.BackColor = Color.Transparent
            Lnd.Font = FontInfo
            ChartObj.Legends.Add(Lnd)

            ChartObj.BackColor = Color.Transparent
            CArea.Name = "CArea"
            CArea.BackColor = Color.Transparent
            CArea.Area3DStyle.Enable3D = Chart3D_yn
            ChartObj.ChartAreas.Add(CArea)

            CArea.AxisX.IsLabelAutoFit = True
            CArea.AxisY.IsLabelAutoFit = False
            CArea.AxisX2.IsLabelAutoFit = True
            CArea.AxisY2.IsLabelAutoFit = True
            CArea.AxisX.LabelStyle.Font = FontInfo
            CArea.AxisY.LabelStyle.Font = FontInfo
            CArea.AxisX.LabelStyle.Interval = 1
            CArea.AxisX.MinorGrid.Enabled = True
            CArea.AxisX.MinorGrid.Interval = 0.5
            CArea.AxisX.MinorGrid.LineColor = Color.DarkGray
            CArea.AxisY.MinorGrid.Enabled = True
            CArea.AxisY.MinorGrid.Interval = CArea.AxisY.LabelStyle.Interval '
            CArea.AxisY.MinorGrid.LineColor = Color.DarkGray
            '          CArea.AxisY.LabelStyle.Interval = 1
            CArea.AxisX.LabelStyle.Angle = 30
            CArea.AxisX.LabelStyle.IsStaggered = True
            'CArea.AxisY.LabelStyle.Interval = 2
            CArea.AxisX.MinorGrid.Enabled = True
            CArea.AxisX.MinorGrid.Interval = 1
            CArea.AxisX.MinorGrid.LineColor = Color.DarkGray
            'CArea.AxisY.MinorGrid.Enabled = True
            'CArea.AxisY.MinorGrid.Interval = CArea.AxisY.LabelStyle.Interval / 2
            'CArea.AxisY.MinorGrid.LineColor = Color.DarkGray
            'CArea.AxisY.LabelStyle.Interval = 1




            If S.ActiveSheet.RowCount < 1 Then
                Exit Function
            End If

            Select Case BarTypeIdx
                Case 0
                    BarType = ""
                Case 1
                    BarType = "DrawingStyle = Wedge"
                Case 2
                    BarType = "DrawingStyle = Cylinder"
                Case 3
                    BarType = "DrawingStyle = Emboss"
                Case 4
                    BarType = "DrawingStyle = LightToDark"
            End Select

            If LastRowHidden = True Then
                Rcnt = S.ActiveSheet.RowCount - 2
            Else
                Rcnt = S.ActiveSheet.RowCount - 1
            End If

            ReDim Xary(Rcnt)
            ReDim Yary(Rcnt)



            For i = 0 To YColidx.Length - 1
                For j = 0 To Rcnt
                    Xary(j) = S.ActiveSheet.GetValue(j, XColidx)
                    Yary(j) = S.ActiveSheet.GetValue(j, YColidx(i))
                Next
                If Seriesitem.Length > 0 Then
                    SeriesNm = Seriesitem(i)
                Else
                    SeriesNm = S.ActiveSheet.ColumnHeader.Columns(YColidx(i)).Label
                End If
                ChartObj.Series.Add(SeriesNm)
                ChartObj.Series.Item(SeriesNm).ChartArea = CArea.Name
                ChartObj.Series.Item(SeriesNm).CustomProperties = BarType
                ChartObj.Series.Item(SeriesNm).ChartType = Chart_Type
                ChartObj.Series.Item(SeriesNm).Font = FontInfo
                ChartObj.Series.Item(SeriesNm).Points.DataBindXY(Xary, Yary)
            Next

        Catch ex As Exception
            MessageBox.Show("Error:" & ex.Message, "Error")
        End Try

    End Function



    Function Chart_Spread_All2(ByVal ChartObj As Windows.Forms.DataVisualization.Charting.Chart, ByVal Chart3D_yn As Boolean, ByVal BarTypeIdx As Integer, ByVal Chart_Type As System.Windows.Forms.DataVisualization.Charting.SeriesChartType, ByVal S As FarPoint.Win.Spread.FpSpread, ByVal Seriesitem As String(), ByVal XColidx As String(), ByVal YColidx As Integer(), ByVal dispRow As Integer()) As Boolean
        ' DrawingStyle : default = 0 , Wedge = 1, Cylinder = 2, Emboss = 3, LightToDark = 4
        ' SeriesItem은 Series의 Item 명을 의미함, 만일 스프레드의 컬럼명과 다른 이름으로 표시하는 경우, new String() {'item1','item2'} 식으로 나열하면 됨
        Try

            ChartObj.ChartAreas.Clear()
            ChartObj.Legends.Clear()
            ChartObj.Series.Clear()

            Dim i, j, Rcnt As Integer
            Dim SeriesNm As String = ""
            Dim Lnd As System.Windows.Forms.DataVisualization.Charting.Legend = New System.Windows.Forms.DataVisualization.Charting.Legend
            Dim CArea As System.Windows.Forms.DataVisualization.Charting.ChartArea = New System.Windows.Forms.DataVisualization.Charting.ChartArea
            Dim Yarys(0, 0) As Integer
            Dim Xary(0) As String
            Dim Yary(0) As Integer
            Dim BarType As String = ""
            Dim FontInfo = New System.Drawing.Font("verdana", 7.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))

            Lnd.Name = "Lnd"
            Lnd.BackColor = Color.Transparent
            Lnd.Font = FontInfo
            ChartObj.Legends.Add(Lnd)

            ChartObj.BackColor = Color.Transparent
            CArea.Name = "CArea"
            CArea.BackColor = Color.Transparent
            CArea.Area3DStyle.Enable3D = Chart3D_yn
            ChartObj.ChartAreas.Add(CArea)

            CArea.AxisX.IsLabelAutoFit = True
            CArea.AxisY.IsLabelAutoFit = False
            CArea.AxisX2.IsLabelAutoFit = True
            CArea.AxisY2.IsLabelAutoFit = True
            CArea.AxisX.LabelStyle.Font = FontInfo
            CArea.AxisY.LabelStyle.Font = FontInfo
            CArea.AxisX.LabelStyle.Interval = 1
            '          CArea.AxisY.LabelStyle.Interval = 1
            'CArea.AxisX.LabelStyle.Angle = 30
            '          CArea.AxisX.LabelStyle.IsStaggered = True
            'CArea.AxisY.LabelStyle.Interval = 2
            CArea.AxisX.MinorGrid.Enabled = True
            CArea.AxisX.MinorGrid.Interval = 1
            CArea.AxisX.MinorGrid.LineColor = Color.DarkGray
            'CArea.AxisY.MinorGrid.Enabled = True
            'CArea.AxisY.MinorGrid.Interval = CArea.AxisY.LabelStyle.Interval / 2
            'CArea.AxisY.MinorGrid.LineColor = Color.DarkGray
            'CArea.AxisY.LabelStyle.Interval = 1




            If S.ActiveSheet.RowCount < 1 Then
                Exit Function
            End If

            Select Case BarTypeIdx
                Case 0
                    BarType = ""
                Case 1
                    BarType = "DrawingStyle = Wedge"
                Case 2
                    BarType = "DrawingStyle = Cylinder"
                Case 3
                    BarType = "DrawingStyle = Emboss"
                Case 4
                    BarType = "DrawingStyle = LightToDark"
            End Select

            'If LastRowHidden = True Then
            '    Rcnt = S.ActiveSheet.RowCount - 2
            'Else
            '    Rcnt = S.ActiveSheet.RowCount - 1
            'End If

            Rcnt = dispRow.Length

            ReDim Xary(YColidx.Length - 1)
            ReDim Yary(YColidx.Length - 1)



            For i = 0 To Rcnt - 1
                For j = 0 To YColidx.Length - 1
                    Xary(j) = XColidx(j)
                    Yary(j) = S.ActiveSheet.GetValue(dispRow(i), YColidx(j))
                Next
                If Seriesitem.Length > 0 Then
                    SeriesNm = Seriesitem(i)
                Else
                    SeriesNm = S.ActiveSheet.ColumnHeader.Columns(YColidx(i)).Label
                End If
                ChartObj.Series.Add(SeriesNm)
                ChartObj.Series.Item(SeriesNm).ChartArea = CArea.Name
                ChartObj.Series.Item(SeriesNm).CustomProperties = BarType
                ChartObj.Series.Item(SeriesNm).ChartType = Chart_Type
                ChartObj.Series.Item(SeriesNm).Font = FontInfo
                ChartObj.Series.Item(SeriesNm).Points.DataBindXY(Xary, Yary)
            Next

        Catch ex As Exception
            MessageBox.Show("Error:" & ex.Message, "Error")
        End Try

    End Function

    Function Chart_Spread_Col(ByVal ChartObj As Windows.Forms.DataVisualization.Charting.Chart, ByVal Chart3D_yn As Boolean, ByVal BarTypeIdx As Integer, ByVal Chart_Type As System.Windows.Forms.DataVisualization.Charting.SeriesChartType, ByVal S As FarPoint.Win.Spread.FpSpread, ByVal Seriesitem As String, ByVal Rowidx As Integer, ByVal Colidx As Integer(), ByVal Headeridx As Integer) As Boolean
        ' DrawingStyle : default = 0 , Wedge = 1, Cylinder = 2, Emboss = 3, LightToDark = 4
        ' SeriesItem은 Series의 Item 명을 의미함, 만일 스프레드의 컬럼명과 다른 이름으로 표시하는 경우, new String() {'item1','item2'} 식으로 나열하면 됨
        Try

            ChartObj.ChartAreas.Clear()
            ChartObj.Legends.Clear()
            ChartObj.Series.Clear()

            Dim i, Rcnt As Integer
            Dim SeriesNm As String = ""
            Dim Lnd As System.Windows.Forms.DataVisualization.Charting.Legend = New System.Windows.Forms.DataVisualization.Charting.Legend
            Dim CArea As System.Windows.Forms.DataVisualization.Charting.ChartArea = New System.Windows.Forms.DataVisualization.Charting.ChartArea
            Dim Xary(0) As String
            Dim Yary(0) As Integer
            Dim BarType As String = ""
            Dim FontInfo = New System.Drawing.Font("Verdana", 7.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))

            Lnd.Name = "Lnd"
            Lnd.BackColor = Color.Transparent
            Lnd.Font = FontInfo
            ChartObj.Legends.Add(Lnd)

            ChartObj.BackColor = Color.Transparent
            CArea.Name = "CArea"
            CArea.BackColor = Color.Transparent
            CArea.Area3DStyle.Enable3D = Chart3D_yn
            ChartObj.ChartAreas.Add(CArea)

            CArea.AxisX.IsLabelAutoFit = False
            CArea.AxisY.IsLabelAutoFit = False
            CArea.AxisX2.IsLabelAutoFit = False
            CArea.AxisY2.IsLabelAutoFit = False
            CArea.AxisX.LabelStyle.Font = FontInfo
            CArea.AxisY.LabelStyle.Font = FontInfo
            CArea.AxisX.LabelStyle.Interval = 1
            '            CArea.AxisY.LabelStyle.Interval = 1
            '           CArea.AxisX.LabelStyle.IsStaggered = True


            If S.ActiveSheet.RowCount < 1 Then
                Exit Function
            End If

            Select Case BarTypeIdx
                Case 0
                    BarType = ""
                Case 1
                    BarType = "DrawingStyle = Wedge"
                Case 2
                    BarType = "DrawingStyle = Cylinder"
                Case 3
                    BarType = "DrawingStyle = Emboss"
                Case 4
                    BarType = "DrawingStyle = LightToDark"
            End Select

            Rcnt = Colidx.Length - 1

            ReDim Xary(Rcnt)
            ReDim Yary(Rcnt)

            For i = 0 To Rcnt
                'If S.ActiveSheet.GetValue(Rowidx, Colidx(i)) > 0 Then
                Xary(i) = S.ActiveSheet.ColumnHeader.Cells(Headeridx, Colidx(i)).Text
                Yary(i) = S.ActiveSheet.GetValue(Rowidx, Colidx(i))
                'End If
            Next

            SeriesNm = Seriesitem
            ChartObj.Series.Add(SeriesNm)
            ChartObj.Series.Item(SeriesNm).ChartArea = CArea.Name
            ChartObj.Series.Item(SeriesNm).CustomProperties = "DrawingStyle = Cylinder"  'BarType
            ChartObj.Series.Item(SeriesNm).ChartType = DataVisualization.Charting.SeriesChartType.Column    'Chart_Type
            ChartObj.Series.Item(SeriesNm).Font = FontInfo
            ChartObj.Series.Item(SeriesNm).Points.DataBindXY(Xary, Yary)

            'ChartObj.Series.Item(SeriesNm).IsValueShownAsLabel = False


        Catch ex As Exception
            MessageBox.Show("Error:" & ex.Message, "Error")
        End Try

    End Function

    Function Chart_Print(ByVal ChartObj As Windows.Forms.DataVisualization.Charting.Chart, ByVal PageLandscapeY As Boolean) As Boolean
        ChartObj.Printing.PrintDocument.DefaultPageSettings.Landscape = PageLandscapeY
        ChartObj.Printing.Print(True)
    End Function

    Function Chart_Query_main(ByVal ChartObj As Windows.Forms.DataVisualization.Charting.Chart, ByVal Chart3D_yn As Boolean, ByVal BarTypeIdx As Integer, ByVal Chart_Type As System.Windows.Forms.DataVisualization.Charting.SeriesChartType, ByVal SeriesItem As String(), ByVal Interval_init As Integer, ByVal Qtext As String) As Boolean
        ' DrawingStyle : default = 0 , Wedge = 1, Cylinder = 2, Emboss = 3, LightToDark = 4
        ' Qtext의 1st 컬럼은 X축, 2nd 이후 컬럼은 Y축에 표시됨
        ' SeriesItem은 Series의 Item 명을 의미함
        Try

            ChartObj.ChartAreas.Clear()
            ChartObj.Legends.Clear()
            ChartObj.Series.Clear()


            Dim Crs, Crs2 As ADODB.Recordset
            Dim i, j As Integer
            Dim Scnt As Integer = SeriesItem.Length
            Dim Lnd As System.Windows.Forms.DataVisualization.Charting.Legend = New System.Windows.Forms.DataVisualization.Charting.Legend
            Dim CArea As System.Windows.Forms.DataVisualization.Charting.ChartArea = New System.Windows.Forms.DataVisualization.Charting.ChartArea
            Dim Yarys(0, 0) As Integer
            Dim Xary(0) As String
            Dim Yary(0) As Integer
            Dim BarType As String = ""
            Dim FontInfo = New System.Drawing.Font("Verdana", 7.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))

            Lnd.Name = "Lnd"
            Lnd.BackColor = Color.Transparent
            Lnd.Font = FontInfo
            ChartObj.Legends.Add(Lnd)


            ChartObj.BackColor = Color.Transparent
            CArea.Name = "CArea"
            CArea.BackColor = Color.Transparent
            CArea.Area3DStyle.Enable3D = Chart3D_yn
            ChartObj.ChartAreas.Add(CArea)

            CArea.AxisX.IsLabelAutoFit = True
            CArea.AxisY.IsLabelAutoFit = False
            CArea.AxisX2.IsLabelAutoFit = True
            CArea.AxisY2.IsLabelAutoFit = False
            'CArea.AxisX.LabelStyle.Font = FontInfo
            CArea.AxisY.LabelStyle.Font = FontInfo
            CArea.AxisX.LabelStyle.Interval = 1
            'CArea.AxisY.LabelStyle.Interval = Interval_init
            CArea.AxisX.MinorGrid.Enabled = True
            CArea.AxisX.MinorGrid.Interval = 0.5
            CArea.AxisX.MinorGrid.LineColor = Color.DarkGray
            CArea.AxisY.MinorGrid.Enabled = True
            CArea.AxisY.MinorGrid.Interval = CArea.AxisY.LabelStyle.Interval / 2
            CArea.AxisY.MinorGrid.LineColor = Color.DarkGray
            '        CArea.AxisX.LabelStyle.IsStaggered = True



            Crs = Query_RS_ALL(Qtext)

            If SeriesItem.Length <> Crs.Fields.Count - 1 Then
                MessageBox.Show("Series's Count must be Query's Column Count", "Validation Error")
                Exit Function
            End If

            If Crs.RecordCount < 1 Then
                MessageBox.Show("Record count 0 ", "Validation Error")
                Exit Function
            End If


            ReDim Xary(Crs.RecordCount - 1)
            ReDim Yary(Crs.RecordCount - 1)
            ReDim Yarys(Scnt - 1, Crs.RecordCount - 1)

            Select Case BarTypeIdx
                Case 0
                    BarType = ""
                Case 1
                    BarType = "DrawingStyle = Wedge"
                Case 2
                    BarType = "DrawingStyle = Cylinder"
                Case 3
                    BarType = "DrawingStyle = Emboss"
                Case 4
                    BarType = "DrawingStyle = LightToDark"
            End Select


            For i = 0 To Crs.RecordCount - 1
                Xary(i) = Crs.Fields(0).Value
                For j = 1 To Scnt
                    Yarys(j - 1, i) = Crs.Fields(j).Value
                Next
                Crs.MoveNext()
            Next

            ChartObj.Legends(0).Docking = DataVisualization.Charting.Docking.Top
            ChartObj.Legends(0).Alignment = StringAlignment.Center
            ChartObj.Legends(0).BorderDashStyle = DataVisualization.Charting.ChartDashStyle.Solid

            For i = 0 To Scnt - 1
                For j = 0 To Yary.Length - 1
                    Yary(j) = Yarys(i, j)
                Next

                ChartObj.Series.Add(SeriesItem(i))
                ChartObj.Series.Item(SeriesItem(i)).ChartArea = CArea.Name
                ChartObj.Series.Item(SeriesItem(i)).CustomProperties = BarType
                ChartObj.Series.Item(SeriesItem(i)).ChartType = Chart_Type
                'ChartObj.Series.Item(SeriesItem(i)).Font = FontInfo
                ChartObj.Series.Item(SeriesItem(i)).Points.DataBindXY(Xary, Yary)
                ChartObj.Series.Item(SeriesItem(i)).XValueType = DataVisualization.Charting.ChartValueType.String


            Next



            '======series2 추가
            Dim Yarys2(0, 0) As Integer
            Dim Xary2(0) As String
            Dim Yary2(0) As Integer

            'CArea2.Name = "CArea2"
            'CArea2.BackColor = Color.Transparent
            'CArea2.Area3DStyle.Enable3D = Chart3D_yn
            'ChartObj.ChartAreas.Add(CArea2)

            'CArea2.AxisX.IsLabelAutoFit = True
            'CArea2.AxisY.IsLabelAutoFit = False
            'CArea2.AxisX2.IsLabelAutoFit = True
            'CArea2.AxisY2.IsLabelAutoFit = False
            ''CArea.AxisX.LabelStyle.Font = FontInfo
            'CArea2.AxisY.LabelStyle.Font = FontInfo
            'CArea2.AxisX.LabelStyle.Interval = 1
            ''CArea.AxisY.LabelStyle.Interval = Interval_init
            'CArea2.AxisX.MinorGrid.Enabled = True
            'CArea2.AxisX.MinorGrid.Interval = 0.5
            'CArea2.AxisX.MinorGrid.LineColor = Color.DarkGray
            'CArea2.AxisY.MinorGrid.Enabled = True
            'CArea2.AxisY.MinorGrid.Interval = CArea.AxisY.LabelStyle.Interval / 2
            'CArea2.AxisY.MinorGrid.LineColor = Color.DarkGray

            Crs2 = Query_RS_ALL("EXEC sp_FRMMAINDISP '" & Site_id & "',8")

            'If SeriesItem.Length <> Crs.Fields.Count - 1 Then
            '    MessageBox.Show("Series's Count must be Query's Column Count", "Validation Error")
            '    Exit Function
            'End If

            'If Crs.RecordCount < 1 Then
            '    MessageBox.Show("Record count 0 ", "Validation Error")
            '    Exit Function
            'End If


            ReDim Xary2(Crs2.RecordCount - 1)
            ReDim Yary2(Crs2.RecordCount - 1)
            ReDim Yarys2(Scnt - 1, Crs2.RecordCount - 1)

            'Select Case BarTypeIdx
            '    Case 0
            '        BarType = ""
            '    Case 1
            '        BarType = "DrawingStyle = Wedge"
            '    Case 2
            BarType = "DrawingStyle = Cylinder"
            '    Case 3
            '        BarType = "DrawingStyle = Emboss"
            '    Case 4
            '        BarType = "DrawingStyle = LightToDark"
            'End Select


            For i = 0 To Crs2.RecordCount - 1
                Xary2(i) = Crs2.Fields(0).Value
                For j = 1 To Scnt
                    Yarys2(j - 1, i) = Crs2.Fields(j).Value
                Next
                Crs2.MoveNext()
            Next

            'ChartObj.Legends(0).Docking = DataVisualization.Charting.Docking.Top
            'ChartObj.Legends(0).Alignment = StringAlignment.Center
            'ChartObj.Legends(0).BorderDashStyle = DataVisualization.Charting.ChartDashStyle.Solid

            For i = 0 To Scnt - 1
                For j = 0 To Yary2.Length - 1
                    Yary2(j) = Yarys2(i, j)
                Next

                Dim aa As String
                aa = ""

                If i = 0 Then
                    aa = "Triage"
                ElseIf i = 1 Then
                    aa = "DISASSEMBLY"
                ElseIf i = 2 Then
                    aa = "ASSEMBLY"
                ElseIf i = 3 Then
                    aa = "DOWNLOAD"
                ElseIf i = 4 Then
                    aa = "QC"
                ElseIf i = 5 Then
                    aa = "PACKING"
                End If

                ChartObj.Series.Add(aa)
                'ChartObj.Series.Item(SeriesItem(i)).ChartArea = CArea.Name & i
                'ChartObj.Series.Item(SeriesItem(i)).CustomProperties = BarType
                'ChartObj.Series.Item(SeriesItem(i)).ChartType = DataVisualization.Charting.SeriesChartType.Column
                ''ChartObj.Series.Item(SeriesItem(i)).Font = FontInfo
                'ChartObj.Series.Item(SeriesItem(i)).Points.DataBindXY(Xary, Yary)
                'ChartObj.Series.Item(SeriesItem(i)).XValueType = DataVisualization.Charting.ChartValueType.String

                ChartObj.Series(aa).CustomProperties = BarType
                ChartObj.Series(aa).ChartType = DataVisualization.Charting.SeriesChartType.Column
                ChartObj.Series(aa).Points.DataBindXY(Xary2, Yary2)
                'ChartObj.Series.Item(SeriesItem(i)).XValueType = DataVisualization.Charting.ChartValueType.String


            Next







        Catch ex As Exception
            MessageBox.Show("Error:" & ex.Message, "Error")
        End Try

    End Function

End Module
