Module Md_Part

    Function Part_Loc(ByVal s As FarPoint.Win.Spread.FpSpread, ByVal wh As String) As Boolean
        With s.ActiveSheet

            Dim H_RS As ADODB.Recordset
            Dim I, J As Integer
            Dim wh_cd As String

            .RowCount = 0

            wh_cd = Query_RS("select code_id from tbl_codemaster where site_id = '" & Site_id & "' and class_id = 'R0073' and code_name = '" & wh & "'")

            H_RS = Query_RS_ALL("select DISTINCT substring(code_id,5,1)  from tbl_codemaster where class_id= 'R0074' and code_id like '" & wh_cd & "%' ORDER BY substring(code_id,5,1)")

            If H_RS Is Nothing Then
                .ColumnCount = 1
                .RowCount = 1
                .SetValue(0, 0, wh_cd)
                .ColumnHeader.Cells(0, 0).Text = wh
                .ColumnHeader.Cells(1, 0).Text = ""


                Exit Function

            End If

            .ColumnCount = H_RS.RecordCount
            .ColumnHeader.RowCount = 2

            For I = 0 To .ColumnCount - 1
                .ColumnHeader.Cells(1, I).Text = H_RS(0).Value
                H_RS.MoveNext()
            Next

            .ColumnHeader.Cells(0, 0).ColumnSpan = .ColumnCount
            .ColumnHeader.Cells(0, 0).Text = wh

            H_RS = Nothing

            For I = 0 To .ColumnCount - 1
                H_RS = Query_RS_ALL("select CODE_NAME  from tbl_codemaster where class_id= 'R0074' AND CODE_ID LIKE '" & wh_cd & .ColumnHeader.Cells(1, I).Text & "%' AND ACTIVE = 'Y' ORDER BY CODE_NAME")

                If H_RS Is Nothing Then
                Else
                    If .RowCount < H_RS.RecordCount Then
                        .RowCount = H_RS.RecordCount
                    End If

                    For J = 0 To H_RS.RecordCount - 1
                        .SetValue(J, I, H_RS(0).Value)
                        H_RS.MoveNext()
                    Next
                End If
            Next

            .Cells(0, 0, .RowCount - 1, .ColumnCount - 1).HorizontalAlignment = FarPoint.Win.Spread.CellHorizontalAlignment.Center
            .Cells(0, 0, .RowCount - 1, .ColumnCount - 1).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
            .Cells(0, 0, .RowCount - 1, .ColumnCount - 1).Locked = True
            Spread_AutoCol(s)

            H_RS = Nothing

        End With

    End Function

End Module
