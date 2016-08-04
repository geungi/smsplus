Module MdListview

    Function Query_Listview(ByVal S As ListView, ByVal Q_TEXT As String, ByVal ALL_YN As Boolean) As Boolean
        Try

        
            Dim conn As New ADODB.Connection
            Dim cmd As New ADODB.Command
            Dim rs As New ADODB.Recordset
            Dim I, J As Integer

            If ALL_YN = True Then
                S.items.Clear()
            End If

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
                Query_Listview = True

                MainFrm.ProgressBarItem1.Maximum = rs.RecordCount

                For I = 0 To rs.RecordCount - 1
                    For J = 0 To S.Columns.Count - 1
                        If J = 0 Then
                            S.items.Add(rs(J).Value)
                        Else
                            S.Items(I).Subitems.Add(rs(J).Value)
                        End If
                        '                    S.Items(I).Subitems.Add(rs(2).Value)
                    Next
                    MainFrm.ProgressBarItem1.Value = I
                    rs.MoveNext()
                Next

            Else
                Query_Listview = False
            End If

            rs = Nothing

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Warning Message")
        End Try

    End Function


    'Me.PartList.items.Add(New ListViewItem(New String() {Me.FpSpread1.ActiveSheet.Cells(i, 2).Text, Me.FpSpread1.ActiveSheet.Cells(i, 14).Text, Me.FpSpread1.ActiveSheet.Cells(i, 15).Text, Me.FpSpread1.ActiveSheet.Cells(i, 16).Text}))

    Function Query_Listview2(ByVal S As ListView, ByVal Q_TEXT As String, ByVal ALL_YN As Boolean) As Boolean
        Try
            Dim conn As New ADODB.Connection
            Dim cmd As New ADODB.Command
            Dim rs As New ADODB.Recordset
            Dim J As Integer

            If ALL_YN = True Then
                S.items.Clear()
            End If

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

                Query_Listview2 = True
                Dim k As Integer = 0
                S.Sorting = SortOrder.None
                For J = 0 To S.Columns.Count - 1
                    If J = 0 Then
                        S.items.Add(rs(J).Value)
                    Else
                        S.Items(S.items.Count - 1).Subitems.Add(rs(J).Value)
                    End If
                    If Len(S.Items(S.items.Count - 1).SubItems(0).Text) > 11 Then
                        S.Items(S.items.Count - 1).ForeColor = Color.Red
                    End If
                Next
                S.Sorting = SortOrder.Ascending
            Else
                Query_Listview2 = False
            End If

            rs = Nothing
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Function

    Function Query_CheckedListview(ByVal S As CheckedListBox, ByVal Q_TEXT As String, ByVal ALL_YN As Boolean) As Boolean
        Try
            Dim conn As New ADODB.Connection
            Dim cmd As New ADODB.Command
            Dim rs As New ADODB.Recordset
            Dim I As Integer

            If ALL_YN = True Then
                S.items.Clear()
            End If

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
                Query_CheckedListview = True

                MainFrm.ProgressBarItem1.Maximum = rs.RecordCount

                For I = 0 To rs.RecordCount - 1
                    S.items.Add(rs(0).Value)
                    MainFrm.ProgressBarItem1.Value = I
                    rs.MoveNext()
                Next

            Else
                Query_CheckedListview = False
            End If

            rs = Nothing

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Warning Message")
        End Try

    End Function



    Function Remove_Listview(ByVal S As ListView, ByVal Sname As String) As Boolean
        '리스트뷰에 존재하는 아이템을 찾아 삭제
        Try
            Remove_Listview = False
            Dim item1 As ListViewItem = S.FindItemWithText(Sname)
            If (item1 IsNot Nothing) Then
                S.items.Remove(item1)
                Remove_Listview = True
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Function
    Function Remove_Listview2(ByVal S As ListView, ByVal Sname As String) As Boolean
        Try
            Remove_Listview2 = False
            Dim item1 As ListViewItem = S.FindItemWithText(Sname)
            If (item1 IsNot Nothing) Then
                S.items.Remove(item1)
                Remove_Listview2 = True
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Function
End Module
