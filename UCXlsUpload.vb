Public Class UCXlsUpload

    Private Sub ButtonX1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonX1.Click
        Me.Cursor = Cursors.WaitCursor
        Dim i As Integer
        If Xls_SpIns(Me.FileOpen, UcXup_SP, TextBoxX1, UcXup_FrmNm) = True Then


            If UcXup_FrmNm = "FrmBOM" Then
                With UcXup_SP.ActiveSheet
                    For i = 0 To .RowCount - 1
                        Dim item1 As ListViewItem = UcXup_ListVw.FindItemWithText(.Cells(i, 2).Text)
                        If (item1 IsNot Nothing) Then
                            .Cells(i, 15).Text = item1.SubItems(1).Text
                            .Cells(i, 16).Text = item1.SubItems(2).Text
                            .Cells(i, 17).Text = item1.SubItems(3).Text
                            UcXup_ListVw.items.Remove(item1)
                        End If
                    Next
                End With
            End If

            TextBoxX1.Text = ""
            Me.Visible = False

        End If

        Me.Cursor = Cursors.Default
    End Sub

    Private Sub GroupPanel1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles GroupPanel1.Click
        If Me.Visible = True Then Me.Visible = False

    End Sub
End Class
