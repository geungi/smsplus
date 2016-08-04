Class FrmModifyPO
    Const CtrlMask As Byte = 8
    Private NowPoNo As String = ""
    Private SelPoNo As String = ""
    Private TotRec As Integer = 0
    Private TotRec2 As Integer = 0

    Private Sub FrmModifyPO_Disposed(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Disposed

    End Sub


    Private Sub FrmModifyPO_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            If Me.LabelX4.Text = "" Then
                Me.Close()
            Else
                Me.DockContainerItem2.Text = Me.LabelX4.Text
            End If

            Me.ComboBoxEx3.Items.Clear()
            Me.ComboBoxEx2.Items.Clear()
            Me.ComboBoxEx1.Items.Clear()
            Query_Combo(Me.ComboBoxEx3, "select model_no from tbl_modelmaster where site_id = '" & Site_id & "' and active='Y'")

            Me.ComboBoxEx1.Items.Clear()
            If Me.LabelX1.Text = "" Then
                Me.ComboBoxEx1.Text = "INSERT"
                Me.ComboBoxEx3.Enabled = True
            Else
                Me.ComboBoxEx1.Text = "MODIFY"
                Me.ComboBoxEx3.Enabled = False
                Me.ComboBoxEx2.Items.Add(Me.LabelX1.Text)
                Me.ComboBoxEx2.Text = Me.LabelX1.Text
            End If

            Me.ComboBoxEx1.Items.AddRange(New String() {"INSERT", "MODIFY"})
            Me.GroupPanel2.Visible = False
        Catch ex As Exception
            MessageBox.Show("Err: " & ex.Message, "Error")
        End Try
    End Sub

    Private Sub ButtonX1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonX1.Click
        Try
            Dim PoRs As ADODB.Recordset
            PoRs = Query_RS_ALL("EXEC SP_FRMMODIFYPO '" & Site_id & "','" & Me.ComboBoxEx2.Text & "','" & Me.DockContainerItem2.Text & "'")

            If PoRs Is Nothing Then
                MessageBox.Show("PartNo is not exist!!!", "Validation Error")
                Me.TextBoxX1.Focus()
                Exit Sub
            End If

            Me.partno.Text = PoRs(0).Value
            Me.partname.Text = PoRs(1).Value
            Me.partqty.Text = CInt(PoRs(2).Value)
            Me.TecNm.Text = PoRs(3).Value & " : "
            Me.TecQty.Text = CInt(PoRs(4).Value)
            Me.OrderQty.Text = CInt(PoRs(5).Value)
            Me.RcvQty.Text = CInt(PoRs(6).Value)
            Me.LabelX5.Text = PoRs(7).Value
            Me.LabelX9.Text = PoRs(8).Value

            PoRs = Nothing

            If CInt(Me.OrderQty.Text) > 0 Then
                Me.ComboBoxEx1.Text = "MODIFY"
            End If

            If Me.ComboBoxEx1.Text = "INSERT" Then
                Me.ButtonX2.Text = "INSERT"
                Me.r_qty.ReadOnly = True
            Else
                Me.ButtonX2.Text = "MODIFY"
            End If
            Me.GroupPanel2.Visible = True
        Catch ex As Exception
            MessageBox.Show("Err: " & ex.Message, "Error")
        End Try
    End Sub


    Private Sub ButtonX2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonX2.Click
        Try

            If o_qty.Text = "" Then
                o_qty.Text = 0
            End If

            If r_qty.Text = "" Then
                r_qty.Text = 0
            End If

            Select Case Me.ButtonX2.Text
                Case "INSERT"
                    If CInt(o_qty.Text) < 1 Then
                        Exit Sub
                    End If
                Case "MODIFY"
                    If CInt(o_qty.Text) < 1 And CInt(r_qty.Text) < 1 Then
                        Exit Sub
                    End If
                    If CInt(r_qty.Text) > CInt(partqty.Text) Then
                        MessageBox.Show("RCV Qty > PARTROOM'S Qty", "Validation Error")
                        Exit Sub
                    End If
            End Select

            Insert_Data("EXEC SP_FRMMODIFYPO_INS '" & Site_id & "','" & Me.LabelX4.Text & "','" & Me.partno.Text & "','" & Me.o_qty.Text & "','" & Me.r_qty.Text & "','" & Emp_No & "','" & Me.LabelX5.Text & "','" & Me.ButtonX2.Text & "'")


           


            MessageBox.Show("Find's Button Must be Click !!", "Success Modify")
            Me.Dispose()

        Catch ex As Exception
            MessageBox.Show("Err: " & ex.Message, "Error")
        End Try
    End Sub


    
    Private Sub o_qty_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles o_qty.KeyPress, r_qty.KeyPress
        If Asc(e.KeyChar) < 48 Or Asc(e.KeyChar) > 57 Then
            e.KeyChar = ""
        End If
    End Sub

    Private Sub ComboBoxEx3_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBoxEx3.SelectedIndexChanged
        If ComboBoxEx3.Enabled = True Then
            Query_Combo(Me.ComboBoxEx2, "select b.lg_partno from tbl_bom a,tbl_partmaster b where a.site_id = '" & Site_id & "' and a.site_id = b.site_id and a.c_no = b.part_no and a.p_no = '" & ComboBoxEx3.Text & "' and a.active='Y' and a.c_no not like '%[_]R'   order by b.lg_partno")
        End If
    End Sub



    Private Sub ComboBoxEx1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBoxEx1.SelectedIndexChanged
        If ComboBoxEx1.Text = "INSERT" Then
            Me.ComboBoxEx2.Items.Clear()
            Me.ComboBoxEx3.Enabled = True
        Else
            Me.ComboBoxEx1.Text = "MODIFY"
            Me.ComboBoxEx2.Items.Clear()
            Me.ComboBoxEx3.Enabled = False
            Me.ComboBoxEx2.Items.Add(Me.LabelX1.Text)
            Me.ComboBoxEx2.Text = Me.LabelX1.Text
        End If
    End Sub
End Class