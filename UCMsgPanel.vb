Public Class UCMsgPanel

    Private initX, initY As Integer

    Private Sub UCMsgPanel_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

    End Sub

    Private Sub CtlMove_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles ctlmove.MouseDown
        initX = Me.Location.X - MousePosition.X
        initY = Me.Location.Y - MousePosition.Y
        Timer1.Enabled = True
    End Sub

    Private Sub CtlMove_MouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles ctlmove.MouseUp
        Timer1.Enabled = False
    End Sub

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        Me.Location = New System.Drawing.Point(MousePosition.X + initX, MousePosition.Y + initY)
    End Sub

    Private Sub Ctlclose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ctlclose.Click
        '237,210 /29
        If Me.LabelX1.Visible = True Then
            Me.LabelX1.Visible = False
            Me.Size = New System.Drawing.Size(Me.Width, 23)
            Me.ctlclose.Image = Global.SMSPLUS.My.Resources.Resources.sPlus
        Else
            Me.LabelX1.Visible = True
            Me.Size = New System.Drawing.Size(Me.Width, 100)
            Me.ctlclose.Image = Global.SMSPLUS.My.Resources.Resources.sMinus
        End If
    End Sub


    Private Sub LabelX1_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles LabelX1.DoubleClick
        Me.LabelX1.Text = 0
    End Sub
End Class
