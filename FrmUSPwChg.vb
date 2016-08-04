Public Class FrmUSPwChg

    Private Sub ButtonX2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonX2.Click
        Me.Close()
    End Sub

    Private Sub ButtonX1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonX1.Click

        If TextBoxX1.Text = "" Then
            Modal_Error("기존 패스워드를 입력하세요.")
            Exit Sub
        End If

        If Query_RS("SELECT PASSWORD FROM TBL_EMPMASTER WHERE SITE = '" & Site_id & "' AND EMP_NO = '" & Emp_No & "'") <> TextBoxX1.Text Then
            Modal_Error("기존 패스워드가 틀립니다.")
            Exit Sub
        End If

        If TextBoxX2.Text = "" Or TextBoxX3.Text = "" Then
            Modal_Error("변경할 패스워드를 입력하세요.")
            Exit Sub
        End If

        If TextBoxX2.Text <> TextBoxX3.Text Then
            Modal_Error("변경할 패스워드가 틀립니다.")
            Exit Sub
        End If

        If Insert_Data("UPDATE TBL_EMPMASTER SET PASSWORD = '" & TextBoxX3.Text & "' WHERE SITE = '" & Site_id & "' AND EMP_NO = '" & Emp_No & "'") = True Then
            MessageBox.Show("패스워드 변경이 완료되었습니다.")
            ButtonX2_Click(sender, e)
        End If

    End Sub

    Private Sub FrmUSPwChg_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        TextBoxX1.Text = ""
        TextBoxX2.Text = ""
        TextBoxX3.Text = ""
    End Sub
End Class