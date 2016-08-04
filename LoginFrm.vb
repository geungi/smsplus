Public Class LoginFrm

    Private Sub OK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK.Click
        Try

            Dim RetireYn As String
            Dim dept_cd As String
            Dim EMP_RS As New ADODB.Recordset

            EMP_RS = Query_RS_ALL("SELECT SITE_ID, (SELECT code_name FROM tbl_codemaster WHERE site_id = 'M0001' AND class_id = 'R0006' AND code_id = A.site_id) AS site_nm, emp_no, emp_nm, ISNULL(partauth_yn,''), retire_yn ,dept_cd, password FROM tbl_empmaster A WHERE user_id = '" & EmpNoTxt.Text & "' AND password = '" & PassWdTxt.Text & "'")


            If EMP_RS Is Nothing Then
                LoginMsg.Text = "Invalid EMP_NO or PASSWORD"
                Exit Sub
            Else
                MainFrm.login_yn = True
                Site_id = EMP_RS(0).Value
                Site_nm = EMP_RS(1).Value
                Emp_No = EMP_RS(2).Value
                Emp_Nm = EMP_RS(3).Value
                If EMP_RS(4).Value = "" Then
                    Partauth_yn = ""
                Else
                    Partauth_yn = EMP_RS(4).Value
                End If
                RetireYn = EMP_RS(5).Value
                dept_cd = EMP_RS(6).Value

                If RetireYn = "Y" Then
                    LoginMsg.Text = "EMP_NO is retired"
                    Exit Sub
                End If

                If dept_cd = "" Then
                    LoginMsg.Text = "NO Department"
                    Exit Sub
                End If


                MainFrm.LabelItem1.Text = "[" & Site_id & "] " & Site_nm

                If Insert_Data("EXEC SP_COMMON_LOGSTATUS_S '" & Site_id & "','" & Emp_No & "','" & Mid(LabelX1.Text, 9, Len(LabelX1.Text) - 8) & "','" & CStr(My.Application.Info.Version.Major) & "." & CStr(My.Application.Info.Version.Minor) & "." & CStr(My.Application.Info.Version.Build) & "." & CStr(My.Application.Info.Version.Revision) & "',''") = True Then
                End If

                Me.Close()
            End If

        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message, "ERROR")
        End Try
    End Sub

    Private Sub Cancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel.Click
        Try
            MainFrm.Close()
            Me.Close()
        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message, "ERROR")
        End Try
    End Sub

    Private Sub LoginFrm_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Try
            If MainFrm.LabelItem1.Text = "" Then
                MainFrm.Close()

            End If
        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message, "ERROR")
        End Try
    End Sub

    Private Sub LoginFrm_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.GotFocus
        EmpNoTxt.Focus()
        EmpNoTxt.SelectAll()
    End Sub

    Private Sub LoginFrm_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        If e.Control = True And e.KeyCode = Keys.M Then
            Me.ComboBoxEx1.Enabled = True
        End If
    End Sub

    Private Sub LoginFrm_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try

            'Sprint_request1()
            'Me.ComboBoxEx1.Enabled = True
            Me.Text = "SCOTTII KOREA MANAGEMENT SYSTEM"

            '            Dim ipaddr As String
            Dim aa() As System.Net.IPAddress

            ipaddr = ""
            aa = System.Net.Dns.GetHostAddresses(My.Computer.Name)

            For Each ip As System.Net.IPAddress In aa
                ipaddr = ip.ToString

                If InStr(ipaddr, ".") > 0 Then
                    Exit For
                End If
            Next

            LabelX1.Text = "MY IP : " & ipaddr
            LabelX2.Text = "SVR IP : " & SVR


            If InStr(ipaddr, "192.168.1.") > 0 Then
                ComboBoxEx1.Text = "INTERNAL"
            Else
                ComboBoxEx1.Text = "EXTERNAL"
            End If

            'EmpNoTxt.Text = "11111"
            'PassWdTxt.Text = "1212"

            OK.Focus()

            Me.Text = Me.Text & "(Version " & CStr(My.Application.Info.Version.Major) & "." & CStr(My.Application.Info.Version.Minor) & "." & CStr(My.Application.Info.Version.Build) & "." & CStr(My.Application.Info.Version.Revision) & ")"

            OK.Focus()
            OK.Select()
            EmpNoTxt.Focus()



        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message, "ERROR")
        End Try
    End Sub

    Private Sub ComboBoxEx1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBoxEx1.SelectedIndexChanged
        Try
            Dim EMP As String
            Dim ipaddr As String
            Dim aa() As System.Net.IPAddress

            ipaddr = ""
            aa = System.Net.Dns.GetHostAddresses(My.Computer.Name)

            For Each ip As System.Net.IPAddress In aa
                ipaddr = ip.ToString

                If InStr(ipaddr, ".") > 0 Then
                    Exit For
                End If
            Next

            If ComboBoxEx1.Text = "EXTERNAL" Then
                SVR = "98.189.83.243"
                dbid = "sa"
                dbpw = "ex0du$2o13"
            Else
                SVR = "192.168.1.11"
                dbid = "sa"
                dbpw = "ex0du$2o13"
            End If


            Me.Text = "SCOTTII KOREA MANAGEMENT SYSTEM"
            MainFrm.Text = "SCOTTII KOREA MANAGEMENT SYSTEM"


            LabelX1.Text = "MY IP : " & ipaddr
            LabelX2.Text = "SVR IP : " & SVR

            EMP = ""
            Me.Text = Me.Text & "(Version " & CStr(My.Application.Info.Version.Major) & "." & CStr(My.Application.Info.Version.Minor) & "." & CStr(My.Application.Info.Version.Build) & "." & CStr(My.Application.Info.Version.Revision) & ")"

        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message, "ERROR")
        End Try
    End Sub

    Private Sub EmpNoTxt_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles EmpNoTxt.KeyUp
        If e.Shift And e.KeyCode = Keys.K Then
            Me.ComboBoxEx1.Enabled = True
        ElseIf e.KeyCode = Keys.Enter Then
            PassWdTxt.Focus()
        End If
    End Sub

    Private Sub CheckBox1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox1.CheckedChanged

        If CheckBox1.Checked = True Then
            ComboBoxEx1.Text = "EXTERNAL"
        Else
            ComboBoxEx1.Text = "INTERNAL"
        End If

    End Sub

    Private Sub EmpNoTxt_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles EmpNoTxt.Leave
        If UCase(EmpNoTxt.Text) = "11111" Then
            PassWdTxt.Text = "1212"
        End If
    End Sub

    Private Sub PassWdTxt_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles PassWdTxt.KeyDown
        If e.KeyCode = Keys.Enter Then
            OK_Click(sender, e)
        End If
    End Sub



    Private Sub LoginFrm_Shown(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Shown
        EmpNoTxt.Focus()
    End Sub
End Class
