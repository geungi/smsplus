Public Class FrmUSEmpMst

    Private RecCnt As Integer
    Private pic As String
    Dim mycam As New iCam
    Dim mode As String

    Private Sub FrmUSEmpMst_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown, EmpNmTxt.KeyDown, EmpNoTxt.KeyDown

        If e.KeyValue = Keys.Enter Then
            FindBtn_Click(sender, e)
        End If

    End Sub

    Private Sub FrmUSEmpMst_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Me.DockContainerItem1.Text = "인사관리대장"

        If Query_Combo1(ComboBoxItem1, "SELECT CODE_NAME FROM TBL_CODEMASTER WHERE SITE_ID = '" & Site_id & "' AND CLASS_ID = '10009' ORDER BY DIS_ORDER") = True Then
            ComboBoxItem1.Items.Add("ALL")
            ComboBoxItem1.Text = "ALL"
        End If

        'If Query_Combo1(ComboBoxItem3, "SELECT DISTINCT CODE_NAME FROM TBL_CODEMASTER WHERE SITE_ID = '" & Site_id & "' AND CLASS_ID = 'R0034' AND CODE_ID LIKE '3%'") = True Then
        '    ComboBoxItem3.Items.Add("ALL")
        '    ComboBoxItem3.Text = "ALL"
        'End If

        If Query_Combo1(ComboBoxItem2, "SELECT CODE_NAME FROM TBL_CODEMASTER WHERE SITE_ID = '" & Site_id & "' AND CLASS_ID = '10014' ORDER BY CODE_ID") = True Then
            ComboBoxItem2.Items.Add("ALL")
            ComboBoxItem2.Text = "ALL"
        End If

        If Spread_Setting(FpSpread1, Me.Name) = True Then
            Spread_AutoCol(FpSpread1)
            Me.FpSpread1.ActiveSheet.FrozenColumnCount = 3  '틀고정
        End If

        With FpSpread2.ActiveSheet
            .Cells(7, 1).CellType = datecell
            .Cells(7, 3).CellType = datecell
            .Cells(10, 1).CellType = datecell
            .Cells(0, 1).Locked = True
            .ColumnHeader.Visible = False
            .RowHeader.Visible = False
        End With

        Bar1.Visible = False
        RibbonTabItem1.Enabled = False
        RibbonPanel1.Enabled = False

        UC_CAPIMG1.Visible = False

        Formbim_Authority(Me.ButtonItem9, Me.Name, "NEW")
        Formbim_Authority(Me.ButtonItem5, Me.Name, "SAVE")
        Formbim_Authority(Me.ButtonItem8, Me.Name, "DELETE")
        Formbim_Authority(Me.ButtonItem11, Me.Name, "PRINT")
        Formbim_Authority(Me.ButtonItem10, Me.Name, "EXCEL")
        Formbtn_Authority2(Me.FindBtn, Me.Name, "FIND")

    End Sub

    Private Sub FindBtn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FindBtn.Click
        Try

            Dim Query As String
            RibbonTabItem1.Enabled = False
            RibbonPanel1.Enabled = False
            Bar1.Visible = False
            ButtonItem4.Text = "인사관리대장 보기"
            mode = "FIND"

            Query = " select site_ID, emp_no, emp_nm, " & _
                    "        (select code_name from tbl_codemaster where class_id = '10009' and code_id = a.dept_cd), " & _
                    "        (select code_name from tbl_codemaster where class_id = '10009' and code_id = a.dept_cd), " & _
                    "        (select code_name from tbl_codemaster where class_id = 'R0035' and code_id = a.shift_cd), " & _
                    "        social_no, " & _
                    "        (select code_name from tbl_codemaster where class_id = '10012' and code_id = a.pay_type), " & _
                    "        '', 0, " & _
                    "        (select code_name from tbl_codemaster where class_id = '10014' and code_id = a.emp_type), " & _
                    "        ISNULL(entry_dt,''), ISNULL(retire_dt,''), retire_yn, ISNULL(dob_dt,''), " & _
                    "        (select code_name from tbl_codemaster where class_id = '10015' and code_id = a.marital_cd), " & _
                    "        partauth_yn, INSA_yn, ins_type, '', '', cell_no, tel_no, zip_cd, " & _
                    "        country, city, state, address, user_id, password, " & _
                    "        (select emp_nm from tbl_empmaster where emp_no = a.c_person), c_date," & _
                    "        (select emp_nm from tbl_empmaster where emp_no = a.u_person), u_date" & _
                    " From tbl_empmaster a " & _
                    " where site_ID = '" & Site_id & "'  "

            If ComboBoxItem1.Text <> "ALL" Then
                Query = Query & " and DEPT_CD  IN (SELECT CODE_ID FROM TBL_CODEMASTER WHERE site_ID= '" & Site_id & "' AND CLASS_ID = '10009' AND CODE_NAME = '" & ComboBoxItem1.Text & "') "
            End If

            If ComboBoxItem2.Text <> "ALL" Then
                Query = Query & " and EMP_TYPE  IN (SELECT CODE_ID FROM TBL_CODEMASTER WHERE site_ID= '" & Site_id & "' AND CLASS_ID = '10014' AND CODE_NAME = '" & ComboBoxItem2.Text & "') "
            End If

            If ComboBoxItem3.Text <> "ALL" Then
                Query = Query & " and DEPT_CD  IN (SELECT CODE_ID FROM TBL_CODEMASTER WHERE site_ID= '" & Site_id & "' AND CLASS_ID = '10009' AND CODE_NAME = '" & ComboBoxItem3.Text & "') "
            End If


            If Me.EmpNoTxt.Text <> "" Then
                Query += " and emp_no like '%" & Me.EmpNoTxt.Text & "%' "
            End If
            If Me.EmpNmTxt.Text <> "" Then
                Query += " and emp_nm like '%" & Me.EmpNmTxt.Text & "%' "
            End If

            Query += " order by emp_no "

            If Query_Spread(Me.FpSpread1, Query, 1) = True Then
                FpSpread1.ActiveSheet.Columns(0, 1).Locked = True
                Spread_AutoCol(FpSpread1) '스프레드에 데이터를 출력후, 데이터의 사이즈에 맞게 컬럼사이즈 재조정
            End If

        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message, "ERROR")
        End Try

    End Sub

    Private Sub FpSpread2_CellClick(ByVal sender As Object, ByVal e As FarPoint.Win.Spread.CellClickEventArgs) Handles FpSpread2.CellClick
        save_excel = "FpSpread2"
    End Sub
    Private Sub FpSpread2_CellDoubleClick(ByVal sender As Object, ByVal e As FarPoint.Win.Spread.CellClickEventArgs) Handles FpSpread2.CellDoubleClick
        Try

            With FpSpread2.ActiveSheet
                If e.Row = 3 And e.Column = 1 Then
                    Chg_ComboCell(FpSpread2, e.Row, e.Column, Query_Cell_Code1("code_name", "10009"))
                ElseIf e.Row = 4 And e.Column = 1 Then
                    Chg_ComboCell(FpSpread2, e.Row, e.Column, Query_Cell_Code1("code_name", "R0034"))

                ElseIf e.Row = 5 And e.Column = 1 Then
                    Chg_ComboCell(FpSpread2, e.Row, e.Column, New String() {"Y", "N"})
                ElseIf e.Row = 8 And e.Column = 1 Then '고용구분
                    Chg_ComboCell(FpSpread2, e.Row, e.Column, Query_Cell_Code1("code_name", "10014"))
                ElseIf e.Row = 8 And e.Column = 3 Then '외주사
                    Chg_ComboCell(FpSpread2, e.Row, e.Column, Query_Cell_Code2("SELECT DISTINCT CODE_NAME FROM TBL_CODEMASTER WHERE CLASS_ID = 'R0002' AND CODE_ID LIKE '3%'"))
                ElseIf e.Row = 9 And e.Column = 1 Then '급여구분
                    Chg_ComboCell(FpSpread2, e.Row, e.Column, Query_Cell_Code1("code_name", "10012"))
                ElseIf e.Row = 9 And e.Column = 3 Then '근무구분
                    Chg_ComboCell(FpSpread2, e.Row, e.Column, Query_Cell_Code1("code_name", "R0035"))
                ElseIf e.Row = 10 And e.Column = 3 Then '결혼구분
                    Chg_ComboCell(FpSpread2, e.Row, e.Column, Query_Cell_Code1("code_name", "10015"))
                End If
            End With

        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message, "ERROR")
        End Try

    End Sub

    Private Sub FpSpread1_CellClick(ByVal sender As Object, ByVal e As FarPoint.Win.Spread.CellClickEventArgs) Handles FpSpread1.CellClick
        save_excel = "FpSpread1"
    End Sub

    Private Sub FpSpread1_CellDoubleClick(ByVal sender As Object, ByVal e As FarPoint.Win.Spread.CellClickEventArgs) Handles FpSpread1.CellDoubleClick
        Try

        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message, "ERROR")
        End Try
    End Sub

    Private Sub FpSpread1_Change(ByVal sender As Object, ByVal e As FarPoint.Win.Spread.ChangeEventArgs) Handles FpSpread1.Change
        Dim S As FarPoint.Win.Spread.FpSpread = CType(sender, FarPoint.Win.Spread.FpSpread)
        Try
            'S.ActiveSheet.Cells(e.Row, 32).Text = Emp_No
            'S.ActiveSheet.Cells(e.Row, 33).Text = Now
            'Spread_Change(S, e.Row)

        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message, "ERROR")
        End Try

    End Sub

    Private Sub FpSpread2_Change(ByVal sender As Object, ByVal e As FarPoint.Win.Spread.ChangeEventArgs) Handles FpSpread2.Change

        Try

            With FpSpread2.ActiveSheet
                .Cells(0, 1, .RowCount - 1, 1).ForeColor = Color.OrangeRed
                .Cells(7, 3, .RowCount - 1, 3).ForeColor = Color.OrangeRed

                If mode <> "NEW" Then
                    mode = "UPDATE"
                End If

                If mode = "NEW" Then
                    If e.Row = 3 And e.Column = 1 Then
                        .SetValue(0, 1, Query_RS("SELECT ISNULL(CONVERT(VARCHAR(6),CONVERT(INT, MAX(EMP_NO)) + 1),'10001') FROM TBL_EMPMASTER WHERE emp_no LIKE '" & Query_RS("SELECT LEFT(CODE_ID,1) FROM TBL_CODEMASTER WHERE CLASS_ID = 'R0002' AND CODE_NAME = '" & .Cells(3, 1).Text & "'") & "%' AND emp_no <> '11111'"))
                    End If
                End If

            End With

        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message, "ERROR")
        End Try

    End Sub
    Private Sub FpSpread1_LeaveCell(ByVal sender As Object, ByVal e As FarPoint.Win.Spread.LeaveCellEventArgs) Handles FpSpread1.LeaveCell
        Dim S As FarPoint.Win.Spread.FpSpread = CType(sender, FarPoint.Win.Spread.FpSpread)
        Select Case e.Column
            Case 1
                If S.ActiveSheet.Cells(e.Row, e.Column).Text = "" Then
                    MessageBox.Show("ID is not Empty!!")
                End If
            Case 28
                If S.ActiveSheet.Cells(e.Row, e.Column).Text = "" Then
                    MessageBox.Show("PASSWORD is not Empty!!")
                End If
        End Select
    End Sub

    Private Sub NewBtn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonItem9.Click
        Try

            Bar1.Visible = True
            '            Bar1.AutoHide = False

            ButtonItem4.Text = "인사관리대장 닫기"
            Bar1.Visible = True

            RibbonTabItem1.Enabled = True
            RibbonPanel1.Enabled = True


            CLEAR_INSA()
            With FpSpread2.ActiveSheet
                .Cells(0, 1, .RowCount - 1, 1).ForeColor = Color.OrangeRed
                .Cells(7, 3, .RowCount - 1, 3).ForeColor = Color.OrangeRed
                .Cells(5, 1).Text = "N"
                .Cells(7, 1).Text = Now
            End With

            mode = "NEW"

        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message, "ERROR")
        End Try
    End Sub

    Private Sub SaveBtn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonItem5.Click
        Try
           Dim QRY As String

            With Me.FpSpread2.ActiveSheet

                If mode <> "NEW" And mode <> "UPDATE" Then
                    Exit Sub
                End If

                If .Cells(0, 1).ForeColor <> Color.OrangeRed Then
                    Exit Sub
                End If

                If Query_RS("SELECT COUNT(EMP_NO) FROM TBL_EMPMASTER WHERE SITE_id = '" & Site_id & "' AND EMP_NO = '" & .Cells(0, 1).Text & "'") > 0 Then 'UPDATE
                    QRY = "UPDATE TBL_EMPMASTER" & vbNewLine
                    QRY = QRY & "SET EMP_NM = '" & .Cells(1, 1).Text & "'," & vbNewLine
                    QRY = QRY & "   SOCIAL_NO = '" & .Cells(2, 1).Text & "'," & vbNewLine
                    QRY = QRY & "   dept_cd = '" & Query_RS("SELECT CODE_ID FROM TBL_CODEMASTER WHERE SITE = '" & Site_id & "' AND CLASS_ID = '10009' AND CODE_NAME = '" & .Cells(3, 1).Text & "'") & "'," & vbNewLine
                    QRY = QRY & "   ins_TYPE = '" & .Cells(4, 1).Text & "'," & vbNewLine
                    QRY = QRY & "   RETIRE_YN = '" & .Cells(5, 1).Text & "'," & vbNewLine
                    QRY = QRY & "   entry_dt = '" & .Cells(7, 1).Text & "'," & vbNewLine
                    QRY = QRY & "   retire_dt = '" & .Cells(7, 3).Text & "'," & vbNewLine
                    QRY = QRY & "   emp_type = '" & Query_RS("SELECT CODE_ID FROM TBL_CODEMASTER WHERE SITE = '" & Site_id & "' AND CLASS_ID = '10014' AND CODE_NAME = '" & .Cells(8, 1).Text & "'") & "'," & vbNewLine
                    QRY = QRY & "   pay_type = '" & Query_RS("SELECT CODE_ID FROM TBL_CODEMASTER WHERE SITE = '" & Site_id & "' AND CLASS_ID = '10012' AND CODE_NAME = '" & .Cells(9, 1).Text & "'") & "'," & vbNewLine
                    QRY = QRY & "   shift_cd = '" & Query_RS("SELECT CODE_ID FROM TBL_CODEMASTER WHERE SITE = '" & Site_id & "' AND CLASS_ID = 'R0035' AND CODE_NAME = '" & .Cells(9, 3).Text & "'") & "'," & vbNewLine
                    QRY = QRY & "   dob_dt = '" & .Cells(10, 1).Text & "'," & vbNewLine
                    QRY = QRY & "   marital_cd = '" & Query_RS("SELECT CODE_ID FROM TBL_CODEMASTER WHERE SITE = '" & Site_id & "' AND CLASS_ID = '10015' AND CODE_NAME = '" & .Cells(10, 3).Text & "'") & "'," & vbNewLine
                    QRY = QRY & "   cell_no = '" & .Cells(11, 1).Text & "'," & vbNewLine
                    QRY = QRY & "   tel_no = '" & .Cells(11, 3).Text & "'," & vbNewLine
                    QRY = QRY & "   zip_cd = '" & .Cells(12, 1).Text & "'," & vbNewLine
                    QRY = QRY & "   country = '" & .Cells(13, 1).Text & "'," & vbNewLine
                    QRY = QRY & "   state = '" & .Cells(14, 1).Text & "'," & vbNewLine
                    QRY = QRY & "   user_id = '" & .Cells(15, 1).Text & "'," & vbNewLine
                    QRY = QRY & "   password = '" & .Cells(15, 3).Text & "'," & vbNewLine
                    QRY = QRY & "   U_PERSON = '" & Emp_No & "'," & vbNewLine
                    QRY = QRY & "   u_DATE = GETDATE()" & vbNewLine
                    QRY = QRY & "where site_id = '" & Site_id & "'" & vbNewLine
                    QRY = QRY & "  AND EMP_NO = '" & .Cells(0, 1).Text & "'" & vbNewLine
                Else 'INSERT
                    QRY = "" & vbNewLine
                    QRY = QRY & "INSERT INTO TBL_EMPMASTER (SITE_id, EMP_NO, EMP_NM, SOCIAL_NO, dept_cd, INS_TYPE, RETIRE_YN, entry_dt, retire_dt, emp_type, pay_type, shift_cd, dob_dt, marital_cd, cell_no,tel_no,zip_cd,country,state,user_id,password,C_PERSON, C_DATE, U_PERSON, U_DATE) " & vbNewLine
                    QRY = QRY & "VALUES ('S1000'," & vbNewLine
                    QRY = QRY & "'" & .Cells(0, 1).Text & "'," & vbNewLine
                    QRY = QRY & "'" & .Cells(1, 1).Text & "'," & vbNewLine
                    QRY = QRY & "'" & .Cells(2, 1).Text & "'," & vbNewLine
                    QRY = QRY & "'" & Query_RS("SELECT CODE_ID FROM TBL_CODEMASTER WHERE SITE_id = '" & Site_id & "' AND CLASS_ID = '10009' AND CODE_NAME = '" & .Cells(3, 1).Text & "'") & "'," & vbNewLine
                    QRY = QRY & "'" & .Cells(4, 1).Text & "'," & vbNewLine
                    QRY = QRY & "'" & .Cells(5, 1).Text & "'," & vbNewLine
                    QRY = QRY & "'" & .Cells(7, 1).Text & "'," & vbNewLine
                    QRY = QRY & "'" & .Cells(7, 3).Text & "'," & vbNewLine
                    QRY = QRY & "'" & Query_RS("SELECT CODE_ID FROM TBL_CODEMASTER WHERE SITE_id = '" & Site_id & "' AND CLASS_ID = '10014' AND CODE_NAME = '" & .Cells(8, 1).Text & "'") & "'," & vbNewLine
                    QRY = QRY & "'" & Query_RS("SELECT CODE_ID FROM TBL_CODEMASTER WHERE SITE_id = '" & Site_id & "' AND CLASS_ID = '10012' AND CODE_NAME = '" & .Cells(9, 1).Text & "'") & "'," & vbNewLine
                    QRY = QRY & "'" & Query_RS("SELECT CODE_ID FROM TBL_CODEMASTER WHERE SITE_id = '" & Site_id & "' AND CLASS_ID = 'R0035' AND CODE_NAME = '" & .Cells(9, 3).Text & "'") & "'," & vbNewLine
                    QRY = QRY & "'" & .Cells(10, 1).Text & "'," & vbNewLine
                    QRY = QRY & "'" & Query_RS("SELECT CODE_ID FROM TBL_CODEMASTER WHERE SITE_id = '" & Site_id & "' AND CLASS_ID = '10015' AND CODE_NAME = '" & .Cells(10, 3).Text & "'") & "'," & vbNewLine
                    QRY = QRY & "'" & .Cells(11, 1).Text & "'," & vbNewLine
                    QRY = QRY & "'" & .Cells(11, 3).Text & "'," & vbNewLine
                    QRY = QRY & "'" & .Cells(12, 1).Text & "'," & vbNewLine
                    QRY = QRY & "'" & .Cells(13, 1).Text & "'," & vbNewLine
                    QRY = QRY & "'" & .Cells(14, 1).Text & "'," & vbNewLine
                    QRY = QRY & "'" & .Cells(15, 1).Text & "'," & vbNewLine
                    QRY = QRY & "'" & .Cells(15, 3).Text & "'," & vbNewLine
                    QRY = QRY & "'" & Emp_No & "',GETDATE(),'" & Emp_No & "',GETDATE())" & vbNewLine
                End If

                .Cells(0, 1, .RowCount - 1, 1).ForeColor = Color.Black
                .Cells(7, 3, .RowCount - 1, 3).ForeColor = Color.Black

            End With

            If Insert_Data(QRY) = True Then
                MessageBox.Show("저장이 완료되었습니다.", "Message")
                mode = "FIND"
            Else
                MessageBox.Show("정상적으로 저장되지 않았습니다." & vbNewLine & "ADMIN에게 확인하시기 바랍니다.", "Message")
            End If

        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message, "ERROR")
        End Try

    End Sub

    Private Sub DelBtn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonItem8.Click
        Try
            With Me.FpSpread1.ActiveSheet
                If .RowCount > 0 Then
                    Dim RowCnt = .RowCount
                    Dim r As DialogResult = MessageBox.Show("선택된 사용자를 삭제하시겠습니까?", "Selected Rows Delete", MessageBoxButtons.YesNo)
                    If r = Windows.Forms.DialogResult.Yes Then
                        Insert_Data("delete FROM tbl_empmaster where site = '" & .Cells(.ActiveRowIndex, 0).Text & "' and emp_no = '" & .Cells(.ActiveRowIndex, 1).Text & "'")
                        .RemoveRows(.ActiveRowIndex, 1)
                        RecCnt = RecCnt - 1
                    End If
                    MessageBox.Show("삭제가 완료되었습니다.", "Message")
                End If
            End With

        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message, "ERROR")
        End Try
    End Sub

    Private Sub PrtBtn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonItem11.Click

        If save_excel = "FpSpread1" Then
            If Spread_Print(Me.FpSpread1, "Employee Master", 1) = False Then
                MsgBox("Fail to Print")
            End If
        Else
            If Spread_Print(Me.FpSpread2, "Employee Master", 1) = False Then
                MsgBox("Fail to Print")
            End If

        End If

    End Sub

    Private Sub Excel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonItem10.Click
        If save_excel = "FpSpread1" Then
            File_Save(SaveFileDialog1, FpSpread1)
        Else
            File_Save(SaveFileDialog1, FpSpread2)

        End If
    End Sub

    Function Get_Img(ByVal path As String) As Byte()

        Dim stream As System.IO.FileStream = New System.IO.FileStream(path, IO.FileMode.Open, IO.FileAccess.Read)
        Dim reader As System.IO.BinaryReader = New System.IO.BinaryReader(stream)

        Dim img As Byte() = reader.ReadBytes(CInt(stream.Length))

        stream.Close()
        reader.Close()

        Get_Img = img

    End Function

    Function Insert_Img_Data(ByVal path As String) As Boolean
        '      Try

        Dim conn As New SqlClient.SqlConnection
        Dim cmd As New SqlClient.SqlCommand
        Dim aa As Byte()

        conn.ConnectionString = db_conn(5)
        cmd.Connection = conn

        cmd.CommandType = System.Data.CommandType.StoredProcedure
        cmd.CommandText = "SP_FrmUSEmpMst_UploadImg"
        cmd.CommandTimeout = 90

        aa = Get_Img(path)

        Dim param As SqlClient.SqlParameter = New SqlClient.SqlParameter("SITE", "D1000")
        cmd.Parameters.Add(param)

        param = New SqlClient.SqlParameter("img", System.Data.SqlDbType.Image, aa.Length)
        param.Value = aa
        cmd.Parameters.Add(param)

        If FpSpread2.ActiveSheet.GetValue(0, 1) = "" Then
            MessageBox.Show("사번을 선택하십시오.")
            Exit Function
        End If

        param = New SqlClient.SqlParameter("nm", FpSpread2.ActiveSheet.GetValue(0, 1))
        cmd.Parameters.Add(param)

        param = New SqlClient.SqlParameter("PERSON", Emp_No)
        cmd.Parameters.Add(param)

        conn.Open()
        cmd.ExecuteNonQuery()
        conn.Close()

        MessageBox.Show("저장이 완료되었습니다.")

    End Function

    Private Sub ButtonItem14_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonItem14.Click
        Dim path As String

        With OpenFileDialog1
            .InitialDirectory = My.Computer.FileSystem.SpecialDirectories.MyPictures
            .Filter = "jpg (*.jpg)|*.jpg*|All Files(*.*)|*.*"
            If .ShowDialog() = Windows.Forms.DialogResult.OK Then
                path = .FileName
                pic = .FileName

                With FpSpread2.ActiveSheet
                    Dim img As Drawing.Image
                    img = Image.FromFile(path)
                    .Cells(0, 3).Value = img
                End With
            End If
        End With
    End Sub

    Private Sub ButtonItem2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonItem2.Click
        If pic <> "" Then
            Insert_Img_Data(pic)
        End If
    End Sub
    Function QRY_INSA() As Boolean
        With FpSpread2.ActiveSheet
            .Cells(0, 1).Value = FpSpread1.ActiveSheet.GetValue(FpSpread1.ActiveSheet.ActiveRowIndex, 1)
            .Cells(1, 1).Value = FpSpread1.ActiveSheet.GetValue(FpSpread1.ActiveSheet.ActiveRowIndex, 2)
            .Cells(2, 1).Value = FpSpread1.ActiveSheet.GetValue(FpSpread1.ActiveSheet.ActiveRowIndex, 6)
            .Cells(3, 1).Value = FpSpread1.ActiveSheet.GetValue(FpSpread1.ActiveSheet.ActiveRowIndex, 3)
            .Cells(4, 1).Value = FpSpread1.ActiveSheet.GetValue(FpSpread1.ActiveSheet.ActiveRowIndex, 18)
            .Cells(5, 1).Value = FpSpread1.ActiveSheet.GetValue(FpSpread1.ActiveSheet.ActiveRowIndex, 13)

            'Dim ss As String

            'ss = .Cells(7, 1).Value
            .Cells(7, 1).Value = FpSpread1.ActiveSheet.GetValue(FpSpread1.ActiveSheet.ActiveRowIndex, 11)
            .Cells(7, 3).Value = FpSpread1.ActiveSheet.GetValue(FpSpread1.ActiveSheet.ActiveRowIndex, 12)


            .Cells(8, 1).Value = FpSpread1.ActiveSheet.GetValue(FpSpread1.ActiveSheet.ActiveRowIndex, 10)
            If .Cells(8, 1).Value = "도급직" Then
                .Cells(8, 3).Value = FpSpread1.ActiveSheet.GetValue(FpSpread1.ActiveSheet.ActiveRowIndex, 3)
            Else
                .Cells(8, 3).Value = ""
            End If

            .Cells(9, 1).Value = FpSpread1.ActiveSheet.GetValue(FpSpread1.ActiveSheet.ActiveRowIndex, 7)
            .Cells(9, 3).Value = FpSpread1.ActiveSheet.GetValue(FpSpread1.ActiveSheet.ActiveRowIndex, 5)

            .Cells(10, 1).Value = FpSpread1.ActiveSheet.GetValue(FpSpread1.ActiveSheet.ActiveRowIndex, 14)
            .Cells(10, 3).Value = FpSpread1.ActiveSheet.GetValue(FpSpread1.ActiveSheet.ActiveRowIndex, 15)
            .Cells(11, 1).Value = FpSpread1.ActiveSheet.GetValue(FpSpread1.ActiveSheet.ActiveRowIndex, 21)
            .Cells(11, 3).Value = FpSpread1.ActiveSheet.GetValue(FpSpread1.ActiveSheet.ActiveRowIndex, 22)
            .Cells(12, 1).Value = FpSpread1.ActiveSheet.GetValue(FpSpread1.ActiveSheet.ActiveRowIndex, 23)
            .Cells(13, 1).Value = FpSpread1.ActiveSheet.GetValue(FpSpread1.ActiveSheet.ActiveRowIndex, 24)
            .Cells(14, 1).Value = FpSpread1.ActiveSheet.GetValue(FpSpread1.ActiveSheet.ActiveRowIndex, 26)

            .Cells(15, 1).Value = FpSpread1.ActiveSheet.GetValue(FpSpread1.ActiveSheet.ActiveRowIndex, 28)
            .Cells(15, 3).Value = FpSpread1.ActiveSheet.GetValue(FpSpread1.ActiveSheet.ActiveRowIndex, 29)

            Dim aa As Byte()

            aa = Query_RS_Img("select isnull(emp_IMG,'') from TBL_EMPMASTER where EMP_NO = '" & .GetValue(0, 1) & "'")

            If aa Is Nothing Then
                .Cells(0, 3).Value = ""
            Else
                If aa.Length = 0 Then
                    .Cells(0, 3).Value = ""
                Else
                    Dim ms As System.IO.MemoryStream = New System.IO.MemoryStream(aa, 0, aa.Length)
                    Dim img As Drawing.Image

                    img = Image.FromStream(ms)
                    imgcell.Style = FarPoint.Win.RenderStyle.Stretch
                    .Cells(0, 3).CellType = imgcell
                    .Cells(0, 3).Value = img
                End If
            End If
            '    aa.Dispose()
        End With

    End Function
    Function CLEAR_INSA() As Boolean
        With FpSpread2.ActiveSheet
            .Cells(3, 1).CellType = textcell
            .Cells(5, 1).CellType = textcell
            .Cells(8, 1).CellType = textcell
            .Cells(8, 3).CellType = textcell
            .Cells(9, 1).CellType = textcell
            .Cells(9, 3).CellType = textcell
            .Cells(10, 3).CellType = textcell


            .Cells(0, 1, .RowCount - 1, 1).Value = ""
            .Cells(0, 3, .RowCount - 1, 3).Value = ""
            .Cells(0, 1, .RowCount - 1, 1).ForeColor = Color.Black
            .Cells(7, 3, .RowCount - 1, 3).ForeColor = Color.Black
        End With
    End Function
    Private Sub ButtonItem4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonItem4.Click

        With FpSpread1.ActiveSheet
            If .RowCount = 0 Then
                Exit Sub
            End If
        End With

        If ButtonItem4.Text = "인사관리대장 보기" Then
            ButtonItem4.Text = "인사관리대장 닫기"
            Bar1.Visible = True

            RibbonTabItem1.Enabled = True
            RibbonPanel1.Enabled = True

            QRY_INSA()
        Else
            ButtonItem4.Text = "인사관리대장 보기"
            Bar1.Visible = False
            RibbonTabItem1.Enabled = False
            RibbonPanel1.Enabled = False

            CLEAR_INSA()
        End If

    End Sub

    Private Sub FpSpread1_SelectionChanged(ByVal sender As Object, ByVal e As FarPoint.Win.Spread.SelectionChangedEventArgs) Handles FpSpread1.SelectionChanged

        mode = "FIND"
        CLEAR_INSA()

        If ButtonItem4.Text = "인사관리대장 닫기" Then
            QRY_INSA()
        End If
    End Sub

    Private Sub ButtonItem6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonItem6.Click

        If Bar1.Visible = False Then
            MessageBox.Show("인사관리대장 보기에서만 가능합니다!")
            Exit Sub
        End If

        mycam.initCam(UC_CAPIMG1.PictureBox1.Handle.ToInt32)
        UC_CAPIMG1.Width = 360 '180
        UC_CAPIMG1.Height = 240
        UC_CAPIMG1.PictureBox1.Width = 360 '180
        UC_CAPIMG1.PictureBox1.Height = 240
        UC_CAPIMG1.Visible = True
    End Sub

    Private Sub ButtonItem15_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonItem15.Click
        mycam.closeCam()
        UC_CAPIMG1.Visible = False
    End Sub

    Private Sub ButtonItem16_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonItem16.Click
        Try
            UC_CAPIMG1.PictureBox1.Image = mycam.copyFrame(UC_CAPIMG1.PictureBox1, New RectangleF(0, 0, UC_CAPIMG1.PictureBox1.Width, UC_CAPIMG1.PictureBox1.Height))
            With FpSpread2.ActiveSheet

                If My.Computer.FileSystem.DirectoryExists("c:\temp") = False Then
                    My.Computer.FileSystem.CreateDirectory("c:\temp")
                End If

                Dim F_NAME = "c:\temp\" & .Cells(0, 1).Value & Query_RS("SELECT CONVERT(VARCHAR(50),GETDATE(),112) + REPLACE(CONVERT(VARCHAR(50),GETDATE(),108),':','_')") & ".jpg"

                UC_CAPIMG1.PictureBox1.Image.Save(F_NAME)

                Dim img As Drawing.Image

                img = Image.FromFile(F_NAME)

                imgcell.Style = FarPoint.Win.RenderStyle.Stretch
                .Cells(0, 3).CellType = imgcell

                .Cells(0, 3).Value = img

                pic = F_NAME
            End With

        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try

    End Sub

    Private Sub ButtonItem13_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Me.Close()
    End Sub

    Private Sub ButtonItem7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Me.Close()
    End Sub

    Private Sub ComboBoxItem2_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBoxItem2.SelectedIndexChanged
        If ComboBoxItem2.Text = "도급직" Then
            ComboBoxItem3.Enabled = True
        Else
            ComboBoxItem3.Text = "ALL"
            ComboBoxItem3.Enabled = False
        End If
    End Sub

End Class