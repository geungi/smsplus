Imports System.Data.SqlClient
Imports System.Net.Mail
Imports System.Web

Module MdData

    Public SVR As String
    Public dbid As String
    Public dbpw As String

    Public Site_id As String
    Public Site_nm As String
    Public Emp_No As String
    Public Emp_Nm As String

    Public OP_No As String
    Public OP_Nm As String

    Public WH_NO As String
    Public WH_NM As String

    Public UcXup_SP As FarPoint.Win.Spread.FpSpread
    Public UcXup_FrmNm As String
    Public UcXup_ListVw As ListView
    Public UcXup_SP2 As FarPoint.Win.Spread.FpSpread
    Public Partauth_yn As String

    Public File_Status As String

    Public ipaddr As String
    Public i_conn As New ADODB.Connection
    Public i_cmd As New ADODB.Command
    Public img_fname As String

    Function db_conn(ByVal flag) As String

        'SVR = "98.189.83.243"
        'dbpw = "ex0du$2o13"
        dbid = "exokor"

        SVR = "211.239.157.71\exokor , 63367" '"192.168.219.105"
        dbpw = "ex0du$2o13"

        'SVR = "localhost"
        'dbpw = "udiatech"
        'dbid = "sa"


        db_conn = ""
        If flag = 2 Then
            db_conn = "Provider=sqloledb;Data Source=98.189.83.243,1433;Network Library=DBMSSOCN;Initial Catalog=SMSPLUS1;User ID=sa;Password=ex0du$2o13;"
        ElseIf flag = 1 Then
            db_conn = "Provider=sqloledb;Data Source=" & SVR & ",1433;Network Library=DBMSSOCN;Initial Catalog=SMSPLUS01;User ID=" & dbid & ";Password=" & dbpw & ";"
        ElseIf flag = 5 Then
            db_conn = "Data Source=" & SVR & ";Initial Catalog=KEMS01;Persist Security Info=True;User ID=" & dbid & ";Password=" & dbpw
        ElseIf flag = 6 Then
            db_conn = "Provider=MSDASQL;Data Source = SQLite3RIT;Use Procedure for Prepare=1;Auto Translate=True;Packet Size=8000;Connect Timeout=90"
        Else 'OleSQLite.SQLiteSource.1
            db_conn = False
            '           db_conn = True
        End If

    End Function

    Function xls_conn(ByVal File_Name As String) As String
        xls_conn = "Provider = MSDASQL;Driver={Microsoft Excel Driver (*.xls)};DBQ=" & File_Name & "; ReadOnly=False;"
    End Function

    Function Query_TEXTBOXDROPDOWN(ByVal S As CheckedListBox, ByVal Q_TEXT As String) As Boolean
        Dim conn As New ADODB.Connection
        Dim cmd As New ADODB.Command
        Dim rs As New ADODB.Recordset
        Dim I As Integer
        '        Dim tvRoot As TreeNode

        '        tvRoot = S.Nodes.Add(MODEL)
        S.items.Clear()

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
            Query_TEXTBOXDROPDOWN = True

            For I = 0 To rs.RecordCount - 1
                S.items.Add(rs(0).Value)
                rs.MoveNext()
            Next
        Else
            Query_TEXTBOXDROPDOWN = False
        End If

        rs = Nothing

    End Function

    Function Query_Combo(ByVal S As ComboBox, ByVal Q_TEXT As String) As Boolean
        Dim conn As New ADODB.Connection
        Dim cmd As New ADODB.Command
        Dim rs As New ADODB.Recordset
        Dim I As Integer

        S.items.Clear()

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
            Query_Combo = True

            For I = 0 To rs.RecordCount - 1
                S.items.Add(rs(0).Value)

                rs.MoveNext()
            Next

        Else
            Query_Combo = False
        End If

        rs = Nothing

    End Function

    Function Query_CheckList(ByVal S As CheckedListBox, ByVal Q_TEXT As String) As Boolean
        Dim conn As New ADODB.Connection
        Dim cmd As New ADODB.Command
        Dim rs As New ADODB.Recordset
        Dim I As Integer

        S.Items.Clear()

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
            Query_CheckList = True

            For I = 0 To rs.RecordCount - 1
                S.Items.Add(rs(0).Value)

                rs.MoveNext()
            Next

        Else
            Query_CheckList = False
        End If

        rs = Nothing

    End Function

    Function Query_Combo1(ByVal S As DevComponents.DotNetBar.ComboBoxItem, ByVal Q_TEXT As String) As Boolean
        Dim conn As New ADODB.Connection
        Dim cmd As New ADODB.Command
        Dim rs As New ADODB.Recordset
        Dim I As Integer
        '        Dim tvRoot As TreeNode

        '        tvRoot = S.Nodes.Add(MODEL)
        S.items.Clear()

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
            Query_Combo1 = True

            For I = 0 To rs.RecordCount - 1
                S.items.Add(rs(0).Value)

                rs.MoveNext()
            Next

        Else
            Query_Combo1 = False
        End If

        rs = Nothing

    End Function

    Function Query_Combo2(ByVal S As DevComponents.DotNetBar.Controls.ComboBoxEx, ByVal TableName As String, ByVal TableDispayCol As String, ByVal TablevalueCol As String, ByVal QueryCondi As String, ByVal ALL_YN As Boolean) As Boolean
        Try
            'Query작성시, 앞에 컬럼은 Valuemember, 두번째 컬럼은 Displaymember가 됨
            'DisplayMember : 콤보박스에 표시할 컬럼값
            'ValueMember : 콤보박스에서 선택하면 실제 넘겨받는 값
            Dim connection As New SqlConnection(db_conn(5))
            Dim Q_TEXT As String = "select distinct " & TablevalueCol & "," & TableDispayCol & " from " & TableName & " where " & QueryCondi
            Dim adapter As New SqlDataAdapter(Q_TEXT, connection)

            Dim Ds As New DataSet
            Ds.Tables.Add(TableName)
            Ds.Tables(0).Columns.Add(TablevalueCol)
            Ds.Tables(0).Columns.Add(TableDispayCol)

            If ALL_YN = True Then
                Dim TbNewRow As DataRow = Ds.Tables(0).NewRow
                TbNewRow(0) = ""
                TbNewRow(1) = "ALL"
                Ds.Tables(0).Rows.Add(TbNewRow)
            End If
            adapter.Fill(Ds.Tables(0))

            S.DataSource = Ds.Tables(0)
            S.DisplayMember = TableDispayCol
            S.ValueMember = TablevalueCol

        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Function

    Function Query_WHCombo(ByVal S As DevComponents.DotNetBar.Controls.ComboBoxEx, ByVal SiteID As String, ByVal ALL_YN As Boolean) As Boolean
        Try
            'Query작성시, 앞에 컬럼은 Valuemember, 두번째 컬럼은 Displaymember가 됨
            'DisplayMember : 콤보박스에 표시할 컬럼값
            'ValueMember : 콤보박스에서 선택하면 실제 넘겨받는 값
            If Query_Combo(S, "SELECT '['+CODE_ID+'] ' +CODE_NAME FROM TBL_CODEMASTER WHERE CLASS_ID = '10024' and active = 'Y' ORDER BY CODE_ID") = True Then
                S.Items.Add("ALL")
                S.Text = "ALL"
            End If

        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Function

    Function Query_WHCombo2(ByVal S As DevComponents.DotNetBar.Controls.ComboBoxEx, ByVal SiteID As String, ByVal ALL_YN As Boolean) As Boolean
        Try
            'Query작성시, 앞에 컬럼은 Valuemember, 두번째 컬럼은 Displaymember가 됨
            'DisplayMember : 콤보박스에 표시할 컬럼값
            'ValueMember : 콤보박스에서 선택하면 실제 넘겨받는 값
            Dim qry As String = "SELECT '['+CODE_ID+'] ' +CODE_NAME as aa FROM TBL_CODEMASTER WHERE CLASS_ID = '10024' and active = 'Y' ORDER BY CODE_ID" & vbNewLine

            If Query_Combo(S, qry) = True Then
                S.Items.Add("[S2014-0001] 엑소더스 일렉트론")
                S.Items.Add("[U2014-0001] 엑소더스 와이어레스")
                S.Items.Add("[S2014-0000] 자체처리")
                S.Items.Add("ALL")
                S.Text = "ALL"
            End If

        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Function


    Function Query_RS(ByVal Q_TEXT) As String
        Dim conn As New ADODB.Connection
        Dim cmd As New ADODB.Command
        Dim rs As New ADODB.Recordset
        Query_RS = ""

        Try
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

            If rs.EOF = True Or rs.BOF = True Then
                Query_RS = ""
            Else
                If rs.RecordCount > 0 And IsDBNull(rs(0).Value) = False Then
                    Query_RS = rs(0).Value
                Else
                    Query_RS = ""
                End If
            End If

            rs = Nothing
            conn.Close()

        Catch ex As Exception

        End Try

    End Function


    Function Query_RS_Img(ByVal Q_TEXT) As Byte()
        Dim conn As New ADODB.Connection
        Dim cmd As New ADODB.Command
        Dim rs As New ADODB.Recordset
        Query_RS_Img = Nothing

        Try
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

            If rs.EOF = True Or rs.BOF = True Then
                Query_RS_Img = Nothing
            Else
                If rs.RecordCount > 0 And IsDBNull(rs(0).Value) = False Then
                    Query_RS_Img = rs(0).Value
                Else
                    Query_RS_Img = Nothing
                End If
            End If

            rs = Nothing
            conn.Close()

        Catch ex As Exception

        End Try

    End Function

    Function Query_RS_Img_rit(ByVal Q_TEXT) As Byte()
        Dim conn As New ADODB.Connection
        Dim cmd As New ADODB.Command
        Dim rs As New ADODB.Recordset
        Query_RS_Img_rit = Nothing

        Try
            conn.Open(db_conn(6))
            cmd.ActiveConnection = conn

            cmd.CommandType = ADODB.CommandTypeEnum.adCmdText
            cmd.CommandText = Q_TEXT

            rs = New ADODB.Recordset
            rs.ActiveConnection = conn
            rs.CursorLocation = ADODB.CursorLocationEnum.adUseClient
            rs.CursorType = ADODB.CursorTypeEnum.adOpenForwardOnly
            rs.LockType = ADODB.LockTypeEnum.adLockReadOnly

            rs.Open(cmd.CommandText)

            If rs.EOF = True Or rs.BOF = True Then
                Query_RS_Img_rit = Nothing
            Else
                If rs.RecordCount > 0 And IsDBNull(rs(0).Value) = False Then
                    Query_RS_Img_rit = rs(0).Value
                Else
                    Query_RS_Img_rit = Nothing
                End If
            End If

            rs = Nothing
            conn.Close()

        Catch ex As Exception

        End Try

    End Function

    Function Query_RS_ALL(ByVal Q_TEXT) As ADODB.Recordset
        Dim conn As New ADODB.Connection
        Dim cmd As New ADODB.Command
        Dim rs As New ADODB.Recordset

        Query_RS_ALL = Nothing
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

        If rs Is Nothing Then
        Else
            If rs.RecordCount > 0 Then
                Query_RS_ALL = rs
            Else
                Query_RS_ALL = Nothing
            End If
        End If

        'rs = Nothing

    End Function

    Function Query_Cell_Code1(ByVal col_nm As String, ByVal class_id As String) As String()
        'tbl_codemaster의 코드 조회

        Try
            'code_master 에서 셀에 뿌려줄 코드들 조회
            Dim conn As New ADODB.Connection
            Dim cmd As New ADODB.Command
            Dim rs As New ADODB.Recordset
            Dim i As Integer = 0
            Dim Arry As String() = Nothing

            conn.Open(db_conn(1))
            cmd.ActiveConnection = conn

            cmd.CommandType = ADODB.CommandTypeEnum.adCmdText
            cmd.CommandText = "Select " & col_nm & " from tbl_codemaster where site_id ='" & Site_id & "' and class_id = '" & class_id & "' and active='y' order by dis_order"
            rs = New ADODB.Recordset
            rs.ActiveConnection = conn
            rs.CursorLocation = ADODB.CursorLocationEnum.adUseClient
            rs.CursorType = ADODB.CursorTypeEnum.adOpenForwardOnly
            rs.LockType = ADODB.LockTypeEnum.adLockReadOnly

            rs.Open(cmd.CommandText)
            ReDim Arry(rs.RecordCount - 1)
            Do Until (rs.EOF = True)
                Arry(i) = rs(0).Value
                rs.MoveNext()
                i += 1
            Loop

            rs = Nothing
            Query_Cell_Code1 = Arry

        Catch ex As Exception
            MessageBox.Show(ex.Message)
            Query_Cell_Code1 = New String() {}
        End Try
    End Function

    Function QryComboCell1(ByVal Qtext As String) As Array
        Try
            'code_master 에서 셀에 뿌려줄 코드들 조회
            Dim conn As New ADODB.Connection
            Dim cmd As New ADODB.Command
            Dim rs As New ADODB.Recordset
            Dim i As Integer = 0
            Dim Arry(1, 10) As String

            conn.Open(db_conn(1))
            cmd.ActiveConnection = conn

            cmd.CommandType = ADODB.CommandTypeEnum.adCmdText
            cmd.CommandText = Qtext
            rs = New ADODB.Recordset
            rs.ActiveConnection = conn
            rs.CursorLocation = ADODB.CursorLocationEnum.adUseClient
            rs.CursorType = ADODB.CursorTypeEnum.adOpenForwardOnly
            rs.LockType = ADODB.LockTypeEnum.adLockReadOnly

            rs.Open(cmd.CommandText)
            ReDim Arry(1, rs.RecordCount - 1)
            Do Until (rs.EOF = True)
                Arry(0, i) = rs(0).Value
                Arry(1, i) = rs(1).Value
                rs.MoveNext()
                i += 1
            Loop
            Arry(0, i) = ""
            Arry(1, i) = ""
            rs = Nothing
            QryComboCell1 = Arry

        Catch ex As Exception
            MessageBox.Show(ex.Message)
            QryComboCell1 = New String() {}
        End Try
    End Function

    Function Query_Cell_Code2(ByVal Q_TEXT) As String()
	try
        '셀에 뿌려줄 코드를 조회
        Dim conn As New ADODB.Connection
        Dim cmd As New ADODB.Command
        Dim rs As New ADODB.Recordset
        Dim i As Integer = 0
        Dim Arry As String() = Nothing

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
            ReDim Arry(rs.RecordCount - 1)
        Do Until (rs.EOF = True)
            Arry(i) = rs(0).Value
            rs.MoveNext()
            i += 1
            Loop

        rs = Nothing
        Query_Cell_Code2 = Arry
	Catch ex As Exception
            MessageBox.Show(ex.Message)
            Query_Cell_Code2 = New String() {}
    End Try
    End Function

    Function Insert_Data(ByVal Q_TEXT As String) As Boolean
        '      Try

        Dim conn As New ADODB.Connection
        Dim cmd As New ADODB.Command
        Dim rs As New ADODB.Recordset

        conn.Open(db_conn(1))
        cmd.ActiveConnection = conn

        cmd.CommandType = ADODB.CommandTypeEnum.adCmdText
        cmd.CommandText = Q_TEXT
        cmd.CommandTimeout = 90

        rs = New ADODB.Recordset
        rs.ActiveConnection = conn
        rs.CursorLocation = ADODB.CursorLocationEnum.adUseClient
        rs.CursorType = ADODB.CursorTypeEnum.adOpenForwardOnly
        rs.LockType = ADODB.LockTypeEnum.adLockReadOnly

        'rs.Open(cmd.CommandText)
        rs = cmd.Execute

        If rs.State = ADODB.ObjectStateEnum.adStateClosed Then
            Insert_Data = True
        Else
            Insert_Data = False
        End If

        rs = Nothing


        'Catch ex As Exception
        '    MessageBox.Show("Error : " & ex.Message, "ERROR")
        'End Try
    End Function

    Function Insert_Data_USA(ByVal Q_TEXT As String) As Boolean

        Dim conn As New ADODB.Connection
        Dim cmd As New ADODB.Command
        Dim rs As New ADODB.Recordset

        conn.Open(db_conn(2))
        cmd.ActiveConnection = conn

        cmd.CommandType = ADODB.CommandTypeEnum.adCmdText
        cmd.CommandText = Q_TEXT
        cmd.CommandTimeout = 600

        rs = New ADODB.Recordset
        rs.ActiveConnection = conn
        rs.CursorLocation = ADODB.CursorLocationEnum.adUseClient
        rs.CursorType = ADODB.CursorTypeEnum.adOpenForwardOnly
        rs.LockType = ADODB.LockTypeEnum.adLockReadOnly

        rs = cmd.Execute

        If rs.State = ADODB.ObjectStateEnum.adStateClosed Then
            Insert_Data_USA = True
        Else
            Insert_Data_USA = False
        End If

        rs = Nothing

    End Function

    Function Insert_Data_COS(ByVal Q_TEXT As String) As Boolean
        '      Try

        Dim conn As New ADODB.Connection
        Dim cmd As New ADODB.Command
        Dim rs As New ADODB.Recordset

        conn.Open(db_conn(1))
        cmd.ActiveConnection = conn

        cmd.CommandType = ADODB.CommandTypeEnum.adCmdText
        cmd.CommandText = Q_TEXT
        cmd.CommandTimeout = 90

        rs = New ADODB.Recordset
        rs.ActiveConnection = conn
        rs.CursorLocation = ADODB.CursorLocationEnum.adUseClient
        rs.CursorType = ADODB.CursorTypeEnum.adOpenForwardOnly
        rs.LockType = ADODB.LockTypeEnum.adLockReadOnly

        'rs.Open(cmd.CommandText)
        rs = cmd.Execute

        If rs.State = ADODB.ObjectStateEnum.adStateClosed Then
            Insert_Data_COS = True
        Else
            If rs(0).Value = "ROLLBACK" Then
                Insert_Data_COS = False
            End If
            Insert_Data_COS = False
        End If

        rs = Nothing


        'Catch ex As Exception
        '    MessageBox.Show("Error : " & ex.Message, "ERROR")
        'End Try
    End Function



    Function DB2XLS(ByVal S As FarPoint.Win.Spread.FpSpread, ByVal fname As String, ByVal SheetName As String, ByVal HD_YN As Boolean, ByVal Comment As String, ByVal Q_TEXT As String) As Boolean

        Dim conn As New ADODB.Connection
        Dim cmd As New ADODB.Command
        Dim rs As New ADODB.Recordset
        Dim i, j As Integer

        Try


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
            If rs.RecordCount < 1 Then
                Exit Function
            End If

            S.ActiveSheet.Rows.Clear()

            If Comment <> "" Then  '데이터를 표시하기전에 코멘트가 있는 경우 표시
                S.ActiveSheet.RowCount += 1
                S.ActiveSheet.SetValue(S.ActiveSheet.RowCount - 1, 0, Comment)
            Else
            End If

            MainFrm.ProgressBarItem1.Maximum = rs.RecordCount

            For i = 0 To rs.RecordCount - 1
                S.ActiveSheet.RowCount += 1
                For j = 0 To rs.Fields.Count - 1
                    S.ActiveSheet.SetValue(S.ActiveSheet.RowCount - 1, j, rs.Fields(j).Value)
                Next
                MainFrm.ProgressBarItem1.Value = i
                rs.MoveNext()
            Next
            'Save the Workbook and Quit Excel
            S.SaveExcel(fname, FarPoint.Excel.ExcelSaveFlags.SaveCustomColumnHeaders)
            'S.ActiveSheet.SaveTextFile(fname, True)

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error")
        End Try
    End Function

    Function DB2XLS_NEW(ByVal S As FarPoint.Win.Spread.FpSpread, ByVal fname As String, ByVal SheetName As String, ByVal HD_YN As Boolean, ByVal Q_TEXT As String) As Boolean

        Dim conn As New ADODB.Connection
        Dim cmd As New ADODB.Command
        Dim rs As New ADODB.Recordset
        Dim i, j As Integer

        Try


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
            If rs.RecordCount < 1 Then
                Exit Function
            End If

            S.ActiveSheet.Rows.Clear()

            S.ActiveSheet.RowCount += 1
            S.ActiveSheet.SetValue(S.ActiveSheet.RowCount - 1, 0, "Part No")
            S.ActiveSheet.SetValue(S.ActiveSheet.RowCount - 1, 1, "Order Qty")

            MainFrm.ProgressBarItem1.Maximum = rs.RecordCount

            For i = 0 To rs.RecordCount - 1
                S.ActiveSheet.RowCount += 1
                For j = 0 To rs.Fields.Count - 1
                    S.ActiveSheet.SetValue(S.ActiveSheet.RowCount - 1, j, rs.Fields(j).Value)
                Next
                MainFrm.ProgressBarItem1.Value = i
                rs.MoveNext()
            Next
            'Save the Workbook and Quit Excel
            S.SaveExcel(fname, FarPoint.Excel.ExcelSaveFlags.SaveCustomColumnHeaders)
            'S.ActiveSheet.SaveTextFile(fname, True)

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error")
        End Try
    End Function

    Function DB2XLS_DIR(ByVal fname As String, ByVal SheetName As String, ByVal HD_YN As Boolean, ByVal Comment As String, ByVal Q_TEXT As String) As Boolean

        Dim conn As New ADODB.Connection
        Dim cmd As New ADODB.Command
        Dim rs As New ADODB.Recordset
        Dim i, j, k As Integer

        Try


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
            If rs.RecordCount < 1 Then
                Exit Function
            End If
            'Create a new workbook in Excel
            Dim oExcel As Object
            Dim oBook As Object
            Dim oSheet As Object
            oExcel = CreateObject("Excel.Application")
            oBook = oExcel.Workbooks.Add
            oSheet = oBook.Worksheets(1)
            oSheet.name = SheetName

            If HD_YN = True Then '헤더표시가 True이면 첫번째행에 컬럼명 표시
                For i = 0 To rs.Fields.Count - 1
                    oSheet.Range(Chr(65 + i) & 1).value = rs.Fields(i).Name
                Next
                k = 2
            Else
                k = 1
            End If

            If Comment <> "" Then  '데이터를 표시하기전에 코멘트가 있는 경우 표시
                oSheet.Range(Chr(65) & k).value = Comment
                k += 1
            Else
            End If

            MainFrm.ProgressBarItem1.Maximum = rs.RecordCount

            For i = 0 To rs.RecordCount - 1
                For j = 0 To rs.Fields.Count - 1
                    oSheet.Range(Chr(65 + j) & i + k).value = rs.Fields(j).Value
                Next
                MainFrm.ProgressBarItem1.Value = i
                rs.MoveNext()
            Next


            'Save the Workbook and Quit Excel
            oBook.SaveAs(fname)
            oExcel.Quit()

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error")
        End Try
    End Function

    Function FSendMail(ByVal MsrvNm As String, ByVal MId As String, ByVal Mpw As String, ByVal Fmail As String, ByVal Tmail As String(), ByVal mTitle As String, ByVal mbody As String) As Boolean
        Try
            FSendMail = False

            Dim mailServerName As String = MsrvNm
            Dim mailClient As SmtpClient = New SmtpClient
            Dim SmtpUser As New System.Net.NetworkCredential()
            Dim message As MailMessage = New MailMessage()
            Dim i As Integer
            'Dim userToken As Object

            message.IsBodyHtml = True
            message.Subject = mTitle
            message.Body = mbody
            message.From = New MailAddress(Fmail)
            For i = 0 To Tmail.GetLength(0) - 1
                message.To.Add(Tmail(i))
            Next

            mailClient.Host = mailServerName
            mailClient.UseDefaultCredentials = False
            mailClient.Credentials = New System.Net.NetworkCredential(MId, Mpw)
            '            mailClient.EnableSsl = True
            mailClient.DeliveryMethod = SmtpDeliveryMethod.Network


            'message.Headers.Add("http://schemas.microsoft.cdo.configuration/smtpauthenticate", 1)
            'message.Headers.Add("http://schemas.microsoft.cdo.configuration/sendusername", MId)
            'message.Headers.Add("http://schemas.microsoft.cdo.configuration/sendpassword", Mpw)
            'message.Headers.Add("http://schemas.microsoft.cdo.configuration/sendusing", 2)
            'message.Headers.Add("http://schemas.microsoft.cdo.configuration/smtpserver", mailServerName)
            'message.Headers.Add("http://schemas.microsoft.cdo.configuration/smtpconnectiontimeout", 10)
            'message.Headers.Add("http://schemas.microsoft.cdo.configuration/smtpserverport", 25)
            'message.Headers.Add("http://schemas.microsoft.cdo.configuration/smtpusessl", False)

            mailClient.Send(message)
            message.Dispose()


            FSendMail = True
            MessageBox.Show("Success Send to mail")

        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message, "ERROR")
        End Try
    End Function

    Function GET_TRG(ByVal ESN As String, ByVal type As String) As String()

        Dim RS As ADODB.Recordset
        Dim QRY As String
        Dim ARR As String() = Nothing

        ReDim ARR(5)
        ARR(0) = "N/A"
        ARR(1) = "N/A"
        ARR(2) = "N/A"
        ARR(3) = "N/A"
        ARR(4) = "N/A"

        QRY = "select DEF_CD from tbl_TRIAGE" & vbNewLine
        If type = "FLIP" Then
            QRY = QRY & "where ESN_OBID = (SELECT LOT_NO FROM TBL_LOTMASTER WHERE LOT_NO = '" & ESN & "')" & vbNewLine
        Else
            QRY = QRY & "where ESN_OBID = (SELECT OBID FROM TBL_ESNMASTER WHERE ESN = '" & ESN & "')" & vbNewLine
        End If
        QRY = QRY & "ORDER BY DEF_CD DESC" & vbNewLine

        RS = Query_RS_ALL(QRY)

        If RS Is Nothing Then
        Else
            For I As Integer = 0 To RS.RecordCount - 1
                ARR(I) = RS(0).Value
                RS.MoveNext()
            Next
        End If
        GET_TRG = ARR

    End Function

    Function CONVERT_MEID(ByVal ESN As DevComponents.DotNetBar.Controls.TextBoxX, ByVal ESN1 As DevComponents.DotNetBar.Controls.TextBoxX, ByVal ESN2 As DevComponents.DotNetBar.Controls.TextBoxX) As String

        If Len(ESN.Text) = 15 Or Len(ESN.Text) = 14 Or Len(ESN.Text) = 10 Then
            CONVERT_MEID = "HEX"
            ESN1.Text = ESN.Text
            ESN2.Text = HEX_TO_DEC2(ESN.Text)
        ElseIf Len(ESN.Text) = 18 Then
            CONVERT_MEID = "DEC"
            ESN1.Text = DEC_TO_HEX2(ESN.Text)
            ESN2.Text = ESN.Text
            ESN.Text = ESN1.Text
        Else
            CONVERT_MEID = "WRONG"
            ESN1.Text = ""
            ESN2.Text = ""
        End If

    End Function

    Function CONVERT_MEID_NC(ByVal ESN As DevComponents.DotNetBar.Controls.TextBoxX, ByVal ESN1 As DevComponents.DotNetBar.Controls.TextBoxX, ByVal ESN2 As DevComponents.DotNetBar.Controls.TextBoxX) As String

        If Len(ESN.Text) = 15 Or Len(ESN.Text) = 14 Or Len(ESN.Text) = 10 Then
            CONVERT_MEID_NC = "HEX"
            ESN1.Text = ESN.Text
            ESN2.Text = HEX_TO_DEC2(ESN.Text)
        ElseIf Len(ESN.Text) = 18 Then
            CONVERT_MEID_NC = "DEC"
            ESN1.Text = DEC_TO_HEX2(ESN.Text)
            ESN2.Text = ESN.Text
            '            ESN.Text = ESN1.Text
        Else
            CONVERT_MEID_NC = "WRONG"
            ESN1.Text = ""
            ESN2.Text = ""
        End If

    End Function

    Function Chk_Sum(ByVal id As String) As String

        Chk_Sum = ""
        Dim mul As Integer = 3
        Dim sum As Integer = 0

        For i As Integer = 1 To 17
            If mul = 3 Then
                sum = sum + CInt(Mid(id, i, 1)) * mul
                mul = 1
            Else
                sum = sum + CInt(Mid(id, i, 1)) * mul
                mul = 3
            End If
        Next

        Dim aa As String = Microsoft.VisualBasic.Right(CStr(sum), 1)

        'If CInt(Microsoft.VisualBasic.Right(CStr(sum), 1)) > 5 Then
        Chk_Sum = 10 - CInt(Microsoft.VisualBasic.Right(CStr(sum), 1))

        If Chk_Sum = 10 Then
            Chk_Sum = 0
        End If
        'Else
        'Chk_Sum = CInt(Microsoft.VisualBasic.Right(CStr(sum), 1))
        'End If

    End Function

    Function MODEL_SELECTED(ByVal S As CheckedListBox) As String
        MODEL_SELECTED = ""
        For i As Integer = 0 To S.Items.Count - 1
            If S.GetItemChecked(i) = True Then
                If S.Items(i).ToString = "ALL" Then
                    MODEL_SELECTED = ""
                    Exit For
                Else
                    If MODEL_SELECTED = "" Then
                        MODEL_SELECTED = "'" & S.Items(i).ToString & "'"
                    Else
                        MODEL_SELECTED = MODEL_SELECTED & ", '" & S.Items(i).ToString & "'"
                    End If
                End If
            End If
        Next


    End Function

    Function WK_SELECTED(ByVal S As CheckedListBox) As String
        WK_SELECTED = ""
        For i As Integer = 0 To S.Items.Count - 1
            If S.GetItemChecked(i) = True Then
                If S.Items(i).ToString = "ALL" Then
                    WK_SELECTED = ""
                    Exit For
                Else
                    If WK_SELECTED = "" Then
                        WK_SELECTED = "'" & Query_RS("SELECT CODE_ID FROM TBL_CODEMASTER WHERE CLASS_ID = 'R0001' AND CODE_NAME = '" & S.Items(i).ToString & "'") & "'"
                    Else
                        WK_SELECTED = WK_SELECTED & ", '" & Query_RS("SELECT CODE_ID FROM TBL_CODEMASTER WHERE CLASS_ID = 'R0001' AND CODE_NAME = '" & S.Items(i).ToString & "'") & "'"
                    End If
                End If
            End If
        Next


    End Function

    Function Chk_Sum_ICC(ByVal id As String) As String

        Chk_Sum_ICC = ""
        Dim mul As Integer = 2
        Dim sum As Integer = 0
        Dim SUM_ICC As Integer = 0

        For i As Integer = 1 To 19
            If mul = 2 Then
                SUM_ICC = CInt(Mid(id, i, 1)) * mul

                If SUM_ICC > 9 Then
                    SUM_ICC = CInt(Mid(CStr(SUM_ICC), 1, 1)) + CInt(Mid(CStr(SUM_ICC), 2, 1))
                End If

                sum = sum + SUM_ICC
                mul = 1
            Else
                SUM_ICC = CInt(Mid(id, i, 1)) * mul

                If SUM_ICC > 9 Then
                    SUM_ICC = CInt(Mid(CStr(SUM_ICC), 1, 1)) + CInt(Mid(CStr(SUM_ICC), 2, 1))
                End If

                sum = sum + SUM_ICC
                mul = 2
            End If
        Next

        Chk_Sum_ICC = 10 - CInt(Microsoft.VisualBasic.Right(CStr(sum), 1))

        If Chk_Sum_ICC = 10 Then
            Chk_Sum_ICC = 0
        End If

    End Function

    Function Get_Img(ByVal path As String) As Byte()

        Dim stream As System.IO.FileStream = New System.IO.FileStream(path, IO.FileMode.Open, IO.FileAccess.Read)
        Dim reader As System.IO.BinaryReader = New System.IO.BinaryReader(stream)

        Dim img As Byte() = reader.ReadBytes(CInt(stream.Length))

        stream.Close()
        reader.Close()

        Get_Img = img

    End Function

    Function Insert_Img_Data(ByVal OBID As String, ByVal C_PRC As String, ByVal SEQ As Integer) As Boolean
        '      Try

        Dim conn As New SqlClient.SqlConnection
        Dim cmd As New SqlClient.SqlCommand
        Dim aa As Byte()

        conn.ConnectionString = db_conn(5)
        cmd.Connection = conn

        cmd.CommandType = System.Data.CommandType.StoredProcedure
        cmd.CommandText = "SP_IMG_SAVE"
        cmd.CommandTimeout = 90

        Dim img As Drawing.Image
        Dim pic As String

        img = Image.FromFile(img_fname)


        pic = img_fname
        aa = Get_Img(pic)
        'Kill(img_fname)

        Dim param As SqlClient.SqlParameter = New SqlClient.SqlParameter("SITE", "S1000")
        cmd.Parameters.Add(param)

        param = New SqlClient.SqlParameter("OBID", OBID)
        cmd.Parameters.Add(param)

        param = New SqlClient.SqlParameter("C_PRC", C_PRC)
        cmd.Parameters.Add(param)

        param = New SqlClient.SqlParameter("SEQ", SEQ)
        cmd.Parameters.Add(param)

        param = New SqlClient.SqlParameter("IMG", System.Data.SqlDbType.Image, aa.Length)
        param.Value = aa
        cmd.Parameters.Add(param)

        param = New SqlClient.SqlParameter("EMP", Emp_No)
        cmd.Parameters.Add(param)

        conn.Open()
        cmd.ExecuteNonQuery()
        conn.Close()

        '        MessageBox.Show("저장이 완료되었습니다.")

        img.Dispose()
        aa = Nothing
    End Function

    Function DISPLAY_IMG(ByVal OBID As String, ByVal S As PictureBox, ByVal seq As Integer) As Boolean

        'Dim aa_rs As ADODB.Recordset

        'aa_rs = Query_RS_ALL("select right(obid,14)+'_'+process, src_obid, seq_no from emsweb.emsweb1.dbo.tbl_img where obid in (select obid from tbl_esnmaster_b where return_dv = 'INWTY' and model not like 'D%' )")

        'For i As Integer = 0 To aa_rs.RecordCount - 1
        '    OBID = aa_rs(1).Value
        '    seq = aa_rs(2).Value

        '    Dim aa As Byte()

        '    aa = Query_RS_Img("SELECT IMG FROM EMSWEB.EMSWEB1.DBO.TBL_IMG where SRC_OBID = '" & OBID & "' and seq_no = " & seq)

        '    If aa Is Nothing Then
        '    Else
        '        If aa.Length = 0 Then
        '        Else
        '            Dim ms As System.IO.MemoryStream = New System.IO.MemoryStream(aa, 0, aa.Length)
        '            Dim img As Drawing.Image

        '            img = Image.FromStream(ms)
        '            S.Image = img
        '        End If
        '    End If

        '    S.Image.Save("c:\temp\inwty\" & aa_rs(0).Value & "_" & seq & ".jpg")
        '    aa_rs.MoveNext()
        'Next


        Dim aa As Byte()

        aa = Query_RS_Img("SELECT IMG FROM EMSWEB.EMSWEB1.DBO.TBL_IMG where SRC_OBID = '" & OBID & "' and seq_no = " & seq)

        If aa Is Nothing Then
        Else
            If aa.Length = 0 Then
            Else
                Dim ms As System.IO.MemoryStream = New System.IO.MemoryStream(aa, 0, aa.Length)
                Dim img As Drawing.Image

                img = Image.FromStream(ms)
                S.Image = img
            End If
        End If
    End Function


    Function Send_Pmail(ByVal sum As String, ByVal det As String) As Boolean '메일 발송
        Try
            Dim kk, msrvnm, mid, mpw, fmail, mtitle As String
            Dim tmail As String()
            'Dim S1 = Me.FpSpread1.ActiveSheet
            Dim TmailRs As ADODB.Recordset
            Dim i As Integer

            msrvnm = "smtp.naver.com" 'Query_RS("select code_name from tbl_codemaster where class_id = 'R1001' and code_id = 'HOST'")
            mid = "udiatech" 'Query_RS("select code_name from tbl_codemaster where class_id = 'R1001' and code_id = 'LOGID'")
            mpw = "ud1atech" 'Query_RS("select code_name from tbl_codemaster where class_id = 'R1001' and code_id = 'LOGPW'")
            fmail = "udiatech@naver.com" 'Query_RS("select code_name from tbl_codemaster where class_id = 'R1001' and code_id = 'FMAIL'")
            '메일 수신자 추가는 tbl_codemaster의 R1002 클래스에 추가하면 됨
            TmailRs = Query_RS_ALL("select code_name from tbl_codemaster where class_id = 'R1002' and active = 'Y'")

            If TmailRs Is Nothing Then
                Exit Function
            End If

            ReDim tmail(TmailRs.RecordCount - 1)

            For i = 0 To TmailRs.RecordCount - 1
                tmail(i) = TmailRs(0).Value
                TmailRs.MoveNext()
            Next

            '메일 본문내용 수정시 아래 부분을 고치면 됨
            If sum = "1" Then
                mtitle = "EXODUS KOREA PURCHASE ORDER[" & det & "] 가 등록되었습니다."
                kk = "EXODUS KOREA PURCHASE ORDER[" & det & "] 의 승인을 요청합니다."
            Else
                mtitle = "EXODUS KOREA PURCHASE ORDER[" & det & "] 가 승인되었습니다."
                kk = "EXODUS KOREA PURCHASE ORDER[" & det & "] 의 승인되었습니다." & vbNewLine
                kk = kk & "입고 진행하시기 바랍니다." & vbNewLine

            End If

            If Send_Mail(msrvnm, mid, mpw, fmail, tmail, mtitle, kk) = True Then
                Send_Pmail = True
            Else
                Send_Pmail = False
            End If


        Catch ex As Exception
            '            MessageBox.Show("Error: " & ex.Message, "ERROR")
        End Try
    End Function


    Function Send_Mail(ByVal MsrvNm As String, ByVal MId As String, ByVal Mpw As String, ByVal Fmail As String, ByVal Tmail As String(), ByVal mTitle As String, ByVal mbody As String) As Boolean
        Try
            Send_Mail = False

            Dim mailServerName As String = MsrvNm
            Dim mailClient As SmtpClient = New SmtpClient
            Dim SmtpUser As New System.Net.NetworkCredential()
            Dim message As MailMessage = New MailMessage()
            Dim i As Integer
            'Dim userToken As Object

            message.IsBodyHtml = True
            message.Subject = mTitle
            message.Body = mbody
            message.From = New MailAddress(Fmail)
            For i = 0 To Tmail.GetLength(0) - 1
                message.To.Add(Tmail(i))
            Next

            mailClient.Host = mailServerName
            mailClient.UseDefaultCredentials = False
            mailClient.Credentials = New System.Net.NetworkCredential(MId, Mpw)
            mailClient.EnableSsl = True
            mailClient.DeliveryMethod = SmtpDeliveryMethod.Network


            'message.Headers.Add("http://schemas.microsoft.cdo.configuration/smtpauthenticate", 1)
            'message.Headers.Add("http://schemas.microsoft.cdo.configuration/sendusername", MId)
            'message.Headers.Add("http://schemas.microsoft.cdo.configuration/sendpassword", Mpw)
            'message.Headers.Add("http://schemas.microsoft.cdo.configuration/sendusing", 2)
            'message.Headers.Add("http://schemas.microsoft.cdo.configuration/smtpserver", mailServerName)
            'message.Headers.Add("http://schemas.microsoft.cdo.configuration/smtpconnectiontimeout", 10)
            'message.Headers.Add("http://schemas.microsoft.cdo.configuration/smtpserverport", 25)
            'message.Headers.Add("http://schemas.microsoft.cdo.configuration/smtpusessl", False)

            mailClient.Send(message)
            message.Dispose()


            Send_Mail = True
            '           MessageBox.Show("Success Send to mail")

        Catch ex As Exception
            '          MessageBox.Show("Error: " & ex.Message, "ERROR")
        End Try
    End Function


End Module
