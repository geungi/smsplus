Imports System.IO
Imports System.Net
Imports FarPoint.Excel

Module MdFile

    Function xml_Open(ByVal f As OpenFileDialog, ByVal t As TextBox, ByVal aa As String) As Boolean
        Try


            f.InitialDirectory = My.Computer.FileSystem.SpecialDirectories.MyDocuments
            f.Filter = "XML File (*.xml)|*.xml*|All Files(*.*)|*.*"
            If f.ShowDialog() = Windows.Forms.DialogResult.OK Then
                t.Text = f.FileName


            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Function


    Function File_Open_Meiddet(ByVal f As OpenFileDialog, ByVal s As FarPoint.Win.Spread.FpSpread, ByVal t As TextBox) As Boolean
        Try
            s.ActiveSheet.RowCount = 0

            f.InitialDirectory = My.Computer.FileSystem.SpecialDirectories.MyDocuments
            f.Filter = "Microsoft Office Excel File (*.xlsx)|*.xlsx*|All Files(*.*)|*.*"
            If f.ShowDialog() = Windows.Forms.DialogResult.OK Then
                t.Text = f.FileName


                s.ActiveSheet.OpenExcel(t.Text, 0)

                '                s.OpenExcel(t.Text, FarPoint.Excel.ExcelOpenFlags.DataOnly)
                s.ActiveSheet.RemoveRows(s.ActiveSheet.NonEmptyRowCount(), s.ActiveSheet.RowCount - s.ActiveSheet.NonEmptyRowCount())
                s.ActiveSheet.RemoveColumns(s.ActiveSheet.NonEmptyColumnCount, s.ActiveSheet.ColumnCount - s.ActiveSheet.NonEmptyColumnCount)
            End If
            File_Open_Meiddet = True
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Function

    Function File_Open_Meiddet1(ByVal path As String, ByVal s As FarPoint.Win.Spread.FpSpread) As Boolean
        Try


            's.ActiveSheet.OpenExcel(path, 0)

            s.OpenExcel(path, ExcelOpenFlags.DataOnly)
            s.ActiveSheet.Rows(0).Remove()
            s.ActiveSheet.RowCount = 1

            '                s.OpenExcel(t.Text, FarPoint.Excel.ExcelOpenFlags.DataOnly)
            's.ActiveSheet.RemoveRows(s.ActiveSheet.NonEmptyRowCount(), s.ActiveSheet.RowCount - s.ActiveSheet.NonEmptyRowCount())
            's.ActiveSheet.RemoveColumns(s.ActiveSheet.NonEmptyColumnCount, s.ActiveSheet.ColumnCount - s.ActiveSheet.NonEmptyColumnCount)

            File_Open_Meiddet1 = True
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Function

    Function File_Open(ByVal f As OpenFileDialog, ByVal s As FarPoint.Win.Spread.FpSpread, ByVal t As TextBox, ByVal aa As String) As Boolean
        Try
            s.ActiveSheet.RowCount = 0

            f.InitialDirectory = My.Computer.FileSystem.SpecialDirectories.MyDocuments
            f.Filter = "Microsoft Office Excel File (*.xls)|*.xls*|All Files(*.*)|*.*"
            If f.ShowDialog() = Windows.Forms.DialogResult.OK Then
                t.Text = f.FileName


                s.OpenExcel(t.Text, FarPoint.Excel.ExcelOpenFlags.DataOnly)
                s.ActiveSheet.RemoveRows(s.ActiveSheet.NonEmptyRowCount(), s.ActiveSheet.RowCount - s.ActiveSheet.NonEmptyRowCount())
                s.ActiveSheet.RemoveColumns(s.ActiveSheet.NonEmptyColumnCount, s.ActiveSheet.ColumnCount - s.ActiveSheet.NonEmptyColumnCount)

                If aa = "FrmOpenra" Then
                    If Spread_Setting(s, "FrmOpenra") = True Then
                        Spread_AutoCol(s)
                    End If
                ElseIf aa = "FrmRAUpload" Then
                    If Spread_Setting(s, "FrmRAUpload") = True Then
                        Spread_AutoCol(s)
                    End If
                ElseIf aa = "FRMWIPPARTINV" Then
                    If Spread_Setting(s, "FRMWIPPARTINV") = True Then
                        ' Spread_AutoCol(s)
                    End If
                End If

            End If
            File_Open = True
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Function



    Function File_Open2(ByVal f As OpenFileDialog, ByVal s As FarPoint.Win.Spread.FpSpread, ByVal t As TextBox, ByVal aa As String) As Boolean
        Try

            s.ActiveSheet.RowCount = 0

            f.InitialDirectory = My.Computer.FileSystem.SpecialDirectories.MyDocuments
            f.Filter = "Microsoft Office Excel File (*.xls)|*.xls*|All Files(*.*)|*.*"
            If f.ShowDialog() = Windows.Forms.DialogResult.OK Then
                t.Text = f.FileName

                s.Open(t.Text)
                '              s.OpenExcel(t.Text, FarPoint.Excel.ExcelOpenFlags.DataOnly)
                s.ActiveSheet.RemoveRows(0, 1)
                s.ActiveSheet.RemoveRows(s.ActiveSheet.NonEmptyRowCount(), s.ActiveSheet.RowCount - s.ActiveSheet.NonEmptyRowCount())
                s.ActiveSheet.RemoveRows(s.ActiveSheet.RowCount - 1, 1)

                If aa = "FrmOpenra" Then
                    If Spread_Setting(s, "FrmOpenra") = True Then
                        Spread_AutoCol(s)
                    End If
                ElseIf aa = "FrmRAUpload" Then
                    If Spread_Setting(s, "FrmRAUpload") = True Then
                        Spread_AutoCol(s)
                    End If
                End If

            End If

            File_Open2 = True

        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Function


    Function File_Open3(ByVal f As OpenFileDialog, ByVal s As FarPoint.Win.Spread.FpSpread, ByVal t As TextBox, ByVal aa As String) As Boolean
        Try

            s.ActiveSheet.RowCount = 0

            f.InitialDirectory = My.Computer.FileSystem.SpecialDirectories.MyDocuments
            f.Filter = "Microsoft Office Excel File (*.xls)|*.xls*|All Files(*.*)|*.*"
            If f.ShowDialog() = Windows.Forms.DialogResult.OK Then
                t.Text = f.FileName

                s.Open(t.Text)
                '              s.OpenExcel(t.Text, FarPoint.Excel.ExcelOpenFlags.DataOnly)
                s.ActiveSheet.RemoveRows(0, 1)
                s.ActiveSheet.RemoveRows(s.ActiveSheet.NonEmptyRowCount(), s.ActiveSheet.RowCount - s.ActiveSheet.NonEmptyRowCount())
                '                s.ActiveSheet.RemoveRows(s.ActiveSheet.RowCount - 1, 1)

            End If

            If aa = "FRMWIPPARTINV" Then
                If Spread_Setting(s, "FRMWIPPARTINV") = True Then
                    Spread_AutoCol(s)
                End If
            End If

            File_Open3 = True

        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Function

    Function File_Open4(ByVal f As OpenFileDialog, ByVal s As FarPoint.Win.Spread.FpSpread, ByVal t As TextBox, ByVal aa As String) As Boolean
        Try

            s.ActiveSheet.RowCount = 0
            f.InitialDirectory = My.Computer.FileSystem.SpecialDirectories.MyDocuments
            f.Filter = "Microsoft Office Excel File (*.xls)|*.xls*|All Files(*.*)|*.*"
            If f.ShowDialog() = Windows.Forms.DialogResult.OK Then
                t.Text = f.FileName

                's.Open(t.Text)
                s.OpenExcel(t.Text, FarPoint.Excel.ExcelOpenFlags.DataOnly)
                's.ActiveSheet.RemoveRows(0, 6)
                s.ActiveSheet.RemoveRows(s.ActiveSheet.NonEmptyRowCount(), s.ActiveSheet.RowCount - s.ActiveSheet.NonEmptyRowCount())

            End If

            File_Open4 = True

        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Function

    Function File_Open5(ByVal f As OpenFileDialog, ByVal s As FarPoint.Win.Spread.FpSpread, ByVal t As TextBox, ByVal aa As String) As Boolean
        Try

            s.ActiveSheet.RowCount = 0

            f.InitialDirectory = My.Computer.FileSystem.SpecialDirectories.MyDocuments
            f.Filter = "Microsoft Office Excel File (*.xls)|*.xls*|All Files(*.*)|*.*"
            If f.ShowDialog() = Windows.Forms.DialogResult.OK Then
                t.Text = f.FileName

                s.Open(t.Text)
                '              s.OpenExcel(t.Text, FarPoint.Excel.ExcelOpenFlags.DataOnly)
                s.ActiveSheet.RemoveRows(0, 2)
                s.ActiveSheet.RemoveRows(s.ActiveSheet.NonEmptyRowCount(), s.ActiveSheet.RowCount - s.ActiveSheet.NonEmptyRowCount())
                '                s.ActiveSheet.RemoveRows(s.ActiveSheet.RowCount - 1, 1)

            End If

            If aa = "FRMWIPPARTINV" Then
                If Spread_Setting(s, "FRMWIPPARTINV") = True Then
                    Spread_AutoCol(s)
                End If
            End If

            File_Open5 = True

        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Function

    Function Xls_SpIns(ByVal f As OpenFileDialog, ByVal s As FarPoint.Win.Spread.FpSpread, ByVal t As TextBox, ByVal aa As String) As Boolean
        Xls_SpIns = False
        Try
            's.ActiveSheet.RowCount = 0

            f.InitialDirectory = My.Computer.FileSystem.SpecialDirectories.MyDocuments
            f.Filter = "Microsoft Office Excel File (*.xls)|*.xls*|All Files(*.*)|*.*"
            If f.ShowDialog() = Windows.Forms.DialogResult.OK Then
                t.Text = f.FileName
                SPREAD_XLSIns(f.FileName, s, aa)
            End If
            Xls_SpIns = True
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Function


    Function File_Save(ByVal f As SaveFileDialog, ByVal s As FarPoint.Win.Spread.FpSpread) As Boolean
        Try

            '        s.ActiveSheet.RowCount = 0
            File_Save = False
            f.DefaultExt = "xls"
            f.InitialDirectory = My.Computer.FileSystem.SpecialDirectories.MyDocuments
            'Dim answer As DialogResult

            'answer = MessageBox.Show("엑셀파일로 저장 하시겠습니까?" & vbNewLine & "NO를 선택하면 PDF파일로 저장됩니다.", "파일저장", MessageBoxButtons.YesNo)

            'If answer = Windows.Forms.DialogResult.Yes Then
            f.Filter = "Microsoft Office Excel File (*.xls)|*.xls*|All Files(*.*)|*.*"
            If f.ShowDialog() = Windows.Forms.DialogResult.OK Then
                s.ActiveSheet.Columns(0, s.ActiveSheet.ColumnCount - 1).Locked = False
                s.ActiveSheet.Protect = False


                'If s.Save(f.FileName, True) = True Then

                If s.SaveExcel(f.FileName, FarPoint.Excel.ExcelSaveFlags.SaveCustomColumnHeaders) = True Then
                    File_Save = True
                End If
            End If
            'Else
            'f.DefaultExt = "pdf"
            'f.Filter = "pdf File (*.pdf)|*.pdf*|All Files(*.*)|*.*"
            'If f.ShowDialog() = Windows.Forms.DialogResult.OK Then

            '    s.ActiveSheet.Columns(0, s.ActiveSheet.ColumnCount - 1).Locked = False
            '    s.ActiveSheet.Protect = False

            '    Dim printset As New FarPoint.Win.Spread.PrintInfo()


            '    printset.PrintToPdf = True
            '    printset.PdfFileName = f.FileName
            '    s.Sheets(0).PrintInfo = printset
            '    s.PrintSheet(0)
            '    File_Save = True

            'End If
            'End If

            s.ActiveSheet.Columns(0, s.ActiveSheet.ColumnCount - 1).Locked = True
            s.ActiveSheet.Protect = True

        Catch ex As Exception
            MessageBox.Show("Error:" & ex.Message, "Error")
        End Try
    End Function

    Function File_Save2(ByVal f As SaveFileDialog, ByVal s As FarPoint.Win.Spread.FpSpread, ByVal pkrow As Integer) As Boolean
        Try


            '        s.ActiveSheet.RowCount = 0
            File_Save2 = False
            f.DefaultExt = "xls"
            f.InitialDirectory = My.Computer.FileSystem.SpecialDirectories.MyDocuments
            f.Filter = "Microsoft Office Excel File (*.xls)|*.xls*|All Files(*.*)|*.*"
            If f.ShowDialog() = Windows.Forms.DialogResult.OK Then
                If pkrow > 15 Then
                    s.ActiveSheet.Columns(0, s.ActiveSheet.ColumnCount - 1).Locked = False
                    s.ActiveSheet.Protect = False
                    s.SaveExcel(f.FileName, FarPoint.Excel.ExcelSaveFlags.NoFlagsSet)
                    s.ActiveSheet.Columns(0, s.ActiveSheet.ColumnCount - 1).Locked = True
                    s.ActiveSheet.Protect = True
                Else
                    s.ActiveSheet.Protect = False
                    s.ActiveSheet.Columns(0, s.ActiveSheet.ColumnCount - 1).Locked = False
                    s.SaveExcel(f.FileName, FarPoint.Excel.ExcelSaveFlags.SaveCustomColumnHeaders)
                    File_Save2 = True
                    s.ActiveSheet.Columns(0, s.ActiveSheet.ColumnCount - 1).Locked = True
                    s.ActiveSheet.Protect = True
                End If
            End If
        Catch ex As Exception
            MessageBox.Show("Error:" & ex.Message, "Error")
        End Try

    End Function

    Function File_Save_1(ByVal f As SaveFileDialog, ByVal s As FarPoint.Win.Spread.FpSpread) As Boolean

        '        s.ActiveSheet.RowCount = 0
        File_Save_1 = False
        f.DefaultExt = "xls"
        f.InitialDirectory = My.Computer.FileSystem.SpecialDirectories.MyDocuments
        f.Filter = "Microsoft Office Excel File (*.xls)|*.xls*|All Files(*.*)|*.*"
        If f.ShowDialog() = Windows.Forms.DialogResult.OK Then
            s.ActiveSheet.Columns(0, s.ActiveSheet.ColumnCount - 1).Locked = False
            s.ActiveSheet.Protect = False
            If s.SaveExcel(f.FileName, FarPoint.Excel.ExcelSaveFlags.SaveBothCustomRowAndColumnHeaders) = True Then
                File_Save_1 = True
            End If
            s.ActiveSheet.Columns(0, s.ActiveSheet.ColumnCount - 1).Locked = True
            s.ActiveSheet.Protect = True
        End If


    End Function


    Function File_Save_meiddet(ByVal f As SaveFileDialog, ByVal s As FarPoint.Win.Spread.FpSpread) As Boolean

        '        s.ActiveSheet.RowCount = 0
        File_Save_meiddet = False
        f.DefaultExt = "xlsx"
        f.InitialDirectory = My.Computer.FileSystem.SpecialDirectories.MyDocuments
        f.Filter = "Microsoft Office Excel File (*.xlsx)|*.xlsx*|All Files(*.*)|*.*"
        If f.ShowDialog() = Windows.Forms.DialogResult.OK Then
            s.ActiveSheet.Columns(0, s.ActiveSheet.ColumnCount - 1).Locked = False
            s.ActiveSheet.Protect = False
            If s.SaveExcel(f.FileName, ExcelSaveFlags.UseOOXMLFormat) = True Then
                File_Save_meiddet = True
            End If
            s.ActiveSheet.Columns(0, s.ActiveSheet.ColumnCount - 1).Locked = True
            s.ActiveSheet.Protect = True
        End If


    End Function

    Function File_XML(ByVal f As SaveFileDialog, ByVal s As FarPoint.Win.Spread.FpSpread) As Boolean

        File_XML = False
        f.DefaultExt = "xml"
        f.InitialDirectory = My.Computer.FileSystem.SpecialDirectories.MyDocuments
        f.Filter = "XML File (*.xml)|*.xml*|All Files(*.*)|*.*"
        If f.ShowDialog() = Windows.Forms.DialogResult.OK Then
            s.ActiveSheet.Columns(0, s.ActiveSheet.ColumnCount - 1).Locked = False
            s.ActiveSheet.Protect = False
            If s.ActiveSheet.SaveXml(f.FileName, "c:\rcexodus_20130828_000008.xml") = True Then
                File_XML = True
            End If
            s.ActiveSheet.Columns(0, s.ActiveSheet.ColumnCount - 1).Locked = True
            s.ActiveSheet.Protect = True
        End If


    End Function




    Function DOC_Save(ByVal f As SaveFileDialog, ByVal doc As String, ByVal doc_path As String) As String
        Try
            DOC_Save = ""
            f.FileName = doc
            f.InitialDirectory = My.Computer.FileSystem.SpecialDirectories.MyDocuments
            f.Filter = "All Files(*.*)|*.*"
            If f.ShowDialog() = Windows.Forms.DialogResult.OK Then

                If ftp_download(doc_path, f.FileName) = True Then
                    DOC_Save = f.FileName
                    MessageBox.Show("다운로드 완료.")
                Else
                    MessageBox.Show("다운로드 실패.")
                End If
            End If

        Catch ex As Exception
            MessageBox.Show(ex.Message)
            DOC_Save = ""
        End Try

    End Function

    Function ftp_download(ByVal o_doc As String, ByVal d_doc As String) As Boolean
        Try

            '다운로드할 ftp주소
            Dim u As New Uri(o_doc)

            'c:에 다운로드
            Dim downFile As String = d_doc

            '리퀘스트작성
            Dim ftpReq As System.Net.FtpWebRequest = CType(System.Net.WebRequest.Create(u), System.Net.FtpWebRequest)
            ftpReq.Credentials = New System.Net.NetworkCredential("admin", "dms123")
            ftpReq.Method = System.Net.WebRequestMethods.Ftp.DownloadFile
            ftpReq.KeepAlive = False
            ftpReq.UseBinary = False
            ftpReq.UsePassive = False

            Dim ftpRes As System.Net.FtpWebResponse = CType(ftpReq.GetResponse(), System.Net.FtpWebResponse)

            '화일 다운로드하기위에 스트림취득
            Dim resStrm As System.IO.Stream = ftpRes.GetResponseStream()

            '다운로드할 화일을 엽니다.
            Dim fs As System.IO.FileStream = New System.IO.FileStream(downFile, System.IO.FileMode.Create, System.IO.FileAccess.Write)
            Dim buffer(1024) As Byte

            While (True)
                Dim readSize As Integer = resStrm.Read(buffer, 0, buffer.Length)
                If (readSize = 0) Then
                    Exit While
                End If

                fs.Write(buffer, 0, readSize)
            End While

            fs.Close()
            resStrm.Close()

            ftp_download = True

        Catch ex As Exception
            MessageBox.Show(ex.Message)
            ftp_download = False
        End Try

    End Function


    Function ftp_upload(ByVal doc As String, ByVal cls As String, ByVal no1 As String, ByVal no2 As String, ByVal CLS2 As String) As Boolean
        Dim fileInf As FileInfo = New FileInfo(doc) '전송할 File을 설정

        Dim sftp As String = "ftp://59.19.211.100/dms/doc/" & cls & "/" & fileInf.Name '전송할 FTP 경로를 설정

        If InStr(ipaddr, "192.168.0.") > 0 Then
            sftp = "ftp://192.168.0.10/dms/doc/" & cls & "/" & fileInf.Name '전송할 FTP 경로를 설정
        Else
            sftp = "ftp://59.19.211.100/dms/doc/" & cls & "/" & fileInf.Name '전송할 FTP 경로를 설정
        End If


        Dim upFTP As FtpWebRequest = FtpWebRequest.Create(New Uri(sftp))
        upFTP.Credentials = New NetworkCredential("admin", "dms123") 'FTP접속 계정 설정
        upFTP.Method = WebRequestMethods.Ftp.UploadFile '작업유형 설정
        upFTP.UsePassive = False
        upFTP.KeepAlive = True
        upFTP.UseBinary = True
        upFTP.ContentLength = fileInf.Length


        Dim bBuffer(1023) As Byte '버퍼설정

        Dim iLength As Integer
        Dim fs As FileStream = fileInf.OpenRead

        Dim stm As Stream = upFTP.GetRequestStream
        iLength = fs.Read(bBuffer, 0, 1024)

        While (iLength <> 0) '전송할 파일 크기만큼 1024씩 전송
            stm.Write(bBuffer, 0, iLength)
            iLength = fs.Read(bBuffer, 0, 1024)
        End While

        stm.Close() '전송이 끝나면 닫기
        fs.Close()

        If Insert_Data("EXEC SP_FrmDocMst_SAVE '" & cls & "','" & no1 & "','" & no2 & "','" & fileInf.Name & "','" & sftp & "','" & Emp_No & "','" & CLS2 & "'") = True Then
        End If


    End Function

    Function doc_Open(ByVal f As OpenFileDialog, ByVal cls As String, ByVal no1 As String, ByVal no2 As String, ByVal CLS2 As String) As Boolean
        Try
            doc_Open = False

            f.InitialDirectory = My.Computer.FileSystem.SpecialDirectories.MyDocuments
            f.Filter = "All Files(*.*)|*.*"
            If f.ShowDialog() = Windows.Forms.DialogResult.OK Then

                If My.Computer.FileSystem.DirectoryExists("c:\temp\" & cls) = False Then
                    MkDir("c:\temp\" & cls)
                End If

                Dim f_name As String = Query_RS("SELECT CONVERT(VARCHAR(50),GETDATE(),112) + REPLACE(CONVERT(VARCHAR(50),GETDATE(),108),':','')") & "_" & f.SafeFileName
                FileCopy(f.FileName, "c:\temp\" & cls & "\" & f_name)
                ftp_upload("c:\temp\" & cls & "\" & f_name, cls, no1, no2, CLS2)
                Kill("c:\temp\" & cls & "\" & f_name)

                doc_Open = True
                MessageBox.Show("업로드 완료되었습니다.")

            End If


        Catch ex As Exception
            MessageBox.Show(ex.Message)
            doc_Open = False
        End Try
    End Function


    Function Csv_Open(ByVal f As OpenFileDialog, ByVal s As FarPoint.Win.Spread.FpSpread, ByVal t As TextBox, ByVal aa As String) As Boolean
        Try


            f.InitialDirectory = My.Computer.FileSystem.SpecialDirectories.MyDocuments
            'f.Filter = "Microsoft Office Excel File (*.xls)|*.xls*|CSV File (*.csv)|*.csv*|TEXT File (*.txt)|*.txt*|All Files(*.*)|*.*"
            f.Filter = "CSV File (*.csv)|*.csv*|All Files(*.*)|*.*"

            If f.ShowDialog() = Windows.Forms.DialogResult.OK Then
                t.Text = f.FileName

                Dim FileHolder As FileInfo = New FileInfo(t.Text)
                Dim ReadFile As StreamReader = FileHolder.OpenText()
                Dim strLine
                Dim i As Integer = 0
                s.ActiveSheet.RowCount = 0
                s.ActiveSheet.ColumnHeader.Columns.Count = 0

                Select Case aa
                    Case "FrmCompLGinvoice"

                        Do Until ReadFile.EndOfStream
                            strLine = ReadFile.ReadLine
                            If i = 1 Then
                                ParseCSV(strLine, s, True)
                            ElseIf i > 2 Then
                                'If strLine = "" Then
                                '    ReadFile.Close()
                                '    ReadFile = Nothing
                                'End If
                                s.ActiveSheet.RowCount += 1
                                ParseCSV(strLine, s, False)
                            End If
                            i += 1
                        Loop
                        ReadFile.Close()
                        ReadFile = Nothing
                End Select

            End If

            s.ActiveSheet.Rows.Remove(s.ActiveSheet.RowCount - 1, 1)
            Csv_Open = True

        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Function

    Public Sub ParseCSV(ByVal CSVstr As String, ByVal s As FarPoint.Win.Spread.FpSpread, ByVal hd_yn As Boolean)

        Dim strLen As Integer
        Dim j, k As Integer
        Dim colvalue As String
        Dim Colary As String()

        colvalue = ""
        strLen = Len(CSVstr)
        k = 0

        If strLen > 0 Then
            CSVstr = Replace(Replace(CSVstr, Chr(34) & ",", Chr(9)), Chr(34), "")
            Colary = Split(CSVstr, Chr(9))
            For j = 0 To Colary.Length - 1
                If (j > 0 And j < 6) Or (j > 13 And j < 20) Or j = 21 Then
                    If hd_yn = True Then
                        s.ActiveSheet.ColumnHeader.Columns.Count += 1
                        s.ActiveSheet.ColumnHeader.Columns(k).Label = Colary(j)
                    Else
                        If j > 14 And j < 19 Then
                            Spread_NumType(s, s.ActiveSheet.RowCount - 1, k)
                        ElseIf j = 19 Then
                            Spread_CurrencyType(s, s.ActiveSheet.RowCount - 1, k)
                        End If
                        s.ActiveSheet.SetValue(s.ActiveSheet.RowCount - 1, k, Colary(j))
                    End If
                    k += 1
                End If
            Next

        End If
    End Sub

    Sub VwHelp_Load()
        Dim Dnm, OriDnm, Fname As String

        Fname = ""

        If Site_id = "S1000" Then

            OriDnm = Query_RS("select code_name from tbl_codemaster where site_id ='" & Site_id & "' and class_id = 'R1001' and code_id = '10001'")
            Dnm = My.Application.Info.DirectoryPath & "\help" '"C:\Program Files\smsplus\ggg"
            Fname = Query_RS("select code_name from tbl_codemaster where site_id ='" & Site_id & "' and class_id = 'R1001' and code_id = '10002'")

            Make_Dir(Dnm)
            If Chk_Version(OriDnm, Dnm, Fname) = True Then
                MessageBox.Show("Manual is Updating, Wait Minutes!!", "Message")
            End If

            System.Windows.Forms.Help.ShowHelp(MainFrm, Dnm & "\" & Fname)

        Else
            Fname = Query_RS("select code_name from tbl_codemaster where site_id ='" & Site_id & "' and class_id = 'R1001' and code_id = '10002'")
            System.Windows.Forms.Help.ShowHelp(MainFrm, "HTTP://97.67.53.35/help/" & Fname)

        End If

    End Sub

    Function Make_Dir(ByVal DirNm As String) As Boolean
        Try
            Make_Dir = False
            Dim ExistTrue As Boolean = False

            ExistTrue = My.Computer.FileSystem.DirectoryExists(DirNm)

            If ExistTrue = False Then
                My.Computer.FileSystem.CreateDirectory(DirNm)
                Make_Dir = True
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error")
        End Try
    End Function

    Function Chk_Version(ByVal OriDNm As String, ByVal DirNm As String, ByVal FNm As String) As Boolean
        Try
            Chk_Version = False
            If My.Computer.FileSystem.FileExists(DirNm & "\" & FNm) = False Then
                My.Computer.FileSystem.CopyFile(OriDNm & "\" & FNm, DirNm & "\" & FNm)
                Chk_Version = True
            Else
                If My.Computer.FileSystem.GetFileInfo(DirNm & "\" & FNm).LastWriteTimeUtc < My.Computer.FileSystem.GetFileInfo(OriDNm & "\" & FNm).LastWriteTimeUtc Then
                    'MsgBox(My.Computer.FileSystem.GetFileInfo(DirNm & "\" & FNm).LastWriteTimeUtc)
                    My.Computer.FileSystem.DeleteFile(DirNm & "\" & FNm)
                    My.Computer.FileSystem.CopyFile(OriDNm & "\" & FNm, DirNm & "\" & FNm)
                    Chk_Version = True
                End If
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error")
        End Try
    End Function



    Function Img_Open(ByVal f As OpenFileDialog, ByVal s As FarPoint.Win.Spread.FpSpread, ByVal row As Integer, ByVal col As Integer, ByVal t As TextBox) As Boolean
        Try

            f.InitialDirectory = My.Computer.FileSystem.SpecialDirectories.MyDocuments
            f.Filter = "jpg File(*.jpg)|*.jpg*|bmp file(*.bmp)|*.bmp*|gif file(*.gif)|*.gif*|All Files(*.*)|*.*"
            If f.ShowDialog() = Windows.Forms.DialogResult.OK Then
                t.Text = f.FileName

                Dim aa As Drawing.Image
                With s.ActiveSheet
                    .Cells(row, col).Value = Nothing

                    aa = Image.FromFile(t.Text, True)
                    .Cells(row, col).Value = aa
                    .Cells(row, col).Note = t.Text
                End With
            End If
            Img_Open = True

        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Function

End Module
