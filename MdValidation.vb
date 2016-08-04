Module MdValidation

    Public save_excel As String

    Function ESN_VERIFY(ByVal S As TextBox, ByVal e As System.Windows.Forms.KeyEventArgs) As Boolean
        If e.KeyCode = Keys.Enter Then
            If Len(S.Text) = 15 Or Len(S.Text) = 14 Or Len(S.Text) = 10 Or Len(S.Text) = 18 Then
                ESN_VERIFY = True
            Else
                ESN_VERIFY = False
            End If
        End If
    End Function

    Function PROD_VERIFY(ByVal S As TextBox, ByVal e As System.Windows.Forms.KeyEventArgs) As Boolean
        If e.KeyCode = Keys.Enter Then
            If Len(S.Text) = 18 Then
                PROD_VERIFY = True
            Else
                PROD_VERIFY = False
            End If
        End If
    End Function


    Function FESN_VERIFY(ByVal S As TextBox, ByVal e As System.Windows.Forms.KeyEventArgs) As Boolean
        If e.KeyCode = Keys.Enter Then
            If Len(S.Text) = 10 Or Len(S.Text) = 12 Then
                FESN_VERIFY = True
            Else
                FESN_VERIFY = False
            End If
        End If
    End Function

    Function LOT_VERIFY(ByVal S As TextBox, ByVal e As System.Windows.Forms.KeyEventArgs) As Boolean
        If e.KeyCode = Keys.Enter Then
            If Len(S.Text) = 14 Then
                LOT_VERIFY = True
            Else
                LOT_VERIFY = False
            End If
        End If
    End Function

    Function DEC_TO_HEX(ByVal AA As String) As String
        Dim bb As String
        DEC_TO_HEX = ""
        bb = CStr(Hex(Mid(AA, 4, Len(AA) - 3)))
        DEC_TO_HEX = Hex(Left(AA, 3)) & bb 'CStr(Hex(Mid(Text1.Text, 4, Len(Text1.Text) - 3)))
    End Function

    Function Emp2Null(ByVal TXT As String) As String ''DB입력시 값이 없는 경우 NULL처리해서 입력(숫자)
        If TXT = "" Then
            Emp2Null = "NULL"
        Else
            Emp2Null = TXT
        End If
    End Function

    Function Str2DBNull(ByVal TXT As String) As String  'DB입력시 값이 없는 경우 NULL처리해서 입력(문자열)
        If TXT = "" Then
            Str2DBNull = "NULL"
        Else
            Str2DBNull = "'" & TXT & "'"
        End If
    End Function

    Function Formbtn_Authority(ByVal b As DevComponents.DotNetBar.ButtonX, ByVal f As String, ByVal s As String) As Boolean

        Dim au_rs As New ADODB.Recordset

        au_rs = Query_RS_ALL("select func from tbl_formautho where emp_no = '" & Emp_No & "' and form = '" & f & "' and func = '" & s & "'")

        If au_rs Is Nothing Then
            b.Enabled = False
            b.Visible = False
        Else
            If au_rs.RecordCount > 0 Then
                b.Enabled = True
                b.Visible = True
            End If
        End If

    End Function

    Function Formbtn_Authority2(ByVal b As DevComponents.DotNetBar.ButtonItem, ByVal f As String, ByVal s As String) As Boolean

        Dim au_rs As New ADODB.Recordset

        au_rs = Query_RS_ALL("select func from tbl_formautho where emp_no = '" & Emp_No & "' and form = '" & f & "' and func = '" & s & "'")

        If au_rs Is Nothing Then
            b.Enabled = False
            b.Visible = False
        Else
            If au_rs.RecordCount > 0 Then
                b.Enabled = True
                b.Visible = True
            End If
        End If

    End Function

    Function Formbim_Authority(ByVal b As DevComponents.DotNetBar.ButtonItem, ByVal f As String, ByVal s As String) As Boolean

        Dim au_rs As New ADODB.Recordset

        au_rs = Query_RS_ALL("select func from tbl_formautho where emp_no = '" & Emp_No & "' and form = '" & f & "' and func = '" & s & "'")

        If au_rs Is Nothing Then
            b.Enabled = False
            b.Visible = False
        Else
            If au_rs.RecordCount > 0 Then
                b.Enabled = True
                b.Visible = True
            End If
        End If

    End Function

    Function Menu_Authority(ByVal m As String) As Boolean

        If Query_RS("select count(menu) from tbl_menuautho where menu = '" & m & "' and emp_no = '" & Emp_No & "'") > 0 Then
            Menu_Authority = True
        Else
            Menu_Authority = False
        End If


    End Function

    Function Check_Valid_FEsn(ByVal esn As String, ByVal form As String) As Boolean
        Dim va_rs As ADODB.Recordset

        va_rs = Query_RS_ALL("EXEC SP_COMMON_CheckValidEsn '" & Site_id & "','" & esn & "', '1'")

        If va_rs Is Nothing Then
            MessageBox.Show("Not Exist ESN")
            Check_Valid_FEsn = False
            Exit Function
        End If

        If va_rs.RecordCount > 0 Then

            'If Check_FWC(form, va_rs(8).Value) = False Then
            '    Check_Valid_FEsn = False
            '    Exit Function
            'End If

            'If Check_RCV(form, va_rs(7).Value) = False Then
            '    Check_Valid_FEsn = False
            '    Exit Function
            'End If

            'If Check_TRG(form, va_rs(5).Value) = False Then
            '    Check_Valid_FEsn = False
            '    Exit Function
            'End If


        End If

        Check_Valid_FEsn = True

    End Function

    Function Check_Valid_LOT(ByVal esn As String, ByVal form As String) As Boolean
        Dim va_rs As ADODB.Recordset

        va_rs = Query_RS_ALL("EXEC SP_COMMON_CheckValidLOT '" & Site_id & "','" & esn & "'")

        If va_rs Is Nothing Then
            MessageBox.Show("Not Exist ESN")
            Check_Valid_LOT = False
            Exit Function
        End If

        If va_rs.RecordCount > 0 Then

            If Check_FWC(form, va_rs(4).Value) = False Then
                Check_Valid_LOT = False
                Exit Function
            End If

            'If Check_RCV(form, va_rs(7).Value) = False Then
            '    Check_Valid_FEsn = False
            '    Exit Function
            'End If

            'If Check_TRG(form, va_rs(5).Value) = False Then
            '    Check_Valid_FEsn = False
            '    Exit Function
            'End If


        End If

        Check_Valid_LOT = True

    End Function

    Function Check_Valid_Esn(ByVal esn As String, ByVal form As String) As Boolean

        Dim va_rs As ADODB.Recordset

        va_rs = Query_RS_ALL("EXEC SP_COMMON_CheckValidEsn '" & Site_id & "','" & esn & "', '1'")

        If va_rs Is Nothing Then
            Modal_Error(esn & vbNewLine & "Not Exist ESN")
            Check_Valid_Esn = False
            Exit Function
        End If

        If va_rs.RecordCount > 0 Then

            If Check_RCV(form, va_rs(4).Value) = False Then
                Check_Valid_Esn = False
                Exit Function
            End If

            'If form = "FrmPacking" Or form = "FrmFinalPass" Then

            '    If va_rs(3).Value = "N" Then
            '        If Check_CAL(form, va_rs(6).Value, va_rs(7).Value) = False Then
            '            Check_Valid_Esn = False
            '            Exit Function
            '        End If

            '        If Check_INFO(form, va_rs(6).Value, va_rs(9).Value) = False Then
            '            Check_Valid_Esn = False
            '            Exit Function
            '        End If

            '    End If
            'End If

            'If form = "FrmQCPass" Then

            '    If va_rs(3).Value = "N" Then
            '        If Check_CAL(form, va_rs(6).Value, va_rs(7).Value) = False Then
            '            Check_Valid_Esn = False
            '            Exit Function
            '        End If
            '    End If
            'End If

            'If form <> "FrmPD" And form <> "FrmPacking" Then
            '    If Check_Good(form, va_rs(6).Value) = False Then
            '        Check_Valid_Esn = False
            '        Exit Function
            '    End If
            'End If

            If Check_WC(form, va_rs(1).Value) = False Then
                Check_Valid_Esn = False
                Exit Function
            End If

            If va_rs(1).Value = "W9800" Then
                If Emp_No = "10032" Or Emp_No = "11111" Or Emp_No = "10074" Then
                Else
                    Check_Valid_Esn = False
                    MessageBox.Show("Quarantine!")
                    Exit Function

                End If

            End If

            Check_Valid_Esn = True
        End If

    End Function
    Function Check_FWC(ByVal form As String, ByVal wc As String) As Boolean

        If form = "FrmReceiving" Then
            If wc <> "K0000" Then
                Modal_Error("공정 오류 :  : " & Query_RS("SELECT CODE_NAME FROM TBL_CODEMASTER WHERE SITE_ID = '" & Site_id & "' AND CLASS_ID = 'R0001' AND CODE_ID = '" & wc & "'"))
                Check_FWC = False
                Exit Function
            End If
        ElseIf form = "FrmTriage" Then
            If wc <> "K1000" Then
                Modal_Error("공정 오류 :  : " & Query_RS("SELECT CODE_NAME FROM TBL_CODEMASTER WHERE SITE_ID = '" & Site_id & "' AND CLASS_ID = 'R0001' AND CODE_ID = '" & wc & "'"))
                Check_FWC = False
                Exit Function
            End If
        ElseIf form = "FrmLevel1" Then
            If wc <> "K2000" Then
                Modal_Error("공정 오류 :  : " & Query_RS("SELECT CODE_NAME FROM TBL_CODEMASTER WHERE SITE_ID = '" & Site_id & "' AND CLASS_ID = 'R0001' AND CODE_ID = '" & wc & "'"))
                Check_FWC = False
                Exit Function
            End If
        ElseIf form = "FrmLinePass" Or form = "FrmQCLine" Then
            If wc <> "K3900" Then
                Modal_Error("공정 오류 :  : " & Query_RS("SELECT CODE_NAME FROM TBL_CODEMASTER WHERE SITE_ID = '" & Site_id & "' AND CLASS_ID = 'R0001' AND CODE_ID = '" & wc & "'"))
                Check_FWC = False
                Exit Function
            End If
        ElseIf form = "FrmLevel2" Then
            If wc <> "K4000" Then
                Modal_Error("공정 오류 :  : " & Query_RS("SELECT CODE_NAME FROM TBL_CODEMASTER WHERE SITE_ID = '" & Site_id & "' AND CLASS_ID = 'R0001' AND CODE_ID = '" & wc & "'"))
                Check_FWC = False
                Exit Function
            End If
        ElseIf form = "FrmQCPass" Or form = "FrmQCFail" Then
            If wc <> "K4900" Then
                Modal_Error("공정 오류 :  : " & Query_RS("SELECT CODE_NAME FROM TBL_CODEMASTER WHERE SITE_ID = '" & Site_id & "' AND CLASS_ID = 'R0001' AND CODE_ID = '" & wc & "'"))
                Check_FWC = False
                Exit Function
            End If
        ElseIf form = "FrmFSHIPPING" And form = "FrmFPacking" Then
            If wc <> "K9000" Then
                Modal_Error("공정 오류 :  : " & Query_RS("SELECT CODE_NAME FROM TBL_CODEMASTER WHERE SITE_ID = '" & Site_id & "' AND CLASS_ID = 'R0001' AND CODE_ID = '" & wc & "'"))
                Check_FWC = False
                Exit Function
            End If
        ElseIf form = "FrmFinalPass" Or form = "FrmFinalFail" Then
            If wc <> "K6000" Then
                Modal_Error("공정 오류 :  : " & Query_RS("SELECT CODE_NAME FROM TBL_CODEMASTER WHERE SITE_ID = '" & Site_id & "' AND CLASS_ID = 'R0001' AND CODE_ID = '" & wc & "'"))
                Check_FWC = False
                Exit Function
            End If
        ElseIf form = "FrmAssy" Then
            If wc <> "K5000" Then
                Modal_Error("공정 오류 :  : " & Query_RS("SELECT CODE_NAME FROM TBL_CODEMASTER WHERE SITE_ID = '" & Site_id & "' AND CLASS_ID = 'R0001' AND CODE_ID = '" & wc & "'"))
                Check_FWC = False
                Exit Function
            End If
        ElseIf form = "FrmPD" Then
            If wc <> "K8300" Then
                Modal_Error("공정 오류 :  : " & Query_RS("SELECT CODE_NAME FROM TBL_CODEMASTER WHERE SITE_ID = '" & Site_id & "' AND CLASS_ID = 'R0001' AND CODE_ID = '" & wc & "'"))
                Check_FWC = False
                Exit Function
            End If
        ElseIf form = "FrmPolRtn" Then
        End If

        Check_FWC = True


    End Function

    Function Check_WC(ByVal form As String, ByVal wc As String) As Boolean

        If form = "FrmTechnician" Then
            If wc <> "W3000" And wc <> "W3100" Then
                Modal_Error("공정 오류 :  : " & Query_RS("SELECT CODE_NAME FROM TBL_CODEMASTER WHERE SITE_ID = '" & Site_id & "' AND CLASS_ID = 'R0001' AND CODE_ID = '" & wc & "'"))
                Check_WC = False
                Exit Function
            End If
        ElseIf form = "FrmBulkInput" Then
            If wc <> "W3000" And wc <> "W3100" Then
                Modal_Error("공정 오류 :  : " & Query_RS("SELECT CODE_NAME FROM TBL_CODEMASTER WHERE SITE_ID = '" & Site_id & "' AND CLASS_ID = 'R0001' AND CODE_ID = '" & wc & "'"))
                Check_WC = False
                Exit Function
            End If
        ElseIf form = "FrmLevel1" Then
            If wc <> "W4000" Then
                Modal_Error("공정 오류 :  : " & Query_RS("SELECT CODE_NAME FROM TBL_CODEMASTER WHERE SITE_ID = '" & Site_id & "' AND CLASS_ID = 'R0001' AND CODE_ID = '" & wc & "'"))
                Check_WC = False
                Exit Function
            End If
        ElseIf form = "FrmReceiving" Then
            If wc <> "W0000" Then
                Modal_Error("공정 오류 :  : " & Query_RS("SELECT CODE_NAME FROM TBL_CODEMASTER WHERE SITE_ID = '" & Site_id & "' AND CLASS_ID = 'R0001' AND CODE_ID = '" & wc & "'"))
                Check_WC = False
                Exit Function
            End If
        ElseIf form = "FrmTriage" Then
            If wc <> "W1000" Then
                Modal_Error("공정 오류 :  : " & Query_RS("SELECT CODE_NAME FROM TBL_CODEMASTER WHERE SITE_ID = '" & Site_id & "' AND CLASS_ID = 'R0001' AND CODE_ID = '" & wc & "'"))
                Check_WC = False
                Exit Function
            End If
        ElseIf form = "FrmPD" Then
            If wc <> "W5000" Then
                Modal_Error("공정 오류 :  : " & Query_RS("SELECT CODE_NAME FROM TBL_CODEMASTER WHERE SITE_ID = '" & Site_id & "' AND CLASS_ID = 'R0001' AND CODE_ID = '" & wc & "'"))
                Check_WC = False
                Exit Function
            End If
        ElseIf form = "FrmPacking" Then
            If wc <> "W9000" Then
                Modal_Error("공정 오류 :  : " & Query_RS("SELECT CODE_NAME FROM TBL_CODEMASTER WHERE SITE_ID = '" & Site_id & "' AND CLASS_ID = 'R0001' AND CODE_ID = '" & wc & "'"))
                Check_WC = False
                Exit Function
            End If
        ElseIf form = "FrmQCPass" Then
            If wc <> "W6000" Then
                Modal_Error("공정 오류 :  : " & Query_RS("SELECT CODE_NAME FROM TBL_CODEMASTER WHERE SITE_ID = '" & Site_id & "' AND CLASS_ID = 'R0001' AND CODE_ID = '" & wc & "'"))
                Check_WC = False
                Exit Function
            End If
        ElseIf form = "FrmQCFail" Then
            If wc <> "W6000" Then
                Modal_Error("공정 오류 :  : " & Query_RS("SELECT CODE_NAME FROM TBL_CODEMASTER WHERE SITE_ID = '" & Site_id & "' AND CLASS_ID = 'R0001' AND CODE_ID = '" & wc & "'"))
                Check_WC = False
                Exit Function
            End If
        ElseIf form = "FrmQCLine" Then
            If wc <> "W3500" Then
                Modal_Error("공정 오류 :  : " & Query_RS("SELECT CODE_NAME FROM TBL_CODEMASTER WHERE SITE_ID = '" & Site_id & "' AND CLASS_ID = 'R0001' AND CODE_ID = '" & wc & "'"))
                Check_WC = False
                Exit Function
            End If
        ElseIf form = "FrmLinePass" Then
            If wc <> "W3500" Then
                Modal_Error("공정 오류 :  : " & Query_RS("SELECT CODE_NAME FROM TBL_CODEMASTER WHERE SITE_ID = '" & Site_id & "' AND CLASS_ID = 'R0001' AND CODE_ID = '" & wc & "'"))
                Check_WC = False
                Exit Function
            End If
        ElseIf form = "FrmFinalPass" Then
            If wc <> "W8500" Then
                Modal_Error("공정 오류 :  : " & Query_RS("SELECT CODE_NAME FROM TBL_CODEMASTER WHERE SITE_ID = '" & Site_id & "' AND CLASS_ID = 'R0001' AND CODE_ID = '" & wc & "'"))
                Check_WC = False
                Exit Function
            End If
        ElseIf form = "FrmFinalFail" Then
            If wc <> "W8500" Then
                Modal_Error("공정 오류 :  : " & Query_RS("SELECT CODE_NAME FROM TBL_CODEMASTER WHERE SITE_ID = '" & Site_id & "' AND CLASS_ID = 'R0001' AND CODE_ID = '" & wc & "'"))
                Check_WC = False
                Exit Function
            End If
        ElseIf form = "FrmPFail" Then
            If wc <> "W9000" Then
                Modal_Error("공정 오류 :  : " & Query_RS("SELECT CODE_NAME FROM TBL_CODEMASTER WHERE SITE_ID = '" & Site_id & "' AND CLASS_ID = 'R0001' AND CODE_ID = '" & wc & "'"))
                Check_WC = False
                Exit Function
            End If
        ElseIf form = "FrmCosmetic" Then
            If wc <> "W4000" Then
                Modal_Error("공정 오류 :  : " & Query_RS("SELECT CODE_NAME FROM TBL_CODEMASTER WHERE SITE_ID = '" & Site_id & "' AND CLASS_ID = 'R0002' AND CODE_ID = '" & wc & "'"))
                Check_WC = False
                Exit Function
            End If
        ElseIf form = "FrmDisAssy" Then
            If wc <> "W1500" Then
                Modal_Error("공정 오류 :  : " & Query_RS("SELECT CODE_NAME FROM TBL_CODEMASTER WHERE SITE_ID = '" & Site_id & "' AND CLASS_ID = 'R0002' AND CODE_ID = '" & wc & "'"))
                Check_WC = False
                Exit Function
            End If
        ElseIf form = "FrmMapping" Or form = "FrmAssy" Then
            If wc <> "W4500" Then
                Modal_Error("공정 오류 :  : " & Query_RS("SELECT CODE_NAME FROM TBL_CODEMASTER WHERE SITE_ID = '" & Site_id & "' AND CLASS_ID = 'R0002' AND CODE_ID = '" & wc & "'"))
                Check_WC = False
                Exit Function
            End If
        ElseIf form = "FrmRFFail" Or form = "FrmRFPass" Then
            If wc <> "W2000" Then
                Modal_Error("공정 오류 :  : " & Query_RS("SELECT CODE_NAME FROM TBL_CODEMASTER WHERE SITE_ID = '" & Site_id & "' AND CLASS_ID = 'R0002' AND CODE_ID = '" & wc & "'"))
                Check_WC = False
                Exit Function
            End If
        ElseIf form = "FrmDLFail" Or form = "FrmDLPass" Then
            If wc <> "W2200" Then
                Modal_Error("공정 오류 :  : " & Query_RS("SELECT CODE_NAME FROM TBL_CODEMASTER WHERE SITE_ID = '" & Site_id & "' AND CLASS_ID = 'R0002' AND CODE_ID = '" & wc & "'"))
                Check_WC = False
                Exit Function
            End If
        ElseIf form = "FrmRS" Or form = "FrmBRS" Then
            If wc <> "W8000" Then
                Modal_Error("공정 오류 :  : " & Query_RS("SELECT CODE_NAME FROM TBL_CODEMASTER WHERE SITE_ID = '" & Site_id & "' AND CLASS_ID = 'R0002' AND CODE_ID = '" & wc & "'"))
                Check_WC = False
                Exit Function
            End If
        ElseIf form = "FrmLabel" Then
            If wc <> "W8300" Then
                Modal_Error("공정 오류 :  : " & Query_RS("SELECT CODE_NAME FROM TBL_CODEMASTER WHERE SITE_ID = '" & Site_id & "' AND CLASS_ID = 'R0002' AND CODE_ID = '" & wc & "'"))
                Check_WC = False
                Exit Function
            End If
        End If

        Check_WC = True

    End Function

    Function Check_RCV(ByVal form As String, ByVal wc As String) As Boolean

        If wc = "N" Then
            Modal_Error("No Receiving")
            Check_RCV = False
            Exit Function
        End If

        Check_RCV = True

    End Function

    Function Check_TRG(ByVal form As String, ByVal wc As String) As Boolean

        If wc = "N" Then
            Modal_Error("No Triage")
            Check_TRG = False
            Exit Function
        End If

        Check_TRG = True

    End Function

    Function Check_Good(ByVal form As String, ByVal wc As String) As Boolean

        If wc <> "GOOD" Then
            Modal_Error("No GOOD JUDGE : " & wc)
            Check_Good = False
            Exit Function
        End If

        Check_Good = True

    End Function

    Function Check_QcRej(ByVal form As String, ByVal wc As String) As Boolean

        If wc <> "GOOD" Then
            Modal_Error("QAREJECT")
            Check_QcRej = False
            Exit Function
        End If

        Check_QcRej = True

    End Function
    Function Check_BASIC(ByVal form As String, ByVal JUDGE As String, ByVal P As String) As Boolean

        If JUDGE = "GOOD" And P = "N" Then
            Modal_Error("시험(BASIC) 결과가 없습니다.")
            Check_BASIC = False
            Exit Function
        End If

        Check_BASIC = True

    End Function
    Function Check_NW(ByVal form As String, ByVal JUDGE As String, ByVal P As String) As Boolean

        If JUDGE = "GOOD" And P = "N" Then
            Modal_Error("시험(N/W) 결과가 없습니다.")
            Check_NW = False
            Exit Function
        End If

        Check_NW = True

    End Function


    Function Check_CAL(ByVal form As String, ByVal JUDGE As String, ByVal P As String) As Boolean

        If JUDGE = "GOOD" And P = "N" Then
            Modal_Error("NO CAL PASS.")
            Check_CAL = False
            Exit Function
        End If

        Check_CAL = True

    End Function

    Function Check_DM(ByVal form As String, ByVal JUDGE As String, ByVal P As String) As Boolean

        If JUDGE = "GOOD" And P = "N" Then
            Modal_Error("NO DL PASS.")
            Check_DM = False
            Exit Function
        End If

        Check_DM = True

    End Function

    Function Check_INFO(ByVal form As String, ByVal JUDGE As String, ByVal P As String) As Boolean

        If JUDGE = "GOOD" And P = "N" Then
            Modal_Error("NO INFO PASS.")
            Check_INFO = False
            Exit Function
        End If

        Check_INFO = True

    End Function

    Function Modal_Error(ByVal errmsg As String) As Boolean
        FrmError.TextBoxX1.Text = errmsg
        FrmError.ShowDialog()

    End Function

    Function Conv_Mon(ByVal intMon As String) As String
        Conv_Mon = ""
        Select Case CInt(intMon)
            Case 1
                Conv_Mon = "January"
            Case 2
                Conv_Mon = "February"
            Case 3
                Conv_Mon = "March"
            Case 4
                Conv_Mon = "April"
            Case 5
                Conv_Mon = "May"
            Case 6
                Conv_Mon = "June"
            Case 7
                Conv_Mon = "July"
            Case 8
                Conv_Mon = "August"
            Case 9
                Conv_Mon = "September"
            Case 10
                Conv_Mon = "October"
            Case 11
                Conv_Mon = "November"
            Case 12
                Conv_Mon = "December"
        End Select
    End Function

    Function Month_Day(ByVal InputDate As String) As String()  '월의 말일 구하기
        Dim Fy, Fy2 As String
        Dim Fm As String
        Dim Nextm, lastday, firstday As String

        Fy = Year(CDate(InputDate))
        Fm = Month(CDate(InputDate))

        If CInt(Fm) = 12 Then
            Nextm = "01"
            Fy2 = CInt(Fy) + 1
        Else
            Nextm = CStr(CInt(Fm) + 1)
            Fy2 = Fy
        End If

        firstday = CDate(Fy & "-" & Fm & "-01")
        lastday = DateAdd(DateInterval.Day, -1, CDate(Fy2 & "-" & Nextm & "-01"))


        Month_Day = New String() {firstday, lastday}



    End Function

    Function Month_Week(ByVal InputDate As String) As Integer()
        Dim fweek, lweek As Integer

        fweek = DatePart(DateInterval.WeekOfYear, CDate(Month_Day(InputDate)(0)), FirstDayOfWeek.Monday)
        lweek = DatePart(DateInterval.WeekOfYear, CDate(Month_Day(InputDate)(1)), FirstDayOfWeek.Monday)

        Month_Week = New Integer() {fweek, lweek}
    End Function

    Function Week_Day(ByVal InputDate As String, ByVal Weekidx As Integer, ByVal Weektot As Integer) As String()
        Dim weekcnt As Integer
        Dim fMday, eMday, fWDay, eWDay As Date
        fMday = Month_Day(InputDate)(0)
        eMday = Month_Day(InputDate)(1)
        weekcnt = DatePart(DateInterval.Weekday, CDate(fMday), FirstDayOfWeek.Monday)

        fWDay = DateAdd(DateInterval.Day, 1 - weekcnt, CDate(fMday))
        eWDay = DateAdd(DateInterval.Day, 7 - weekcnt, CDate(fMday))
        If Weekidx = 1 Then
            Week_Day = New String() {fMday, eWDay}
        ElseIf Weektot = Weekidx Then
            Week_Day = New String() {DateAdd(DateInterval.Day, 7 * (Weekidx - 1), fWDay), eMday}
        Else
            Week_Day = New String() {DateAdd(DateInterval.Day, 7 * (Weekidx - 1), fWDay), DateAdd(DateInterval.Day, 7 * (Weekidx - 1), eWDay)}
        End If
    End Function

    Function Week_FEday(ByVal InputDate As String) As String()  '주의 시작일(월) 과 마지막일(일) 구하기
        Dim weekcnt As Integer
        Dim fWDay, eWDay As Date

        weekcnt = DatePart(DateInterval.Weekday, CDate(InputDate), FirstDayOfWeek.Monday)

        fWDay = DateAdd(DateInterval.Day, 1 - weekcnt, CDate(InputDate))
        eWDay = DateAdd(DateInterval.Day, 7 - weekcnt, CDate(InputDate))
        Week_FEday = New String() {fWDay, eWDay}
    End Function

    Function Check_PartLoc(ByVal WH As String, ByVal LOC As String) As Boolean
        Dim wwh_cd, wloc_cd As String

        wwh_cd = Query_RS("Select code_id from tbl_codemaster where site_id = '" & Site_id & "' and Class_id = 'R0073' and Code_name = '" & WH & "'")
        wloc_cd = Query_RS("Select isnull(code_id,'N/A') from tbl_codemaster where site_id = '" & Site_id & "' and Class_id = 'R0074' and Code_name = '" & LOC & "'")

        If LOC = "" Then
            wloc_cd = wwh_cd
        End If

        If wwh_cd = Microsoft.VisualBasic.Left(wloc_cd, 4) Then
            Check_PartLoc = True
        Else
            Check_PartLoc = False
        End If

    End Function


    Function DUP_COLOR(ByVal S As FarPoint.Win.Spread.FpSpread, ByVal POS As Integer, ByVal S1 As ListView, ByVal POS1 As Integer) As Boolean

        Dim I, J As Integer

        If S.ActiveSheet.RowCount > 0 Then
            For I = 0 To S.ActiveSheet.RowCount - 1
                For J = 0 To S1.items.Count - 1
                    '                MessageBox.Show(S.ActiveSheet.GetValue(I, POS) & ":" & S1.Items(J).Text)

                    If S.ActiveSheet.GetValue(I, POS) = S1.Items(J).Text Then
                        If S.ActiveSheet.GetValue(I, POS1) > 0 Then
                            S1.Items(J).BackColor = Color.OrangeRed
                        End If
                    End If
                Next
            Next
        Else

        End If

        DUP_COLOR = True
    End Function

    Function Chk_Pack(ByVal esn As String) As Boolean

        If Query_RS("select isnull(inboxid,'') from tbl_esnmaster where site_id = '" & Site_id & "' and esn = '" & esn & "'") <> "" Then
            Chk_Pack = False
            Modal_Error("Already Packed ESN!!!")
        Else
            Chk_Pack = True
        End If

    End Function

    Function DEC_TO_HEX2(ByVal AA As String) As String
        Dim bb As String
        DEC_TO_HEX2 = ""
        bb = CStr(Hex(Mid(AA, 11, 8)))
        '        cc = Strings(6 - Len(bb), "0") & bb
        If Len(bb) = 5 Then
            bb = "0" & bb
        ElseIf (Len(bb) = 4) Then
            bb = "00" & bb
        End If

        'If Mid(AA, 11, 3) = "000" Then
        '    DEC_TO_HEX2 = Hex(Left(AA, 10)) & bb 'CStr(Hex(Mid(Text1.Text, 4, Len(Text1.Text) - 3)))
        'Else
        DEC_TO_HEX2 = Hex(Left(AA, 10)) & Microsoft.VisualBasic.Right("0000000" & bb, 6) 'CStr(Hex(Mid(Text1.Text, 4, Len(Text1.Text) - 3)))
        'End If

        'Else
        '    DEC_TO_HEX = CLng("&H" & Left(AA, 2))
        'DEC_TO_HEX = String(3 - Len(bb), "0") & bb

        '    cc = CLng("&H" & Mid(AA, 3, Len(AA) - 2))
        'cc = String(8 - Len(cc), "0") & cc
        '    Text1.Text = bb & cc 'CLng("&H" & Left(Text2.Text, 2)) & CLng("&H" & Mid(Text2.Text, 3, Len(Text2.Text) - 2))
        'End If

    End Function

    Function HEX_TO_DEC2(ByVal AA As String) As String
        Dim bb, str As String
        Dim i As Integer
        Dim cc, SUM As Long

        '  Dim CC1, SUM1 As Long

        HEX_TO_DEC2 = ""

        For i = 1 To 8
            str = Mid(AA, i, 1)
            Select Case str
                Case "0" : cc = 0
                Case "1" : cc = 1
                Case "2" : cc = 2
                Case "3" : cc = 3
                Case "4" : cc = 4
                Case "5" : cc = 5
                Case "6" : cc = 6
                Case "7" : cc = 7
                Case "8" : cc = 8
                Case "9" : cc = 9
                Case "A" : cc = 10
                Case "B" : cc = 11
                Case "C" : cc = 12
                Case "D" : cc = 13
                Case "E" : cc = 14
                Case "F" : cc = 15
            End Select
            cc = cc * (16 ^ (8 - i))
            SUM = SUM + cc
        Next




        bb = Microsoft.VisualBasic.Right("000" & SUM, 10) & Microsoft.VisualBasic.Right("00000000" & Convert.ToInt32(Mid(AA, 9, 6), 16), 8)
        ' bb = Microsoft.VisualBasic.Right("000" & SUM, 10) & Microsoft.VisualBasic.Right("000" & Val((Mid(AA, 9, 6))), 8)



        'If Mid(AA, 9, 2) = "01" Then
        '    bb = SUM & "000" & Val("&H" & (Mid(AA, 9, 6)))
        '    HEX_TO_DEC2 = bb
        '    Exit Function
        'ElseIf Mid(AA, 9, 1) = "0" Then
        '    bb = SUM & "00" & Val("&H" & (Mid(AA, 9, 6)))

        'Else
        '    bb = SUM & "0" & Val("&H" & (Mid(AA, 9, 6)))
        'End If

        'Val()

        HEX_TO_DEC2 = bb

    End Function

End Module
