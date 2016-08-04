Module MdBarcode

    Function Print_BBarcode(ByVal esn As String, ByVal model As String, ByVal def As String) As Boolean

        Dim FN

        FN = Microsoft.VisualBasic.FileSystem.FreeFile
        Microsoft.VisualBasic.FileOpen(FN, "C:\ESN.ZPL", OpenMode.Output)
        '        Microsoft.VisualBasic.FileSystem.Write(FN, "/ ^XA" & vbNewLine & "/ ^LT0" & vbNewLine & "/ ^LH40,0" & vbNewLine & "/ ^FO1,100^A0N,13,25^FD" & model & "-" & def & "^FS" & vbNewLine & "/ ^FO1,50^BY1" & vbNewLine & "/ ^B3N,N,30,Y,N" & vbNewLine & "/ ^FD" & esn & "^FS" & vbNewLine & "/ ^XZ")

        If Site_id = "S1000" Then
            Microsoft.VisualBasic.FileSystem.Write(FN, "/ ^XA" & vbNewLine & "/ ^LT0" & vbNewLine & "/ ^LH40,0" & vbNewLine & "/ ^FO1,100^A0N,13,25^FD" & model & "-" & def & "^FS" & vbNewLine & "/ ^FO1,50^BY1" & vbNewLine & "/ ^B3N,N,30,Y,N" & vbNewLine & "/ ^FD" & esn & "^FS" & vbNewLine & "/ ^XZ")
            Microsoft.VisualBasic.FileClose()
            Shell("c:\windows\system32\cmd.exe /c print c:\esn.zpl /d:com1")
        Else
            '            Microsoft.VisualBasic.FileSystem.Write(FN, "/ ^XA" & vbNewLine & "/ ^LT-40" & vbNewLine & "/ ^LH10" & vbNewLine & "/ ^FO1,30^A0N,10,20^FD" & model & "^FS" & vbNewLine & "/ ^FO1,50^BY1,2,50" & vbNewLine & "/ ^B3N,N,40,Y,Y" & vbNewLine & "/ ^FD" & esn & "^FS" & vbNewLine & "/ ^XZ")
            Microsoft.VisualBasic.FileSystem.Write(FN, "/ ^XA" & vbNewLine & "/ ^LT-40" & vbNewLine & "/ ^LH10" & vbNewLine & "/ ^FO1,30^A0N,10,20^FD" & model & "-" & def & "^FS" & vbNewLine & "/ ^FO1,50^BY1,2,50" & vbNewLine & "/ ^B3N,N,40,Y,N" & vbNewLine & "/ ^FD" & esn & "^FS" & vbNewLine & "/ ^XZ")
            Microsoft.VisualBasic.FileClose()
            Shell("c:\windows\system32\cmd.exe /c print c:\esn.zpl /d:LPT1", AppWinStyle.Hide)
        End If
        'Microsoft.VisualBasic.FileClose()
        ''Shell("c:\windows\system32\cmd.exe /c print c:\esn.zpl /d:com1")
        'Shell("c:\windows\system32\cmd.exe /c print c:\esn.zpl /d:lpt1", AppWinStyle.Hide)

    End Function

    Function Print_BBarcode_rcv(ByVal esn As String, ByVal model As String, ByVal def As String, ByVal rcv As String) As Boolean

        Dim FN

        FN = Microsoft.VisualBasic.FileSystem.FreeFile
        Microsoft.VisualBasic.FileOpen(FN, "C:\ESN.ZPL", OpenMode.Output)

        If Site_id = "S1000" Then
            Microsoft.VisualBasic.FileSystem.Write(FN, "/ ^XA" & vbNewLine & "/ ^LT0" & vbNewLine & "/ ^LH40,0" & vbNewLine & "/ ^FO1,100^A0N,13,25^FD" & model & "-" & def & "-" & rcv & "^FS" & vbNewLine & "/ ^FO1,50^BY1" & vbNewLine & "/ ^B3N,N,30,Y,N" & vbNewLine & "/ ^FD" & esn & "^FS" & vbNewLine & "/ ^XZ")
            Microsoft.VisualBasic.FileClose()
            Shell("c:\windows\system32\cmd.exe /c print c:\esn.zpl /d:com1")
        End If

    End Function

    Function Print_BOx_Barcode1(ByVal LOT As String, ByVal cusnm As String, ByVal lcusnm As String, ByVal QTY As Integer, ByVal Rt As RichTextBox) As Boolean
        '미국 스카티용
        'Dim FN
        Dim bcode As String

        Rt.Text = ""

        'FN = Microsoft.VisualBasic.FileSystem.FreeFile
        'Microsoft.VisualBasic.FileOpen(FN, "C:\TEMP\LOT.PRN", OpenMode.Output)
        ''        Microsoft.VisualBasic.FileSystem.Write(FN, "/ ^XA" & vbNewLine & "/ ^LT0" & vbNewLine & "/ ^LH40,0" & vbNewLine & "/ ^FO1,100^A0N,13,25^FD" & model & "-" & def & "^FS" & vbNewLine & "/ ^FO1,50^BY1" & vbNewLine & "/ ^B3N,N,30,Y,N" & vbNewLine & "/ ^FD" & esn & "^FS" & vbNewLine & "/ ^XZ")

        If Site_id = "S1000" Then

            bcode = "<xpml><page quantity='0' pitch='50.0 mm'></xpml>SIZE 97.5 mm, 50 mm" & vbNewLine
            bcode += "DIRECTION 0,0" & vbNewLine
            bcode += "REFERENCE 0,0" & vbNewLine
            bcode += "OFFSET 0 mm" & vbNewLine
            bcode += "SET PEEL OFF" & vbNewLine
            bcode += "SET CUTTER OFF" & vbNewLine
            bcode += "SET PARTIAL_CUTTER OFF" & vbNewLine
            bcode += "<xpml></page></xpml><xpml><page quantity='1' pitch='50.0 mm'></xpml>SET TEAR ON" & vbNewLine
            bcode += "CLS" & vbNewLine
            bcode += "CODEPAGE 1254" & vbNewLine
            bcode += "TEXT 712,379,""0"",180,10,10,""SCOTTII KOREA""" & vbNewLine
            bcode += "BOX 56,345,725,388,3" & vbNewLine
            bcode += "TEXT 697,78,""0"",180,22,22,""QTY""" & vbNewLine
            bcode += "BOX 57,260,550,347,3" & vbNewLine
            bcode += "BARCODE 484,339,""128M"",42,0,180,2,4,""!104PK!099" & Mid(LOT, 3, 8) & "!100-!099" & Right(LOT, 4) & """" & vbNewLine
            bcode += "TEXT 431,291,""0"",180,11,11,""" & LOT & """" & vbNewLine
            bcode += "BOX 548,260,725,347,3" & vbNewLine
            bcode += "TEXT 690,330,""0"",180,22,22,""LOT""" & vbNewLine
            bcode += "BOX 548,90,725,177,3" & vbNewLine
            bcode += "BOX 57,90,550,177,3" & vbNewLine
            bcode += "BOX 548,4,725,91,3" & vbNewLine
            bcode += "BOX 57,4,550,91,3" & vbNewLine
            bcode += "TEXT 354,74,""0"",180,20,20,""" & QTY & """" & vbNewLine
            bcode += "BOX 57,175,550,262,3" & vbNewLine
            bcode += "BOX 548,175,725,262,3" & vbNewLine
            bcode += "TEXT 701,240,""0"",180,18,18,""고객명""" & vbNewLine
            bcode += "TEXT 537,232,""0"",180,14,14,""" & cusnm & """" & vbNewLine
            bcode += "TEXT 710,154,""0"",180,18,18,""출하처""" & vbNewLine
            bcode += "TEXT 537,148,""0"",180,14,14,""" & lcusnm & """" & vbNewLine
            bcode += "PRINT 1,1" & vbNewLine
            bcode += "<xpml></page></xpml><xpml><End/></xpml>" & vbNewLine


            Rt.Text = bcode

            'Microsoft.VisualBasic.FileSystem.Write(FN, bcode)
            'Microsoft.VisualBasic.FileClose()

            Rt.SaveFile("C:\TEMP\LOT.PRN", RichTextBoxStreamType.PlainText)

            Shell("c:\windows\system32\cmd.exe /c COPY c:\TEMP\LOT.PRN \\" & My.Computer.Name & "\BP04")

        End If

    End Function

    Function Print_BOx_Barcode2(ByVal LOT As String, ByVal cusno As String, ByVal model As String, ByVal modelnm As String, ByVal QTY As Integer, ByVal idx As Integer, ByVal Rt As RichTextBox) As Boolean
        '미국 스카티용
        'Dim FN
        Dim bcode As String

        Rt.Text = ""

        'FN = Microsoft.VisualBasic.FileSystem.FreeFile
        'Microsoft.VisualBasic.FileOpen(FN, "C:\TEMP\LOT.PRN", OpenMode.Output)
        ''        Microsoft.VisualBasic.FileSystem.Write(FN, "/ ^XA" & vbNewLine & "/ ^LT0" & vbNewLine & "/ ^LH40,0" & vbNewLine & "/ ^FO1,100^A0N,13,25^FD" & model & "-" & def & "^FS" & vbNewLine & "/ ^FO1,50^BY1" & vbNewLine & "/ ^B3N,N,30,Y,N" & vbNewLine & "/ ^FD" & esn & "^FS" & vbNewLine & "/ ^XZ")

        If Site_id = "S1000" Then


            bcode = "<xpml><page quantity='0' pitch='50.0 mm'></xpml>SIZE 97.5 mm, 50 mm" & vbNewLine
            bcode += "DIRECTION 0,0" & vbNewLine
            bcode += "REFERENCE 0,0" & vbNewLine
            bcode += "OFFSET 0 mm" & vbNewLine
            bcode += "SET PEEL OFF" & vbNewLine
            bcode += "SET CUTTER OFF" & vbNewLine
            bcode += "SET PARTIAL_CUTTER OFF" & vbNewLine
            bcode += "<xpml></page></xpml><xpml><page quantity='1' pitch='50.0 mm'></xpml>SET TEAR ON" & vbNewLine
            bcode += "CLS" & vbNewLine
            'bcode += "BARCODE 415,78,""128M"",33,0,180,3,6,""!10510!1000""" & vbNewLine
            bcode += "CODEPAGE 1254" & vbNewLine
            bcode += "BARCODE 455,163,""UPCA"",36,0,180,3,6,""" & cusno & """" & vbNewLine
            bcode += "TEXT 421,121,""0"",180,10,10,""" & Mid(cusno, 2, 1) & """" & vbNewLine
            bcode += "TEXT 401,121,""0"",180,10,10,""" & Mid(cusno, 3, 1) & """" & vbNewLine
            bcode += "TEXT 379,121,""0"",180,10,10,""" & Mid(cusno, 4, 1) & """" & vbNewLine
            bcode += "TEXT 356,121,""0"",180,10,10,""" & Mid(cusno, 5, 1) & """" & vbNewLine
            bcode += "TEXT 336,121,""0"",180,10,10,""" & Mid(cusno, 6, 1) & """" & vbNewLine
            bcode += "TEXT 306,121,""0"",180,10,10,""" & Mid(cusno, 7, 1) & """" & vbNewLine
            bcode += "TEXT 285,121,""0"",180,10,10,""" & Mid(cusno, 8, 1) & """" & vbNewLine
            bcode += "TEXT 263,121,""0"",180,10,10,""" & Mid(cusno, 9, 1) & """" & vbNewLine
            bcode += "TEXT 242,121,""0"",180,10,10,""" & Mid(cusno, 10, 1) & """" & vbNewLine
            bcode += "TEXT 220,121,""0"",180,10,10,""" & Mid(cusno, 11, 1) & """" & vbNewLine
            bcode += "TEXT 482,116,""0"",180,7,7,""" & Mid(cusno, 1, 1) & """" & vbNewLine
            bcode += "TEXT 153,116,""0"",180,7,7,""" & Mid(cusno, 12, 1) & """" & vbNewLine
            bcode += "TEXT 696,158,""0"",180,22,22,""UPC""" & vbNewLine
            bcode += "TEXT 712,378,""0"",180,10,10,""" & modelnm & """" & vbNewLine
            bcode += "BOX 56,344,725,387,3" & vbNewLine
            bcode += "TEXT 697,76,""0"",180,22,22,""QTY""" & vbNewLine
            bcode += "BOX 57,259,550,346,3" & vbNewLine
            bcode += "BARCODE 484,337,""128M"",42,0,180,2,4,""!104PK!099" & Mid(LOT, 3, 8) & "!100-!099" & Right(LOT, 4) & """" & vbNewLine
            bcode += "TEXT 431,290,""0"",180,11,11,""" & LOT & """" & vbNewLine
            bcode += "TEXT 694,245,""0"",180,22,22,""SKU""" & vbNewLine
            bcode += "BARCODE 493,253,""128M"",41,0,180,3,6,""!104" & model & """" & vbNewLine
            bcode += "TEXT 385,207,""0"",180,12,12,""" & model & """" & vbNewLine
            bcode += "BOX 548,259,725,346,3" & vbNewLine
            bcode += "TEXT 690,329,""0"",180,22,22,""LOT""" & vbNewLine
            bcode += "BOX 548,174,725,261,3" & vbNewLine
            bcode += "BOX 57,174,550,261,3" & vbNewLine
            bcode += "BOX 548,88,725,175,3" & vbNewLine
            bcode += "BOX 57,88,550,175,3" & vbNewLine
            bcode += "BOX 548,3,725,90,3" & vbNewLine
            bcode += "BOX 57,3,550,90,3" & vbNewLine
            bcode += "TEXT 354, 73, ""0"", 180, 20, 20,""" & QTY & """" & vbNewLine
            bcode += "PRINT 1,1" & vbNewLine
            bcode += "<xpml></page></xpml><xpml><End/></xpml>" & vbNewLine

            Rt.Text = bcode

            'Microsoft.VisualBasic.FileSystem.Write(FN, bcode)
            'Microsoft.VisualBasic.FileClose()

            Rt.SaveFile("C:\TEMP\LOT" & idx & ".PRN", RichTextBoxStreamType.PlainText)

            Shell("c:\windows\system32\cmd.exe /c COPY c:\TEMP\LOT" & idx & ".PRN \\" & My.Computer.Name & "\BP04")

        End If

    End Function

    Function Print_LOTBarcode(ByVal LOT As String, ByVal model As String, ByVal QTY As Integer, ByVal Rt As RichTextBox, ByVal DEF As String) As Boolean

        'Dim FN
        Dim bcode As String

        Rt.Text = ""

        'FN = Microsoft.VisualBasic.FileSystem.FreeFile
        'Microsoft.VisualBasic.FileOpen(FN, "C:\TEMP\LOT.PRN", OpenMode.Output)
        ''        Microsoft.VisualBasic.FileSystem.Write(FN, "/ ^XA" & vbNewLine & "/ ^LT0" & vbNewLine & "/ ^LH40,0" & vbNewLine & "/ ^FO1,100^A0N,13,25^FD" & model & "-" & def & "^FS" & vbNewLine & "/ ^FO1,50^BY1" & vbNewLine & "/ ^B3N,N,30,Y,N" & vbNewLine & "/ ^FD" & esn & "^FS" & vbNewLine & "/ ^XZ")

        If Site_id = "S1000" Then


            bcode = "<xpml><page quantity='0' pitch='50.0 mm'></xpml>SIZE 99.10 mm, 50 mm" & vbNewLine
            bcode = bcode & "DIRECTION 0,0" & vbNewLine
            bcode = bcode & "REFERENCE 0,0" & vbNewLine
            bcode = bcode & "OFFSET 0 mm" & vbNewLine
            bcode = bcode & "SET PEEL OFF" & vbNewLine
            bcode = bcode & "SET CUTTER OFF" & vbNewLine

            bcode = bcode & "SET PARTIAL_CUTTER OFF" & vbNewLine
            bcode = bcode & "<xpml></page></xpml><xpml><page quantity='1' pitch='50.0 mm'></xpml>SET TEAR ON" & vbNewLine
            bcode = bcode & "CLS" & vbNewLine
            bcode = bcode & "CODEPAGE 949" & vbNewLine
            bcode = bcode & "TEXT 732,356," & """0""" & ",180,10,10," & """" & "EXODUS WIRELESS KOREA" & """" & vbNewLine

            bcode = bcode & "BARCODE 727,211," & """" & "128M" & """" & ",102,0,180,3,6," & """" & "!104" & Mid(LOT, 1, 1) & "!099" & Mid(LOT, 2, 8) & "!100-!099" & Mid(LOT, 11, 4) & """" & vbNewLine
            bcode = bcode & "TEXT 582,104," & """" & "0" & """" & ",180,9,9," & """" & LOT & """" & vbNewLine
            bcode = bcode & "TEXT 732,315," & """" & "0" & """" & ",180,10,10," & """" & "MODEL : " & model & """" & vbNewLine

            bcode = bcode & "TEXT 732,276," & """" & "HYWULM.TTF" & """" & ",180,10,10," & """" & "불량 : " & Query_RS("SELECT CODE_NAME FROM TBL_DEFECT WHERE CODE_ID = '" & DEF & "'") & "(" & DEF & ")" & """" & vbNewLine
            bcode = bcode & "TEXT 732,235," & """" & "0" & """" & ",180,10,10," & """" & "LOT QTY : " & QTY & """" & vbNewLine
            bcode = bcode & "PRINT 1,1" & vbNewLine
            bcode = bcode & "<xpml></page></xpml><xpml><end/></xpml>" & vbNewLine

            Rt.Text = bcode

            'Microsoft.VisualBasic.FileSystem.Write(FN, bcode)
            'Microsoft.VisualBasic.FileClose()

            Rt.SaveFile("C:\TEMP\LOT.PRN", RichTextBoxStreamType.PlainText)

            Shell("c:\windows\system32\cmd.exe /c COPY c:\TEMP\LOT.PRN \\" & My.Computer.Name & "\BP02")

        End If

    End Function

    Function Print_PalletBarcode(ByVal INV As String, ByVal PALLET As String, ByVal REV As String, ByVal Rt As RichTextBox) As Boolean

        'Dim FN
        Dim bcode As String

        Rt.Text = ""

        If Site_id = "S1000" Then


            bcode = "<xpml><page quantity='0' pitch='50.0 mm'></xpml>SIZE 99.10 mm, 50 mm" & vbNewLine
            bcode = bcode & "DIRECTION 0,0" & vbNewLine
            bcode = bcode & "REFERENCE 0,0" & vbNewLine
            bcode = bcode & "OFFSET 0 mm" & vbNewLine
            bcode = bcode & "SET PEEL OFF" & vbNewLine
            bcode = bcode & "SET CUTTER OFF" & vbNewLine

            bcode = bcode & "SET PARTIAL_CUTTER OFF" & vbNewLine
            bcode = bcode & "<xpml></page></xpml><xpml><page quantity='1' pitch='50.0 mm'></xpml>SET TEAR ON" & vbNewLine
            bcode = bcode & "CLS" & vbNewLine
            bcode = bcode & "BOX 560,215,737,328,3" & vbNewLine
            bcode = bcode & "BOX 68,215,562,328,3" & vbNewLine
            bcode = bcode & "CODEPAGE 1254" & vbNewLine

            bcode = bcode & "TEXT 724,300," & """" & "ROMAN.TTF" & """" & ",180,1,12," & """" & "INVOICE NO." & """" & vbNewLine
            bcode = bcode & "BARCODE 516,314," & """" & "128M" & """" & ",56,0,180,2,4," & """" & "!104" & INV & """" & vbNewLine
            bcode = bcode & "TEXT 393,252," & """" & "ROMAN.TTF" & """" & ",180,1,12," & """" & INV & """" & vbNewLine

            bcode = bcode & "TEXT 724,87," & """" & "ROMAN.TTF" & """" & ",180,1,12," & """" & "PALLET NO." & """" & vbNewLine
            bcode = bcode & "TEXT 370,85," & """" & "ROMAN.TTF" & """" & ",180,1,22," & """" & PALLET & """" & vbNewLine

            bcode = bcode & "TEXT 724,189," & """" & "ROMAN.TTF" & """" & ",180,1,12," & """" & "REVISION" & """" & vbNewLine
            bcode = bcode & "TEXT 340,180," & """" & "ROMAN.TTF" & """" & ",180,1,22," & """" & REV & """" & vbNewLine

            bcode = bcode & "TEXT 724,369," & """" & "ROMAN.TTF" & """" & ",180,1,12," & """" & "SCOTTII KOREA" & """" & vbNewLine

            bcode = bcode & "BOX 560,104,737,217,3" & vbNewLine
            bcode = bcode & "BOX 68,104,562,217,3" & vbNewLine
            bcode = bcode & "BOX 560,12,737,106,3" & vbNewLine
            bcode = bcode & "BOX 68,326,737,388,3" & vbNewLine
            bcode = bcode & "BOX 68,12,562,106,3" & vbNewLine

            bcode = bcode & "PRINT 1,1" & vbNewLine
            bcode = bcode & "<xpml></page></xpml><xpml><end/></xpml>" & vbNewLine

            Rt.Text = bcode

            'Microsoft.VisualBasic.FileSystem.Write(FN, bcode)
            'Microsoft.VisualBasic.FileClose()

            Rt.SaveFile("C:\TEMP\LOT.PRN", RichTextBoxStreamType.PlainText)

            Shell("c:\windows\system32\cmd.exe /c COPY c:\TEMP\LOT.PRN \\" & My.Computer.Name & "\BP04", AppWinStyle.Hide, True)

        End If

    End Function


    Function Print_BOXBarcode(ByVal LOT As String, ByVal model As String, ByVal QTY As Integer, ByVal Rt As RichTextBox) As Boolean

        'Dim FN
        Dim bcode As String

        Rt.Text = ""

        If Site_id = "S1000" Then


            bcode = "<xpml><page quantity='0' pitch='50.0 mm'></xpml>SIZE 99.10 mm, 50 mm" & vbNewLine
            bcode = bcode & "DIRECTION 0,0" & vbNewLine
            bcode = bcode & "REFERENCE 0,0" & vbNewLine
            bcode = bcode & "OFFSET 0 mm" & vbNewLine
            bcode = bcode & "SET PEEL OFF" & vbNewLine
            bcode = bcode & "SET CUTTER OFF" & vbNewLine

            bcode = bcode & "SET PARTIAL_CUTTER OFF" & vbNewLine
            bcode = bcode & "<xpml></page></xpml><xpml><page quantity='1' pitch='50.0 mm'></xpml>SET TEAR ON" & vbNewLine
            bcode = bcode & "CLS" & vbNewLine
            bcode = bcode & "BOX 560,215,737,328,3" & vbNewLine
            bcode = bcode & "BOX 68,215,562,328,3" & vbNewLine
            bcode = bcode & "CODEPAGE 1254" & vbNewLine

            bcode = bcode & "TEXT 706,300," & """" & "ROMAN.TTF" & """" & ",180,1,22," & """" & "SKU" & """" & vbNewLine
            bcode = bcode & "BARCODE 506,314," & """" & "128M" & """" & ",56,0,180,3,6," & """" & "!104" & model & """" & vbNewLine
            bcode = bcode & "TEXT 393,252," & """" & "ROMAN.TTF" & """" & ",180,1,12," & """" & model & """" & vbNewLine

            bcode = bcode & "TEXT 704,87," & """" & "ROMAN.TTF" & """" & ",180,1,22," & """" & "QTY" & """" & vbNewLine
            bcode = bcode & "BARCODE 417,100," & """" & "128M" & """" & ",48,0,180,3,6," & """" & "!10510!1000" & """" & vbNewLine
            bcode = bcode & "TEXT 340,47," & """" & "ROMAN.TTF" & """" & ",180,1,12," & """" & QTY & """" & vbNewLine

            bcode = bcode & "TEXT 706,189," & """" & "ROMAN.TTF" & """" & ",180,1,22," & """" & "UPC" & """" & vbNewLine
            bcode = bcode & "BARCODE 457,207," & """" & "UPCA" & """" & ",56,0,180,3,6," & """" & LOT & """" & vbNewLine
            bcode = bcode & "TEXT 423,145," & """" & "ROMAN.TTF" & """" & ",180,1,10," & """" & "1" & """" & vbNewLine

            bcode = bcode & "TEXT 724,369," & """" & "ROMAN.TTF" & """" & ",180,1,12," & """" & "TPU Case Dark Blue 6/6S R1" & """" & vbNewLine

            bcode = bcode & "TEXT 403,145," & """" & "ROMAN.TTF" & """" & ",180,1,10," & """" & "5" & """" & vbNewLine
            bcode = bcode & "TEXT 381,145," & """" & "ROMAN.TTF" & """" & ",180,1,10," & """" & "6" & """" & vbNewLine
            bcode = bcode & "TEXT 358,145," & """" & "ROMAN.TTF" & """" & ",180,1,10," & """" & "1" & """" & vbNewLine
            bcode = bcode & "TEXT 338,145," & """" & "ROMAN.TTF" & """" & ",180,1,10," & """" & "5" & """" & vbNewLine
            bcode = bcode & "TEXT 307,145," & """" & "ROMAN.TTF" & """" & ",180,1,10," & """" & "0" & """" & vbNewLine
            bcode = bcode & "TEXT 286,145," & """" & "ROMAN.TTF" & """" & ",180,1,10," & """" & "2" & """" & vbNewLine
            bcode = bcode & "TEXT 264,145," & """" & "ROMAN.TTF" & """" & ",180,1,10," & """" & "0" & """" & vbNewLine
            bcode = bcode & "TEXT 242,145," & """" & "ROMAN.TTF" & """" & ",180,1,10," & """" & "1" & """" & vbNewLine
            bcode = bcode & "TEXT 221,145," & """" & "ROMAN.TTF" & """" & ",180,1,10," & """" & "0" & """" & vbNewLine
            bcode = bcode & "TEXT 484,145," & """" & "ROMAN.TTF" & """" & ",180,1,7," & """" & "8" & """" & vbNewLine
            bcode = bcode & "TEXT 155,145," & """" & "ROMAN.TTF" & """" & ",180,1,7," & """" & "3" & """" & vbNewLine

            bcode = bcode & "BOX 560,104,737,217,3" & vbNewLine
            bcode = bcode & "BOX 68,104,562,217,3" & vbNewLine
            bcode = bcode & "BOX 560,12,737,106,3" & vbNewLine
            bcode = bcode & "BOX 68,326,737,388,3" & vbNewLine
            bcode = bcode & "BOX 68,12,562,106,3" & vbNewLine

            bcode = bcode & "PRINT 1,1" & vbNewLine
            bcode = bcode & "<xpml></page></xpml><xpml><end/></xpml>" & vbNewLine

            Rt.Text = bcode

            'Microsoft.VisualBasic.FileSystem.Write(FN, bcode)
            'Microsoft.VisualBasic.FileClose()

            Rt.SaveFile("C:\TEMP\LOT.PRN", RichTextBoxStreamType.PlainText)

            Shell("c:\windows\system32\cmd.exe /c COPY c:\TEMP\LOT.PRN \\" & My.Computer.Name & "\BP02")

        End If

    End Function


    Function Print_ShipBarcode(ByVal esn As String, ByVal model As String, ByVal Rt As RichTextBox) As Boolean
        Dim bcode As String

        Rt.Text = ""


        bcode = "<xpml><page quantity='0' pitch='50.0 mm'></xpml>SIZE 47.5 mm, 13 mm" & vbNewLine
        bcode = bcode & "DIRECTION 0,0" & vbNewLine
        bcode = bcode & "REFERENCE 0,0" & vbNewLine
        bcode = bcode & "OFFSET 0 mm" & vbNewLine
        bcode = bcode & "SET PEEL OFF" & vbNewLine
        bcode = bcode & "SET CUTTER OFF" & vbNewLine

        bcode = bcode & "SET PARTIAL_CUTTER OFF" & vbNewLine
        bcode = bcode & "<xpml></page></xpml><xpml><page quantity='1' pitch='13.0 mm'></xpml>SET TEAR ON" & vbNewLine
        bcode = bcode & "CLS" & vbNewLine
        bcode = bcode & "CODEPAGE 1254" & vbNewLine
        bcode = bcode & "TEXT 353,92," & """0""" & ",180,8,8," & """" & "EXODUS" & """" & vbNewLine
        bcode = bcode & "TEXT 220,92," & """" & "0" & """" & ",180,8,8," & """" & model & """" & vbNewLine

        bcode = bcode & "BARCODE 356,66," & """" & "128M" & """" & ",32,0,180,2,4," & """" & "!104" & "K" & "!099" & Mid(esn, 2, 14) & "!100-" & Mid(esn, 17, 1) & """" & vbNewLine
        bcode = bcode & "TEXT 274,28," & """" & "0" & """" & ",180,7,7," & """" & "K" & Mid(esn, 2, Len(esn) - 1) & """" & vbNewLine

        bcode = bcode & "PRINT 1,1" & vbNewLine
        bcode = bcode & "<xpml></page></xpml><xpml><end/></xpml>" & vbNewLine

        Rt.Text = bcode

        'Microsoft.VisualBasic.FileSystem.Write(FN, bcode)
        'Microsoft.VisualBasic.FileClose()

        Rt.SaveFile("C:\TEMP\ESN.PRN", RichTextBoxStreamType.PlainText)

        Shell("c:\windows\system32\cmd.exe /c COPY c:\TEMP\ESN.PRN \\" & My.Computer.Name & "\BP01")

    End Function

    Function Print_SMSBarcode(ByVal esn As String, ByVal Rt As RichTextBox) As Boolean
        'Product serial Barcode
        Dim bcode As String

        Rt.Text = ""

        bcode = "<xpml><page quantity='0' pitch='20.0 mm'></xpml>SIZE 37.5 mm, 20 mm" & vbNewLine
        bcode = bcode & "DIRECTION 0,0" & vbNewLine
        bcode = bcode & "REFERENCE 0,0" & vbNewLine
        bcode = bcode & "OFFSET 0 mm" & vbNewLine
        bcode = bcode & "SET PEEL OFF" & vbNewLine
        bcode = bcode & "SET CUTTER OFF" & vbNewLine

        bcode = bcode & "SET PARTIAL_CUTTER OFF" & vbNewLine
        bcode = bcode & "<xpml></page></xpml><xpml><page quantity='1' pitch='20.0 mm'></xpml>SET TEAR ON" & vbNewLine
        bcode = bcode & "CLS" & vbNewLine
        bcode = bcode & "CODEPAGE 1254" & vbNewLine

        bcode = bcode & "BARCODE 278,138," & """" & "128M" & """" & ",85,0,180,1,2," & """" & "!104" & esn & """" & vbNewLine
        bcode = bcode & "TEXT 258,48," & """" & "0" & """" & ",180,8,8," & """" & esn & """" & vbNewLine

        bcode = bcode & "PRINT 1,1" & vbNewLine
        bcode = bcode & "<xpml></page></xpml><xpml><end/></xpml>" & vbNewLine

        Rt.Text = bcode

        Rt.SaveFile("C:\TEMP\ESN.PRN", RichTextBoxStreamType.PlainText)

        Shell("c:\windows\system32\cmd.exe /c COPY c:\TEMP\ESN.PRN \\" & My.Computer.Name & "\BP02", AppWinStyle.MinimizedFocus, True)


        Rt.Clear()

    End Function


    Function Print_NewBarcode(ByVal esn As String, ByVal model As String, ByVal def As String, ByVal dt As String) As Boolean

        Dim FN
        Dim bcode As String

        Dim CUS, Color As String
        Dim ba_rs As New ADODB.Recordset
        ba_rs = Query_RS_ALL("SELECT SW_VER, color FROM TBL_MODELMASTER WHERE MODEL_NO = (SELECT MODEL FROM TBL_FESNMASTER_K WHERE OUT_ESN = '" & esn & "')")
        CUS = ba_rs(0).Value
        Color = ba_rs(1).Value
        ba_rs = Nothing

        'Dim ARR As String() = GET_TRG(esn, "FLIP")
        'ba_rs = Query_RS_ALL("SELECT def_cd FROM TBL_triage WHERE esn_obid = (SELECT obid FROM TBL_FESNMASTER WHERE ESN = '" & esn & "') order by c_date")

        'Dim def1 As String = ARR(4)
        'Dim def2 As String = ARR(3)
        'Dim def3 As String = ARR(2)
        'Dim def4 As String = ARR(1)
        'Dim def5 As String = ARR(0)

        FN = Microsoft.VisualBasic.FileSystem.FreeFile
        Microsoft.VisualBasic.FileOpen(FN, "C:\TEMP\ESN.ZPL", OpenMode.Output)
        '        Microsoft.VisualBasic.FileSystem.Write(FN, "/ ^XA" & vbNewLine & "/ ^LT0" & vbNewLine & "/ ^LH40,0" & vbNewLine & "/ ^FO1,100^A0N,13,25^FD" & model & "-" & def & "^FS" & vbNewLine & "/ ^FO1,50^BY1" & vbNewLine & "/ ^B3N,N,30,Y,N" & vbNewLine & "/ ^FD" & esn & "^FS" & vbNewLine & "/ ^XZ")

        If Site_id = "S1000" Then

            esn = Replace(esn, "F", "K")

            bcode = "^XA~TA000~JSN^LT0^MMT^MNW^MTT^PON^PMN^LH0,0^JMA^PR2,2^MD10^JUS^LRN^CI0^XZ" & vbNewLine
            bcode = bcode & "^XA^LL0150" & vbNewLine
            bcode = bcode & "^PW510" & vbNewLine
            bcode = bcode & "^BY1,3,30^FT100,90^BCN,30,N,N" & vbNewLine
            '            bcode = bcode & "^FD>:" & Mid(esn, 1, 1) & ">5" & Mid(esn, 2, Len(esn) - 4) & ">6" & Mid(esn, Len(esn) - 2, 3) & "^FS" & vbNewLine
            bcode = bcode & "^FD>:" & esn & "^FS" & vbNewLine
            bcode = bcode & "^FT100,50^A0N,15,20^FH\^FD" & esn & "^FS" & vbNewLine

            bcode = bcode & "^FT386,59^A0N,25,24^FH\^FD" & "" & "^FS" & vbNewLine
            bcode = bcode & "^FT331,59^A0N,25,24^FH\^FD" & "" & "^FS" & vbNewLine
            bcode = bcode & "^FT386,26^A0N,25,24^FH\^FD" & "" & "^FS" & vbNewLine
            bcode = bcode & "^FT441,27^A0N,25,24^FH\^FD" & "" & "^FS" & vbNewLine
            bcode = bcode & "^FT332,27^A0N,25,24^FH\^FD" & "" & "^FS" & vbNewLine

            bcode = bcode & "^FT200,30^A0N,15,20^FH\^FD" & model & "^FS" & vbNewLine
            '            bcode = bcode & "^FT250,56^A0N,29,28^FH\^FD" & Color & "^FS" & vbNewLine

            bcode = bcode & "^FT100,30^A0N,15,20^FH\^FD" & "EXODUS" & "^FS" & vbNewLine
            'bcode = bcode & "^FT439,59^A0N,42,40^FH\^FD:^FS" & vbNewLine
            'bcode = bcode & "^PQ1,0,1,Y^XZ" & vbNewLine
            bcode = bcode & "^PQ1,0,1,Y^XZ" & vbNewLine

            Microsoft.VisualBasic.FileSystem.Write(FN, bcode)
            Microsoft.VisualBasic.FileClose()

            'Dim PORT As String = "" 'Query_RS("SELECT CN_TYPE FROM tbl_bcconfig WHERE MODEL = (SELECT MODEL FROM TBL_FESNMASTER WHERE ESN = '" & esn & "') AND BC_TYPE = 'RECEIVING'")
            'If PORT = "" Then
            '    PORT = "COM1"
            'End If

            '            Shell("c:\windows\system32\cmd.exe /c print c:\TEMP\esn.zpl /d:" & PORT)
            Shell("c:\windows\system32\cmd.exe /c COPY c:\TEMP\esn.zpl \\" & My.Computer.Name & "\BP01")
            '            Shell("c:\windows\system32\cmd.exe /c print c:\esn.zpl /d:com1")
        End If

    End Function

    Function Print_UPCBarcode(ByVal model As String, ByVal TYPE As String, ByVal Rt As RichTextBox) As Boolean

        Dim FN
        Dim bcode As String
        Dim UPC As String = ""
        Dim RS As New ADODB.Recordset

        Rt.Text = ""

        Dim QRY As String = ""
        QRY = "SELECT SKU_NM, "

        If TYPE = "AMAZON" Then
            QRY = QRY & "AMAZON "
        ElseIf TYPE = "GTIN14" Then
            QRY = QRY & "GTIN14 "
        ElseIf TYPE = "GTIN12" Then
            QRY = QRY & "GTIN12 "
        End If
        QRY = QRY & "FROM TBL_BARCODE WHERE MODEL = '" & model & "'"

        RS = Query_RS_ALL(QRY)

        If RS Is Nothing Then
            Modal_Error("바코드가 등록되지 않은 모델입니다.")
            Return False
            Exit Function
        End If

        '        FN = Microsoft.VisualBasic.FileSystem.FreeFile
        '       Microsoft.VisualBasic.FileOpen(FN, "C:\TEMP\ESN1.ZPL", OpenMode.Output)
        '        Microsoft.VisualBasic.FileSystem.Write(FN, "/ ^XA" & vbNewLine & "/ ^LT0" & vbNewLine & "/ ^LH40,0" & vbNewLine & "/ ^FO1,100^A0N,13,25^FD" & model & "-" & def & "^FS" & vbNewLine & "/ ^FO1,50^BY1" & vbNewLine & "/ ^B3N,N,30,Y,N" & vbNewLine & "/ ^FD" & esn & "^FS" & vbNewLine & "/ ^XZ")

        bcode = "CT~~CD,~CC^~CT~" & vbNewLine
        bcode = bcode & "^XA~TA000~JSN^LT0^MNW^MTT^PON^PMN^LH0,0^JMA^PR4,4~SD15^JUS^LRN^CI0^XZ" & vbNewLine
        bcode = bcode & "^XA" & vbNewLine
        bcode = bcode & "^MMT" & vbNewLine
        bcode = bcode & "^PW591" & vbNewLine
        bcode = bcode & "^LL0295" & vbNewLine
        bcode = bcode & "^LS0" & vbNewLine
        bcode = bcode & "^BY4,2,152^FT105,184^BUN,,Y,N" & vbNewLine
        bcode = bcode & "^FD" & RS(1).Value & "^FS" & vbNewLine
        bcode = bcode & "^FT29,260^A0N,37,36^FH\^FD" & RS(0).Value & "^FS" & vbNewLine
        bcode = bcode & "^PQ1,0,1,Y^XZ"


        Rt.Text = bcode

        Rt.SaveFile("C:\TEMP\ESN1.ZPL", RichTextBoxStreamType.PlainText)
        'Microsoft.VisualBasic.FileSystem.Write(FN, bcode)
        'Microsoft.VisualBasic.FileClose()

        Shell("c:\windows\system32\cmd.exe /c PRINT c:\TEMP\ESN1.ZPL /D:\\" & My.Computer.Name & "\BP03", AppWinStyle.MinimizedFocus, True)


    End Function

    Function Print_NewBarcode_rcv(ByVal esn As String, ByVal model As String, ByVal def As String, ByVal dt As String) As Boolean

        Dim FN
        Dim bcode As String

        FN = Microsoft.VisualBasic.FileSystem.FreeFile
        Microsoft.VisualBasic.FileOpen(FN, "C:\ESN.ZPL", OpenMode.Output)

        If Site_id = "S1000" Then
            bcode = "^XA~TA000~JSN^LT0^MMT^MNW^MTT^PON^PMN^LH0,0^JMA^PR4,4^MD0^JUS^LRN^CI0^XZ" & vbNewLine
            bcode = bcode & "^XA^LL0226" & vbNewLine
            bcode = bcode & "^BY1,3,29^FT5,91^BCN,,N,N" & vbNewLine
            '            bcode = bcode & "^FD>:" & Mid(esn, 1, 2) & ">5" & Mid(esn, 3, 8) & ">6" & Mid(esn, 11, 7) & "^FS" & vbNewLine
            bcode = bcode & "^FD>:" & esn & "^FS" & vbNewLine
            bcode = bcode & "^FT5,53^A0N,29,28^FH\^FD" & esn & "^FS" & vbNewLine
            bcode = bcode & "^FT57,96^A0N,17,16^FH\^FD" & "                                                   " & Mid(dt, 5, 2) & "/" & Mid(dt, 7, 2) & "/" & Mid(dt, 1, 4) & "^FS" & vbNewLine
            bcode = bcode & "^FT59,76^A0N,17,16^FH\^FD" & "                                                   " & model & "^FS" & vbNewLine
            '            bcode = bcode & "^FT100,57^A0N,42,25^FH\^FD" & "                      " & def & "^FS" & vbNewLine
            bcode = bcode & "^FT100,57^A0N,42,25^FH\^FD" & "                                    " & def & "^FS" & vbNewLine
            bcode = bcode & "^PQ1,0,1,Y^XZ" & vbNewLine

            Microsoft.VisualBasic.FileSystem.Write(FN, bcode)
            Microsoft.VisualBasic.FileClose()
            Shell("c:\windows\system32\cmd.exe /c print c:\esn.zpl /d:com1")
        End If

    End Function

    Function Print_FailBarcode(ByVal esn As String, ByVal model As String, ByVal def As String, ByVal dt As String, ByVal fw As String, ByVal tw As String) As Boolean

        Dim FN
        Dim bcode As String

        FN = Microsoft.VisualBasic.FileSystem.FreeFile
        Microsoft.VisualBasic.FileOpen(FN, "C:\ESN.ZPL", OpenMode.Output)
        '        Microsoft.VisualBasic.FileSystem.Write(FN, "/ ^XA" & vbNewLine & "/ ^LT0" & vbNewLine & "/ ^LH40,0" & vbNewLine & "/ ^FO1,100^A0N,13,25^FD" & model & "-" & def & "^FS" & vbNewLine & "/ ^FO1,50^BY1" & vbNewLine & "/ ^B3N,N,30,Y,N" & vbNewLine & "/ ^FD" & esn & "^FS" & vbNewLine & "/ ^XZ")

        If Site_id = "S1000" Then


            bcode = "^XA~TA000~JSN^LT0^MMT^MNW^MTT^PON^PMN^LH0,0^JMA^PR4,4^MD0^JUS^LRN^CI0^XZ" & vbNewLine
            bcode = bcode & "^XA^LL0226" & vbNewLine
            bcode = bcode & "^FT5,54^A0N,30,28^FH\^FD" & esn & "^FS" & vbNewLine
            bcode = bcode & "^FT55,95^A0N,17,16^FH\^FD" & "                                                   " & dt & "^FS" & vbNewLine
            bcode = bcode & "^FT56,76^A0N,17,16^FH\^FD" & "                                                   " & model & "^FS" & vbNewLine
            bcode = bcode & "^FT100,57^A0N,42,40^FH\^FD" & "                      " & def & "^FS" & vbNewLine
            bcode = bcode & "^FT5,91^A0N,33,33^FH\^FD" & fw & "^FS" & vbNewLine
            bcode = bcode & "^FT4,89^A0N,33,33^FH\^FD" & "            " & tw & "^FS" & vbNewLine
            bcode = bcode & "^FT5,87^ABN,22,14^FH\^FD" & "     >" & "^FS" & vbNewLine
            bcode = bcode & "^PQ1,0,1,Y^XZ" & esn & "^FS" & vbNewLine

            Microsoft.VisualBasic.FileSystem.Write(FN, bcode)
            Microsoft.VisualBasic.FileClose()
            Shell("c:\windows\system32\cmd.exe /c print c:\esn.zpl /d:com1")
        End If

    End Function



    Function Print_Barcode_SPC(ByVal esn As String, ByVal spc As String, ByVal model As String, ByVal rcv As String) As Boolean

        Dim FN
        Dim bcode As String

        FN = Microsoft.VisualBasic.FileSystem.FreeFile
        Microsoft.VisualBasic.FileOpen(FN, "C:\ESN.ZPL", OpenMode.Output)

        If Site_id = "S1000" Then
            bcode = "^XA~TA000~JSN^LT0^MMT^MNW^MTT^PON^PMN^LH0,0^JMA^PR4,4^MD0^JUS^LRN^CI0^XZ" & vbNewLine
            bcode = bcode & "^XA^LL0226" & vbNewLine
            bcode = bcode & "^FT56,49^A0N,29,28^FH\^FD" & esn & "^FS" & vbNewLine
            bcode = bcode & "^FT56,99^AAN,36,15^FH\^FD" & spc & "^FS" & vbNewLine
            bcode = bcode & "^FT450,98^A0N,29,28^FH\^FD" & model & "^FS" & vbNewLine
            bcode = bcode & "^PQ1,0,1,Y^XZ" & vbNewLine

            Microsoft.VisualBasic.FileSystem.Write(FN, bcode)
            Microsoft.VisualBasic.FileClose()
            Shell("c:\windows\system32\cmd.exe /c print c:\esn.zpl /d:com1")
        End If


    End Function

    Function Print_Barcode_Coffin(ByVal dec As String, ByVal hex As String, ByVal model As String) As Boolean

        Dim FN
        Dim bcode As String

        FN = Microsoft.VisualBasic.FileSystem.FreeFile
        Microsoft.VisualBasic.FileOpen(FN, "C:\COFFIN.ZPL", OpenMode.Output)

        bcode = "^XA" & vbNewLine
        bcode = bcode & "^SZ2^JMA" & vbNewLine
        bcode = bcode & "^MCY^PMN" & vbNewLine
        bcode = bcode & "^PW988~JSN" & vbNewLine
        bcode = bcode & "^JZY" & vbNewLine
        bcode = bcode & "^LH0,0^LRN" & vbNewLine
        bcode = bcode & "^XZ" & vbNewLine
        '        bcode = bcode & "~DGR:SSGFX000.GRF,21079,107,:Z64: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:5B56" & vbNewLine
        bcode = bcode & "^XA" & vbNewLine
        bcode = bcode & "^DFE:SSFMT000.ZPL^FS" & vbNewLine
        bcode = bcode & "^FO60,149" & vbNewLine
        bcode = bcode & "^BY2^BCN,95,N,Y^FD>:" & model & "^FS" & vbNewLine
        bcode = bcode & "^FT040,120^A0N,32,60^FH\^FD" & model & "^FS" & vbNewLine

        bcode = bcode & "^FO495,54" & vbNewLine
        bcode = bcode & "^BY3^BCN,57,N,N^FD>;" & dec & "^FS" & vbNewLine
        bcode = bcode & "^FT0495,150^A0N,32,38^FH\^FDDEC " & dec & "^FS" & vbNewLine

        bcode = bcode & "^FO496,168" & vbNewLine
        bcode = bcode & "^BY3^BCN,57,N,N^FD>;" & hex & "^FS" & vbNewLine
        bcode = bcode & "^FT0495,260^A0N,32,38^FH\^FDHEX " & hex & "^FS" & vbNewLine

        bcode = bcode & "^FO44,63" & vbNewLine
        '        bcode = bcode & "^XGR:SSGFX000.GRF,1,1^FS" & vbNewLine
        bcode = bcode & "^XZ" & vbNewLine
        bcode = bcode & "^XA" & vbNewLine
        bcode = bcode & "^XFE:SSFMT000.ZPL^FS" & vbNewLine
        bcode = bcode & "^PQ1,0,1,Y" & vbNewLine
        bcode = bcode & "^XZ" & vbNewLine


        Microsoft.VisualBasic.FileSystem.Write(FN, bcode)
        Microsoft.VisualBasic.FileClose()

        Dim PORT As String = Query_RS("SELECT CN_TYPE FROM tbl_bcconfig WHERE MODEL = (SELECT MODEL FROM TBL_ESNMASTER WHERE ESN = '" & hex & "') AND BC_TYPE = 'COFFIN BOX'")
        If PORT = "" Then
            PORT = "COM3"
        End If

        Shell("c:\windows\system32\cmd.exe /c print c:\COFFIN.zpl /d:" & PORT)

        '        Shell("c:\windows\system32\cmd.exe /c print c:\COFFIN.zpl /d:COM3")

    End Function

    Function Print_Barcode_Carton(ByVal boxid As String, ByVal model As String, ByVal prl As String, ByVal model1 As String, ByVal cnt As Integer) As Boolean

        Dim FN
        Dim bcode As String
        Dim cb_rs As ADODB.Recordset = Query_RS_ALL("select reserv5 from tbl_esnmaster where inboxid = '" & boxid & "' order by pos")
        Dim j As Integer = 0

        If cb_rs Is Nothing Then
            Exit Function
        End If

        FN = Microsoft.VisualBasic.FileSystem.FreeFile
        Microsoft.VisualBasic.FileOpen(FN, "C:\CARTON.ZPL", OpenMode.Output)

        bcode = "^XA" & vbNewLine
        bcode = bcode & "^SZ2^JMA" & vbNewLine
        bcode = bcode & "^MCY^PMN" & vbNewLine
        bcode = bcode & "^PW1353~JSN" & vbNewLine
        bcode = bcode & "^JZY" & vbNewLine
        bcode = bcode & "^LH0,0^LRN" & vbNewLine
        bcode = bcode & "^XZ" & vbNewLine
        '        bcode = bcode & "~DGR:SSGFX000.GRF,226100,133,:Z64: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:8CAF" & vbNewLine
        bcode = bcode & "^XA" & vbNewLine
        bcode = bcode & "^DFE:SSFMT000.ZPL^FS" & vbNewLine

        bcode = bcode & "^FT0100,100^A0N,30,60^FH\^FDSKU : " & model & "^FS" & vbNewLine

        bcode = bcode & "^FO725,86" & vbNewLine
        bcode = bcode & "^BY2^BCN,57,N,N^FD>:" & model & "^FS" & vbNewLine

        bcode = bcode & "^FT0725,70^A0N,30,28^FH\^FDSKU : " & model & "^FS" & vbNewLine

        bcode = bcode & "^FO154,219" & vbNewLine
        bcode = bcode & "^BY4^BCN,57,N,N^FD>;" & cnt & "^FS" & vbNewLine

        bcode = bcode & "^FT0154,200^A0N,30,28^FH\^FDCarton Qty : " & cnt & "^FS" & vbNewLine

        bcode = bcode & "^FO550,219" & vbNewLine
        bcode = bcode & "^BY3^BCN,57,N,N^FD>;" & Mid(prl, 1, Len(prl) - 1) & ">6" & Mid(prl, Len(prl), 1) & ", " & model1 & "^FS" & vbNewLine

        bcode = bcode & "^FT0550,200^A0N,30,28^FH\^FDPRL/SW Ver : " & prl & ", " & model1 & "^FS" & vbNewLine

        bcode = bcode & "^FO74,160" & vbNewLine
        bcode = bcode & "^GB1111,0,3^FS" & vbNewLine
        bcode = bcode & "^FO646,29" & vbNewLine
        bcode = bcode & "^GB0,133,3^FS" & vbNewLine
        bcode = bcode & "^FO69,293" & vbNewLine
        bcode = bcode & "^GB1121,0,3^FS" & vbNewLine
        bcode = bcode & "^FO494,161" & vbNewLine
        bcode = bcode & "^GB0,133,3^FS" & vbNewLine

        bcode = bcode & "^FT0350,350^A0N,50,100^FH\^FDReconditioned^FS" & vbNewLine
        bcode = bcode & "^FT0150,400^A0N,30,40^FH\^FDDEC :^FS" & vbNewLine
        bcode = bcode & "^^FT0636,400^A0N,30,40^FH\^FDDEC :^FS" & vbNewLine


        Dim pos As Integer = 457
        For i As Integer = 0 To cb_rs.RecordCount - 1
            If j = 0 Then
                bcode = bcode & "^FO150," & pos & vbNewLine
                bcode = bcode & "^BCN,38,Y,N^FD>;" & cb_rs(0).Value & "^FS" & vbNewLine
                j = 1
            Else
                bcode = bcode & "^FO636," & pos & vbNewLine
                bcode = bcode & "^BCN,38,Y,N^FD>;" & cb_rs(0).Value & "^FS" & vbNewLine
                pos = pos + 77
                j = 0
            End If
            cb_rs.MoveNext()
        Next

        bcode = bcode & "^FO344,1634" & vbNewLine
        bcode = bcode & "^BY4^BCN,76,N,N^FD>;>800" & boxid & "^FS" & vbNewLine
        bcode = bcode & "^FT0344,1750^A0N,40,50^FH\^FD(00)" & boxid & "^FS" & vbNewLine
        bcode = bcode & "^FO81,55" & vbNewLine
        '       bcode = bcode & "^XGR:SSGFX000.GRF,1,1^FS" & vbNewLine
        bcode = bcode & "^XZ" & vbNewLine
        bcode = bcode & "^XA" & vbNewLine
        bcode = bcode & "^XFE:SSFMT000.ZPL^FS" & vbNewLine
        bcode = bcode & "^PQ1,0,1,Y" & vbNewLine
        bcode = bcode & "^XZ" & vbNewLine

        Microsoft.VisualBasic.FileSystem.Write(FN, bcode)
        Microsoft.VisualBasic.FileClose()

        Dim PORT As String = Query_RS("SELECT CN_TYPE FROM tbl_bcconfig WHERE MODEL = (SELECT TOP 1 MODEL FROM TBL_ESNMASTER WHERE INBOXID = '" & boxid & "') AND BC_TYPE = 'CARTON BOX'")
        If PORT = "" Then
            PORT = "LPT3"
        End If

        Shell("c:\windows\system32\cmd.exe /c print c:\CARTON.zpl /d:" & PORT)
        System.Threading.Thread.Sleep(50)
        Shell("c:\windows\system32\cmd.exe /c print c:\CARTON.zpl /d:" & PORT)


        'Shell("c:\windows\system32\cmd.exe /c print c:\CARTON.zpl /d:LPT3")
        'Shell("c:\windows\system32\cmd.exe /c print c:\CARTON.zpl /d:LPT3")

    End Function


    Function Print_Barcode_Carton_ship(ByVal boxid As String, ByVal model As String, ByVal prl As String, ByVal model1 As String, ByVal cnt As Integer) As Boolean

        Dim FN
        Dim bcode As String
        Dim cb_rs As ADODB.Recordset = Query_RS_ALL("select reserv5 from tbl_esnmaster_b where inboxid = '" & boxid & "' order by pos")
        Dim j As Integer = 0

        If cb_rs Is Nothing Then
            Exit Function
        End If

        FN = Microsoft.VisualBasic.FileSystem.FreeFile
        Microsoft.VisualBasic.FileOpen(FN, "C:\CARTON.ZPL", OpenMode.Output)

        bcode = "^XA" & vbNewLine
        bcode = bcode & "^SZ2^JMA" & vbNewLine
        bcode = bcode & "^MCY^PMN" & vbNewLine
        bcode = bcode & "^PW1353~JSN" & vbNewLine
        bcode = bcode & "^JZY" & vbNewLine
        bcode = bcode & "^LH0,0^LRN" & vbNewLine
        bcode = bcode & "^XZ" & vbNewLine
        '        bcode = bcode & "~DGR:SSGFX000.GRF,226100,133,:Z64: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:8CAF" & vbNewLine
        bcode = bcode & "^XA" & vbNewLine
        bcode = bcode & "^DFE:SSFMT000.ZPL^FS" & vbNewLine

        bcode = bcode & "^FT0100,100^A0N,30,60^FH\^FDSKU : " & model & "^FS" & vbNewLine

        bcode = bcode & "^FO725,86" & vbNewLine
        bcode = bcode & "^BY2^BCN,57,N,N^FD>:" & model & "^FS" & vbNewLine

        bcode = bcode & "^FT0725,70^A0N,30,28^FH\^FDSKU : " & model & "^FS" & vbNewLine

        bcode = bcode & "^FO154,219" & vbNewLine
        bcode = bcode & "^BY4^BCN,57,N,N^FD>;" & cnt & "^FS" & vbNewLine

        bcode = bcode & "^FT0154,200^A0N,30,28^FH\^FDCarton Qty : " & cnt & "^FS" & vbNewLine

        bcode = bcode & "^FO550,219" & vbNewLine
        bcode = bcode & "^BY3^BCN,57,N,N^FD>;" & Mid(prl, 1, Len(prl) - 1) & ">6" & Mid(prl, Len(prl), 1) & ", " & model1 & "^FS" & vbNewLine

        bcode = bcode & "^FT0550,200^A0N,30,28^FH\^FDPRL/SW Ver : " & prl & ", " & model1 & "^FS" & vbNewLine

        bcode = bcode & "^FO74,160" & vbNewLine
        bcode = bcode & "^GB1111,0,3^FS" & vbNewLine
        bcode = bcode & "^FO646,29" & vbNewLine
        bcode = bcode & "^GB0,133,3^FS" & vbNewLine
        bcode = bcode & "^FO69,293" & vbNewLine
        bcode = bcode & "^GB1121,0,3^FS" & vbNewLine
        bcode = bcode & "^FO494,161" & vbNewLine
        bcode = bcode & "^GB0,133,3^FS" & vbNewLine

        bcode = bcode & "^FT0350,350^A0N,50,100^FH\^FDReconditioned^FS" & vbNewLine
        bcode = bcode & "^FT0150,400^A0N,30,40^FH\^FDDEC :^FS" & vbNewLine
        bcode = bcode & "^^FT0636,400^A0N,30,40^FH\^FDDEC :^FS" & vbNewLine


        Dim pos As Integer = 457
        For i As Integer = 0 To cb_rs.RecordCount - 1
            If j = 0 Then
                bcode = bcode & "^FO150," & pos & vbNewLine
                bcode = bcode & "^BCN,38,Y,N^FD>;" & cb_rs(0).Value & "^FS" & vbNewLine
                j = 1
            Else
                bcode = bcode & "^FO636," & pos & vbNewLine
                bcode = bcode & "^BCN,38,Y,N^FD>;" & cb_rs(0).Value & "^FS" & vbNewLine
                pos = pos + 77
                j = 0
            End If
            cb_rs.MoveNext()
        Next

        bcode = bcode & "^FO344,1634" & vbNewLine
        bcode = bcode & "^BY4^BCN,76,N,N^FD>;>800" & boxid & "^FS" & vbNewLine
        bcode = bcode & "^FT0344,1750^A0N,40,50^FH\^FD(00)" & boxid & "^FS" & vbNewLine
        bcode = bcode & "^FO81,55" & vbNewLine
        '       bcode = bcode & "^XGR:SSGFX000.GRF,1,1^FS" & vbNewLine
        bcode = bcode & "^XZ" & vbNewLine
        bcode = bcode & "^XA" & vbNewLine
        bcode = bcode & "^XFE:SSFMT000.ZPL^FS" & vbNewLine
        bcode = bcode & "^PQ1,0,1,Y" & vbNewLine
        bcode = bcode & "^XZ" & vbNewLine

        Microsoft.VisualBasic.FileSystem.Write(FN, bcode)
        Microsoft.VisualBasic.FileClose()

        Dim PORT As String = Query_RS("SELECT CN_TYPE FROM tbl_bcconfig WHERE MODEL = (SELECT TOP 1 MODEL FROM TBL_ESNMASTER WHERE INBOXID = '" & boxid & "') AND BC_TYPE = 'CARTON BOX'")
        If PORT = "" Then
            PORT = "LPT3"
        End If

        Shell("c:\windows\system32\cmd.exe /c print c:\CARTON.zpl /d:" & PORT)
        System.Threading.Thread.Sleep(50)
        Shell("c:\windows\system32\cmd.exe /c print c:\CARTON.zpl /d:" & PORT)


        'Shell("c:\windows\system32\cmd.exe /c print c:\CARTON.zpl /d:LPT3")
        'Shell("c:\windows\system32\cmd.exe /c print c:\CARTON.zpl /d:LPT3")

    End Function






    Function Print_Barcode_Pallet(ByVal shipno As String, ByVal cnt As Integer, ByVal c_cnt As Integer, ByVal carton As String, ByVal MODEL As String) As Boolean

        Try
            Dim FN
            Dim bcode As String = ""

            Dim RS As ADODB.Recordset = Query_RS_ALL("SELECT AC_NM,ATTN,ADDRESS1,ADDRESS2,CITY+ ' ' +STATE+' '+ ZIPCODE,TELNO FROM TBL_ACINFO WHERE AC_NO = '" & Query_RS("SELECT TOP 1 CLAIM_NO FROM TBL_ESNMASTER_B WHERE INV_NO = '" & carton & "'") & "' AND AC_TP = 'SHIP'")

            If RS Is Nothing Then
                Modal_Error("No Shipping Information.")
                Exit Function
            End If

            FN = Microsoft.VisualBasic.FileSystem.FreeFile
            Microsoft.VisualBasic.FileOpen(FN, "C:\PALLET.ZPL", OpenMode.Output)



            bcode = bcode & "^XA" & vbNewLine
            bcode = bcode & "^SZ2^JMA" & vbNewLine
            bcode = bcode & "^MCY^PMN" & vbNewLine
            bcode = bcode & "^PW1353~JSN" & vbNewLine
            bcode = bcode & "^JZY" & vbNewLine
            bcode = bcode & "^LH0,0^LRN" & vbNewLine
            bcode = bcode & "^XZ" & vbNewLine
            bcode = bcode & "^XA" & vbNewLine
            bcode = bcode & "^DFE:SSFMT000.ZPL^FS" & vbNewLine
            bcode = bcode & "^FT78,148" & vbNewLine
            bcode = bcode & "^CI0" & vbNewLine
            bcode = bcode & "^A0N,42,58^FDShip To:^FS" & vbNewLine
            bcode = bcode & "^FT78,202" & vbNewLine
            bcode = bcode & "^A0N,34,47^FD" & RS(0).Value & "^FS" & vbNewLine
            bcode = bcode & "^FT78,252" & vbNewLine
            bcode = bcode & "^A0N,34,47^FD" & RS(2).Value & "^FS" & vbNewLine
            bcode = bcode & "^FT78,301" & vbNewLine
            bcode = bcode & "^A0N,34,47^FD" & RS(3).Value & "^FS" & vbNewLine
            bcode = bcode & "^FT78,351" & vbNewLine
            bcode = bcode & "^A0N,34,47^FD" & RS(1).Value & "^FS" & vbNewLine
            bcode = bcode & "^FT78,400" & vbNewLine
            bcode = bcode & "^A0N,34,47^FD" & RS(4).Value & "^FS" & vbNewLine
            bcode = bcode & "^FT743,148" & vbNewLine
            bcode = bcode & "^A0N,42,58^FDShip From:^FS" & vbNewLine
            bcode = bcode & "^FT743,203" & vbNewLine
            bcode = bcode & "^A0N,34,47^FDExodus Wireless Corp.^FS" & vbNewLine
            bcode = bcode & "^FT743,252" & vbNewLine
            bcode = bcode & "^A0N,34,47^FD14352 Chambers Road^FS" & vbNewLine
            bcode = bcode & "^FT743,302" & vbNewLine
            bcode = bcode & "^A0N,34,47^FDTustin, CA 92780^FS" & vbNewLine
            bcode = bcode & "^FT173,606" & vbNewLine
            bcode = bcode & "^A0N,51,70^FDSKU :^FS" & vbNewLine
            bcode = bcode & "^FT401,609" & vbNewLine
            bcode = bcode & "^A0N,59,81^FD" & MODEL & "^FS" & vbNewLine
            bcode = bcode & "^FT401,744" & vbNewLine
            bcode = bcode & "^A0N,59,81^FDRECONDITIONED^FS" & vbNewLine
            bcode = bcode & "^FT173,884" & vbNewLine
            bcode = bcode & "^A0N,51,70^FDPO :^FS" & vbNewLine
            bcode = bcode & "^FT401,900" & vbNewLine
            bcode = bcode & "^A0N,93,128^FD" & shipno & "^FS" & vbNewLine
            '        bcode = bcode & "^A0N,93,128^FD" & shipno & "^FS" & vbNewLine
            bcode = bcode & "^FT173,1019" & vbNewLine
            bcode = bcode & "^A0N,51,70^FDItem Count : ^FS" & vbNewLine
            bcode = bcode & "^FT173,1155" & vbNewLine
            bcode = bcode & "^A0N,51,70^FDTotal Cartons :^FS" & vbNewLine
            bcode = bcode & "^FT173,1290" & vbNewLine
            bcode = bcode & "^A0N,51,70^FDShip Date :^FS" & vbNewLine
            bcode = bcode & "^FO173,1501" & vbNewLine
            bcode = bcode & "^BY6^BCN,152,N,N^FD>;>800" & carton & "^FS" & vbNewLine
            bcode = bcode & "^FT173,1699" & vbNewLine
            bcode = bcode & "^A0N,51,70^FD(00)" & carton & "^FS" & vbNewLine
            bcode = bcode & "^FT612,1023" & vbNewLine
            bcode = bcode & "^A0N,59,81^FD" & cnt & "^FS" & vbNewLine
            bcode = bcode & "^FT651,1158" & vbNewLine
            bcode = bcode & "^A0N,59,81^FD" & c_cnt & "^FS" & vbNewLine
            bcode = bcode & "^FT515,1293" & vbNewLine
            bcode = bcode & "^A0N,59,81^FD" & Query_RS("Select REPLACE(CONVERT(VARCHAR(10), GETDATE(),101), '.', '/')") & "^FS" & vbNewLine
            bcode = bcode & "^XZ" & vbNewLine
            bcode = bcode & "^XA" & vbNewLine
            bcode = bcode & "^XFE:SSFMT000.ZPL^FS" & vbNewLine
            bcode = bcode & "^PQ1,0,1,Y" & vbNewLine
            bcode = bcode & "^XZ" & vbNewLine


            Microsoft.VisualBasic.FileSystem.Write(FN, bcode)
            Microsoft.VisualBasic.FileClose()

            Dim PORT As String = Query_RS("SELECT CN_TYPE FROM tbl_bcconfig WHERE MODEL = (SELECT TOP 1 MODEL FROM TBL_ESNMASTER_B WHERE INV_NO = '" & carton & "') AND BC_TYPE = 'PALLET'")
            If PORT = "" Then
                PORT = "LPT3"
            End If


            Shell("c:\windows\system32\cmd.exe /c print c:\PALLET.zpl /d:" & PORT)
            System.Threading.Thread.Sleep(50)
            Shell("c:\windows\system32\cmd.exe /c print c:\PALLET.zpl /d:" & PORT)
            System.Threading.Thread.Sleep(50)
            Shell("c:\windows\system32\cmd.exe /c print c:\PALLET.zpl /d:" & PORT)
            System.Threading.Thread.Sleep(50)
            Shell("c:\windows\system32\cmd.exe /c print c:\PALLET.zpl /d:" & PORT)
            System.Threading.Thread.Sleep(50)
            Shell("c:\windows\system32\cmd.exe /c print c:\PALLET.zpl /d:" & PORT)

        Catch ex As Exception
            Modal_Error(ex.Message)
        End Try



        'Shell("c:\windows\system32\cmd.exe /c print c:\PALLET.zpl /d:LPT3")
        'Shell("c:\windows\system32\cmd.exe /c print c:\PALLET.zpl /d:LPT3")
        'Shell("c:\windows\system32\cmd.exe /c print c:\PALLET.zpl /d:LPT3")
        'Shell("c:\windows\system32\cmd.exe /c print c:\PALLET.zpl /d:LPT3")
        'Shell("c:\windows\system32\cmd.exe /c print c:\PALLET.zpl /d:LPT3")

    End Function

    Function Print_Barcode_Label(ByVal dec As String, ByVal hex As String, ByVal model As String) As Boolean

        Dim FN
        Dim bcode As String

        Dim PORT As String = Query_RS("SELECT CN_TYPE FROM tbl_bcconfig WHERE MODEL = '" & model & "' AND BC_TYPE = 'LABEL'")
        If PORT = "" Then
            PORT = "COM3"
        End If

        FN = Microsoft.VisualBasic.FileSystem.FreeFile
        Microsoft.VisualBasic.FileOpen(FN, "C:\label.ZPL", OpenMode.Output)


        'bcode = ""
        'bcode = bcode & "^XA" & vbNewLine
        'bcode = bcode & "^SZ2^JMA" & vbNewLine
        'bcode = bcode & "^MCY^PMN" & vbNewLine
        'bcode = bcode & "^PW471~JSN" & vbNewLine
        'bcode = bcode & "^JZY" & vbNewLine
        'bcode = bcode & "^LH0,0^LRN" & vbNewLine
        'bcode = bcode & "^XZ" & vbNewLine
        ''        bcode = bcode & "~DGR:SSGFX000.GRF,1036,14,:Z64:eJztzzFqw0AQBdAPAU+3ewGDrqFAwIfIBXyEBTVytTIuXASsCxhyDZebyodwsymC3QTkbgqhzSxaOZLTpvTvXjH8P5rqtuhgP0LFgK5FAfYzbNpBlb2GbRe1Q7GtzNVRnbRxZuWI4NWbaO3Ms1Mz0UnuosIOd8K9bndTzUSvQ4PWomLcrorRMt8rre7Vwq7jRx6PPPJPacpwXHhqM9YVwCxqqMt44ZKYQsbWR3nKS9JzlE3SC9EexihcDqI5Pe3hlwrfh3CMemePqfBHt7uplsBXPjQ00nfOx+3n/HeZ6pVWJzni+JHCDxCoxyg=:93D6" & vbNewLine
        'bcode = bcode & "^XA" & vbNewLine
        'bcode = bcode & "^FO32,17" & vbNewLine
        ''       bcode = bcode & "^XGR:SSGFX000.GRF,1,1^FS" & vbNewLine
        'bcode = bcode & "^ISR:SS_TEMP.GRF,N^XZ" & vbNewLine
        ''bcode = bcode & "~DGR:SSGFX001.GRF,2280,30,:Z64:eJztlbFu2zAQho/gwCUV0bEALQ19gqBDaESBX0mqBzuAEbPokKVI5gBC8xoZaXjQlryAUdDw4M2h4SEeVKl3lAT4FQr4BoHffYQo3j+IuVoY5qbcQJYyy5wWAJmSwG3NrPC1tNxXZO8IZxJXKjHCtTahhzDga94i+DT0uJXaKieGO2mY3wrE3cAyP7uxMrPCRkOrXTQsE8P2JWGpHfMV9YItNESsHBm+t7K1fB9sIS0M62sTrctvhuerT4hk84JsjXhbj6zcl1+NyOsBIlqRq2AR8zey79xEeT1ERBtl6sbJ2zdEL3A1bbi5yN8vLezRXuB9nTiKSxyOwM0arcrL3LI1WtVZRBvshJvkbzn3bIE2qVTk+IuYe04YadFZ3llw+PVzL36UYRqdDYg29BDZbuDE7Sul8I7HAE4SUxjaKN8xK49tCmTrDikj6hHOThMk7BOcEerT9An79DWzcK5znetc/0nF0k0hFkaD4FZBzE0GglkPqXQK5tJVMJIWrSAU9oh2cYS71q4r2jIFKV3aouSL0Evh8Iu9wvi3O0ISZTM4PKxXkIjlCj4uYAabgj/gI0erMsLxAyRyST2yT0sDTo2PMFJQtXaUkH2yFfDnzQvs1FjD1cCugJO9+kL22a1APG8qaNR3DZMEjxRkJ5+DJVQHtLHXsE3wKKHQbtnyD3wUhwIkjHH1iLaMDwok2ZItapibMQ4HEDdoWakmiGhZZxFVsPeZZs1q0oAi2yzsDq4NYkqv2pjepp2N8OsRp8EWvZ22NkwDMYvXf2BSOEwh8gqyiO7BlgVsY0yheWxTQCt7xIxCD/HeniaI2CcoMO7mpzlNH7FPH//X8A+HE71A:11D4" & vbNewLine
        'bcode = bcode & "^XA" & vbNewLine
        'bcode = bcode & "^ILR:SS_TEMP.GRF^FS" & vbNewLine
        'bcode = bcode & "^FO75,37" & vbNewLine
        'bcode = bcode & "^BY2^BCN,24,N,N^FD>;" & dec & "^FS" & vbNewLine
        'bcode = bcode & "^FT045,30^A0N,20,20^FH\^FDMEID DEC : " & Mid(dec, 1, 3) & " " & Mid(dec, 4, 3) & " " & Mid(dec, 7, 3) & " " & Mid(dec, 10, 3) & " " & Mid(dec, 13, 3) & " " & Mid(dec, 16, 3) & " " & "^FS" & vbNewLine

        'bcode = bcode & "^FO70,95" & vbNewLine
        'bcode = bcode & "^BCN,24,N,N^FD>;" & dec & "^FS" & vbNewLine
        'bcode = bcode & "^FT045,85^A0N,20,20^FH\^FDMEID DEC : " & Mid(dec, 1, 3) & " " & Mid(dec, 4, 3) & " " & Mid(dec, 7, 3) & " " & Mid(dec, 10, 3) & " " & Mid(dec, 13, 3) & " " & Mid(dec, 16, 3) & " " & "^FS" & vbNewLine

        'bcode = bcode & "^FO148,15" & vbNewLine
        ''bcode = bcode & "^XGR:SSGFX001.GRF,1,1^FS" & vbNewLine
        'bcode = bcode & "^PQ1,0,1,Y" & vbNewLine
        'bcode = bcode & "^XZ" & vbNewLine
        'bcode = bcode & "^XA" & vbNewLine

        'bcode = bcode & "^IDR:SSGFX000.GRF^XZ" & vbNewLine
        'bcode = bcode & "^XA" & vbNewLine
        'bcode = bcode & "^IDR:SSGFX001.GRF^XZ" & vbNewLine
        'bcode = bcode & "^XA" & vbNewLine
        'bcode = bcode & "^IDR:SS_TEMP.GRF^XZ" & vbNewLine



        bcode = ""
        bcode = bcode & "^XA" & vbNewLine
        bcode = bcode & "^SZ2^JMA" & vbNewLine
        bcode = bcode & "^MCY^PMN" & vbNewLine
        bcode = bcode & "^PW471~JSN" & vbNewLine
        bcode = bcode & "^JZY" & vbNewLine
        bcode = bcode & "^LH0,0^LRN" & vbNewLine
        bcode = bcode & "^XZ" & vbNewLine
        '        bcode = bcode & "~DGR:SSGFX000.GRF,1036,14,:Z64:eJztzzFqw0AQBdAPAU+3ewGDrqFAwIfIBXyEBTVytTIuXASsCxhyDZebyodwsymC3QTkbgqhzSxaOZLTpvTvXjH8P5rqtuhgP0LFgK5FAfYzbNpBlb2GbRe1Q7GtzNVRnbRxZuWI4NWbaO3Ms1Mz0UnuosIOd8K9bndTzUSvQ4PWomLcrorRMt8rre7Vwq7jRx6PPPJPacpwXHhqM9YVwCxqqMt44ZKYQsbWR3nKS9JzlE3SC9EexihcDqI5Pe3hlwrfh3CMemePqfBHt7uplsBXPjQ00nfOx+3n/HeZ6pVWJzni+JHCDxCoxyg=:93D6" & vbNewLine
        bcode = bcode & "^XA" & vbNewLine
        bcode = bcode & "^FO32,17" & vbNewLine
        '       bcode = bcode & "^XGR:SSGFX000.GRF,1,1^FS" & vbNewLine
        bcode = bcode & "^ISR:SS_TEMP.GRF,N^XZ" & vbNewLine
        'bcode = bcode & "~DGR:SSGFX001.GRF,2280,30,:Z64:eJztlbFu2zAQho/gwCUV0bEALQ19gqBDaESBX0mqBzuAEbPokKVI5gBC8xoZaXjQlryAUdDw4M2h4SEeVKl3lAT4FQr4BoHffYQo3j+IuVoY5qbcQJYyy5wWAJmSwG3NrPC1tNxXZO8IZxJXKjHCtTahhzDga94i+DT0uJXaKieGO2mY3wrE3cAyP7uxMrPCRkOrXTQsE8P2JWGpHfMV9YItNESsHBm+t7K1fB9sIS0M62sTrctvhuerT4hk84JsjXhbj6zcl1+NyOsBIlqRq2AR8zey79xEeT1ERBtl6sbJ2zdEL3A1bbi5yN8vLezRXuB9nTiKSxyOwM0arcrL3LI1WtVZRBvshJvkbzn3bIE2qVTk+IuYe04YadFZ3llw+PVzL36UYRqdDYg29BDZbuDE7Sul8I7HAE4SUxjaKN8xK49tCmTrDikj6hHOThMk7BOcEerT9An79DWzcK5znetc/0nF0k0hFkaD4FZBzE0GglkPqXQK5tJVMJIWrSAU9oh2cYS71q4r2jIFKV3aouSL0Evh8Iu9wvi3O0ISZTM4PKxXkIjlCj4uYAabgj/gI0erMsLxAyRyST2yT0sDTo2PMFJQtXaUkH2yFfDnzQvs1FjD1cCugJO9+kL22a1APG8qaNR3DZMEjxRkJ5+DJVQHtLHXsE3wKKHQbtnyD3wUhwIkjHH1iLaMDwok2ZItapibMQ4HEDdoWakmiGhZZxFVsPeZZs1q0oAi2yzsDq4NYkqv2pjepp2N8OsRp8EWvZ22NkwDMYvXf2BSOEwh8gqyiO7BlgVsY0yheWxTQCt7xIxCD/HeniaI2CcoMO7mpzlNH7FPH//X8A+HE71A:11D4" & vbNewLine
        bcode = bcode & "^XA^LRY" & vbNewLine
        bcode = bcode & "^ILR:SS_TEMP.GRF^FS" & vbNewLine
        bcode = bcode & "^FO30,20^GB340,24,20,B^FS" & vbNewLine
        bcode = bcode & "^FO65,20" & vbNewLine
        bcode = bcode & "^BCN,24,N,N^FD>;" & dec & "^FS" & vbNewLine
        bcode = bcode & "^FT030,15^A0N,20,20^FH\^FDMEID DEC : " & Mid(dec, 1, 3) & "  " & Mid(dec, 4, 3) & "  " & Mid(dec, 7, 3) & "  " & Mid(dec, 10, 3) & "  " & Mid(dec, 13, 3) & "  " & Mid(dec, 16, 3) & " " & "^FS" & vbNewLine

        bcode = bcode & "^FO30,78^GB340,24,20,B^FS" & vbNewLine
        bcode = bcode & "^FO65,78" & vbNewLine
        bcode = bcode & "^BCN,24,N,N^FD>;" & dec & "^FS" & vbNewLine
        bcode = bcode & "^FT030,73^A0N,20,20^FH\^FDMEID DEC : " & Mid(dec, 1, 3) & "  " & Mid(dec, 4, 3) & "  " & Mid(dec, 7, 3) & "  " & Mid(dec, 10, 3) & "  " & Mid(dec, 13, 3) & "  " & Mid(dec, 16, 3) & " " & "^FS" & vbNewLine

        bcode = bcode & "^FO148,15" & vbNewLine
        'bcode = bcode & "^XGR:SSGFX001.GRF,1,1^FS" & vbNewLine
        bcode = bcode & "^PQ1,0,1,Y" & vbNewLine
        bcode = bcode & "^XZ" & vbNewLine
        bcode = bcode & "^XA" & vbNewLine

        bcode = bcode & "^IDR:SSGFX000.GRF^XZ" & vbNewLine
        bcode = bcode & "^XA" & vbNewLine
        bcode = bcode & "^IDR:SSGFX001.GRF^XZ" & vbNewLine
        bcode = bcode & "^XA" & vbNewLine
        bcode = bcode & "^IDR:SS_TEMP.GRF^XZ" & vbNewLine


        'Microsoft.VisualBasic.FileSystem.Write(FN, bcode)
        'Microsoft.VisualBasic.FileClose()
        'Shell("c:\windows\system32\cmd.exe /c print c:\label.zpl /d:" & PORT)
        ''            Shell("c:\windows\system32\cmd.exe /c print c:\label.zpl /d:LPT3")


        If model = "LS980SPRWH" Then
            'bcode = ""
            'bcode = bcode & "^XA" & vbNewLine
            'bcode = bcode & "^SZ2^JMA" & vbNewLine
            'bcode = bcode & "^MCY^PMN" & vbNewLine
            'bcode = bcode & "^PW471~JSN" & vbNewLine
            'bcode = bcode & "^JZY" & vbNewLine
            'bcode = bcode & "^LH0,0^LRN" & vbNewLine
            'bcode = bcode & "^XZ" & vbNewLine
            ''        bcode = bcode & "~DGR:SSGFX000.GRF,1036,14,:Z64:eJztzzFqw0AQBdAPAU+3ewGDrqFAwIfIBXyEBTVytTIuXASsCxhyDZebyodwsymC3QTkbgqhzSxaOZLTpvTvXjH8P5rqtuhgP0LFgK5FAfYzbNpBlb2GbRe1Q7GtzNVRnbRxZuWI4NWbaO3Ms1Mz0UnuosIOd8K9bndTzUSvQ4PWomLcrorRMt8rre7Vwq7jRx6PPPJPacpwXHhqM9YVwCxqqMt44ZKYQsbWR3nKS9JzlE3SC9EexihcDqI5Pe3hlwrfh3CMemePqfBHt7uplsBXPjQ00nfOx+3n/HeZ6pVWJzni+JHCDxCoxyg=:93D6" & vbNewLine
            'bcode = bcode & "^XA" & vbNewLine
            'bcode = bcode & "^FO32,17" & vbNewLine
            ''       bcode = bcode & "^XGR:SSGFX000.GRF,1,1^FS" & vbNewLine
            'bcode = bcode & "^ISR:SS_TEMP.GRF,N^XZ" & vbNewLine
            ''bcode = bcode & "~DGR:SSGFX001.GRF,2280,30,:Z64:eJztlbFu2zAQho/gwCUV0bEALQ19gqBDaESBX0mqBzuAEbPokKVI5gBC8xoZaXjQlryAUdDw4M2h4SEeVKl3lAT4FQr4BoHffYQo3j+IuVoY5qbcQJYyy5wWAJmSwG3NrPC1tNxXZO8IZxJXKjHCtTahhzDga94i+DT0uJXaKieGO2mY3wrE3cAyP7uxMrPCRkOrXTQsE8P2JWGpHfMV9YItNESsHBm+t7K1fB9sIS0M62sTrctvhuerT4hk84JsjXhbj6zcl1+NyOsBIlqRq2AR8zey79xEeT1ERBtl6sbJ2zdEL3A1bbi5yN8vLezRXuB9nTiKSxyOwM0arcrL3LI1WtVZRBvshJvkbzn3bIE2qVTk+IuYe04YadFZ3llw+PVzL36UYRqdDYg29BDZbuDE7Sul8I7HAE4SUxjaKN8xK49tCmTrDikj6hHOThMk7BOcEerT9An79DWzcK5znetc/0nF0k0hFkaD4FZBzE0GglkPqXQK5tJVMJIWrSAU9oh2cYS71q4r2jIFKV3aouSL0Evh8Iu9wvi3O0ISZTM4PKxXkIjlCj4uYAabgj/gI0erMsLxAyRyST2yT0sDTo2PMFJQtXaUkH2yFfDnzQvs1FjD1cCugJO9+kL22a1APG8qaNR3DZMEjxRkJ5+DJVQHtLHXsE3wKKHQbtnyD3wUhwIkjHH1iLaMDwok2ZItapibMQ4HEDdoWakmiGhZZxFVsPeZZs1q0oAi2yzsDq4NYkqv2pjepp2N8OsRp8EWvZ22NkwDMYvXf2BSOEwh8gqyiO7BlgVsY0yheWxTQCt7xIxCD/HeniaI2CcoMO7mpzlNH7FPH//X8A+HE71A:11D4" & vbNewLine
            'bcode = bcode & "^XA^LRY" & vbNewLine
            'bcode = bcode & "^ILR:SS_TEMP.GRF^FS" & vbNewLine
            'bcode = bcode & "^FO30,23^GB340,24,20,B^FS" & vbNewLine
            'bcode = bcode & "^FO65,23" & vbNewLine
            'bcode = bcode & "^BCN,24,N,N^FD>;" & dec & "^FS" & vbNewLine
            'bcode = bcode & "^FT030,16^A0N,20,20^FH\^FDMEID DEC : " & Mid(dec, 1, 3) & "  " & Mid(dec, 4, 3) & "  " & Mid(dec, 7, 3) & "  " & Mid(dec, 10, 3) & "  " & Mid(dec, 13, 3) & "  " & Mid(dec, 16, 3) & " " & "^FS" & vbNewLine

            'bcode = bcode & "^FO30,81^GB340,24,20,B^FS" & vbNewLine
            'bcode = bcode & "^FO65,81" & vbNewLine
            'bcode = bcode & "^BCN,24,N,N^FD>;" & dec & "^FS" & vbNewLine
            'bcode = bcode & "^FT030,71^A0N,20,20^FH\^FDMEID DEC : " & Mid(dec, 1, 3) & "  " & Mid(dec, 4, 3) & "  " & Mid(dec, 7, 3) & "  " & Mid(dec, 10, 3) & "  " & Mid(dec, 13, 3) & "  " & Mid(dec, 16, 3) & " " & "^FS" & vbNewLine

            'bcode = bcode & "^FO148,15" & vbNewLine
            ''bcode = bcode & "^XGR:SSGFX001.GRF,1,1^FS" & vbNewLine
            'bcode = bcode & "^PQ1,0,1,Y" & vbNewLine
            'bcode = bcode & "^XZ" & vbNewLine
            'bcode = bcode & "^XA" & vbNewLine

            'bcode = bcode & "^IDR:SSGFX000.GRF^XZ" & vbNewLine
            'bcode = bcode & "^XA" & vbNewLine
            'bcode = bcode & "^IDR:SSGFX001.GRF^XZ" & vbNewLine
            'bcode = bcode & "^XA" & vbNewLine
            'bcode = bcode & "^IDR:SS_TEMP.GRF^XZ" & vbNewLine


            bcode = ""
            bcode = bcode & "^XA" & vbNewLine
            bcode = bcode & "^SZ2^JMA" & vbNewLine
            bcode = bcode & "^MCY^PMN" & vbNewLine
            bcode = bcode & "^PW471~JSN" & vbNewLine
            bcode = bcode & "^JZY" & vbNewLine
            bcode = bcode & "^LH0,0^LRN" & vbNewLine
            bcode = bcode & "^XZ" & vbNewLine
            '        bcode = bcode & "~DGR:SSGFX000.GRF,1036,14,:Z64:eJztzzFqw0AQBdAPAU+3ewGDrqFAwIfIBXyEBTVytTIuXASsCxhyDZebyodwsymC3QTkbgqhzSxaOZLTpvTvXjH8P5rqtuhgP0LFgK5FAfYzbNpBlb2GbRe1Q7GtzNVRnbRxZuWI4NWbaO3Ms1Mz0UnuosIOd8K9bndTzUSvQ4PWomLcrorRMt8rre7Vwq7jRx6PPPJPacpwXHhqM9YVwCxqqMt44ZKYQsbWR3nKS9JzlE3SC9EexihcDqI5Pe3hlwrfh3CMemePqfBHt7uplsBXPjQ00nfOx+3n/HeZ6pVWJzni+JHCDxCoxyg=:93D6" & vbNewLine
            bcode = bcode & "^XA" & vbNewLine
            bcode = bcode & "^FO32,17" & vbNewLine
            '       bcode = bcode & "^XGR:SSGFX000.GRF,1,1^FS" & vbNewLine
            bcode = bcode & "^ISR:SS_TEMP.GRF,N^XZ" & vbNewLine
            'bcode = bcode & "~DGR:SSGFX001.GRF,2280,30,:Z64:eJztlbFu2zAQho/gwCUV0bEALQ19gqBDaESBX0mqBzuAEbPokKVI5gBC8xoZaXjQlryAUdDw4M2h4SEeVKl3lAT4FQr4BoHffYQo3j+IuVoY5qbcQJYyy5wWAJmSwG3NrPC1tNxXZO8IZxJXKjHCtTahhzDga94i+DT0uJXaKieGO2mY3wrE3cAyP7uxMrPCRkOrXTQsE8P2JWGpHfMV9YItNESsHBm+t7K1fB9sIS0M62sTrctvhuerT4hk84JsjXhbj6zcl1+NyOsBIlqRq2AR8zey79xEeT1ERBtl6sbJ2zdEL3A1bbi5yN8vLezRXuB9nTiKSxyOwM0arcrL3LI1WtVZRBvshJvkbzn3bIE2qVTk+IuYe04YadFZ3llw+PVzL36UYRqdDYg29BDZbuDE7Sul8I7HAE4SUxjaKN8xK49tCmTrDikj6hHOThMk7BOcEerT9An79DWzcK5znetc/0nF0k0hFkaD4FZBzE0GglkPqXQK5tJVMJIWrSAU9oh2cYS71q4r2jIFKV3aouSL0Evh8Iu9wvi3O0ISZTM4PKxXkIjlCj4uYAabgj/gI0erMsLxAyRyST2yT0sDTo2PMFJQtXaUkH2yFfDnzQvs1FjD1cCugJO9+kL22a1APG8qaNR3DZMEjxRkJ5+DJVQHtLHXsE3wKKHQbtnyD3wUhwIkjHH1iLaMDwok2ZItapibMQ4HEDdoWakmiGhZZxFVsPeZZs1q0oAi2yzsDq4NYkqv2pjepp2N8OsRp8EWvZ22NkwDMYvXf2BSOEwh8gqyiO7BlgVsY0yheWxTQCt7xIxCD/HeniaI2CcoMO7mpzlNH7FPH//X8A+HE71A:11D4" & vbNewLine
            bcode = bcode & "^XA" & vbNewLine
            bcode = bcode & "^ILR:SS_TEMP.GRF^FS" & vbNewLine
            bcode = bcode & "^FO78,37" & vbNewLine
            bcode = bcode & "^BY2^BCN,24,N,N^FD>;" & dec & "^FS" & vbNewLine
            bcode = bcode & "^FT045,34^A0N,20,20^FH\^FDMEID DEC : " & Mid(dec, 1, 3) & "  " & Mid(dec, 4, 3) & "  " & Mid(dec, 7, 3) & "  " & Mid(dec, 10, 3) & "  " & Mid(dec, 13, 3) & "  " & Mid(dec, 16, 3) & " " & "^FS" & vbNewLine

            bcode = bcode & "^FO78,97" & vbNewLine
            bcode = bcode & "^BCN,24,N,N^FD>;" & dec & "^FS" & vbNewLine
            bcode = bcode & "^FT045,92^A0N,20,20^FH\^FDMEID DEC : " & Mid(dec, 1, 3) & "  " & Mid(dec, 4, 3) & "  " & Mid(dec, 7, 3) & "  " & Mid(dec, 10, 3) & "  " & Mid(dec, 13, 3) & "  " & Mid(dec, 16, 3) & " " & "^FS" & vbNewLine

            bcode = bcode & "^FO148,15" & vbNewLine
            'bcode = bcode & "^XGR:SSGFX001.GRF,1,1^FS" & vbNewLine
            bcode = bcode & "^PQ1,0,1,Y" & vbNewLine
            bcode = bcode & "^XZ" & vbNewLine
            bcode = bcode & "^XA" & vbNewLine

            bcode = bcode & "^IDR:SSGFX000.GRF^XZ" & vbNewLine
            bcode = bcode & "^XA" & vbNewLine
            bcode = bcode & "^IDR:SSGFX001.GRF^XZ" & vbNewLine
            bcode = bcode & "^XA" & vbNewLine
            bcode = bcode & "^IDR:SS_TEMP.GRF^XZ" & vbNewLine



            Microsoft.VisualBasic.FileSystem.Write(FN, bcode)
            Microsoft.VisualBasic.FileClose()
            Shell("c:\windows\system32\cmd.exe /c print c:\label.zpl /d:" & PORT)
            '            Shell("c:\windows\system32\cmd.exe /c print c:\label.zpl /d:LPT3")

        ElseIf model = "LG870BSTBK" Then

            bcode = ""

            bcode = bcode & "CT~~CD,~CC^~CT~" & vbNewLine
            bcode = bcode & "^XA~TA000~JSN^LT0^MNW^MTT^PON^PMN^LH0,0^JMA^PR2,2~SD25^JUS^LRN^CI0^XZ" & vbNewLine
            bcode = bcode & "^XA" & vbNewLine
            bcode = bcode & "^MMT" & vbNewLine
            bcode = bcode & "^PW526" & vbNewLine
            bcode = bcode & "^LL0300" & vbNewLine
            bcode = bcode & "^LS0" & vbNewLine
            bcode = bcode & "^FT187,275^ACI,26,13^FH\^FDLG870^FS" & vbNewLine
            bcode = bcode & "^FT294,253^ACI,26,13^FH\^FDFCC ID:^FS" & vbNewLine
            bcode = bcode & "^FT205,253^ACI,26,13^FH\^FDZNFLG870^FS" & vbNewLine
            bcode = bcode & "^FT495,135^ACI,26,13^FH\^FDMEID^FS" & vbNewLine
            bcode = bcode & "^FT495,222^ACI,26,13^FH\^FDMEID^FS" & vbNewLine
            bcode = bcode & "^BY2,1,40^FT440,200^BCI,,N,N" & vbNewLine
            bcode = bcode & "^FD>;" & dec & "^FS" & vbNewLine
            bcode = bcode & "^FT480,160^A0I,35,35^FH\^FD" & Mid(dec, 1, 3) & "^FS" & vbNewLine
            bcode = bcode & "^FT420,160^A0I,35,35^FH\^FD" & Mid(dec, 4, 3) & "^FS" & vbNewLine
            bcode = bcode & "^FT360,160^A0I,35,35^FH\^FD" & Mid(dec, 7, 3) & "^FS" & vbNewLine
            bcode = bcode & "^FT300,160^A0I,35,35^FH\^FD" & Mid(dec, 10, 3) & "^FS" & vbNewLine
            bcode = bcode & "^FT240,160^A0I,35,35^FH\^FD" & Mid(dec, 13, 3) & "^FS" & vbNewLine
            bcode = bcode & "^FT180,160^A0I,35,35^FH\^FD" & Mid(dec, 16, 3) & "^FS" & vbNewLine
            bcode = bcode & "^BY2,3,41^FT440,110^BCI,,N,N" & vbNewLine
            bcode = bcode & "^FD>;" & hex & "^FS" & vbNewLine
            bcode = bcode & "^FT460,70^A0I,35,35^FH\^FD" & Mid(hex, 1, 3) & "^FS" & vbNewLine
            bcode = bcode & "^FT400,70^A0I,35,35^FH\^FD" & Mid(hex, 4, 3) & "^FS" & vbNewLine
            bcode = bcode & "^FT340,70^A0I,35,35^FH\^FD" & Mid(hex, 7, 3) & "^FS" & vbNewLine
            bcode = bcode & "^FT280,70^A0I,35,35^FH\^FD" & Mid(hex, 10, 3) & "^FS" & vbNewLine
            bcode = bcode & "^FT450,47^ACI,18,10^FH\^FDS/N:" & Query_RS("SELECT TOP 1 ISNULL(SERIAL_NO,'') FROM TBL_SERIAL WHERE MEID = '" & hex & "' ORDER BY C_DATE DESC") & "^FS" & vbNewLine
            bcode = bcode & "^FT220,70^A0I,35,35^FH\^FD" & Mid(hex, 13, 2) & "^FS" & vbNewLine
            bcode = bcode & "^FT254,24^ACI,26,13^FH\^FDMADE IN KOREA^FS" & vbNewLine
            bcode = bcode & "^FT495,205^ACI,26,13^FH\^FDDEC^FS" & vbNewLine
            bcode = bcode & "^FT495,115^ACI,26,13^FH\^FDHEX^FS" & vbNewLine
            bcode = bcode & "^FT294,275^ADI,26,13^FH\^FDMODEL:^FS" & vbNewLine
            bcode = bcode & "^PQ1,0,1,Y^XZ" & vbNewLine

            Microsoft.VisualBasic.FileSystem.Write(FN, bcode)
            Microsoft.VisualBasic.FileClose()

            Shell("c:\windows\system32\cmd.exe /c print c:\label.zpl /d:" & PORT)
            '           Shell("c:\windows\system32\cmd.exe /c print c:\label.zpl /d:COM3")
        Else
            Microsoft.VisualBasic.FileSystem.Write(FN, bcode)
            Microsoft.VisualBasic.FileClose()
            Shell("c:\windows\system32\cmd.exe /c print c:\label.zpl /d:" & PORT)
            '            Shell("c:\windows\system32\cmd.exe /c print c:\label.zpl /d:COM3")
        End If


    End Function



    Function Print_NewBarcode_NoTrg(ByVal esn As String, ByVal model As String, ByVal def As String, ByVal dt As String) As Boolean

        Dim FN
        Dim bcode As String

        Dim CUS, Color As String
        Dim ba_rs As New ADODB.Recordset
        ba_rs = Query_RS_ALL("SELECT SW_VER, color FROM TBL_MODELMASTER WHERE MODEL_NO = (SELECT MODEL FROM TBL_FESNMASTER WHERE ESN = '" & esn & "')")
        CUS = ba_rs(0).Value
        Color = ba_rs(1).Value
        ba_rs = Nothing

        Dim ARR As String() = GET_TRG(esn, "FLIP")
        'ba_rs = Query_RS_ALL("SELECT def_cd FROM TBL_triage WHERE esn_obid = (SELECT obid FROM TBL_FESNMASTER WHERE ESN = '" & esn & "') order by c_date")

        Dim def1 As String = ARR(4)
        Dim def2 As String = ARR(3)
        Dim def3 As String = ARR(2)
        Dim def4 As String = ARR(1)
        Dim def5 As String = ARR(0)

        FN = Microsoft.VisualBasic.FileSystem.FreeFile
        Microsoft.VisualBasic.FileOpen(FN, "C:\ESN.ZPL", OpenMode.Output)
        '        Microsoft.VisualBasic.FileSystem.Write(FN, "/ ^XA" & vbNewLine & "/ ^LT0" & vbNewLine & "/ ^LH40,0" & vbNewLine & "/ ^FO1,100^A0N,13,25^FD" & model & "-" & def & "^FS" & vbNewLine & "/ ^FO1,50^BY1" & vbNewLine & "/ ^B3N,N,30,Y,N" & vbNewLine & "/ ^FD" & esn & "^FS" & vbNewLine & "/ ^XZ")

        If Site_id = "S1000" Then

            bcode = "^XA~TA000~JSN^LT0^MMT^MNW^MTT^PON^PMN^LH0,0^JMA^PR2,2^MD10^JUS^LRN^CI0^XZ" & vbNewLine
            bcode = bcode & "^XA^LL0150" & vbNewLine
            bcode = bcode & "^PW510" & vbNewLine
            bcode = bcode & "^BY2,3,34^FT47,145^BCN,,N,N" & vbNewLine
            '            bcode = bcode & "^FD>:" & Mid(esn, 1, 1) & ">5" & Mid(esn, 2, Len(esn) - 4) & ">6" & Mid(esn, Len(esn) - 2, 3) & "^FS" & vbNewLine
            bcode = bcode & "^FD>:" & esn & "^FS" & vbNewLine
            bcode = bcode & "^FT49,98^A0N,33,33^FH\^FD" & esn & "^FS" & vbNewLine

            'bcode = bcode & "^FT386,59^A0N,25,24^FH\^FD" & def5 & "^FS" & vbNewLine
            'bcode = bcode & "^FT331,59^A0N,25,24^FH\^FD" & def4 & "^FS" & vbNewLine
            'bcode = bcode & "^FT386,26^A0N,25,24^FH\^FD" & def3 & "^FS" & vbNewLine
            'bcode = bcode & "^FT441,27^A0N,25,24^FH\^FD" & def2 & "^FS" & vbNewLine
            'bcode = bcode & "^FT332,27^A0N,25,24^FH\^FD" & def1 & "^FS" & vbNewLine

            bcode = bcode & "^FT154,55^A0N,29,28^FH\^FD" & model & "^FS" & vbNewLine
            '            bcode = bcode & "^FT250,56^A0N,29,28^FH\^FD" & Color & "^FS" & vbNewLine

            bcode = bcode & "^FT48,56^A0N,29,28^FH\^FD" & CUS & "^FS" & vbNewLine
            'bcode = bcode & "^FT439,59^A0N,42,40^FH\^FD:^FS" & vbNewLine
            'bcode = bcode & "^PQ1,0,1,Y^XZ" & vbNewLine
            bcode = bcode & "^PQ1,0,1,Y^XZ" & vbNewLine

            Microsoft.VisualBasic.FileSystem.Write(FN, bcode)
            Microsoft.VisualBasic.FileClose()

            Dim PORT As String = Query_RS("SELECT CN_TYPE FROM tbl_bcconfig WHERE MODEL = (SELECT MODEL FROM TBL_FESNMASTER WHERE ESN = '" & esn & "') AND BC_TYPE = 'RECEIVING'")
            If PORT = "" Then
                PORT = "COM1"
            End If

            Shell("c:\windows\system32\cmd.exe /c print c:\esn.zpl /d:" & PORT)

            '            Shell("c:\windows\system32\cmd.exe /c print c:\esn.zpl /d:com1")
        End If

    End Function

    Function Print_NewBBarcode_NoTrg(ByVal esn As String, ByVal model As String, ByVal def As String, ByVal dt As String) As Boolean

        Dim FN
        Dim bcode As String

        Dim CUS, Color As String
        Dim ba_rs As New ADODB.Recordset
        ba_rs = Query_RS_ALL("SELECT SW_VER, color FROM TBL_MODELMASTER WHERE MODEL_NO = (SELECT MODEL FROM TBL_ESNMASTER WHERE ESN = '" & esn & "')")
        CUS = ba_rs(0).Value
        Color = ba_rs(1).Value
        ba_rs = Nothing

        '       Dim ARR As String() = GET_TRG(esn, "BORDER")
        'ba_rs = Query_RS_ALL("SELECT def_cd FROM TBL_triage WHERE esn_obid = (SELECT obid FROM TBL_FESNMASTER WHERE ESN = '" & esn & "') order by c_date")

        'Dim def1 As String = ARR(4)
        'Dim def2 As String = ARR(3)
        'Dim def3 As String = ARR(2)
        'Dim def4 As String = ARR(1)
        'Dim def5 As String = ARR(0)

        'If ARR Is Nothing Then
        'Else
        '    For I As Integer = 0 To ARR.Count
        '        If I = 0 Then

        '        End If
        '    Next
        'End If

        FN = Microsoft.VisualBasic.FileSystem.FreeFile
        Microsoft.VisualBasic.FileOpen(FN, "C:\ESN1.ZPL", OpenMode.Output)
        '        Microsoft.VisualBasic.FileSystem.Write(FN, "/ ^XA" & vbNewLine & "/ ^LT0" & vbNewLine & "/ ^LH40,0" & vbNewLine & "/ ^FO1,100^A0N,13,25^FD" & model & "-" & def & "^FS" & vbNewLine & "/ ^FO1,50^BY1" & vbNewLine & "/ ^B3N,N,30,Y,N" & vbNewLine & "/ ^FD" & esn & "^FS" & vbNewLine & "/ ^XZ")

        If Site_id = "S1000" Then

            bcode = "^XA~TA000~JSN^LT0^MMT^MNW^MTT^PON^PMN^LH0,0^JMA^PR2,2^MD10^JUS^LRN^CI0^XZ" & vbNewLine
            bcode = bcode & "^XA^LL0150" & vbNewLine
            bcode = bcode & "^PW510" & vbNewLine
            bcode = bcode & "^BY2,3,34^FT47,145^BCN,,N,N" & vbNewLine
            bcode = bcode & "^FD>:" & esn & "^FS" & vbNewLine
            bcode = bcode & "^FT49,98^A0N,33,33^FH\^FD" & esn & "^FS" & vbNewLine

            'bcode = bcode & "^FT386,59^A0N,25,24^FH\^FD" & def5 & "^FS" & vbNewLine
            'bcode = bcode & "^FT331,59^A0N,25,24^FH\^FD" & def4 & "^FS" & vbNewLine
            'bcode = bcode & "^FT386,26^A0N,25,24^FH\^FD" & def3 & "^FS" & vbNewLine
            'bcode = bcode & "^FT441,27^A0N,25,24^FH\^FD" & def2 & "^FS" & vbNewLine
            'bcode = bcode & "^FT332,27^A0N,25,24^FH\^FD" & def1 & "^FS" & vbNewLine

            bcode = bcode & "^FT154,55^A0N,29,28^FH\^FD" & model & "^FS" & vbNewLine
            '            bcode = bcode & "^FT250,56^A0N,29,28^FH\^FD" & Color & "^FS" & vbNewLine

            bcode = bcode & "^FT48,56^A0N,29,28^FH\^FD" & CUS & "^FS" & vbNewLine
            'bcode = bcode & "^FT439,59^A0N,42,40^FH\^FD:^FS" & vbNewLine
            'bcode = bcode & "^PQ1,0,1,Y^XZ" & vbNewLine
            bcode = bcode & "^PQ1,0,1,Y^XZ" & vbNewLine

            Microsoft.VisualBasic.FileSystem.Write(FN, bcode)
            Microsoft.VisualBasic.FileClose()

            Dim PORT As String = Query_RS("SELECT CN_TYPE FROM tbl_bcconfig WHERE MODEL = (SELECT MODEL FROM TBL_ESNMASTER WHERE ESN = '" & esn & "') AND BC_TYPE = 'RECEIVING'")
            If PORT = "" Then
                PORT = "COM1"
            End If

            Shell("c:\windows\system32\cmd.exe /c print c:\esn1.zpl /d:" & PORT)

            '            Shell("c:\windows\system32\cmd.exe /c print c:\esn1.zpl /d:com1")
        End If

    End Function


End Module
