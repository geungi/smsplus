Public Class FrmBarcode

    Private Sub FrmBarcode_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        DockContainerItem1.Text = "실행 메뉴"
        DockContainerItem2.Text = "조회 조건"

        ComboBoxEx1.Items.Add("아마존바코드")
        '        ComboBoxEx1.Items.Add("출고플립ID")

        ComboBoxEx1.Text = "아마존바코드"

        If Query_Combo(Me.ComboBoxEx2, "SELECT model_no FROM tbl_modelmaster WHERE site_id = '" & Site_id & "' and active = 'Y' ORDER BY model_no") = True Then
        End If



        If Spread_Setting(FpSpread1, Me.Name) = True Then
            FpSpread1.ActiveSheet.RowCount = 0
            Spread_AutoCol(FpSpread1)
        End If

        Formbim_Authority(Me.ButtonItem1, Me.Name, "NEW")
        Formbtn_Authority(Me.NewBtn1, Me.Name, "NEW")
        Formbim_Authority(Me.SaveBtn, Me.Name, "SAVE")
        Formbtn_Authority(Me.SaveBtn1, Me.Name, "SAVE")
        Formbim_Authority(Me.DelBtn, Me.Name, "DELETE")
        Formbtn_Authority(Me.DelBtn1, Me.Name, "DELETE")
        Formbim_Authority(Me.PrtBtn, Me.Name, "PRINT")
        Formbtn_Authority(Me.PrtBtn1, Me.Name, "PRINT")
        Formbim_Authority(Me.ExcelBtn, Me.Name, "EXCEL")
        Formbtn_Authority(Me.XlsBtn1, Me.Name, "EXCEL")
        Formbim_Authority(Me.FindBtn, Me.Name, "FIND")

    End Sub

    



    Private Sub SaveBtn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SaveBtn.Click

        If ComboBoxEx2.Text = "" Then
            System.Console.Beep(3000, 400)
            System.Console.Beep(3000, 400)
            Modal_Error("모델을 선택하세요!!")
            Exit Sub
        End If


        If IntegerInput1.Text = "" Then
            System.Console.Beep(3000, 400)
            System.Console.Beep(3000, 400)
            Modal_Error("인쇄할 수량을 입력하세요!!")
            Exit Sub
        End If

        For I As Integer = 1 To IntegerInput1.Text
            Shell("c:\windows\system32\cmd.exe /c COPY c:\BCODE\" & ComboBoxEx2.Text & ".PRN \\" & My.Computer.Name & "\BP03")
            'System.Threading.Thread.Sleep(10)
        Next


    End Sub


End Class