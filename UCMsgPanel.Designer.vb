<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class UCMsgPanel
    Inherits System.Windows.Forms.UserControl

    'UserControl은 Dispose를 재정의하여 구성 요소 목록을 정리합니다.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Windows Form 디자이너에 필요합니다.
    Private components As System.ComponentModel.IContainer

    '참고: 다음 프로시저는 Windows Form 디자이너에 필요합니다.
    '수정하려면 Windows Form 디자이너를 사용하십시오.  
    '코드 편집기를 사용하여 수정하지 마십시오.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Me.Timer1 = New System.Windows.Forms.Timer(Me.components)
        Me.Bar1 = New DevComponents.DotNetBar.Bar
        Me.ctlmove = New DevComponents.DotNetBar.ButtonItem
        Me.ctlclose = New DevComponents.DotNetBar.ButtonItem
        Me.LabelX1 = New DevComponents.DotNetBar.LabelX
        CType(Me.Bar1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Timer1
        '
        Me.Timer1.Interval = 1
        '
        'Bar1
        '
        Me.Bar1.Dock = System.Windows.Forms.DockStyle.Top
        Me.Bar1.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Bar1.items.AddRange(New DevComponents.DotNetBar.BaseItem() {Me.ctlmove, Me.ctlclose})
        Me.Bar1.Location = New System.Drawing.Point(0, 0)
        Me.Bar1.Name = "Bar1"
        Me.Bar1.Size = New System.Drawing.Size(100, 19)
        Me.Bar1.Stretch = True
        Me.Bar1.Style = DevComponents.DotNetBar.eDotNetBarStyle.Office2003
        Me.Bar1.TabIndex = 0
        Me.Bar1.TabStop = False
        Me.Bar1.Text = "Bar1"
        '
        'ctlmove
        '
        Me.ctlmove.Image = Global.SMSPLUS.My.Resources.Resources.sCheckMark
        Me.ctlmove.ImageFixedSize = New System.Drawing.Size(10, 10)
        Me.ctlmove.ImagePaddingHorizontal = 8
        Me.ctlmove.Name = "ctlmove"
        Me.ctlmove.Text = "ButtonItem1"
        '
        'ctlclose
        '
        Me.ctlclose.Image = Global.SMSPLUS.My.Resources.Resources.sMinus
        Me.ctlclose.ImageFixedSize = New System.Drawing.Size(10, 10)
        Me.ctlclose.ImagePaddingHorizontal = 8
        Me.ctlclose.ItemAlignment = DevComponents.DotNetBar.eItemAlignment.Far
        Me.ctlclose.Name = "ctlclose"
        Me.ctlclose.Text = "ButtonItem2"
        '
        'LabelX1
        '
        Me.LabelX1.BackColor = System.Drawing.Color.Transparent
        Me.LabelX1.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None
        '
        '
        '
        Me.LabelX1.BackgroundStyle.BackColor = System.Drawing.Color.Transparent
        Me.LabelX1.BackgroundStyle.BackColor2 = System.Drawing.Color.Transparent
        Me.LabelX1.BackgroundStyle.BackColorGradientType = DevComponents.DotNetBar.eGradientType.Radial
        Me.LabelX1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.LabelX1.Font = New System.Drawing.Font("Impact", 30.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LabelX1.Location = New System.Drawing.Point(0, 19)
        Me.LabelX1.Name = "LabelX1"
        Me.LabelX1.Size = New System.Drawing.Size(100, 81)
        Me.LabelX1.TabIndex = 1
        Me.LabelX1.Text = "0"
        Me.LabelX1.TextAlignment = System.Drawing.StringAlignment.Center
        '
        'UCMsgPanel
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 14.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.Transparent
        Me.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None
        Me.Controls.Add(Me.LabelX1)
        Me.Controls.Add(Me.Bar1)
        Me.Font = New System.Drawing.Font("Verdana", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Name = "UCMsgPanel"
        Me.Size = New System.Drawing.Size(100, 100)
        CType(Me.Bar1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents Timer1 As System.Windows.Forms.Timer
    Friend WithEvents Bar1 As DevComponents.DotNetBar.Bar
    Friend WithEvents ctlmove As DevComponents.DotNetBar.ButtonItem
    Friend WithEvents ctlclose As DevComponents.DotNetBar.ButtonItem
    Friend WithEvents LabelX1 As DevComponents.DotNetBar.LabelX

End Class
