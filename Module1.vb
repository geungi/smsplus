Option Strict On

Imports System.Drawing.Drawing2D
Imports System.Drawing.Text
Imports SMSPLUS.DrawHelpers

'<info>
' --------------------Fader Theme--------------------
' Creator - SaketSaket (HF)
' UID - 1869668
' Inspiration & Credits to all Theme creators of HF
' Version - 1.0
' Date Created - 1st December 2014
' Date Modified - 12th December 2014
'
'
'Special Thanks to Aeonhack for RoundRect Functions...
'AlertBox Control idea taken from iSynthesis' Flat UI theme
'
'
' For bugs & Constructive Criticism contact me on HF
' If you like it & want to DONATE then pm me on HF
' --------------------Fader Theme--------------------
'<info>

'Please Leave Credits in Source, Do not redistribute

Enum MouseState As Byte
    None = 0
    Over = 1
    Down = 2
End Enum

Public Enum AlignmentStyle
    Center
    Left
    Right
End Enum

Module Draw
    'Special Thanks to Aeonhack for RoundRect Functions... ;)
    Public Function RoundRect(ByVal rectangle As Rectangle, ByVal curve As Integer) As GraphicsPath
        Dim p As GraphicsPath = New GraphicsPath()
        Dim arcRectangleWidth As Integer = curve * 2
        p.AddArc(New Rectangle(rectangle.X, rectangle.Y, arcRectangleWidth, arcRectangleWidth), -180, 90)
        p.AddArc(New Rectangle(rectangle.Width - arcRectangleWidth + rectangle.X, rectangle.Y, arcRectangleWidth, arcRectangleWidth), -90, 90)
        p.AddArc(New Rectangle(rectangle.Width - arcRectangleWidth + rectangle.X, rectangle.Height - arcRectangleWidth + rectangle.Y, arcRectangleWidth, arcRectangleWidth), 0, 90)
        p.AddArc(New Rectangle(rectangle.X, rectangle.Height - arcRectangleWidth + rectangle.Y, arcRectangleWidth, arcRectangleWidth), 90, 90)
        p.AddLine(New Point(rectangle.X, rectangle.Height - arcRectangleWidth + rectangle.Y), New Point(rectangle.X, curve + rectangle.Y))
        Return p
    End Function
End Module

Public Class FaderTheme : Inherits ContainerControl
    Private _mousepos As Point = New Point(0, 0)
    Private _drag As Boolean = False
    Private _icon As Icon

    Public Property Icon() As Icon
        Get
            Return _icon
        End Get
        Set(ByVal value As Icon)
            _icon = value
            Invalidate()
        End Set
    End Property

    Private _showicon As Boolean = True
    Public Property ShowIcon As Boolean
        Get
            Return _showicon
        End Get
        Set(value As Boolean)
            _showicon = value
            Invalidate()
        End Set
    End Property

    Private _showheader As Boolean = True
    Public Property ShowHeader As Boolean
        Get
            Return _showheader
        End Get
        Set(value As Boolean)
            _showheader = value
            Invalidate()
        End Set
    End Property

    Private _headerAlignment As AlignmentStyle = AlignmentStyle.Center
    Public Property HeaderAlignment As AlignmentStyle
        Get
            Return _headerAlignment
        End Get
        Set(value As AlignmentStyle)
            _headerAlignment = value
            Invalidate()
        End Set
    End Property

    Protected Overrides Sub OnMouseDown(ByVal e As MouseEventArgs)
        MyBase.OnMouseDown(e)
        If e.Button = Windows.Forms.MouseButtons.Left And New Rectangle(15, 10, Width - 31, 45).Contains(e.Location) Then
            _drag = True : _mousepos = e.Location
        End If
    End Sub

    Protected Overrides Sub OnMouseUp(ByVal e As MouseEventArgs)
        MyBase.OnMouseUp(e)
        _drag = False
    End Sub

    Protected Overrides Sub OnMouseMove(ByVal e As MouseEventArgs)
        MyBase.OnMouseMove(e)
        If _drag Then
            Parent.Location = New Point(MousePosition.X - _mousepos.X, MousePosition.Y - _mousepos.Y)
        End If
    End Sub

    Sub New()
        MyBase.New()
        SetStyle(ControlStyles.AllPaintingInWmPaint Or ControlStyles.OptimizedDoubleBuffer Or ControlStyles.ResizeRedraw Or ControlStyles.UserPaint, True)
        DoubleBuffered = True
    End Sub

    Protected Overrides Sub OnCreateControl()
        MyBase.OnCreateControl()
        ParentForm.FormBorderStyle = FormBorderStyle.None
        ParentForm.TransparencyKey = Color.Fuchsia
        Dock = DockStyle.Fill
    End Sub

    Protected Overrides Sub OnPaint(ByVal e As PaintEventArgs)
        Dim b As Bitmap = New Bitmap(Width, Height)
        Dim g As Graphics = Graphics.FromImage(b)

        Dim bodyrect As Rectangle = New Rectangle(15, 10, Width - 31, Height - 21)
        Dim headerrect As Rectangle = New Rectangle(15, 10, Width - 31, 50)
        Dim footerrect As Rectangle = New Rectangle(15, Height - 25, Width - 31, 15)
        Dim leftborder As Rectangle = New Rectangle(0, 0, 15, Height - 1)
        Dim rightborder As Rectangle = New Rectangle(Width - 16, 0, 15, Height - 1)

        MyBase.OnPaint(e)
        g.Clear(Color.Fuchsia)

        Dim bodygb As New LinearGradientBrush(bodyrect, Color.FromArgb(61, 61, 61), Color.FromArgb(41, 41, 41), 90S)
        g.FillRectangle(bodygb, bodyrect)
        g.DrawRectangle(New Pen(Brushes.Black), bodyrect)

        Dim headergb As New LinearGradientBrush(headerrect, Color.FromArgb(61, 61, 61), Color.FromArgb(41, 41, 41), 90S)
        g.FillRectangle(headergb, headerrect)
        g.DrawRectangle(New Pen(Brushes.Black), headerrect)

        Dim footergb As New LinearGradientBrush(footerrect, Color.FromArgb(61, 61, 61), Color.FromArgb(41, 41, 41), 90S)
        g.FillRectangle(footergb, footerrect)
        g.DrawRectangle(New Pen(Brushes.Black), footerrect)

        Dim leftbordergb As New LinearGradientBrush(leftborder, Color.FromArgb(53, 53, 53), Color.FromArgb(62, 62, 62), LinearGradientMode.Horizontal)
        g.FillPath(leftbordergb, SMSPLUS.Draw.RoundRect(leftborder, 7))
        g.DrawPath(New Pen(Brushes.Black), SMSPLUS.Draw.RoundRect(leftborder, 7))

        Dim rightbordergb As New LinearGradientBrush(rightborder, Color.FromArgb(53, 53, 53), Color.FromArgb(62, 62, 62), LinearGradientMode.Horizontal)
        g.FillPath(rightbordergb, SMSPLUS.Draw.RoundRect(rightborder, 7))
        g.DrawPath(New Pen(Brushes.Black), SMSPLUS.Draw.RoundRect(rightborder, 7))

        If ShowIcon = True Then
            Try
                g.DrawIcon(_icon, New Rectangle(20, 17, 38, 38))
            Catch : End Try
        End If

        If ShowHeader = True Then
            Select Case _headerAlignment
                Case AlignmentStyle.Left
                    g.DrawString(Text, New Font("Segoe UI", 15, FontStyle.Bold), New SolidBrush(Color.FromArgb(220, 220, 220)), New Rectangle(65, 10, Width - 31, 45), New StringFormat() With {.Alignment = StringAlignment.Near, .LineAlignment = StringAlignment.Center})
                Case AlignmentStyle.Center
                    g.DrawString(Text, New Font("Segoe UI", 15, FontStyle.Bold), New SolidBrush(Color.FromArgb(220, 220, 220)), headerrect, New StringFormat() With {.Alignment = StringAlignment.Center, .LineAlignment = StringAlignment.Center})
                Case AlignmentStyle.Right
                    g.DrawString(Text, New Font("Segoe UI", 15, FontStyle.Bold), New SolidBrush(Color.FromArgb(220, 220, 220)), New Rectangle(0, 10, Width - 120, 45), New StringFormat() With {.Alignment = StringAlignment.Far, .LineAlignment = StringAlignment.Center})
            End Select
        End If

        e.Graphics.DrawImage(b, New Point(0, 0))
        g.Dispose() : b.Dispose()
    End Sub
End Class

Public Class FaderMinimalTheme : Inherits ContainerControl
    Private _mousepos As Point = New Point(0, 0)
    Private _drag As Boolean = False

    Protected Overrides Sub OnMouseDown(ByVal e As MouseEventArgs)
        MyBase.OnMouseDown(e)
        If e.Button = Windows.Forms.MouseButtons.Left And New Rectangle(15, 10, Width - 31, 45).Contains(e.Location) Then
            _drag = True : _mousepos = e.Location
        End If
    End Sub

    Protected Overrides Sub OnMouseUp(ByVal e As MouseEventArgs)
        MyBase.OnMouseUp(e)
        _drag = False
    End Sub

    Protected Overrides Sub OnMouseMove(ByVal e As MouseEventArgs)
        MyBase.OnMouseMove(e)
        If _drag Then
            Parent.Location = New Point(MousePosition.X - _mousepos.X, MousePosition.Y - _mousepos.Y)
        End If
    End Sub

    Sub New()
        MyBase.New()
        SetStyle(ControlStyles.AllPaintingInWmPaint Or ControlStyles.OptimizedDoubleBuffer Or ControlStyles.ResizeRedraw Or ControlStyles.UserPaint, True)
        DoubleBuffered = True
    End Sub

    Protected Overrides Sub OnCreateControl()
        MyBase.OnCreateControl()
        ParentForm.FormBorderStyle = FormBorderStyle.None
        ParentForm.TransparencyKey = Color.Fuchsia
        Dock = DockStyle.Fill
    End Sub

    Protected Overrides Sub OnPaint(ByVal e As PaintEventArgs)
        Dim b As Bitmap = New Bitmap(Width, Height)
        Dim g As Graphics = Graphics.FromImage(b)

        Dim bodyrect As Rectangle = New Rectangle(15, 10, Width - 31, Height - 21)
        Dim headerrect As Rectangle = New Rectangle(15, 10, Width - 31, 15)
        Dim footerrect As Rectangle = New Rectangle(15, Height - 25, Width - 31, 15)
        Dim leftborder As Rectangle = New Rectangle(0, 0, 15, Height - 1)
        Dim rightborder As Rectangle = New Rectangle(Width - 16, 0, 15, Height - 1)

        MyBase.OnPaint(e)
        g.Clear(Color.Fuchsia)

        Dim bodygb As New LinearGradientBrush(bodyrect, Color.FromArgb(61, 61, 61), Color.FromArgb(41, 41, 41), 90S)
        g.FillRectangle(bodygb, bodyrect)
        g.DrawRectangle(New Pen(Brushes.Black), bodyrect)

        Dim headergb As New LinearGradientBrush(headerrect, Color.FromArgb(61, 61, 61), Color.FromArgb(41, 41, 41), 90S)
        g.FillRectangle(headergb, headerrect)
        g.DrawRectangle(New Pen(Brushes.Black), headerrect)

        Dim footergb As New LinearGradientBrush(footerrect, Color.FromArgb(61, 61, 61), Color.FromArgb(41, 41, 41), 90S)
        g.FillRectangle(footergb, footerrect)
        g.DrawRectangle(New Pen(Brushes.Black), footerrect)

        Dim leftbordergb As New LinearGradientBrush(leftborder, Color.FromArgb(53, 53, 53), Color.FromArgb(62, 62, 62), LinearGradientMode.Horizontal)
        g.FillPath(leftbordergb, SMSPLUS.Draw.RoundRect(leftborder, 7))
        g.DrawPath(New Pen(Brushes.Black), SMSPLUS.Draw.RoundRect(leftborder, 7))

        Dim rightbordergb As New LinearGradientBrush(rightborder, Color.FromArgb(53, 53, 53), Color.FromArgb(62, 62, 62), LinearGradientMode.Horizontal)
        g.FillPath(rightbordergb, SMSPLUS.Draw.RoundRect(rightborder, 7))
        g.DrawPath(New Pen(Brushes.Black), SMSPLUS.Draw.RoundRect(rightborder, 7))

        e.Graphics.DrawImage(b, New Point(0, 0))
        g.Dispose() : b.Dispose()
    End Sub
End Class

Public Class FaderControlBox : Inherits Control
    Dim _state As MouseState = MouseState.None
    Dim _x As Integer
    ReadOnly _minrect As New Rectangle(5, 2, 24, 24)
    ReadOnly _maxrect As New Rectangle(32, 2, 24, 24)
    ReadOnly _closerect As New Rectangle(59, 2, 24, 24)
    ReadOnly _mintextrect As New Rectangle(5, 4, 24, 24)
    ReadOnly _maxtextrect As New Rectangle(32, 4, 24, 24)
    ReadOnly _closetextrect As New Rectangle(59, 4, 24, 24)

    Private _minDisable As Boolean = False
    Public Property MinimumDisable As Boolean
        Get
            Return _minDisable
        End Get
        Set(value As Boolean)
            _minDisable = value
            Invalidate()
        End Set
    End Property

    Private _maxDisable As Boolean = False
    Public Property MaximumDisable As Boolean
        Get
            Return _maxDisable
        End Get
        Set(value As Boolean)
            _maxDisable = value
            Invalidate()
        End Set
    End Property

    Protected Overrides Sub OnMouseDown(ByVal e As Windows.Forms.MouseEventArgs)
        MyBase.OnMouseDown(e)

        If _x > 5 AndAlso _x < 29 Then
            If MinimumDisable = False Then
                FindForm.WindowState = FormWindowState.Minimized
            End If
        ElseIf _x > 32 AndAlso _x < 56 Then
            If MaximumDisable = False Then
                If FindForm.WindowState = FormWindowState.Maximized Then
                    FindForm.WindowState = FormWindowState.Minimized
                    FindForm.WindowState = FormWindowState.Normal
                Else
                    FindForm.WindowState = FormWindowState.Minimized
                    FindForm.WindowState = FormWindowState.Maximized
                End If
            End If
        ElseIf _x > 59 AndAlso _x < 83 Then
            FindForm.Close()
        End If

        _state = MouseState.Down : Invalidate()
    End Sub

    Protected Overrides Sub OnMouseUp(ByVal e As Windows.Forms.MouseEventArgs)
        MyBase.OnMouseUp(e)
        _state = MouseState.Over : Invalidate()
    End Sub

    Protected Overrides Sub OnMouseEnter(ByVal e As EventArgs)
        MyBase.OnMouseEnter(e)
        _state = MouseState.Over : Invalidate()
    End Sub

    Protected Overrides Sub OnMouseLeave(ByVal e As EventArgs)
        MyBase.OnMouseLeave(e)
        _state = MouseState.None : Invalidate()
    End Sub

    Protected Overrides Sub OnMouseMove(ByVal e As Windows.Forms.MouseEventArgs)
        MyBase.OnMouseMove(e)
        _x = e.Location.X
        Invalidate()
    End Sub

    Sub New()
        MyBase.New()
        SetStyle(ControlStyles.AllPaintingInWmPaint Or ControlStyles.OptimizedDoubleBuffer Or ControlStyles.ResizeRedraw Or ControlStyles.UserPaint Or ControlStyles.SupportsTransparentBackColor, True)
        DoubleBuffered = True
        BackColor = Color.Transparent
        Width = 85
        Height = 30
        Anchor = AnchorStyles.Top Or AnchorStyles.Right
    End Sub

    Protected Overrides Sub OnPaint(ByVal e As PaintEventArgs)
        Dim b As Bitmap = New Bitmap(Width, Height)
        Dim g As Graphics = Graphics.FromImage(b)

        MyBase.OnPaint(e)
        g.SmoothingMode = SmoothingMode.HighQuality
        g.FillRectangle(New SolidBrush(Color.FromArgb(161, 161, 161)), _minrect)
        g.DrawRectangle(Pens.Black, _minrect)
        g.FillRectangle(New SolidBrush(Color.FromArgb(161, 161, 161)), _maxrect)
        g.DrawRectangle(Pens.Black, _maxrect)
        g.FillRectangle(New SolidBrush(Color.FromArgb(161, 161, 161)), _closerect)
        g.DrawRectangle(Pens.Black, _closerect)

        g.DrawString("0", New Font("Marlett", 11.5), New SolidBrush(Color.FromArgb(0, 0, 0)), _mintextrect, New StringFormat() With {.Alignment = StringAlignment.Center, .LineAlignment = StringAlignment.Center})
        If FindForm.WindowState = FormWindowState.Maximized Then
            g.DrawString("2", New Font("Marlett", 11.5), New SolidBrush(Color.FromArgb(0, 0, 0)), _maxtextrect, New StringFormat() With {.Alignment = StringAlignment.Center, .LineAlignment = StringAlignment.Center})
        Else
            g.DrawString("1", New Font("Marlett", 11.5), New SolidBrush(Color.FromArgb(0, 0, 0)), _maxtextrect, New StringFormat() With {.Alignment = StringAlignment.Center, .LineAlignment = StringAlignment.Center})
        End If
        g.DrawString("r", New Font("Marlett", 11.5), New SolidBrush(Color.FromArgb(0, 0, 0)), _closetextrect, New StringFormat() With {.Alignment = StringAlignment.Center, .LineAlignment = StringAlignment.Center})

        Select Case _state
            Case MouseState.None
                g.FillRectangle(New SolidBrush(Color.FromArgb(161, 161, 161)), _minrect)
                g.DrawRectangle(Pens.Black, _minrect)
                g.FillRectangle(New SolidBrush(Color.FromArgb(161, 161, 161)), _maxrect)
                g.DrawRectangle(Pens.Black, _maxrect)
                g.FillRectangle(New SolidBrush(Color.FromArgb(161, 161, 161)), _closerect)
                g.DrawRectangle(Pens.Black, _closerect)

                g.DrawString("0", New Font("Marlett", 11.5), New SolidBrush(Color.FromArgb(0, 0, 0)), _mintextrect, New StringFormat() With {.Alignment = StringAlignment.Center, .LineAlignment = StringAlignment.Center})
                If FindForm.WindowState = FormWindowState.Maximized Then
                    g.DrawString("2", New Font("Marlett", 11.5), New SolidBrush(Color.FromArgb(0, 0, 0)), _maxtextrect, New StringFormat() With {.Alignment = StringAlignment.Center, .LineAlignment = StringAlignment.Center})
                Else
                    g.DrawString("1", New Font("Marlett", 11.5), New SolidBrush(Color.FromArgb(0, 0, 0)), _maxtextrect, New StringFormat() With {.Alignment = StringAlignment.Center, .LineAlignment = StringAlignment.Center})
                End If
                g.DrawString("r", New Font("Marlett", 11.5), New SolidBrush(Color.FromArgb(0, 0, 0)), _closetextrect, New StringFormat() With {.Alignment = StringAlignment.Center, .LineAlignment = StringAlignment.Center})

            Case MouseState.Over
                If _x > 5 AndAlso _x < 29 Then
                    If MinimumDisable = False Then
                        g.FillRectangle(New SolidBrush(Color.FromArgb(121, 121, 121)), _minrect)
                        g.DrawRectangle(Pens.Black, _minrect)
                        g.DrawString("0", New Font("Marlett", 11.5), New SolidBrush(Color.FromArgb(255, 255, 255)), _mintextrect, New StringFormat() With {.Alignment = StringAlignment.Center, .LineAlignment = StringAlignment.Center})
                    End If
                ElseIf _x > 32 AndAlso _x < 56 Then
                    If MaximumDisable = False Then
                        g.FillRectangle(New SolidBrush(Color.FromArgb(121, 121, 121)), _maxrect)
                        g.DrawRectangle(Pens.Black, _maxrect)
                    End If

                    If FindForm.WindowState = FormWindowState.Maximized Then
                        g.DrawString("2", New Font("Marlett", 11.5), New SolidBrush(Color.FromArgb(255, 255, 255)), _maxtextrect, New StringFormat() With {.Alignment = StringAlignment.Center, .LineAlignment = StringAlignment.Center})
                    Else
                        g.DrawString("1", New Font("Marlett", 11.5), New SolidBrush(Color.FromArgb(255, 255, 255)), _maxtextrect, New StringFormat() With {.Alignment = StringAlignment.Center, .LineAlignment = StringAlignment.Center})
                    End If
                ElseIf _x > 59 AndAlso _x < 83 Then
                    g.FillRectangle(New SolidBrush(Color.FromArgb(121, 121, 121)), _closerect)
                    g.DrawRectangle(Pens.Black, _closerect)
                    g.DrawString("r", New Font("Marlett", 11.5), New SolidBrush(Color.FromArgb(255, 255, 255)), _closetextrect, New StringFormat() With {.Alignment = StringAlignment.Center, .LineAlignment = StringAlignment.Center})
                End If
            Case Else
                If MinimumDisable = False Then
                    g.FillRectangle(New SolidBrush(Color.FromArgb(81, 81, 81)), _minrect)
                    g.DrawRectangle(Pens.Black, _minrect)
                End If

                If MaximumDisable = False Then
                    g.FillRectangle(New SolidBrush(Color.FromArgb(81, 81, 81)), _maxrect)
                    g.DrawRectangle(Pens.Black, _maxrect)
                End If
                g.FillRectangle(New SolidBrush(Color.FromArgb(81, 81, 81)), _closerect)
                g.DrawRectangle(Pens.Black, _closerect)

                g.DrawString("0", New Font("Marlett", 11.5), New SolidBrush(Color.FromArgb(0, 0, 0)), _mintextrect, New StringFormat() With {.Alignment = StringAlignment.Center, .LineAlignment = StringAlignment.Center})
                If FindForm.WindowState = FormWindowState.Maximized Then
                    g.DrawString("2", New Font("Marlett", 11.5), New SolidBrush(Color.FromArgb(0, 0, 0)), _maxtextrect, New StringFormat() With {.Alignment = StringAlignment.Center, .LineAlignment = StringAlignment.Center})
                Else
                    g.DrawString("1", New Font("Marlett", 11.5), New SolidBrush(Color.FromArgb(0, 0, 0)), _maxtextrect, New StringFormat() With {.Alignment = StringAlignment.Center, .LineAlignment = StringAlignment.Center})
                End If
                g.DrawString("r", New Font("Marlett", 11.5), New SolidBrush(Color.FromArgb(0, 0, 0)), _closetextrect, New StringFormat() With {.Alignment = StringAlignment.Center, .LineAlignment = StringAlignment.Center})
        End Select

        e.Graphics.DrawImage(b, New Point(0, 0))
        g.Dispose() : b.Dispose()
    End Sub
End Class

Public Class FaderRadioButton : Inherits Control
    Dim _state As MouseState = MouseState.None
    Private _check As Boolean
    Public Property Checked As Boolean
        Get
            Return _check
        End Get
        Set(value As Boolean)
            _check = value
            Invalidate()
        End Set
    End Property

    Sub New()
        MyBase.New()
        SetStyle(ControlStyles.AllPaintingInWmPaint Or ControlStyles.OptimizedDoubleBuffer Or ControlStyles.ResizeRedraw Or ControlStyles.UserPaint Or ControlStyles.SupportsTransparentBackColor, True)
        BackColor = Color.Transparent
        DoubleBuffered = True
        Size = New Size(180, 25)
    End Sub

    Protected Overrides Sub OnTextChanged(ByVal e As EventArgs)
        MyBase.OnTextChanged(e)
        Invalidate()
    End Sub

    Protected Overrides Sub OnClick(ByVal e As EventArgs)
        If Not Checked Then Checked = True
        For Each ctrl As FaderRadioButton In From ctrl1 In Parent.Controls.OfType(Of FaderRadioButton)() Where ctrl1.Handle <> Handle Where ctrl1.Enabled
            ctrl.Checked = False
        Next
        MyBase.OnClick(e)
    End Sub

    Protected Overrides Sub OnMouseDown(e As MouseEventArgs)
        MyBase.OnMouseDown(e)
        _state = MouseState.Down : Invalidate()
    End Sub
    Protected Overrides Sub OnMouseUp(e As MouseEventArgs)
        MyBase.OnMouseUp(e)
        _state = MouseState.Over : Invalidate()
    End Sub
    Protected Overrides Sub OnMouseEnter(e As EventArgs)
        MyBase.OnMouseEnter(e)
        _state = MouseState.Over : Invalidate()
    End Sub
    Protected Overrides Sub OnMouseLeave(e As EventArgs)
        MyBase.OnMouseLeave(e)
        _state = MouseState.None : Invalidate()
    End Sub

    Protected Overrides Sub OnPaint(ByVal e As PaintEventArgs)
        Dim b As Bitmap = New Bitmap(Width, Height)
        Dim g As Graphics = Graphics.FromImage(b)

        Dim selectionrect As Rectangle = New Rectangle(3, 3, 18, 18)
        Dim innerselectionrect As Rectangle = New Rectangle(4, 4, 17, 17)
        Dim selectrect As Rectangle = New Rectangle(8, 8, 8, 8)

        MyBase.OnPaint(e)
        g.SmoothingMode = SmoothingMode.HighQuality
        g.InterpolationMode = InterpolationMode.HighQualityBicubic
        g.TextRenderingHint = Drawing.Text.TextRenderingHint.AntiAlias
        g.Clear(BackColor)

        Select Case _state
            Case MouseState.Over
                selectionrect.Inflate(1, 1)
            Case MouseState.Down
                selectionrect.Inflate(-1, -1)
        End Select

        g.DrawString(Text, New Font("Segoe UI", 11, FontStyle.Bold), New SolidBrush(Color.FromArgb(245, 245, 245)), New Rectangle(25, 4, Width, 16), New StringFormat() With {.Alignment = StringAlignment.Near, .LineAlignment = StringAlignment.Center})

        If Checked Then
            g.FillEllipse(New SolidBrush(Color.FromArgb(0, 0, 0)), selectionrect)
            g.FillEllipse(New SolidBrush(Color.FromArgb(40, 37, 33)), innerselectionrect)
            g.FillEllipse(New SolidBrush(Color.FromArgb(245, 245, 245)), selectrect)
        Else
            g.FillEllipse(New SolidBrush(Color.FromArgb(0, 0, 0)), selectionrect)
            g.FillEllipse(New SolidBrush(Color.FromArgb(40, 37, 33)), innerselectionrect)
        End If

        e.Graphics.DrawImage(b, New Point(0, 0))
        g.Dispose() : b.Dispose()
    End Sub
End Class

Public Class FaderCheckBox : Inherits Control
    Dim _state As MouseState = MouseState.None

    Private _check As Boolean
    Public Property Checked As Boolean
        Get
            Return _check
        End Get
        Set(value As Boolean)
            _check = value
            Invalidate()
        End Set
    End Property

    Sub New()
        MyBase.New()
        SetStyle(ControlStyles.AllPaintingInWmPaint Or ControlStyles.OptimizedDoubleBuffer Or ControlStyles.ResizeRedraw Or ControlStyles.UserPaint Or ControlStyles.SupportsTransparentBackColor, True)
        BackColor = Color.Transparent
        DoubleBuffered = True
        Size = New Size(180, 25)
    End Sub

    Protected Overrides Sub OnTextChanged(ByVal e As EventArgs)
        MyBase.OnTextChanged(e)
        Invalidate()
    End Sub

    Protected Overrides Sub OnClick(ByVal e As EventArgs)
        If Not Checked Then Checked = True Else Checked = False
        MyBase.OnClick(e)
    End Sub

    Protected Overrides Sub OnMouseDown(e As MouseEventArgs)
        MyBase.OnMouseDown(e)
        _state = MouseState.Down : Invalidate()
    End Sub
    Protected Overrides Sub OnMouseUp(e As MouseEventArgs)
        MyBase.OnMouseUp(e)
        _state = MouseState.Over : Invalidate()
    End Sub
    Protected Overrides Sub OnMouseEnter(e As EventArgs)
        MyBase.OnMouseEnter(e)
        _state = MouseState.Over : Invalidate()
    End Sub
    Protected Overrides Sub OnMouseLeave(e As EventArgs)
        MyBase.OnMouseLeave(e)
        _state = MouseState.None : Invalidate()
    End Sub

    Protected Overrides Sub OnPaint(ByVal e As PaintEventArgs)
        Dim b As Bitmap = New Bitmap(Width, Height)
        Dim g As Graphics = Graphics.FromImage(b)

        Dim selectionrect As Rectangle = New Rectangle(3, 3, 18, 18)
        Dim innerselectionrect As Rectangle = New Rectangle(4, 4, 17, 17)
        Dim selectrect As Rectangle = New Rectangle(6, 6, 15, 15)

        MyBase.OnPaint(e)
        g.SmoothingMode = SmoothingMode.HighQuality
        g.InterpolationMode = InterpolationMode.HighQualityBicubic
        g.TextRenderingHint = Drawing.Text.TextRenderingHint.AntiAlias
        g.Clear(BackColor)

        Select Case _state
            Case MouseState.Over
                selectionrect.Inflate(1, 1)
            Case MouseState.Down
                selectionrect.Inflate(-1, -1)
        End Select

        g.DrawString(Text, New Font("Segoe UI", 11, FontStyle.Bold), New SolidBrush(Color.FromArgb(245, 245, 245)), New Rectangle(25, 4, Width, 16), New StringFormat() With {.Alignment = StringAlignment.Near, .LineAlignment = StringAlignment.Center})

        If Checked Then
            g.FillRectangle(New SolidBrush(Color.FromArgb(0, 0, 0)), selectionrect)
            g.FillRectangle(New SolidBrush(Color.FromArgb(40, 37, 33)), innerselectionrect)
            g.DrawString("b", New Font("Marlett", 15, FontStyle.Regular), New SolidBrush(Color.FromArgb(245, 245, 245)), selectrect, New StringFormat() With {.Alignment = StringAlignment.Center, .LineAlignment = StringAlignment.Center})
        Else
            g.FillRectangle(New SolidBrush(Color.FromArgb(0, 0, 0)), selectionrect)
            g.FillRectangle(New SolidBrush(Color.FromArgb(40, 37, 33)), innerselectionrect)
        End If

        e.Graphics.DrawImage(b, New Point(0, 0))
        g.Dispose() : b.Dispose()
    End Sub
End Class

Public Class FaderButton : Inherits Control
    Dim _state As MouseState
    Protected Overrides Sub OnMouseDown(ByVal e As MouseEventArgs)
        MyBase.OnMouseDown(e)
        _state = MouseState.Down
        Invalidate()
    End Sub

    Protected Overrides Sub OnMouseUp(ByVal e As MouseEventArgs)
        MyBase.OnMouseUp(e)
        _state = MouseState.Over
        Invalidate()
    End Sub

    Protected Overrides Sub OnMouseEnter(ByVal e As EventArgs)
        MyBase.OnMouseEnter(e)
        _state = MouseState.Over
        Invalidate()
    End Sub

    Protected Overrides Sub OnMouseLeave(ByVal e As EventArgs)
        MyBase.OnMouseLeave(e)
        _state = MouseState.None
        Invalidate()
    End Sub

    Sub New()
        SetStyle(ControlStyles.AllPaintingInWmPaint Or ControlStyles.OptimizedDoubleBuffer Or ControlStyles.ResizeRedraw Or ControlStyles.UserPaint Or ControlStyles.SupportsTransparentBackColor, True)
        BackColor = Color.Transparent
        DoubleBuffered = True
        Size = New Size(160, 35)
        _state = MouseState.None
    End Sub

    Protected Overrides Sub OnPaint(ByVal e As PaintEventArgs)
        Dim b As Bitmap = New Bitmap(Width, Height)
        Dim g As Graphics = Graphics.FromImage(b)

        Dim bodyrect As Rectangle = New Rectangle(1, 5, Width - 3, Height - 11)
        Dim topborderrect As Rectangle = New Rectangle(0, 0, Width - 1, 5)
        Dim bottomborderrect As Rectangle = New Rectangle(0, Height - 6, Width - 1, 5)
        Dim btnfont As New Font("Segoe UI", 11, FontStyle.Bold)

        MyBase.OnPaint(e)
        g.SmoothingMode = SmoothingMode.HighQuality
        g.InterpolationMode = InterpolationMode.HighQualityBicubic
        g.TextRenderingHint = Drawing.Text.TextRenderingHint.AntiAlias
        g.Clear(BackColor)

        Dim bodyrectgb As New LinearGradientBrush(bodyrect, Color.FromArgb(61, 61, 61), Color.FromArgb(41, 41, 41), LinearGradientMode.Vertical)
        g.FillRectangle(bodyrectgb, bodyrect)
        g.DrawRectangle(New Pen(Brushes.Black), bodyrect)

        Dim topborderrectgb As New LinearGradientBrush(topborderrect, Color.FromArgb(53, 53, 53), Color.FromArgb(62, 62, 62), LinearGradientMode.Vertical)
        g.FillRectangle(topborderrectgb, topborderrect)
        g.DrawRectangle(New Pen(Brushes.Black), topborderrect)

        Dim bottomborderrectgb As New LinearGradientBrush(bottomborderrect, Color.FromArgb(53, 53, 53), Color.FromArgb(62, 62, 62), LinearGradientMode.Vertical)
        g.FillRectangle(bottomborderrectgb, bottomborderrect)
        g.DrawRectangle(New Pen(Brushes.Black), bottomborderrect)

        Select Case _state
            Case MouseState.None
                Dim bodyrectnonegb As New LinearGradientBrush(bodyrect, Color.FromArgb(61, 61, 61), Color.FromArgb(41, 41, 41), LinearGradientMode.Vertical)
                g.FillRectangle(bodyrectnonegb, bodyrect)
                g.DrawRectangle(New Pen(Brushes.Black), bodyrect)

                Dim topborderrectnonegb As New LinearGradientBrush(topborderrect, Color.FromArgb(53, 53, 53), Color.FromArgb(62, 62, 62), LinearGradientMode.Vertical)
                g.FillRectangle(topborderrectnonegb, topborderrect)
                g.DrawRectangle(New Pen(Brushes.Black), topborderrect)

                Dim bottomborderrectnonegb As New LinearGradientBrush(bottomborderrect, Color.FromArgb(53, 53, 53), Color.FromArgb(62, 62, 62), LinearGradientMode.Vertical)
                g.FillRectangle(bottomborderrectnonegb, bottomborderrect)
                g.DrawRectangle(New Pen(Brushes.Black), bottomborderrect)
                g.DrawString(Text, btnfont, New SolidBrush(Color.FromArgb(245, 245, 245)), bodyrect, New StringFormat() With {.Alignment = StringAlignment.Center, .LineAlignment = StringAlignment.Center})
            Case MouseState.Down
                g.TranslateTransform(1, 1)
                g.DrawString(Text, btnfont, New SolidBrush(Color.FromArgb(245, 245, 245)), bodyrect, New StringFormat() With {.Alignment = StringAlignment.Center, .LineAlignment = StringAlignment.Center})
            Case MouseState.Over
                Dim bodyrectovergb As New LinearGradientBrush(bodyrect, Color.FromArgb(41, 41, 41), Color.FromArgb(61, 61, 61), LinearGradientMode.Vertical)
                g.FillRectangle(bodyrectovergb, bodyrect)
                g.DrawRectangle(New Pen(Brushes.Black), bodyrect)
                g.DrawString(Text, btnfont, New SolidBrush(Color.FromArgb(245, 245, 245)), bodyrect, New StringFormat() With {.Alignment = StringAlignment.Center, .LineAlignment = StringAlignment.Center})
        End Select

        e.Graphics.DrawImage(b, New Point(0, 0))
        g.Dispose() : b.Dispose()
    End Sub
End Class

Public Class FaderPanel : Inherits ContainerControl

    Protected Overrides Sub OnResize(ByVal e As EventArgs)
        MyBase.OnResize(e)
        Invalidate()
    End Sub

    Sub New()
        MyBase.New()
        SetStyle(ControlStyles.AllPaintingInWmPaint Or ControlStyles.OptimizedDoubleBuffer Or ControlStyles.ResizeRedraw Or ControlStyles.UserPaint Or ControlStyles.SupportsTransparentBackColor, True)
        BackColor = Color.Transparent
        DoubleBuffered = True
        Size = New Size(240, 160)
    End Sub

    Protected Overrides Sub OnPaint(ByVal e As PaintEventArgs)
        Dim b As Bitmap = New Bitmap(Width, Height)
        Dim g As Graphics = Graphics.FromImage(b)

        Dim rect As Rectangle = New Rectangle(0, 0, Width - 1, Height - 1)

        MyBase.OnPaint(e)
        g.SmoothingMode = SmoothingMode.HighQuality
        g.InterpolationMode = InterpolationMode.HighQualityBicubic
        g.Clear(BackColor)

        Dim panelgb As New LinearGradientBrush(rect, Color.FromArgb(61, 61, 61), Color.FromArgb(41, 41, 41), 90S)
        g.FillRectangle(panelgb, rect)
        g.DrawRectangle(New Pen(Brushes.Black, 2), rect)

        e.Graphics.DrawImage(b, New Point(0, 0))
        g.Dispose() : b.Dispose()
    End Sub
End Class

Public Class FaderGroupBox : Inherits ContainerControl
    Private _showHeader As Boolean = True
    Public Property ShowHeader() As Boolean
        Get
            Return _showHeader
        End Get
        Set(ByVal v As Boolean)
            _showHeader = v
            Invalidate()
        End Set
    End Property

    Protected Overrides Sub OnResize(ByVal e As EventArgs)
        MyBase.OnResize(e)
        Invalidate()
    End Sub

    Sub New()
        MyBase.New()
        SetStyle(ControlStyles.AllPaintingInWmPaint Or ControlStyles.OptimizedDoubleBuffer Or ControlStyles.ResizeRedraw Or ControlStyles.UserPaint Or ControlStyles.SupportsTransparentBackColor, True)
        BackColor = Color.Transparent
        DoubleBuffered = True
        Size = New Size(240, 160)
    End Sub

    Protected Overrides Sub OnPaint(ByVal e As PaintEventArgs)
        Dim b As Bitmap = New Bitmap(Width, Height)
        Dim g As Graphics = Graphics.FromImage(b)

        Dim underlinepen As New Pen(Color.FromArgb(255, 255, 255), 2)
        Dim rect As Rectangle = New Rectangle(0, 0, Width - 1, Height - 1)

        MyBase.OnPaint(e)
        g.SmoothingMode = SmoothingMode.HighQuality
        g.InterpolationMode = InterpolationMode.HighQualityBicubic
        g.Clear(BackColor)

        Dim groupgb As New LinearGradientBrush(rect, Color.FromArgb(61, 61, 61), Color.FromArgb(41, 41, 41), 90S)
        g.FillRectangle(groupgb, rect)
        g.DrawRectangle(New Pen(Brushes.Black, 2), rect)

        If _showHeader Then
            g.DrawString(Text, New Font("Segoe UI", 11, FontStyle.Bold), New SolidBrush(Color.FromArgb(245, 245, 245)), New Rectangle(0, 3, Width - 1, 30), New StringFormat() With {.Alignment = StringAlignment.Center, .LineAlignment = StringAlignment.Center})
            g.DrawLine(underlinepen, 10, 30, Width - 11, 30)
        End If

        e.Graphics.DrawImage(b, New Point(0, 0))
        g.Dispose() : b.Dispose()
    End Sub
End Class

Public Class FaderProgressBar : Inherits Control
    Private _val As Integer
    Public Property Value() As Integer
        Get
            Return _val
        End Get
        Set(ByVal v As Integer)
            If v > _max Then
                _val = _max
            ElseIf v < 0 Then
                _val = 0
            Else
                _val = v
            End If
            Invalidate()
        End Set
    End Property

    Private _max As Integer
    Public Property Maximum() As Integer
        Get
            Return _max
        End Get
        Set(ByVal v As Integer)
            If v < 1 Then
                _max = 1
            Else
                _max = v
            End If
            If v < _val Then
                _val = _max
            End If
            Invalidate()
        End Set
    End Property

    Private _showPercentage As Boolean = False
    Public Property ShowPercentage() As Boolean
        Get
            Return _showPercentage
        End Get
        Set(ByVal v As Boolean)
            _showPercentage = v
            Invalidate()
        End Set
    End Property

    Sub New()
        MyBase.New()
        SetStyle(ControlStyles.AllPaintingInWmPaint Or ControlStyles.OptimizedDoubleBuffer Or ControlStyles.ResizeRedraw Or ControlStyles.UserPaint Or ControlStyles.SupportsTransparentBackColor, True)
        BackColor = Color.Transparent
        Size = New Size(250, 20)
        DoubleBuffered = True
        _max = 100
    End Sub

    Protected Overrides Sub OnPaint(e As Windows.Forms.PaintEventArgs)
        Dim b As Bitmap = New Bitmap(Width, Height)
        Dim g As Graphics = Graphics.FromImage(b)

        Dim percent As Integer = CInt((Width - 1) * (_val / _max))
        Dim outerrect As Rectangle = New Rectangle(0, 0, Width - 1, Height - 1)
        Dim innerrect As Rectangle = New Rectangle(4, 4, percent - 9, Height - 9)

        MyBase.OnPaint(e)
        g.Clear(BackColor)
        g.SmoothingMode = SmoothingMode.HighQuality
        g.FillPath(New SolidBrush(Color.FromArgb(41, 41, 41)), SMSPLUS.Draw.RoundRect(outerrect, 5))
        g.DrawPath(New Pen(Color.FromArgb(0, 0, 0), 2), SMSPLUS.Draw.RoundRect(outerrect, 5))

        If percent <> 0 Then
            g.FillPath(New SolidBrush(Color.FromArgb(128, 128, 128)), SMSPLUS.Draw.RoundRect(innerrect, 7))
        End If

        If _showPercentage Then
            g.DrawString(String.Format("{0}%", _val), New Font("Segoe UI", 10, FontStyle.Bold), New SolidBrush(Color.FromArgb(245, 245, 245)), New Rectangle(10, 1, Width - 1, Height - 1), New StringFormat With {.Alignment = StringAlignment.Center, .LineAlignment = StringAlignment.Center})
        End If

        e.Graphics.DrawImage(b, New Point(0, 0))
        g.Dispose() : b.Dispose()
    End Sub
End Class

Public Class FaderLabel : Inherits Control
    Private _border As Boolean = True
    Public Property Border As Boolean
        Get
            Return _border
        End Get
        Set(value As Boolean)
            _border = value
            Invalidate()
        End Set
    End Property

    Sub New()
        SetStyle(ControlStyles.AllPaintingInWmPaint Or ControlStyles.OptimizedDoubleBuffer Or ControlStyles.ResizeRedraw Or ControlStyles.UserPaint Or ControlStyles.SupportsTransparentBackColor, True)
        BackColor = Color.Transparent
        DoubleBuffered = True
        Size = New Size(150, 30)
    End Sub

    Protected Overrides Sub OnPaint(ByVal e As Windows.Forms.PaintEventArgs)
        Dim b As Bitmap = New Bitmap(Width, Height)
        Dim g As Graphics = Graphics.FromImage(b)

        Dim rect As Rectangle = New Rectangle(0, 0, Width - 1, Height - 1)

        MyBase.OnPaint(e)
        g.Clear(BackColor)
        g.SmoothingMode = SmoothingMode.HighQuality
        g.InterpolationMode = InterpolationMode.HighQualityBicubic
        g.FillPath(New SolidBrush(Color.FromArgb(41, 41, 41)), SMSPLUS.Draw.RoundRect(rect, 5))

        If Border = True Then
            g.FillPath(New SolidBrush(Color.FromArgb(41, 41, 41)), SMSPLUS.Draw.RoundRect(rect, 5))
            g.DrawPath(New Pen(Color.FromArgb(21, 21, 21), 2), SMSPLUS.Draw.RoundRect(rect, 5))
        End If

        g.DrawString(Text, New Font("Segoe UI", 11, FontStyle.Regular), New SolidBrush(Color.FromArgb(255, 255, 255)), rect, New StringFormat() With {.Alignment = StringAlignment.Center, .LineAlignment = StringAlignment.Center})

        e.Graphics.DrawImage(b, New Point(0, 0))
        g.Dispose() : b.Dispose()
    End Sub
End Class

Public Class FaderToggle : Inherits Control
    Private _check As Boolean
    Public Property Checked As Boolean
        Get
            Return _check
        End Get
        Set(value As Boolean)
            _check = value
            Invalidate()
        End Set
    End Property

    Sub New()
        MyBase.New()
        SetStyle(ControlStyles.AllPaintingInWmPaint Or ControlStyles.OptimizedDoubleBuffer Or ControlStyles.ResizeRedraw Or ControlStyles.UserPaint Or ControlStyles.SupportsTransparentBackColor, True)
        DoubleBuffered = True
        BackColor = Color.Transparent
        Size = New Size(80, 25)
    End Sub

    Protected Overrides Sub OnClick(ByVal e As EventArgs)
        If Not Checked Then Checked = True Else Checked = False
        MyBase.OnClick(e)
    End Sub

    Protected Overrides Sub OnPaint(ByVal e As PaintEventArgs)
        Dim b As Bitmap = New Bitmap(Width, Height)
        Dim g As Graphics = Graphics.FromImage(b)

        Dim outerrect As Rectangle = New Rectangle(0, 0, Width - 1, Height - 1)

        MyBase.OnPaint(e)
        g.Clear(BackColor)
        g.SmoothingMode = SmoothingMode.HighQuality
        g.InterpolationMode = InterpolationMode.HighQualityBicubic
        g.FillPath(New SolidBrush(Color.FromArgb(41, 41, 41)), SMSPLUS.Draw.RoundRect(outerrect, 5))
        g.DrawPath(New Pen(Color.FromArgb(0, 0, 0), 2), SMSPLUS.Draw.RoundRect(outerrect, 5))

        If Checked Then
            g.FillPath(New SolidBrush(Color.FromArgb(128, 128, 128)), SMSPLUS.Draw.RoundRect(New Rectangle(3, 3, CInt((Width / 2) - 3), Height - 7), 3))
            g.DrawString("ON", New Font("Segoe UI", 11, FontStyle.Bold), New SolidBrush(Color.FromArgb(0, 0, 0)), New Rectangle(2, 3, CInt((Width / 2) - 1), Height - 5), New StringFormat() With {.Alignment = StringAlignment.Center, .LineAlignment = StringAlignment.Center})
        Else
            g.FillPath(New SolidBrush(Color.FromArgb(128, 128, 128)), SMSPLUS.Draw.RoundRect(New Rectangle(CInt((Width / 2) - 2), 3, CInt((Width / 2) - 2), Height - 7), 3))
            g.DrawString("OFF", New Font("Segoe UI", 11, FontStyle.Bold), New SolidBrush(Color.FromArgb(0, 0, 0)), New Rectangle(CInt((Width / 2) - 2), 3, CInt((Width / 2) - 1), Height - 5), New StringFormat() With {.Alignment = StringAlignment.Center, .LineAlignment = StringAlignment.Center})
        End If

        e.Graphics.DrawImage(b, New Point(0, 0))
        g.Dispose() : b.Dispose()
    End Sub
End Class

Public Class FaderVerticalProgressBar : Inherits Control
    Private _val As Integer
    Public Property Value() As Integer
        Get
            Return _val
        End Get
        Set(ByVal v As Integer)
            If v > _max Then
                _val = _max
            ElseIf v < 0 Then
                _val = 0
            Else
                _val = v
            End If
            Invalidate()
        End Set
    End Property

    Private _max As Integer
    Public Property Maximum() As Integer
        Get
            Return _max
        End Get
        Set(ByVal v As Integer)
            If v < 1 Then
                _max = 1
            Else
                _max = v
            End If
            If v < _val Then
                _val = _max
            End If
            Invalidate()
        End Set
    End Property

    Private _showPercentage As Boolean = False
    Public Property ShowPercentage() As Boolean
        Get
            Return _showPercentage
        End Get
        Set(ByVal v As Boolean)
            _showPercentage = v
            Invalidate()
        End Set
    End Property

    Sub New()
        MyBase.New()
        SetStyle(ControlStyles.AllPaintingInWmPaint Or ControlStyles.OptimizedDoubleBuffer Or ControlStyles.ResizeRedraw Or ControlStyles.UserPaint Or ControlStyles.SupportsTransparentBackColor, True)
        BackColor = Color.Transparent
        Size = New Size(20, 250)
        DoubleBuffered = True
        _max = 100
    End Sub

    Protected Overrides Sub OnPaint(e As Windows.Forms.PaintEventArgs)
        Dim b As Bitmap = New Bitmap(Width, Height)
        Dim g As Graphics = Graphics.FromImage(b)

        Dim percent As Integer = CInt((Height - 1) * (_val / _max))
        Dim outerrect As Rectangle = New Rectangle(0, 0, Width - 1, Height - 1)
        Dim innerrect As Rectangle = New Rectangle(4, (Height - percent) + 4, Width - 9, percent - 9)

        MyBase.OnPaint(e)
        g.Clear(BackColor)
        g.SmoothingMode = SmoothingMode.HighQuality
        g.FillPath(New SolidBrush(Color.FromArgb(41, 41, 41)), SMSPLUS.Draw.RoundRect(outerrect, 5))
        g.DrawPath(New Pen(Color.FromArgb(0, 0, 0)), SMSPLUS.Draw.RoundRect(outerrect, 5))

        If percent <> 0 Then
            Try
                g.FillPath(New SolidBrush(Color.FromArgb(128, 128, 128)), SMSPLUS.Draw.RoundRect(innerrect, 7))
            Catch : End Try
        End If

        If _showPercentage Then
            g.DrawString(String.Format("{0}%", _val), New Font("Segoe UI", 8, FontStyle.Bold), New SolidBrush(Color.FromArgb(245, 245, 245)), outerrect, New StringFormat With {.Alignment = StringAlignment.Center, .LineAlignment = StringAlignment.Center})
        End If

        e.Graphics.DrawImage(b, New Point(0, 0))
        g.Dispose() : b.Dispose()
    End Sub
End Class

Public Class FaderHorizontalSeperator : Inherits Control
    Sub New()
        SetStyle(ControlStyles.AllPaintingInWmPaint Or ControlStyles.OptimizedDoubleBuffer Or ControlStyles.ResizeRedraw Or ControlStyles.UserPaint Or ControlStyles.SupportsTransparentBackColor, True)
        DoubleBuffered = True
        BackColor = Color.Transparent
        Size = New Size(200, 3)
    End Sub

    Protected Overrides Sub OnResize(ByVal e As EventArgs)
        MyBase.OnResize(e)
        Height = 3
    End Sub

    Protected Overrides Sub OnPaint(e As PaintEventArgs)
        Dim b As Bitmap = New Bitmap(Width, Height)
        Dim g As Graphics = Graphics.FromImage(b)

        Dim rect As Rectangle = New Rectangle(0, 0, Width - 1, Height - 1)

        MyBase.OnPaint(e)
        g.Clear(BackColor)
        g.SmoothingMode = SmoothingMode.HighQuality
        g.InterpolationMode = InterpolationMode.HighQualityBicubic
        g.FillPath(New SolidBrush(Color.FromArgb(41, 41, 41)), SMSPLUS.Draw.RoundRect(rect, 2))
        g.DrawPath(New Pen(Color.FromArgb(0, 0, 0)), SMSPLUS.Draw.RoundRect(rect, 2))

        e.Graphics.DrawImage(b, New Point(0, 0))
        g.Dispose() : b.Dispose()
    End Sub
End Class

Public Class FaderVerticalSeperator : Inherits Control
    Sub New()
        SetStyle(ControlStyles.AllPaintingInWmPaint Or ControlStyles.OptimizedDoubleBuffer Or ControlStyles.ResizeRedraw Or ControlStyles.UserPaint Or ControlStyles.SupportsTransparentBackColor, True)
        DoubleBuffered = True
        BackColor = Color.Transparent
        Size = New Size(3, 200)
    End Sub

    Protected Overrides Sub OnResize(ByVal e As EventArgs)
        MyBase.OnResize(e)
        Width = 3
    End Sub

    Protected Overrides Sub OnPaint(e As PaintEventArgs)
        Dim b As Bitmap = New Bitmap(Width, Height)
        Dim g As Graphics = Graphics.FromImage(b)

        Dim rect As Rectangle = New Rectangle(0, 0, Width - 1, Height - 1)

        MyBase.OnPaint(e)
        g.Clear(BackColor)
        g.SmoothingMode = SmoothingMode.HighQuality
        g.InterpolationMode = InterpolationMode.HighQualityBicubic
        g.FillPath(New SolidBrush(Color.FromArgb(41, 41, 41)), SMSPLUS.Draw.RoundRect(rect, 2))
        g.DrawPath(New Pen(Color.FromArgb(0, 0, 0)), SMSPLUS.Draw.RoundRect(rect, 2))

        e.Graphics.DrawImage(b, New Point(0, 0))
        g.Dispose() : b.Dispose()
    End Sub
End Class

Public Class FaderProgressButton : Inherits Control
    Dim _state As MouseState
    Private _val As Integer
    Public Property Value() As Integer
        Get
            Return _val
        End Get
        Set(ByVal v As Integer)
            If v > _max Then
                _val = _max
            ElseIf v < 0 Then
                _val = 0
            Else
                _val = v
            End If
            Invalidate()
        End Set
    End Property

    Private _max As Integer
    Public Property Maximum() As Integer
        Get
            Return _max
        End Get
        Set(ByVal v As Integer)
            If v < 1 Then
                _max = 1
            Else
                _max = v
            End If
            If v < _val Then
                _val = _max
            End If
            Invalidate()
        End Set
    End Property

    Protected Overrides Sub OnMouseDown(ByVal e As MouseEventArgs)
        MyBase.OnMouseDown(e)
        _state = MouseState.Down
        Invalidate()
    End Sub

    Protected Overrides Sub OnMouseUp(ByVal e As MouseEventArgs)
        MyBase.OnMouseUp(e)
        _state = MouseState.Over
        Invalidate()
    End Sub

    Protected Overrides Sub OnMouseEnter(ByVal e As EventArgs)
        MyBase.OnMouseEnter(e)
        _state = MouseState.Over
        Invalidate()
    End Sub

    Protected Overrides Sub OnMouseLeave(ByVal e As EventArgs)
        MyBase.OnMouseLeave(e)
        _state = MouseState.None
        Invalidate()
    End Sub
    Sub New()
        MyBase.New()
        SetStyle(ControlStyles.AllPaintingInWmPaint Or ControlStyles.OptimizedDoubleBuffer Or ControlStyles.ResizeRedraw Or ControlStyles.UserPaint Or ControlStyles.SupportsTransparentBackColor, True)
        BackColor = Color.Transparent
        Size = New Size(160, 50)
        DoubleBuffered = True
        _max = 100
    End Sub

    Protected Overrides Sub OnPaint(e As Windows.Forms.PaintEventArgs)
        Dim b As Bitmap = New Bitmap(Width, Height)
        Dim g As Graphics = Graphics.FromImage(b)

        Dim percent As Integer = CInt((Width - 1) * (_val / _max))
        Dim bodyrect As Rectangle = New Rectangle(1, 5, Width - 3, Height - 11)
        Dim topborderrect As Rectangle = New Rectangle(0, 0, Width - 1, 5)
        Dim bottomborderrect As Rectangle = New Rectangle(0, Height - 6, Width - 1, 5)
        Dim progressinnerrect As Rectangle = New Rectangle(1, 5, percent - 3, Height - 11)
        Dim btnfont As New Font("Segoe UI", 11, FontStyle.Bold)

        MyBase.OnPaint(e)
        g.SmoothingMode = SmoothingMode.HighQuality
        g.InterpolationMode = InterpolationMode.HighQualityBicubic
        g.TextRenderingHint = Drawing.Text.TextRenderingHint.AntiAlias
        g.Clear(BackColor)

        Dim bodyrectgb As New LinearGradientBrush(bodyrect, Color.FromArgb(61, 61, 61), Color.FromArgb(41, 41, 41), LinearGradientMode.Vertical)
        g.FillRectangle(bodyrectgb, bodyrect)
        g.DrawRectangle(New Pen(Brushes.Black), bodyrect)

        Dim topborderrectgb As New LinearGradientBrush(topborderrect, Color.FromArgb(53, 53, 53), Color.FromArgb(62, 62, 62), LinearGradientMode.Vertical)
        g.FillRectangle(topborderrectgb, topborderrect)
        g.DrawRectangle(New Pen(Brushes.Black), topborderrect)

        Dim bottomborderrectgb As New LinearGradientBrush(bottomborderrect, Color.FromArgb(53, 53, 53), Color.FromArgb(62, 62, 62), LinearGradientMode.Vertical)
        g.FillRectangle(bottomborderrectgb, bottomborderrect)
        g.DrawRectangle(New Pen(Brushes.Black), bottomborderrect)

        Select Case _state
            Case MouseState.None
                Dim bodyrectnonegb As New LinearGradientBrush(bodyrect, Color.FromArgb(61, 61, 61), Color.FromArgb(41, 41, 41), LinearGradientMode.Vertical)
                g.FillRectangle(bodyrectnonegb, bodyrect)
                g.DrawRectangle(New Pen(Brushes.Black), bodyrect)

                Dim topborderrectnonegb As New LinearGradientBrush(topborderrect, Color.FromArgb(53, 53, 53), Color.FromArgb(62, 62, 62), LinearGradientMode.Vertical)
                g.FillRectangle(topborderrectnonegb, topborderrect)
                g.DrawRectangle(New Pen(Brushes.Black), topborderrect)

                Dim bottomborderrectnonegb As New LinearGradientBrush(bottomborderrect, Color.FromArgb(53, 53, 53), Color.FromArgb(62, 62, 62), LinearGradientMode.Vertical)
                g.FillRectangle(bottomborderrectnonegb, bottomborderrect)
                g.DrawRectangle(New Pen(Brushes.Black), bottomborderrect)

                g.DrawString(Text, btnfont, Brushes.White, bodyrect, New StringFormat() With {.Alignment = StringAlignment.Center, .LineAlignment = StringAlignment.Center})
            Case MouseState.Down
                g.TranslateTransform(1, 1)
                g.DrawString(Text, btnfont, Brushes.White, bodyrect, New StringFormat() With {.Alignment = StringAlignment.Center, .LineAlignment = StringAlignment.Center})
            Case MouseState.Over
                Dim bodyrectovergb As New LinearGradientBrush(bodyrect, Color.FromArgb(41, 41, 41), Color.FromArgb(61, 61, 61), LinearGradientMode.Vertical)
                g.FillRectangle(bodyrectovergb, bodyrect)
                g.DrawRectangle(New Pen(Brushes.Black), bodyrect)
                g.DrawString(Text, btnfont, Brushes.White, bodyrect, New StringFormat() With {.Alignment = StringAlignment.Center, .LineAlignment = StringAlignment.Center})
        End Select

        If percent <> 0 Then
            Dim progressgb As New LinearGradientBrush(progressinnerrect, Color.FromArgb(61, 61, 61), Color.FromArgb(41, 41, 41), 180S)
            g.FillPath(progressgb, SMSPLUS.Draw.RoundRect(progressinnerrect, 3))
            g.DrawPath(New Pen(Color.FromArgb(0, 0, 0)), SMSPLUS.Draw.RoundRect(progressinnerrect, 3))
            g.DrawString(Text, btnfont, Brushes.White, bodyrect, New StringFormat() With {.Alignment = StringAlignment.Center, .LineAlignment = StringAlignment.Center})
        End If

        e.Graphics.DrawImage(b, New Point(0, 0))
        g.Dispose() : b.Dispose()
    End Sub
End Class

Public Class FaderTabControl : Inherits TabControl

    Sub New()
        SetStyle(ControlStyles.AllPaintingInWmPaint Or ControlStyles.OptimizedDoubleBuffer Or ControlStyles.ResizeRedraw Or ControlStyles.UserPaint Or ControlStyles.SupportsTransparentBackColor, True)
        BackColor = Color.Transparent
        DoubleBuffered = True
        ItemSize = New Size(100, 35)
    End Sub

    Protected Overrides Sub OnPaint(ByVal e As Windows.Forms.PaintEventArgs)
        Dim b As Bitmap = New Bitmap(Width, Height)
        Dim g As Graphics = Graphics.FromImage(b)

        Dim rect As Rectangle = New Rectangle(0, 0, Width - 1, Height - 1)
        Dim selectedtabgb As New LinearGradientBrush(rect, Color.FromArgb(61, 61, 61), Color.FromArgb(41, 41, 41), 90S)
        Dim nonselectedtabgb As New LinearGradientBrush(rect, Color.FromArgb(81, 81, 81), Color.FromArgb(61, 61, 61), 90S)

        Try : SelectedTab.BackColor = Color.FromArgb(61, 61, 61) : Catch : End Try

        MyBase.OnPaint(e)
        g.Clear(BackColor)
        g.FillRectangle(selectedtabgb, rect)
        g.DrawRectangle(New Pen(Brushes.Black, 2), rect)

        For i = 0 To TabCount - 1
            Dim textRectangle As Rectangle = New Rectangle(New Point(GetTabRect(i).Location.X + 1, GetTabRect(i).Location.Y), New Size(GetTabRect(i).Width - 1, GetTabRect(i).Height))
            If i = SelectedIndex Then
                Dim tabrect As Rectangle = New Rectangle(New Point(GetTabRect(i).Location.X + 1, GetTabRect(i).Location.Y + 1), New Size(GetTabRect(i).Width - 1, GetTabRect(i).Height - 2))
                g.FillPath(selectedtabgb, SMSPLUS.Draw.RoundRect(tabrect, 2))
                g.DrawPath(New Pen(Brushes.Black), SMSPLUS.Draw.RoundRect(tabrect, 2))
                g.DrawString(TabPages(i).Text, New Font("Segoe UI", 10, FontStyle.Bold), New SolidBrush(Color.FromArgb(245, 245, 245)), textRectangle, New StringFormat With {.LineAlignment = StringAlignment.Center, .Alignment = StringAlignment.Center})
            Else
                Dim tabrect As Rectangle = New Rectangle(New Point(GetTabRect(i).Location.X + 1, GetTabRect(i).Location.Y + 4), New Size(GetTabRect(i).Width - 1, GetTabRect(i).Height - 5))
                g.FillPath(nonselectedtabgb, SMSPLUS.Draw.RoundRect(tabrect, 2))
                g.DrawPath(New Pen(Brushes.Black), SMSPLUS.Draw.RoundRect(tabrect, 2))
                g.DrawString(TabPages(i).Text, New Font("Segoe UI", 10, FontStyle.Regular), New SolidBrush(Color.FromArgb(245, 245, 245)), textRectangle, New StringFormat With {.LineAlignment = StringAlignment.Center, .Alignment = StringAlignment.Center})
            End If
        Next

        e.Graphics.DrawImage(b, New Point(0, 0))
        g.Dispose() : b.Dispose()
    End Sub
End Class

Public Class FaderVerticalTabControl : Inherits TabControl

    Sub New()
        SetStyle(ControlStyles.AllPaintingInWmPaint Or ControlStyles.OptimizedDoubleBuffer Or ControlStyles.ResizeRedraw Or ControlStyles.UserPaint Or ControlStyles.SupportsTransparentBackColor, True)
        BackColor = Color.Transparent
        DoubleBuffered = True
        SizeMode = TabSizeMode.Fixed
        Alignment = TabAlignment.Left
        ItemSize = New Size(35, 100)
    End Sub

    Protected Overrides Sub OnPaint(ByVal e As Windows.Forms.PaintEventArgs)
        Dim b As Bitmap = New Bitmap(Width, Height)
        Dim g As Graphics = Graphics.FromImage(b)

        Dim rect As Rectangle = New Rectangle(0, 0, Width - 1, Height - 1)
        Dim selectedtabgb As New LinearGradientBrush(rect, Color.FromArgb(61, 61, 61), Color.FromArgb(41, 41, 41), 90S)
        Dim nonselectedtabgb As New LinearGradientBrush(rect, Color.FromArgb(81, 81, 81), Color.FromArgb(61, 61, 61), 90S)

        Try : SelectedTab.BackColor = Color.FromArgb(61, 61, 61) : Catch : End Try

        MyBase.OnPaint(e)
        g.Clear(BackColor)
        g.FillRectangle(selectedtabgb, rect)
        g.DrawRectangle(New Pen(Brushes.Black, 2), rect)

        For i = 0 To TabCount - 1
            Dim textRectangle As Rectangle = New Rectangle(New Point(GetTabRect(i).Location.X + 1, GetTabRect(i).Location.Y), New Size(GetTabRect(i).Width - 1, GetTabRect(i).Height))
            If i = SelectedIndex Then
                Dim tabrect As Rectangle = New Rectangle(New Point(GetTabRect(i).Location.X + 1, GetTabRect(i).Location.Y + 1), New Size(GetTabRect(i).Width - 1, GetTabRect(i).Height - 1))
                g.FillPath(selectedtabgb, SMSPLUS.Draw.RoundRect(tabrect, 2))
                g.DrawPath(New Pen(Brushes.Black), SMSPLUS.Draw.RoundRect(tabrect, 2))
                g.DrawString(TabPages(i).Text, New Font("Segoe UI", 10, FontStyle.Bold), New SolidBrush(Color.FromArgb(245, 245, 245)), textRectangle, New StringFormat With {.LineAlignment = StringAlignment.Center, .Alignment = StringAlignment.Center})
            Else
                Dim tabrect As Rectangle = New Rectangle(New Point(GetTabRect(i).Location.X + 4, GetTabRect(i).Location.Y + 1), New Size(GetTabRect(i).Width - 4, GetTabRect(i).Height - 1))
                g.FillPath(nonselectedtabgb, SMSPLUS.Draw.RoundRect(tabrect, 2))
                g.DrawPath(New Pen(Brushes.Black), SMSPLUS.Draw.RoundRect(tabrect, 2))
                g.DrawString(TabPages(i).Text, New Font("Segoe UI", 10, FontStyle.Regular), New SolidBrush(Color.FromArgb(245, 245, 245)), textRectangle, New StringFormat With {.LineAlignment = StringAlignment.Center, .Alignment = StringAlignment.Center})
            End If
        Next

        e.Graphics.DrawImage(b, New Point(0, 0))
        g.Dispose() : b.Dispose()
    End Sub
End Class

Public Class FaderTextBox : Inherits Control
    Dim WithEvents _tb As New TextBox
    Private _allowpassword As Boolean = False
    Public Shadows Property UseSystemPasswordChar() As Boolean
        Get
            Return _allowpassword
        End Get
        Set(ByVal value As Boolean)
            _tb.UseSystemPasswordChar = UseSystemPasswordChar
            _allowpassword = value
            Invalidate()
        End Set
    End Property

    Private _maxChars As Integer = 32767
    Public Shadows Property MaxLength() As Integer
        Get
            Return _maxChars
        End Get
        Set(ByVal value As Integer)
            _maxChars = value
            _tb.MaxLength = MaxLength
            Invalidate()
        End Set
    End Property

    Private _textAlignment As HorizontalAlignment
    Public Shadows Property TextAlign() As HorizontalAlignment
        Get
            Return _textAlignment
        End Get
        Set(ByVal value As HorizontalAlignment)
            _textAlignment = value
            Invalidate()
        End Set
    End Property

    Private _multiLine As Boolean = False
    Public Shadows Property MultiLine() As Boolean
        Get
            Return _multiLine
        End Get
        Set(ByVal value As Boolean)
            _multiLine = value
            _tb.Multiline = value
            OnResize(EventArgs.Empty)
            Invalidate()
        End Set
    End Property

    Private _readOnly As Boolean = False
    Public Shadows Property [ReadOnly]() As Boolean
        Get
            Return _readOnly
        End Get
        Set(ByVal value As Boolean)
            _readOnly = value
            If _tb IsNot Nothing Then
                _tb.ReadOnly = value
            End If
        End Set
    End Property

    Protected Overrides Sub OnTextChanged(ByVal e As EventArgs)
        MyBase.OnTextChanged(e)
        Invalidate()
    End Sub

    Protected Overrides Sub OnBackColorChanged(ByVal e As EventArgs)
        MyBase.OnBackColorChanged(e)
        Invalidate()
    End Sub

    Protected Overrides Sub OnForeColorChanged(ByVal e As EventArgs)
        MyBase.OnForeColorChanged(e)
        _tb.ForeColor = ForeColor
        Invalidate()
    End Sub

    Protected Overrides Sub OnFontChanged(ByVal e As EventArgs)
        MyBase.OnFontChanged(e)
        _tb.Font = Font
    End Sub

    Protected Overrides Sub OnGotFocus(ByVal e As EventArgs)
        MyBase.OnGotFocus(e)
        _tb.Focus()
    End Sub

    Private Sub TextChangeTb() Handles _tb.TextChanged
        Text = _tb.Text
    End Sub

    Private Sub TextChng() Handles MyBase.TextChanged
        _tb.Text = Text
    End Sub

    Public Sub NewTextBox()
        With _tb
            .Text = String.Empty
            .BackColor = Color.FromArgb(61, 61, 61)
            .ForeColor = ForeColor
            .TextAlign = HorizontalAlignment.Center
            .BorderStyle = BorderStyle.None
            .Location = New Point(3, 3)
            .Font = New Font("Segoe UI", 11, FontStyle.Regular)
            .Size = New Size(Width - 3, Height - 3)
            .UseSystemPasswordChar = UseSystemPasswordChar
        End With
    End Sub

    Sub New()
        MyBase.New()
        NewTextBox()
        Controls.Add(_tb)
        SetStyle(ControlStyles.AllPaintingInWmPaint Or ControlStyles.OptimizedDoubleBuffer Or ControlStyles.ResizeRedraw Or ControlStyles.UserPaint Or ControlStyles.SupportsTransparentBackColor, True)
        DoubleBuffered = True
        BackColor = Color.FromArgb(61, 61, 61)
        ForeColor = Color.FromArgb(245, 245, 245)
        Font = New Font("Segoe UI", 11, FontStyle.Regular)
        Size = New Size(200, 30)
    End Sub

    Protected Overrides Sub OnPaint(ByVal e As Windows.Forms.PaintEventArgs)
        Dim b As Bitmap = New Bitmap(Width, Height)
        Dim g As Graphics = Graphics.FromImage(b)

        Dim rect As Rectangle = New Rectangle(0, 0, Width - 1, Height - 1)

        MyBase.OnPaint(e)
        g.SmoothingMode = SmoothingMode.HighQuality
        g.InterpolationMode = InterpolationMode.HighQualityBicubic
        g.TextRenderingHint = Drawing.Text.TextRenderingHint.AntiAlias

        With _tb
            .TextAlign = TextAlign
            .UseSystemPasswordChar = UseSystemPasswordChar
        End With

        g.FillPath(New SolidBrush(Color.FromArgb(61, 61, 61)), SMSPLUS.Draw.RoundRect(rect, 1))
        g.DrawPath(New Pen(Brushes.Black, 2), SMSPLUS.Draw.RoundRect(rect, 1))

        e.Graphics.DrawImage(b, New Point(0, 0))
        g.Dispose() : b.Dispose()
    End Sub

    Protected Overrides Sub OnResize(e As EventArgs)
        MyBase.OnResize(e)
        If Not MultiLine Then
            Dim tbheight As Integer = _tb.Height
            _tb.Location = New Point(10, CType(((Height / 2) - (tbheight / 2) - 1), Integer))
            _tb.Size = New Size(Width - 20, tbheight)
        Else
            _tb.Location = New Point(10, 10)
            _tb.Size = New Size(Width - 20, Height - 20)
        End If
    End Sub
End Class

Public Class FaderComboBox : Inherits ComboBox
    Private _startIndex As Integer = 0
    Private Property StartIndex As Integer
        Get
            Return _startIndex
        End Get
        Set(ByVal value As Integer)
            _startIndex = value
            Try
                SelectedIndex = value
            Catch
            End Try
            Invalidate()
        End Set
    End Property

    Sub ReplaceItem(ByVal sender As System.Object, ByVal e As Windows.Forms.DrawItemEventArgs) Handles Me.DrawItem
        e.DrawBackground()
        Dim sitemrect As New LinearGradientBrush(New Rectangle(e.Bounds.X, e.Bounds.Y, e.Bounds.Width, e.Bounds.Height), Color.FromArgb(61, 61, 61), Color.FromArgb(41, 41, 41), 90S)
        Dim itemrect As New LinearGradientBrush(New Rectangle(e.Bounds.X, e.Bounds.Y, e.Bounds.Width, e.Bounds.Height), Color.FromArgb(81, 81, 81), Color.FromArgb(61, 61, 61), 90S)

        Try
            If (e.State And DrawItemState.Selected) = DrawItemState.Selected Then
                e.Graphics.FillPath(sitemrect, SMSPLUS.Draw.RoundRect(New Rectangle(e.Bounds.X + 3, e.Bounds.Y, e.Bounds.Width - 7, e.Bounds.Height), 2))
                e.Graphics.DrawPath(New Pen(Brushes.Black), SMSPLUS.Draw.RoundRect(New Rectangle(e.Bounds.X + 3, e.Bounds.Y, e.Bounds.Width - 7, e.Bounds.Height), 2))
            Else
                e.Graphics.FillPath(itemrect, SMSPLUS.Draw.RoundRect(New Rectangle(e.Bounds.X, e.Bounds.Y, e.Bounds.Width, e.Bounds.Height), 2))
                e.Graphics.DrawPath(New Pen(Brushes.Black), SMSPLUS.Draw.RoundRect(New Rectangle(e.Bounds.X, e.Bounds.Y, e.Bounds.Width, e.Bounds.Height), 2))
            End If
            e.Graphics.DrawString(GetItemText(Items(e.Index)), e.Font, New SolidBrush(Color.FromArgb(245, 245, 245)), New Rectangle(e.Bounds.X, e.Bounds.Y, e.Bounds.Width, e.Bounds.Height), New StringFormat() With {.Alignment = StringAlignment.Center, .LineAlignment = StringAlignment.Center})
        Catch : End Try
    End Sub

    Sub New()
        MyBase.New()
        SetStyle(ControlStyles.AllPaintingInWmPaint Or ControlStyles.OptimizedDoubleBuffer Or ControlStyles.ResizeRedraw Or ControlStyles.UserPaint Or ControlStyles.SupportsTransparentBackColor, True)
        DrawMode = Windows.Forms.DrawMode.OwnerDrawFixed
        BackColor = Color.Transparent
        DropDownStyle = ComboBoxStyle.DropDownList
        StartIndex = 0
        ItemHeight = 25
        DoubleBuffered = True
        Width = 200
    End Sub

    Protected Overrides Sub OnResize(e As EventArgs)
        MyBase.OnResize(e)
        Height = 20
    End Sub

    Protected Overrides Sub OnPaint(ByVal e As PaintEventArgs)
        Dim b As Bitmap = New Bitmap(Width, Height)
        Dim g As Graphics = Graphics.FromImage(b)

        Dim rect As Rectangle = New Rectangle(0, 0, Width - 1, Height - 1)

        MyBase.OnPaint(e)
        g.Clear(BackColor)

        Dim rectgb As New LinearGradientBrush(rect, Color.FromArgb(61, 61, 61), Color.FromArgb(41, 41, 41), 90S)
        g.FillPath(rectgb, SMSPLUS.Draw.RoundRect(rect, 3))
        g.DrawPath(New Pen(Brushes.Black, 2), SMSPLUS.Draw.RoundRect(rect, 3))
        g.SetClip(SMSPLUS.Draw.RoundRect(rect, 3))
        g.FillPath(rectgb, SMSPLUS.Draw.RoundRect(rect, 3))
        g.DrawPath(New Pen(Brushes.Black, 2), SMSPLUS.Draw.RoundRect(rect, 3))
        g.ResetClip()

        g.DrawLine(New Pen(Color.FromArgb(245, 245, 245)), Width - 9, 10, Width - 22, 10)
        g.DrawLine(New Pen(Color.FromArgb(245, 245, 245)), Width - 9, 11, Width - 22, 11)
        g.DrawLine(New Pen(Color.FromArgb(245, 245, 245)), Width - 9, 15, Width - 22, 15)
        g.DrawLine(New Pen(Color.FromArgb(245, 245, 245)), Width - 9, 16, Width - 22, 16)
        g.DrawLine(New Pen(Color.FromArgb(245, 245, 245)), Width - 9, 20, Width - 22, 20)
        g.DrawLine(New Pen(Color.FromArgb(245, 245, 245)), Width - 9, 21, Width - 22, 21)
        g.DrawLine(New Pen(Color.FromArgb(245, 245, 245)), New Point(Width - 29, 7), New Point(Width - 29, Height - 7))
        g.DrawLine(New Pen(Color.FromArgb(245, 245, 245)), New Point(Width - 30, 7), New Point(Width - 30, Height - 7))

        Try
            g.DrawString(Text, New Font("Segoe UI", 11, FontStyle.Bold), New SolidBrush(Color.FromArgb(245, 245, 245)), rect, New StringFormat() With {.Alignment = StringAlignment.Center, .LineAlignment = StringAlignment.Center})
        Catch : End Try

        e.Graphics.DrawImage(b, New Point(0, 0))
        g.Dispose() : b.Dispose()
    End Sub
End Class

Public Class FaderListBox : Inherits ListBox
    Sub New()
        SetStyle(ControlStyles.AllPaintingInWmPaint Or ControlStyles.OptimizedDoubleBuffer Or
                 ControlStyles.SupportsTransparentBackColor, True)
        BackColor = Color.Transparent
        DoubleBuffered = True
        DrawMode = Windows.Forms.DrawMode.OwnerDrawFixed
        ForeColor = Color.White
        BackColor = Color.FromArgb(61, 61, 61)
        BorderStyle = Windows.Forms.BorderStyle.None
        ItemHeight = 20
    End Sub
    Protected Overrides Sub OnPaint(e As PaintEventArgs)
        Dim b As Bitmap = New Bitmap(Width, Height)
        Dim g As Graphics = Graphics.FromImage(b)

        Dim rect As Rectangle = New Rectangle(0, 0, Width - 1, Height - 1)

        MyBase.OnPaint(e)
        g.Clear(Color.Transparent)
        g.FillPath(New SolidBrush(Color.FromArgb(61, 61, 61)), SMSPLUS.Draw.RoundRect(rect, 3))
        g.DrawPath(New Pen(Color.Black, 2), SMSPLUS.Draw.RoundRect(rect, 3))

        e.Graphics.DrawImage(b, New Point(0, 0))
        g.Dispose() : b.Dispose()
    End Sub
    Protected Overrides Sub OnDrawItem(e As DrawItemEventArgs)
        Dim b As Bitmap = New Bitmap(Width, Height)
        Dim g As Graphics = Graphics.FromImage(b)

        g.TextRenderingHint = TextRenderingHint.AntiAlias
        g.SmoothingMode = SmoothingMode.HighQuality

        g.SetClip(SMSPLUS.Draw.RoundRect(New Rectangle(0, 0, Width, Height), 3))
        g.Clear(Color.Transparent)
        g.FillRectangle(New SolidBrush(BackColor), New Rectangle(e.Bounds.X, e.Bounds.Y - 1, e.Bounds.Width, e.Bounds.Height + 3))

        If e.State.ToString().Contains("Selected,") Then
            Dim selectgb As New LinearGradientBrush(New Rectangle(e.Bounds.X, e.Bounds.Y + 1, e.Bounds.Width, e.Bounds.Height), Color.FromArgb(61, 61, 61), Color.FromArgb(41, 41, 31), 90S)
            g.FillRectangle(selectgb, New Rectangle(e.Bounds.X, e.Bounds.Y + 1, e.Bounds.Width, e.Bounds.Height))
            g.DrawRectangle(New Pen(Color.FromArgb(128, 128, 128)), New Rectangle(e.Bounds.X, e.Bounds.Y + 1, e.Bounds.Width, e.Bounds.Height))
            Try
                g.DrawString(Items(e.Index).ToString(), New Font("Segoe UI", 10, FontStyle.Bold), New SolidBrush(Color.FromArgb(245, 245, 245)), New Rectangle(e.Bounds.X + 3, e.Bounds.Y + 1, e.Bounds.Width, e.Bounds.Height), New StringFormat With {.LineAlignment = StringAlignment.Center, .Alignment = StringAlignment.Center})
            Catch : End Try
        Else
            Dim nonselectgb As New LinearGradientBrush(New Rectangle(e.Bounds.X, e.Bounds.Y + 1, e.Bounds.Width, e.Bounds.Height), Color.FromArgb(81, 81, 81), Color.FromArgb(61, 61, 61), 90S)
            g.FillRectangle(nonselectgb, e.Bounds)
            Try
                g.DrawString(Items(e.Index).ToString(), New Font("Segoe UI", 10, FontStyle.Regular), New SolidBrush(Color.FromArgb(245, 245, 245)), New Rectangle(e.Bounds.X + 3, e.Bounds.Y + 1, e.Bounds.Width, e.Bounds.Height), New StringFormat With {.LineAlignment = StringAlignment.Center, .Alignment = StringAlignment.Center})
            Catch : End Try
        End If

        g.DrawPath(New Pen(Color.FromArgb(61, 61, 61), 2), SMSPLUS.Draw.RoundRect(New Rectangle(0, 0, Width - 1, Height - 1), 1))

        e.Graphics.DrawImage(b, New Point(0, 0))
        g.Dispose() : b.Dispose()
    End Sub
End Class

Public Class FaderAlertBox : Inherits Control
    Dim WithEvents _mytimer As Windows.Forms.Timer

    Enum AlertStyle
        [Info]
        [Success]
        [Error]
    End Enum

    Private _style As AlertStyle
    Public Property Style As AlertStyle
        Get
            Return _style
        End Get
        Set(value As AlertStyle)
            _style = value
        End Set
    End Property

    Private _text As String
    Overrides Property Text As String
        Get
            Return MyBase.Text
        End Get
        Set(ByVal value As String)
            MyBase.Text = value
            If _text IsNot Nothing Then
                _text = value
            End If
        End Set
    End Property

    Shadows Property Visible As Boolean
        Get
            Return MyBase.Visible = False
        End Get
        Set(value As Boolean)
            MyBase.Visible = value
        End Set
    End Property

    Protected Overrides Sub OnTextChanged(e As EventArgs)
        MyBase.OnTextChanged(e) : Invalidate()
    End Sub

    Public Sub ShowAlertBox(mystyle As AlertStyle, str As String, interval As Integer)
        _style = mystyle
        Text = str
        Visible = True
        _mytimer.Interval = interval
        _mytimer.Enabled = True
    End Sub

    Sub New()
        MyBase.New()
        SetStyle(ControlStyles.UserPaint Or ControlStyles.ResizeRedraw Or ControlStyles.SupportsTransparentBackColor, True)
        BackColor = Color.Transparent
        Size = New Size(500, 30)
        DoubleBuffered = True
        _mytimer = New Timer()
    End Sub

    Protected Overrides Sub OnPaint(e As PaintEventArgs)
        Dim b As Bitmap = New Bitmap(Width, Height)
        Dim g As Graphics = Graphics.FromImage(b)

        Dim rect As New Rectangle(0, 0, Width - 1, Height - 1)
        Dim textrect As New Rectangle(20, 5, Width - 21, Height - 11)
        Dim logorect As New Rectangle(7, 7, 20, 20)
        Dim logocirclerect As New Rectangle(5, 5, 20, 20)

        g.TextRenderingHint = TextRenderingHint.AntiAlias
        g.SmoothingMode = SmoothingMode.HighQuality
        g.Clear(Color.Transparent)

        MyBase.OnPaint(e)
        Select Case _style
            Case AlertStyle.Success
                g.FillRectangle(New SolidBrush(Color.FromArgb(30, 0, 255, 0)), rect)
                g.DrawRectangle(New Pen(Color.FromArgb(0, 0, 0)), rect)
                g.DrawString(Text, New Font("Segoe UI", 11, FontStyle.Regular), New SolidBrush(Color.FromArgb(245, 245, 245)), textrect, New StringFormat() With {.Alignment = StringAlignment.Center, .LineAlignment = StringAlignment.Center})
                g.DrawString("ü", New Font("Wingdings", 14), New SolidBrush(Color.FromArgb(245, 245, 245)), logorect, New StringFormat() With {.Alignment = StringAlignment.Center, .LineAlignment = StringAlignment.Center})
            Case AlertStyle.Error
                g.FillRectangle(New SolidBrush(Color.FromArgb(30, 255, 0, 0)), rect)
                g.DrawRectangle(New Pen(Color.FromArgb(0, 0, 0)), rect)
                g.DrawString(Text, New Font("Segoe UI", 11, FontStyle.Regular), New SolidBrush(Color.FromArgb(245, 245, 245)), textrect, New StringFormat() With {.Alignment = StringAlignment.Center, .LineAlignment = StringAlignment.Center})
                g.DrawString("r", New Font("Marlett", 10), New SolidBrush(Color.FromArgb(245, 245, 245)), logorect, New StringFormat() With {.Alignment = StringAlignment.Center, .LineAlignment = StringAlignment.Center})
            Case AlertStyle.Info
                g.FillRectangle(New SolidBrush(Color.FromArgb(30, 0, 0, 255)), rect)
                g.DrawRectangle(New Pen(Color.FromArgb(0, 0, 0)), rect)
                g.DrawString(Text, New Font("Segoe UI", 11, FontStyle.Regular), New SolidBrush(Color.FromArgb(245, 245, 245)), textrect, New StringFormat() With {.Alignment = StringAlignment.Center, .LineAlignment = StringAlignment.Center})
                g.DrawString("i", New Font("Segoe UI", 12), New SolidBrush(Color.FromArgb(245, 245, 245)), logorect, New StringFormat() With {.Alignment = StringAlignment.Center, .LineAlignment = StringAlignment.Center})
        End Select

        g.DrawEllipse(New Pen(Color.FromArgb(255, 255, 255)), logocirclerect)

        e.Graphics.DrawImage(b, New Point(0, 0))
        g.Dispose() : b.Dispose()
    End Sub

    Private Sub MyTimer_Tick(ByVal sender As Object, ByVal e As EventArgs) Handles _mytimer.Tick
        Visible = False
        _mytimer.Enabled = False
    End Sub
End Class

Public Class FaderNotifyButton : Inherits Control
    Dim _state As MouseState

    Private _showNotification As Boolean = True
    Public Property ShowNotification As Boolean
        Get
            Return _showNotification
        End Get
        Set(value As Boolean)
            _showNotification = value
            Invalidate()
        End Set
    End Property

    Private _notifyValue As UInteger = 0
    Public Property NotifyValue As UInteger
        Get
            Return _notifyValue
        End Get
        Set(value As UInteger)
            _notifyValue = value
            Invalidate()
        End Set
    End Property
    Protected Overrides Sub OnMouseDown(ByVal e As MouseEventArgs)
        MyBase.OnMouseDown(e)
        _state = MouseState.Down
        Invalidate()
    End Sub

    Protected Overrides Sub OnMouseUp(ByVal e As MouseEventArgs)
        MyBase.OnMouseUp(e)
        _state = MouseState.Over
        Invalidate()
    End Sub

    Protected Overrides Sub OnMouseEnter(ByVal e As EventArgs)
        MyBase.OnMouseEnter(e)
        _state = MouseState.Over
        Invalidate()
    End Sub

    Protected Overrides Sub OnMouseLeave(ByVal e As EventArgs)
        MyBase.OnMouseLeave(e)
        _state = MouseState.None
        Invalidate()
    End Sub

    Sub New()
        SetStyle(ControlStyles.AllPaintingInWmPaint Or ControlStyles.OptimizedDoubleBuffer Or ControlStyles.ResizeRedraw Or ControlStyles.UserPaint Or ControlStyles.SupportsTransparentBackColor, True)
        BackColor = Color.Transparent
        DoubleBuffered = True
        Size = New Size(160, 40)
        _state = MouseState.None
    End Sub

    Protected Overrides Sub OnPaint(ByVal e As PaintEventArgs)
        Dim b As Bitmap = New Bitmap(Width, Height)
        Dim g As Graphics = Graphics.FromImage(b)

        Dim rect As Rectangle = New Rectangle(0, 10, Width - 1, Height - 11)
        Dim btnfont As New Font("Segoe UI", 11, FontStyle.Bold)

        MyBase.OnPaint(e)
        g.SmoothingMode = SmoothingMode.HighQuality
        g.InterpolationMode = InterpolationMode.HighQualityBicubic
        g.TextRenderingHint = Drawing.Text.TextRenderingHint.AntiAlias
        g.Clear(BackColor)

        Dim rectgb As New LinearGradientBrush(rect, Color.FromArgb(61, 61, 61), Color.FromArgb(41, 41, 41), LinearGradientMode.Vertical)
        g.FillRectangle(rectgb, rect)
        g.DrawRectangle(New Pen(Brushes.Black), rect)

        Select Case _state
            Case MouseState.None
                Dim rectnonegb As New LinearGradientBrush(rect, Color.FromArgb(61, 61, 61), Color.FromArgb(41, 41, 41), LinearGradientMode.Vertical)
                g.FillRectangle(rectnonegb, rect)
                g.DrawRectangle(New Pen(Brushes.Black), rect)
                g.DrawString(Text, btnfont, New SolidBrush(Color.FromArgb(245, 245, 245)), rect, New StringFormat() With {.Alignment = StringAlignment.Center, .LineAlignment = StringAlignment.Center})
            Case MouseState.Down
                g.TranslateTransform(1, 1)
                g.DrawString(Text, btnfont, New SolidBrush(Color.FromArgb(245, 245, 245)), rect, New StringFormat() With {.Alignment = StringAlignment.Center, .LineAlignment = StringAlignment.Center})
            Case MouseState.Over
                Dim rectovergb As New LinearGradientBrush(rect, Color.FromArgb(41, 41, 41), Color.FromArgb(61, 61, 61), LinearGradientMode.Vertical)
                g.FillRectangle(rectovergb, rect)
                g.DrawRectangle(New Pen(Brushes.Black), rect)
                g.DrawString(Text, btnfont, New SolidBrush(Color.FromArgb(245, 245, 245)), rect, New StringFormat() With {.Alignment = StringAlignment.Center, .LineAlignment = StringAlignment.Center})
        End Select

        If ShowNotification Then
            Select Case NotifyValue
                Case 0 To 9
                    g.FillEllipse(New SolidBrush(Color.FromArgb(81, 81, 81)), New Rectangle(Width - 25, 0, 20, 20))
                    g.DrawEllipse(New Pen(Brushes.Black), New Rectangle(Width - 25, 0, 20, 20))
                    g.DrawString(NotifyValue.ToString(), btnfont, New SolidBrush(Color.FromArgb(245, 245, 245)), New Rectangle(Width - 25, 0, 20, 20), New StringFormat() With {.Alignment = StringAlignment.Center, .LineAlignment = StringAlignment.Center})
                Case 10 To 99
                    g.FillEllipse(New SolidBrush(Color.FromArgb(81, 81, 81)), New Rectangle(Width - 45, 0, 40, 20))
                    g.DrawEllipse(New Pen(Brushes.Black), New Rectangle(Width - 45, 0, 40, 20))
                    g.DrawString(NotifyValue.ToString(), btnfont, New SolidBrush(Color.FromArgb(245, 245, 245)), New Rectangle(Width - 45, 0, 40, 20), New StringFormat() With {.Alignment = StringAlignment.Center, .LineAlignment = StringAlignment.Center})
                Case 100 To 999
                    g.FillEllipse(New SolidBrush(Color.FromArgb(81, 81, 81)), New Rectangle(Width - 65, 0, 60, 20))
                    g.DrawEllipse(New Pen(Brushes.Black), New Rectangle(Width - 65, 0, 60, 20))
                    g.DrawString(NotifyValue.ToString(), btnfont, New SolidBrush(Color.FromArgb(245, 245, 245)), New Rectangle(Width - 65, 0, 60, 20), New StringFormat() With {.Alignment = StringAlignment.Center, .LineAlignment = StringAlignment.Center})
            End Select
        End If

        e.Graphics.DrawImage(b, New Point(0, 0))
        g.Dispose() : b.Dispose()
    End Sub
End Class
