Public Class DisButton
    Inherits System.Windows.Forms.Button

    Private _ForeColorBackup As Color = SystemColors.ControlText
    Private _BackColorBackup As Color = Color.Transparent

    Private _BackColorDisabled As Color = Color.DimGray
    Private _ForeColorDisabled As Color = Color.White

    Private Const WM_ENABLE As Integer = &HA

    Private Sub DisButton_VisibleChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.VisibleChanged

        SetColors() ' Change to the Enabled/Disabled colors specified by the user
    End Sub

    Private Sub SetColors()
        If Me.Enabled Then
            Me.ForeColor = Me._ForeColorBackup
            Me.BackColor = Me._BackColorBackup
        Else
            Me.ForeColor = Me.ForeColorDisabled
            Me.BackColor = Me.BackColorDisabled
        End If
    End Sub

    Protected Overrides Sub OnEnabledChanged(ByVal e As System.EventArgs)
        MyBase.OnEnabledChanged(e)

        SetColors() ' change colors whenever the Enabled() state changes
    End Sub

    Public ReadOnly Property BackColorDisabled() As System.Drawing.Color
        Get
            Return _BackColorDisabled
        End Get
    End Property

    Public ReadOnly Property ForeColorDisabled() As System.Drawing.Color
        Get
            Return _ForeColorDisabled
        End Get
    End Property

    Protected Overrides ReadOnly Property CreateParams As System.Windows.Forms.CreateParams
        Get
            Dim cp As System.Windows.Forms.CreateParams
            If Not Me.Enabled Then ' If the window starts out in a disabled state...
                ' Prevent window being initialized in a disabled state:
                Me.Enabled = True ' temporary ENABLED state
                cp = MyBase.CreateParams ' create window in ENABLED state
                Me.Enabled = False ' toggle it back to DISABLED state 
            Else
                cp = MyBase.CreateParams
            End If
            Return cp
        End Get
    End Property

    Protected Overrides Sub WndProc(ByRef m As System.Windows.Forms.Message)
        Select Case m.Msg
            Case WM_ENABLE
                ' Prevent the message from reaching the control,
                ' so the colors don't get changed by the default procedure.
                Exit Sub ' <-- suppress WM_ENABLE message

        End Select

        MyBase.WndProc(m)
    End Sub

End Class

Public Class DisTextBox
    Inherits System.Windows.Forms.TextBox

    Private _ForeColorBackup As Color = SystemColors.WindowText
    Private _BackColorBackup As Color = SystemColors.Control

    Private _BackColorDisabled As Color = Color.DarkGray 'SystemColors.Control
    Private _ForeColorDisabled As Color = Color.White 'SystemColors.WindowText

    Private Const WM_ENABLE As Integer = &HA

    Public Sub New()
        MyBase.New()
    End Sub

    Private Sub DisTextBox_VisibleChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.VisibleChanged
        SetColors() ' Change to the Enabled/Disabled colors specified by the user
    End Sub

    Private Sub SetColors()
        If Me.Enabled Then
            Me.ForeColor = Me._ForeColorBackup
            Me.BackColor = Me._BackColorBackup
        Else
            Me.ForeColor = Me.ForeColorDisabled
            Me.BackColor = Me.BackColorDisabled
        End If
    End Sub

    Protected Overrides Sub OnEnabledChanged(ByVal e As System.EventArgs)
        MyBase.OnEnabledChanged(e)

        SetColors() ' change colors whenever the Enabled() state changes
    End Sub

    Public ReadOnly Property BackColorDisabled() As System.Drawing.Color
        Get
            Return _BackColorDisabled
        End Get
    End Property

    Public ReadOnly Property ForeColorDisabled() As System.Drawing.Color
        Get
            Return _ForeColorDisabled
        End Get
    End Property

    Protected Overrides ReadOnly Property CreateParams As System.Windows.Forms.CreateParams
        Get
            Dim cp As System.Windows.Forms.CreateParams
            If Not Me.Enabled Then ' If the window starts out in a disabled state...
                ' Prevent window being initialized in a disabled state:
                Me.Enabled = True ' temporary ENABLED state
                cp = MyBase.CreateParams ' create window in ENABLED state
                Me.Enabled = False ' toggle it back to DISABLED state 
            Else
                cp = MyBase.CreateParams
            End If
            Return cp
        End Get
    End Property

    Protected Overrides Sub WndProc(ByRef m As System.Windows.Forms.Message)
        Select Case m.Msg
            Case WM_ENABLE
                ' Prevent the message from reaching the control,
                ' so the colors don't get changed by the default procedure.
                Exit Sub ' <-- suppress WM_ENABLE message

        End Select

        MyBase.WndProc(m)
    End Sub

End Class

Class ChaoTextBox
    Inherits DisTextBox

    Public Sub New()
        MyBase.New()
    End Sub

    Public Sub ChaoTextBox_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        Dim nonNumberEntered = False

        ' Determine whether the keystroke is a number from the top of the keyboard. 
        If e.KeyCode < Keys.D0 OrElse e.KeyCode > Keys.D9 And Not e.KeyCode = Keys.Decimal And Not e.KeyCode = Keys.OemPeriod Then
            ' Determine whether the keystroke is a number from the keypad. 
            If e.KeyCode < Keys.NumPad0 OrElse e.KeyCode > Keys.NumPad9 Then
                ' Determine whether the keystroke is a backspace. 
                If e.KeyCode <> Keys.Back And e.KeyCode <> Keys.Shift And e.KeyCode <> Keys.Alt And e.KeyCode <> Keys.Control And e.KeyCode <> Keys.LWin And e.KeyCode <> Keys.RWin Then
                    ' A non-numerical keystroke was pressed.  
                    ' Set the flag to true and evaluate in KeyPress event.
                    MsgBox("只能輸入數字!")
                    Me.Text = ""
                End If
            End If
        End If
        'If shift key was pressed, it's not a number. 
        If Control.ModifierKeys = Keys.Shift Then
            MsgBox("只能輸入數字!")
            Me.Text = ""
        End If
    End Sub
End Class

Public Class ColorLabel
    Inherits Control

    Public Sub New()
        MyBase.New()
    End Sub

    Public Sub ClearXYZ()
        _X = 0
        _Y = 0
        _Z = 0
    End Sub

    Private _X As Decimal = 0

    Property X() As Decimal
        Get
            Return _X
        End Get
        Set(ByVal value As Decimal)
            _X = value
            Me.Invalidate()
        End Set
    End Property

    Private _Y As Decimal = 0

    Property Y() As Decimal
        Get
            Return _Y
        End Get
        Set(ByVal value As Decimal)
            _Y = value
            Me.Invalidate()
        End Set
    End Property

    Private _Z As Decimal = 0

    Property Z() As Decimal
        Get
            Return _Z
        End Get
        Set(ByVal value As Decimal)
            _Z = value
            Me.Invalidate()
        End Set
    End Property

    Private _P As String = "P"

    Property P() As String
        Get
            Return _P
        End Get
        Set(ByVal value As String)
            _P = value
            Me.Invalidate()
        End Set
    End Property

    Protected Overrides Sub OnPaint(ByVal e As PaintEventArgs)
        MyBase.OnPaint(e)
        Me.Size = New Size(120, 100)
        Me.Font = New System.Drawing.Font("Microsoft Sans Serif", 15.0!)
        If _X = 0 And _Y = 0 And _Z = 0 Then
            e.Graphics.Clear(SystemColors.Control)
            Return
        End If

        Dim xText As String = String.Format("X: {0:N1}", _X)
        Dim yText As String = String.Format("Y: {0:N1}", _Y)
        Dim zText As String = String.Format("Z: {0:N1}", _Z)
        Dim pHeight As Integer = TextRenderer.MeasureText(e.Graphics, _P, Me.Font).Height
        Dim pWidth As Integer = TextRenderer.MeasureText(e.Graphics, _P, Me.Font).Width


        TextRenderer.DrawText(e.Graphics, _P, Me.Font, New Point(0, 0), Color.DimGray)
        TextRenderer.DrawText(e.Graphics, xText, Me.Font, New Point(pWidth + 2, 0), Color.Blue)
        TextRenderer.DrawText(e.Graphics, yText, Me.Font, New Point(pWidth + 2, 2 + pHeight), Color.Red)
        TextRenderer.DrawText(e.Graphics, zText, Me.Font, New Point(pWidth + 2, (2 + pHeight) * 2), Color.DarkGreen)
        

    End Sub

End Class