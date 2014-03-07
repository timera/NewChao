Public Class DisLabel
    Inherits System.Windows.Forms.Label

    Private _ForeColorBackup As Color = Color.Black
    Private _BackColorBackup As Color = SystemColors.Control
    Private _SettingColors As Boolean = False

    Private _BackColorDisabled As Color = SystemColors.Control
    Private _ForeColorDisabled As Color = Color.White 'SystemColors.WindowText

    Private Const WM_ENABLE As Integer = &HA

    Public Sub New()
        MyBase.New()
    End Sub

    Private Sub DisLabel_VisibleChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.VisibleChanged
        SetColors() ' Change to the Enabled/Disabled colors specified by the user
    End Sub

    Protected Overrides Sub OnForeColorChanged(ByVal e As System.EventArgs)
        MyBase.OnForeColorChanged(e)

        ' If the color is being set from OUTSIDE our control,
        ' then save the current ForeColor and set the specified color
        If Not _SettingColors Then
            _ForeColorBackup = Me.ForeColor
            SetColors()
        End If
    End Sub

    Protected Overrides Sub OnBackColorChanged(ByVal e As System.EventArgs)
        MyBase.OnBackColorChanged(e)

        ' If the color is being set from OUTSIDE our control,
        ' then save the current BackColor and set the specified color
        If Not _SettingColors Then
            _BackColorBackup = Me.BackColor
            SetColors()
        End If
    End Sub

    Private Sub SetColors()
        ' Don't change colors until the original ones have been saved,
        ' since we would lose what the original Enabled colors are supposed to be
            _SettingColors = True
            If Me.Enabled Then
                Me.ForeColor = Me._ForeColorBackup
                Me.BackColor = Me._BackColorBackup
            Else
                Me.ForeColor = Me.ForeColorDisabled
                Me.BackColor = Me.BackColorDisabled
            End If
            _SettingColors = False
    End Sub

    Protected Overrides Sub OnEnabledChanged(ByVal e As System.EventArgs)
        MyBase.OnEnabledChanged(e)

        SetColors() ' change colors whenever the Enabled() state changes
    End Sub

    Public Property BackColorDisabled() As System.Drawing.Color
        Get
            Return _BackColorDisabled
        End Get
        Set(ByVal Value As System.Drawing.Color)
            If Not Value.Equals(Color.Empty) Then
                _BackColorDisabled = Value
            End If
            SetColors()
        End Set
    End Property

    Public Property ForeColorDisabled() As System.Drawing.Color
        Get
            Return _ForeColorDisabled
        End Get
        Set(ByVal Value As System.Drawing.Color)
            If Not Value.Equals(Color.Empty) Then
                _ForeColorDisabled = Value
            End If
            SetColors()
        End Set
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

Public Class DisButton
    Inherits System.Windows.Forms.Button

    Private _ForeColorBackup As Color = SystemColors.WindowText
    Private _BackColorBackup As Color = System.Drawing.Color.Transparent
    Private _ColorsSaved As Boolean = False
    Private _SettingColors As Boolean = False

    Private _BackColorDisabled As Color = Color.DimGray 'SystemColors.Control
    Private _ForeColorDisabled As Color = Color.White 'SystemColors.WindowText

    Private Const WM_ENABLE As Integer = &HA

    Public Sub New()
        MyBase.New()
    End Sub

    Private Sub DisButton_VisibleChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.VisibleChanged
        If Not Me._ColorsSaved AndAlso Me.Visible Then
            ' Save the ForeColor/BackColor so we can switch back to them later
            _ForeColorBackup = Me.ForeColor
            _BackColorBackup = Me.BackColor
            _ColorsSaved = True

            If Not Me.Enabled Then ' If the window starts out in a Disabled state...
                ' Force the TextBox to initialize properly in an Enabled state,
                ' then switch it back to a Disabled state
                Me.Enabled = True
                Me.Enabled = False
            End If

            SetColors() ' Change to the Enabled/Disabled colors specified by the user
        End If
    End Sub

    Protected Overrides Sub OnForeColorChanged(ByVal e As System.EventArgs)
        MyBase.OnForeColorChanged(e)

        ' If the color is being set from OUTSIDE our control,
        ' then save the current ForeColor and set the specified color
        If Not _SettingColors Then
            _ForeColorBackup = Me.ForeColor
            SetColors()
        End If
    End Sub

    Protected Overrides Sub OnBackColorChanged(ByVal e As System.EventArgs)
        MyBase.OnBackColorChanged(e)

        ' If the color is being set from OUTSIDE our control,
        ' then save the current BackColor and set the specified color
        If Not _SettingColors Then
            _BackColorBackup = Me.BackColor
            SetColors()
        End If
    End Sub

    Private Sub SetColors()
        ' Don't change colors until the original ones have been saved,
        ' since we would lose what the original Enabled colors are supposed to be
        If _ColorsSaved Then
            _SettingColors = True
            If Me.Enabled Then
                Me.ForeColor = Me._ForeColorBackup
                Me.BackColor = Me._BackColorBackup
            Else
                Me.ForeColor = Me.ForeColorDisabled
                Me.BackColor = Me.BackColorDisabled
            End If
            _SettingColors = False
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

    Private _ForeColorBackup As Color
    Private _BackColorBackup As Color
    Private _ColorsSaved As Boolean = False
    Private _SettingColors As Boolean = False

    Private _BackColorDisabled As Color = Color.DarkGray 'SystemColors.Control
    Private _ForeColorDisabled As Color = Color.White 'SystemColors.WindowText

    Private Const WM_ENABLE As Integer = &HA

    Public Sub New()
        MyBase.New()
    End Sub

    Private Sub DisTextBox_VisibleChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.VisibleChanged
        If Not Me._ColorsSaved AndAlso Me.Visible Then
            ' Save the ForeColor/BackColor so we can switch back to them later
            _ForeColorBackup = Me.ForeColor
            _BackColorBackup = Me.BackColor
            _ColorsSaved = True

            If Not Me.Enabled Then ' If the window starts out in a Disabled state...
                ' Force the TextBox to initialize properly in an Enabled state,
                ' then switch it back to a Disabled state
                Me.Enabled = True
                Me.Enabled = False
            End If

            SetColors() ' Change to the Enabled/Disabled colors specified by the user
        End If
    End Sub

    Protected Overrides Sub OnForeColorChanged(ByVal e As System.EventArgs)
        MyBase.OnForeColorChanged(e)

        ' If the color is being set from OUTSIDE our control,
        ' then save the current ForeColor and set the specified color
        If Not _SettingColors Then
            _ForeColorBackup = Me.ForeColor
            SetColors()
        End If
    End Sub

    Protected Overrides Sub OnBackColorChanged(ByVal e As System.EventArgs)
        MyBase.OnBackColorChanged(e)

        ' If the color is being set from OUTSIDE our control,
        ' then save the current BackColor and set the specified color
        If Not _SettingColors Then
            _BackColorBackup = Me.BackColor
            SetColors()
        End If
    End Sub

    Private Sub SetColors()
        ' Don't change colors until the original ones have been saved,
        ' since we would lose what the original Enabled colors are supposed to be
        If _ColorsSaved Then
            _SettingColors = True
            If Me.Enabled Then
                Me.ForeColor = Me._ForeColorBackup
                Me.BackColor = Me._BackColorBackup
            Else
                Me.ForeColor = Me.ForeColorDisabled
                Me.BackColor = Me.BackColorDisabled
            End If
            _SettingColors = False
        End If
    End Sub

    Protected Overrides Sub OnEnabledChanged(ByVal e As System.EventArgs)
        MyBase.OnEnabledChanged(e)

        SetColors() ' change colors whenever the Enabled() state changes
    End Sub

    Public Property BackColorDisabled() As System.Drawing.Color
        Get
            Return _BackColorDisabled
        End Get
        Set(ByVal Value As System.Drawing.Color)
            If Not Value.Equals(Color.Empty) Then
                _BackColorDisabled = Value
            End If
            SetColors()
        End Set
    End Property

    Public Property ForeColorDisabled() As System.Drawing.Color
        Get
            Return _ForeColorDisabled
        End Get
        Set(ByVal Value As System.Drawing.Color)
            If Not Value.Equals(Color.Empty) Then
                _ForeColorDisabled = Value
            End If
            SetColors()
        End Set
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
