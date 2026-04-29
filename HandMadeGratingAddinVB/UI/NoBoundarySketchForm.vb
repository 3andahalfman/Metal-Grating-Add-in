'////////////////////////////////////////////////////////////////////
' Metal Bar Grating Addin - VB.NET
'
' NoBoundarySketchForm: Non-modal instruction dialog shown when no
' boundary sketch is available.  Unlike a MessageBox this form
' allows the user to interact with Inventor (select / edit a sketch)
' while the instructions remain visible on screen.
'
' The caller uses ShowNonModal() which displays the form modeless
' and pumps messages until the user clicks "Try Again" or "Cancel".
'////////////////////////////////////////////////////////////////////

Imports System.Drawing
Imports System.Windows.Forms

''' <summary>
''' Non-modal instruction form shown when the Create Grating command
''' cannot find a named or selected boundary sketch.  The user can
''' interact with Inventor behind this window, then click
''' "Try Again" to re-check.
''' </summary>
Public Class NoBoundarySketchForm
    Inherits Form

    ''' <summary>Action chosen by the user.</summary>
    Public Enum UserAction
        ''' <summary>User wants to re-check for a boundary sketch.</summary>
        TryAgain = 0
        ''' <summary>User cancelled the workflow.</summary>
        Cancel = 1
    End Enum

    Private _action As UserAction = UserAction.Cancel
    Private _dismissed As Boolean = False

    ' --- Public result ---

    ''' <summary>
    ''' The action the user chose when the form was dismissed.
    ''' Only meaningful after ShowNonModal returns.
    ''' </summary>
    Public ReadOnly Property SelectedAction As UserAction
        Get
            Return _action
        End Get
    End Property

    ''' <summary>Creates the form with all controls.</summary>
    Public Sub New()
        InitializeControls()
    End Sub

    ' ==================================================================
    '  Non-modal wait loop
    ' ==================================================================

    ''' <summary>
    ''' Shows the form modeless (non-blocking) so the user can interact
    ''' with Inventor behind it, then waits in a DoEvents loop until
    ''' the user clicks "Try Again" or "Cancel".
    ''' </summary>
    Public Function ShowNonModal() As UserAction
        _dismissed = False
        _action = UserAction.Cancel
        Me.TopMost = True
        Me.Show()

        While Not _dismissed
            Application.DoEvents()
            Threading.Thread.Sleep(50)
        End While

        Return _action
    End Function

    ' ==================================================================
    '  Prevent accidental closure via X button without setting action
    ' ==================================================================

    Protected Overrides Sub OnFormClosing(e As FormClosingEventArgs)
        If Not _dismissed Then
            ' User clicked the X button — treat as Cancel
            _action = UserAction.Cancel
            _dismissed = True
        End If
        MyBase.OnFormClosing(e)
    End Sub

    ' ==================================================================
    '  Event handlers
    ' ==================================================================

    Private Sub OnTryAgainClick(sender As Object, e As EventArgs)
        _action = UserAction.TryAgain
        _dismissed = True
        Me.Close()
    End Sub

    Private Sub OnCancelClick(sender As Object, e As EventArgs)
        _action = UserAction.Cancel
        _dismissed = True
        Me.Close()
    End Sub

    ' ==================================================================
    '  Control layout
    ' ==================================================================

    Private Sub InitializeControls()
        Me.Text = "Metal Bar Grating — No Boundary Sketch Available"
        Me.FormBorderStyle = FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.StartPosition = FormStartPosition.CenterScreen
        Me.ShowInTaskbar = True
        Me.TopMost = True

        Dim pad As Integer = 18
        Dim yPos As Integer = pad
        Dim contentWidth As Integer = 430

        ' --- Info icon + heading ---
        Dim lblHeading As New Label()
        lblHeading.Text = "No grating boundary sketch is available."
        lblHeading.Font = New Font(Me.Font.FontFamily, Me.Font.Size, FontStyle.Bold)
        lblHeading.Location = New Point(pad, yPos)
        lblHeading.AutoSize = True
        Me.Controls.Add(lblHeading)
        yPos += 28

        ' --- Instructions ---
        Dim instructions As String =
            "Select or create a boundary sketch in Inventor while " &
            "this window is open, then click Try Again." & vbCrLf & vbCrLf &
            "Option A — create a named boundary sketch:" & vbCrLf &
            "   Draw a closed sketch and name it """ &
            BoundarySourceService.PrimaryName & """." & vbCrLf & vbCrLf &
            "Option B — use a manually selected sketch:" & vbCrLf &
            "   1. Double-click a sketch to enter Edit mode, or" & vbCrLf &
            "   2. Single-click a sketch in the browser, or" & vbCrLf &
            "   3. Select sketch curves in the graphics area."

        Dim lblInstructions As New Label()
        lblInstructions.Text = instructions
        lblInstructions.Location = New Point(pad, yPos)
        lblInstructions.Size = New Size(contentWidth, 160)
        Me.Controls.Add(lblInstructions)
        yPos += 168

        ' --- Buttons ---
        Dim btnTryAgain As New Button()
        btnTryAgain.Text = "Try Again"
        btnTryAgain.Size = New Size(100, 30)
        btnTryAgain.Location = New Point(contentWidth + pad - 210, yPos)
        AddHandler btnTryAgain.Click, AddressOf OnTryAgainClick
        Me.Controls.Add(btnTryAgain)

        Dim btnCancel As New Button()
        btnCancel.Text = "Cancel"
        btnCancel.Size = New Size(100, 30)
        btnCancel.Location = New Point(contentWidth + pad - 104, yPos)
        AddHandler btnCancel.Click, AddressOf OnCancelClick
        Me.Controls.Add(btnCancel)

        Me.AcceptButton = btnTryAgain
        Me.CancelButton = btnCancel

        Me.ClientSize = New Size(contentWidth + pad * 2, yPos + 44)
    End Sub

End Class
