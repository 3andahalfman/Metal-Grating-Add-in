'////////////////////////////////////////////////////////////////////
' Metal Bar Grating Addin - VB.NET
'
' BoundarySourceDialog: Modal dialog shown at the start of the
' Create Grating command. Lets the user choose how the grating
' perimeter boundary is sourced.
'
' Phase 11: Only "Use currently selected sketch" is active.
' Phase 12: "Use existing named sketch" enabled when BoundarySourceService
'           finds GRATING_BOUNDARY (or project alias) in the active
'           Part document. Preferred over selected sketch when available.
'////////////////////////////////////////////////////////////////////

Imports System.Diagnostics
Imports System.Drawing
Imports System.Windows.Forms

''' <summary>
''' Modal dialog that lets the user choose the grating boundary source
''' before perimeter extraction begins.
''' </summary>
Public Class BoundarySourceDialog
    Inherits Form

    ' --- Controls ---
    Private _txtProjectName As TextBox
    Private _rbSelectedSketch As RadioButton
    Private _rbNamedSketch As RadioButton
    Private _btnContinue As Button
    Private _btnCancel As Button

    ''' <summary>
    ''' Phase 12: lookup result from BoundarySourceService, used to
    ''' enable/disable and label the named sketch radio button.
    ''' </summary>
    Private ReadOnly _namedSketchLookup As NamedSketchLookupResult
    Private ReadOnly _app As Inventor.Application

    ''' <summary>
    ''' The user-entered project name. Only meaningful when DialogResult is OK.
    ''' </summary>
    Public Property ProjectName As String

    ''' <summary>
    ''' The boundary source type chosen by the user.
    ''' Only meaningful when DialogResult is OK.
    ''' </summary>
    Public Property SelectedSourceType As BoundarySourceType

    ''' <summary>
    ''' Creates the dialog with an optional default project name and an
    ''' optional Phase 12 named sketch lookup result.
    ''' When <paramref name="namedSketchLookup"/> reports Found=True the
    ''' named sketch option is enabled and pre-selected.
    ''' </summary>
    Public Sub New(Optional defaultProjectName As String = "Grating",
                   Optional namedSketchLookup As NamedSketchLookupResult = Nothing,
                   Optional app As Inventor.Application = Nothing)
        ProjectName = If(defaultProjectName, "Grating")
        SelectedSourceType = BoundarySourceType.SelectedSketch
        _namedSketchLookup = namedSketchLookup
        _app = app
        InitializeControls()
        ConfigureNamedSketchOption()
    End Sub

    ' ==================================================================
    '  Control layout (code-only, no designer)
    ' ==================================================================

    Private Sub InitializeControls()
        Me.Text = "Metal Bar Grating — Boundary Source"
        Me.FormBorderStyle = FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.StartPosition = FormStartPosition.CenterParent
        Me.ShowInTaskbar = False

        Dim labelX As Integer = 18
        Dim controlX As Integer = 140
        Dim controlWidth As Integer = 270
        Dim rowHeight As Integer = 28
        Dim yPos As Integer = 18

        ' ==============================================================
        '  Project name
        ' ==============================================================

        Dim lblProject As New Label()
        lblProject.Text = "Project Name:"
        lblProject.Location = New Point(labelX, yPos + 3)
        lblProject.AutoSize = True
        Me.Controls.Add(lblProject)

        _txtProjectName = New TextBox()
        _txtProjectName.Text = ProjectName
        _txtProjectName.Location = New Point(controlX, yPos)
        _txtProjectName.Size = New Size(controlWidth, 21)
        Me.Controls.Add(_txtProjectName)
        yPos += rowHeight + 10

        ' ==============================================================
        '  Separator
        ' ==============================================================

        Dim separator As New Label()
        separator.BorderStyle = BorderStyle.Fixed3D
        separator.Location = New Point(labelX, yPos)
        separator.Size = New Size(controlX + controlWidth - labelX, 2)
        Me.Controls.Add(separator)
        yPos += 12

        ' ==============================================================
        '  Prompt
        ' ==============================================================

        Dim lblPrompt As New Label()
        lblPrompt.Text = "How would you like to define the grating perimeter?"
        lblPrompt.Location = New Point(labelX, yPos)
        lblPrompt.AutoSize = True
        lblPrompt.Font = New Font(Me.Font, FontStyle.Bold)
        Me.Controls.Add(lblPrompt)
        yPos += 24

        ' ==============================================================
        '  Radio buttons — source options
        ' ==============================================================

        ' --- Selected sketch (active, Phase 11) ---
        _rbSelectedSketch = New RadioButton()
        _rbSelectedSketch.Text = "Use currently selected sketch"
        _rbSelectedSketch.Location = New Point(labelX + 8, yPos)
        _rbSelectedSketch.AutoSize = True
        _rbSelectedSketch.Checked = True
        _rbSelectedSketch.Enabled = True
        Me.Controls.Add(_rbSelectedSketch)
        yPos += rowHeight

        ' --- Named sketch (Phase 12 — enabled/labeled by ConfigureNamedSketchOption) ---
        _rbNamedSketch = New RadioButton()
        _rbNamedSketch.Text = "Use existing named sketch (" &
                              BoundarySourceService.PrimaryName & ")"
        _rbNamedSketch.Location = New Point(labelX + 8, yPos)
        _rbNamedSketch.AutoSize = True
        _rbNamedSketch.Enabled = False  ' overridden by ConfigureNamedSketchOption
        Me.Controls.Add(_rbNamedSketch)
        yPos += rowHeight + 12

        ' ==============================================================
        '  Hint
        ' ==============================================================

        Dim lblHint As New Label()
        lblHint.Text = "Tip: Edit or select a closed sketch before clicking Create Grating."
        lblHint.Location = New Point(labelX, yPos)
        lblHint.AutoSize = True
        lblHint.ForeColor = SystemColors.GrayText
        Me.Controls.Add(lblHint)
        yPos += 24

        ' ==============================================================
        '  Continue / Cancel
        ' ==============================================================

        Dim formWidth As Integer = controlX + controlWidth + 24

        _btnContinue = New Button()
        _btnContinue.Text = "Continue"
        _btnContinue.Size = New Size(90, 28)
        _btnContinue.Location = New Point(formWidth - 200, yPos)
        AddHandler _btnContinue.Click, AddressOf OnContinueClick
        Me.Controls.Add(_btnContinue)

        _btnCancel = New Button()
        _btnCancel.Text = "Cancel"
        _btnCancel.Size = New Size(90, 28)
        _btnCancel.Location = New Point(formWidth - 104, yPos)
        _btnCancel.DialogResult = DialogResult.Cancel
        Me.Controls.Add(_btnCancel)

        Me.AcceptButton = _btnContinue
        Me.CancelButton = _btnCancel

        Me.ClientSize = New Size(formWidth, yPos + 42)
    End Sub

    ' ==================================================================
    '  Phase 12: Named sketch option configuration
    ' ==================================================================

    ''' <summary>
    ''' Configures the named sketch radio button based on the lookup result
    ''' passed to the constructor.  Called once after InitializeControls.
    '''
    ''' Found=True  → enabled, shows the exact sketch name, pre-selected.
    ''' Found=False → disabled, shows "(not found in document)".
    ''' </summary>
    Private Sub ConfigureNamedSketchOption()
        If _namedSketchLookup IsNot Nothing AndAlso _namedSketchLookup.Found Then
            _rbNamedSketch.Text = "Use existing named sketch: " &
                                  _namedSketchLookup.SketchName
            _rbNamedSketch.Enabled = True
            ' Prefer named sketch when available
            _rbSelectedSketch.Checked = False
            _rbNamedSketch.Checked = True
            Trace.TraceInformation(
                ": HMG BoundarySourceDialog: Named sketch option enabled — '" &
                _namedSketchLookup.SketchName & "'.")
        Else
            _rbNamedSketch.Text = "Use existing named sketch (" &
                                  BoundarySourceService.PrimaryName &
                                  ")    — not found in document"
            _rbNamedSketch.Enabled = False
            Trace.TraceInformation(
                ": HMG BoundarySourceDialog: Named sketch option disabled — not found.")
        End If
    End Sub

    ' ==================================================================
    '  Event handlers
    ' ==================================================================

    Private Sub OnContinueClick(sender As Object, e As EventArgs)
        ' Validate project name
        Dim name As String = _txtProjectName.Text.Trim()
        If String.IsNullOrEmpty(name) Then
            MessageBox.Show("Please enter a project name.",
                            "Metal Bar Grating",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Warning)
            _txtProjectName.Focus()
            Return
        End If

        ProjectName = name

        ' Determine selected source type
        If _rbSelectedSketch.Checked Then
            SelectedSourceType = BoundarySourceType.SelectedSketch
        ElseIf _rbNamedSketch.Checked Then
            SelectedSourceType = BoundarySourceType.NamedSketch
        Else
            SelectedSourceType = BoundarySourceType.SelectedSketch
        End If

        Trace.TraceInformation(": HMG: BoundarySourceDialog — project=""" &
            ProjectName & """, source=" & SelectedSourceType.ToString())

        If Not ValidateSelectedSource() Then
            Return
        End If

        Me.DialogResult = DialogResult.OK
        Me.Close()
    End Sub

    Private Function ValidateSelectedSource() As Boolean
        If SelectedSourceType = BoundarySourceType.NamedSketch Then
            If _namedSketchLookup IsNot Nothing AndAlso _namedSketchLookup.Found Then
                Return True
            End If
            MessageBox.Show(
                "The named boundary sketch """ &
                BoundarySourceService.PrimaryName &
                """ was not found in this Part document.",
                "Metal Bar Grating",
                MessageBoxButtons.OK,
                MessageBoxIcon.Exclamation)
            Return False
        End If

        If _app Is Nothing Then
            Return True
        End If

        Dim check As SelectionResult =
            New PerimeterSelectionService(_app).ValidateSketchResolvable()
        If check.Success Then Return True

        MessageBox.Show(
            check.ErrorMessage,
            "Metal Bar Grating",
            MessageBoxButtons.OK,
            MessageBoxIcon.Exclamation)
        Return False
    End Function

End Class
