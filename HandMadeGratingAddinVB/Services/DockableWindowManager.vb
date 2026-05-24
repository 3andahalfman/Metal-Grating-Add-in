'////////////////////////////////////////////////////////////////////
' Metal Bar Grating Addin - VB.NET
'
' DockableWindowManager: Creates and manages an Inventor DockableWindow
' that hosts the HmgPanelControl.  This gives the add-in a docked
' panel experience identical to iLogic or Design Accelerator —
' the panel docks to the right side of Inventor (like the Fusion 360
' Interoperability panel) and persists its position across sessions.
'
' Lifecycle
' ---------
'  Initialize()  — called during add-in Activate; creates or retrieves
'                   the DockableWindow and hosts the .NET UserControl.
'  HidePanel()   — hides the dockable window.
'  Cleanup()     — called during add-in Deactivate; disposes the
'                   UserControl and releases the COM reference.
'
' DockableWindow persistence
' --------------------------
'  Inventor persists DockableWindow positions and visibility in the
'  workspace layout.  If the user moves or resizes the panel, those
'  settings survive Inventor restarts.  The InternalName must be
'  stable across versions to preserve the layout.
'
' Note on AddChild
' ----------------
'  AddChild() takes an HWND (IntPtr).  Accessing UserControl.Handle
'  forces the native window to be created.  With Option Strict Off
'  the COM interop layer handles IntPtr → Integer marshalling.
'////////////////////////////////////////////////////////////////////

Imports System.Diagnostics
Imports Inventor

''' <summary>
''' Manages an Inventor DockableWindow that hosts the HMG panel control.
''' </summary>
Public Class DockableWindowManager

    ''' <summary>
    ''' Stable internal name used by Inventor to persist window position.
    ''' Do not change between versions or the user's layout resets.
    ''' Changed from "HMG_DockPanel" to "HMG_DockPanel_R" to create
    ''' a fresh window on the right side, independent of the Model Browser.
    ''' </summary>
    Private Const InternalName As String = "HMG_DockPanel_R"

    ''' <summary>Previous internal name — removed on startup.</summary>
    Private Const OldInternalName As String = "HMG_DockPanel"

    ''' <summary>Caption shown in the dockable window title bar.</summary>
    Private Const Caption As String = "Metal Bar Grating"

    Private _app As Application
    Private _clientId As String
    Private _dockableWindow As DockableWindow
    Private _panel As HmgPanelControl
    Private _initialized As Boolean = False
    Private _materialized As Boolean = False

    ''' <summary>
    ''' Per-document summary cache keyed by Document.InternalName.
    ''' Allows the user to generate grating in multiple assemblies
    ''' and switch between them without losing any summary.
    ''' </summary>
    Private ReadOnly _summaryCache As New Dictionary(Of String, CachedSummary)(
        StringComparer.OrdinalIgnoreCase)

    ''' <summary>Holds the cached summary + project for one document.</summary>
    Private Class CachedSummary
        Public Property SummaryText As String
        Public Property Project As GratingProject
    End Class

    ''' <summary>
    ''' Held reference to prevent the event sink from being garbage-collected.
    ''' </summary>
    Private _uiEvents As UserInterfaceEvents

    ''' <summary>
    ''' Held reference for document activation events.
    ''' </summary>
    Private _appEvents As ApplicationEvents

    ''' <summary>
    ''' Held reference for dockable window show/hide events.  Used to
    ''' detect when the user closes the panel via the X button so that
    ''' auto-restore in environment/document handlers doesn't pop it
    ''' back open immediately.
    ''' </summary>
    Private _dwEvents As DockableWindowsEvents

    ''' <summary>
    ''' Set to True while our own code is about to hide the dockable
    ''' window so the OnHide handler does not treat that as a user
    ''' dismissal.
    ''' </summary>
    Private _suppressHideEvent As Boolean = False

    ''' <summary>
    ''' Documents (by InternalName) for which the user has explicitly
    ''' dismissed the panel.  Auto-restore is skipped while the doc
    ''' is in this set.  Entries are cleared when a new generation
    ''' produces a fresh summary for the doc.
    ''' </summary>
    Private ReadOnly _userDismissedDocs As New HashSet(Of String)(
        StringComparer.OrdinalIgnoreCase)

    ''' <summary>
    ''' Callback invoked whenever the answer to
    ''' <see cref="CanShowGenerationSummary"/> may have changed
    ''' (document switch, environment change, summary cached, user
    ''' dismissal).  The ribbon's "View Generation Summary" button
    ''' subscribes to this to toggle its enabled state.
    ''' </summary>
    Public Property SummaryAvailabilityChanged As Action

    ''' <summary>
    ''' True when the active document is an assembly **AND** was
    ''' produced by this add-in (identified by the "Metal Bar Grating"
    ''' iProperty set stamped during generation by
    ''' <see cref="GratingAssemblyGenerator"/>).  Other assemblies the
    ''' user happens to have open keep the button greyed out — the
    ''' marker is what distinguishes a grating assembly.
    ''' </summary>
    Public ReadOnly Property CanShowGenerationSummary As Boolean
        Get
            If Not _initialized OrElse _app Is Nothing Then Return False
            Try
                Dim activeDoc As Document = _app.ActiveDocument
                If activeDoc Is Nothing Then Return False
                If activeDoc.DocumentType <>
                   DocumentTypeEnum.kAssemblyDocumentObject Then Return False
                Return DocHasGratingMarker(activeDoc)
            Catch
                Return False
            End Try
        End Get
    End Property

    ''' <summary>
    ''' True when the active document is a Part document.  Drives the
    ''' Create Grating ribbon button's enabled state — that command
    ''' requires a part-level perimeter sketch as input, so it is
    ''' greyed out in assemblies / drawings / the Home page.
    ''' </summary>
    Public ReadOnly Property CanCreateGrating As Boolean
        Get
            If Not _initialized OrElse _app Is Nothing Then Return False
            Try
                Dim activeDoc As Document = _app.ActiveDocument
                If activeDoc Is Nothing Then Return False
                Return activeDoc.DocumentType =
                       DocumentTypeEnum.kPartDocumentObject
            Catch
                Return False
            End Try
        End Get
    End Property

    ''' <summary>
    ''' True if the given document carries the Metal Bar Grating
    ''' iProperty marker — i.e. it was generated by this add-in.
    ''' Looks for the "HMG_Type" property in the
    ''' "Metal Bar Grating" property set
    ''' (see <see cref="GratingAssemblyGenerator.HmgPropertySetName"/>
    ''' and <see cref="GratingAssemblyGenerator.HmgMarkerPropertyName"/>).
    ''' </summary>
    Private Function DocHasGratingMarker(doc As Document) As Boolean
        Try
            Dim propSet As PropertySet =
                doc.PropertySets.Item(GratingAssemblyGenerator.HmgPropertySetName)
            Dim marker As [Property] =
                propSet.Item(GratingAssemblyGenerator.HmgMarkerPropertyName)
            Return marker IsNot Nothing
        Catch
            Return False
        End Try
    End Function

    ''' <summary>
    ''' True iff the active document has a cached generation summary.
    ''' Used by the click handler to decide whether to restore the
    ''' panel or show a "no summary yet" message.  Independent of
    ''' user-dismissal state — a dismissed-but-cached doc still
    ''' returns True.
    ''' </summary>
    Public ReadOnly Property HasCachedSummaryForActiveDoc As Boolean
        Get
            If Not _initialized OrElse _app Is Nothing Then Return False
            Try
                Dim activeDoc As Document = _app.ActiveDocument
                If activeDoc Is Nothing Then Return False
                Return _summaryCache.ContainsKey(activeDoc.InternalName)
            Catch
                Return False
            End Try
        End Get
    End Property

    ''' <summary>
    ''' Re-shows the cached summary panel for the active document.
    ''' Lazily materializes the dockable window if it has not been
    ''' created yet, and clears any prior user-dismissal so the auto-
    ''' restore handlers will continue to work afterwards.  Safe no-op
    ''' when no summary is cached for the active document.
    ''' </summary>
    Public Sub RestoreCurrentDocSummary()
        Dim cached As CachedSummary = TryGetCachedSummary()
        If cached Is Nothing Then Return

        Try
            Dim docKey As String = _app.ActiveDocument.InternalName
            _userDismissedDocs.Remove(docKey)
        Catch
        End Try

        If Not _materialized Then MaterializeDockableWindow()
        If Not _materialized Then Return

        RestoreSummary(cached)
        NotifySummaryAvailabilityChanged()
    End Sub

    Private Sub NotifySummaryAvailabilityChanged()
        Try
            Dim cb As Action = SummaryAvailabilityChanged
            If cb IsNot Nothing Then cb.Invoke()
        Catch ex As Exception
            Trace.TraceWarning(
                ": HMG DockableWindowManager: SummaryAvailabilityChanged " &
                "callback threw — " & ex.Message)
        End Try
    End Sub

    ''' <summary>
    ''' True if the DockableWindow was created/retrieved successfully
    ''' and the panel control is hosted.  When False, callers should
    ''' fall back to standalone forms.
    ''' </summary>
    Public ReadOnly Property IsAvailable As Boolean
        Get
            ' Trigger lazy creation on first access if not yet materialized
            If Not _materialized AndAlso _initialized Then
                MaterializeDockableWindow()
            End If
            Return _materialized AndAlso
                   _dockableWindow IsNot Nothing AndAlso
                   _panel IsNot Nothing
        End Get
    End Property

    ' ==================================================================
    '  Lifecycle
    ' ==================================================================

    ''' <summary>
    ''' Stores the Inventor references for later use.  Call once during
    ''' add-in Activate.  The DockableWindow is NOT created here — it is
    ''' deferred to the first time the panel is actually needed so that
    ''' nothing flashes visible during Inventor startup.
    '''
    ''' Safe to call multiple times — subsequent calls are no-ops.
    ''' </summary>
    Public Sub Initialize(app As Application, clientId As String)
        If _initialized Then Return

        _app = app
        _clientId = clientId
        _initialized = True

        ' Hide any DockableWindow that Inventor auto-restored from a
        ' previous session to prevent a brief flash on the wrong side.
        Try
            Dim uiMgr As UserInterfaceManager = _app.UserInterfaceManager
            Dim existingWin As DockableWindow = uiMgr.DockableWindows.Item(InternalName)
            If existingWin.Visible Then
                existingWin.Visible = False
            End If
        Catch
        End Try

        ' Also hide the old-name window if Inventor restored it
        Try
            Dim uiMgr2 As UserInterfaceManager = _app.UserInterfaceManager
            Dim oldWin As DockableWindow = uiMgr2.DockableWindows.Item(OldInternalName)
            oldWin.Visible = False
        Catch
        End Try

        ' Subscribe to environment change events so the summary
        ' panel is only shown in the assembly environment.
        Try
            _uiEvents = _app.UserInterfaceEvents
            AddHandler _uiEvents.OnEnvironmentChange,
                AddressOf OnEnvironmentChange
        Catch
        End Try

        ' Subscribe to document activation events so the panel
        ' resets when the user switches to a different document.
        ' OnDeactivateDocument (kBefore) hides the panel immediately
        ' so the stale tab does not linger during the switch repaint.
        Try
            _appEvents = _app.ApplicationEvents
            AddHandler _appEvents.OnActivateDocument,
                AddressOf OnActivateDocument
            AddHandler _appEvents.OnDeactivateDocument,
                AddressOf OnDeactivateDocument
        Catch
        End Try

        ' Subscribe to dockable window hide events so we can record
        ' explicit user dismissals (X button) and stop auto-restoring
        ' the panel for that document.
        Try
            _dwEvents = _app.UserInterfaceManager.DockableWindowsEvents
            AddHandler _dwEvents.OnHide, AddressOf OnDockableWindowHide
        Catch
        End Try

        Trace.TraceInformation(
            ": HMG DockableWindowManager: Initialized (deferred). " &
            "DockableWindow will be created on first use.")
    End Sub

    ''' <summary>
    ''' Creates the DockableWindow and hosts the panel control.
    ''' Called lazily on first access — never during Inventor startup.
    ''' </summary>
    Private Sub MaterializeDockableWindow()
        If _materialized Then Return
        If _app Is Nothing OrElse _clientId Is Nothing Then Return

        Try
            ' Create the .NET UserControl that will be hosted
            _panel = New HmgPanelControl()

            ' Force native window creation so we have a valid HWND
            Dim hwnd As IntPtr = _panel.Handle

            Dim uiMgr As UserInterfaceManager = _app.UserInterfaceManager

            ' Remove the old dockable window (was tabbed with Model Browser
            ' on the left).  Hiding + deleting prevents it from lingering.
            Try
                Dim oldWin As DockableWindow = uiMgr.DockableWindows.Item(OldInternalName)
                oldWin.Visible = False
                oldWin.Delete()
                Trace.TraceInformation(
                    ": HMG DockableWindowManager: Removed old DockableWindow ('" &
                    OldInternalName & "').")
            Catch
                ' Old window doesn't exist — nothing to clean up
            End Try

            ' Retrieve existing DockableWindow (survives Inventor sessions)
            ' or create a new one
            Dim isNew As Boolean = False

            Try
                _dockableWindow = uiMgr.DockableWindows.Item(InternalName)
                Trace.TraceInformation(
                    ": HMG DockableWindowManager: Retrieved existing DockableWindow.")
            Catch
                _dockableWindow = uiMgr.DockableWindows.Add(
                    _clientId, InternalName, Caption)
                isNew = True
                Trace.TraceInformation(
                    ": HMG DockableWindowManager: Created new DockableWindow.")
            End Try

            ' Hide the window BEFORE hosting the child control.
            ' Inventor persists DockableWindow visibility across sessions;
            ' if the user previously left it open, the restored window
            ' would otherwise trigger resize/paint events on the panel
            ' during AddChild — before the add-in is fully initialized.
            Try
                _suppressHideEvent = True
                _dockableWindow.Visible = False
            Catch
            Finally
                _suppressHideEvent = False
            End Try

            ' Host the .NET control inside the Inventor DockableWindow
            _dockableWindow.AddChild(hwnd)

            ' Always enforce right-side docking so the panel never
            ' appears on the wrong side after an Inventor session restore.
            Try
                If _dockableWindow.DockingState <> DockingStateEnum.kDockRight Then
                    _dockableWindow.DockingState = DockingStateEnum.kDockRight
                End If
            Catch
            End Try

            ' Ensure hidden after AddChild (belt-and-suspenders).
            _suppressHideEvent = True
            Try
                _dockableWindow.Visible = False
            Finally
                _suppressHideEvent = False
            End Try

            _materialized = True

            Trace.TraceInformation(
                ": HMG DockableWindowManager: Materialized DockableWindow. " &
                "HWND=" & hwnd.ToString())

        Catch ex As Exception
            ' Non-fatal: the add-in works without docking (falls back to forms)
            Trace.TraceWarning(
                ": HMG DockableWindowManager: Materialization failed — " &
                ex.Message & ". Docked panel will not be available.")
            CleanupPartial()
        End Try
    End Sub

    ''' <summary>
    ''' Releases resources.  Call during add-in Deactivate.
    ''' </summary>
    Public Sub Cleanup()
        ' Unsubscribe from environment change events
        Try
            If _uiEvents IsNot Nothing Then
                RemoveHandler _uiEvents.OnEnvironmentChange,
                    AddressOf OnEnvironmentChange
                _uiEvents = Nothing
            End If
        Catch
        End Try

        ' Unsubscribe from document activation events
        Try
            If _appEvents IsNot Nothing Then
                RemoveHandler _appEvents.OnActivateDocument,
                    AddressOf OnActivateDocument
                RemoveHandler _appEvents.OnDeactivateDocument,
                    AddressOf OnDeactivateDocument
                _appEvents = Nothing
            End If
        Catch
        End Try

        ' Unsubscribe from dockable window hide events
        Try
            If _dwEvents IsNot Nothing Then
                RemoveHandler _dwEvents.OnHide, AddressOf OnDockableWindowHide
                _dwEvents = Nothing
            End If
        Catch
        End Try

        Try
            If _dockableWindow IsNot Nothing Then
                _suppressHideEvent = True
                _dockableWindow.Visible = False
                _suppressHideEvent = False
            End If
        Catch
            _suppressHideEvent = False
        End Try

        _summaryCache.Clear()
        _userDismissedDocs.Clear()
        CleanupPartial()
        _initialized = False
        _materialized = False
        Trace.TraceInformation(": HMG DockableWindowManager: Cleaned up.")
    End Sub

    Private Sub CleanupPartial()
        If _panel IsNot Nothing Then
            _panel.Dispose()
            _panel = Nothing
        End If
        _dockableWindow = Nothing
        _materialized = False
    End Sub

    ' ==================================================================
    '  Boundary source workflow
    ' ==================================================================

    ''' <summary>
    ''' Shows the boundary source panel in the docked window and blocks
    ''' until the user clicks Continue or Cancel.
    ''' Returns True if the user clicked Continue. The caller reads
    ''' <see cref="HmgPanelControl.ResultProjectName"/> and
    ''' <see cref="HmgPanelControl.ResultSourceType"/> from the panel.
    ''' </summary>
    Public Function ShowBoundarySourcePanel(
            Optional defaultProjectName As String = "Grating",
            Optional namedSketchLookup As NamedSketchLookupResult = Nothing
        ) As Boolean
        If Not IsAvailable Then Return False

        _panel.EnterBoundarySourceMode(defaultProjectName, namedSketchLookup)
        _dockableWindow.Visible = True

        Trace.TraceInformation(
            ": HMG DockableWindowManager: Showing boundary source panel.")

        Do
            While Not _panel.IsActionTaken
                System.Windows.Forms.Application.DoEvents()
                System.Threading.Thread.Sleep(50)
            End While

            If _panel.LastAction = HmgPanelControl.UserAction.Cancel Then
                Trace.TraceInformation(
                    ": HMG DockableWindowManager: Boundary source cancelled.")
                _panel.ResetToIdle()
                Return False
            End If

            If ValidateBoundarySourceSelection(
                    _panel.ResultSourceType, namedSketchLookup) Then
                Trace.TraceInformation(
                    ": HMG DockableWindowManager: Boundary source accepted — " &
                    "project=""" & _panel.ResultProjectName & """, source=" &
                    _panel.ResultSourceType.ToString())
                Return True
            End If

            ' Sketch/source not ready — keep panel open for another Continue
            Trace.TraceInformation(
                ": HMG DockableWindowManager: Boundary source validation failed — " &
                "waiting for user to fix selection and click Continue again.")
            _panel.ResetBoundarySourceAction()
        Loop
    End Function

    ''' <summary>
    ''' Returns False and shows a message when the chosen boundary source
    ''' cannot be used yet (e.g. no sketch selected).
    ''' </summary>
    Private Function ValidateBoundarySourceSelection(
            sourceType As BoundarySourceType,
            namedSketchLookup As NamedSketchLookupResult) As Boolean

        Select Case sourceType
            Case BoundarySourceType.SelectedSketch
                Dim check As SelectionResult =
                    New PerimeterSelectionService(_app).ValidateSketchResolvable()
                If check.Success Then Return True

                System.Windows.Forms.MessageBox.Show(
                    check.ErrorMessage,
                    "Metal Bar Grating",
                    System.Windows.Forms.MessageBoxButtons.OK,
                    System.Windows.Forms.MessageBoxIcon.Exclamation)
                Return False

            Case BoundarySourceType.NamedSketch
                If namedSketchLookup IsNot Nothing AndAlso namedSketchLookup.Found Then
                    Return True
                End If

                System.Windows.Forms.MessageBox.Show(
                    "The named boundary sketch """ &
                    BoundarySourceService.PrimaryName &
                    """ was not found in this Part document.",
                    "Metal Bar Grating",
                    System.Windows.Forms.MessageBoxButtons.OK,
                    System.Windows.Forms.MessageBoxIcon.Exclamation)
                Return False

            Case Else
                System.Windows.Forms.MessageBox.Show(
                    "The """ & sourceType.ToString() &
                    """ boundary source is not yet available.",
                    "Metal Bar Grating",
                    System.Windows.Forms.MessageBoxButtons.OK,
                    System.Windows.Forms.MessageBoxIcon.Information)
                Return False
        End Select
    End Function

    ''' <summary>The panel control for reading results after boundary source selection.</summary>
    Public ReadOnly Property Panel As HmgPanelControl
        Get
            Return _panel
        End Get
    End Property

    ' ==================================================================
    '  Parameter input workflow
    ' ==================================================================

    ''' <summary>
    ''' Shows the docked panel in parameter input mode pre-populated with
    ''' the given defaults and blocks until the user clicks OK or Cancel.
    '''
    ''' Returns the validated <see cref="GratingParameters"/> if OK,
    ''' or Nothing if cancelled.
    ''' </summary>
    Public Function ShowParameterInputPanel(
            defaults As GratingParameters) As GratingParameters
        If Not IsAvailable Then Return Nothing

        _panel.EnterParameterInputMode(defaults)
        _dockableWindow.Visible = True

        Trace.TraceInformation(
            ": HMG DockableWindowManager: Showing parameter input panel.")

        While Not _panel.IsActionTaken
            System.Windows.Forms.Application.DoEvents()
            System.Threading.Thread.Sleep(50)
        End While

        Dim result As GratingParameters = Nothing
        If _panel.LastAction = HmgPanelControl.UserAction.Done Then
            result = _panel.ResultParameters
        End If

        Trace.TraceInformation(
            ": HMG DockableWindowManager: Parameter input dismissed — " &
            If(result IsNot Nothing, "OK", "Cancel"))

        _panel.ResetToIdle()
        Return result
    End Function

    ' ==================================================================
    '  Panel visibility
    ' ==================================================================

    ''' <summary>
    ''' Switches the panel to progress mode with a progress bar.
    ''' Call <see cref="UpdateProgress"/> to advance the bar.
    ''' Call <see cref="ShowPanel"/> or <see cref="ResetToIdle"/>
    ''' when generation is complete.
    ''' </summary>
    Public Sub ShowProgressPanel(Optional title As String = "Generating Grating...")
        If Not IsAvailable Then Return
        _panel.EnterProgressMode(title)
        _dockableWindow.Visible = True
        Trace.TraceInformation(
            ": HMG DockableWindowManager: Showing progress panel.")
    End Sub

    ''' <summary>
    ''' Updates the progress bar percentage and step text.
    ''' </summary>
    Public Sub UpdateProgress(percent As Integer, stepText As String)
        If Not IsAvailable Then Return
        _panel.UpdateProgress(percent, stepText)
        ' Pump messages so the UI repaints immediately
        System.Windows.Forms.Application.DoEvents()
    End Sub

    ''' <summary>
    ''' Resets the panel to idle after progress completes.
    ''' </summary>
    Public Sub ResetToIdle()
        If Not IsAvailable Then Return
        _panel.ResetToIdle()
    End Sub

    ''' <summary>
    ''' Switches the panel to summary mode showing the generation
    ''' result text.  The panel stays visible so the user can review
    ''' the results while inspecting the assembly.
    ''' </summary>
    Public Sub ShowSummaryPanel(summaryText As String)
        ShowSummaryPanel(summaryText, Nothing)
    End Sub

    ''' <summary>
    ''' Switches the panel to summary mode and wires the "Create Drawing"
    ''' button to generate a fabrication drawing from the given project.
    ''' </summary>
    Public Sub ShowSummaryPanel(summaryText As String, project As GratingProject)
        If Not IsAvailable Then Return

        ' Store per-document so we can restore when returning.
        ' Also clear any prior user-dismissal for this doc — a brand
        ' new summary is the one case where we DO want the panel back.
        Try
            Dim docKey As String = _app.ActiveDocument.InternalName
            _summaryCache(docKey) = New CachedSummary With {
                .SummaryText = summaryText,
                .Project = project
            }
            _userDismissedDocs.Remove(docKey)
        Catch
        End Try

        ' Wire up the Create Drawing button callback
        If project IsNot Nothing Then
            _panel.SetCreateDrawingCallback(
                Sub()
                    Try
                        Dim drawingGen As New DrawingGenerator(_app)
                        project.DrawingResult = drawingGen.Generate(project)
                        If project.DrawingResult IsNot Nothing AndAlso
                           project.DrawingResult.Success Then
                            Trace.TraceInformation(
                                ": HMG: Drawing created — " &
                                project.DrawingResult.DrawingFilePath)
                        Else
                            Dim msg = If(project.DrawingResult?.ErrorMessage,
                                         "Unknown error")
                            Trace.TraceWarning(
                                ": HMG: Drawing generation failed — " & msg)
                            System.Windows.Forms.MessageBox.Show(
                                "Drawing generation failed:" & vbCrLf & msg,
                                "Metal Bar Grating",
                                System.Windows.Forms.MessageBoxButtons.OK,
                                System.Windows.Forms.MessageBoxIcon.Warning)
                        End If
                    Catch ex As Exception
                        Trace.TraceWarning(
                            ": HMG: Drawing generation error — " & ex.Message)
                        Throw
                    End Try
                End Sub)
        Else
            _panel.SetCreateDrawingCallback(Nothing)
        End If

        _panel.EnterSummaryMode(summaryText)
        _dockableWindow.Visible = True
        Trace.TraceInformation(
            ": HMG DockableWindowManager: Showing summary panel.")
        NotifySummaryAvailabilityChanged()
    End Sub

    ''' <summary>Shows the docked panel (idle content).</summary>
    Public Sub ShowPanel()
        If Not IsAvailable Then Return
        _panel.ResetToIdle()
        _dockableWindow.Visible = True
    End Sub

    ''' <summary>Hides the docked panel.</summary>
    Public Sub HidePanel()
        If Not IsAvailable Then Return
        Try
            _suppressHideEvent = True
            _dockableWindow.Visible = False
        Catch
        Finally
            _suppressHideEvent = False
        End Try
    End Sub

    ' ==================================================================
    '  Environment-aware summary
    ' ==================================================================

    ''' <summary>
    ''' Returns True when the active document is an assembly.
    ''' </summary>
    Private Function IsAssemblyEnvironment() As Boolean
        Try
            Dim activeDoc As Document = _app.ActiveDocument
            If activeDoc Is Nothing Then Return False
            Return activeDoc.DocumentType =
                DocumentTypeEnum.kAssemblyDocumentObject
        Catch
            Return False
        End Try
    End Function

    ''' <summary>
    ''' Returns True when the active document is a drawing (.idw/.dwg).
    ''' </summary>
    Private Function IsDrawingEnvironment() As Boolean
        Try
            Dim activeDoc As Document = _app.ActiveDocument
            If activeDoc Is Nothing Then Return False
            Return activeDoc.DocumentType =
                DocumentTypeEnum.kDrawingDocumentObject
        Catch
            Return False
        End Try
    End Function

    ''' <summary>
    ''' Fired by Inventor whenever any DockableWindow becomes hidden.
    ''' When our panel is hidden by the user (X button), record the
    ''' dismissal for the active document so auto-restore in the
    ''' environment/document handlers stops re-showing it.
    ''' </summary>
    Private Sub OnDockableWindowHide(
            DockableWindow As DockableWindow,
            BeforeOrAfter As EventTimingEnum,
            Context As NameValueMap,
            ByRef HandlingCode As HandlingCodeEnum)

        If BeforeOrAfter <> EventTimingEnum.kAfter Then Return
        If _suppressHideEvent Then Return

        Try
            If DockableWindow Is Nothing Then Return
            If DockableWindow.InternalName <> InternalName Then Return
        Catch
            Return
        End Try

        Try
            Dim docKey As String = _app.ActiveDocument.InternalName
            _userDismissedDocs.Add(docKey)
            Trace.TraceInformation(
                ": HMG DockableWindowManager: User dismissed panel for doc '" &
                docKey & "'. Auto-restore suppressed until next generation.")
        Catch ex As Exception
            Trace.TraceWarning(
                ": HMG DockableWindowManager: Hide event tracking failed — " & ex.Message)
        End Try
        ' Cache is unchanged but the "user can re-open" state is — the
        ' ribbon button reads cache presence (not dismissal), so this
        ' specific notification is a no-op for it.  Kept for symmetry.
        NotifySummaryAvailabilityChanged()
    End Sub

    ''' <summary>
    ''' Handles Inventor environment changes.  Restores the summary
    ''' panel when the user returns to an assembly document and hides
    ''' it when they leave.
    ''' </summary>
    Private Sub OnEnvironmentChange(
            environment As Inventor.Environment,
            environmentState As EnvironmentStateEnum,
            beforeOrAfter As EventTimingEnum,
            context As NameValueMap,
            ByRef handlingCode As HandlingCodeEnum)

        ' Only act after the new environment is fully active
        If beforeOrAfter <> EventTimingEnum.kAfter Then Return
        If environmentState <> EnvironmentStateEnum.kActivateEnvironmentState Then Return

        ' Notify the ribbon button regardless of panel materialization —
        ' the assembly-environment check changes here even when the panel
        ' itself was never opened.
        NotifySummaryAvailabilityChanged()

        ' Never trigger materialization from an environment event —
        ' only act if the panel was already created by user action.
        If Not _materialized Then Return
        If _panel Is Nothing OrElse _dockableWindow Is Nothing Then Return

        ' If the user explicitly closed the panel for this doc, do not
        ' auto-restore it on any subsequent environment events.
        If IsCurrentDocDismissed() Then Return

        Dim cached As CachedSummary = TryGetCachedSummary()
        If IsAssemblyEnvironment() AndAlso cached IsNot Nothing Then
            ' Entering assembly environment — restore this doc's summary
            RestoreSummary(cached)
            Trace.TraceInformation(
                ": HMG DockableWindowManager: Restored summary " &
                "in assembly environment.")
        ElseIf IsDrawingEnvironment() Then
            ' Entering drawing environment — show drawing panel
            _panel.EnterDrawingMode()
            _dockableWindow.Visible = True
            Trace.TraceInformation(
                ": HMG DockableWindowManager: Entered drawing environment.")
        ElseIf _panel.Mode = HmgPanelControl.PanelMode.Summary Then
            ' Leaving assembly environment — hide summary
            _panel.ResetToIdle()
            Trace.TraceInformation(
                ": HMG DockableWindowManager: Hid summary — " &
                "left assembly environment.")
        End If
    End Sub

    ''' <summary>
    ''' Handles document activation.  When the user switches away from
    ''' the assembly that generated the summary, the panel resets to
    ''' idle but the cached summary is preserved.  Switching back to
    ''' the assembly restores the summary and "Create Drawing" button.
    ''' </summary>
    Private Sub OnActivateDocument(
            DocumentObject As _Document,
            BeforeOrAfter As EventTimingEnum,
            Context As NameValueMap,
            ByRef HandlingCode As HandlingCodeEnum)

        If BeforeOrAfter <> EventTimingEnum.kAfter Then Return

        ' Notify the ribbon button regardless of panel materialization —
        ' the active-doc identity changed, so the button's enabled state
        ' may need to flip even if the dockable window was never opened.
        NotifySummaryAvailabilityChanged()

        ' Don't trigger materialization from document activation events.
        ' Only act if the panel has already been created by user action.
        If Not _materialized Then Return
        If _panel Is Nothing OrElse _dockableWindow Is Nothing Then Return

        ' If the user explicitly closed the panel for the new active
        ' document, do not auto-restore it.
        If IsCurrentDocDismissed() Then Return

        Try
            Dim cached As CachedSummary = TryGetCachedSummary()
            If cached IsNot Nothing Then
                ' This document has a cached summary — restore it
                RestoreSummary(cached)
                Trace.TraceInformation(
                    ": HMG DockableWindowManager: Restored summary — " &
                    "returned to assembly document.")
            ElseIf IsDrawingEnvironment() Then
                ' Drawing document — show drawing panel
                _panel.EnterDrawingMode()
                _dockableWindow.Visible = True
                Trace.TraceInformation(
                    ": HMG DockableWindowManager: Switched to drawing document.")
            Else
                ' No cached summary and not a drawing — the user is on
                ' a doc/Home with nothing for us to show.  Hide the
                ' window outright so the stale tab does not linger.
                If _panel.Mode = HmgPanelControl.PanelMode.Summary OrElse
                   _panel.Mode = HmgPanelControl.PanelMode.DrawingView Then
                    _panel.ResetToIdle()
                End If
                Try
                    _suppressHideEvent = True
                    _dockableWindow.Visible = False
                Catch
                Finally
                    _suppressHideEvent = False
                End Try
                Trace.TraceInformation(
                    ": HMG DockableWindowManager: Switched to document " &
                    "without cached summary — hidden.")
            End If
        Catch
        End Try
    End Sub

    ''' <summary>
    ''' Fires before a document is deactivated (switch away, close,
    ''' or transition to the Home page).  Hides the dockable window
    ''' immediately so the previous document's panel state does not
    ''' linger visually while Inventor repaints the new environment.
    ''' Auto-restore (if applicable) happens in OnActivateDocument or
    ''' OnEnvironmentChange once the new context is fully active.
    ''' </summary>
    Private Sub OnDeactivateDocument(
            DocumentObject As _Document,
            BeforeOrAfter As EventTimingEnum,
            Context As NameValueMap,
            ByRef HandlingCode As HandlingCodeEnum)

        If BeforeOrAfter <> EventTimingEnum.kBefore Then Return

        ' Notify the ribbon button — the active doc is about to change.
        NotifySummaryAvailabilityChanged()

        If Not _materialized Then Return
        If _dockableWindow Is Nothing Then Return

        Try
            _suppressHideEvent = True
            _dockableWindow.Visible = False
        Catch
        Finally
            _suppressHideEvent = False
        End Try
    End Sub

    ' ==================================================================
    '  Cache helpers
    ' ==================================================================

    ''' <summary>
    ''' Returns True if the user has explicitly dismissed the panel
    ''' for the currently active document (X button).  Used by the
    ''' auto-restore handlers to skip re-showing the panel.
    ''' </summary>
    Private Function IsCurrentDocDismissed() As Boolean
        Try
            Dim docKey As String = _app.ActiveDocument.InternalName
            Return _userDismissedDocs.Contains(docKey)
        Catch
            Return False
        End Try
    End Function

    ''' <summary>
    ''' Returns the cached summary for the currently active document,
    ''' or Nothing if no summary is cached for it.
    ''' </summary>
    Private Function TryGetCachedSummary() As CachedSummary
        Try
            Dim docKey As String = _app.ActiveDocument.InternalName
            Dim result As CachedSummary = Nothing
            If _summaryCache.TryGetValue(docKey, result) Then
                Return result
            End If
        Catch
        End Try
        Return Nothing
    End Function

    ''' <summary>
    ''' Restores a cached summary onto the panel, re-wiring the
    ''' Create Drawing callback if a project is available.
    ''' </summary>
    Private Sub RestoreSummary(cached As CachedSummary)
        If cached.Project IsNot Nothing Then
            ShowSummaryPanel(cached.SummaryText, cached.Project)
        Else
            _panel.EnterSummaryMode(cached.SummaryText)
            _dockableWindow.Visible = True
        End If
    End Sub

End Class
