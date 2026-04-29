'////////////////////////////////////////////////////////////////////
' Metal Bar Grating Addin - VB.NET
'
' GratingProjectStorage: Saves and loads GratingProject metadata to
' and from the active Inventor Part document using a custom iProperties
' property set named "Metal Bar Grating".
'
' Phase 14: Project persistence via Inventor iProperties.
'
' Design notes
' ------------
'  - All values are stored as strings in a single named property set.
'    This avoids COM type-mismatch issues across Inventor versions.
'  - Save writes all properties; missing keys are added, existing ones
'    are updated in-place.
'  - TryLoad is fully defensive: any missing or malformed property falls
'    back to a safe default without crashing the workflow.
'  - Boundary re-resolution delegates to BoundarySourceService, which
'    already handles the Phase 12 named-sketch pipeline.  ImportedDwg
'    projects are treated identically (they use the same GRATING_BOUNDARY
'    sketch after import).
'
' Property set name:  "Metal Bar Grating"
' All property keys are prefixed "HMG." to avoid name clashes with
' standard or third-party iProperties.
'////////////////////////////////////////////////////////////////////

Imports System.Diagnostics
Imports System.Globalization
Imports Inventor

''' <summary>
''' Saves and restores GratingProject metadata via Inventor custom
''' iProperties on the source Part document.
''' </summary>
Public Class GratingProjectStorage

    ' ==================================================================
    '  Property set and key constants
    ' ==================================================================

    ''' <summary>
    ''' Name of the Inventor custom iProperty set written to the Part.
    ''' Visible in File → iProperties → Custom tab as "Metal Bar Grating".
    ''' </summary>
    Public Const PropertySetName As String = "Metal Bar Grating"

    ' Project-level keys
    Private Const K_ProjectName As String = "HMG.ProjectName"
    Private Const K_SourceType As String = "HMG.BoundarySourceType"
    Private Const K_SketchName As String = "HMG.BoundarySketchName"
    Private Const K_ImportedPath As String = "HMG.ImportedFilePath"
    Private Const K_SavedUtc As String = "HMG.SavedUtc"

    ' Bearing bar parameter keys
    Private Const K_SpanDir As String = "HMG.SpanDirection"
    Private Const K_BarDepth As String = "HMG.BarDepth"
    Private Const K_BarWidth As String = "HMG.BarWidth"
    Private Const K_BarOC As String = "HMG.OnCenterSpacing"
    Private Const K_NamingPrefix As String = "HMG.NamingPrefix"
    Private Const K_OutputFolder As String = "HMG.OutputFolder"

    ' Cross bar parameter keys
    Private Const K_CrossBarType As String = "HMG.CrossBarType"
    Private Const K_CrossBarOC As String = "HMG.CrossBarOnCenter"
    Private Const K_CrossBarOffset As String = "HMG.FirstCrossBarOffset"

    ' Surface and banding keys
    Private Const K_Surface As String = "HMG.SurfaceProfile"
    Private Const K_Banding As String = "HMG.Banding"

    ' Invariant culture for numeric serialization
    Private Shared ReadOnly Inv As CultureInfo = CultureInfo.InvariantCulture

    Private ReadOnly _app As Application

    Public Sub New(app As Application)
        _app = app
    End Sub

    ' ==================================================================
    '  Save
    ' ==================================================================

    ''' <summary>
    ''' Writes the project metadata and parameters to the active Part's
    ''' custom iProperties.  Creates the "Metal Bar Grating" property set
    ''' if it does not already exist.
    ''' </summary>
    ''' <param name="partDoc">The Part document to write to.</param>
    ''' <param name="project">The project to persist.</param>
    Public Sub Save(partDoc As PartDocument, project As GratingProject)
        If partDoc Is Nothing Then
            Trace.TraceWarning(": HMG Storage: Save — partDoc is Nothing, skipped.")
            Return
        End If
        If project Is Nothing Then
            Trace.TraceWarning(": HMG Storage: Save — project is Nothing, skipped.")
            Return
        End If

        Try
            Dim propSet As Object = GetOrCreatePropertySet(partDoc)

            ' --- Project-level ---
            WriteProp(propSet, K_ProjectName,
                      If(project.ProjectName, "Grating"))

            Dim bs As BoundarySourceInfo = project.BoundarySource
            If bs IsNot Nothing Then
                WriteProp(propSet, K_SourceType,
                          CInt(bs.SourceType).ToString(Inv))
                WriteProp(propSet, K_SketchName,
                          If(bs.SketchName, ""))
                WriteProp(propSet, K_ImportedPath,
                          If(bs.ImportedFilePath, ""))
            End If

            WriteProp(propSet, K_SavedUtc,
                      DateTime.UtcNow.ToString("O", Inv))

            ' --- Parameters ---
            Dim p As GratingParameters = project.Parameters
            If p IsNot Nothing Then
                WriteProp(propSet, K_SpanDir,
                          CInt(p.SpanDirection).ToString(Inv))
                WriteProp(propSet, K_BarDepth,
                          p.BarDepth.ToString("R", Inv))
                WriteProp(propSet, K_BarWidth,
                          p.BarWidth.ToString("R", Inv))
                WriteProp(propSet, K_BarOC,
                          p.OnCenterSpacing.ToString("R", Inv))
                WriteProp(propSet, K_NamingPrefix,
                          If(p.NamingPrefix, ""))
                WriteProp(propSet, K_OutputFolder,
                          If(p.OutputFolder, ""))
                WriteProp(propSet, K_CrossBarType,
                          CInt(p.CrossBar).ToString(Inv))
                WriteProp(propSet, K_CrossBarOC,
                          p.CrossBarOnCenter.ToString("R", Inv))
                WriteProp(propSet, K_CrossBarOffset,
                          p.FirstCrossBarOffset.ToString("R", Inv))
                WriteProp(propSet, K_Surface,
                          CInt(p.SurfaceProfile).ToString(Inv))
                WriteProp(propSet, K_Banding,
                          CInt(p.Banding).ToString(Inv))
            End If

            Trace.TraceInformation(
                ": HMG Storage: Saved project """ &
                project.ProjectName & """ to iProperties on " &
                partDoc.DisplayName & ".")

        Catch ex As Exception
            Trace.TraceError(
                ": HMG Storage: Save failed on " &
                partDoc.DisplayName & " — " & ex.Message)
            Throw
        End Try
    End Sub

    ' ==================================================================
    '  TryLoad
    ' ==================================================================

    ''' <summary>
    ''' Attempts to read a previously saved GratingProject from the Part
    ''' document's iProperties.  Always returns a result; never throws.
    ''' </summary>
    ''' <param name="partDoc">The Part document to read from.</param>
    ''' <returns>
    ''' A <see cref="GratingProjectLoadResult"/> describing what was found.
    ''' </returns>
    Public Function TryLoad(partDoc As PartDocument) As GratingProjectLoadResult
        If partDoc Is Nothing Then
            Trace.TraceWarning(": HMG Storage: TryLoad — partDoc is Nothing.")
            Return GratingProjectLoadResult.NotFound()
        End If

        ' --- Locate property set ---
        Dim propSet As Object = Nothing
        Try
            propSet = partDoc.PropertySets.Item(PropertySetName)
        Catch
            ' Property set does not exist on this document
            Trace.TraceInformation(
                ": HMG Storage: No '" & PropertySetName & "' property set on " &
                partDoc.DisplayName & ".")
            Return GratingProjectLoadResult.NotFound()
        End Try

        Trace.TraceInformation(
            ": HMG Storage: Found '" & PropertySetName &
            "' property set on " & partDoc.DisplayName & ". Reading...")

        ' --- Read project-level fields ---
        Dim projectName As String = ReadProp(propSet, K_ProjectName, "Grating")
        Dim sourceTypeStr As String = ReadProp(propSet, K_SourceType, "0")
        Dim sketchName As String = ReadProp(propSet, K_SketchName, "")
        Dim importedPath As String = ReadProp(propSet, K_ImportedPath, "")
        Dim savedUtcStr As String = ReadProp(propSet, K_SavedUtc, "")

        ' Parse boundary source type
        Dim sourceType As BoundarySourceType = BoundarySourceType.NamedSketch
        Try
            sourceType = CType(CInt(sourceTypeStr), BoundarySourceType)
        Catch
            Trace.TraceWarning(
                ": HMG Storage: Unrecognised BoundarySourceType '" &
                sourceTypeStr & "', defaulting to NamedSketch.")
        End Try

        Trace.TraceInformation(
            ": HMG Storage: Project='" & projectName &
            "', SourceType=" & sourceType.ToString() &
            ", Sketch='" & sketchName & "'" &
            If(String.IsNullOrEmpty(savedUtcStr), "", ", Saved=" & savedUtcStr))

        ' --- Read parameters ---
        Dim parms As GratingParameters = ReadParameters(propSet)

        ' --- Build BoundarySourceInfo ---
        Dim source As BoundarySourceInfo
        Select Case sourceType
            Case BoundarySourceType.ImportedDwg
                source = BoundarySourceInfo.FromImportedDwg(
                    If(sketchName, BoundarySourceService.PrimaryName), importedPath)
            Case BoundarySourceType.SelectedSketch
                source = BoundarySourceInfo.FromSelectedSketch(
                    If(sketchName, "Sketch"))
            Case BoundarySourceType.NewSketch
                source = BoundarySourceInfo.FromNewSketch(
                    If(sketchName, BoundarySourceService.PrimaryName))
            Case Else  ' NamedSketch
                source = BoundarySourceInfo.FromNamedSketch(
                    If(sketchName, BoundarySourceService.PrimaryName))
        End Select

        ' --- Attempt boundary re-resolution ---
        Dim boundaryResolved As Boolean = False
        Dim resolvedResult As SelectionResult = Nothing
        Dim resolveMessage As String = ""

        Select Case sourceType
            Case BoundarySourceType.NamedSketch, BoundarySourceType.ImportedDwg
                ' Both use a named sketch (GRATING_BOUNDARY after import) —
                ' delegate to BoundarySourceService for standard Phase 12 pipeline.
                Dim resolveSketch As String =
                    If(Not String.IsNullOrEmpty(sketchName),
                       sketchName, BoundarySourceService.PrimaryName)
                Try
                    Dim svc As New BoundarySourceService(_app)
                    resolvedResult = svc.ResolveFromNamedSketch(resolveSketch)
                    If resolvedResult.Success Then
                        boundaryResolved = True
                        resolveMessage = "Boundary sketch '" & resolveSketch &
                                         "' resolved — " &
                                         resolvedResult.Perimeter.EdgeCount & " edges."
                        Trace.TraceInformation(": HMG Storage: " & resolveMessage)
                    Else
                        resolveMessage = "Boundary sketch '" & resolveSketch &
                                         "' not resolved — " & resolvedResult.ErrorMessage
                        Trace.TraceWarning(": HMG Storage: " & resolveMessage)
                    End If
                Catch ex As Exception
                    resolveMessage = "Exception resolving sketch '" &
                                     resolveSketch & "': " & ex.Message
                    Trace.TraceWarning(": HMG Storage: " & resolveMessage)
                End Try

            Case Else
                ' SelectedSketch and NewSketch cannot be auto-resolved at load time.
                resolveMessage = "Source type " & sourceType.ToString() &
                                 " cannot be auto-resolved at load time."
                Trace.TraceInformation(": HMG Storage: " & resolveMessage)
        End Select

        ' --- Build GratingProject ---
        Dim proj As GratingProject
        If boundaryResolved AndAlso resolvedResult IsNot Nothing Then
            proj = GratingProject.Create(projectName, source, resolvedResult)
        Else
            ' Metadata only — Perimeter will be established through normal flow
            proj = New GratingProject With {
                .ProjectName = projectName,
                .BoundarySource = source,
                .CreatedUtc = DateTime.UtcNow,
                .LastModifiedUtc = DateTime.UtcNow
            }
        End If

        Dim summaryMsg As String =
            "project='" & projectName & "', " &
            "sourceType=" & sourceType.ToString() &
            If(boundaryResolved, ", boundary resolved", ", boundary NOT resolved") &
            " — " & resolveMessage

        Return GratingProjectLoadResult.Loaded(
            proj, parms, boundaryResolved, summaryMsg)
    End Function

    ' ==================================================================
    '  Parameters deserialization helper
    ' ==================================================================

    ''' <summary>
    ''' Reads all parameter keys from the property set.
    ''' Any missing or malformed value falls back to the program default
    ''' without throwing.
    ''' </summary>
    Private Shared Function ReadParameters(propSet As Object) As GratingParameters
        Dim d As GratingParameters = GratingParameters.CreateDefaults()

        d.SpanDirection = CType(
            ParseInt(ReadProp(propSet, K_SpanDir, "0"), CInt(d.SpanDirection)),
            SpanDirectionType)

        d.BarDepth = ParseDouble(
            ReadProp(propSet, K_BarDepth, ""), d.BarDepth)

        d.BarWidth = ParseDouble(
            ReadProp(propSet, K_BarWidth, ""), d.BarWidth)

        d.OnCenterSpacing = ParseDouble(
            ReadProp(propSet, K_BarOC, ""), d.OnCenterSpacing)

        d.NamingPrefix = ReadProp(propSet, K_NamingPrefix, d.NamingPrefix)
        d.OutputFolder = ReadProp(propSet, K_OutputFolder, "")

        d.CrossBar = CType(
            ParseInt(ReadProp(propSet, K_CrossBarType, "0"), CInt(d.CrossBar)),
            CrossBarType)

        d.CrossBarOnCenter = ParseDouble(
            ReadProp(propSet, K_CrossBarOC, ""), d.CrossBarOnCenter)

        d.FirstCrossBarOffset = ParseDouble(
            ReadProp(propSet, K_CrossBarOffset, ""), d.FirstCrossBarOffset)

        d.SurfaceProfile = CType(
            ParseInt(ReadProp(propSet, K_Surface, "0"), CInt(d.SurfaceProfile)),
            SurfaceProfileType)

        d.Banding = CType(
            ParseInt(ReadProp(propSet, K_Banding, "0"), CInt(d.Banding)),
            BandingOptionType)

        Return d
    End Function

    ' ==================================================================
    '  iProperty read / write helpers
    ' ==================================================================

    ''' <summary>
    ''' Gets the "Metal Bar Grating" property set, creating it if absent.
    ''' </summary>
    Private Shared Function GetOrCreatePropertySet(
            partDoc As PartDocument) As Object
        Try
            Return partDoc.PropertySets.Item(PropertySetName)
        Catch
            ' Property set not found — create a new one
            Trace.TraceInformation(
                ": HMG Storage: Creating new property set '" &
                PropertySetName & "'.")
            Return partDoc.PropertySets.Add(PropertySetName)
        End Try
    End Function

    ''' <summary>
    ''' Reads a string property by key.  Returns defaultVal if absent or
    ''' unreadable.  Logs a note for any missing key.
    ''' </summary>
    Private Shared Function ReadProp(propSet As Object,
                                     key As String,
                                     Optional defaultVal As String = "") As String
        Try
            Dim v As Object = propSet.Item(key).Value
            If v IsNot Nothing Then Return CStr(v)
        Catch
            Trace.TraceInformation(
                ": HMG Storage: Property '" & key &
                "' not found — using default '" & defaultVal & "'.")
        End Try
        Return defaultVal
    End Function

    ''' <summary>
    ''' Writes a string property.  Updates the value if the property
    ''' already exists; adds a new property if it does not.
    ''' Logs a warning if both attempts fail.
    ''' </summary>
    Private Shared Sub WriteProp(propSet As Object,
                                 key As String,
                                 value As String)
        Try
            propSet.Item(key).Value = value
            Return  ' updated successfully
        Catch
        End Try
        Try
            propSet.Add(value, key)
        Catch ex As Exception
            Trace.TraceWarning(
                ": HMG Storage: Cannot write property '" & key & "' — " &
                ex.Message)
        End Try
    End Sub

    ' ==================================================================
    '  Parsing helpers
    ' ==================================================================

    Private Shared Function ParseDouble(s As String, fallback As Double) As Double
        If String.IsNullOrEmpty(s) Then Return fallback
        Dim result As Double
        If Double.TryParse(s, Globalization.NumberStyles.Float,
                           Inv, result) Then
            Return result
        End If
        Trace.TraceWarning(
            ": HMG Storage: Cannot parse double '" & s &
            "', using " & fallback.ToString(Inv) & ".")
        Return fallback
    End Function

    Private Shared Function ParseInt(s As String, fallback As Integer) As Integer
        If String.IsNullOrEmpty(s) Then Return fallback
        Dim result As Integer
        If Integer.TryParse(s, result) Then Return result
        Trace.TraceWarning(
            ": HMG Storage: Cannot parse integer '" & s &
            "', using " & fallback.ToString() & ".")
        Return fallback
    End Function

End Class
