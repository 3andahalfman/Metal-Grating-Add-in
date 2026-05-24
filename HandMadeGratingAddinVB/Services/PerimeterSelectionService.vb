'////////////////////////////////////////////////////////////////////
' Metal Bar Grating Addin - VB.NET
'
' PerimeterSelectionService: Orchestrates the perimeter selection
' workflow — document validation, sketch resolution, profile
' validation, and delegation to PerimeterExtractor.
'
' Sketch resolution strategies (tried in order):
'   1. Active edit sketch  — user is inside a sketch
'   2. Pre-selected object — user selected a sketch or sketch curve
'      before clicking the button
'
' Supported document types: Part (.ipt) only.
' Unsupported: Assembly, Drawing, Presentation, and ZeroDoc.
'////////////////////////////////////////////////////////////////////

Imports System.Diagnostics
Imports Inventor

''' <summary>
''' Manages the perimeter selection and validation workflow.
''' Delegates geometry extraction to <see cref="PerimeterExtractor"/>.
''' </summary>
Public Class PerimeterSelectionService

    Private ReadOnly _app As Application

    Public Sub New(app As Application)
        _app = app
    End Sub

    ''' <summary>
    ''' Runs the full perimeter selection workflow.
    ''' </summary>
    Public Function Execute() As SelectionResult

        ' --- 1. Document validation ---
        Dim doc As Document = GetActiveDocument()

        If doc Is Nothing Then
            Return SelectionResult.Failed(
                "No document is open." & vbCrLf &
                "Please open a Part document first.")
        End If

        If doc.DocumentType <> DocumentTypeEnum.kPartDocumentObject Then
            Dim typeName As String = GetDocumentTypeName(doc.DocumentType)
            Return SelectionResult.Failed(
                "Grating creation is only supported in Part documents." & vbCrLf &
                "Current document type: " & typeName & vbCrLf & vbCrLf &
                "Supported: Part (.ipt)")
        End If

        Trace.TraceInformation(": HMG: Document validated: " & doc.DisplayName)

        ' --- 2. Sketch resolution ---
        Dim sketch As PlanarSketch = ResolveSketch(doc)

        If sketch Is Nothing Then
            Return SelectionResult.Failed(
                "No sketch found." & vbCrLf & vbCrLf &
                "To select a grating perimeter, do one of the following" &
                " then click Create Grating:" & vbCrLf &
                "  1. Double-click a sketch to edit it, or" & vbCrLf &
                "  2. Single-click a sketch in the browser, or" & vbCrLf &
                "  3. Select sketch curves in the graphics area.")
        End If

        Trace.TraceInformation(": HMG: Sketch resolved: " & sketch.Name)

        ' --- 3. Profile validation ---
        '     Try to compute closed profiles. AddForSolid creates a Profile
        '     containing all closed regions. If the sketch has no closed
        '     loops this call throws.
        Dim profile As Profile = Nothing
        Try
            profile = sketch.Profiles.AddForSolid()
        Catch
            Return SelectionResult.Failed(
                "The sketch '" & sketch.Name &
                "' does not contain any closed regions." & vbCrLf & vbCrLf &
                "Ensure all curves connect end-to-end to form at " &
                "least one closed loop.")
        End Try

        If profile Is Nothing Then
            Return SelectionResult.Failed(
                "Could not compute profiles from sketch '" &
                sketch.Name & "'.")
        End If

        ' Find the outer boundary path (the one that adds material)
        Dim outerPath As ProfilePath = FindOuterPath(profile)

        If outerPath Is Nothing Then
            Return SelectionResult.Failed(
                "No outer boundary found in the sketch profiles." & vbCrLf &
                "The sketch must contain a closed loop that defines " &
                "the grating outline.")
        End If

        Dim pathCount As Integer = SafeCount(profile)
        Trace.TraceInformation(": HMG: Profile has " & pathCount &
                               " path(s). Outer boundary located.")

        ' --- 4. Geometry extraction ---
        Dim extractor As New PerimeterExtractor()
        Dim data As PerimeterData = extractor.ExtractFromPath(sketch, outerPath)

        If data Is Nothing Then
            Return SelectionResult.Failed(
                "Failed to extract perimeter geometry from sketch '" &
                sketch.Name & "'.")
        End If

        Trace.TraceInformation(": HMG: Perimeter extracted — " &
            data.EdgeCount & " edges, " &
            data.OuterLoopVertices.Count & " vertices.")

        Return SelectionResult.Succeeded(data)
    End Function

    ''' <summary>
    ''' Lightweight check used before boundary-source Continue is accepted.
    ''' Confirms the active Part has a resolvable sketch without extracting
    ''' perimeter geometry (so the user can select a sketch and retry).
    ''' </summary>
    Public Function ValidateSketchResolvable() As SelectionResult
        Dim doc As Document = GetActiveDocument()

        If doc Is Nothing Then
            Return SelectionResult.Failed(
                "No document is open." & vbCrLf &
                "Please open a Part document first.")
        End If

        If doc.DocumentType <> DocumentTypeEnum.kPartDocumentObject Then
            Dim typeName As String = GetDocumentTypeName(doc.DocumentType)
            Return SelectionResult.Failed(
                "Grating creation is only supported in Part documents." & vbCrLf &
                "Current document type: " & typeName & vbCrLf & vbCrLf &
                "Supported: Part (.ipt)")
        End If

        Dim sketch As PlanarSketch = ResolveSketch(doc)
        If sketch Is Nothing Then
            Return SelectionResult.Failed(
                "No sketch found." & vbCrLf & vbCrLf &
                "To select a grating perimeter, do one of the following" &
                " then click Continue:" & vbCrLf &
                "  1. Double-click a sketch to edit it, or" & vbCrLf &
                "  2. Single-click a sketch in the browser, or" & vbCrLf &
                "  3. Select sketch curves in the graphics area.")
        End If

        Return SelectionResult.Succeeded(Nothing)
    End Function

#Region "Document helpers"

    Private Function GetActiveDocument() As Document
        Try
            Return _app.ActiveDocument
        Catch
            Return Nothing
        End Try
    End Function

    Private Shared Function GetDocumentTypeName(docType As DocumentTypeEnum) As String
        Select Case docType
            Case DocumentTypeEnum.kAssemblyDocumentObject : Return "Assembly (.iam)"
            Case DocumentTypeEnum.kDrawingDocumentObject : Return "Drawing (.idw/.dwg)"
            Case DocumentTypeEnum.kPresentationDocumentObject : Return "Presentation (.ipn)"
            Case Else : Return "Unknown"
        End Select
    End Function

#End Region

#Region "Sketch resolution"

    ''' <summary>
    ''' Resolves a PlanarSketch using two strategies:
    '''   1. Active edit sketch (user is inside a sketch)
    '''   2. Pre-selected sketch or sketch entity in SelectSet
    ''' </summary>
    Private Function ResolveSketch(doc As Document) As PlanarSketch
        ' Strategy 1: Active edit sketch
        Dim sketch As PlanarSketch = TryGetActiveEditSketch()
        If sketch IsNot Nothing Then
            Trace.TraceInformation(": HMG: Using active edit sketch.")
            Return sketch
        End If

        ' Strategy 2: Pre-selected object
        sketch = TryGetSketchFromSelection(doc)
        If sketch IsNot Nothing Then
            Trace.TraceInformation(": HMG: Using pre-selected sketch.")
            Return sketch
        End If

        Return Nothing
    End Function

    Private Function TryGetActiveEditSketch() As PlanarSketch
        Try
            Dim activeObj As Object = _app.ActiveEditObject
            If TypeOf activeObj Is PlanarSketch Then
                Return CType(activeObj, PlanarSketch)
            End If
        Catch
        End Try
        Return Nothing
    End Function

    Private Function TryGetSketchFromSelection(doc As Document) As PlanarSketch
        Try
            Dim selectSet As SelectSet = doc.SelectSet
            For i As Integer = 1 To selectSet.Count
                Dim sketch As PlanarSketch =
                    TryResolveSketchFromObject(selectSet.Item(i))
                If sketch IsNot Nothing Then Return sketch
            Next
        Catch
        End Try
        Return Nothing
    End Function

    ''' <summary>
    ''' Attempts to walk up from any Inventor object to a PlanarSketch.
    ''' Supports direct PlanarSketch, sketch entities (via .Parent),
    ''' and deeper children (via .Parent.Parent).
    ''' Uses late binding (Option Strict Off) for COM compatibility.
    ''' </summary>
    Private Function TryResolveSketchFromObject(obj As Object) As PlanarSketch
        ' Direct sketch
        If TypeOf obj Is PlanarSketch Then
            Return CType(obj, PlanarSketch)
        End If

        ' Walk up via late-bound .Parent (handles SketchLine, SketchArc,
        ' SketchCircle, Profile, etc.)
        Try
            Dim parent As Object = obj.Parent
            If TypeOf parent Is PlanarSketch Then
                Return CType(parent, PlanarSketch)
            End If

            ' One more level (e.g., ProfilePath → Profile → PlanarSketch)
            Dim grandparent As Object = parent.Parent
            If TypeOf grandparent Is PlanarSketch Then
                Return CType(grandparent, PlanarSketch)
            End If
        Catch
        End Try

        Return Nothing
    End Function

#End Region

#Region "Profile validation helpers"

    ''' <summary>
    ''' Finds the first ProfilePath that adds material (outer boundary).
    ''' Falls back to the first path if AddsMaterial is not accessible.
    ''' </summary>
    Private Shared Function FindOuterPath(profile As Profile) As ProfilePath
        Try
            For i As Integer = 1 To profile.Count
                Dim path As ProfilePath = profile.Item(i)
                Try
                    If path.AddsMaterial Then Return path
                Catch
                End Try
            Next
        Catch
        End Try

        ' Fallback: first path is typically the outer boundary
        Try
            If profile.Count > 0 Then Return profile.Item(1)
        Catch
        End Try

        Return Nothing
    End Function

    Private Shared Function SafeCount(profile As Profile) As Integer
        Try
            Return profile.Count
        Catch
            Return 0
        End Try
    End Function

#End Region

End Class
