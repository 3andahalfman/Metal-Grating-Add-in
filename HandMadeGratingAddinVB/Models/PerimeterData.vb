'////////////////////////////////////////////////////////////////////
' Metal Bar Grating Addin - VB.NET
'
' PerimeterData: Internal model representing a validated grating
' perimeter extracted from an Inventor sketch.
' Holds both extracted vertex data (for inspection/logging) and a
' transient COM reference to the source sketch (for downstream phases).
'////////////////////////////////////////////////////////////////////

Imports Inventor

''' <summary>
''' Holds the extracted perimeter geometry from a validated sketch.
''' Used by downstream phases for grating generation.
''' </summary>
Public Class PerimeterData

    ''' <summary>Name of the source sketch.</summary>
    Public Property SketchName As String

    ''' <summary>Number of edges (profile entities) in the outer perimeter.</summary>
    Public Property EdgeCount As Integer

    ''' <summary>
    ''' Extracted 2D vertices of the outer perimeter loop in sketch coordinates.
    ''' Each element is a Double array {X, Y}.
    ''' </summary>
    Public Property OuterLoopVertices As List(Of Double())

    ''' <summary>
    ''' Metadata for arc entities found in the perimeter.
    ''' Each entry maps a range of tessellated vertices back to
    ''' the original arc geometry (centre, radius, angles).
    ''' May be Nothing or empty if no arcs are present.
    ''' </summary>
    Public Property ArcSegments As List(Of PerimeterArcInfo)

    ''' <summary>
    ''' Transient reference to the source PlanarSketch COM object.
    ''' Valid only within the current Inventor session. Callers must
    ''' check for Nothing before use.
    ''' </summary>
    Public Property SourceSketch As PlanarSketch

End Class
