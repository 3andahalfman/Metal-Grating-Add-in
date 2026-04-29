'////////////////////////////////////////////////////////////////////
' Metal Bar Grating Addin - VB.NET
'
' PerimeterArcInfo: Records metadata for a single arc entity
' found during perimeter extraction.  This allows the band bar
' generator to produce a single curved .ipt instead of many
' short straight segments for arc perimeter edges.
'
' All coordinates and dimensions are in inches (converted from
' Inventor's internal centimetres during extraction).
'////////////////////////////////////////////////////////////////////

''' <summary>
''' Metadata for one arc segment discovered during perimeter extraction.
''' </summary>
Public Class PerimeterArcInfo

    ''' <summary>Arc centre X coordinate in inches.</summary>
    Public Property CenterX As Double

    ''' <summary>Arc centre Y coordinate in inches.</summary>
    Public Property CenterY As Double

    ''' <summary>Arc radius in inches.</summary>
    Public Property Radius As Double

    ''' <summary>
    ''' Entry angle in radians (angle where the path enters the arc,
    ''' accounting for forward/reversed traversal).
    ''' </summary>
    Public Property EntryAngle As Double

    ''' <summary>
    ''' Signed sweep angle in radians.  Positive = CCW, negative = CW
    ''' relative to path traversal direction.
    ''' </summary>
    Public Property SweepAngle As Double

    ''' <summary>
    ''' Index of the first tessellated vertex emitted for this arc
    ''' in the OuterLoopVertices list.
    ''' </summary>
    Public Property FirstVertexIndex As Integer

    ''' <summary>
    ''' Number of tessellated vertices emitted for this arc
    ''' (entry point + intermediate points, excluding exit).
    ''' </summary>
    Public Property VertexCount As Integer

End Class
