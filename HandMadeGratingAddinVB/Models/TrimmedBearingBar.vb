'////////////////////////////////////////////////////////////////////
' Metal Bar Grating Addin - VB.NET
'
' TrimmedBearingBar: Represents a single bearing bar after layout
' generation and trimming to the perimeter boundary.
'////////////////////////////////////////////////////////////////////

''' <summary>
''' One bearing bar trimmed to the perimeter. Carries enough data
''' for downstream phases to create an Inventor part or assembly member.
''' </summary>
Public Class TrimmedBearingBar

    ''' <summary>1-based bar index within the layout.</summary>
    Public Property BarIndex As Integer

    ''' <summary>Optional mark/label for the bar (e.g. "BB-1").</summary>
    Public Property Mark As String

    ''' <summary>Start point {X, Y} in sketch coordinates (inches).</summary>
    Public Property StartPoint As Double()

    ''' <summary>End point {X, Y} in sketch coordinates (inches).</summary>
    Public Property EndPoint As Double()

    ''' <summary>Computed bar length between start and end (inches).</summary>
    Public Property Length As Double

    ''' <summary>The span direction this bar follows.</summary>
    Public Property SpanDirection As SpanDirectionType

    ''' <summary>
    ''' The lateral position (coordinate perpendicular to the span)
    ''' at which this bar sits (inches). Useful for spacing validation.
    ''' </summary>
    Public Property LateralPosition As Double

    ''' <summary>Readable summary for trace/logging.</summary>
    Public Overrides Function ToString() As String
        Return Mark & ": (" &
               StartPoint(0).ToString("F4") & ", " & StartPoint(1).ToString("F4") &
               ") -> (" &
               EndPoint(0).ToString("F4") & ", " & EndPoint(1).ToString("F4") &
               ") L=" & Length.ToString("F4")
    End Function

End Class
