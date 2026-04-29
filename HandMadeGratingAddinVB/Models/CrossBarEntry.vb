'////////////////////////////////////////////////////////////////////
' Metal Bar Grating Addin - VB.NET
'
' CrossBarEntry: Represents a single cross bar in the grating layout.
' Computed from the bearing bar notch positions and layout geometry.
'
' Cross bars run perpendicular to bearing bars and pass through the
' notches.  Each entry carries the absolute position along the
' bearing bar span direction and the computed length across the
' lateral (perpendicular) span of bearing bars at that position.
'////////////////////////////////////////////////////////////////////

''' <summary>
''' One cross bar in the grating with position and length.
''' All dimensions in inches.
''' </summary>
Public Class CrossBarEntry

    ''' <summary>1-based index within the cross bar layout.</summary>
    Public Property Index As Integer

    ''' <summary>
    ''' Absolute position along the bearing bar span direction (inches).
    ''' For AlongX bars this is an X coordinate; for AlongY, a Y coordinate.
    ''' </summary>
    Public Property AbsolutePosition As Double

    ''' <summary>
    ''' Computed length of the cross bar (inches).
    ''' Open-ended: outer face of first bearing bar to outer face of last.
    ''' Banded: inner face of first band bar to inner face of last band bar.
    ''' </summary>
    Public Property Length As Double

    ''' <summary>
    ''' Number of bearing bars this cross bar passes through.
    ''' </summary>
    Public Property BarsCrossed As Integer

    ''' <summary>
    ''' Actual start coordinate of the cross bar along the lateral axis
    ''' in world space (inches). Used directly for assembly placement.
    ''' </summary>
    Public Property LateralMin As Double

    ''' <summary>
    ''' Actual end coordinate of the cross bar along the lateral axis
    ''' in world space (inches).
    ''' </summary>
    Public Property LateralMax As Double

    ''' <summary>
    ''' Optional mark/label for the cross bar (e.g. "CB-001").
    ''' </summary>
    Public Property Mark As String

    ''' <summary>Readable summary for trace/logging.</summary>
    Public Overrides Function ToString() As String
        Return If(Mark, "CB-" & Index.ToString("000")) &
               ": pos=" & AbsolutePosition.ToString("F4") &
               " L=" & Length.ToString("F4") &
               " (" & BarsCrossed & " bars)"
    End Function

End Class
