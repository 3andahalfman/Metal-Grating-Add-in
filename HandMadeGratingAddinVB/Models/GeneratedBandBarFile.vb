'////////////////////////////////////////////////////////////////////
' Metal Bar Grating Addin - VB.NET
'
' GeneratedBandBarFile: Records the output of a single band bar
' part generation — the saved file path and source segment data.
'////////////////////////////////////////////////////////////////////

''' <summary>
''' Tracks one generated band bar .ipt file.
''' </summary>
Public Class GeneratedBandBarFile

    ''' <summary>1-based segment index around the perimeter.</summary>
    Public Property SegmentIndex As Integer

    ''' <summary>Mark/label for the band bar (e.g. "BAND-001").</summary>
    Public Property Mark As String

    ''' <summary>Length of the band bar segment in inches.</summary>
    Public Property Length As Double

    ''' <summary>Start point of the perimeter edge {X, Y} in inches.</summary>
    Public Property StartPoint As Double()

    ''' <summary>End point of the perimeter edge {X, Y} in inches.</summary>
    Public Property EndPoint As Double()

    ''' <summary>
    ''' Polygon winding sign (+1 CCW, -1 CW). Used at assembly placement so
    ''' the outer face lands on the perimeter for both sketch windings.
    ''' </summary>
    Public Property PerpSign As Double

    ''' <summary>Full path of the saved .ipt file.</summary>
    Public Property FilePath As String

    ''' <summary>File name only (no directory).</summary>
    Public Property FileName As String

    ''' <summary>True if the file was saved successfully.</summary>
    Public Property Saved As Boolean

    ''' <summary>Error message if Saved is False.</summary>
    Public Property ErrorMessage As String

    ''' <summary>
    ''' Non-fatal warning emitted during generation even when
    ''' <see cref="Saved"/> is True (e.g. when the mitered profile was
    ''' rejected and the bar was generated as a plain rectangle).
    ''' </summary>
    Public Property WarningMessage As String

    ''' <summary>True if this band bar is a curved arc bar.</summary>
    Public Property IsArc As Boolean

    ''' <summary>Arc centre X coordinate in inches (valid when IsArc).</summary>
    Public Property ArcCenterX As Double

    ''' <summary>Arc centre Y coordinate in inches (valid when IsArc).</summary>
    Public Property ArcCenterY As Double

    ''' <summary>Readable summary for trace/logging.</summary>
    Public Overrides Function ToString() As String
        If Saved Then
            Return Mark & " -> " & FileName &
                   " (L=" & Length.ToString("F4") & """)"
        Else
            Return Mark & " FAILED: " & ErrorMessage
        End If
    End Function

End Class
