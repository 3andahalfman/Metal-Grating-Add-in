'////////////////////////////////////////////////////////////////////
' Metal Bar Grating Addin - VB.NET
'
' ScheduleRow: Represents one row in a fabrication schedule table.
' Groups identical items (by mark/type and length) and tracks the
' count/quantity for each.
'
' Phase 19 — Fabrication schedule data layer.
'////////////////////////////////////////////////////////////////////

''' <summary>
''' Identifies which component type a schedule row represents.
''' </summary>
Public Enum ScheduleComponentType
    ''' <summary>Bearing bar (longitudinal member).</summary>
    BearingBar = 0
    ''' <summary>Cross bar (transverse member).</summary>
    CrossBar = 1
    ''' <summary>Band bar (perimeter member).</summary>
    BandBar = 2
End Enum

''' <summary>
''' One row in a fabrication schedule table.  Represents a group of
''' identical components sharing the same type, mark, and length.
''' All dimensions in inches.
''' </summary>
Public Class ScheduleRow

    ''' <summary>Component type (BearingBar, CrossBar, BandBar).</summary>
    Public Property ComponentType As ScheduleComponentType

    ''' <summary>
    ''' Mark or label that identifies this group (e.g. "CB-001", "BAND-003").
    ''' For grouped cross bars this is the representative mark or mark range.
    ''' </summary>
    Public Property Mark As String

    ''' <summary>
    ''' Cross bar type display name (e.g. "3/8"" Round").
    ''' Populated only for CrossBar rows; Nothing for others.
    ''' </summary>
    Public Property TypeDescription As String

    ''' <summary>Length of the component in inches.</summary>
    Public Property Length As Double

    ''' <summary>Number of identical components at this length.</summary>
    Public Property Quantity As Integer

    ''' <summary>
    ''' Individual marks of all items in this group when they were
    ''' merged by length.  Useful for mark-range display.
    ''' </summary>
    Public Property IndividualMarks As List(Of String)

    ''' <summary>
    ''' True if the file(s) for this row were saved successfully.
    ''' </summary>
    Public Property AllSaved As Boolean

    ''' <summary>
    ''' Human-readable mark range (e.g. "CB-001 – CB-005").
    ''' </summary>
    Public ReadOnly Property MarkRange As String
        Get
            If IndividualMarks Is Nothing OrElse IndividualMarks.Count = 0 Then
                Return If(Mark, "")
            End If
            If IndividualMarks.Count = 1 Then Return IndividualMarks(0)
            Return IndividualMarks(0) & " – " & IndividualMarks(IndividualMarks.Count - 1)
        End Get
    End Property

    ''' <summary>Length formatted as a fractional-inch string with quote.</summary>
    Public ReadOnly Property LengthDisplay As String
        Get
            Return Length.ToString("F4") & """"
        End Get
    End Property

    ''' <summary>Readable summary for trace/logging.</summary>
    Public Overrides Function ToString() As String
        Return ComponentType.ToString() & " | " &
               If(Mark, "?") & " | " &
               LengthDisplay & " x " & Quantity
    End Function

End Class
