'////////////////////////////////////////////////////////////////////
' Metal Bar Grating Addin - VB.NET
'
' BearingBarLayoutResult: Aggregates the output of the bearing bar
' layout engine — the list of trimmed bars plus summary metadata.
'////////////////////////////////////////////////////////////////////

''' <summary>
''' Result of the bearing bar layout generation.
''' </summary>
Public Class BearingBarLayoutResult

    ''' <summary>True if at least one bar was generated.</summary>
    Public Property Success As Boolean

    ''' <summary>Human-readable error when Success is False.</summary>
    Public Property ErrorMessage As String

    ''' <summary>Generated bearing bars (empty when Success is False).</summary>
    Public Property Bars As List(Of TrimmedBearingBar)

    ''' <summary>Non-fatal warnings encountered during layout.</summary>
    Public Property Warnings As List(Of String)

    ''' <summary>Span direction used for layout.</summary>
    Public Property SpanDirection As SpanDirectionType

    ''' <summary>Bounding box min {X, Y} of the perimeter (inches).</summary>
    Public Property BoundsMin As Double()

    ''' <summary>Bounding box max {X, Y} of the perimeter (inches).</summary>
    Public Property BoundsMax As Double()

    ''' <summary>On-center spacing that was applied (inches).</summary>
    Public Property AppliedSpacing As Double

    ''' <summary>
    ''' Perimeter edges (identified by the {x,y} pair of their two
    ''' endpoints) whose band bar was eliminated by the galvanize-gap
    ''' rule.  These edges must be skipped by the band bar generator so
    ''' the assembly does not double up a band bar against the bearing
    ''' bar / cross bar that the rule chose to keep.  See
    ''' BearingBarLayoutService.ApplyMinGalvanizeGap for the rule.
    '''
    ''' Each entry is {ax, ay, bx, by}.  Match by coordinate so the list
    ''' is robust to whether the polygon vertex list is closed or open.
    ''' </summary>
    Public Property EliminatedBandBarEdges As List(Of Double())

    ''' <summary>
    ''' Span-axis coordinates of cross bar positions that were eliminated
    ''' by the galvanize-gap rule (cross bar landing within 1/4" of a
    ''' perpendicular notch-wall band bar).  Cross bar generation must
    ''' skip these positions.
    ''' </summary>
    Public Property EliminatedCrossBarPositions As List(Of Double)

    ''' <summary>
    ''' Number of distinct lateral scan positions used (after the galvanize-
    ''' gap filter).  At a polygon notch each position may produce two or
    ''' more <see cref="TrimmedBearingBar"/> segments; <see cref="Bars"/>.Count
    ''' reports those segments (= fabrication pieces), whereas this property
    ''' reports the count the user sees on a design drawing.
    ''' </summary>
    Public Property UniqueLateralPositions As Integer

    ' --- Factory helpers ---

    Public Shared Function Succeeded(bars As List(Of TrimmedBearingBar),
                                     direction As SpanDirectionType,
                                     spacing As Double,
                                     boundsMin As Double(),
                                     boundsMax As Double(),
                                     warnings As List(Of String),
                                     Optional eliminatedEdges As List(Of Double()) = Nothing,
                                     Optional eliminatedCrossBarPositions As List(Of Double) = Nothing,
                                     Optional uniqueLateralPositions As Integer = 0) As BearingBarLayoutResult
        Return New BearingBarLayoutResult With {
            .Success = True,
            .Bars = bars,
            .SpanDirection = direction,
            .AppliedSpacing = spacing,
            .BoundsMin = boundsMin,
            .BoundsMax = boundsMax,
            .Warnings = If(warnings, New List(Of String)),
            .EliminatedBandBarEdges = If(eliminatedEdges, New List(Of Double())),
            .EliminatedCrossBarPositions = If(eliminatedCrossBarPositions, New List(Of Double)),
            .UniqueLateralPositions = uniqueLateralPositions
        }
    End Function

    Public Shared Function Failed(message As String) As BearingBarLayoutResult
        Return New BearingBarLayoutResult With {
            .Success = False,
            .ErrorMessage = message,
            .Bars = New List(Of TrimmedBearingBar),
            .Warnings = New List(Of String),
            .EliminatedBandBarEdges = New List(Of Double()),
            .EliminatedCrossBarPositions = New List(Of Double),
            .UniqueLateralPositions = 0
        }
    End Function

    ''' <summary>Builds a multi-line summary string for UI/logging.</summary>
    Public Function ToSummary() As String
        If Not Success Then Return "Layout failed: " & ErrorMessage

        Dim sb As New System.Text.StringBuilder()
        sb.AppendLine("Bearing Bar Layout Summary")
        sb.AppendLine("——————————————————————————")
        sb.AppendLine("Span direction : " & SpanDirection.ToString())
        If UniqueLateralPositions > 0 AndAlso
           UniqueLateralPositions <> Bars.Count Then
            sb.AppendLine("Bearing bars   : " & UniqueLateralPositions &
                          " (positions) · " & Bars.Count & " pieces")
        Else
            sb.AppendLine("Bearing bars   : " & Bars.Count)
        End If
        sb.AppendLine("On-center spacing: " & AppliedSpacing.ToString("F4") & " in")
        sb.AppendLine("Bounds min     : (" & BoundsMin(0).ToString("F4") & ", " & BoundsMin(1).ToString("F4") & ")")
        sb.AppendLine("Bounds max     : (" & BoundsMax(0).ToString("F4") & ", " & BoundsMax(1).ToString("F4") & ")")

        If Bars.Count > 0 Then
            sb.AppendLine()
            Dim preview As Integer = Math.Min(Bars.Count, 5)
            sb.AppendLine("First " & preview & " bars:")
            For i As Integer = 0 To preview - 1
                sb.AppendLine("  " & Bars(i).ToString())
            Next
            If Bars.Count > preview Then
                sb.AppendLine("  ... (" & (Bars.Count - preview) & " more)")
            End If
        End If

        If Warnings.Count > 0 Then
            sb.AppendLine()
            sb.AppendLine("Warnings:")
            For Each w As String In Warnings
                sb.AppendLine("  • " & w)
            Next
        End If

        Return sb.ToString()
    End Function

End Class
