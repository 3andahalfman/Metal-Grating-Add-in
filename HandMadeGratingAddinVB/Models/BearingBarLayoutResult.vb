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

    ' --- Factory helpers ---

    Public Shared Function Succeeded(bars As List(Of TrimmedBearingBar),
                                     direction As SpanDirectionType,
                                     spacing As Double,
                                     boundsMin As Double(),
                                     boundsMax As Double(),
                                     warnings As List(Of String)) As BearingBarLayoutResult
        Return New BearingBarLayoutResult With {
            .Success = True,
            .Bars = bars,
            .SpanDirection = direction,
            .AppliedSpacing = spacing,
            .BoundsMin = boundsMin,
            .BoundsMax = boundsMax,
            .Warnings = If(warnings, New List(Of String))
        }
    End Function

    Public Shared Function Failed(message As String) As BearingBarLayoutResult
        Return New BearingBarLayoutResult With {
            .Success = False,
            .ErrorMessage = message,
            .Bars = New List(Of TrimmedBearingBar),
            .Warnings = New List(Of String)
        }
    End Function

    ''' <summary>Builds a multi-line summary string for UI/logging.</summary>
    Public Function ToSummary() As String
        If Not Success Then Return "Layout failed: " & ErrorMessage

        Dim sb As New System.Text.StringBuilder()
        sb.AppendLine("Bearing Bar Layout Summary")
        sb.AppendLine("——————————————————————————")
        sb.AppendLine("Span direction : " & SpanDirection.ToString())
        sb.AppendLine("Bars generated : " & Bars.Count)
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
