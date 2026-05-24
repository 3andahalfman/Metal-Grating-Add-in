'////////////////////////////////////////////////////////////////////
' Metal Bar Grating Addin - VB.NET
'
' BearingBarGenerationResult: Aggregates the output of the bearing
' bar part generation phase — all generated files plus summary data.
'////////////////////////////////////////////////////////////////////

''' <summary>
''' Result of the Phase 6 bearing bar .ipt generation pass.
''' </summary>
Public Class BearingBarGenerationResult

    ''' <summary>True if at least one file was generated.</summary>
    Public Property Success As Boolean

    ''' <summary>Human-readable error when Success is False.</summary>
    Public Property ErrorMessage As String

    ''' <summary>Per-bar generation records.</summary>
    Public Property Files As List(Of GeneratedBearingBarFile)

    ''' <summary>Non-fatal warnings.</summary>
    Public Property Warnings As List(Of String)

    ''' <summary>Resolved output folder that was used.</summary>
    Public Property OutputFolder As String

    ''' <summary>
    ''' Number of distinct lateral scan positions used by the layout.
    ''' At a polygon notch each position may produce multiple
    ''' fabrication pieces (one per polygon entry/exit pair); this
    ''' property reports the count the user sees on a design drawing.
    ''' Forwarded from <see cref="BearingBarLayoutResult.UniqueLateralPositions"/>.
    ''' </summary>
    Public Property UniqueLateralPositions As Integer

    ''' <summary>
    ''' Span direction used by the layout. Surfaced in the generation
    ''' summary so the user can confirm their orientation selection.
    ''' Forwarded from <see cref="BearingBarLayoutResult.SpanDirection"/>.
    ''' </summary>
    Public Property SpanDirection As SpanDirectionType

    ' --- Computed helpers ---

    Public ReadOnly Property TotalRequested As Integer
        Get
            Return If(Files IsNot Nothing, Files.Count, 0)
        End Get
    End Property

    Public ReadOnly Property TotalSaved As Integer
        Get
            If Files Is Nothing Then Return 0
            Dim count As Integer = 0
            For Each f In Files
                If f.Saved Then count += 1
            Next
            Return count
        End Get
    End Property

    Public ReadOnly Property TotalFailed As Integer
        Get
            Return TotalRequested - TotalSaved
        End Get
    End Property

    Public ReadOnly Property TotalNotches As Integer
        Get
            If Files Is Nothing Then Return 0
            Dim count As Integer = 0
            For Each f In Files
                If f.Saved Then count += f.NotchCount
            Next
            Return count
        End Get
    End Property

    ' --- Factory helpers ---

    Public Shared Function Succeeded(files As List(Of GeneratedBearingBarFile),
                                     outputFolder As String,
                                     warnings As List(Of String),
                                     Optional uniqueLateralPositions As Integer = 0,
                                     Optional spanDirection As SpanDirectionType = SpanDirectionType.AlongY) As BearingBarGenerationResult
        Return New BearingBarGenerationResult With {
            .Success = True,
            .Files = files,
            .OutputFolder = outputFolder,
            .Warnings = If(warnings, New List(Of String)),
            .UniqueLateralPositions = uniqueLateralPositions,
            .SpanDirection = spanDirection
        }
    End Function

    Public Shared Function Failed(message As String) As BearingBarGenerationResult
        Return New BearingBarGenerationResult With {
            .Success = False,
            .ErrorMessage = message,
            .Files = New List(Of GeneratedBearingBarFile),
            .Warnings = New List(Of String)
        }
    End Function

    ''' <summary>
    ''' Human-readable label for the span direction, matching the
    ''' parameter-panel combo wording (v1.5.7+).
    ''' </summary>
    Private Function SpanDirectionLabel() As String
        Select Case SpanDirection
            Case SpanDirectionType.AlongY
                Return "Vertical (along Y)"
            Case SpanDirectionType.AlongX
                Return "Horizontal (along X)"
            Case Else
                Return SpanDirection.ToString()
        End Select
    End Function

    ''' <summary>Builds a multi-line summary string for UI/logging.</summary>
    Public Function ToSummary() As String
        If Not Success Then Return "Part generation failed: " & ErrorMessage

        Dim sb As New System.Text.StringBuilder()
        sb.AppendLine("Bearing Bars")
        sb.AppendLine("———————————————————————————————————")
        sb.AppendLine("Span direction : " & SpanDirectionLabel())
        ' Show unique scan positions (matches the design-drawing
        ' convention the user sees in CAD references) alongside the
        ' fabrication-piece total (= count of .ipt files).
        If UniqueLateralPositions > 0 AndAlso
           UniqueLateralPositions <> TotalSaved Then
            sb.AppendLine("Bars (positions): " & UniqueLateralPositions)
            sb.AppendLine("Fabrication pieces: " & TotalSaved & " / " & TotalRequested)
        Else
            sb.AppendLine("Total bars     : " & TotalSaved & " / " & TotalRequested)
        End If
        sb.AppendLine("Total notches  : " & TotalNotches)

        If Warnings.Count > 0 Then
            sb.AppendLine()
            For Each w As String In Warnings
                sb.AppendLine("  • " & w)
            Next
        End If

        Return sb.ToString()
    End Function

End Class
