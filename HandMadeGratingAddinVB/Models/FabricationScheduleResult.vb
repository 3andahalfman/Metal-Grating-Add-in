'////////////////////////////////////////////////////////////////////
' Metal Bar Grating Addin - VB.NET
'
' FabricationScheduleResult: Aggregates all schedule data extracted
' from a completed grating project.  Contains cross bar, band bar,
' and bearing bar schedule tables ready for documentation/PDF export.
'
' Phase 19 — Fabrication schedule data layer.
'////////////////////////////////////////////////////////////////////

Imports System.Text

''' <summary>
''' Complete fabrication schedule data for a grating project.
''' Produced by <see cref="FabricationScheduleBuilder"/> after all
''' generation phases have completed.
''' </summary>
Public Class FabricationScheduleResult

    ''' <summary>True if the schedule was built successfully.</summary>
    Public Property Success As Boolean

    ''' <summary>Human-readable error when Success is False.</summary>
    Public Property ErrorMessage As String

    ''' <summary>Cross bar schedule rows grouped by length.</summary>
    Public Property CrossBarSchedule As List(Of ScheduleRow)

    ''' <summary>Band bar schedule rows (one per segment or grouped by length).</summary>
    Public Property BandBarSchedule As List(Of ScheduleRow)

    ''' <summary>
    ''' Bearing bar schedule rows grouped by length.
    ''' Future-ready for documentation use.
    ''' </summary>
    Public Property BearingBarSchedule As List(Of ScheduleRow)

    ''' <summary>Non-fatal warnings encountered during schedule extraction.</summary>
    Public Property Warnings As List(Of String)

    ''' <summary>Project name for labeling.</summary>
    Public Property ProjectName As String

    ''' <summary>Cross bar type description (e.g. "3/8"" Round Plain").</summary>
    Public Property CrossBarTypeDescription As String

    ''' <summary>UTC timestamp when the schedule was generated.</summary>
    Public Property GeneratedUtc As DateTime

    ' --- Computed helpers ---

    ''' <summary>Total number of unique cross bar length groups.</summary>
    Public ReadOnly Property UniqueCrossBarLengths As Integer
        Get
            Return If(CrossBarSchedule IsNot Nothing, CrossBarSchedule.Count, 0)
        End Get
    End Property

    ''' <summary>Total quantity of cross bars across all groups.</summary>
    Public ReadOnly Property TotalCrossBarQuantity As Integer
        Get
            If CrossBarSchedule Is Nothing Then Return 0
            Dim total As Integer = 0
            For Each row In CrossBarSchedule
                total += row.Quantity
            Next
            Return total
        End Get
    End Property

    ''' <summary>Total number of band bar schedule entries.</summary>
    Public ReadOnly Property TotalBandBarEntries As Integer
        Get
            Return If(BandBarSchedule IsNot Nothing, BandBarSchedule.Count, 0)
        End Get
    End Property

    ''' <summary>Total quantity of band bars across all groups.</summary>
    Public ReadOnly Property TotalBandBarQuantity As Integer
        Get
            If BandBarSchedule Is Nothing Then Return 0
            Dim total As Integer = 0
            For Each row In BandBarSchedule
                total += row.Quantity
            Next
            Return total
        End Get
    End Property

    ''' <summary>Total combined band bar length in inches.</summary>
    Public ReadOnly Property TotalBandBarLength As Double
        Get
            If BandBarSchedule Is Nothing Then Return 0
            Dim sum As Double = 0
            For Each row In BandBarSchedule
                sum += row.Length * row.Quantity
            Next
            Return sum
        End Get
    End Property

    ''' <summary>Total number of unique bearing bar length groups.</summary>
    Public ReadOnly Property UniqueBearingBarLengths As Integer
        Get
            Return If(BearingBarSchedule IsNot Nothing, BearingBarSchedule.Count, 0)
        End Get
    End Property

    ''' <summary>Total quantity of bearing bars across all groups.</summary>
    Public ReadOnly Property TotalBearingBarQuantity As Integer
        Get
            If BearingBarSchedule Is Nothing Then Return 0
            Dim total As Integer = 0
            For Each row In BearingBarSchedule
                total += row.Quantity
            Next
            Return total
        End Get
    End Property

    ' --- Factory helpers ---

    Public Shared Function Succeeded(
            crossBars As List(Of ScheduleRow),
            bandBars As List(Of ScheduleRow),
            bearingBars As List(Of ScheduleRow),
            projectName As String,
            crossBarTypeDesc As String,
            warnings As List(Of String)) As FabricationScheduleResult

        Return New FabricationScheduleResult With {
            .Success = True,
            .CrossBarSchedule = If(crossBars, New List(Of ScheduleRow)),
            .BandBarSchedule = If(bandBars, New List(Of ScheduleRow)),
            .BearingBarSchedule = If(bearingBars, New List(Of ScheduleRow)),
            .ProjectName = projectName,
            .CrossBarTypeDescription = crossBarTypeDesc,
            .Warnings = If(warnings, New List(Of String)),
            .GeneratedUtc = DateTime.UtcNow
        }
    End Function

    Public Shared Function Failed(message As String) As FabricationScheduleResult
        Return New FabricationScheduleResult With {
            .Success = False,
            .ErrorMessage = message,
            .CrossBarSchedule = New List(Of ScheduleRow),
            .BandBarSchedule = New List(Of ScheduleRow),
            .BearingBarSchedule = New List(Of ScheduleRow),
            .Warnings = New List(Of String),
            .GeneratedUtc = DateTime.UtcNow
        }
    End Function

    ''' <summary>
    ''' Builds a multi-line summary string for UI/logging.
    ''' </summary>
    Public Function ToSummary() As String
        If Not Success Then
            Return "Fabrication Schedule: FAILED — " & ErrorMessage
        End If

        Dim sb As New StringBuilder()
        sb.AppendLine("Fabrication Schedule")
        sb.AppendLine("—————————————————————————————————")

        ' Cross bars
        sb.AppendLine("Cross bar type : " & If(CrossBarTypeDescription, "N/A"))
        sb.AppendLine("Cross bar groups : " & UniqueCrossBarLengths)
        sb.AppendLine("Cross bars total : " & TotalCrossBarQuantity)

        ' Band bars
        If TotalBandBarEntries > 0 Then
            sb.AppendLine("Band bar groups : " & TotalBandBarEntries)
            sb.AppendLine("Band bars total : " & TotalBandBarQuantity)
            sb.AppendLine("Band bar length : " & TotalBandBarLength.ToString("F4") & """")
        Else
            sb.AppendLine("Band bars : None (open ended)")
        End If

        ' Bearing bars
        sb.AppendLine("Bearing bar groups : " & UniqueBearingBarLengths)
        sb.AppendLine("Bearing bars total : " & TotalBearingBarQuantity)

        ' Warnings
        If Warnings IsNot Nothing AndAlso Warnings.Count > 0 Then
            sb.AppendLine()
            For Each w In Warnings
                sb.AppendLine("• " & w)
            Next
        End If

        Return sb.ToString()
    End Function

End Class
