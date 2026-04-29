'////////////////////////////////////////////////////////////////////
' Metal Bar Grating Addin - VB.NET
'
' FabricationScheduleBuilder: Extracts structured schedule/table
' data from the generated grating project results.
'
' Produces grouped schedule rows for:
'   - Cross bars  (grouped by length, with quantity and mark range)
'   - Band bars   (grouped by length, with quantity)
'   - Bearing bars (grouped by length, with quantity — future-ready)
'
' The output is a FabricationScheduleResult that can be consumed by
' a later drawing/PDF export phase without re-scanning geometry.
'
' Phase 19 — Fabrication schedule data layer.
'////////////////////////////////////////////////////////////////////

Imports System.Diagnostics

''' <summary>
''' Builds fabrication schedule data from completed generation results.
''' Call <see cref="Build"/> after all generation phases have finished.
''' </summary>
Public Class FabricationScheduleBuilder

    ''' <summary>Length tolerance for grouping — bars within this
    ''' tolerance (inches) are considered the same length.</summary>
    Private Const LengthTolerance As Double = 0.0005

    ''' <summary>Number of decimal places for rounding lengths when
    ''' building the group key.</summary>
    Private Const RoundingDecimals As Integer = 4

    ''' <summary>
    ''' Builds the complete fabrication schedule from a finished project.
    ''' </summary>
    ''' <param name="project">
    ''' A GratingProject with completed GenerationResult, CrossBarResult,
    ''' and BandBarResult.
    ''' </param>
    ''' <returns>A FabricationScheduleResult with all schedule tables.</returns>
    Public Function Build(project As GratingProject) As FabricationScheduleResult
        If project Is Nothing Then
            Return FabricationScheduleResult.Failed("Project is Nothing.")
        End If

        Dim warnings As New List(Of String)

        Try
            ' --- Cross bar schedule ---
            Dim crossBarRows As List(Of ScheduleRow) =
                BuildCrossBarSchedule(project, warnings)

            ' --- Band bar schedule ---
            Dim bandBarRows As List(Of ScheduleRow) =
                BuildBandBarSchedule(project, warnings)

            ' --- Bearing bar schedule ---
            Dim bearingBarRows As List(Of ScheduleRow) =
                BuildBearingBarSchedule(project, warnings)

            ' --- Cross bar type description ---
            Dim cbTypeDesc As String = Nothing
            If project.Parameters IsNot Nothing Then
                cbTypeDesc = CrossBarTypeHelper.GetDisplayName(
                    project.Parameters.CrossBar)
                If project.Parameters.SurfaceProfile = SurfaceProfileType.Serrated Then
                    cbTypeDesc &= " Serrated"
                Else
                    cbTypeDesc &= " Plain"
                End If
            End If

            ' --- Trace summary ---
            Trace.TraceInformation(
                ": HMG Phase 19: Schedule built — " &
                crossBarRows.Count & " cross bar group(s), " &
                bandBarRows.Count & " band bar group(s), " &
                bearingBarRows.Count & " bearing bar group(s).")

            Dim totalCB As Integer = 0
            For Each r In crossBarRows : totalCB += r.Quantity : Next
            Trace.TraceInformation(
                ": HMG Phase 19:   Cross bars — " &
                crossBarRows.Count & " unique length(s), " &
                totalCB & " total quantity.")

            Dim totalBB As Integer = 0
            For Each r In bandBarRows : totalBB += r.Quantity : Next
            Trace.TraceInformation(
                ": HMG Phase 19:   Band bars — " &
                bandBarRows.Count & " unique length(s), " &
                totalBB & " total quantity.")

            Dim totalBearing As Integer = 0
            For Each r In bearingBarRows : totalBearing += r.Quantity : Next
            Trace.TraceInformation(
                ": HMG Phase 19:   Bearing bars — " &
                bearingBarRows.Count & " unique length(s), " &
                totalBearing & " total quantity.")

            If warnings.Count > 0 Then
                For Each w In warnings
                    Trace.TraceWarning(": HMG Phase 19: Warning — " & w)
                Next
            End If

            Return FabricationScheduleResult.Succeeded(
                crossBarRows, bandBarRows, bearingBarRows,
                If(project.ProjectName, "Grating"),
                cbTypeDesc, warnings)

        Catch ex As Exception
            Trace.TraceError(
                ": HMG Phase 19: Schedule build failed — " & ex.Message)
            Return FabricationScheduleResult.Failed(
                "Schedule extraction failed: " & ex.Message)
        End Try
    End Function

    ' ==================================================================
    '  Cross bar schedule
    ' ==================================================================

    ''' <summary>
    ''' Builds cross bar schedule rows grouped by length.
    ''' Uses the generated CrossBarResult entries (actual computed lengths).
    ''' </summary>
    Private Function BuildCrossBarSchedule(
            project As GratingProject,
            warnings As List(Of String)) As List(Of ScheduleRow)

        Dim rows As New List(Of ScheduleRow)

        If project.CrossBarResult Is Nothing Then
            warnings.Add("Cross bar result is not available — schedule skipped.")
            Return rows
        End If

        If Not project.CrossBarResult.Success Then
            warnings.Add("Cross bar generation was not successful — schedule skipped.")
            Return rows
        End If

        Dim entries As List(Of CrossBarEntry) = project.CrossBarResult.Entries
        If entries Is Nothing OrElse entries.Count = 0 Then
            warnings.Add("No cross bar entries found.")
            Return rows
        End If

        ' Group entries by rounded length
        Dim groups As New SortedDictionary(Of Double, List(Of CrossBarEntry))

        For Each entry In entries
            Dim key As Double = Math.Round(entry.Length, RoundingDecimals)
            If Not groups.ContainsKey(key) Then
                groups(key) = New List(Of CrossBarEntry)
            End If
            groups(key).Add(entry)
        Next

        ' Build schedule rows from groups
        For Each kvp In groups
            Dim groupEntries As List(Of CrossBarEntry) = kvp.Value
            Dim marks As New List(Of String)
            For Each e In groupEntries
                marks.Add(If(e.Mark, "CB-" & e.Index.ToString("000")))
            Next

            Dim row As New ScheduleRow()
            row.ComponentType = ScheduleComponentType.CrossBar
            row.Mark = marks(0)
            row.Length = kvp.Key
            row.Quantity = groupEntries.Count
            row.IndividualMarks = marks
            row.AllSaved = True  ' Cross bars grouped by file; assume saved

            ' Type description from parameters
            If project.Parameters IsNot Nothing Then
                row.TypeDescription = CrossBarTypeHelper.GetDisplayName(
                    project.Parameters.CrossBar)
            End If

            rows.Add(row)
        Next

        Return rows
    End Function

    ' ==================================================================
    '  Band bar schedule
    ' ==================================================================

    ''' <summary>
    ''' Builds band bar schedule rows grouped by length.
    ''' Uses the generated BandBarResult files (actual segment lengths).
    ''' </summary>
    Private Function BuildBandBarSchedule(
            project As GratingProject,
            warnings As List(Of String)) As List(Of ScheduleRow)

        Dim rows As New List(Of ScheduleRow)

        If project.BandBarResult Is Nothing Then
            warnings.Add("Band bar result is not available — schedule skipped.")
            Return rows
        End If

        If project.BandBarResult.Skipped Then
            ' Open-ended — no band bars to schedule
            Return rows
        End If

        If Not project.BandBarResult.Success Then
            warnings.Add("Band bar generation was not successful — schedule skipped.")
            Return rows
        End If

        Dim files As List(Of GeneratedBandBarFile) = project.BandBarResult.Files
        If files Is Nothing OrElse files.Count = 0 Then
            Return rows
        End If

        ' Only include saved segments
        Dim savedFiles As New List(Of GeneratedBandBarFile)
        For Each f In files
            If f.Saved Then savedFiles.Add(f)
        Next

        If savedFiles.Count = 0 Then
            warnings.Add("No band bar segments were saved successfully.")
            Return rows
        End If

        ' Group by rounded length
        Dim groups As New SortedDictionary(Of Double, List(Of GeneratedBandBarFile))

        For Each f In savedFiles
            Dim key As Double = Math.Round(f.Length, RoundingDecimals)
            If Not groups.ContainsKey(key) Then
                groups(key) = New List(Of GeneratedBandBarFile)
            End If
            groups(key).Add(f)
        Next

        ' Build schedule rows from groups
        For Each kvp In groups
            Dim groupFiles As List(Of GeneratedBandBarFile) = kvp.Value
            Dim marks As New List(Of String)
            For Each f In groupFiles
                marks.Add(If(f.Mark, "BAND-" & f.SegmentIndex.ToString("000")))
            Next

            Dim row As New ScheduleRow()
            row.ComponentType = ScheduleComponentType.BandBar
            row.Mark = marks(0)
            row.Length = kvp.Key
            row.Quantity = groupFiles.Count
            row.IndividualMarks = marks
            row.AllSaved = True

            rows.Add(row)
        Next

        Return rows
    End Function

    ' ==================================================================
    '  Bearing bar schedule
    ' ==================================================================

    ''' <summary>
    ''' Builds bearing bar schedule rows grouped by length.
    ''' Uses the generated BearingBarGenerationResult files
    ''' (actual bar lengths from the layout engine).
    ''' </summary>
    Private Function BuildBearingBarSchedule(
            project As GratingProject,
            warnings As List(Of String)) As List(Of ScheduleRow)

        Dim rows As New List(Of ScheduleRow)

        If project.GenerationResult Is Nothing Then
            warnings.Add("Bearing bar generation result is not available — schedule skipped.")
            Return rows
        End If

        If Not project.GenerationResult.Success Then
            warnings.Add("Bearing bar generation was not successful — schedule skipped.")
            Return rows
        End If

        Dim files As List(Of GeneratedBearingBarFile) = project.GenerationResult.Files
        If files Is Nothing OrElse files.Count = 0 Then
            warnings.Add("No bearing bar files found.")
            Return rows
        End If

        ' Only include saved files
        Dim savedFiles As New List(Of GeneratedBearingBarFile)
        For Each f In files
            If f.Saved Then savedFiles.Add(f)
        Next

        If savedFiles.Count = 0 Then
            warnings.Add("No bearing bar files were saved successfully.")
            Return rows
        End If

        ' Group by rounded length
        Dim groups As New SortedDictionary(Of Double, List(Of GeneratedBearingBarFile))

        For Each f In savedFiles
            Dim barLength As Double = f.SourceBar.Length
            Dim key As Double = Math.Round(barLength, RoundingDecimals)
            If Not groups.ContainsKey(key) Then
                groups(key) = New List(Of GeneratedBearingBarFile)
            End If
            groups(key).Add(f)
        Next

        ' Build schedule rows from groups
        For Each kvp In groups
            Dim groupFiles As List(Of GeneratedBearingBarFile) = kvp.Value
            Dim marks As New List(Of String)
            For Each f In groupFiles
                marks.Add(If(f.SourceBar.Mark, "BB-" & f.SourceBar.BarIndex.ToString("000")))
            Next

            Dim row As New ScheduleRow()
            row.ComponentType = ScheduleComponentType.BearingBar
            row.Mark = marks(0)
            row.Length = kvp.Key
            row.Quantity = groupFiles.Count
            row.IndividualMarks = marks
            row.AllSaved = True

            rows.Add(row)
        Next

        Return rows
    End Function

End Class
