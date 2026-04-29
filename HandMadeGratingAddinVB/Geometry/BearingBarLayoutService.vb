'////////////////////////////////////////////////////////////////////
' Metal Bar Grating Addin - VB.NET
'
' BearingBarLayoutService: Core bearing bar layout engine.
' Generates trimmed bearing bar centerlines by scanning across
' the perimeter polygon at on-center spacing intervals.
'
' Supports: rectangles, skewed rectangles, simple cutouts.
' Uses scan-line / polygon-edge intersection clipping.
'////////////////////////////////////////////////////////////////////

Imports System.Diagnostics

''' <summary>
''' Generates bearing bar layout from a validated perimeter and parameters.
''' Pure geometry — no Inventor API dependency, no UI.
''' </summary>
Public Class BearingBarLayoutService

    Private Const Tolerance As Double = 0.0001 ' inches

    ''' <summary>
    ''' Minimum allowed edge-to-edge gap between two bars, in inches.
    ''' Gaps below this value cannot be reliably hot-dip galvanized
    ''' (zinc bridges across the gap and traps debris).  Any bar that
    ''' would land within this distance of an adjacent bar is dropped
    ''' from the layout — see <see cref="ApplyMinGalvanizeGap"/>.
    ''' </summary>
    Private Const MinGalvanizeGap As Double = 0.25 ' inches

    ''' <summary>
    ''' Generates the bearing bar layout.
    ''' </summary>
    Public Function Generate(perimeter As PerimeterData,
                             params As GratingParameters) As BearingBarLayoutResult
        Try
            Dim vertices As List(Of Double()) = perimeter.OuterLoopVertices
            If vertices Is Nothing OrElse vertices.Count < 3 Then
                Return BearingBarLayoutResult.Failed(
                    "Perimeter must have at least 3 vertices (found " &
                    If(vertices Is Nothing, "0", vertices.Count.ToString()) & ").")
            End If

            Trace.TraceInformation(": HMG Layout: Starting layout — " &
                vertices.Count & " vertices, direction=" &
                params.SpanDirection.ToString() &
                ", spacing=" & params.OnCenterSpacing)

            ' Ensure polygon is closed (duplicate first vertex at end if needed)
            Dim poly As List(Of Double()) = EnsureClosed(vertices)

            ' Compute bounding box
            Dim boundsMin As Double() = Nothing
            Dim boundsMax As Double() = Nothing
            ComputeBounds(poly, boundsMin, boundsMax)

            Trace.TraceInformation(": HMG Layout: Bounds min=(" &
                boundsMin(0).ToString("F4") & ", " & boundsMin(1).ToString("F4") &
                ") max=(" &
                boundsMax(0).ToString("F4") & ", " & boundsMax(1).ToString("F4") & ")")

            ' Determine axes
            ' lateralAxis: the axis we step across (perpendicular to span)
            ' spanAxis: the axis bars run along
            Dim lateralIdx As Integer ' 0=X, 1=Y
            If params.SpanDirection = SpanDirectionType.AlongX Then
                lateralIdx = 1 ' step in Y, bars run in X
            Else
                lateralIdx = 0 ' step in X, bars run in Y
            End If
            Dim spanIdx As Integer = 1 - lateralIdx

            Dim latMin As Double = boundsMin(lateralIdx)
            Dim latMax As Double = boundsMax(lateralIdx)
            Dim latSpan As Double = latMax - latMin

            If latSpan < Tolerance Then
                Return BearingBarLayoutResult.Failed(
                    "Perimeter has zero extent in the lateral direction.")
            End If

            ' Warnings accumulator (declared here so the galvanize-gap
            ' filter below can append before scan/clip starts).
            Dim warnings As New List(Of String)

            ' Generate scan positions centered within the perimeter bounds
            Dim positions As List(Of Double)

            If params.Banding = BandingOptionType.Banded Then
                ' Banded mode: band bars sit at latMin and latMax, replacing
                ' the first and last bearing bars. The remaining bearing bars
                ' are centered between the two band bars so that the edge gap
                ' (band bar to nearest bearing bar) is equal on both sides.
                ' The on-center spacing between bearing bars is preserved.
                positions = GenerateBandedScanPositions(
                    latMin, latMax, params.OnCenterSpacing)
            Else
                positions = GenerateScanPositions(
                    latMin, latMax, params.OnCenterSpacing)
            End If

            ' Galvanize gap rule: drop any bearing bar that would sit within
            ' MinGalvanizeGap of an adjacent bar (or band bar in banded
            ' mode).  See ApplyMinGalvanizeGap for details.
            Dim positionsBefore As Integer = positions.Count
            positions = ApplyMinGalvanizeGap(
                positions, params.BarWidth, lateralIdx, spanIdx,
                poly, params.Banding)
            If positions.Count <> positionsBefore Then
                warnings.Add(
                    (positionsBefore - positions.Count).ToString() &
                    " bar(s) removed — galvanize gap rule (< " &
                    MinGalvanizeGap.ToString("F2") & """ between bars).")
            End If

            Trace.TraceInformation(": HMG Layout: " & positions.Count &
                " scan positions from " &
                If(positions.Count > 0, positions(0).ToString("F4"), "?") &
                " to " &
                If(positions.Count > 0, positions(positions.Count - 1).ToString("F4"), "?"))

            ' Build polygon edge list
            Dim edges As List(Of Double()()) = BuildEdges(poly)

            ' Scan and clip
            Dim bars As New List(Of TrimmedBearingBar)
            Dim barIndex As Integer = 0

            For Each latPos As Double In positions
                Dim intersections As List(Of ScanIntersection) =
                    FindScanIntersections(edges, lateralIdx, spanIdx, latPos,
                                          params.BarWidth)

                If intersections.Count < 2 Then
                    warnings.Add("Scan at lateral=" & latPos.ToString("F4") &
                                 " found " & intersections.Count &
                                 " intersection(s) — skipped.")
                    Continue For
                End If

                ' Sort along the span axis
                intersections.Sort(
                    Function(a, b) a.SpanHit.CompareTo(b.SpanHit))

                ' When banded, inset every entry/exit intersection by the
                ' per-edge band-bar inset so bearing bars stop at the
                ' inner face of the adjacent band bar.  The inset depends
                ' on the angle between the perimeter edge and the span
                ' axis: for an edge perpendicular to the span axis the
                ' inset equals BarWidth, but for slanted edges (and edges
                ' approaching parallel to the span axis) the along-span
                ' inset grows by 1/sin(theta).  Each intersection carries
                ' its own EdgeInset so notch walls, slanted boundary
                ' edges, and rectangular sides all cut back correctly.
                If params.Banding = BandingOptionType.Banded AndAlso
                   intersections.Count >= 2 Then
                    For ix As Integer = 0 To intersections.Count - 1
                        Dim hit As ScanIntersection = intersections(ix)
                        Dim adjusted As Double
                        If ix Mod 2 = 0 Then
                            ' Entry — shift forward along span
                            adjusted = hit.SpanHit + hit.EdgeInset
                        Else
                            ' Exit — shift backward along span
                            adjusted = hit.SpanHit - hit.EdgeInset
                        End If
                        intersections(ix) = New ScanIntersection(
                            adjusted, hit.EdgeInset)
                    Next
                End If

                ' Pair entry/exit (even index = entry, odd = exit)
                For p As Integer = 0 To intersections.Count - 2 Step 2
                    Dim spanStart As Double = intersections(p).SpanHit
                    Dim spanEnd As Double = intersections(p + 1).SpanHit
                    Dim barLength As Double = spanEnd - spanStart

                    If barLength < Tolerance Then
                        warnings.Add("Zero-length bar at lateral=" &
                                     latPos.ToString("F4") & " — skipped.")
                        Continue For
                    End If

                    barIndex += 1

                    Dim startPt As Double()
                    Dim endPt As Double()

                    If params.SpanDirection = SpanDirectionType.AlongX Then
                        startPt = New Double() {spanStart, latPos}
                        endPt = New Double() {spanEnd, latPos}
                    Else
                        startPt = New Double() {latPos, spanStart}
                        endPt = New Double() {latPos, spanEnd}
                    End If

                    Dim bar As New TrimmedBearingBar()
                    bar.BarIndex = barIndex
                    bar.Mark = params.ResolvedPrefix & "-BB-" & barIndex
                    bar.StartPoint = startPt
                    bar.EndPoint = endPt
                    bar.Length = barLength
                    bar.SpanDirection = params.SpanDirection
                    bar.LateralPosition = latPos

                    bars.Add(bar)
                Next
            Next

            If bars.Count = 0 Then
                Return BearingBarLayoutResult.Failed(
                    "No bearing bars could be generated. " &
                    "The scan lines did not intersect the perimeter. " &
                    "Check span direction and spacing.")
            End If

            Trace.TraceInformation(": HMG Layout: Generated " & bars.Count &
                " bearing bars. First: " & bars(0).ToString())

            Return BearingBarLayoutResult.Succeeded(
                bars, params.SpanDirection, params.OnCenterSpacing,
                boundsMin, boundsMax, warnings)

        Catch ex As Exception
            Trace.TraceError(": HMG Layout: Unexpected error — " & ex.ToString())
            Return BearingBarLayoutResult.Failed(
                "Layout engine error: " & ex.Message)
        End Try
    End Function

#Region "Scan position generation"

    ''' <summary>
    ''' Generates evenly-spaced scan positions centered within [latMin, latMax].
    ''' First bar is placed at half-spacing inset from the boundary, then at
    ''' on-center intervals until the opposite boundary is reached.
    ''' </summary>
    Private Function GenerateScanPositions(latMin As Double,
                                           latMax As Double,
                                           spacing As Double) As List(Of Double)
        Dim positions As New List(Of Double)

        ' Inset the first bar by half the spacing from the minimum edge
        Dim first As Double = latMin + (spacing / 2.0)

        Dim pos As Double = first
        Do While pos < latMax - Tolerance
            positions.Add(pos)
            pos += spacing
        Loop

        Return positions
    End Function

    ''' <summary>
    ''' Generates bearing bar positions for banded grating.
    '''
    ''' Band bars sit at the perimeter edges (latMin and latMax),
    ''' replacing the outermost bearing bars. Bearing bars are placed
    ''' at exactly one on-center spacing from the near band bar and
    ''' continue at OC intervals. The spacing is adjusted slightly so
    ''' that bars are evenly distributed between both band bars.
    ''' </summary>
    Private Function GenerateBandedScanPositions(latMin As Double,
                                                  latMax As Double,
                                                  spacing As Double) As List(Of Double)
        Dim positions As New List(Of Double)
        Dim totalSpan As Double = latMax - latMin

        If totalSpan < spacing * 2 Then
            ' Span is too narrow for any bearing bars between the band bars.
            Trace.TraceInformation(
                ": HMG Layout: Banded — span (" &
                totalSpan.ToString("F4") &
                """) too narrow for bearing bars between band bars.")
            Return positions
        End If

        ' Determine the number of bearing bars between the two band bars.
        ' We want N bars such that (N+1) even spaces fill the total span.
        ' N = round(totalSpan/spacing) - 1  gives the closest integer
        ' count to the requested OC.
        Dim N As Integer = CInt(Math.Round(totalSpan / spacing)) - 1
        If N < 1 Then N = 1

        ' Compute the actual spacing so all gaps (including edge gaps
        ' from band bar to nearest bearing bar) are exactly equal.
        Dim actualSpacing As Double = totalSpan / CDbl(N + 1)

        Trace.TraceInformation(
            ": HMG Layout: Banded — " & N & " bearing bars, " &
            "requestedOC=" & spacing.ToString("F4") &
            """, actualOC=" & actualSpacing.ToString("F4") &
            """, span=" & totalSpan.ToString("F4") & """")

        For i As Integer = 1 To N
            positions.Add(latMin + CDbl(i) * actualSpacing)
        Next

        Return positions
    End Function

    ''' <summary>
    ''' Removes any bearing bar position whose edge-to-edge gap to a
    ''' neighbor (another bearing bar, or — in banded mode — a band bar
    ''' along a perimeter edge parallel to the span axis) would be less
    ''' than <see cref="MinGalvanizeGap"/>.
    '''
    ''' Hot-dip galvanizing requires every gap between adjacent bars to
    ''' be at least 0.25"; smaller gaps trap zinc and produce defective
    ''' parts.  Common scenarios where this rule fires:
    '''   • Banded mode + a notch/cutout whose walls are parallel to the
    '''     bearing bars — the wall becomes a band bar that lands very
    '''     close to a regular bearing bar at a standard scan position.
    '''   • Banded mode with a tight totalSpan / N ratio.
    '''
    ''' Band bars are treated as fixed (cannot be removed) since they
    ''' are structural perimeter framing.  Bearing bars are removed
    ''' wherever they would conflict.
    ''' </summary>
    Private Function ApplyMinGalvanizeGap(
            positions As List(Of Double),
            barWidth As Double,
            lateralIdx As Integer,
            spanIdx As Integer,
            polygon As List(Of Double()),
            banding As BandingOptionType) As List(Of Double)

        If positions Is Nothing OrElse positions.Count = 0 Then
            Return positions
        End If

        Dim minCenterDist As Double = barWidth + MinGalvanizeGap

        ' Build a candidate list including band bars (in banded mode),
        ' tagged so band bars cannot be removed.
        Dim candidates As New List(Of Candidate)
        If banding = BandingOptionType.Banded AndAlso polygon IsNot Nothing Then
            For Each bp As Double In FindParallelEdgeBandBarPositions(
                    polygon, lateralIdx, spanIdx, barWidth)
                candidates.Add(New Candidate(bp, True))
            Next
        End If
        For Each p As Double In positions
            candidates.Add(New Candidate(p, False))
        Next

        ' Stable-sort by lateral position
        candidates.Sort(Function(a, b) a.Position.CompareTo(b.Position))

        ' Single linear walk — for each adjacent pair within minCenterDist,
        ' remove one.  Band bars cannot be removed; if both candidates in
        ' a conflicting pair are bearing bars, drop the latter (so the
        ' kept bar stays closer to its neighbor on the other side).
        Dim removed As New HashSet(Of Integer)
        Dim lastKept As Integer = 0
        For i As Integer = 1 To candidates.Count - 1
            Dim prev As Candidate = candidates(lastKept)
            Dim cur As Candidate = candidates(i)
            Dim centerDist As Double = cur.Position - prev.Position

            If centerDist < minCenterDist - Tolerance Then
                If Not cur.IsBandBar Then
                    ' Drop the current bearing bar.
                    removed.Add(i)
                ElseIf Not prev.IsBandBar Then
                    ' Current is a band bar (immovable); drop the previous
                    ' bearing bar and replace lastKept with the band bar.
                    removed.Add(lastKept)
                    lastKept = i
                Else
                    ' Both are band bars — geometric edge case; keep both.
                    lastKept = i
                End If
            Else
                lastKept = i
            End If
        Next

        ' Emit only the surviving bearing-bar positions, in order.
        Dim result As New List(Of Double)
        For i As Integer = 0 To candidates.Count - 1
            If removed.Contains(i) Then Continue For
            If Not candidates(i).IsBandBar Then
                result.Add(candidates(i).Position)
            End If
        Next

        If removed.Count > 0 Then
            Trace.TraceInformation(
                ": HMG Layout: Galvanize-gap rule removed " &
                removed.Count.ToString() & " bar(s); " &
                result.Count.ToString() & " bearing bar(s) remain.")
        End If

        Return result
    End Function

    ''' <summary>
    ''' Internal record used by <see cref="ApplyMinGalvanizeGap"/> to track
    ''' a candidate bar's lateral position and whether it is a band bar
    ''' (band bars are structural and may not be removed).
    ''' </summary>
    Private Structure Candidate
        Public ReadOnly Position As Double
        Public ReadOnly IsBandBar As Boolean
        Public Sub New(pos As Double, bandBar As Boolean)
            Position = pos
            IsBandBar = bandBar
        End Sub
    End Structure

    ''' <summary>
    ''' Walks the polygon edges and returns the centerline lateral
    ''' coordinate of every band bar that runs parallel to the bearing
    ''' bars (i.e. every perimeter edge whose lateral coordinate is
    ''' constant).  Each such edge becomes a band bar in banded mode,
    ''' and that band bar can conflict with adjacent bearing bars.
    '''
    ''' For an outer rectangular perimeter this returns the latMin and
    ''' latMax positions.  For a perimeter with a notch whose walls are
    ''' parallel to the bearing bars, it additionally returns the lateral
    ''' coordinate of each notch wall — exactly the case where bearing
    ''' bars at standard scan positions can land within 0.25" of the
    ''' notch-wall band bar.
    '''
    ''' The returned position is the band bar's centerline, offset half a
    ''' bar width from the perimeter edge into the polygon material so
    ''' the gap calculation against adjacent bearing bars matches the
    ''' physical part placement.  Inside-direction is determined by the
    ''' polygon's signed area (CCW vs CW winding).
    ''' </summary>
    Private Function FindParallelEdgeBandBarPositions(
            polygon As List(Of Double()),
            lateralIdx As Integer,
            spanIdx As Integer,
            barWidth As Double) As List(Of Double)

        Dim result As New List(Of Double)
        If polygon Is Nothing OrElse polygon.Count < 2 Then Return result

        ' Compute signed area to determine winding (CCW => positive).
        Dim signedArea As Double = 0.0
        For i As Integer = 0 To polygon.Count - 2
            Dim a = polygon(i)
            Dim b = polygon(i + 1)
            signedArea += a(0) * b(1) - b(0) * a(1)
        Next
        Dim isCcw As Boolean = signedArea > 0.0

        Dim half As Double = barWidth / 2.0

        For i As Integer = 0 To polygon.Count - 2
            Dim a = polygon(i)
            Dim b = polygon(i + 1)
            Dim dLat As Double = b(lateralIdx) - a(lateralIdx)

            ' Edge is parallel to span axis when its lateral coord is constant.
            If Math.Abs(dLat) >= Tolerance Then Continue For

            Dim dSpan As Double = b(spanIdx) - a(spanIdx)
            If Math.Abs(dSpan) < Tolerance Then Continue For ' degenerate

            ' For a CCW polygon, "inside" is to the LEFT of the walking
            ' direction.  Rotating the edge direction by +90° gives the
            ' inside-pointing perpendicular.  In (X, Y) the rotation is
            '   (dx, dy) -> (-dy, dx).
            ' The sign of the lateral component depends on which axis
            ' is the lateral one:
            '   latIdx = 0 (X is lateral, Y is span):
            '     dx = 0, dy = dSpan, perp = (-dSpan, 0).
            '     lateral component = -sign(dSpan).
            '   latIdx = 1 (Y is lateral, X is span):
            '     dx = dSpan, dy = 0, perp = (0, dSpan).
            '     lateral component = +sign(dSpan).
            ' For CW polygons the sign is flipped.
            Dim insideSign As Integer
            If lateralIdx = 0 Then
                insideSign = If(dSpan > 0, -1, 1)
            Else
                insideSign = If(dSpan > 0, 1, -1)
            End If
            If Not isCcw Then insideSign = -insideSign

            Dim edgeLat As Double = a(lateralIdx)
            result.Add(edgeLat + insideSign * half)
        Next

        Return result
    End Function

#End Region

#Region "Polygon helpers"

    ''' <summary>
    ''' Ensures the vertex list forms a closed loop by appending the first
    ''' vertex if it does not already match the last.
    ''' </summary>
    Private Function EnsureClosed(vertices As List(Of Double())) As List(Of Double())
        Dim result As New List(Of Double())(vertices)
        Dim first As Double() = result(0)
        Dim last As Double() = result(result.Count - 1)

        If Math.Abs(first(0) - last(0)) > Tolerance OrElse
           Math.Abs(first(1) - last(1)) > Tolerance Then
            result.Add(New Double() {first(0), first(1)})
        End If

        Return result
    End Function

    ''' <summary>
    ''' Builds the list of directed edges {startVertex, endVertex} from the
    ''' closed polygon vertex list.
    ''' </summary>
    Private Function BuildEdges(poly As List(Of Double())) As List(Of Double()())
        Dim edges As New List(Of Double()())
        For i As Integer = 0 To poly.Count - 2
            edges.Add(New Double()() {poly(i), poly(i + 1)})
        Next
        Return edges
    End Function

    ''' <summary>
    ''' Computes the axis-aligned bounding box of the polygon.
    ''' </summary>
    Private Sub ComputeBounds(poly As List(Of Double()),
                              ByRef boundsMin As Double(),
                              ByRef boundsMax As Double())
        Dim xMin As Double = Double.MaxValue
        Dim yMin As Double = Double.MaxValue
        Dim xMax As Double = Double.MinValue
        Dim yMax As Double = Double.MinValue

        For Each v As Double() In poly
            If v(0) < xMin Then xMin = v(0)
            If v(1) < yMin Then yMin = v(1)
            If v(0) > xMax Then xMax = v(0)
            If v(1) > yMax Then yMax = v(1)
        Next

        boundsMin = New Double() {xMin, yMin}
        boundsMax = New Double() {xMax, yMax}
    End Sub

#End Region

#Region "Scan-line / edge intersection"

    ''' <summary>
    ''' One scan-line / polygon-edge intersection.  Carries both the
    ''' span-axis coordinate of the hit and the per-edge along-span inset
    ''' that should be applied in banded mode so the bearing bar stops at
    ''' the inner face of the band bar covering this edge.
    ''' </summary>
    Private Structure ScanIntersection
        Public ReadOnly SpanHit As Double
        Public ReadOnly EdgeInset As Double
        Public Sub New(spanHit_ As Double, edgeInset_ As Double)
            SpanHit = spanHit_
            EdgeInset = edgeInset_
        End Sub
    End Structure

    ''' <summary>
    ''' Finds all intersection coordinates (along the span axis) where the
    ''' scan line at the given lateral position crosses polygon edges, and
    ''' computes the per-edge along-span inset to use in banded mode.
    '''
    ''' Inset derivation: a band bar sits inside the perimeter with its
    ''' outer face on the edge and width <paramref name="barWidth"/>
    ''' measured perpendicular to the edge.  For a bearing bar to terminate
    ''' exactly at the inner face of that band bar, we shift its endpoint
    ''' along the span axis by:
    '''
    '''     inset = barWidth * sqrt(dSpan² + dLat²) / |dLat|
    '''
    ''' which equals barWidth/sin(theta), where theta is the angle between
    ''' the edge and the span axis.  At theta = 90° (edge perpendicular to
    ''' span) the inset reduces to <paramref name="barWidth"/>; for
    ''' slanted edges the along-span inset is correspondingly larger.
    ''' </summary>
    ''' <param name="edges">Polygon edge list.</param>
    ''' <param name="latIdx">Index of the lateral axis (0=X, 1=Y).</param>
    ''' <param name="spanIdx">Index of the span axis (0=X, 1=Y).</param>
    ''' <param name="latPos">The scan coordinate on the lateral axis.</param>
    ''' <param name="barWidth">Bar width in inches (used for inset).</param>
    Private Function FindScanIntersections(edges As List(Of Double()()),
                                           latIdx As Integer,
                                           spanIdx As Integer,
                                           latPos As Double,
                                           barWidth As Double) As List(Of ScanIntersection)
        Dim hits As New List(Of ScanIntersection)

        For Each edge As Double()() In edges
            Dim a As Double() = edge(0)
            Dim b As Double() = edge(1)

            Dim aLat As Double = a(latIdx)
            Dim bLat As Double = b(latIdx)

            ' Skip edges that are entirely on one side of the scan line.
            ' Use half-open interval [min, max) to avoid double-counting
            ' the scan line at a shared vertex.
            Dim minLat As Double = Math.Min(aLat, bLat)
            Dim maxLat As Double = Math.Max(aLat, bLat)

            If latPos < minLat - Tolerance OrElse latPos >= maxLat - Tolerance Then
                Continue For
            End If

            ' Edge is parallel to scan line (degenerate; would have
            ' been filtered by the half-open range check above unless
            ' the edge has zero length).
            Dim dLat As Double = bLat - aLat
            If Math.Abs(dLat) < Tolerance Then Continue For

            Dim dSpan As Double = b(spanIdx) - a(spanIdx)

            ' Linear interpolation to find span coordinate at latPos
            Dim t As Double = (latPos - aLat) / dLat
            Dim spanHit As Double = a(spanIdx) + t * dSpan

            ' Per-edge along-span inset for banded-mode trimming.  Equals
            ' barWidth / |sin(theta)| where theta is angle between the
            ' edge and the span axis.  For edges perpendicular to span
            ' (|dLat| = edgeLen) the inset equals barWidth.
            Dim edgeLen As Double = Math.Sqrt(dSpan * dSpan + dLat * dLat)
            Dim edgeInset As Double = barWidth * edgeLen / Math.Abs(dLat)

            hits.Add(New ScanIntersection(spanHit, edgeInset))
        Next

        Return hits
    End Function

#End Region

End Class
