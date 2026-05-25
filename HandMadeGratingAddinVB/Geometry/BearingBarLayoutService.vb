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
    ''' Extra slack when deciding if a perimeter wall lies on the panel
    ''' bounding box (outer frame). DWG/import noise can leave edges
    ''' slightly off exact latMin/latMax; a too-tight tolerance mis-tagged
    ''' outer bands as notch walls, so the gap rule eliminated them and
    ''' band parts were never generated.
    ''' 0.02" (~0.5 mm) plus a small fraction of panel span covers typical
    ''' sketch/translation error without swallowing real interior notch walls.
    ''' </summary>
    Private Const OuterBBoxEpsilonInches As Double = 0.02

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

            ' Eliminated edge / cross bar trackers populated by the gap rule.
            Dim eliminatedEdges As New List(Of Double())
            Dim eliminatedCrossBarPositions As New List(Of Double)

            ' Generate scan positions.  Both banded and non-banded modes
            ' place the outermost bearing bars FLUSH with the panel lateral
            ' edges (outer face on latMin / latMax).  Band bars only exist
            ' at the SPAN ends (top/bottom of the panel) — the outer bearing
            ' bars themselves form the lateral side edges, so no separate
            ' band bar is needed at latMin or latMax.  This was the design
            ' established in v1.4.1 and is what the user expects.
            Dim positions As List(Of Double) = GenerateScanPositions(
                latMin, latMax, params.OnCenterSpacing, params.BarWidth)

            ' Galvanize gap rule: drop any bearing bar that would sit within
            ' MinGalvanizeGap of an adjacent bar (or band bar in banded
            ' mode).  Notch-wall band bars are themselves eliminated per the
            ' "band bar is eliminated and cut is enlarged to the next closest
            ' bearing bar" rule (see ApplyMinGalvanizeGap for details).
            Dim positionsBefore As Integer = positions.Count
            positions = ApplyMinGalvanizeGap(
                positions, params.BarWidth, lateralIdx, spanIdx,
                poly, params.Banding, latMin, latMax, eliminatedEdges)
            If positions.Count <> positionsBefore Then
                warnings.Add(
                    (positionsBefore - positions.Count).ToString() &
                    " bar(s) removed — galvanize gap rule (< " &
                    MinGalvanizeGap.ToString("F2") & """ between bars).")
            End If

            ' Inner notch-wall rule (per the user's PDF spec, v1.5.11):
            ' For each parallel-to-span notch wall, measure the lateral
            ' distance from the cut line (the wall's perimeter coord) to
            ' the nearest *kept* bearing bar centerline.  When that
            ' distance is < MinGalvanizeGap (1/4"), the wall band bar is
            ' eliminated; the next-closest bearing bar stands in as the
            ' wall.  When the distance is ≥ 1/4", the wall band bar is
            ' added (BandBarGenerator emits the segment).  This is
            ' simpler and more lenient than the symmetric edge-to-edge
            ' rule applied in ApplyMinGalvanizeGap to outer band bars.
            If params.Banding = BandingOptionType.Banded AndAlso
               poly IsNot Nothing Then
                Dim wallsEliminatedBefore As Integer = eliminatedEdges.Count
                ApplyInnerWallEliminationRule(
                    poly, positions, lateralIdx, params.BarWidth,
                    latMin, latMax, eliminatedEdges)
                Dim wallsEliminated As Integer =
                    eliminatedEdges.Count - wallsEliminatedBefore
                If wallsEliminated > 0 Then
                    warnings.Add(
                        wallsEliminated.ToString() &
                        " inner-notch-wall band bar(s) eliminated — " &
                        "cut line within " &
                        MinGalvanizeGap.ToString("F2") &
                        """ of a bearing bar.")
                End If
            End If

            ' Perpendicular-axis galvanize-gap rule: cross-bar vs perpendicular
            ' notch-wall band bar.  Applies the same "eliminate band bar and
            ' enlarge cut" rule symmetrically along the span axis.
            Dim eliminatedEdgesBefore As Integer = eliminatedEdges.Count
            If params.Banding = BandingOptionType.Banded AndAlso
               params.CrossBarOnCenter > 0 Then
                ApplyPerpendicularGap(poly, params, latMin, latMax,
                                      lateralIdx, spanIdx,
                                      eliminatedEdges,
                                      eliminatedCrossBarPositions)
                Dim perpEliminated As Integer =
                    eliminatedEdges.Count - eliminatedEdgesBefore
                If perpEliminated > 0 Then
                    warnings.Add(
                        perpEliminated.ToString() &
                        " perpendicular wall band bar(s) eliminated — " &
                        "galvanize gap rule (< " &
                        MinGalvanizeGap.ToString("F2") &
                        """ to cross bar).")
                End If
            End If

            Trace.TraceInformation(": HMG Layout: " & positions.Count &
                " scan positions from " &
                If(positions.Count > 0, positions(0).ToString("F4"), "?") &
                " to " &
                If(positions.Count > 0, positions(positions.Count - 1).ToString("F4"), "?"))

            ' Build polygon edge list
            Dim edges As List(Of Double()()) = BuildEdges(poly)

            ' Arc-aware trim setup: build an edge-index → inner-face-circle
            ' map so the scan-line clip lands on the curved band bar's inner
            ' face (radius R ± barWidth from arc centre) instead of on the
            ' tessellated chord that lies on the arc itself.  Without this
            ' override, bearing bars terminate ~on the arc and overlap the
            ' curved band bar's body (the chord-normal banded inset is only
            ' approximate and zero in non-banded mode).
            Dim arcEdgeMap As Dictionary(Of Integer, ArcEdgeContext) =
                BuildArcEdgeMap(perimeter, poly, params.BarWidth, edges.Count)

            ' Scan and clip
            Dim bars As New List(Of TrimmedBearingBar)
            Dim barIndex As Integer = 0

            For Each latPos As Double In positions
                Dim intersections As List(Of ScanIntersection) =
                    FindScanIntersections(edges, lateralIdx, spanIdx, latPos,
                                          params.BarWidth, arcEdgeMap)

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

                ' Pair entry/exit (even index = entry, odd = exit), then
                ' apply body-aware arc cuts so bars whose centerline
                ' misses an arc — but whose body crosses the inner-face
                ' circle — are split around the arc instead of running
                ' straight through it.
                For p As Integer = 0 To intersections.Count - 2 Step 2
                    Dim spanStart As Double = intersections(p).SpanHit
                    Dim spanEnd As Double = intersections(p + 1).SpanHit

                    Dim subSegments As List(Of Double()) =
                        ApplyArcBodyCuts(
                            spanStart, spanEnd, latPos, params.BarWidth,
                            arcEdgeMap, lateralIdx, spanIdx)

                    For Each subSeg As Double() In subSegments
                        Dim subStart As Double = subSeg(0)
                        Dim subEnd As Double = subSeg(1)
                        Dim subLength As Double = subEnd - subStart

                        If subLength < Tolerance Then
                            Continue For
                        End If

                        barIndex += 1

                        Dim startPt As Double()
                        Dim endPt As Double()

                        If params.SpanDirection = SpanDirectionType.AlongX Then
                            startPt = New Double() {subStart, latPos}
                            endPt = New Double() {subEnd, latPos}
                        Else
                            startPt = New Double() {latPos, subStart}
                            endPt = New Double() {latPos, subEnd}
                        End If

                        Dim bar As New TrimmedBearingBar()
                        bar.BarIndex = barIndex
                        bar.Mark = params.ResolvedPrefix & "-BB-" & barIndex
                        bar.StartPoint = startPt
                        bar.EndPoint = endPt
                        bar.Length = subLength
                        bar.SpanDirection = params.SpanDirection
                        bar.LateralPosition = latPos

                        bars.Add(bar)
                    Next
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
                boundsMin, boundsMax, warnings,
                eliminatedEdges, eliminatedCrossBarPositions,
                positions.Count)

        Catch ex As Exception
            Trace.TraceError(": HMG Layout: Unexpected error — " & ex.ToString())
            Return BearingBarLayoutResult.Failed(
                "Layout engine error: " & ex.Message)
        End Try
    End Function

#Region "Scan position generation"

    ''' <summary>
    ''' Generates bearing bar positions with the outermost bars FLUSH at
    ''' the panel lateral edges — i.e. the first bar's outer face sits
    ''' on <paramref name="latMin"/> and the last bar's outer face sits
    ''' on <paramref name="latMax"/>.  In banded mode the outer bearing
    ''' bars themselves form the lateral side edges of the panel; band
    ''' bars only exist at the SPAN ends (top/bottom).
    '''
    ''' Nominal count: N = round(innerSpan / spacing) + 1, where
    ''' innerSpan = latMax − latMin − barWidth (center-to-center distance
    ''' between the two flush outer bars).  N is reduced if any adjacent
    ''' edge-to-edge gap would fall below MinGalvanizeGap (¼").
    ''' </summary>
    Private Function GenerateScanPositions(latMin As Double,
                                           latMax As Double,
                                           spacing As Double,
                                           barWidth As Double) As List(Of Double)
        Dim positions As New List(Of Double)
        Dim half As Double = barWidth / 2.0

        ' Center-to-center distance between the two flush outer bars.
        Dim innerSpan As Double = latMax - latMin - barWidth

        If innerSpan < -Tolerance Then Return positions

        ' Only one bar fits (panel too narrow for two flush bars).
        If innerSpan < Tolerance Then
            positions.Add((latMin + latMax) / 2.0)
            Return positions
        End If

        Dim N As Integer = Math.Max(1, CInt(Math.Round(innerSpan / spacing)) + 1)

        ' Reduce N until every adjacent edge-to-edge gap meets MinGalvanizeGap.
        Do While N >= 2
            Dim actualSpacing As Double = innerSpan / CDbl(N - 1)
            If actualSpacing - barWidth >= MinGalvanizeGap - Tolerance Then Exit Do
            N -= 1
            Trace.TraceInformation(
                ": HMG Layout: Bar-to-bar gap too small, reducing to N=" & N)
        Loop

        Dim finalSpacing As Double = If(N > 1, innerSpan / CDbl(N - 1), 0.0)

        Trace.TraceInformation(
            ": HMG Layout: " & N & " bearing bars, " &
            "requestedOC=" & spacing.ToString("F4") &
            """, actualOC=" & finalSpacing.ToString("F4") &
            """, innerSpan=" & innerSpan.ToString("F4") & """")

        For i As Integer = 0 To N - 1
            positions.Add(latMin + half + CDbl(i) * finalSpacing)
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
    ''' parts.  Outer-perimeter band bars (along the bounding-box
    ''' latMin / latMax edges) are immovable — they are structural
    ''' framing of the panel.  Notch-wall band bars (parallel-to-span
    ''' edges anywhere else on the perimeter) follow the PDF rule:
    ''' when one would land within MinGalvanizeGap of a bearing bar,
    ''' the band bar is eliminated and the cut is treated as enlarged
    ''' to the next closest bearing bar — so the conflicting bearing
    ''' bar is also dropped, and the edge is recorded in
    ''' <paramref name="eliminatedEdges"/> so the band-bar generator
    ''' skips producing a part for that wall.
    ''' </summary>
    Private Function ApplyMinGalvanizeGap(
            positions As List(Of Double),
            barWidth As Double,
            lateralIdx As Integer,
            spanIdx As Integer,
            polygon As List(Of Double()),
            banding As BandingOptionType,
            latMin As Double,
            latMax As Double,
            eliminatedEdges As List(Of Double())) As List(Of Double)

        If positions Is Nothing OrElse positions.Count = 0 Then
            Return positions
        End If

        Dim minCenterDist As Double = barWidth + MinGalvanizeGap

        ' Build a candidate list.  Only OUTER-perimeter band bars are
        ' candidates here — they protect bearing bars and so participate
        ' in the BB-vs-band conflict resolution.  Inner notch-wall band
        ' bars are evaluated in a separate pass below using the user's
        ' rule (cut-line to nearest bearing-bar centerline < 1/4").
        Dim candidates As New List(Of Candidate)
        If banding = BandingOptionType.Banded AndAlso polygon IsNot Nothing Then
            Dim wallCandidates As List(Of WallBandBar) =
                FindParallelEdgeBandBarPositions(
                    polygon, lateralIdx, spanIdx, barWidth)
            For Each wb In wallCandidates
                Dim isOuter As Boolean =
                    IsOuterPerimeterBandBar(wb, lateralIdx, barWidth, latMin, latMax)
                If Not isOuter Then Continue For
                candidates.Add(New Candidate(wb.Position, True, isOuter, wb.Edge))
                Trace.TraceInformation(
                    ": HMG Layout: Outer wall band-bar candidate at lat=" &
                    wb.Position.ToString("F4") & ".")
            Next
        End If
        For Each p As Double In positions
            candidates.Add(New Candidate(p, False, False, Nothing))
        Next

        ' Stable-sort by lateral position
        candidates.Sort(Function(a, b) a.Position.CompareTo(b.Position))

        ' Single linear walk — for each adjacent pair within minCenterDist,
        ' decide what to drop.
        '   • Bearing bar vs bearing bar  -> drop the latter
        '   • Bearing bar vs outer band   -> drop the bearing bar
        '   • Bearing bar vs notch band   -> drop the bearing bar AND
        '                                    eliminate the notch band bar
        '   • Outer vs outer band         -> keep both (geometric edge)
        Dim removed As New HashSet(Of Integer)
        Dim eliminatedIdx As New HashSet(Of Integer)
        Dim lastKept As Integer = 0
        For i As Integer = 1 To candidates.Count - 1
            Dim prev As Candidate = candidates(lastKept)
            Dim cur As Candidate = candidates(i)
            Dim centerDist As Double = cur.Position - prev.Position

            If centerDist < minCenterDist - Tolerance Then
                ' Identify the bearing bar / band bar in this conflict pair.
                Dim bearingIdx As Integer = -1
                Dim bandIdx As Integer = -1
                If cur.IsBandBar AndAlso Not prev.IsBandBar Then
                    bearingIdx = lastKept : bandIdx = i
                ElseIf prev.IsBandBar AndAlso Not cur.IsBandBar Then
                    bearingIdx = i : bandIdx = lastKept
                End If

                If bearingIdx >= 0 AndAlso bandIdx >= 0 Then
                    ' Bearing-bar vs band-bar conflict.
                    If candidates(bandIdx).IsOuterBand Then
                        ' Outer-perimeter band bar (panel framing) is
                        ' immovable — drop the bearing bar.
                        removed.Add(bearingIdx)
                        lastKept = bandIdx
                    Else
                        ' Notch-wall band bar: eliminate the band bar,
                        ' KEEP the bearing bar.  The cut is treated as
                        ' enlarged to the bearing bar position, so the
                        ' bearing bar itself forms the new wall edge.
                        eliminatedIdx.Add(bandIdx)
                        lastKept = bearingIdx
                    End If
                ElseIf Not cur.IsBandBar Then
                    ' Bearing vs bearing — drop the latter.
                    removed.Add(i)
                Else
                    ' Band vs band (outer/outer geometric edge) — keep both.
                    lastKept = i
                End If
            Else
                lastKept = i
            End If
        Next

        ' Record eliminated notch-wall edges for the band bar generator.
        For Each idx In eliminatedIdx
            Dim wbEdge As Double() = candidates(idx).Edge
            If wbEdge IsNot Nothing Then eliminatedEdges.Add(wbEdge)
        Next

        ' Emit only the surviving bearing-bar positions, in order.
        Dim result As New List(Of Double)
        For i As Integer = 0 To candidates.Count - 1
            If removed.Contains(i) Then Continue For
            If Not candidates(i).IsBandBar Then
                result.Add(candidates(i).Position)
            End If
        Next

        If removed.Count > 0 OrElse eliminatedIdx.Count > 0 Then
            Trace.TraceInformation(
                ": HMG Layout: Galvanize-gap rule dropped " &
                removed.Count.ToString() & " bearing bar(s) and eliminated " &
                eliminatedIdx.Count.ToString() & " notch-wall band bar(s); " &
                result.Count.ToString() & " bearing bar(s) remain.")
        End If

        Return result
    End Function

    ''' <summary>
    ''' User's inner-notch-wall rule (PDF spec, v1.5.11+):
    ''' For each non-outer parallel-to-span perimeter edge (= an inner
    ''' notch wall), measure the lateral distance from the cut line
    ''' (the wall's lateral coord) to the nearest kept bearing-bar
    ''' centerline.  When that distance is &lt; MinGalvanizeGap (1/4"),
    ''' the wall's band bar is eliminated (added to <paramref name="eliminatedEdges"/>);
    ''' the bearing bar at that position takes the wall's role.
    ''' When the distance is ≥ 1/4", the band bar is kept — BandBarGenerator
    ''' will emit the segment.
    '''
    ''' This is a different measurement from the symmetric edge-to-edge
    ''' rule in <see cref="ApplyMinGalvanizeGap"/>: the cut-line measure
    ''' ignores the band bar's own thickness, so a wall and a flush-
    ''' adjacent bearing bar (centerline at wall + halfWidth) read as
    ''' ~halfWidth apart, well below the 1/4" threshold — band bar
    ''' eliminated.  A wall whose nearest bearing bar is a full
    ''' on-center step away reads as ~OC apart — well above 1/4", band
    ''' bar kept.
    ''' </summary>
    Private Sub ApplyInnerWallEliminationRule(
            polygon As List(Of Double()),
            scanPositions As List(Of Double),
            lateralIdx As Integer,
            barWidth As Double,
            latMin As Double,
            latMax As Double,
            eliminatedEdges As List(Of Double()))

        If polygon Is Nothing OrElse scanPositions Is Nothing Then Return
        If scanPositions.Count = 0 Then Return

        Dim spanIdx As Integer = 1 - lateralIdx

        Dim allWalls As List(Of WallBandBar) =
            FindParallelEdgeBandBarPositions(
                polygon, lateralIdx, spanIdx, barWidth)

        For Each wb In allWalls
            Dim isOuter As Boolean =
                IsOuterPerimeterBandBar(wb, lateralIdx, barWidth, latMin, latMax)
            If isOuter Then Continue For ' outer band bars handled elsewhere

            Dim wallLat As Double =
                If(lateralIdx = 0, wb.Edge(0), wb.Edge(1))

            Dim minDist As Double = Double.MaxValue
            Dim closest As Double = Double.NaN
            For Each scanPos As Double In scanPositions
                Dim d As Double = Math.Abs(scanPos - wallLat)
                If d < minDist Then
                    minDist = d
                    closest = scanPos
                End If
            Next

            If minDist < MinGalvanizeGap - Tolerance Then
                eliminatedEdges.Add(wb.Edge)
                Trace.TraceInformation(
                    ": HMG Layout: Inner wall eliminated — cut line " &
                    wallLat.ToString("F4") & " is " &
                    minDist.ToString("F4") &
                    """ from bearing bar at " &
                    closest.ToString("F4") & " (< " &
                    MinGalvanizeGap.ToString("F2") & """).")
            Else
                Trace.TraceInformation(
                    ": HMG Layout: Inner wall KEPT — cut line " &
                    wallLat.ToString("F4") & " is " &
                    minDist.ToString("F4") &
                    """ from nearest bearing bar at " &
                    closest.ToString("F4") & " (≥ " &
                    MinGalvanizeGap.ToString("F2") & """).")
            End If
        Next
    End Sub

    ''' <summary>
    ''' Internal record used by <see cref="ApplyMinGalvanizeGap"/>.
    ''' Position = lateral centerline coordinate.  IsBandBar = true for
    ''' candidates coming from a perimeter edge parallel to the span
    ''' axis.  IsOuterBand = true only when that edge is on the panel's
    ''' bounding box (latMin or latMax); outer band bars are immovable.
    ''' Edge = {ax, ay, bx, by} so the caller can record eliminated
    ''' notch-wall edges by coordinate.
    ''' </summary>
    Private Structure Candidate
        Public ReadOnly Position As Double
        Public ReadOnly IsBandBar As Boolean
        Public ReadOnly IsOuterBand As Boolean
        Public ReadOnly Edge As Double()
        Public Sub New(pos As Double, bandBar As Boolean,
                       outerBand As Boolean, edge_ As Double())
            Position = pos
            IsBandBar = bandBar
            IsOuterBand = outerBand
            Edge = edge_
        End Sub
    End Structure

    ''' <summary>
    ''' One parallel-to-span perimeter edge interpreted as a candidate
    ''' band bar.  Position is the band bar centerline (offset half a
    ''' bar width into the polygon material from the wall).  Edge holds
    ''' {ax, ay, bx, by} of the originating perimeter edge so the gap
    ''' rule can mark it for elimination.
    ''' </summary>
    Private Structure WallBandBar
        Public ReadOnly Position As Double
        Public ReadOnly Edge As Double()
        Public Sub New(pos As Double, edge_ As Double())
            Position = pos
            Edge = edge_
        End Sub
    End Structure

    ''' <summary>
    ''' True if a parallel-to-span band wall lies on the panel bounding
    ''' box in the lateral direction (latMin / latMax framing).
    ''' Outer band bars are structural and must not be eliminated by the
    ''' gap rule. Uses <see cref="OuterBBoxEpsilonInches"/> so slightly
    ''' noisy geometry still classifies as outer.
    ''' </summary>
    Private Function IsOuterPerimeterBandBar(
            wb As WallBandBar,
            lateralIdx As Integer,
            barWidth As Double,
            latMin As Double,
            latMax As Double) As Boolean

        If wb.Edge Is Nothing OrElse wb.Edge.Length < 4 Then Return False

        Dim latA As Double = If(lateralIdx = 0, wb.Edge(0), wb.Edge(1))
        Dim latB As Double = If(lateralIdx = 0, wb.Edge(2), wb.Edge(3))
        Dim edgeLat As Double = (latA + latB) / 2.0

        Dim latSpan As Double = latMax - latMin
        Dim tol As Double = Math.Max(
            OuterBBoxEpsilonInches,
            Math.Max(Tolerance, latSpan * 0.0005 + Tolerance))

        Return Math.Abs(edgeLat - latMin) < tol OrElse
               Math.Abs(edgeLat - latMax) < tol
    End Function

    ''' <summary>
    ''' True if a constant-span coordinate is on the panel bbox in span
    ''' direction (immovable perpendicular band bars for ApplyPerpendicularGap).
    ''' </summary>
    Private Function IsOuterSpanCoordinate(
            edgeSpan As Double,
            spanMin As Double,
            spanMax As Double) As Boolean
        Dim sSpan As Double = spanMax - spanMin
        Dim tol As Double = Math.Max(
            OuterBBoxEpsilonInches,
            Math.Max(Tolerance, sSpan * 0.0005 + Tolerance))
        Return Math.Abs(edgeSpan - spanMin) < tol OrElse
               Math.Abs(edgeSpan - spanMax) < tol
    End Function

    ''' <summary>
    ''' Walks the polygon edges and returns one <see cref="WallBandBar"/>
    ''' per perimeter edge that runs parallel to the bearing bars (i.e.
    ''' every edge whose lateral coordinate is constant).
    '''
    ''' For an outer rectangular perimeter this returns the latMin and
    ''' latMax edges.  For a perimeter with a notch whose walls are
    ''' parallel to the bearing bars, it additionally returns each notch
    ''' wall — exactly the case where bearing bars at standard scan
    ''' positions can land within MinGalvanizeGap of a wall band bar.
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
            barWidth As Double) As List(Of WallBandBar)

        Dim result As New List(Of WallBandBar)
        If polygon Is Nothing OrElse polygon.Count < 2 Then Return result

        Dim isCcw As Boolean = ComputeSignedArea(polygon) > 0.0
        Dim half As Double = barWidth / 2.0

        ' Outer lateral boundary edges (polyLatMin / polyLatMax) are NOT
        ' band bar walls — the outermost bearing bars sit flush there and
        ' form the side edges of the panel themselves.  Compute the
        ' extremes so we can skip those edges below.
        Dim polyLatMin As Double = Double.MaxValue
        Dim polyLatMax As Double = Double.MinValue
        For Each v As Double() In polygon
            If v(lateralIdx) < polyLatMin Then polyLatMin = v(lateralIdx)
            If v(lateralIdx) > polyLatMax Then polyLatMax = v(lateralIdx)
        Next

        For i As Integer = 0 To polygon.Count - 2
            Dim a = polygon(i)
            Dim b = polygon(i + 1)
            Dim dLat As Double = b(lateralIdx) - a(lateralIdx)

            ' Edge is parallel to span axis when its lateral coord is constant.
            If Math.Abs(dLat) >= Tolerance Then Continue For

            Dim dSpan As Double = b(spanIdx) - a(spanIdx)
            If Math.Abs(dSpan) < Tolerance Then Continue For ' degenerate

            ' Skip outer lateral boundary edges — these are not band bar walls.
            Dim edgeLatCoord As Double = a(lateralIdx)
            If Math.Abs(edgeLatCoord - polyLatMin) < Tolerance OrElse
               Math.Abs(edgeLatCoord - polyLatMax) < Tolerance Then Continue For

            Dim insideSign As Integer =
                ComputeInsideSign(lateralIdx, dSpan, isCcw)

            Dim edgeCoords As Double() =
                New Double() {a(0), a(1), b(0), b(1)}
            result.Add(New WallBandBar(edgeLatCoord + insideSign * half, edgeCoords))
        Next

        Return result
    End Function

    ''' <summary>Signed area of the closed polygon (positive => CCW).</summary>
    Private Function ComputeSignedArea(polygon As List(Of Double())) As Double
        Dim signedArea As Double = 0.0
        For i As Integer = 0 To polygon.Count - 2
            Dim a = polygon(i)
            Dim b = polygon(i + 1)
            signedArea += a(0) * b(1) - b(0) * a(1)
        Next
        Return signedArea
    End Function

    ''' <summary>
    ''' For a perimeter edge parallel to one axis, returns +1 or -1 so
    ''' that (edgeCoord + sign * halfWidth) lies inside the polygon
    ''' material.  See FindParallelEdgeBandBarPositions for the
    ''' geometric derivation.
    ''' </summary>
    ''' <param name="constantAxisIdx">
    ''' Index of the axis on which the edge has constant coordinate
    ''' (the axis we are offsetting along to find the inside direction).
    ''' </param>
    ''' <param name="dAlongEdge">
    ''' Vector component along the edge direction on the OTHER axis.
    ''' </param>
    Private Function ComputeInsideSign(
            constantAxisIdx As Integer,
            dAlongEdge As Double,
            isCcw As Boolean) As Integer
        Dim insideSign As Integer
        If constantAxisIdx = 0 Then
            insideSign = If(dAlongEdge > 0, -1, 1)
        Else
            insideSign = If(dAlongEdge > 0, 1, -1)
        End If
        If Not isCcw Then insideSign = -insideSign
        Return insideSign
    End Function

    ''' <summary>
    ''' Symmetric counterpart of <see cref="ApplyMinGalvanizeGap"/> along
    ''' the span axis: for every perimeter edge perpendicular to the
    ''' span axis (i.e. parallel to the cross bars), check whether the
    ''' wall band bar would land within MinGalvanizeGap of a cross bar
    ''' at one of the standard CrossBarOnCenter positions.  When it
    ''' does, the cross bar position is recorded for elimination by the
    ''' cross bar generator AND the wall edge is recorded so the band
    ''' bar generator does not produce a part for it.  Outer-perimeter
    ''' edges (along spanMin / spanMax of the bearing bar layout) are
    ''' immovable and are skipped.
    ''' </summary>
    Private Sub ApplyPerpendicularGap(
            polygon As List(Of Double()),
            params As GratingParameters,
            latMin As Double,
            latMax As Double,
            lateralIdx As Integer,
            spanIdx As Integer,
            eliminatedEdges As List(Of Double()),
            eliminatedCrossBarPositions As List(Of Double))

        If polygon Is Nothing OrElse polygon.Count < 2 Then Return

        ' Compute the bounding box on the span axis so we can identify
        ' which perpendicular edges are outer (immovable).
        Dim spanMin As Double = Double.MaxValue
        Dim spanMax As Double = Double.MinValue
        For Each v In polygon
            If v(spanIdx) < spanMin Then spanMin = v(spanIdx)
            If v(spanIdx) > spanMax Then spanMax = v(spanIdx)
        Next

        Dim isCcw As Boolean = ComputeSignedArea(polygon) > 0.0
        Dim half As Double = params.BarWidth / 2.0
        Dim minCenterDist As Double = params.BarWidth + MinGalvanizeGap

        ' Cross bar positions: spanMin + firstOffset + N*OC, same formula
        ' the cross bar generator uses.
        Dim cbPositions As New List(Of Double)
        Dim pos As Double = spanMin + params.FirstCrossBarOffset
        Do While pos <= spanMax + Tolerance
            cbPositions.Add(pos)
            pos += params.CrossBarOnCenter
        Loop

        For i As Integer = 0 To polygon.Count - 2
            Dim a = polygon(i)
            Dim b = polygon(i + 1)
            Dim dSpan As Double = b(spanIdx) - a(spanIdx)

            ' Perpendicular to span axis when the edge has constant span coord.
            If Math.Abs(dSpan) >= Tolerance Then Continue For

            Dim dLat As Double = b(lateralIdx) - a(lateralIdx)
            If Math.Abs(dLat) < Tolerance Then Continue For ' degenerate

            Dim edgeSpan As Double = a(spanIdx)

            ' Skip outer perimeter edges (immovable frame).
            If IsOuterSpanCoordinate(edgeSpan, spanMin, spanMax) Then Continue For

            Dim insideSign As Integer =
                ComputeInsideSign(spanIdx, dLat, isCcw)
            Dim wallBandBarPos As Double = edgeSpan + insideSign * half

            ' Find the closest cross bar position to this wall band bar.
            Dim conflictedCb As Double = Double.NaN
            Dim closestDist As Double = Double.MaxValue
            For Each cb In cbPositions
                Dim d As Double = Math.Abs(cb - wallBandBarPos)
                If d < closestDist Then
                    closestDist = d
                    conflictedCb = cb
                End If
            Next

            If closestDist < minCenterDist - Tolerance AndAlso
               Not Double.IsNaN(conflictedCb) Then
                ' Notch-wall band bar conflicts with a cross bar: eliminate
                ' the band bar (the cut snaps to the cross bar position),
                ' keep the cross bar.
                Dim edgeCoords As Double() =
                    New Double() {a(0), a(1), b(0), b(1)}
                eliminatedEdges.Add(edgeCoords)
                Trace.TraceInformation(
                    ": HMG Layout: Perpendicular gap — wall band bar at " &
                    "span=" & wallBandBarPos.ToString("F4") &
                    " conflicts with cross bar at " &
                    conflictedCb.ToString("F4") &
                    " (gap=" & closestDist.ToString("F4") &
                    """); eliminating band bar, keeping cross bar.")
            End If
        Next
    End Sub

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
                                           barWidth As Double,
                                           arcEdgeMap As Dictionary(Of Integer, ArcEdgeContext)) As List(Of ScanIntersection)
        Dim hits As New List(Of ScanIntersection)

        For ei As Integer = 0 To edges.Count - 1
            Dim edge As Double()() = edges(ei)
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
            Dim chordSpanHit As Double = a(spanIdx) + t * dSpan

            ' Arc-aware override: if this edge is a tessellated chord of
            ' an arc, replace the chord interpolation with the exact
            ' intersection of the scan line and the inner-face circle
            ' (radius R ± barWidth centred on the arc).  The hit's
            ' EdgeInset is set to zero — the circle root already sits on
            ' the curved band bar's inner face, no banded shift needed.
            If arcEdgeMap IsNot Nothing AndAlso arcEdgeMap.ContainsKey(ei) Then
                Dim ctx As ArcEdgeContext = arcEdgeMap(ei)
                Dim cLat As Double = If(latIdx = 0, ctx.CenterX, ctx.CenterY)
                Dim cSpan As Double = If(spanIdx = 0, ctx.CenterX, ctx.CenterY)
                ' Use the bar's NEAR EDGE (latPos ± halfBarWidth, whichever
                ' is closer to cLat), not its centerline.  Otherwise the
                ' bar's corner protrudes ~halfBarWidth past the inner face
                ' into the curved band bar's body.  When the bar straddles
                ' cLat (|latPos − cLat| ≤ halfBW), the closest point on
                ' the bar to the arc centre lies on its edge at Y = cLat,
                ' so nearEdgeDy collapses to zero and X = cSpan ± innerR.
                Dim halfBW As Double = barWidth / 2.0
                Dim dyCenter As Double = latPos - cLat
                Dim absDy As Double = Math.Abs(dyCenter)
                Dim nearEdgeDy As Double =
                    If(absDy <= halfBW, 0.0, absDy - halfBW)
                Dim disc As Double =
                    ctx.InnerFaceRadius * ctx.InnerFaceRadius -
                    nearEdgeDy * nearEdgeDy
                If disc >= 0 Then
                    Dim root As Double = Math.Sqrt(disc)
                    Dim r1 As Double = cSpan + root
                    Dim r2 As Double = cSpan - root
                    ' Pick whichever circle root is closer to the chord's
                    ' interpolated hit — correct for both single-arc
                    ' cutouts (one chord hit) and circular holes (two
                    ' chord hits, each mapping to its own root).
                    Dim arcHit As Double =
                        If(Math.Abs(r1 - chordSpanHit) <=
                           Math.Abs(r2 - chordSpanHit), r1, r2)
                    hits.Add(New ScanIntersection(arcHit, 0))
                    Continue For
                End If
                ' Discriminant < 0: scan line lies outside the inner-face
                ' circle.  Fall through to chord-based behaviour so the
                ' bearing bar is still trimmed (better than dropping it).
            End If

            ' Per-edge along-span inset for banded-mode trimming.  Equals
            ' barWidth / |sin(theta)| where theta is angle between the
            ' edge and the span axis.  For edges perpendicular to span
            ' (|dLat| = edgeLen) the inset equals barWidth.
            Dim edgeLen As Double = Math.Sqrt(dSpan * dSpan + dLat * dLat)
            Dim edgeInset As Double = barWidth * edgeLen / Math.Abs(dLat)

            hits.Add(New ScanIntersection(chordSpanHit, edgeInset))
        Next

        Return hits
    End Function

    ''' <summary>
    ''' Inner-face geometry for one tessellated chord of an arc.
    ''' Used by <see cref="FindScanIntersections"/> to replace the
    ''' chord interpolation with the exact circle intersection at the
    ''' curved band bar's inner face.
    ''' </summary>
    Private Structure ArcEdgeContext
        Public CenterX As Double
        Public CenterY As Double
        Public ArcRadius As Double
        Public InnerFaceRadius As Double
    End Structure

    ''' <summary>
    ''' Builds a mapping from polygon edge index to inner-face circle
    ''' geometry for every chord edge that belongs to an arc.  Edges
    ''' not in the map are straight perimeter edges and use the
    ''' existing chord-based clip.
    ''' </summary>
    Private Function BuildArcEdgeMap(perimeter As PerimeterData,
                                     poly As List(Of Double()),
                                     barWidth As Double,
                                     edgeCount As Integer) As Dictionary(Of Integer, ArcEdgeContext)
        Dim map As New Dictionary(Of Integer, ArcEdgeContext)
        If perimeter Is Nothing OrElse perimeter.ArcSegments Is Nothing OrElse
           perimeter.ArcSegments.Count = 0 Then
            Return map
        End If

        ' Signed area → +1 for CCW polygon (perpendicular -ey,ex points
        ' inward), −1 for CW.  ArcExtendsOutward classification matches
        ' BandBarGenerator: sweep × perpSign < 0 ⇒ cutout/outward arc.
        Dim signedArea As Double = 0
        Dim n As Integer = poly.Count - 1 ' last vertex duplicates first
        For vi As Integer = 0 To n - 1
            Dim vj As Integer = (vi + 1) Mod n
            signedArea += poly(vi)(0) * poly(vj)(1) -
                          poly(vj)(0) * poly(vi)(1)
        Next
        Dim perpSign As Double = If(signedArea > 0, 1.0, -1.0)

        For Each arc As PerimeterArcInfo In perimeter.ArcSegments
            Dim outward As Boolean = (arc.SweepAngle * perpSign < 0)
            Dim innerFaceR As Double =
                If(outward, arc.Radius + barWidth,
                            arc.Radius - barWidth)
            If innerFaceR <= 0 Then Continue For ' degenerate, skip

            Dim ctx As New ArcEdgeContext() With {
                .CenterX = arc.CenterX,
                .CenterY = arc.CenterY,
                .ArcRadius = arc.Radius,
                .InnerFaceRadius = innerFaceR}

            ' Arc occupies VertexCount vertices (entry + intermediate);
            ' its chord edges are FirstVertexIndex .. FirstVertexIndex +
            ' VertexCount − 1 (the last chord lands on the exit vertex
            ' shared with the next perimeter section).
            Dim lo As Integer = arc.FirstVertexIndex
            Dim hi As Integer = arc.FirstVertexIndex + arc.VertexCount - 1
            For ei As Integer = lo To hi
                If ei >= 0 AndAlso ei < edgeCount Then
                    map(ei) = ctx
                End If
            Next

            Trace.TraceInformation(
                ": HMG Layout: Arc clip — centre=(" &
                arc.CenterX.ToString("F3") & "," &
                arc.CenterY.ToString("F3") & ") R=" &
                arc.Radius.ToString("F3") & " outward=" &
                outward.ToString() & " innerFaceR=" &
                innerFaceR.ToString("F3") & " edges=" &
                lo & ".." & hi)
        Next

        Return map
    End Function

    ''' <summary>
    ''' Splits a single (spanStart, spanEnd) bearing-bar segment around
    ''' every arc whose body intersects the bar's body but whose
    ''' centerline the chord scan missed.  Handles the case the user
    ''' showed in v1.5.17 where one bar still ran through the arc: its
    ''' lateral centerline sat just outside [Cx-R, Cx+R], so no arc
    ''' chord was crossed, but the bar's near edge protruded into the
    ''' band bar.  Result: 0, 1, or 2 sub-segments per input segment.
    '''
    ''' Centerline-IN-arc cases (|latPos - cLat| ≤ R) are already
    ''' trimmed by the v1.5.17 chord-arc replacement in
    ''' FindScanIntersections, so they're skipped here to avoid
    ''' double-cutting.
    ''' </summary>
    Private Function ApplyArcBodyCuts(
            spanStart As Double, spanEnd As Double,
            latPos As Double, barWidth As Double,
            arcEdgeMap As Dictionary(Of Integer, ArcEdgeContext),
            latIdx As Integer, spanIdx As Integer) As List(Of Double())

        Dim segments As New List(Of Double())
        segments.Add(New Double() {spanStart, spanEnd})

        If arcEdgeMap Is Nothing OrElse arcEdgeMap.Count = 0 Then
            Return segments
        End If

        Dim halfBW As Double = barWidth / 2.0
        Dim seen As New HashSet(Of String)

        For Each ctx As ArcEdgeContext In arcEdgeMap.Values
            Dim key As String =
                ctx.CenterX.ToString("F6") & "|" &
                ctx.CenterY.ToString("F6") & "|" &
                ctx.InnerFaceRadius.ToString("F6")
            If Not seen.Add(key) Then Continue For

            Dim cLat As Double = If(latIdx = 0, ctx.CenterX, ctx.CenterY)
            Dim cSpan As Double = If(spanIdx = 0, ctx.CenterX, ctx.CenterY)
            Dim absDy As Double = Math.Abs(latPos - cLat)

            ' Skip arcs the chord scan already trimmed (v1.5.17 path):
            ' the bar's centerline sits inside the arc radius, so chord
            ' edges were crossed and the inner-face circle was already
            ' substituted for them in FindScanIntersections.
            If absDy <= ctx.ArcRadius + Tolerance Then Continue For

            ' Body-cut: cut zone exists when the bar's near edge sits
            ' inside the inner-face circle.  nearEdgeDy uses the bar's
            ' near edge (corner-aware), matching the v1.5.17 nearEdgeDy
            ' derivation.
            Dim nearEdgeDy As Double = absDy - halfBW
            If nearEdgeDy < 0 Then nearEdgeDy = 0
            Dim disc As Double =
                ctx.InnerFaceRadius * ctx.InnerFaceRadius -
                nearEdgeDy * nearEdgeDy
            If disc <= 0 Then Continue For ' bar's body is clear of arc

            Dim root As Double = Math.Sqrt(disc)
            Dim cutLow As Double = cSpan - root
            Dim cutHigh As Double = cSpan + root

            ' Subtract (cutLow, cutHigh) from each current segment.
            Dim newSegments As New List(Of Double())
            For Each seg As Double() In segments
                Dim s As Double = seg(0)
                Dim e As Double = seg(1)
                If e <= cutLow OrElse s >= cutHigh Then
                    newSegments.Add(seg) ' no overlap, keep as-is
                ElseIf s >= cutLow AndAlso e <= cutHigh Then
                    ' Segment lies entirely inside the cut zone — drop.
                ElseIf s < cutLow AndAlso e <= cutHigh Then
                    newSegments.Add(New Double() {s, cutLow})
                ElseIf s >= cutLow AndAlso e > cutHigh Then
                    newSegments.Add(New Double() {cutHigh, e})
                Else
                    ' Cut zone strictly inside segment — split into two.
                    newSegments.Add(New Double() {s, cutLow})
                    newSegments.Add(New Double() {cutHigh, e})
                End If
            Next
            segments = newSegments
        Next

        Return segments
    End Function

#End Region

End Class
