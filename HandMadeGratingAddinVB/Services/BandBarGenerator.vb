'////////////////////////////////////////////////////////////////////
' Metal Bar Grating Addin - VB.NET
'
' BandBarGenerator: Generates band bar .ipt part files for banded
' grating projects.
'
' Phase 18: Band bars follow the perimeter edge of the grating.
' Each edge of the perimeter polygon produces one band bar segment
' with the same cross-section as the bearing bars (BarDepth x BarWidth).
'
' First implementation assumptions:
'   - Each perimeter edge = one straight band bar segment.
'   - Cross-section is a simple rectangle (BarWidth x BarDepth).
'   - Convex corners use 45° plan-view miters so outer faces stay
'     flush on the perimeter; re-entrant corners use square ends.
'   - Band bars that are shorter than 0.1" are skipped (degenerate).
'   - Curved perimeter edges are approximated by their chord length
'     (the straight-line distance between vertices).
'
' Geometry per band bar:
'   Rectangle sketch on XY plane (BarWidth x BarDepth), extruded +Z
'   by the segment length.
'////////////////////////////////////////////////////////////////////

Imports System.Diagnostics
Imports System.IO
Imports Inventor

''' <summary>
''' Generates band bar .ipt part files from the perimeter edges.
''' </summary>
Public Class BandBarGenerator

    Private Const MinSegmentLength As Double = 0.1 ' inches — skip degenerate edges

    Private ReadOnly _app As Application
    Private ReadOnly _pathService As New OutputPathService()

    Public Sub New(app As Application)
        _app = app
    End Sub

    ''' <summary>
    ''' Generates band bar parts if the project is Banded.
    ''' Returns a Skipped result if the project is OpenEnded.
    '''
    ''' <paramref name="eliminatedEdges"/> lists perimeter edges (as
    ''' {ax, ay, bx, by} coordinate quads) whose band bar was eliminated
    ''' by the galvanize-gap rule in the bearing bar layout / cross bar
    ''' stages.  Those edges are skipped so the assembly does not double
    ''' up a band bar against the bearing bar or cross bar that the rule
    ''' chose to keep.
    ''' </summary>
    Public Function Generate(perimeter As PerimeterData,
                             params As GratingParameters,
                             Optional eliminatedEdges As List(Of Double()) = Nothing) As BandBarGenerationResult
        Try
            ' --- Guard: OpenEnded projects skip band bar generation ---
            If params.Banding = BandingOptionType.OpenEnded Then
                Trace.TraceInformation(
                    ": HMG BandBarGen: Banding=OpenEnded — skipping band bar generation.")
                Return BandBarGenerationResult.SkippedOpenEnded()
            End If

            Trace.TraceInformation(
                ": HMG BandBarGen: Banding=Banded — starting band bar generation.")

            ' Validate perimeter
            If perimeter Is Nothing OrElse
               perimeter.OuterLoopVertices Is Nothing OrElse
               perimeter.OuterLoopVertices.Count < 3 Then
                Return BandBarGenerationResult.Failed(
                    "No valid perimeter available for band bar generation.")
            End If

            ' Resolve output folder
            Dim outputFolder As String = _pathService.ResolveOutputFolder(
                params.OutputFolder, _app)
            If String.IsNullOrEmpty(outputFolder) Then
                Return BandBarGenerationResult.Failed(
                    "Cannot resolve output folder.")
            End If
            If Not Directory.Exists(outputFolder) Then
                Try
                    Directory.CreateDirectory(outputFolder)
                Catch ex As Exception
                    Return BandBarGenerationResult.Failed(
                        "Cannot create output folder: " & ex.Message)
                End Try
            End If

            ' --- Step 1: Compute band bar segments from perimeter edges ---
            Dim segments As List(Of BandBarSegment) =
                ComputeSegments(perimeter, params.ResolvedPrefix,
                                params.BarWidth, params.SpanDirection,
                                eliminatedEdges)

            If segments.Count = 0 Then
                Return BandBarGenerationResult.Failed(
                    "No band bar segments could be computed from the perimeter.")
            End If

            Trace.TraceInformation(": HMG BandBarGen: " &
                segments.Count & " segment(s) computed from " &
                perimeter.OuterLoopVertices.Count & " perimeter vertices.")

            For Each seg In segments
                Trace.TraceInformation(": HMG BandBarGen:   " & seg.ToString())
            Next

            ' --- Step 2: Generate one .ipt per segment ---
            Dim files As New List(Of GeneratedBandBarFile)
            Dim warnings As New List(Of String)

            ' Capture the original active document to restore after generation
            Dim originalDoc As Document = Nothing
            Try
                originalDoc = _app.ActiveDocument
            Catch
            End Try

            For Each seg In segments
                Dim genFile As GeneratedBandBarFile

                If seg.IsArc Then
                    genFile = GenerateArcBandBar(seg, params, outputFolder)
                Else
                    genFile = GenerateSingleBandBar(seg, params, outputFolder)
                End If

                files.Add(genFile)

                If genFile.Saved Then
                    Trace.TraceInformation(": HMG BandBarGen:   OK — " &
                        genFile.ToString())
                    ' Surface non-fatal warnings (e.g. fallback rectangle
                    ' when the mitered profile was rejected) so they
                    ' appear in the Generation Summary panel.
                    If Not String.IsNullOrEmpty(genFile.WarningMessage) Then
                        warnings.Add(seg.Mark & ": " & genFile.WarningMessage)
                    End If
                Else
                    Trace.TraceWarning(": HMG BandBarGen:   FAIL — " &
                        genFile.ToString())
                    warnings.Add(seg.Mark & ": " & genFile.ErrorMessage)
                End If
            Next

            ' Restore original active document
            If originalDoc IsNot Nothing Then
                Try
                    originalDoc.Activate()
                Catch
                End Try
            End If

            Dim savedCount As Integer = 0
            For Each f In files
                If f.Saved Then savedCount += 1
            Next

            If savedCount = 0 Then
                Return BandBarGenerationResult.Failed(
                    "All band bar files failed to generate.")
            End If

            Trace.TraceInformation(": HMG BandBarGen: Complete — " &
                savedCount & "/" & files.Count & " files saved.")

            Return BandBarGenerationResult.Succeeded(files, outputFolder, warnings)

        Catch ex As Exception
            Trace.TraceError(": HMG BandBarGen: Unexpected error — " & ex.ToString())
            Return BandBarGenerationResult.Failed(
                "Band bar generation error: " & ex.Message)
        End Try
    End Function

    ' ==================================================================
    '  Step 1: Compute band bar segments from perimeter
    ' ==================================================================

    ''' <summary>
    ''' Extracts perimeter edges as band bar segments.
    ''' Each consecutive pair of vertices defines one band bar.
    ''' Closes the polygon by connecting the last vertex back to the first.
    ''' Arc edges (identified via PerimeterArcInfo) produce a single curved
    ''' segment instead of many short straight segments.
    ''' </summary>
    Private Function ComputeSegments(perimeter As PerimeterData,
                                     prefix As String,
                                     barWidth As Double,
                                     spanDirection As SpanDirectionType,
                                     eliminatedEdges As List(Of Double())) As List(Of BandBarSegment)
        Dim segments As New List(Of BandBarSegment)
        Dim verts As List(Of Double()) = perimeter.OuterLoopVertices
        Dim n As Integer = verts.Count

        ' Walk the perimeter edges. The vertex list may or may not be closed
        ' (last vertex duplicating the first). Detect and handle both cases.
        Dim isClosed As Boolean = False
        If n >= 2 Then
            Dim first As Double() = verts(0)
            Dim last As Double() = verts(n - 1)
            If Math.Abs(first(0) - last(0)) < 0.0001 AndAlso
               Math.Abs(first(1) - last(1)) < 0.0001 Then
                isClosed = True
            End If
        End If

        ' Number of edges
        Dim edgeCount As Integer
        If isClosed Then
            edgeCount = n - 1  ' last vertex is duplicate of first
        Else
            edgeCount = n      ' need to close: add edge from last to first
        End If

        ' Compute polygon signed area to determine winding direction.
        ' Positive = CCW (perpendicular (-ey,ex) points inward).
        ' Negative = CW  (perpendicular (-ey,ex) points outward).
        Dim signedArea As Double = 0
        Dim nVerts As Integer = If(isClosed, n - 1, n)
        For vi As Integer = 0 To nVerts - 1
            Dim vj As Integer = (vi + 1) Mod nVerts
            signedArea += verts(vi)(0) * verts(vj)(1) - verts(vj)(0) * verts(vi)(1)
        Next
        Dim perpSign As Double = If(signedArea > 0, 1.0, -1.0)

        ' Compute polygon bounds along the lateral axis (perpendicular to
        ' the bearing bars).  Parallel-to-span edges whose lateral coord
        ' matches polyLatMin / polyLatMax are the OUTER panel side edges
        ' and get no band bar — the outermost flush bearing bars form
        ' those sides.  All other parallel-to-span edges (i.e. inner
        ' notch walls) DO get a band bar, subject to the galvanize-gap
        ' rule applied in BearingBarLayoutService.ApplyMinGalvanizeGap.
        Dim lateralIdx As Integer =
            If(spanDirection = SpanDirectionType.AlongY, 0, 1)
        Dim polyLatMin As Double = Double.MaxValue
        Dim polyLatMax As Double = Double.MinValue
        For vi As Integer = 0 To nVerts - 1
            Dim latV As Double = verts(vi)(lateralIdx)
            If latV < polyLatMin Then polyLatMin = latV
            If latV > polyLatMax Then polyLatMax = latV
        Next

        ' Classify each vertex corner as convex or concave AND compute
        ' the interior angle (in radians) so corner shrinkage can be
        ' angle-corrected.  At an interior angle α, two adjacent band
        ' bars must each step back from the vertex by bw / tan(α/2)
        ' for their inner faces to meet at one point.  A fixed bw
        ' shrinkage is only correct at α = 90°; non-perpendicular
        ' corners (e.g. slants) need the angle-corrected value.
        Dim cornerConcave(nVerts - 1) As Boolean
        Dim cornerInteriorAngle(nVerts - 1) As Double
        For vi As Integer = 0 To nVerts - 1
            Dim prevV As Integer = (vi - 1 + nVerts) Mod nVerts
            Dim nxtV As Integer = (vi + 1) Mod nVerts
            Dim dxP As Double = verts(vi)(0) - verts(prevV)(0)
            Dim dyP As Double = verts(vi)(1) - verts(prevV)(1)
            Dim dxN As Double = verts(nxtV)(0) - verts(vi)(0)
            Dim dyN As Double = verts(nxtV)(1) - verts(vi)(1)
            Dim cross As Double = dxP * dyN - dyP * dxN
            cornerConcave(vi) = (perpSign * cross < -0.001)

            ' Interior angle from incoming (eA) and outgoing (eB) unit vectors.
            Dim lenP As Double = Math.Sqrt(dxP * dxP + dyP * dyP)
            Dim lenN As Double = Math.Sqrt(dxN * dxN + dyN * dyN)
            If lenP < 0.0001 OrElse lenN < 0.0001 Then
                ' Degenerate edge length — treat as 180° (no corner).
                cornerInteriorAngle(vi) = Math.PI
            Else
                Dim eAx As Double = dxP / lenP, eAy As Double = dyP / lenP
                Dim eBx As Double = dxN / lenN, eBy As Double = dyN / lenN
                Dim turnCross As Double = eAx * eBy - eAy * eBx
                Dim turnDot As Double = eAx * eBx + eAy * eBy
                Dim turn As Double = Math.Atan2(turnCross, turnDot)
                If perpSign < 0 Then turn = -turn ' CW polygon: flip sign
                cornerInteriorAngle(vi) = Math.PI - turn
            End If
        Next

        Dim safePfx As String = If(prefix, "Grating")
        Dim idx As Integer = 0

        ' Build a vertex-index → arc-index mapping so arc tessellation
        ' edges can be collapsed into a single curved band bar segment.
        Dim arcVertexMap As New Dictionary(Of Integer, Integer)
        Dim arcs As List(Of PerimeterArcInfo) = perimeter.ArcSegments
        If arcs IsNot Nothing Then
            For ai As Integer = 0 To arcs.Count - 1
                Dim arc As PerimeterArcInfo = arcs(ai)
                For vi As Integer = arc.FirstVertexIndex To _
                        arc.FirstVertexIndex + arc.VertexCount - 1
                    arcVertexMap(vi) = ai
                Next
            Next
        End If

        For i As Integer = 0 To edgeCount - 1
            Dim j As Integer = (i + 1) Mod n
            If isClosed AndAlso j = n - 1 Then j = 0

            ' --- Arc detection: skip intermediate arc edges, emit one
            '     curved segment when we reach an arc's first vertex ---
            Dim mappedVert As Integer = i Mod nVerts
            If arcVertexMap.ContainsKey(mappedVert) Then
                Dim arcIdx As Integer = arcVertexMap(mappedVert)
                Dim arcInf As PerimeterArcInfo = arcs(arcIdx)
                If mappedVert = arcInf.FirstVertexIndex Then
                    ' First edge of this arc — emit a curved band bar
                    idx += 1
                    Dim arcSeg As New BandBarSegment()
                    arcSeg.Index = idx
                    arcSeg.Mark = safePfx & "-BAND-" & idx.ToString("000")
                    arcSeg.IsArc = True
                    arcSeg.ArcCenterX = arcInf.CenterX
                    arcSeg.ArcCenterY = arcInf.CenterY
                    arcSeg.ArcRadius = arcInf.Radius
                    arcSeg.ArcEntryAngle = arcInf.EntryAngle
                    arcSeg.ArcSweepAngle = arcInf.SweepAngle

                    ' Arc length = radius × |sweep|
                    arcSeg.Length = arcInf.Radius * Math.Abs(arcInf.SweepAngle)

                    ' Start/End points on the arc (for logging / assembly)
                    arcSeg.StartPoint = New Double() {
                        arcInf.CenterX + arcInf.Radius *
                            Math.Cos(arcInf.EntryAngle),
                        arcInf.CenterY + arcInf.Radius *
                            Math.Sin(arcInf.EntryAngle)}
                    Dim exitAngle As Double =
                        arcInf.EntryAngle + arcInf.SweepAngle
                    arcSeg.EndPoint = New Double() {
                        arcInf.CenterX + arcInf.Radius *
                            Math.Cos(exitAngle),
                        arcInf.CenterY + arcInf.Radius *
                            Math.Sin(exitAngle)}

                    ' Determine inset direction:
                    ' sweep * perpSign < 0 → cutout arc, bar extends OUTWARD
                    ' sweep * perpSign > 0 → convex arc, bar extends INWARD
                    arcSeg.ArcExtendsOutward =
                        (arcInf.SweepAngle * perpSign < 0)

                    arcSeg.IsParallel = False
                    segments.Add(arcSeg)

                    Trace.TraceInformation(": HMG BandBarGen: Arc segment " &
                        arcSeg.Mark & " R=" &
                        arcInf.Radius.ToString("F4") & """ sweep=" &
                        (arcInf.SweepAngle * 180 / Math.PI).ToString("F1") &
                        "° outward=" & arcSeg.ArcExtendsOutward.ToString())
                End If
                ' Skip all arc tessellation edges (both first and intermediate)
                Continue For
            End If

            Dim p0 As Double() = verts(i)
            Dim p1 As Double() = verts(j)

            ' Skip edges marked as eliminated by the galvanize-gap rule
            ' (notch walls whose band bar collided with a bearing bar or
            ' cross bar within MinGalvanizeGap).
            If IsEdgeEliminated(p0, p1, eliminatedEdges) Then
                Trace.TraceInformation(
                    ": HMG BandBarGen: Skipping edge " & i &
                    " — eliminated by galvanize-gap rule.")
                Continue For
            End If

            Dim dx As Double = p1(0) - p0(0)
            Dim dy As Double = p1(1) - p0(1)
            Dim segLength As Double = Math.Sqrt(dx * dx + dy * dy)

            If segLength < MinSegmentLength Then
                Trace.TraceWarning(
                    ": HMG BandBarGen: Skipping degenerate edge " &
                    i & " (length=" & segLength.ToString("F4") & """)")
                Continue For
            End If

            Dim halfW As Double = barWidth / 2.0
            Dim edx As Double = dx / segLength  ' unit edge direction
            Dim edy As Double = dy / segLength

            ' Perpendicular direction: (-edy, edx). For CCW polygons this
            ' points inward; for CW it points outward. perpSign corrects
            ' for winding so that (perpSign * px, perpSign * py) always
            ' points toward the polygon interior.
            Dim px As Double = -edy
            Dim py As Double = edx

            ' Edge classification.
            ' "Parallel" here means parallel to the BEARING BARS (i.e.
            ' the edge runs along the span direction).  Side edges on
            ' the OUTER panel boundary (polyLatMin / polyLatMax) are
            ' skipped — the outermost flush bearing bars form those
            ' sides.  Parallel edges OFF the outer boundary (i.e. inner
            ' notch walls) DO get a band bar per the PDF rule;
            ' BearingBarLayoutService.ApplyMinGalvanizeGap eliminates
            ' the wall edge ahead of time if its band bar would land
            ' within 1/4" of an adjacent bearing bar, in which case
            ' IsEdgeEliminated above already skipped this iteration.
            ' Tight tolerance so steep diagonals (slants) count as
            ' end edges, not side edges.
            Const LateralTolerance As Double = 0.01
            Dim isParallel As Boolean
            If spanDirection = SpanDirectionType.AlongX Then
                isParallel = (Math.Abs(edy) < LateralTolerance)
            Else
                isParallel = (Math.Abs(edx) < LateralTolerance)
            End If

            If isParallel Then
                Dim edgeLat As Double = p0(lateralIdx)
                Dim isOuterPanelSide As Boolean =
                    Math.Abs(edgeLat - polyLatMin) < 0.0001 OrElse
                    Math.Abs(edgeLat - polyLatMax) < 0.0001
                If isOuterPanelSide Then Continue For
            End If

            ' Determine whether the neighbouring perimeter edge at each
            ' corner is itself a "skip" (side) edge.  When it is, this
            ' bar's miter at that corner should be 0 — the flush bearing
            ' bar's full width covers the corner, no inner-face cut needed.
            Dim prevEdgeI As Integer = (i - 1 + edgeCount) Mod edgeCount
            Dim nextEdgeI As Integer = (i + 1) Mod edgeCount
            Dim prevP0 As Double() = verts(prevEdgeI Mod nVerts)
            Dim prevP1 As Double() = verts((prevEdgeI + 1) Mod nVerts)
            Dim nextP0 As Double() = verts(nextEdgeI Mod nVerts)
            Dim nextP1 As Double() = verts((nextEdgeI + 1) Mod nVerts)
            Dim startNeighborIsParallel As Boolean =
                IsEdgeParallelToSpan(prevP0, prevP1, spanDirection, LateralTolerance)
            Dim endNeighborIsParallel As Boolean =
                IsEdgeParallelToSpan(nextP0, nextP1, spanDirection, LateralTolerance)

            ' Corner handling:
            '   Convex corner (interior < 180°):
            '     Keep the outer face on the full perimeter edge
            '     (through the vertex) and miter-cut the inner corner
            '     in the .ipt — both incident bars get a miter trim.
            '   Concave corner (interior > 180°, re-entrant inner corner
            '     of an L-shape / notch):
            '     Square ends at the vertex (miter = 0).  The
            '     perpendicular bar additionally EXTENDS past the
            '     vertex by barWidth / tan((2π − α)/2) so its width
            '     covers the otherwise-unfilled square at the inside
            '     of the L.  Same pattern the pre-v1.4.6 code used,
            '     now angle-corrected.
            Dim startVertIdx As Integer = i Mod nVerts
            Dim endVertIdx As Integer = j Mod nVerts

            Dim startMiter As Double = 0.0
            Dim endMiter As Double = 0.0
            Dim startExt As Double = 0.0
            Dim endExt As Double = 0.0
            If cornerConcave(startVertIdx) Then
                ' Concave corner — only the PERPENDICULAR bar extends
                ' past the vertex.  Its width covers the L-pocket and
                ' the adjacent PARALLEL bar (inner notch wall, if
                ' present) butts square against the perpendicular
                ' bar's inner face.  Without this isParallel guard,
                ' both bars would extend into the corner and overlap.
                If Not isParallel Then
                    startExt = ConcaveCornerExtension(
                        cornerInteriorAngle(startVertIdx), barWidth)
                End If
            ElseIf Not startNeighborIsParallel Then
                ' Convex corner with another band bar on the other side:
                ' miter the inner corner.  When the neighbor is a side
                ' edge (skipped), leave the end square at the vertex —
                ' the flush bearing bar covers that corner.
                startMiter = CornerMiterTrim(
                    cornerInteriorAngle(startVertIdx), barWidth)
            End If
            If cornerConcave(endVertIdx) Then
                ' Only the perpendicular bar extends at concave corners
                ' (see startVertIdx branch above for rationale).
                If Not isParallel Then
                    endExt = ConcaveCornerExtension(
                        cornerInteriorAngle(endVertIdx), barWidth)
                End If
            ElseIf Not endNeighborIsParallel Then
                endMiter = CornerMiterTrim(
                    cornerInteriorAngle(endVertIdx), barWidth)
            End If

            Dim extStart As Double() = New Double() {
                p0(0) - startExt * edx, p0(1) - startExt * edy}
            Dim extEnd As Double() = New Double() {
                p1(0) + endExt * edx, p1(1) + endExt * edy}
            Dim extLength As Double = segLength + startExt + endExt

            ' Shift the segment inward by halfWidth perpendicular to the
            ' edge.  Combined with the placement matrix offset (-halfW * perp)
            ' this places the bar outer face exactly on the perimeter edge
            ' with the full bar width sitting inside the perimeter.
            Dim insetX As Double = perpSign * halfW * px
            Dim insetY As Double = perpSign * halfW * py
            extStart(0) += insetX
            extStart(1) += insetY
            extEnd(0) += insetX
            extEnd(1) += insetY

            idx += 1
            Dim seg As New BandBarSegment()
            seg.Index = idx
            seg.Mark = safePfx & "-BAND-" & idx.ToString("000")
            seg.StartPoint = extStart
            seg.EndPoint = extEnd
            seg.Length = extLength
            seg.IsParallel = isParallel
            seg.PerpSign = perpSign
            seg.StartMiterTrim = startMiter
            seg.EndMiterTrim = endMiter
            segments.Add(seg)
        Next

        Return segments
    End Function

    ''' <summary>
    ''' True if the perimeter edge from <paramref name="p0"/> to
    ''' <paramref name="p1"/> runs (nearly) parallel to the bearing bars'
    ''' span direction — i.e. it is a "side edge" that gets no band bar.
    ''' Tolerance is on the unit-vector component, so steep slants still
    ''' count as end edges.
    ''' </summary>
    Private Function IsEdgeParallelToSpan(p0 As Double(),
                                          p1 As Double(),
                                          spanDirection As SpanDirectionType,
                                          tolerance As Double) As Boolean
        Dim dx As Double = p1(0) - p0(0)
        Dim dy As Double = p1(1) - p0(1)
        Dim segLen As Double = Math.Sqrt(dx * dx + dy * dy)
        If segLen < 0.0001 Then Return False
        Dim edx As Double = dx / segLen
        Dim edy As Double = dy / segLen
        If spanDirection = SpanDirectionType.AlongX Then
            Return Math.Abs(edy) < tolerance
        Else
            Return Math.Abs(edx) < tolerance
        End If
    End Function

    ''' <summary>
    ''' Plan-view miter trim length at a convex corner (inches along the bar).
    ''' Equals barWidth / tan(interiorAngle/2) — barWidth at 90° corners.
    ''' </summary>
    Private Function CornerMiterTrim(interiorAngle As Double,
                                       barWidth As Double) As Double
        If interiorAngle <= 0.001 OrElse interiorAngle >= Math.PI - 0.001 Then
            Return 0.0
        End If
        If interiorAngle > Math.PI Then Return 0.0 ' re-entrant — no miter

        Dim halfAngle As Double = interiorAngle / 2.0
        Dim s As Double = Math.Sin(halfAngle)
        If Math.Abs(s) < 0.001 Then Return 0.0
        Return barWidth * Math.Cos(halfAngle) / s
    End Function

    ''' <summary>
    ''' Past-the-vertex extension length for a band bar at a concave
    ''' (re-entrant) corner.  At a concave corner the perpendicular bar
    ''' must extend along its edge direction so its width covers the
    ''' otherwise-unfilled square on the inside of the L-shape.
    '''
    ''' For interior angle α &gt; π (concave), the exterior angle is
    ''' (2π − α); the bar extends by:
    '''   barWidth / tan((2π − α)/2)
    ''' At α = 270° (a 90° re-entrant corner) this collapses to barWidth.
    ''' Returns 0 at non-concave corners.
    ''' </summary>
    Private Function ConcaveCornerExtension(interiorAngle As Double,
                                             barWidth As Double) As Double
        If interiorAngle <= Math.PI + 0.001 Then Return 0.0
        If interiorAngle >= 2 * Math.PI - 0.001 Then Return 0.0

        Dim halfExterior As Double = (2.0 * Math.PI - interiorAngle) / 2.0
        Dim s As Double = Math.Sin(halfExterior)
        If Math.Abs(s) < 0.001 Then Return 0.0
        Return barWidth * Math.Cos(halfExterior) / s
    End Function

    ''' <summary>
    ''' True if the edge from <paramref name="p0"/> to <paramref name="p1"/>
    ''' matches an entry in <paramref name="eliminatedEdges"/> (in either
    ''' direction).  Eliminated entries are stored as {ax, ay, bx, by}.
    ''' </summary>
    Private Function IsEdgeEliminated(p0 As Double(), p1 As Double(),
                                      eliminatedEdges As List(Of Double())) As Boolean
        If eliminatedEdges Is Nothing OrElse eliminatedEdges.Count = 0 Then
            Return False
        End If
        Const eps As Double = 0.0001
        For Each e In eliminatedEdges
            If e Is Nothing OrElse e.Length < 4 Then Continue For
            Dim forwardMatch As Boolean =
                Math.Abs(e(0) - p0(0)) < eps AndAlso
                Math.Abs(e(1) - p0(1)) < eps AndAlso
                Math.Abs(e(2) - p1(0)) < eps AndAlso
                Math.Abs(e(3) - p1(1)) < eps
            Dim reverseMatch As Boolean =
                Math.Abs(e(0) - p1(0)) < eps AndAlso
                Math.Abs(e(1) - p1(1)) < eps AndAlso
                Math.Abs(e(2) - p0(0)) < eps AndAlso
                Math.Abs(e(3) - p0(1)) < eps
            If forwardMatch OrElse reverseMatch Then Return True
        Next
        Return False
    End Function

    ' ==================================================================
    '  Step 3: Generate a single band bar .ipt
    ' ==================================================================

    ''' <summary>
    ''' Creates a single band bar .ipt file.
    ''' Geometry convention matches bearing bars:
    '''   XY plane rectangle (Length x Width), extruded +Z by Depth.
    '''   Part +X = bar length, +Y = bar width, +Z = bar depth.
    ''' Band bars receive no notch cuts — bearing bars and cross rods
    ''' are cut back to the inner face of the band bar instead.
    ''' </summary>
    Private Function GenerateSingleBandBar(
            seg As BandBarSegment,
            params As GratingParameters,
            outputFolder As String) As GeneratedBandBarFile

        Dim result As New GeneratedBandBarFile()
        result.SegmentIndex = seg.Index
        result.Mark = seg.Mark
        result.Length = seg.Length
        result.StartPoint = seg.StartPoint
        result.EndPoint = seg.EndPoint
        result.PerpSign = seg.PerpSign

        Dim partDoc As PartDocument = Nothing

        Try
            ' Build file path: {Prefix}_BAND{Index:000}.ipt
            Dim safePfx As String = SanitizeFileName(
                If(params.ResolvedPrefix, "Grating"))
            Dim baseName As String = safePfx & "_BAND" & seg.Index.ToString("000")
            Dim fileName As String = baseName & ".ipt"
            Dim fullPath As String = IO.Path.Combine(outputFolder, fileName)

            ' Avoid overwriting
            If IO.File.Exists(fullPath) Then
                Dim counter As Integer = 1
                Do
                    fileName = baseName & "_" & counter & ".ipt"
                    fullPath = IO.Path.Combine(outputFolder, fileName)
                    counter += 1
                Loop While IO.File.Exists(fullPath) AndAlso counter < 1000
            End If

            result.FilePath = fullPath
            result.FileName = fileName

            ' Convert to cm (Inventor internal units)
            Dim lengthCm As Double = seg.Length * 2.54
            Dim widthCm As Double = params.BarWidth * 2.54
            Dim depthCm As Double = params.BarDepth * 2.54

            ' Create part document
            Dim templatePath As String = GetPartTemplatePath()
            partDoc = CType(
                _app.Documents.Add(DocumentTypeEnum.kPartDocumentObject,
                                    templatePath, False),
                PartDocument)

            Dim compDef As PartComponentDefinition = partDoc.ComponentDefinition
            Dim tg As TransientGeometry = _app.TransientGeometry

            Dim startMiterCm As Double = seg.StartMiterTrim * 2.54
            Dim endMiterCm As Double = seg.EndMiterTrim * 2.54

            ' One info line up front so any subsequent failure trace can
            ' be tied back to the exact input values that produced it.
            Trace.TraceInformation(
                ": HMG BandBarGen: Generating bar " & seg.Mark &
                " — L=" & lengthCm.ToString("F4") &
                "cm W=" & widthCm.ToString("F4") &
                "cm D=" & depthCm.ToString("F4") &
                "cm sMiter=" & startMiterCm.ToString("F4") &
                "cm eMiter=" & endMiterCm.ToString("F4") & "cm")

            ' Sketch on XY: +X = length along edge, +Y = width toward interior,
            ' origin = outer-face corner at segment start (flush perimeter).
            Dim sketch As PlanarSketch =
                compDef.Sketches.Add(compDef.WorkPlanes.Item(3))

            ' Build the mitered (or rectangular-when-both-miters-zero)
            ' trapezoid profile.  If Inventor rejects it — happens on
            ' certain arc-adjacent geometries with E_INVALIDARG — fall
            ' back to a plain rectangle so the bar still generates.
            Dim profile As Profile = Nothing
            Try
                AddBandBarPlanProfile(
                    sketch, tg, lengthCm, widthCm, startMiterCm, endMiterCm)
                profile = sketch.Profiles.AddForSolid()
                Trace.TraceInformation(
                    ": HMG BandBarGen:   mitered profile OK — " & seg.Mark)
            Catch profileEx As Exception
                Trace.TraceWarning(
                    ": HMG BandBarGen:   mitered profile FAILED — " &
                    seg.Mark & " — " & profileEx.Message &
                    ".  Falling back to rectangle.")
                ' Clear any partial sketch geometry before retrying.
                Try
                    For i As Integer = sketch.SketchLines.Count To 1 Step -1
                        sketch.SketchLines.Item(i).Delete()
                    Next
                Catch
                End Try
                sketch.SketchLines.AddAsTwoPointRectangle(
                    tg.CreatePoint2d(0, 0),
                    tg.CreatePoint2d(lengthCm, widthCm))
                profile = sketch.Profiles.AddForSolid()
                result.WarningMessage =
                    "Miter omitted (Inventor rejected the trapezoid profile); " &
                    "bar generated as a plain rectangle."
                Trace.TraceInformation(
                    ": HMG BandBarGen:   rectangle fallback OK — " & seg.Mark)
            End Try

            ' Extrude along +Z by bar depth
            Dim extDef As ExtrudeDefinition =
                compDef.Features.ExtrudeFeatures.CreateExtrudeDefinition(
                    profile, PartFeatureOperationEnum.kNewBodyOperation)
            extDef.SetDistanceExtent(
                depthCm,
                PartFeatureExtentDirectionEnum.kPositiveExtentDirection)
            compDef.Features.ExtrudeFeatures.Add(extDef)
            Trace.TraceInformation(
                ": HMG BandBarGen:   extruded OK — " & seg.Mark)

            ' Set iProperties for downstream identification
            SetCustomProperty(partDoc, "HMG_Type", "BandBar")
            SetCustomProperty(partDoc, "HMG_Mark", seg.Mark)
            SetCustomProperty(partDoc, "HMG_SegmentIndex", seg.Index.ToString())
            SetCustomProperty(partDoc, "HMG_Length_in",
                seg.Length.ToString("F4"))
            SetCustomProperty(partDoc, "HMG_BarWidth_in",
                params.BarWidth.ToString("F4"))
            SetCustomProperty(partDoc, "HMG_BarDepth_in",
                params.BarDepth.ToString("F4"))
            SetCustomProperty(partDoc, "HMG_IsParallel",
                seg.IsParallel.ToString())
            SetCustomProperty(partDoc, "HMG_StartPoint",
                seg.StartPoint(0).ToString("F4") & "," &
                seg.StartPoint(1).ToString("F4"))
            SetCustomProperty(partDoc, "HMG_EndPoint",
                seg.EndPoint(0).ToString("F4") & "," &
                seg.EndPoint(1).ToString("F4"))

            ' Save
            partDoc.SaveAs(fullPath, False)
            result.Saved = True
            Trace.TraceInformation(
                ": HMG BandBarGen:   saved OK — " & seg.Mark &
                " → " & fileName)

        Catch ex As Exception
            result.Saved = False
            result.ErrorMessage = ex.Message
        Finally
            If partDoc IsNot Nothing Then
                Try
                    partDoc.Close(True)
                Catch
                End Try
            End If
        End Try

        Return result
    End Function

    ' ==================================================================
    '  Arc band bar .ipt generation
    ' ==================================================================

    ''' <summary>
    ''' Creates a single curved band bar .ipt file for an arc perimeter edge.
    ''' Geometry: two concentric arcs (inner/outer radius) with radial
    ''' closing lines at start and end angles, extruded +Z by BarDepth.
    ''' Part origin is at the arc centre so assembly placement is a
    ''' simple translation to the centre coordinates.
    ''' </summary>
    Private Function GenerateArcBandBar(
            seg As BandBarSegment,
            params As GratingParameters,
            outputFolder As String) As GeneratedBandBarFile

        Dim result As New GeneratedBandBarFile()
        result.SegmentIndex = seg.Index
        result.Mark = seg.Mark
        result.Length = seg.Length
        result.StartPoint = seg.StartPoint
        result.EndPoint = seg.EndPoint
        result.IsArc = True
        result.ArcCenterX = seg.ArcCenterX
        result.ArcCenterY = seg.ArcCenterY

        Dim partDoc As PartDocument = Nothing

        Try
            ' Build file path
            Dim safePfx As String = SanitizeFileName(
                If(params.ResolvedPrefix, "Grating"))
            Dim baseName As String = safePfx & "_BAND" & seg.Index.ToString("000")
            Dim fileName As String = baseName & ".ipt"
            Dim fullPath As String = IO.Path.Combine(outputFolder, fileName)

            If IO.File.Exists(fullPath) Then
                Dim counter As Integer = 1
                Do
                    fileName = baseName & "_" & counter & ".ipt"
                    fullPath = IO.Path.Combine(outputFolder, fileName)
                    counter += 1
                Loop While IO.File.Exists(fullPath) AndAlso counter < 1000
            End If

            result.FilePath = fullPath
            result.FileName = fileName

            ' Compute inner/outer radii (inches)
            Dim rInner As Double
            Dim rOuter As Double
            If seg.ArcExtendsOutward Then
                ' Cutout arc: bar sits outside the arc radius
                rInner = seg.ArcRadius
                rOuter = seg.ArcRadius + params.BarWidth
            Else
                ' Convex arc: bar sits inside the arc radius
                rInner = seg.ArcRadius - params.BarWidth
                rOuter = seg.ArcRadius
            End If
            If rInner < 0 Then rInner = 0.001

            ' Convert to cm (Inventor internal units)
            Dim rInnerCm As Double = rInner * 2.54
            Dim rOuterCm As Double = rOuter * 2.54
            Dim depthCm As Double = params.BarDepth * 2.54
            Dim entryAngle As Double = seg.ArcEntryAngle
            Dim sweepAngle As Double = seg.ArcSweepAngle

            ' Create part document
            Dim templatePath As String = GetPartTemplatePath()
            partDoc = CType(
                _app.Documents.Add(DocumentTypeEnum.kPartDocumentObject,
                                    templatePath, False),
                PartDocument)

            Dim compDef As PartComponentDefinition = partDoc.ComponentDefinition
            Dim tg As TransientGeometry = _app.TransientGeometry

            ' Sketch on XY plane — origin at arc centre
            Dim sketch As PlanarSketch =
                compDef.Sketches.Add(compDef.WorkPlanes.Item(3))

            ' Centre point (origin)
            Dim centerPt As Point2d = tg.CreatePoint2d(0, 0)

            ' Draw inner arc
            Dim innerArc As SketchArc = sketch.SketchArcs.AddByCenterStartSweepAngle(
                centerPt, rInnerCm, entryAngle, sweepAngle)

            ' Draw outer arc
            Dim outerArc As SketchArc = sketch.SketchArcs.AddByCenterStartSweepAngle(
                centerPt, rOuterCm, entryAngle, sweepAngle)

            ' Closing radial lines at start and end using sketch endpoints
            sketch.SketchLines.AddByTwoPoints(
                innerArc.StartSketchPoint, outerArc.StartSketchPoint)
            sketch.SketchLines.AddByTwoPoints(
                innerArc.EndSketchPoint, outerArc.EndSketchPoint)

            ' Profile and extrude along +Z by bar depth
            Dim profile As Profile = sketch.Profiles.AddForSolid()
            Dim extDef As ExtrudeDefinition =
                compDef.Features.ExtrudeFeatures.CreateExtrudeDefinition(
                    profile, PartFeatureOperationEnum.kNewBodyOperation)
            extDef.SetDistanceExtent(
                depthCm,
                PartFeatureExtentDirectionEnum.kPositiveExtentDirection)
            compDef.Features.ExtrudeFeatures.Add(extDef)

            ' Set iProperties
            SetCustomProperty(partDoc, "HMG_Type", "BandBar_Arc")
            SetCustomProperty(partDoc, "HMG_Mark", seg.Mark)
            SetCustomProperty(partDoc, "HMG_SegmentIndex", seg.Index.ToString())
            SetCustomProperty(partDoc, "HMG_ArcLength_in",
                seg.Length.ToString("F4"))
            SetCustomProperty(partDoc, "HMG_ArcRadius_in",
                seg.ArcRadius.ToString("F4"))
            SetCustomProperty(partDoc, "HMG_ArcSweep_deg",
                (seg.ArcSweepAngle * 180 / Math.PI).ToString("F2"))
            SetCustomProperty(partDoc, "HMG_BarWidth_in",
                params.BarWidth.ToString("F4"))
            SetCustomProperty(partDoc, "HMG_BarDepth_in",
                params.BarDepth.ToString("F4"))
            SetCustomProperty(partDoc, "HMG_ArcCenter",
                seg.ArcCenterX.ToString("F4") & "," &
                seg.ArcCenterY.ToString("F4"))

            ' Save
            partDoc.SaveAs(fullPath, False)
            result.Saved = True

            Trace.TraceInformation(": HMG BandBarGen:   Arc bar " &
                seg.Mark & " saved — R=" &
                seg.ArcRadius.ToString("F4") & """ rInner=" &
                rInner.ToString("F4") & """ rOuter=" &
                rOuter.ToString("F4") & """")

        Catch ex As Exception
            result.Saved = False
            result.ErrorMessage = ex.Message
        Finally
            If partDoc IsNot Nothing Then
                Try
                    partDoc.Close(True)
                Catch
                End Try
            End If
        End Try

        Return result
    End Function

    ' ==================================================================
    '  Helpers
    ' ==================================================================

    ''' <summary>
    ''' Returns the best available Part template file path.
    ''' Same strategy as BearingBarPartGenerator.
    ''' </summary>
    Private Function GetPartTemplatePath() As String
        Try
            Dim p As String = _app.FileManager.GetTemplateFile(
                DocumentTypeEnum.kPartDocumentObject,
                SystemOfMeasureEnum.kEnglishSystemOfMeasure)
            If Not String.IsNullOrEmpty(p) AndAlso IO.File.Exists(p) Then
                Return p
            End If
        Catch
        End Try

        Try
            Dim p As String = _app.FileManager.GetTemplateFile(
                DocumentTypeEnum.kPartDocumentObject,
                SystemOfMeasureEnum.kMetricSystemOfMeasure)
            If Not String.IsNullOrEmpty(p) AndAlso IO.File.Exists(p) Then
                Return p
            End If
        Catch
        End Try

        Try
            Dim templateDir As String =
                _app.DesignProjectManager.ActiveDesignProject.TemplateDir
            If Not String.IsNullOrEmpty(templateDir) AndAlso
               Directory.Exists(templateDir) Then
                For Each candidate As String In {"Standard (in).ipt", "Standard.ipt"}
                    Dim full As String = IO.Path.Combine(templateDir, candidate)
                    If IO.File.Exists(full) Then Return full
                Next
            End If
        Catch
        End Try

        Return ""
    End Function

    ''' <summary>
    ''' Sets a custom iProperty on the part document.
    ''' Creates the property if it does not exist.
    ''' </summary>
    Private Sub SetCustomProperty(doc As PartDocument,
                                  propName As String,
                                  propValue As String)
        Try
            Dim customSet As PropertySet =
                doc.PropertySets.Item("Inventor User Defined Properties")
            Try
                customSet.Item(propName).Value = propValue
            Catch
                customSet.Add(propValue, propName)
            End Try
        Catch
        End Try
    End Sub

    ''' <summary>
    ''' Removes characters that are invalid in file names.
    ''' </summary>
    Private Function SanitizeFileName(name As String) As String
        Dim invalid() As Char = IO.Path.GetInvalidFileNameChars()
        Dim result As New System.Text.StringBuilder()
        For Each c As Char In name
            If Array.IndexOf(invalid, c) < 0 Then
                result.Append(c)
            Else
                result.Append("_"c)
            End If
        Next
        If result.Length = 0 Then Return "Part"
        Return result.ToString()
    End Function

    ''' <summary>
    ''' Draws the band bar plan profile with optional miters at the start
    ''' and/or end.  Outer face lies on Y=0; miters trim the inner (Y=W) side.
    ''' </summary>
    Private Sub AddBandBarPlanProfile(sketch As PlanarSketch,
                                      tg As TransientGeometry,
                                      lengthCm As Double,
                                      widthCm As Double,
                                      startMiterCm As Double,
                                      endMiterCm As Double)
        Const eps As Double = 0.0001

        If startMiterCm < eps AndAlso endMiterCm < eps Then
            sketch.SketchLines.AddAsTwoPointRectangle(
                tg.CreatePoint2d(0, 0),
                tg.CreatePoint2d(lengthCm, widthCm))
            Return
        End If

        Dim maxMiter As Double = Math.Max(eps, lengthCm * 0.49)
        If startMiterCm > maxMiter Then startMiterCm = maxMiter
        If endMiterCm > maxMiter Then endMiterCm = maxMiter

        ' Unified 4-line trapezoid: connect the four corner points in CCW
        ' order with one SketchLine per edge.  Avoids the conditional
        ' "no-miter inserts an extra L-step" branches that previously
        ' produced zero-length segments at the start side when
        ' startMiterCm = 0 (xInnerStart = x0) — Inventor rejected those
        ' with E_FAIL / E_INVALIDARG and the band bar's .ipt never saved.
        '
        '       (x0, y0)──────────────────(xL, y0)        outer (perimeter)
        '          │                          ╲                  end miter
        '          │                       (xInnerEnd, yW)
        '          │                              │
        '       (xInnerStart, yW)──────────────────         inner (interior)
        '          ╲ start miter
        '       (x0, y0)
        '
        ' When startMiterCm = 0: xInnerStart = x0, the "start miter"
        ' line collapses to a vertical (x0, yW) → (x0, y0).
        ' When endMiterCm = 0: xInnerEnd = xL, the "end miter" line
        ' collapses to a vertical (xL, y0) → (xL, yW).
        Dim x0 As Double = 0.0
        Dim y0 As Double = 0.0
        Dim xL As Double = lengthCm
        Dim yW As Double = widthCm
        Dim xInnerStart As Double = startMiterCm
        Dim xInnerEnd As Double = lengthCm - endMiterCm

        ' 1. Outer edge: (x0, y0) → (xL, y0)
        sketch.SketchLines.AddByTwoPoints(
            tg.CreatePoint2d(x0, y0),
            tg.CreatePoint2d(xL, y0))

        ' 2. End side (miter or vertical): (xL, y0) → (xInnerEnd, yW)
        sketch.SketchLines.AddByTwoPoints(
            tg.CreatePoint2d(xL, y0),
            tg.CreatePoint2d(xInnerEnd, yW))

        ' 3. Inner edge: (xInnerEnd, yW) → (xInnerStart, yW)
        sketch.SketchLines.AddByTwoPoints(
            tg.CreatePoint2d(xInnerEnd, yW),
            tg.CreatePoint2d(xInnerStart, yW))

        ' 4. Start side (miter or vertical): (xInnerStart, yW) → (x0, y0)
        sketch.SketchLines.AddByTwoPoints(
            tg.CreatePoint2d(xInnerStart, yW),
            tg.CreatePoint2d(x0, y0))
    End Sub

    ' ==================================================================
    '  Internal segment model
    ' ==================================================================

    ''' <summary>
    ''' Internal representation of a perimeter edge segment used during
    ''' band bar computation. Not exposed outside the generator.
    ''' </summary>
    Private Class BandBarSegment
        Public Property Index As Integer
        Public Property Mark As String
        Public Property StartPoint As Double()
        Public Property EndPoint As Double()
        Public Property Length As Double
        Public Property IsParallel As Boolean

        ''' <summary>+1 for CCW perimeter, -1 for CW (matches ComputeSegments).</summary>
        Public Property PerpSign As Double

        ''' <summary>Plan-view miter trim at start (inches), 0 = square end.</summary>
        Public Property StartMiterTrim As Double

        ''' <summary>Plan-view miter trim at end (inches), 0 = square end.</summary>
        Public Property EndMiterTrim As Double

        ' Arc-specific properties
        Public Property IsArc As Boolean
        Public Property ArcCenterX As Double
        Public Property ArcCenterY As Double
        Public Property ArcRadius As Double
        Public Property ArcEntryAngle As Double
        Public Property ArcSweepAngle As Double
        Public Property ArcExtendsOutward As Boolean

        Public Overrides Function ToString() As String
            If IsArc Then
                Return Mark & ": ARC R=" & ArcRadius.ToString("F4") & """" &
                       " sweep=" & (ArcSweepAngle * 180 / Math.PI).ToString("F1") & "°" &
                       " center=(" & ArcCenterX.ToString("F2") & "," &
                       ArcCenterY.ToString("F2") & ")"
            End If
            Return Mark & ": L=" & Length.ToString("F4") & """" &
                   " (" & StartPoint(0).ToString("F2") & "," &
                   StartPoint(1).ToString("F2") & ") -> (" &
                   EndPoint(0).ToString("F2") & "," &
                   EndPoint(1).ToString("F2") & ")"
        End Function
    End Class

End Class
