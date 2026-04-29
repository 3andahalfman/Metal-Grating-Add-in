'////////////////////////////////////////////////////////////////////
' Metal Bar Grating Addin - VB.NET
'
' PerimeterExtractor: Extracts perimeter geometry data from a
' validated Inventor sketch profile path into the internal
' PerimeterData model.  Arc entities are tessellated into
' polyline segments so the downstream polygon-intersection
' algorithms handle curved boundaries correctly.
'
' Uses late-bound property access (Option Strict Off) for robust
' COM interop with varying sketch entity types.
'
' NOTE: Inventor's internal coordinate system is centimetres.
'       All extracted vertices are converted to inches to match
'       the grating parameter unit convention (1 in = 2.54 cm).
'////////////////////////////////////////////////////////////////////

Imports System.Diagnostics
Imports Inventor

''' <summary>
''' Extracts vertex data from a validated outer profile path.
''' </summary>
Public Class PerimeterExtractor

    ''' <summary>Conversion factor: centimetres to inches.</summary>
    Private Const CmPerInch As Double = 2.54

    ''' <summary>
    ''' Extracts perimeter data from a pre-validated outer profile path.
    ''' </summary>
    ''' <param name="sketch">Source sketch (for metadata).</param>
    ''' <param name="outerPath">The outer boundary ProfilePath.</param>
    ''' <returns>PerimeterData, or Nothing if extraction fails.</returns>
    Public Function ExtractFromPath(sketch As PlanarSketch,
                                    outerPath As ProfilePath) As PerimeterData
        Try
            Dim arcInfos As New List(Of PerimeterArcInfo)
            Dim vertices As List(Of Double()) = ExtractVertices(outerPath, arcInfos)
            Dim edgeCount As Integer = SafePathCount(outerPath)

            Dim data As New PerimeterData()
            data.SketchName = sketch.Name
            data.EdgeCount = edgeCount
            data.OuterLoopVertices = vertices
            data.ArcSegments = arcInfos
            data.SourceSketch = sketch

            Trace.TraceInformation(": HMG Extractor: Extracted " &
                edgeCount & " edges, " & vertices.Count & " vertices, " &
                arcInfos.Count & " arc(s).")

            Return data
        Catch ex As Exception
            Trace.TraceError(": HMG Extractor: Extraction failed: " &
                             ex.Message)
            Return Nothing
        End Try
    End Function

#Region "Vertex extraction"

    ''' <summary>
    ''' Tolerance (in cm) for matching shared endpoints between consecutive
    ''' profile entities.  Inventor stores sketch coordinates in cm with
    ''' high precision, but DWG-imported geometry can carry tiny rounding
    ''' errors at the connection points.
    ''' </summary>
    Private Const ConnectionToleranceCm As Double = 0.001

    ''' <summary>
    ''' Walks each ProfileEntity in the path and extracts vertex coordinates.
    ''' Arc metadata is recorded in the arcInfos list.
    '''
    ''' Two-pass algorithm:
    '''   1. Pre-read every entity's start/end endpoints (cm).
    '''   2. Determine each entity's traversal direction by checking which
    '''      endpoint connects to the previous entity's exit point.  The
    '''      first entity's direction is determined by lookahead to entity 2.
    '''   3. Emit entry vertices in path order using the determined
    '''      direction.  Arcs are tessellated with the matching sweep sign.
    '''
    ''' Replaces the previous "compare each entity's endpoints to the last
    ''' emitted vertex" heuristic, which produced wrong vertex order when
    '''   • the first entity in the path was geometrically reversed
    '''     (caused a quadrilateral to collapse into a triangle), or
    '''   • the polygon was concave (closer-of-two endpoints picked the
    '''     wrong vertex on inset/notch edges).
    ''' </summary>
    Private Function ExtractVertices(outerPath As ProfilePath,
                                     arcInfos As List(Of PerimeterArcInfo)) As List(Of Double())
        Dim vertices As New List(Of Double())

        Try
            Dim count As Integer = outerPath.Count
            If count = 0 Then Return vertices

            ' --- Pass 1: collect endpoints of every entity (in cm) ---
            Dim startsX(count - 1) As Double
            Dim startsY(count - 1) As Double
            Dim endsX(count - 1) As Double
            Dim endsY(count - 1) As Double
            Dim hasEnds(count - 1) As Boolean

            For i As Integer = 0 To count - 1
                Try
                    Dim sketchEnt As Object = outerPath.Item(i + 1).SketchEntity
                    Dim sg As Object = sketchEnt.StartSketchPoint.Geometry
                    Dim eg As Object = sketchEnt.EndSketchPoint.Geometry
                    startsX(i) = CDbl(sg.X) : startsY(i) = CDbl(sg.Y)
                    endsX(i) = CDbl(eg.X) : endsY(i) = CDbl(eg.Y)
                    hasEnds(i) = True
                Catch
                    hasEnds(i) = False
                End Try
            Next

            ' --- Pass 2: determine each entity's traversal direction ---
            Dim isReversed(count - 1) As Boolean

            ' First entity: direction determined by lookahead to entity 2.
            ' If e0.End connects to either endpoint of e1 → forward direction.
            ' Else if e0.Start connects → reversed.
            If count >= 2 AndAlso hasEnds(0) AndAlso hasEnds(1) Then
                Dim endE0ConnectsE1 As Boolean =
                    PointsClose(endsX(0), endsY(0), startsX(1), startsY(1)) OrElse
                    PointsClose(endsX(0), endsY(0), endsX(1), endsY(1))
                Dim startE0ConnectsE1 As Boolean =
                    PointsClose(startsX(0), startsY(0), startsX(1), startsY(1)) OrElse
                    PointsClose(startsX(0), startsY(0), endsX(1), endsY(1))

                If startE0ConnectsE1 AndAlso Not endE0ConnectsE1 Then
                    isReversed(0) = True
                Else
                    isReversed(0) = False
                End If
            Else
                isReversed(0) = False
            End If

            ' Subsequent entities: match to previous entity's exit point.
            For i As Integer = 1 To count - 1
                If Not hasEnds(i) OrElse Not hasEnds(i - 1) Then
                    isReversed(i) = False
                    Continue For
                End If

                Dim prevExitX As Double, prevExitY As Double
                If isReversed(i - 1) Then
                    prevExitX = startsX(i - 1) : prevExitY = startsY(i - 1)
                Else
                    prevExitX = endsX(i - 1) : prevExitY = endsY(i - 1)
                End If

                Dim startMatches As Boolean =
                    PointsClose(startsX(i), startsY(i), prevExitX, prevExitY)
                Dim endMatches As Boolean =
                    PointsClose(endsX(i), endsY(i), prevExitX, prevExitY)

                If startMatches AndAlso Not endMatches Then
                    isReversed(i) = False
                ElseIf endMatches AndAlso Not startMatches Then
                    isReversed(i) = True
                ElseIf startMatches Then
                    ' Both match (degenerate) — keep forward
                    isReversed(i) = False
                Else
                    ' Disconnected — choose the closer endpoint
                    Dim dStart As Double =
                        (startsX(i) - prevExitX) * (startsX(i) - prevExitX) +
                        (startsY(i) - prevExitY) * (startsY(i) - prevExitY)
                    Dim dEnd As Double =
                        (endsX(i) - prevExitX) * (endsX(i) - prevExitX) +
                        (endsY(i) - prevExitY) * (endsY(i) - prevExitY)
                    isReversed(i) = dEnd < dStart
                End If
            Next

            ' --- Pass 3: emit vertices in path order using determined directions ---
            For i As Integer = 0 To count - 1
                Dim entity As ProfileEntity = outerPath.Item(i + 1)
                ExtractEntityVertex(entity, vertices, arcInfos, isReversed(i))
            Next
        Catch ex As Exception
            Trace.TraceWarning(": HMG Extractor: Vertex walk partial: " &
                               ex.Message)
        End Try

        Return vertices
    End Function

    ''' <summary>
    ''' Returns True if the two cm-space points are within
    ''' <see cref="ConnectionToleranceCm"/> of each other.
    ''' </summary>
    Private Shared Function PointsClose(ax As Double, ay As Double,
                                         bx As Double, by As Double) As Boolean
        Return Math.Abs(ax - bx) < ConnectionToleranceCm AndAlso
               Math.Abs(ay - by) < ConnectionToleranceCm
    End Function

    ''' <summary>
    ''' Extracts the entry vertex from a single profile entity using a
    ''' pre-determined traversal direction.  For arcs, emits tessellated
    ''' points along the curve; for lines/splines, emits a single vertex.
    '''
    ''' When <paramref name="reversed"/> is True, the entity's geometric
    ''' EndSketchPoint is the path entry; otherwise the StartSketchPoint is.
    ''' Coordinates are converted from Inventor's internal cm to inches.
    ''' </summary>
    Private Sub ExtractEntityVertex(entity As ProfileEntity,
                                    vertices As List(Of Double()),
                                    arcInfos As List(Of PerimeterArcInfo),
                                    reversed As Boolean)
        Dim sketchEnt As Object = Nothing
        Try
            sketchEnt = entity.SketchEntity
        Catch
            Trace.TraceWarning(": HMG Extractor: Could not access " &
                               "SketchEntity — skipped.")
            Return
        End Try

        ' Try arc tessellation (emits multiple vertices along the curve)
        If TryTessellateArc(sketchEnt, vertices, arcInfos, reversed) Then Return

        ' Line / spline: emit start or end based on the determined direction.
        Try
            Dim startGeom As Object = sketchEnt.StartSketchPoint.Geometry
            Dim sx As Double = CDbl(startGeom.X)
            Dim sy As Double = CDbl(startGeom.Y)

            Dim hasEnd As Boolean = False
            Dim ex As Double = 0
            Dim ey As Double = 0
            Try
                Dim endGeom As Object = sketchEnt.EndSketchPoint.Geometry
                ex = CDbl(endGeom.X)
                ey = CDbl(endGeom.Y)
                hasEnd = True
            Catch
            End Try

            If reversed AndAlso hasEnd Then
                vertices.Add(New Double() {ex / CmPerInch, ey / CmPerInch})
            Else
                vertices.Add(New Double() {sx / CmPerInch, sy / CmPerInch})
            End If
            Return
        Catch
        End Try

        ' Fallback: CenterSketchPoint (full circles, ellipses)
        Try
            Dim centerPt As Object = sketchEnt.CenterSketchPoint
            Dim geom As Object = centerPt.Geometry
            vertices.Add(New Double() {CDbl(geom.X) / CmPerInch,
                                       CDbl(geom.Y) / CmPerInch})
            Return
        Catch
        End Try

        Trace.TraceWarning(": HMG Extractor: Could not extract vertex " &
                           "from entity at index — skipped.")
    End Sub

    ''' <summary>
    ''' Attempts to tessellate an arc entity into multiple polygon vertices.
    ''' Emits the entry point and intermediate points along the arc, but
    ''' NOT the exit point (which is the next entity's start point).
    ''' Uses the supplied <paramref name="reversed"/> flag to decide
    ''' which geometric endpoint is the path entry and the corresponding
    ''' sweep sign.
    ''' Returns True if the entity was recognized as an arc.
    ''' </summary>
    Private Function TryTessellateArc(sketchEnt As Object,
                                      vertices As List(Of Double()),
                                      arcInfos As List(Of PerimeterArcInfo),
                                      reversed As Boolean) As Boolean
        Try
            ' Access arc-specific properties — throws for non-arcs
            Dim centerGeom As Object = sketchEnt.CenterSketchPoint.Geometry
            Dim cx As Double = CDbl(centerGeom.X)
            Dim cy As Double = CDbl(centerGeom.Y)

            Dim arcGeom As Object = sketchEnt.Geometry
            Dim sweepAngle As Double = CDbl(arcGeom.SweepAngle)

            Dim startGeom As Object = sketchEnt.StartSketchPoint.Geometry
            Dim endGeom As Object = sketchEnt.EndSketchPoint.Geometry
            Dim sx As Double = CDbl(startGeom.X)
            Dim sy As Double = CDbl(startGeom.Y)
            Dim ex As Double = CDbl(endGeom.X)
            Dim ey As Double = CDbl(endGeom.Y)

            Dim radius As Double = Math.Sqrt((sx - cx) * (sx - cx) +
                                              (sy - cy) * (sy - cy))
            If radius < 0.0001 Then Return False

            ' Direction was pre-determined in the calling pass.
            Dim entryAngle As Double
            Dim sweep As Double

            If reversed Then
                ' Path enters at the geometric end — reverse sweep.
                entryAngle = Math.Atan2(ey - cy, ex - cx)
                sweep = -sweepAngle
            Else
                entryAngle = Math.Atan2(sy - cy, sx - cx)
                sweep = sweepAngle
            End If

            ' ~5° per segment (min 4, max 72)
            Dim absSweep As Double = Math.Abs(sweep)
            Dim segments As Integer = Math.Max(4,
                CInt(Math.Ceiling(absSweep / (Math.PI / 36.0))))
            segments = Math.Min(segments, 72)

            Dim stepAngle As Double = sweep / segments

            ' Record arc metadata before emitting vertices
            Dim arcInfo As New PerimeterArcInfo()
            arcInfo.CenterX = cx / CmPerInch
            arcInfo.CenterY = cy / CmPerInch
            arcInfo.Radius = radius / CmPerInch
            arcInfo.EntryAngle = entryAngle
            arcInfo.SweepAngle = sweep
            arcInfo.FirstVertexIndex = vertices.Count
            arcInfo.VertexCount = segments
            arcInfos.Add(arcInfo)

            ' Emit entry point + intermediate points (NOT exit —
            ' exit is the next entity's start point)
            For s As Integer = 0 To segments - 1
                Dim angle As Double = entryAngle + s * stepAngle
                Dim px As Double = cx + radius * Math.Cos(angle)
                Dim py As Double = cy + radius * Math.Sin(angle)
                vertices.Add(New Double() {px / CmPerInch, py / CmPerInch})
            Next

            Trace.TraceInformation(": HMG Extractor: Tessellated arc — " &
                segments & " segments, radius=" &
                (radius / CmPerInch).ToString("F4") & """")

            Return True
        Catch
            Return False
        End Try
    End Function

#End Region

#Region "Helpers"

    Private Shared Function SafePathCount(path As ProfilePath) As Integer
        Try
            Return path.Count
        Catch
            Return 0
        End Try
    End Function

#End Region

End Class
