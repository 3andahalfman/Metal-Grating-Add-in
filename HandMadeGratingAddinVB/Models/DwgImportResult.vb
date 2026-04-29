'////////////////////////////////////////////////////////////////////
' Metal Bar Grating Addin - VB.NET
'
' DwgImportResult: Outcome of a DWG/DXF boundary import performed
' by DwgImportService. Carries the name of the sketch that was
' created or replaced in the active Part document, plus diagnostic
' metadata consumed by GratingCommand and the trace log.
'
' Phase 13: DWG/DXF boundary import.
'////////////////////////////////////////////////////////////////////

''' <summary>
''' Outcome of a DWG/DXF boundary import.
''' Created by DwgImportService; consumed by GratingCommand.
''' </summary>
Public Class DwgImportResult

    ''' <summary>True when the import succeeded and a sketch was created.</summary>
    Public Property Success As Boolean

    ''' <summary>
    ''' Name of the Part sketch that holds the imported boundary.
    ''' Always BoundarySourceService.PrimaryName ("GRATING_BOUNDARY")
    ''' when Success is True.
    ''' </summary>
    Public Property SketchName As String

    ''' <summary>Full path of the source DWG/DXF file.</summary>
    Public Property SourceFilePath As String

    ''' <summary>
    ''' Number of line segments written into the boundary sketch.
    ''' Meaningful only when Success is True.
    ''' </summary>
    Public Property SegmentCount As Integer

    ''' <summary>Human-readable error detail. Empty when Success is True.</summary>
    Public Property ErrorMessage As String

    ' --- Factories ---

    ''' <summary>Creates a successful import result.</summary>
    Public Shared Function Succeeded(sketchName As String,
                                     filePath As String,
                                     segmentCount As Integer) As DwgImportResult
        Return New DwgImportResult With {
            .Success = True,
            .SketchName = sketchName,
            .SourceFilePath = filePath,
            .SegmentCount = segmentCount,
            .ErrorMessage = String.Empty
        }
    End Function

    ''' <summary>Creates a failed import result with a diagnostic message.</summary>
    Public Shared Function Failed(message As String,
                                  Optional filePath As String = "") As DwgImportResult
        Return New DwgImportResult With {
            .Success = False,
            .ErrorMessage = message,
            .SourceFilePath = filePath
        }
    End Function

    Public Overrides Function ToString() As String
        If Success Then
            Return "DWG import OK — " & SegmentCount & " segments → " & SketchName
        Else
            Return "DWG import FAILED — " & ErrorMessage
        End If
    End Function

End Class
