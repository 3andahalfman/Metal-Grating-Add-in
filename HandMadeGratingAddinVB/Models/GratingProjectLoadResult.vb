'////////////////////////////////////////////////////////////////////
' Metal Bar Grating Addin - VB.NET
'
' GratingProjectLoadResult: Outcome of GratingProjectStorage.TryLoad.
' Carries the reconstructed GratingProject metadata, the deserialized
' GratingParameters (used as form defaults even when boundary cannot
' be auto-resolved), and a flag indicating whether the boundary sketch
' was successfully re-resolved.
'
' Phase 14: Project persistence via Inventor iProperties.
'////////////////////////////////////////////////////////////////////

''' <summary>
''' Outcome of a GratingProjectStorage.TryLoad call.
''' </summary>
Public Class GratingProjectLoadResult

    ''' <summary>
    ''' True when HMG iProperties were found on the document and at least
    ''' the project name and boundary source type could be read.
    ''' </summary>
    Public Property Found As Boolean

    ''' <summary>
    ''' The partially-reconstructed GratingProject.
    ''' <list type="bullet">
    '''   <item>When <see cref="BoundaryResolved"/> is True the project
    '''         contains both metadata and a valid Perimeter.</item>
    '''   <item>When False the Perimeter is Nothing; metadata and
    '''         <see cref="SavedParameters"/> are still usable.</item>
    ''' </list>
    ''' Nothing when <see cref="Found"/> is False.
    ''' </summary>
    Public Property Project As GratingProject

    ''' <summary>
    ''' True when the saved boundary sketch was located in the active
    ''' Part document and its perimeter successfully extracted.
    ''' </summary>
    Public Property BoundaryResolved As Boolean

    ''' <summary>
    ''' The saved GratingParameters deserialized from iProperties.
    ''' Use as the defaults for GratingInputForm even if the boundary
    ''' could not be re-resolved.  Nothing when Found is False.
    ''' </summary>
    Public Property SavedParameters As GratingParameters

    ''' <summary>Human-readable diagnostic string for the trace log.</summary>
    Public Property Message As String

    ' --- Factories ---

    ''' <summary>Returns a result indicating no HMG data on the document.</summary>
    Public Shared Function NotFound() As GratingProjectLoadResult
        Return New GratingProjectLoadResult With {
            .Found = False,
            .BoundaryResolved = False,
            .Message = "No HMG iProperties found on document."
        }
    End Function

    ''' <summary>
    ''' Returns a successful load result.
    ''' </summary>
    ''' <param name="project">Reconstructed project (Perimeter set only when
    ''' boundaryResolved is True).</param>
    ''' <param name="savedParameters">Parameters deserialized from iProperties.</param>
    ''' <param name="boundaryResolved">True if the boundary sketch was located
    ''' and its perimeter extracted.</param>
    ''' <param name="message">Diagnostic string.</param>
    Public Shared Function Loaded(project As GratingProject,
                                  savedParameters As GratingParameters,
                                  boundaryResolved As Boolean,
                                  message As String) As GratingProjectLoadResult
        Return New GratingProjectLoadResult With {
            .Found = True,
            .Project = project,
            .SavedParameters = savedParameters,
            .BoundaryResolved = boundaryResolved,
            .Message = message
        }
    End Function

    Public Overrides Function ToString() As String
        If Not Found Then Return "Not found"
        Dim status As String = If(BoundaryResolved, "boundary resolved", "boundary NOT resolved")
        Return "Loaded — " & status & " — " & Message
    End Function

End Class
