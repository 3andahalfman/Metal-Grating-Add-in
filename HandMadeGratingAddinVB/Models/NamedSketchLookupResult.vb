'////////////////////////////////////////////////////////////////////
' Metal Bar Grating Addin - VB.NET
'
' NamedSketchLookupResult: Carries the outcome of a named boundary
' sketch search performed by BoundarySourceService before the
' Boundary Source Dialog is shown.
'
' Phase 12: Named sketch boundary source.
'////////////////////////////////////////////////////////////////////

Imports Inventor

''' <summary>
''' Result of a search for a named grating boundary sketch in the
''' active Part document. Produced by BoundarySourceService and
''' consumed by BoundarySourceDialog and GratingCommand.
''' </summary>
Public Class NamedSketchLookupResult

    ''' <summary>True if a matching sketch was found.</summary>
    Public Property Found As Boolean

    ''' <summary>
    ''' The exact name of the sketch that was found.
    ''' Nothing when Found is False.
    ''' </summary>
    Public Property SketchName As String

    ''' <summary>
    ''' Transient COM reference to the found sketch.
    ''' Nothing when Found is False. Valid only within the current
    ''' Inventor session — do not store long-term.
    ''' </summary>
    Public Property Sketch As PlanarSketch

    ''' <summary>Human-readable message for trace output.</summary>
    Public Property Message As String

    ' --- Factories ---

    ''' <summary>Creates a not-found result with a diagnostic message.</summary>
    Public Shared Function NotFound(message As String) As NamedSketchLookupResult
        Return New NamedSketchLookupResult With {
            .Found = False,
            .Message = message
        }
    End Function

    ''' <summary>Creates a found result.</summary>
    Public Shared Function Succeeded(sketchName As String,
                                     sketch As PlanarSketch) As NamedSketchLookupResult
        Return New NamedSketchLookupResult With {
            .Found = True,
            .SketchName = sketchName,
            .Sketch = sketch,
            .Message = "Found named boundary sketch: " & sketchName
        }
    End Function

End Class
