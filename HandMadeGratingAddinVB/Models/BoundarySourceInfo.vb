'////////////////////////////////////////////////////////////////////
' Metal Bar Grating Addin - VB.NET
'
' BoundarySourceInfo: Records how the grating boundary perimeter
' was obtained — imported DWG, named sketch, selected sketch, or
' newly created sketch. Stored on GratingProject for re-open logic.
'////////////////////////////////////////////////////////////////////

''' <summary>
''' How the grating boundary perimeter was sourced.
''' </summary>
Public Enum BoundarySourceType
    ''' <summary>User pointed to an existing sketch named GRATING_BOUNDARY.</summary>
    NamedSketch = 0
    ''' <summary>Add-in imported a DWG/DXF file into a new sketch.</summary>
    ImportedDwg = 1
    ''' <summary>Add-in created a fresh empty sketch for the user to draw in.</summary>
    NewSketch = 2
    ''' <summary>User selected an arbitrary existing sketch (legacy workflow).</summary>
    SelectedSketch = 3
End Enum

''' <summary>
''' Metadata describing how the grating perimeter was sourced.
''' </summary>
Public Class BoundarySourceInfo

    ''' <summary>Strategy used to obtain the boundary.</summary>
    Public Property SourceType As BoundarySourceType

    ''' <summary>Name of the sketch that holds the boundary geometry.</summary>
    Public Property SketchName As String

    ''' <summary>
    ''' Original file path when SourceType is ImportedDwg. Nothing otherwise.
    ''' </summary>
    Public Property ImportedFilePath As String

    ''' <summary>UTC timestamp when the boundary was established.</summary>
    Public Property CreatedUtc As DateTime

    ' --- Factory helpers ---

    ''' <summary>Creates info for a legacy selected-sketch workflow.</summary>
    Public Shared Function FromSelectedSketch(sketchName As String) As BoundarySourceInfo
        Return New BoundarySourceInfo With {
            .SourceType = BoundarySourceType.SelectedSketch,
            .SketchName = sketchName,
            .CreatedUtc = DateTime.UtcNow
        }
    End Function

    ''' <summary>Creates info for a named GRATING_BOUNDARY sketch.</summary>
    Public Shared Function FromNamedSketch(sketchName As String) As BoundarySourceInfo
        Return New BoundarySourceInfo With {
            .SourceType = BoundarySourceType.NamedSketch,
            .SketchName = sketchName,
            .CreatedUtc = DateTime.UtcNow
        }
    End Function

    ''' <summary>Creates info for an imported DWG/DXF file.</summary>
    Public Shared Function FromImportedDwg(sketchName As String, filePath As String) As BoundarySourceInfo
        Return New BoundarySourceInfo With {
            .SourceType = BoundarySourceType.ImportedDwg,
            .SketchName = sketchName,
            .ImportedFilePath = filePath,
            .CreatedUtc = DateTime.UtcNow
        }
    End Function

    ''' <summary>Creates info for a newly created empty sketch.</summary>
    Public Shared Function FromNewSketch(sketchName As String) As BoundarySourceInfo
        Return New BoundarySourceInfo With {
            .SourceType = BoundarySourceType.NewSketch,
            .SketchName = sketchName,
            .CreatedUtc = DateTime.UtcNow
        }
    End Function

    Public Overrides Function ToString() As String
        Select Case SourceType
            Case BoundarySourceType.ImportedDwg
                Return "Imported from " & If(ImportedFilePath, "DWG") & " → " & SketchName
            Case BoundarySourceType.NamedSketch
                Return "Named sketch: " & SketchName
            Case BoundarySourceType.NewSketch
                Return "New sketch: " & SketchName
            Case Else
                Return "Selected sketch: " & SketchName
        End Select
    End Function

End Class
