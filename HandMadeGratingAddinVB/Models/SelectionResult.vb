'////////////////////////////////////////////////////////////////////
' Metal Bar Grating Addin - VB.NET
'
' SelectionResult: Result type for the perimeter selection workflow.
' Carries either a successful PerimeterData or an error message.
'////////////////////////////////////////////////////////////////////

''' <summary>
''' Encapsulates the outcome of a perimeter selection attempt.
''' </summary>
Public Class SelectionResult

    ''' <summary>True if a valid perimeter was selected and extracted.</summary>
    Public Property Success As Boolean

    ''' <summary>Human-readable error message when Success is False.</summary>
    Public Property ErrorMessage As String

    ''' <summary>Extracted perimeter data when Success is True.</summary>
    Public Property Perimeter As PerimeterData

    Public Shared Function Succeeded(data As PerimeterData) As SelectionResult
        Return New SelectionResult With {
            .Success = True,
            .Perimeter = data
        }
    End Function

    Public Shared Function Failed(message As String) As SelectionResult
        Return New SelectionResult With {
            .Success = False,
            .ErrorMessage = message
        }
    End Function

End Class
