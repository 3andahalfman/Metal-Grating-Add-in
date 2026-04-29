'////////////////////////////////////////////////////////////////////
' Metal Bar Grating Addin - VB.NET
'
' GeneratedCrossBarFile: Records the output of a single cross bar
' part generation — the saved file path and source cross bar data.
'////////////////////////////////////////////////////////////////////

''' <summary>
''' Tracks one generated cross bar .ipt file.
''' </summary>
Public Class GeneratedCrossBarFile

    ''' <summary>The source cross bar entry from the layout.</summary>
    Public Property SourceEntry As CrossBarEntry

    ''' <summary>Full path of the saved .ipt file.</summary>
    Public Property FilePath As String

    ''' <summary>File name only (no directory).</summary>
    Public Property FileName As String

    ''' <summary>True if the file was saved successfully.</summary>
    Public Property Saved As Boolean

    ''' <summary>Error message if Saved is False.</summary>
    Public Property ErrorMessage As String

    ''' <summary>
    ''' Number of identical cross bars at this length.
    ''' When parts are grouped by unique length, this indicates
    ''' how many cross bars share this .ipt geometry.
    ''' </summary>
    Public Property Quantity As Integer

    Public Overrides Function ToString() As String
        If Saved Then
            Return SourceEntry.Mark & " -> " & FileName &
                   " (L=" & SourceEntry.Length.ToString("F4") &
                   ", qty=" & Quantity & ")"
        Else
            Return SourceEntry.Mark & " FAILED: " & ErrorMessage
        End If
    End Function

End Class
