'////////////////////////////////////////////////////////////////////
' Metal Bar Grating Addin - VB.NET
'
' GeneratedBearingBarFile: Records the output of a single bearing
' bar part generation — the saved file path and source bar data.
'////////////////////////////////////////////////////////////////////

''' <summary>
''' Tracks one generated .ipt file created from a TrimmedBearingBar.
''' </summary>
Public Class GeneratedBearingBarFile

    ''' <summary>The source bar from the layout engine.</summary>
    Public Property SourceBar As TrimmedBearingBar

    ''' <summary>Full path of the saved .ipt file.</summary>
    Public Property FilePath As String

    ''' <summary>File name only (no directory).</summary>
    Public Property FileName As String

    ''' <summary>True if the file was saved successfully.</summary>
    Public Property Saved As Boolean

    ''' <summary>Error message if Saved is False.</summary>
    Public Property ErrorMessage As String

    ''' <summary>Number of notch slots cut in this bar (Phase 9).</summary>
    Public Property NotchCount As Integer

    Public Overrides Function ToString() As String
        If Saved Then
            Dim notchInfo As String = If(NotchCount > 0,
                " (" & NotchCount & " notches)", "")
            Return SourceBar.Mark & " -> " & FileName & notchInfo
        Else
            Return SourceBar.Mark & " FAILED: " & ErrorMessage
        End If
    End Function

End Class
