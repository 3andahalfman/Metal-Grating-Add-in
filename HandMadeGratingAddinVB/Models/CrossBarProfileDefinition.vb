'////////////////////////////////////////////////////////////////////
' Metal Bar Grating Addin - VB.NET
'
' CrossBarProfileDefinition: Describes a single standard cross bar
' profile entry in the profile library.  Each entry maps a
' CrossBarType + SurfaceProfileType combination to a real DWG source
' file and carries the physical dimensions needed for later geometry
' generation.
'
' Phase 15: Cross bar profile library and resolver.
'////////////////////////////////////////////////////////////////////

Imports System.IO

''' <summary>
''' Broad shape category for a cross bar profile.
''' </summary>
Public Enum CrossBarShapeFamily
    ''' <summary>Solid round rod (circular cross-section).</summary>
    Round = 0
    ''' <summary>Rectangular flat bar (rectangular cross-section).</summary>
    Rectangular = 1
End Enum

''' <summary>
''' A single entry in the cross bar profile library.
''' Ties a user selection (CrossBarType + SurfaceProfileType) to a
''' physical profile definition and the DWG source file that carries
''' its cross-section geometry.
''' </summary>
Public Class CrossBarProfileDefinition

    ' --- Identity ---

    ''' <summary>Human-readable name shown in logs and future UI.</summary>
    Public Property DisplayName As String

    ''' <summary>Cross bar type this definition covers.</summary>
    Public Property CrossBar As CrossBarType

    ''' <summary>Surface treatment this definition covers.</summary>
    Public Property SurfaceProfile As SurfaceProfileType

    ''' <summary>Shape family (Round or Rectangular).</summary>
    Public Property ShapeFamily As CrossBarShapeFamily

    ' --- Physical dimensions (inches) ---

    ''' <summary>
    ''' Nominal diameter in inches.  Meaningful only when
    ''' ShapeFamily is Round; 0.0 for Rectangular profiles.
    ''' </summary>
    Public Property DiameterInches As Double

    ''' <summary>
    ''' Bar thickness in inches (the short cross-section dimension).
    ''' Meaningful only when ShapeFamily is Rectangular; 0.0 for Round.
    ''' </summary>
    Public Property ThicknessInches As Double

    ''' <summary>
    ''' Bar height in inches (the tall cross-section dimension).
    ''' Meaningful only when ShapeFamily is Rectangular; 0.0 for Round.
    ''' </summary>
    Public Property HeightInches As Double

    ' --- Source file ---

    ''' <summary>
    ''' File name of the DWG source file (name and extension only).
    ''' Example: "3-8 RND PLAIN.dwg"
    ''' </summary>
    Public Property SourceFileName As String

    ''' <summary>
    ''' Absolute path to the DWG source file once resolved by the
    ''' library.  Empty string if the file has not yet been located.
    ''' </summary>
    Public Property SourceFilePath As String

    ' --- Computed helpers ---

    ''' <summary>
    ''' True when SourceFilePath has been set and the file exists on disk.
    ''' </summary>
    Public ReadOnly Property IsFileFound As Boolean
        Get
            Return Not String.IsNullOrEmpty(SourceFilePath) AndAlso
                   File.Exists(SourceFilePath)
        End Get
    End Property

    ''' <summary>
    ''' Compact description for trace logging.
    ''' Example: "3/8"" Round Plain  →  3-8 RND PLAIN.dwg (found)"
    ''' </summary>
    Public Overrides Function ToString() As String
        Dim filePart As String
        If Not String.IsNullOrEmpty(SourceFilePath) Then
            filePart = SourceFileName &
                       If(IsFileFound, " (found)", " (NOT FOUND)")
        ElseIf Not String.IsNullOrEmpty(SourceFileName) Then
            filePart = SourceFileName & " (path not resolved)"
        Else
            filePart = "(no source file)"
        End If
        Return DisplayName & "  →  " & filePart
    End Function

End Class
