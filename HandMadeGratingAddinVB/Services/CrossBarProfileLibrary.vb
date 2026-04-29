'////////////////////////////////////////////////////////////////////
' Metal Bar Grating Addin - VB.NET
'
' CrossBarProfileLibrary: Static catalogue of all standard cross bar
' profiles and a resolver that maps a user's CrossBarType +
' SurfaceProfileType selection to the correct CrossBarProfileDefinition.
'
' Phase 15: Cross bar profile library and resolver.
'
' Source files
' ------------
'  Profile DWG files must be deployed in a "CrossBarProfiles" folder
'  placed next to the compiled add-in DLL.  The library resolves paths
'  at resolve-time; a missing file is reported as a warning, not an error.
'
'  Folder expected at runtime:
'    {addin DLL folder}\CrossBarProfiles\
'      3-8 RND PLAIN.dwg
'      3-8 RND SERRATED.dwg
'      1-2 RND PLAIN.dwg
'      1-2 RND SERRATED.dwg
'      1-4 X 1 REC PLAIN.dwg
'      1-4 X 1 REC SERRATED.dwg
'      1-4 X 1 1-4 REC PLAIN.dwg
'      1-4 X 1 1-4 REC SERRATED.dwg
'      3-8 X 1 REC PLAIN.dwg
'      3-8 X 1 REC SERRATED.dwg
'      3-8 X 1 1-4 REC PLAIN.dwg
'      3-8 X 1 1-4 REC SERRATED.dwg
'      1-4 RND PLAIN.dwg
'      1-4 RND SERRATED.dwg
'
' Resolve result
' --------------
'  CrossBarProfileResolveResult (defined in this file) carries:
'   - Success        : True if a matching library entry was found
'   - Definition     : the resolved CrossBarProfileDefinition (Nothing on failure)
'   - FileFound      : True if the source DWG file exists on disk
'   - Message        : diagnostic string for trace logging
'////////////////////////////////////////////////////////////////////

Imports System.Diagnostics
Imports System.IO
Imports System.Linq

' ==================================================================
'  Resolve result (companion type — defined here for locality)
' ==================================================================

''' <summary>
''' Outcome of a CrossBarProfileLibrary.Resolve call.
''' </summary>
Public Class CrossBarProfileResolveResult

    ''' <summary>True when a matching profile entry was found in the library.</summary>
    Public Property Success As Boolean

    ''' <summary>
    ''' The resolved profile definition.
    ''' Nothing when Success is False.
    ''' </summary>
    Public Property Definition As CrossBarProfileDefinition

    ''' <summary>
    ''' True when Success is True AND the source DWG file exists on disk.
    ''' </summary>
    Public Property FileFound As Boolean

    ''' <summary>Diagnostic string for trace logging.</summary>
    Public Property Message As String

    ' Factories
    Public Shared Function Resolved(def As CrossBarProfileDefinition,
                                    message As String) As CrossBarProfileResolveResult
        Return New CrossBarProfileResolveResult With {
            .Success = True,
            .Definition = def,
            .FileFound = def.IsFileFound,
            .Message = message
        }
    End Function

    Public Shared Function NotFound(message As String) As CrossBarProfileResolveResult
        Return New CrossBarProfileResolveResult With {
            .Success = False,
            .FileFound = False,
            .Message = message
        }
    End Function

    Public Overrides Function ToString() As String
        If Not Success Then Return "Not found — " & Message
        Return "Resolved — " & Definition.ToString() &
               If(FileFound, "", " [source file MISSING]")
    End Function

End Class

' ==================================================================
'  Library
' ==================================================================

''' <summary>
''' Catalogue of standard HMG cross bar profiles.
''' Maps CrossBarType + SurfaceProfileType to a
''' CrossBarProfileDefinition that carries physical dimensions
''' and the DWG source file reference.
''' </summary>
Public Class CrossBarProfileLibrary

    ''' <summary>
    ''' Subfolder name, relative to the add-in DLL directory, where
    ''' the profile DWG files are deployed.
    ''' </summary>
    Public Const ProfilesSubfolder As String = "CrossBarProfiles"

    ' ------------------------------------------------------------------
    '  Static library definition
    ' ------------------------------------------------------------------

    ''' <summary>
    ''' All library entries in display-name order.
    ''' File paths are populated at resolve-time by the instance methods.
    ''' </summary>
    Private Shared ReadOnly _entries As CrossBarProfileDefinition() = BuildEntries()

    ''' <summary>Builds the static catalogue array.</summary>
    Private Shared Function BuildEntries() As CrossBarProfileDefinition()
        Dim list As New List(Of CrossBarProfileDefinition)

        ' ---- Round rod -----------------------------------------------

        Dim e As New CrossBarProfileDefinition
        e.DisplayName = "3/8"" dia. Round — Plain"
        e.CrossBar = CrossBarType.Round_3_8
        e.SurfaceProfile = SurfaceProfileType.Plain
        e.ShapeFamily = CrossBarShapeFamily.Round
        e.DiameterInches = 0.375
        e.SourceFileName = "3-8 RND PLAIN.dwg"
        list.Add(e)

        e = New CrossBarProfileDefinition
        e.DisplayName = "3/8"" dia. Round — Serrated"
        e.CrossBar = CrossBarType.Round_3_8
        e.SurfaceProfile = SurfaceProfileType.Serrated
        e.ShapeFamily = CrossBarShapeFamily.Round
        e.DiameterInches = 0.375
        e.SourceFileName = "3-8 RND SERRATED.dwg"
        list.Add(e)

        e = New CrossBarProfileDefinition
        e.DisplayName = "1/2"" dia. Round — Plain"
        e.CrossBar = CrossBarType.Round_1_2
        e.SurfaceProfile = SurfaceProfileType.Plain
        e.ShapeFamily = CrossBarShapeFamily.Round
        e.DiameterInches = 0.5
        e.SourceFileName = "1-2 RND PLAIN.dwg"
        list.Add(e)

        e = New CrossBarProfileDefinition
        e.DisplayName = "1/2"" dia. Round — Serrated"
        e.CrossBar = CrossBarType.Round_1_2
        e.SurfaceProfile = SurfaceProfileType.Serrated
        e.ShapeFamily = CrossBarShapeFamily.Round
        e.DiameterInches = 0.5
        e.SourceFileName = "1-2 RND SERRATED.dwg"
        list.Add(e)

        ' ---- 1/4" thick flat bar -------------------------------------

        e = New CrossBarProfileDefinition
        e.DisplayName = "1/4"" x 1"" Flat Bar — Plain"
        e.CrossBar = CrossBarType.Flat_1_4_x_1
        e.SurfaceProfile = SurfaceProfileType.Plain
        e.ShapeFamily = CrossBarShapeFamily.Rectangular
        e.ThicknessInches = 0.25
        e.HeightInches = 1.0
        e.SourceFileName = "1-4 X 1 REC PLAIN.dwg"
        list.Add(e)

        e = New CrossBarProfileDefinition
        e.DisplayName = "1/4"" x 1"" Flat Bar — Serrated"
        e.CrossBar = CrossBarType.Flat_1_4_x_1
        e.SurfaceProfile = SurfaceProfileType.Serrated
        e.ShapeFamily = CrossBarShapeFamily.Rectangular
        e.ThicknessInches = 0.25
        e.HeightInches = 1.0
        e.SourceFileName = "1-4 X 1 REC SERRATED.dwg"
        list.Add(e)

        e = New CrossBarProfileDefinition
        e.DisplayName = "1/4"" x 1-1/4"" Flat Bar — Plain"
        e.CrossBar = CrossBarType.Flat_1_4_x_1_1_4
        e.SurfaceProfile = SurfaceProfileType.Plain
        e.ShapeFamily = CrossBarShapeFamily.Rectangular
        e.ThicknessInches = 0.25
        e.HeightInches = 1.25
        e.SourceFileName = "1-4 X 1 1-4 REC PLAIN.dwg"
        list.Add(e)

        e = New CrossBarProfileDefinition
        e.DisplayName = "1/4"" x 1-1/4"" Flat Bar — Serrated"
        e.CrossBar = CrossBarType.Flat_1_4_x_1_1_4
        e.SurfaceProfile = SurfaceProfileType.Serrated
        e.ShapeFamily = CrossBarShapeFamily.Rectangular
        e.ThicknessInches = 0.25
        e.HeightInches = 1.25
        e.SourceFileName = "1-4 X 1 1-4 REC SERRATED.dwg"
        list.Add(e)

        ' ---- 3/8" thick flat bar -------------------------------------

        e = New CrossBarProfileDefinition
        e.DisplayName = "3/8"" x 1"" Flat Bar — Plain"
        e.CrossBar = CrossBarType.Flat_3_8_x_1
        e.SurfaceProfile = SurfaceProfileType.Plain
        e.ShapeFamily = CrossBarShapeFamily.Rectangular
        e.ThicknessInches = 0.375
        e.HeightInches = 1.0
        e.SourceFileName = "3-8 X 1 REC PLAIN.dwg"
        list.Add(e)

        e = New CrossBarProfileDefinition
        e.DisplayName = "3/8"" x 1"" Flat Bar — Serrated"
        e.CrossBar = CrossBarType.Flat_3_8_x_1
        e.SurfaceProfile = SurfaceProfileType.Serrated
        e.ShapeFamily = CrossBarShapeFamily.Rectangular
        e.ThicknessInches = 0.375
        e.HeightInches = 1.0
        e.SourceFileName = "3-8 X 1 REC SERRATED.dwg"
        list.Add(e)

        e = New CrossBarProfileDefinition
        e.DisplayName = "3/8"" x 1-1/4"" Flat Bar — Plain"
        e.CrossBar = CrossBarType.Flat_3_8_x_1_1_4
        e.SurfaceProfile = SurfaceProfileType.Plain
        e.ShapeFamily = CrossBarShapeFamily.Rectangular
        e.ThicknessInches = 0.375
        e.HeightInches = 1.25
        e.SourceFileName = "3-8 X 1 1-4 REC PLAIN.dwg"
        list.Add(e)

        e = New CrossBarProfileDefinition
        e.DisplayName = "3/8"" x 1-1/4"" Flat Bar — Serrated"
        e.CrossBar = CrossBarType.Flat_3_8_x_1_1_4
        e.SurfaceProfile = SurfaceProfileType.Serrated
        e.ShapeFamily = CrossBarShapeFamily.Rectangular
        e.ThicknessInches = 0.375
        e.HeightInches = 1.25
        e.SourceFileName = "3-8 X 1 1-4 REC SERRATED.dwg"
        list.Add(e)

        Return list.ToArray()
    End Function

    ' ------------------------------------------------------------------
    '  Instance state
    ' ------------------------------------------------------------------

    ''' <summary>
    ''' The folder on disk where profile DWG files are located.
    ''' Passed in from GratingCommand (resolved relative to the DLL).
    ''' Empty string means path resolution will be skipped and
    ''' IsFileFound will always be False.
    ''' </summary>
    Private ReadOnly _profilesFolder As String

    ''' <summary>
    ''' Creates the library with an explicit profiles folder path.
    ''' </summary>
    ''' <param name="profilesFolder">
    ''' Absolute path to the folder containing profile DWG files.
    ''' Pass an empty string or Nothing to skip path resolution
    ''' (profile definitions are still returned; IsFileFound = False).
    ''' </param>
    Public Sub New(Optional profilesFolder As String = "")
        _profilesFolder = If(profilesFolder, "")
    End Sub

    ' ------------------------------------------------------------------
    '  Public API
    ' ------------------------------------------------------------------

    ''' <summary>
    ''' Returns a read-only view of all library entries (no paths resolved).
    ''' Useful for enumeration and UI listing.
    ''' </summary>
    Public Shared Function GetAll() As IEnumerable(Of CrossBarProfileDefinition)
        Return _entries
    End Function

    ''' <summary>
    ''' Resolves a profile definition from a GratingParameters instance.
    ''' Delegates to Resolve(CrossBarType, SurfaceProfileType).
    ''' </summary>
    Public Function Resolve(params As GratingParameters) As CrossBarProfileResolveResult
        If params Is Nothing Then
            Return CrossBarProfileResolveResult.NotFound(
                "GratingParameters is Nothing — cannot resolve profile.")
        End If
        Return Resolve(params.CrossBar, params.SurfaceProfile)
    End Function

    ''' <summary>
    ''' Resolves a profile definition for the given cross bar type and
    ''' surface treatment.  The returned definition has SourceFilePath
    ''' set to the absolute path if the file can be located on disk.
    ''' </summary>
    ''' <returns>
    ''' A <see cref="CrossBarProfileResolveResult"/> that is always safe
    ''' to inspect; Success is False only when no library entry matches
    ''' the supplied combination (not when the file is merely missing).
    ''' </returns>
    Public Function Resolve(crossBar As CrossBarType,
                            surface As SurfaceProfileType) As CrossBarProfileResolveResult

        ' Find matching entry in the static catalogue
        Dim match As CrossBarProfileDefinition = Nothing
        For Each entry As CrossBarProfileDefinition In _entries
            If entry.CrossBar = crossBar AndAlso entry.SurfaceProfile = surface Then
                match = entry
                Exit For
            End If
        Next

        If match Is Nothing Then
            Dim msg As String =
                "No library entry for CrossBarType=" & crossBar.ToString() &
                ", SurfaceProfile=" & surface.ToString() & "."
            Trace.TraceWarning(": HMG ProfileLibrary: " & msg)
            Return CrossBarProfileResolveResult.NotFound(msg)
        End If

        ' Clone the entry so we can stamp SourceFilePath without mutating
        ' the shared static catalogue.
        Dim def As CrossBarProfileDefinition = CloneEntry(match)

        ' Resolve the file path
        Dim resolvedPath As String = ResolveFilePath(def.SourceFileName)
        def.SourceFilePath = resolvedPath

        Dim fileStatus As String
        If def.IsFileFound Then
            fileStatus = "file found at: " & resolvedPath
        ElseIf Not String.IsNullOrEmpty(_profilesFolder) Then
            fileStatus = "file NOT FOUND — expected at: " &
                         Path.Combine(_profilesFolder, def.SourceFileName)
        Else
            fileStatus = "file path not resolved (profiles folder not set)"
        End If

        Dim message As String =
            "CrossBarType=" & crossBar.ToString() &
            ", Surface=" & surface.ToString() &
            " → """ & def.DisplayName & """  [" & fileStatus & "]"

        Trace.TraceInformation(": HMG ProfileLibrary: " & message)

        If Not def.IsFileFound Then
            Trace.TraceWarning(
                ": HMG ProfileLibrary: Source file missing — " &
                def.SourceFileName &
                ". Cross bar geometry generation will fail until the file is deployed.")
        End If

        Return CrossBarProfileResolveResult.Resolved(def, message)
    End Function

    ' ------------------------------------------------------------------
    '  Path resolution
    ' ------------------------------------------------------------------

    ''' <summary>
    ''' Attempts to build the full path to a profile DWG file using
    ''' the configured profiles folder.  Returns the path string even if
    ''' the file does not exist (the caller checks IsFileFound).
    ''' Returns an empty string if no profiles folder is configured.
    ''' </summary>
    Private Function ResolveFilePath(fileName As String) As String
        If String.IsNullOrEmpty(_profilesFolder) OrElse
           String.IsNullOrEmpty(fileName) Then
            Return String.Empty
        End If
        Return Path.Combine(_profilesFolder, fileName)
    End Function

    ' ------------------------------------------------------------------
    '  Private helpers
    ' ------------------------------------------------------------------

    ''' <summary>
    ''' Returns a shallow copy of a library entry so the caller can set
    ''' SourceFilePath without mutating the shared static catalogue.
    ''' </summary>
    Private Shared Function CloneEntry(
            src As CrossBarProfileDefinition) As CrossBarProfileDefinition
        Return New CrossBarProfileDefinition With {
            .DisplayName = src.DisplayName,
            .CrossBar = src.CrossBar,
            .SurfaceProfile = src.SurfaceProfile,
            .ShapeFamily = src.ShapeFamily,
            .DiameterInches = src.DiameterInches,
            .ThicknessInches = src.ThicknessInches,
            .HeightInches = src.HeightInches,
            .SourceFileName = src.SourceFileName,
            .SourceFilePath = String.Empty
        }
    End Function

End Class
