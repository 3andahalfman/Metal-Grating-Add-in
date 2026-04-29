'////////////////////////////////////////////////////////////////////
' Metal Bar Grating Addin - VB.NET
'
' NotchProfileData: Stores a user-drawn notch profile as a sequence
' of sketch entities (lines and arcs) with coordinates in centimeters
' relative to the source sketch coordinate system.
'
' The profile is drawn once in a temporary part document on the XZ
' origin plane (sketch X = world X, sketch Y = world Z).  The stored
' entities are replayed at each notch position during bearing bar
' generation by shifting all X coordinates to the notch position.
'
' Coordinate convention (in the XZ-plane sketch):
'   Sketch X = along bar length (world X)
'   Sketch Y = along bar depth direction (world Z)
'   All values in cm (Inventor internal units).
'////////////////////////////////////////////////////////////////////

''' <summary>
''' Stores a user-drawn notch profile for replay during bearing bar
''' generation.  All coordinates are in centimeters in the XZ-plane
''' sketch coordinate system.
''' </summary>
Public Class NotchProfileData

    ''' <summary>Ordered list of sketch entities forming the closed notch profile.</summary>
    Public Property Entities As New List(Of NotchSketchEntity)

    ''' <summary>True if at least two entities were captured.</summary>
    Public ReadOnly Property IsValid As Boolean
        Get
            Return Entities IsNot Nothing AndAlso Entities.Count >= 2
        End Get
    End Property

    ''' <summary>X position of the notch center in the source sketch (cm).
    ''' Used to compute relative offsets during replay.</summary>
    Public Property SourceCenterX As Double

    ''' <summary>Description of how the profile was created.</summary>
    Public Property Description As String

    Public Overrides Function ToString() As String
        Return If(Description, "NotchProfile") & " — " &
               If(Entities IsNot Nothing, Entities.Count, 0) & " entities"
    End Function

End Class

''' <summary>Type of sketch entity in a notch profile.</summary>
Public Enum NotchEntityType
    ''' <summary>Straight line between two points.</summary>
    Line = 0
    ''' <summary>Arc defined by three points (start, mid, end).</summary>
    ThreePointArc = 1
End Enum

''' <summary>
''' A single sketch entity (line or arc) within a notch profile.
''' All coordinates are in centimeters (Inventor internal units).
''' </summary>
Public Class NotchSketchEntity

    ''' <summary>Type of entity.</summary>
    Public Property EntityType As NotchEntityType

    ''' <summary>Start point X (cm).</summary>
    Public Property StartX As Double

    ''' <summary>Start point Y (cm).</summary>
    Public Property StartY As Double

    ''' <summary>End point X (cm).</summary>
    Public Property EndX As Double

    ''' <summary>End point Y (cm).</summary>
    Public Property EndY As Double

    ''' <summary>Mid point X — only for ThreePointArc (cm).</summary>
    Public Property MidX As Double

    ''' <summary>Mid point Y — only for ThreePointArc (cm).</summary>
    Public Property MidY As Double

End Class
