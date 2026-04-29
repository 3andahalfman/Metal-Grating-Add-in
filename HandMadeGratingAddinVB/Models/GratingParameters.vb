'////////////////////////////////////////////////////////////////////
' Metal Bar Grating Addin - VB.NET
'
' GratingParameters: Full parameter model for metal bar grating
' generation. Combines V1 bearing bar settings with V2 cross bar,
' surface profile, and banding options.
'
' Units: all dimensions are in inches.
'////////////////////////////////////////////////////////////////////

' ===================================================================
'  Enumerations
' ===================================================================

''' <summary>
''' Direction the bearing bars span across the grating perimeter.
''' </summary>
Public Enum SpanDirectionType
    ''' <summary>Bearing bars run parallel to the sketch X axis.</summary>
    AlongX = 0
    ''' <summary>Bearing bars run parallel to the sketch Y axis.</summary>
    AlongY = 1
End Enum

''' <summary>
''' Cross bar shape and size. Each member maps to a specific
''' bar profile from the HMG specification.
''' </summary>
Public Enum CrossBarType
    ''' <summary>3/8" diameter round rod.</summary>
    Round_3_8 = 0
    ''' <summary>1/2" diameter round rod.</summary>
    Round_1_2 = 1
    ''' <summary>1/4" x 1" rectangular flat bar.</summary>
    Flat_1_4_x_1 = 2
    ''' <summary>1/4" x 1-1/4" rectangular flat bar.</summary>
    Flat_1_4_x_1_1_4 = 3
    ''' <summary>3/8" x 1" rectangular flat bar.</summary>
    Flat_3_8_x_1 = 4
    ''' <summary>3/8" x 1-1/4" rectangular flat bar.</summary>
    Flat_3_8_x_1_1_4 = 5
End Enum

''' <summary>
''' Surface treatment for the cross bars.
''' </summary>
Public Enum SurfaceProfileType
    ''' <summary>Smooth / plain surface.</summary>
    Plain = 0
    ''' <summary>Serrated surface for slip resistance.</summary>
    Serrated = 1
End Enum

''' <summary>
''' Whether the grating has band bars around the perimeter.
''' </summary>
Public Enum BandingOptionType
    ''' <summary>Band bars follow the perimeter edge.</summary>
    Banded = 0
    ''' <summary>No band bars — bearing bars have open ends.</summary>
    OpenEnded = 1
End Enum

' ===================================================================
'  Helper: display-name lookups for enums
' ===================================================================

''' <summary>
''' Provides human-readable display names for cross bar types.
''' </summary>
Public Class CrossBarTypeHelper

    ''' <summary>Ordered display names matching CrossBarType values.</summary>
    Public Shared ReadOnly DisplayNames() As String = {
        "3/8"" dia. Round",
        "1/2"" dia. Round",
        "1/4"" x 1"" Flat Bar",
        "1/4"" x 1-1/4"" Flat Bar",
        "3/8"" x 1"" Flat Bar",
        "3/8"" x 1-1/4"" Flat Bar"
    }

    ''' <summary>Returns the display name for a CrossBarType value.</summary>
    Public Shared Function GetDisplayName(t As CrossBarType) As String
        Dim idx As Integer = CInt(t)
        If idx >= 0 AndAlso idx < DisplayNames.Length Then
            Return DisplayNames(idx)
        End If
        Return t.ToString()
    End Function

End Class

' ===================================================================
'  Parameter model
' ===================================================================

''' <summary>
''' Stores all grating input parameters (V1 bearing bar + V2 cross bar,
''' surface profile, and banding options).
''' </summary>
Public Class GratingParameters

    ' --- V1: Bearing bar settings ---

    ''' <summary>Direction the bearing bars span.</summary>
    Public Property SpanDirection As SpanDirectionType

    ''' <summary>Bearing bar depth (height of the bar cross-section) in inches.</summary>
    Public Property BarDepth As Double

    ''' <summary>Bearing bar width (thickness of the bar cross-section) in inches.</summary>
    Public Property BarWidth As Double

    ''' <summary>On-center spacing between bearing bars in inches.</summary>
    Public Property OnCenterSpacing As Double

    ''' <summary>
    ''' Output folder path for generated files.
    ''' Empty string means use the active document's folder.
    ''' </summary>
    Public Property OutputFolder As String

    ''' <summary>Optional naming prefix for output files.
    ''' When empty, <see cref="ResolvedPrefix"/> generates a
    ''' descriptive prefix from the grating parameters.</summary>
    Public Property NamingPrefix As String

    ''' <summary>
    ''' Returns <see cref="NamingPrefix"/> if the user specified one,
    ''' otherwise builds a descriptive prefix from the parameters.
    ''' </summary>
    Public ReadOnly Property ResolvedPrefix As String
        Get
            If Not String.IsNullOrWhiteSpace(NamingPrefix) Then
                Return NamingPrefix
            End If
            Return BuildDescriptivePrefix()
        End Get
    End Property

    ' --- V2: Cross bar settings ---

    ''' <summary>Cross bar shape and size.</summary>
    Public Property CrossBar As CrossBarType

    ''' <summary>Cross bar on-center spacing in inches (typically 2 or 4).</summary>
    Public Property CrossBarOnCenter As Double

    ''' <summary>
    ''' Distance from the bearing bar end to the first cross bar
    ''' center in inches. Must be positive.
    ''' </summary>
    Public Property FirstCrossBarOffset As Double

    ' --- V2: Surface and banding ---

    ''' <summary>Surface treatment (plain or serrated).</summary>
    Public Property SurfaceProfile As SurfaceProfileType

    ''' <summary>Whether the grating is banded or open-ended.</summary>
    Public Property Banding As BandingOptionType

    ' --- V3: Notch dimension overrides (optional, Nothing = use default) ---

    ''' <summary>Override for notch slot width in inches. Nothing = use cross bar default.</summary>
    Public Property NotchSlotWidthOverride As Double?

    ''' <summary>Override for total notch slot depth in inches. Nothing = use cross bar default.</summary>
    Public Property NotchSlotDepthOverride As Double?

    ''' <summary>Override for straight wall depth before arc (round only) in inches. Nothing = use default.</summary>
    Public Property NotchStraightDepthOverride As Double?

    ''' <summary>Override for bottom radius (round only) in inches. Nothing = use default.</summary>
    Public Property NotchBottomRadiusOverride As Double?

    ' --- V3: Serration dimension overrides (optional, Nothing = use default) ---

    ''' <summary>Override for serration scallop chord width in inches. Nothing = use default.</summary>
    Public Property SerrationChordOverride As Double?

    ''' <summary>Override for flat land width between scallops in inches. Nothing = use default.</summary>
    Public Property SerrationFlatWidthOverride As Double?

    ''' <summary>Override for serration scallop arc radius in inches. Nothing = use default.</summary>
    Public Property SerrationArcRadiusOverride As Double?

    ''' <summary>Override for clearance from notch slot edge to nearest scallop in inches. Nothing = use default.</summary>
    Public Property SerrationMarginOverride As Double?

    ' --- Descriptive naming ---

    ''' <summary>
    ''' Builds a descriptive naming prefix from the current parameters.
    ''' Pattern: {BarWidth}x{BarDepth}-{OC}-{CrossBar}-{CBOC}-{Surface}{Banding}
    ''' Example: 3-16x2-1-3_16-R38-2-PB  (3/16" x 2", 1-3/16" OC, Round 3/8, 2" CB OC, Plain, Banded)
    ''' </summary>
    Public Function BuildDescriptivePrefix() As String
        Dim sb As New System.Text.StringBuilder()

        ' Bearing bar: WidthxDepth  (e.g. 3-16x2)
        sb.Append(FractionTag(BarWidth))
        sb.Append("x")
        sb.Append(FractionTag(BarDepth))

        ' Bar O.C. spacing  (e.g. -1-3_16)
        sb.Append("-")
        sb.Append(FractionTag(OnCenterSpacing))

        ' Cross bar code  (e.g. -R38, -R12, -F14x1, -F38x1-14)
        sb.Append("-")
        sb.Append(CrossBarTag(CrossBar))

        ' Cross bar O.C.  (e.g. -2 or -4)
        sb.Append("-")
        sb.Append(FractionTag(CrossBarOnCenter))

        ' Surface + Banding  (e.g. -PB, -SB, -PO, -SO)
        sb.Append("-")
        sb.Append(If(SurfaceProfile = SurfaceProfileType.Serrated, "S", "P"))
        sb.Append(If(Banding = BandingOptionType.OpenEnded, "O", "B"))

        Return sb.ToString()
    End Function

    ''' <summary>
    ''' Converts a decimal-inch value to a compact file-safe string.
    ''' Uses common fractional representations when possible.
    ''' Examples: 0.1875 → "3-16", 1.1875 → "1-3_16", 2.0 → "2",
    '''           0.375 → "3-8", 0.5 → "1-2", 0.25 → "1-4"
    ''' </summary>
    Private Shared Function FractionTag(value As Double) As String
        ' Known fraction lookup (denominator 16ths)
        Dim whole As Integer = CInt(Math.Floor(value))
        Dim frac As Double = value - whole

        ' Common fractions (tolerance 0.001)
        Dim fracStr As String = ""
        If Math.Abs(frac) < 0.001 Then
            fracStr = ""
        ElseIf Math.Abs(frac - 0.0625) < 0.001 Then
            fracStr = "1-16"
        ElseIf Math.Abs(frac - 0.125) < 0.001 Then
            fracStr = "1-8"
        ElseIf Math.Abs(frac - 0.1875) < 0.001 Then
            fracStr = "3-16"
        ElseIf Math.Abs(frac - 0.25) < 0.001 Then
            fracStr = "1-4"
        ElseIf Math.Abs(frac - 0.3125) < 0.001 Then
            fracStr = "5-16"
        ElseIf Math.Abs(frac - 0.375) < 0.001 Then
            fracStr = "3-8"
        ElseIf Math.Abs(frac - 0.4375) < 0.001 Then
            fracStr = "7-16"
        ElseIf Math.Abs(frac - 0.5) < 0.001 Then
            fracStr = "1-2"
        ElseIf Math.Abs(frac - 0.5625) < 0.001 Then
            fracStr = "9-16"
        ElseIf Math.Abs(frac - 0.625) < 0.001 Then
            fracStr = "5-8"
        ElseIf Math.Abs(frac - 0.6875) < 0.001 Then
            fracStr = "11-16"
        ElseIf Math.Abs(frac - 0.75) < 0.001 Then
            fracStr = "3-4"
        ElseIf Math.Abs(frac - 0.8125) < 0.001 Then
            fracStr = "13-16"
        ElseIf Math.Abs(frac - 0.875) < 0.001 Then
            fracStr = "7-8"
        ElseIf Math.Abs(frac - 0.9375) < 0.001 Then
            fracStr = "15-16"
        Else
            ' Non-standard fraction — use decimal
            Return value.ToString("G")
        End If

        If whole > 0 AndAlso fracStr.Length > 0 Then
            Return whole.ToString() & "-" & fracStr
        ElseIf whole > 0 Then
            Return whole.ToString()
        Else
            Return fracStr
        End If
    End Function

    ''' <summary>
    ''' Returns a compact cross bar type tag for file naming.
    ''' R38 = Round 3/8, R12 = Round 1/2, F14x1 = Flat 1/4 x 1, etc.
    ''' </summary>
    Private Shared Function CrossBarTag(cb As CrossBarType) As String
        Select Case cb
            Case CrossBarType.Round_3_8 : Return "R38"
            Case CrossBarType.Round_1_2 : Return "R12"
            Case CrossBarType.Flat_1_4_x_1 : Return "F14x1"
            Case CrossBarType.Flat_1_4_x_1_1_4 : Return "F14x1-14"
            Case CrossBarType.Flat_3_8_x_1 : Return "F38x1"
            Case CrossBarType.Flat_3_8_x_1_1_4 : Return "F38x1-14"
            Case Else : Return "CB"
        End Select
    End Function

    ' --- Factory ---

    ''' <summary>
    ''' Creates default parameters matching common grating:
    '''   2" deep x 3/16" wide bars at 1-3/16" on-center,
    '''   3/8" round cross bars at 2" O.C., plain, banded.
    ''' </summary>
    Public Shared Function CreateDefaults() As GratingParameters
        Return New GratingParameters With {
            .SpanDirection = SpanDirectionType.AlongX,
            .BarDepth = 2.0,
            .BarWidth = 0.1875,
            .OnCenterSpacing = 1.1875,
            .OutputFolder = "",
            .NamingPrefix = "",
            .CrossBar = CrossBarType.Round_3_8,
            .CrossBarOnCenter = 2.0,
            .FirstCrossBarOffset = 1.0,
            .SurfaceProfile = SurfaceProfileType.Plain,
            .Banding = BandingOptionType.Banded
        }
    End Function

End Class
