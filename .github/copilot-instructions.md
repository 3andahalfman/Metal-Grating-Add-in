# Copilot Instructions

## Project Guidelines
- Metal Bar Grating (HMG) plugin full specification:

PRODUCT: Bearing bars are laser-cut from plate with notches in the edge to accept cross rods. Bars are placed in a fixture, cross rods manually welded. Used when bar sizes exceed machine-weldable panel limits.

BEARING BARS: Size range 1/8"x1" to 1/2"x8". OC spacings: 7/16", 11/16", 15/16", 1-3/16", 1-3/8", 1-7/8", 2-3/8". Each bar must be individual .ipt with unique name for laser cutter. Bars have notches cut for cross rods. Cutouts/skew can produce 30+ unique bar lengths.

CROSS BARS: Round (3/8" dia, 1/2" dia) or Rectangular (1/4"x1", 1/4"x1-1/4", 3/8"x1", 3/8"x1-1/4"). Each type can be plain or serrated. OC spacings: 2" or 4". First cross rod distance from bearing bar end is adjustable. DWG profiles exist in HMG Description/dwg folder.

BAND BARS: Optional (banding or open-ended). Follow perimeter edge of grating.

OUTPUTS: Individual .ipt files per bearing bar (with notches). PDF fabrication drawing showing: visual grate layout, part locations, cross bar length table, band bar length table.

PERIMETER GEOMETRY: Imported from 2D AutoCAD sketches. Can be rectangle, skewed, cutouts, curved edges. Sample shapes shown in spec PDF.

BAR MARKS: "ALL BAR MARKS ARE TO BE 1 BAR UNLESS NOTED AFTER BAR PIECE MARK" — notation like "7a 1-1", "7a 2-1" with individual lengths.

## Versioning
- Every change/release must have a new version number. Bump the version in `Installer\MetalBarGrating.iss` (`#define MyAppVersion`) before each new installer build.