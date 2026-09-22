Attribute VB_Name = "mdlTypes"
Option Explicit

' ===========================================================================
' mdlTypes
' Shared data structures for the "Select Similar Objects" macro.
' ===========================================================================

Public Type TCriteria
    ObjectType          As Boolean
    FullColor           As Boolean
    OutlineColor        As Boolean
    OutlineWidth        As Boolean
    SameSize            As Boolean
    NodeCount           As Boolean

    WidthOnly           As Boolean
    HeightOnly          As Boolean
    AspectRatio         As Boolean
    Rotation            As Boolean
    Transparency        As Boolean
    FountainFill        As Boolean
    PatternFill         As Boolean
    NoFillPresent       As Boolean
    NoOutlinePresent    As Boolean
    IgnoreRotationSize  As Boolean

    TextFont            As Boolean
    TextSize            As Boolean
    TextContent         As Boolean
    TextStyle           As Boolean

    BitmapResolution    As Boolean
    BitmapWidth         As Boolean
    BitmapHeight        As Boolean
    BitmapColorMode     As Boolean

    SizeTolerance       As Double   ' mm
    OutlineTolerance    As Double   ' mm
    RotationTolerance   As Double   ' degrees
End Type

Public Type TScanOptions
    Scope                As Integer  ' 0=Current Page,1=All Pages,2=Current Layer,3=All Layers,4=Current Selection
    SearchInsideGroups   As Boolean
    IncludeLocked        As Boolean
    IncludeHidden        As Boolean
    IncludeReference     As Boolean
    AddToExisting        As Boolean
End Type

Public Const SCOPE_CURRENT_PAGE      As Integer = 0
Public Const SCOPE_ALL_PAGES         As Integer = 1
Public Const SCOPE_CURRENT_LAYER     As Integer = 2
Public Const SCOPE_ALL_LAYERS        As Integer = 3
Public Const SCOPE_CURRENT_SELECTION As Integer = 4

