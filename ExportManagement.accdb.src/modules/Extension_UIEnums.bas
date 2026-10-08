Attribute VB_Name = "Extension_UIEnums"
Option Compare Database
Option Explicit
'@Folder "Extensions.UI"

'====================================================================
' MODULE    : Extension_UIEnums
' PURPOSE   : UI Design Tokens, Enumerations & Style Constants
'====================================================================

'============================ CONSTANTS ===========================
Public Const TWIPS_PER_INCH = 1440

'============================ CONSTANTS: UI ENUMERATIONS ===========================

' --- HorizontalAnchor ---
Public Const HorizontalAnchor_Left As Integer = 0
Public Const HorizontalAnchor_Right As Integer = 1
Public Const HorizontalAnchor_Both As Integer = 2

' --- VerticalAnchor ---
Public Const VerticalAnchor_Top As Integer = 0
Public Const VerticalAnchor_Bottom As Integer = 1
Public Const VerticalAnchor_Both As Integer = 2

' --- BorderStyle ---
Public Const BorderStyle_Transparent As Integer = 0
Public Const BorderStyle_Solid As Integer = 1
Public Const BorderStyle_Dashes As Integer = 2
Public Const BorderStyle_Shortdashes As Integer = 3
Public Const BorderStyle_Dots As Integer = 4
Public Const BorderStyle_Sparsedots As Integer = 5
Public Const BorderStyle_Dashdot As Integer = 6
Public Const BorderStyle_Dashdotdot As Integer = 7

' --- ScrollBars ---
Public Const ScrollBars_Neither As Integer = 0
Public Const ScrollBars_HorizontalOnly As Integer = 1
Public Const ScrollBars_VerticalOnly As Integer = 2
Public Const ScrollBars_Both As Integer = 3

' --- ScrollBarAlign ---
Public Const ScrollBarAlign_System As Integer = 0
Public Const ScrollBarAlign_Left As Integer = 1
Public Const ScrollBarAlign_Right As Integer = 2

' --- DefaultView ---
Public Const DefaultView_SingleForm As Integer = 0
Public Const DefaultView_ContinuousForms As Integer = 1
Public Const DefaultView_Datasheet As Integer = 2
Public Const DefaultView_PivotTable As Integer = 3
Public Const DefaultView_PivotChart As Integer = 4
Public Const DefaultView_SplitForm As Integer = 5

' --- WindowBorderStyle ---
Public Const WindowBorderStyle_None As Integer = 0
Public Const WindowBorderStyle_Thin As Integer = 1
Public Const WindowBorderStyle_Sizable As Integer = 2
Public Const WindowBorderStyle_Dialog As Integer = 3

' --- MinMaxButton ---
Public Const MinMaxButtons_Both As Integer = 3
Public Const MinMaxButtons_Max As Integer = 2
Public Const MinMaxButtons_Min As Integer = 1
Public Const MinMaxButtons_None As Integer = 0

' --- PictureType ---
Public Const PictureType_Embedded As Integer = 0
Public Const PictureType_Linked As Integer = 1
Public Const PictureType_Shared As Integer = 2


' --- [ BUTTON SHAPES ] ---
Public Const ButtonShape_Rectangle As Integer = 0
Public Const ButtonShape_Rounded As Integer = 1
Public Const ButtonShape_SnipSingle As Integer = 2
Public Const ButtonShape_SnipSameSide As Integer = 3
Public Const ButtonShape_SnipDiagonal As Integer = 4
Public Const ButtonShape_SnipRoundSingle As Integer = 5
Public Const ButtonShape_RoundSingle As Integer = 6
Public Const ButtonShape_RoundSameSide As Integer = 7
Public Const ButtonShape_RoundDiagonal As Integer = 8
Public Const ButtonShape_Oval As Integer = 9

' --- [ BUTTON BEVELS ] ---
Public Const ButtonBevel_None As Integer = 0
Public Const ButtonBevel_Circle As Integer = 1
Public Const ButtonBevel_RelaxedInset As Integer = 2
Public Const ButtonBevel_Cross As Integer = 3
Public Const ButtonBevel_CoolSlant As Integer = 4
Public Const ButtonBevel_Angle As Integer = 5
Public Const ButtonBevel_SoftRound As Integer = 6
Public Const ButtonBevel_Convex As Integer = 7
Public Const ButtonBevel_Slope As Integer = 8
Public Const ButtonBevel_Divot As Integer = 9
Public Const ButtonBevel_Riblet As Integer = 10
Public Const ButtonBevel_HardEdge As Integer = 11
Public Const ButtonBevel_ArtDeco As Integer = 12

' --- [ BUTTON GLOWS ] ---
Public Const ButtonGlow_None As Integer = 0
Public Const ButtonGlow_Blue_5pt As Integer = 1
Public Const ButtonGlow_Red_5pt As Integer = 2
Public Const ButtonGlow_Olive_5pt As Integer = 3
Public Const ButtonGlow_Purple_5pt As Integer = 4
Public Const ButtonGlow_Aqua_5pt As Integer = 5
Public Const ButtonGlow_Orange_5pt As Integer = 6
Public Const ButtonGlow_Blue_8pt As Integer = 7
Public Const ButtonGlow_Red_8pt As Integer = 8
Public Const ButtonGlow_Olive_8pt As Integer = 9
Public Const ButtonGlow_Purple_8pt As Integer = 10
Public Const ButtonGlow_Aqua_8pt As Integer = 11
Public Const ButtonGlow_Orange_8pt As Integer = 12
Public Const ButtonGlow_Blue_11pt As Integer = 13
Public Const ButtonGlow_Red_11pt As Integer = 14
Public Const ButtonGlow_Olive_11pt As Integer = 15
Public Const ButtonGlow_Purple_11pt As Integer = 16
Public Const ButtonGlow_Aqua_11pt As Integer = 17
Public Const ButtonGlow_Orange_11pt As Integer = 18
Public Const ButtonGlow_Blue_18pt As Integer = 19
Public Const ButtonGlow_Red_18pt As Integer = 20
Public Const ButtonGlow_Olive_18pt As Integer = 21
Public Const ButtonGlow_Purple_18pt As Integer = 22
Public Const ButtonGlow_Aqua_18pt As Integer = 23
Public Const ButtonGlow_Orange_18pt As Integer = 24

' --- [ BUTTON SOFT EDGES ] ---
Public Const ButtonSoftEdges_None As Integer = 0
Public Const ButtonSoftEdges_1pt As Integer = 1
Public Const ButtonSoftEdges_2_5pt As Integer = 2
Public Const ButtonSoftEdges_5pt As Integer = 3
Public Const ButtonSoftEdges_10pt As Integer = 4
Public Const ButtonSoftEdges_25pt As Integer = 5
Public Const ButtonSoftEdges_50pt As Integer = 6

' --- [ BUTTON GRADIENTS ] ---
Public Const ButtonGradient_None As Integer = 0
Public Const ButtonGradient_LinearDiagonal_TLBR_Light As Integer = 1
Public Const ButtonGradient_LinearDown_Light As Integer = 2
Public Const ButtonGradient_LinearDiagonal_TRBL_Light As Integer = 3
Public Const ButtonGradient_FromBottomRight_Light As Integer = 4
Public Const ButtonGradient_FromBottomLeft_Light As Integer = 5
Public Const ButtonGradient_LinearLeft_Light As Integer = 6
Public Const ButtonGradient_LinearCenter_Light As Integer = 7
Public Const ButtonGradient_LinearRight_Light As Integer = 8
Public Const ButtonGradient_FromTopRight_Light As Integer = 9
Public Const ButtonGradient_FromTopLeft_Light As Integer = 10
Public Const ButtonGradient_LinearDiagonal_BLTR_Light As Integer = 11
Public Const ButtonGradient_LinearUp_Light As Integer = 12
Public Const ButtonGradient_LinearDiagonal_BRTL_Light As Integer = 13
Public Const ButtonGradient_LinearDiagonal_TLBR_Dark As Integer = 14
Public Const ButtonGradient_LinearDown_Dark As Integer = 15
Public Const ButtonGradient_LinearDiagonal_TRBL_Dark As Integer = 16
Public Const ButtonGradient_FromBottomRight_Dark As Integer = 17
Public Const ButtonGradient_FromBottomLeft_Dark As Integer = 18
Public Const ButtonGradient_LinearLeft_Dark As Integer = 19
Public Const ButtonGradient_LinearCenter_Dark As Integer = 20
Public Const ButtonGradient_LinearRight_Dark As Integer = 21
Public Const ButtonGradient_FromTopRight_Dark As Integer = 22
Public Const ButtonGradient_FromTopLeft_Dark As Integer = 23
Public Const ButtonGradient_LinearDiagonal_BLTR_Dark As Integer = 24
Public Const ButtonGradient_LinearUp_Dark As Integer = 25
Public Const ButtonGradient_LinearDiagonal_BRTL_Light_V2 As Integer = 26

' --- PictureAlignment ---
Public Const PictureAlignment_TopLeft As Integer = 0
Public Const PictureAlignment_TopRight As Integer = 1
Public Const PictureAlignment_Center As Integer = 2
Public Const PictureAlignment_BottomLeft As Integer = 3
Public Const PictureAlignment_BottomRight As Integer = 4
Public Const PictureAlignment_FormCenter As Integer = 5

' --- PictureSizeMode ---
Public Const PictureSizeMode_Clip As Integer = 0
Public Const PictureSizeMode_Stretch As Integer = 1
Public Const PictureSizeMode_Zoom As Integer = 3
Public Const PictureSizeMode_StretchHorizontal As Integer = 4
Public Const PictureSizeMode_StretchVertical As Integer = 5

' --- Orientation ---
Public Const Orientation_LeftToRight As Integer = 0
Public Const Orientation_RightToLeft As Integer = 1

' --- RecordsetType ---
Public Const RecordsetType_Dynaset As Integer = 0
Public Const RecordsetType_DynasetInconsistent As Integer = 1
Public Const RecordsetType_Snapshot As Integer = 2

' --- RecordLocks ---
Public Const RecordLocks_NoLock As Integer = 0
Public Const RecordLocks_AllRecords As Integer = 1
Public Const RecordLocks_EditedRecord As Integer = 2

' --- Cycle ---
Public Const Cycle_AllRecords As Integer = 0
Public Const Cycle_CurrentRecord As Integer = 1
Public Const Cycle_CurrentPage As Integer = 2

' --- SpecialEffect ---
Public Const SpecialEffect_Flat As Integer = 0
Public Const SpecialEffect_Raised As Integer = 1
Public Const SpecialEffect_Sunken As Integer = 2
Public Const SpecialEffect_Etched As Integer = 3
Public Const SpecialEffect_Shadowed As Integer = 4
Public Const SpecialEffect_Chiseled As Integer = 5

' --- [ TEXT ALIGNMENT ] ---
Public Const TextAlign_General As Integer = 0
Public Const TextAlign_Left As Integer = 1
Public Const TextAlign_Center As Integer = 2
Public Const TextAlign_Right As Integer = 3

' --- DisplayWhen ---
Public Const DisplayWhen_Always As Integer = 0
Public Const DisplayWhen_PrintOnly As Integer = 1
Public Const DisplayWhen_ScreenOnly As Integer = 2

' --- PrintSection ---
Public Const PrintSection_None As Integer = 0
Public Const PrintSection_BeforeSection As Integer = 1
Public Const PrintSection_AfterSection As Integer = 2
Public Const PrintSection_BeforeAfter As Integer = 3

' --- BackStyle ---
Public Const BackStyle_Normal As Integer = 1
Public Const BackStyle_Transparent As Integer = 0

' --- FontAlignment ---
Public Const FontAlignment_General As Integer = 0
Public Const FontAlignment_Left As Integer = 1
Public Const FontAlignment_Center As Integer = 2
Public Const FontAlignment_Right As Integer = 3
Public Const FontAlignment_Distribute As Integer = 4

' --- FontWeight ---
Public Const FontWeight_Thin As Integer = 100
Public Const FontWeight_ExtraLight As Integer = 200
Public Const FontWeight_Light As Integer = 300
Public Const FontWeight_Normal As Integer = 400
Public Const FontWeight_Medium As Integer = 500
Public Const FontWeight_SemiBold As Integer = 600
Public Const FontWeight_Bold As Integer = 700
Public Const FontWeight_ExtraBold As Integer = 800
Public Const FontWeight_Heavy As Integer = 900

' --- ReadingOrder ---
Public Const ReadingOrder_Context As Integer = 0
Public Const ReadingOrder_LeftToRight As Integer = 1
Public Const ReadingOrder_RightToLeft As Integer = 2

' --- ViewMode ---
Public Const ViewMode_None As Integer = -1
Public Const ViewMode_Design As Integer = 0
Public Const ViewMode_Form As Integer = 1
Public Const ViewMode_Datasheet As Integer = 2
Public Const ViewMode_Layout As Integer = 7

' --- ShadowStyle ---
Public Const ShadowStyle_None As Integer = 0
Public Const ShadowStyle_DropShadow As Integer = 5

' --- QuickStyle ---
Public Const QuickStyle_None As Integer = 0
Public Const QuickStyle_ModernBorder As Integer = 28

' --- ThemeColorIndex ---
Public Const ThemeColorIndex_None As Integer = -1
Public Const ThemeColorIndex_Text1 As Integer = 0
Public Const ThemeColorIndex_Background1 As Integer = 1
Public Const ThemeColorIndex_Accent1 As Integer = 4
