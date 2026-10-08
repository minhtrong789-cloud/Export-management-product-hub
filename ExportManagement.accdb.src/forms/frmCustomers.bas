Version =20
VersionRequired =20
Begin Form
    AutoResize = NotDefault
    PopUp = NotDefault
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    DefaultView =0
    ScrollBars =0
    ViewsAllowed =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =15810
    DatasheetFontHeight =11
    ItemSuffix =11
    Left =5530
    Top =3030
    Right =13220
    Bottom =9490
    OnUnload ="[Event Procedure]"
    RecSrcDt = Begin
        0xad0c829f7796e640
    End
    Caption ="Customers Master Hub"
    OnOpen ="[Event Procedure]"
    OnClose ="[Event Procedure]"
    DatasheetFontName ="Segoe UI"
    OnLoad ="[Event Procedure]"
    AllowDatasheetView =0
    FilterOnLoad =0
    NavigationCaption ="Customers"
    ShowPageMargins =0
    DisplayOnSharePointSite =1
    AllowLayoutView =0
    DatasheetAlternateBackColor =15921906
    DatasheetGridlinesColor12 =0
    DatasheetBackThemeColorIndex =1
    BorderThemeColorIndex =3
    ThemeFontIndex =1
    ForeThemeColorIndex =0
    AlternateBackThemeColorIndex =1
    AlternateBackShade =95.0
    Begin
        Begin Label
            BackStyle =0
            FontSize =11
            FontName ="Segoe UI"
            ThemeFontIndex =1
            BackThemeColorIndex =1
            BorderThemeColorIndex =0
            BorderTint =50.0
            ForeThemeColorIndex =0
            ForeTint =60.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin Rectangle
            SpecialEffect =3
            BackStyle =0
            BorderLineStyle =0
            BackThemeColorIndex =1
            BorderThemeColorIndex =1
            BorderShade =65.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin CommandButton
            FontSize =11
            FontWeight =400
            FontName ="Segoe UI"
            ForeThemeColorIndex =0
            ForeTint =75.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
            UseTheme =1
            Shape =1
            Gradient =12
            BackThemeColorIndex =4
            BackTint =60.0
            BorderLineStyle =0
            BorderThemeColorIndex =4
            BorderTint =60.0
            ThemeFontIndex =1
            HoverThemeColorIndex =4
            HoverTint =40.0
            PressedThemeColorIndex =4
            PressedShade =75.0
            HoverForeThemeColorIndex =0
            HoverForeTint =75.0
            PressedForeThemeColorIndex =0
            PressedForeTint =75.0
        End
        Begin TextBox
            AddColon = NotDefault
            FELineBreak = NotDefault
            BorderLineStyle =0
            LabelX =-1800
            FontSize =11
            FontName ="Segoe UI"
            AsianLineBreak =1
            BackThemeColorIndex =1
            BorderThemeColorIndex =1
            BorderShade =65.0
            ThemeFontIndex =1
            ForeThemeColorIndex =0
            ForeTint =75.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin Subform
            BorderLineStyle =0
            BorderThemeColorIndex =1
            GridlineThemeColorIndex =1
            GridlineShade =65.0
            BorderShade =65.0
            ShowPageHeaderAndPageFooter =1
        End
        Begin ToggleButton
            FontName ="Segoe UI"
            ForeThemeColorIndex =0
            ForeTint =75.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
            UseTheme =1
            Shape =2
            Bevel =1
            BackThemeColorIndex =4
            BackTint =60.0
            OldBorderStyle =0
            BorderLineStyle =0
            BorderThemeColorIndex =4
            BorderTint =60.0
            ThemeFontIndex =1
            HoverThemeColorIndex =4
            HoverTint =40.0
            PressedThemeColorIndex =4
            PressedShade =75.0
            HoverForeThemeColorIndex =0
            HoverForeTint =75.0
            PressedForeThemeColorIndex =1
        End
        Begin Section
            Height =8640
            Name ="Detail"
            AutoHeight =1
            AlternateBackColor =16777215
            AlternateBackShade =95.0
            Begin
                Begin Rectangle
                    SpecialEffect =0
                    BackStyle =1
                    OldBorderStyle =0
                    OverlapFlags =93
                    Width =15810
                    BackColor =8272943
                    BorderColor =8272943
                    Name ="boxHeaderNavy"
                    HorizontalAnchor =2
                    LayoutCachedWidth =15810
                    LayoutCachedHeight =720
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                End
                Begin Label
                    OverlapFlags =215
                    TextAlign =1
                    Left =288
                    Top =144
                    Width =2880
                    Height =432
                    FontSize =16
                    FontWeight =700
                    LeftMargin =72
                    TopMargin =30
                    BorderColor =16777215
                    ForeColor =16777215
                    Name ="lblTitle"
                    Caption ="CUSTOMER HUB"
                    LayoutCachedLeft =288
                    LayoutCachedTop =144
                    LayoutCachedWidth =3168
                    LayoutCachedHeight =576
                    ThemeFontIndex =-1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                End
                Begin TextBox
                    BorderWidth =1
                    OverlapFlags =215
                    TextAlign =1
                    IMESentenceMode =3
                    Left =3888
                    Top =144
                    Width =3024
                    Height =432
                    TabIndex =1
                    LeftMargin =72
                    TopMargin =30
                    BorderColor =14737632
                    ForeColor =8272943
                    Name ="txtSearch"
                    Format ="@;\" Search customers...\""
                    OnKeyDown ="[Event Procedure]"
                    ControlTipText ="Search by Customer Name, Contact or Phone (Press Enter)"

                    LayoutCachedLeft =3888
                    LayoutCachedTop =144
                    LayoutCachedWidth =6912
                    LayoutCachedHeight =576
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                End
                Begin CommandButton
                    OverlapFlags =215
                    PictureType =2
                    Left =7056
                    Top =144
                    Width =576
                    Height =432
                    FontSize =9
                    FontWeight =700
                    ForeColor =16777215
                    Name ="btnClearFilter"
                    Caption ="CLER"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Clear current search keyword"
                    Picture ="clear-filter"

                    LayoutCachedLeft =7056
                    LayoutCachedTop =144
                    LayoutCachedWidth =7632
                    LayoutCachedHeight =576
                    ForeThemeColorIndex =-1
                    Shape =0
                    Gradient =0
                    BackColor =8272943
                    BackThemeColorIndex =-1
                    OldBorderStyle =0
                    BorderColor =8272943
                    BorderThemeColorIndex =-1
                    ThemeFontIndex =-1
                    HoverColor =9852488
                    HoverThemeColorIndex =-1
                    PressedColor =5580060
                    PressedThemeColorIndex =-1
                    HoverForeColor =16777215
                    HoverForeThemeColorIndex =-1
                    PressedForeColor =16777215
                    PressedForeThemeColorIndex =-1
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =215
                    PictureType =2
                    Left =9216
                    Top =144
                    Width =576
                    Height =432
                    FontSize =9
                    FontWeight =700
                    TabIndex =2
                    ForeColor =16777215
                    Name ="cmdRefresh"
                    Caption ="REFR"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Refresh dataset from database"
                    Picture ="refresh"

                    LayoutCachedLeft =9216
                    LayoutCachedTop =144
                    LayoutCachedWidth =9792
                    LayoutCachedHeight =576
                    ForeThemeColorIndex =-1
                    Shape =0
                    Gradient =0
                    BackColor =8272943
                    BackThemeColorIndex =-1
                    OldBorderStyle =0
                    BorderColor =8272943
                    BorderThemeColorIndex =-1
                    ThemeFontIndex =-1
                    HoverColor =9852488
                    HoverThemeColorIndex =-1
                    PressedColor =5580060
                    PressedThemeColorIndex =-1
                    HoverForeColor =16777215
                    HoverForeThemeColorIndex =-1
                    PressedForeColor =16777215
                    PressedForeThemeColorIndex =-1
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =215
                    Left =9936
                    Top =144
                    Height =432
                    FontSize =9
                    FontWeight =700
                    TabIndex =3
                    ForeColor =16777215
                    Name ="cmdSave"
                    Caption ="[ SAVE ]"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Commit and save pending changes"

                    LayoutCachedLeft =9936
                    LayoutCachedTop =144
                    LayoutCachedWidth =11376
                    LayoutCachedHeight =576
                    ForeThemeColorIndex =-1
                    Shape =0
                    Gradient =0
                    BackColor =6775289
                    BackThemeColorIndex =-1
                    OldBorderStyle =0
                    BorderColor =6775289
                    BorderThemeColorIndex =-1
                    ThemeFontIndex =-1
                    HoverColor =8354554
                    HoverThemeColorIndex =-1
                    PressedColor =5327840
                    PressedThemeColorIndex =-1
                    HoverForeColor =16777215
                    HoverForeThemeColorIndex =-1
                    PressedForeColor =16777215
                    PressedForeThemeColorIndex =-1
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =215
                    PictureType =2
                    Left =15206
                    Top =144
                    Width =518
                    Height =432
                    FontSize =9
                    FontWeight =700
                    TabIndex =4
                    ForeColor =16777215
                    Name ="cmdMaximize"
                    Caption ="MAXM"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Maximize / Restore workspace view"
                    Picture ="max-window-filled"

                    LayoutCachedLeft =15206
                    LayoutCachedTop =144
                    LayoutCachedWidth =15724
                    LayoutCachedHeight =576
                    ForeThemeColorIndex =-1
                    Shape =0
                    Gradient =0
                    BackColor =8272943
                    BackThemeColorIndex =-1
                    OldBorderStyle =0
                    BorderColor =8272943
                    BorderThemeColorIndex =-1
                    ThemeFontIndex =-1
                    HoverColor =9852488
                    HoverThemeColorIndex =-1
                    PressedColor =5580060
                    PressedThemeColorIndex =-1
                    HoverForeColor =16777215
                    HoverForeThemeColorIndex =-1
                    PressedForeColor =16777215
                    PressedForeThemeColorIndex =-1
                    Overlaps =1
                End
                Begin Subform
                    OverlapFlags =215
                    OldBorderStyle =0
                    Top =720
                    Width =15810
                    Height =7920
                    TabIndex =5
                    BorderColor =14737632
                    Name ="sfWorkspace"
                    SourceObject ="Form.frmCustomerGridView"
                    HorizontalAnchor =2
                    VerticalAnchor =2

                    LayoutCachedTop =720
                    LayoutCachedWidth =15810
                    LayoutCachedHeight =8640
                    BorderThemeColorIndex =-1
                End
                Begin CommandButton
                    OverlapFlags =215
                    Left =3312
                    Top =144
                    Width =504
                    Height =432
                    FontSize =9
                    FontWeight =700
                    TabIndex =6
                    ForeColor =16777215
                    Name ="cmdSelectAll"
                    Caption ="ALL"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Select or Deselect all customers"

                    LayoutCachedLeft =3312
                    LayoutCachedTop =144
                    LayoutCachedWidth =3816
                    LayoutCachedHeight =576
                    ForeThemeColorIndex =-1
                    Shape =0
                    Gradient =0
                    BackColor =8272943
                    BackThemeColorIndex =-1
                    OldBorderStyle =0
                    BorderColor =8272943
                    BorderThemeColorIndex =-1
                    ThemeFontIndex =-1
                    HoverColor =9852488
                    HoverThemeColorIndex =-1
                    PressedColor =5580060
                    PressedThemeColorIndex =-1
                    HoverForeColor =16777215
                    HoverForeThemeColorIndex =-1
                    PressedForeColor =16777215
                    PressedForeThemeColorIndex =-1
                    Overlaps =1
                End
                Begin ToggleButton
                    OverlapFlags =215
                    PictureType =2
                    Left =14544
                    Top =144
                    Width =576
                    Height =432
                    FontSize =9
                    FontWeight =600
                    TabIndex =7
                    ForeColor =5197311
                    Name ="togTrashbin"
                    Caption ="TRSH"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Switch between Active and Trashbin archive"
                    Picture ="trash"
                    HorizontalAnchor =1

                    LayoutCachedLeft =14544
                    LayoutCachedTop =144
                    LayoutCachedWidth =15120
                    LayoutCachedHeight =576
                    ForeThemeColorIndex =-1
                    Shape =0
                    Bevel =0
                    BackColor =8272943
                    BackThemeColorIndex =-1
                    BorderWidth =1
                    BorderColor =13421772
                    BorderThemeColorIndex =-1
                    ThemeFontIndex =-1
                    HoverColor =10111547
                    HoverThemeColorIndex =-1
                    PressedColor =5055002
                    PressedThemeColorIndex =-1
                    HoverForeColor =16777215
                    HoverForeThemeColorIndex =-1
                    PressedForeThemeColorIndex =-1
                    Overlaps =1
                End
            End
        End
    End
End
CodeBehindForm
' See "frmCustomers.cls"
