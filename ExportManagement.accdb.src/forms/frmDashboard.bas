Version =20
VersionRequired =20
Begin Form
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
    Width =18000
    DatasheetFontHeight =11
    ItemSuffix =99
    Left =5240
    Top =3040
    Right =13220
    Bottom =9490
    OnUnload ="[Event Procedure]"
    RecSrcDt = Begin
        0x336dff8a357be640
    End
    Caption ="Dashboard"
    OnOpen ="[Event Procedure]"
    DatasheetFontName ="Segoe UI"
    OnResize ="[Event Procedure]"
    OnLoad ="[Event Procedure]"
    OnError ="[Event Procedure]"
    AllowDatasheetView =0
    FilterOnLoad =0
    SplitFormOrientation =1
    NavigationCaption ="Dashboard"
    SplitFormOrientation =1
    ShowPageMargins =0
    DisplayOnSharePointSite =1
    AllowLayoutView =0
    DatasheetAlternateBackColor =15921906
    DatasheetGridlinesColor12 =0
    FitToScreen =255
    DatasheetBackThemeColorIndex =1
    BorderThemeColorIndex =3
    ThemeFontIndex =1
    ForeThemeColorIndex =0
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
        Begin Image
            BackStyle =0
            OldBorderStyle =0
            BorderLineStyle =0
            SizeMode =3
            PictureAlignment =2
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
        Begin Subform
            BorderLineStyle =0
            BorderThemeColorIndex =1
            GridlineThemeColorIndex =1
            GridlineShade =65.0
            BorderShade =65.0
            ShowPageHeaderAndPageFooter =1
        End
        Begin EmptyCell
            Height =240
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin NavigationControl
            BorderWidth =1
            BorderLineStyle =0
            ForeColor =4138256
            BackThemeColorIndex =1
            BorderThemeColorIndex =1
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin NavigationButton
            GridlineThemeColorIndex =1
            GridlineShade =65.0
            HoverThemeColorIndex =2
            HoverTint =20.0
            PressedThemeColorIndex =2
            PressedTint =60.0
            HoverForeThemeColorIndex =0
            HoverForeTint =75.0
            PressedForeThemeColorIndex =0
            PressedForeTint =75.0
            BackThemeColorIndex =1
            OldBorderStyle =0
            BorderLineStyle =0
            BorderThemeColorIndex =3
            BorderShade =90.0
            ThemeFontIndex =1
            TextFontCharSet =177
            TextFontFamily =0
            FontName ="Segoe UI"
            FontWeight =400
            FontSize =11
            ForeThemeColorIndex =0
            ForeTint =75.0
        End
        Begin FormHeader
            Visible = NotDefault
            Height =0
            Name ="FormHeader"
            AutoHeight =255
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackTint =20.0
        End
        Begin PageHeader
            DisplayWhen =1
            Height =0
            Name ="PageHeaderSection"
            AutoHeight =255
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
        End
        Begin Section
            CanGrow = NotDefault
            Height =8640
            Name ="Detail"
            AutoHeight =255
            AlternateBackColor =16777215
            AlternateBackShade =95.0
            Begin
                Begin Rectangle
                    SpecialEffect =0
                    BackStyle =1
                    OldBorderStyle =0
                    OverlapFlags =93
                    Width =2160
                    Height =8640
                    BackColor =15461371
                    BorderColor =15461371
                    Name ="recNavi"
                    VerticalAnchor =2
                    LayoutCachedWidth =2160
                    LayoutCachedHeight =8640
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                End
                Begin Subform
                    CanShrink = NotDefault
                    OverlapFlags =85
                    OldBorderStyle =0
                    Left =2189
                    Width =15811
                    Height =8640
                    TabIndex =1
                    BorderColor =14737632
                    Name ="sfDashboard"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    HorizontalAnchor =2
                    VerticalAnchor =2

                    LayoutCachedLeft =2189
                    LayoutCachedWidth =18000
                    LayoutCachedHeight =8640
                    BorderThemeColorIndex =-1
                End
                Begin NavigationControl
                    OldBorderStyle =0
                    BorderWidth =0
                    OverlapFlags =215
                    BackStyle =0
                    Left =70
                    Top =2016
                    Width =2040
                    Height =4752
                    BackColor =0
                    BorderColor =0
                    Name ="navctrlDashboard"

                    LayoutCachedLeft =70
                    LayoutCachedTop =2016
                    LayoutCachedWidth =2110
                    LayoutCachedHeight =6768
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    NavigationSubform ="sfDashboard"
                    Span =1
                    Begin
                        Begin NavigationButton
                            Enabled = NotDefault
                            OverlapFlags =215
                            Left =70
                            Top =2050
                            Width =2040
                            Height =720
                            TabIndex =1
                            ForeColor =0
                            Name ="navbtnOrders"
                            Caption ="ORDERS"
                            OnMouseDown ="=RouteSelection()"
                            GroupTable =1
                            LeftPadding =0
                            RightPadding =0

                            LayoutCachedLeft =70
                            LayoutCachedTop =2050
                            LayoutCachedWidth =2110
                            LayoutCachedHeight =2770
                            LayoutGroup =1
                            NavigationTargetType =32768
                            HoverColor =14211320
                            HoverThemeColorIndex =-1
                            HoverTint =100.0
                            PressedColor =8272943
                            PressedThemeColorIndex =-1
                            PressedTint =100.0
                            HoverForeColor =8272943
                            HoverForeThemeColorIndex =-1
                            HoverForeTint =85.0
                            PressedForeColor =16777215
                            PressedForeThemeColorIndex =-1
                            PressedForeTint =100.0
                            BackColor =15461371
                            BackThemeColorIndex =-1
                            BorderColor =15461371
                            BorderThemeColorIndex =-1
                            ThemeFontIndex =-1
                            GroupTable =1
                            TextFontCharSet =0
                            TextFontFamily =34
                            FontWeight =700
                            FontSize =9
                            ForeThemeColorIndex =-1
                            ForeTint =85.0
                            Overlaps =1
                        End
                        Begin NavigationButton
                            Enabled = NotDefault
                            OverlapFlags =215
                            Left =70
                            Top =2830
                            Width =2040
                            Height =720
                            TabIndex =2
                            ForeColor =0
                            Name ="navbtnCustomers"
                            Caption ="CUSTOMERS"
                            OnMouseDown ="=RouteSelection()"
                            Tag ="CUSTOMERS"
                            GroupTable =1
                            LeftPadding =0
                            RightPadding =0

                            LayoutCachedLeft =70
                            LayoutCachedTop =2830
                            LayoutCachedWidth =2110
                            LayoutCachedHeight =3550
                            RowStart =1
                            RowEnd =1
                            LayoutGroup =1
                            NavigationTargetType =32768
                            HoverColor =14211320
                            HoverThemeColorIndex =-1
                            HoverTint =100.0
                            PressedColor =8272943
                            PressedThemeColorIndex =-1
                            PressedTint =100.0
                            HoverForeColor =8272943
                            HoverForeThemeColorIndex =-1
                            HoverForeTint =85.0
                            PressedForeColor =16777215
                            PressedForeThemeColorIndex =-1
                            PressedForeTint =100.0
                            BackColor =15461371
                            BackThemeColorIndex =-1
                            BorderColor =15461371
                            BorderThemeColorIndex =-1
                            ThemeFontIndex =-1
                            GroupTable =1
                            TextFontCharSet =0
                            TextFontFamily =34
                            FontWeight =700
                            FontSize =9
                            ForeThemeColorIndex =-1
                            ForeTint =85.0
                            Overlaps =1
                        End
                        Begin NavigationButton
                            OverlapFlags =215
                            Left =70
                            Top =3610
                            Width =2040
                            Height =720
                            ForeColor =0
                            Name ="navbtnProducts"
                            Caption ="PRODUCTS"
                            OnMouseDown ="=RouteSelection()"
                            Tag ="PRODUCTS"
                            GroupTable =1
                            LeftPadding =0
                            RightPadding =0

                            LayoutCachedLeft =70
                            LayoutCachedTop =3610
                            LayoutCachedWidth =2110
                            LayoutCachedHeight =4330
                            RowStart =2
                            RowEnd =2
                            LayoutGroup =1
                            NavigationTargetName ="frmProducts"
                            NavigationTargetType =32768
                            HoverColor =14211320
                            HoverThemeColorIndex =-1
                            HoverTint =100.0
                            PressedColor =8272943
                            PressedThemeColorIndex =-1
                            PressedTint =100.0
                            HoverForeColor =8272943
                            HoverForeThemeColorIndex =-1
                            HoverForeTint =85.0
                            PressedForeColor =16777215
                            PressedForeThemeColorIndex =-1
                            PressedForeTint =100.0
                            BackColor =15461371
                            BackThemeColorIndex =-1
                            BorderColor =15461371
                            BorderThemeColorIndex =-1
                            ThemeFontIndex =-1
                            GroupTable =1
                            TextFontCharSet =0
                            TextFontFamily =34
                            FontWeight =700
                            FontSize =9
                            ForeThemeColorIndex =-1
                            ForeTint =85.0
                            Overlaps =1
                        End
                        Begin NavigationButton
                            Enabled = NotDefault
                            OverlapFlags =215
                            Left =70
                            Top =4390
                            Width =2040
                            Height =720
                            TabIndex =3
                            ForeColor =0
                            Name ="navbtnOthers"
                            Caption ="OTHERS"
                            OnMouseDown ="=RouteSelection()"
                            GroupTable =1
                            LeftPadding =0
                            RightPadding =0

                            LayoutCachedLeft =70
                            LayoutCachedTop =4390
                            LayoutCachedWidth =2110
                            LayoutCachedHeight =5110
                            RowStart =3
                            RowEnd =3
                            LayoutGroup =1
                            HoverColor =14211320
                            HoverThemeColorIndex =-1
                            HoverTint =100.0
                            PressedColor =8272943
                            PressedThemeColorIndex =-1
                            PressedTint =100.0
                            HoverForeColor =8272943
                            HoverForeThemeColorIndex =-1
                            HoverForeTint =85.0
                            PressedForeColor =16777215
                            PressedForeThemeColorIndex =-1
                            PressedForeTint =100.0
                            BackColor =15461371
                            BackThemeColorIndex =-1
                            BorderColor =15461371
                            BorderThemeColorIndex =-1
                            ThemeFontIndex =-1
                            GroupTable =1
                            TextFontCharSet =0
                            TextFontFamily =34
                            FontWeight =700
                            FontSize =9
                            ForeThemeColorIndex =-1
                            ForeTint =85.0
                            Overlaps =1
                        End
                        Begin NavigationButton
                            OverlapFlags =215
                            Left =70
                            Top =5170
                            Width =2040
                            Height =720
                            TabIndex =4
                            ForeColor =0
                            Name ="NavigationButton1"
                            Caption ="[Add New]"
                            GroupTable =1
                            LeftPadding =0
                            RightPadding =0

                            LayoutCachedLeft =70
                            LayoutCachedTop =5170
                            LayoutCachedWidth =2110
                            LayoutCachedHeight =5890
                            RowStart =4
                            RowEnd =4
                            LayoutGroup =1
                            DefaultInsertButton =1
                            HoverColor =14211320
                            HoverThemeColorIndex =-1
                            HoverTint =100.0
                            PressedColor =8272943
                            PressedThemeColorIndex =-1
                            PressedTint =100.0
                            HoverForeColor =8272943
                            HoverForeThemeColorIndex =-1
                            HoverForeTint =85.0
                            PressedForeColor =16777215
                            PressedForeThemeColorIndex =-1
                            PressedForeTint =100.0
                            BackColor =15461371
                            BackThemeColorIndex =-1
                            BorderColor =15461371
                            BorderThemeColorIndex =-1
                            ThemeFontIndex =-1
                            GroupTable =1
                            TextFontCharSet =0
                            TextFontFamily =34
                            FontWeight =700
                            FontSize =9
                            ForeThemeColorIndex =-1
                            ForeTint =85.0
                            Overlaps =1
                        End
                    End
                End
                Begin Image
                    PictureType =2
                    Left =360
                    Top =173
                    BorderColor =13421772
                    Name ="imaSoftWareLogo"
                    Picture ="ExportManagement_RmBG"

                    LayoutCachedLeft =360
                    LayoutCachedTop =173
                    LayoutCachedWidth =1800
                    LayoutCachedHeight =1613
                    TabIndex =3
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                End
                Begin CommandButton
                    OverlapFlags =215
                    PictureType =2
                    Left =72
                    Top =8064
                    Width =461
                    Height =504
                    FontSize =10
                    FontWeight =600
                    TabIndex =2
                    ForeColor =6176175
                    Name ="cmdCallAccessWindow"
                    OnClick ="[Event Procedure]"
                    Picture ="back_40x40"
                    VerticalAnchor =1
                    BackStyle =0

                    LayoutCachedLeft =72
                    LayoutCachedTop =8064
                    LayoutCachedWidth =533
                    LayoutCachedHeight =8568
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    Shape =0
                    Gradient =0
                    BackColor =0
                    BackThemeColorIndex =-1
                    OldBorderStyle =0
                    BorderColor =0
                    BorderThemeColorIndex =-1
                    ThemeFontIndex =-1
                    HoverColor =14211320
                    HoverThemeColorIndex =-1
                    PressedColor =8355711
                    PressedThemeColorIndex =-1
                    HoverForeColor =6176175
                    HoverForeThemeColorIndex =-1
                    HoverForeTint =100.0
                    PressedForeColor =6176175
                    PressedForeThemeColorIndex =-1
                    PressedForeTint =100.0
                    Overlaps =1
                End
            End
        End
        Begin PageFooter
            DisplayWhen =1
            Height =0
            Name ="PageFooterSection"
            AutoHeight =255
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
        End
        Begin FormFooter
            Visible = NotDefault
            Height =0
            Name ="FormFooter"
            AutoHeight =255
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
        End
    End
End
CodeBehindForm
' See "frmDashboard.cls"
