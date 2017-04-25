Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    AllowAdditions = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    ScrollBars =2
    TabularFamily =124
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    Cycle =1
    GridX =24
    GridY =24
    Width =13320
    DatasheetFontHeight =9
    ItemSuffix =65
    Left =60
    Top =1350
    Right =11700
    Bottom =10395
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0xb5100b474c2ee340
    End
    RecordSource ="qry_Quadrat_Transect"
    Caption ="frm_Canopy_Transect"
    OnCurrent ="[Event Procedure]"
    OnOpen ="[Event Procedure]"
    OnClose ="[Event Procedure]"
    DatasheetFontName ="Arial"
    PrtMip = Begin
        0xa0050000a0050000a0050000a005000000000000201c0000e010000001000000 ,
        0x010000006801000000000000a10700000100000001000000
    End
    FilterOnLoad =0
    DatasheetGridlinesColor12 =12632256
    Begin
        Begin Label
            BackStyle =0
            BackColor =-2147483633
            ForeColor =-2147483630
        End
        Begin Rectangle
            SpecialEffect =3
            BackStyle =0
            BorderLineStyle =0
        End
        Begin Image
            BackStyle =0
            OldBorderStyle =0
            BorderLineStyle =0
            PictureAlignment =2
        End
        Begin CommandButton
            FontSize =8
            FontWeight =400
            FontName ="MS Sans Serif"
            BorderLineStyle =0
        End
        Begin OptionButton
            SpecialEffect =2
            BorderLineStyle =0
            LabelX =230
            LabelY =-30
        End
        Begin CheckBox
            SpecialEffect =2
            BorderLineStyle =0
            LabelX =230
            LabelY =-30
        End
        Begin OptionGroup
            SpecialEffect =3
            BorderLineStyle =0
        End
        Begin BoundObjectFrame
            SpecialEffect =2
            OldBorderStyle =0
            BorderLineStyle =0
            BackStyle =0
        End
        Begin TextBox
            FELineBreak = NotDefault
            SpecialEffect =2
            BorderLineStyle =0
            BackColor =-2147483643
            ForeColor =-2147483640
            AsianLineBreak =255
        End
        Begin ListBox
            SpecialEffect =2
            BorderLineStyle =0
            BackColor =-2147483643
            ForeColor =-2147483640
        End
        Begin ComboBox
            SpecialEffect =2
            BorderLineStyle =0
            BackColor =-2147483643
            ForeColor =-2147483640
        End
        Begin Subform
            SpecialEffect =2
            BorderLineStyle =0
        End
        Begin UnboundObjectFrame
            SpecialEffect =2
            OldBorderStyle =1
        End
        Begin ToggleButton
            FontSize =8
            FontWeight =400
            FontName ="MS Sans Serif"
            BorderLineStyle =0
        End
        Begin Tab
            BackStyle =0
            BorderLineStyle =0
        End
        Begin EmptyCell
            Height =240
            GridlineColor =12632256
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin FormHeader
            Height =0
            BackColor =-2147483633
            Name ="FormHeader"
        End
        Begin Section
            CanGrow = NotDefault
            Height =8325
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin Subform
                    Visible = NotDefault
                    OverlapFlags =93
                    Left =5820
                    Top =1020
                    Width =7200
                    Height =6717
                    TabIndex =55
                    Name ="fsub_Species_2008"
                    SourceObject ="Form.fsub_Species_2008"
                    LinkChildFields ="Transect_ID"
                    LinkMasterFields ="Transect_ID"

                    LayoutCachedLeft =5820
                    LayoutCachedTop =1020
                    LayoutCachedWidth =13020
                    LayoutCachedHeight =7737
                End
                Begin Subform
                    Visible = NotDefault
                    OverlapFlags =255
                    Left =5580
                    Top =1020
                    Width =7320
                    Height =6598
                    TabIndex =56
                    Name ="fsub_Species_2009"
                    SourceObject ="Form.fsub_Species_2009"
                    LinkChildFields ="Transect_ID"
                    LinkMasterFields ="Transect_ID"

                    LayoutCachedLeft =5580
                    LayoutCachedTop =1020
                    LayoutCachedWidth =12900
                    LayoutCachedHeight =7618
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =12360
                    Top =360
                    Width =630
                    Height =180
                    ColumnWidth =2310
                    Name ="Transect_ID"
                    ControlSource ="Transect_ID"
                    StatusBarText ="Unique record identifier - primary key"

                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =12300
                    Top =660
                    Width =630
                    Height =180
                    ColumnWidth =2310
                    TabIndex =1
                    Name ="Event_ID"
                    ControlSource ="Event_ID"
                    StatusBarText ="M. Link to tbl_Locations (Loc_ID)"

                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    SpecialEffect =0
                    OldBorderStyle =0
                    OverlapFlags =87
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =900
                    Top =60
                    Width =1080
                    ColumnWidth =465
                    FontWeight =700
                    TabIndex =2
                    Name ="Transect"
                    ControlSource ="Transect"
                    StatusBarText ="Transect number - 1, 2, or 3"

                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =60
                            Top =60
                            Width =840
                            Height =240
                            FontWeight =700
                            Name ="Transect_Label"
                            Caption ="Transect"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =87
                    IMESentenceMode =3
                    Left =4620
                    Top =60
                    Width =960
                    ColumnWidth =1035
                    TabIndex =3
                    Name ="tbxVisitDate"
                    ControlSource ="Start_Time"
                    Format ="Short Time"
                    StatusBarText ="Date of visit."
                    InputMask ="00:00;0;_"
                    OnKeyDown ="[Event Procedure]"

                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =3660
                            Top =60
                            Width =960
                            Height =240
                            FontWeight =700
                            Name ="Visit_Date_Label"
                            Caption ="Start Time"
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =2220
                    Top =60
                    Width =306
                    Height =306
                    TabIndex =4
                    Name ="btnPrevious"
                    Caption ="Command14"
                    OnClick ="[Event Procedure]"
                    PictureData = Begin
                        0x2800000010000000100000000100040000000000800000000000000000000000 ,
                        0x0000000000000000000000000000800000800000008080008000000080008000 ,
                        0x8080000080808000c0c0c0000000ff00c0c0c00000ffff00ff000000c0c0c000 ,
                        0xffff0000ffffff00dadadadadadadadaadadadadadadadaddadadadadadadada ,
                        0xadadadadadadadaddadadadad1dadadaadadadad11adadaddadadad111dadada ,
                        0xadadad1111adadaddadad11111dadadaadadad1111adadaddadadad111dadada ,
                        0xadadadad11adadaddadadadad1dadadaadadadadadadadaddadadadadadadada ,
                        0xadadadadadadadad000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000
                    End
                    OnKeyDown ="[Event Procedure]"
                    ObjectPalette = Begin
                        0x000301000000000000000000
                    End
                    ControlTipText ="Previous Record"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =2580
                    Top =60
                    Width =306
                    Height =306
                    TabIndex =5
                    Name ="btnNext"
                    Caption ="Command15"
                    OnClick ="[Event Procedure]"
                    PictureData = Begin
                        0x2800000010000000100000000100040000000000800000000000000000000000 ,
                        0x0000000000000000000000000000800000800000008080008000000080008000 ,
                        0x8080000080808000c0c0c0000000ff00c0c0c00000ffff00ff000000c0c0c000 ,
                        0xffff0000ffffff00dadadadadadadadaadadadadadadadaddadadadadadadada ,
                        0xadadadadadadadaddadada1adadadadaadadad11adadadaddadada111adadada ,
                        0xadadad1111adadaddadada11111adadaadadad1111adadaddadada111adadada ,
                        0xadadad11adadadaddadada1adadadadaadadadadadadadaddadadadadadadada ,
                        0xadadadadadadadad000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000
                    End
                    OnKeyDown ="[Event Procedure]"
                    ObjectPalette = Begin
                        0x000301000000000000000000
                    End
                    ControlTipText ="Next Record"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =12360
                    Top =60
                    Width =840
                    Height =180
                    TabIndex =6
                    Name ="GPS_Time"
                    ControlSource ="GPS_Time"
                    Format ="Long Time"
                    StatusBarText ="Recording time"

                    LayoutCachedLeft =12360
                    LayoutCachedTop =60
                    LayoutCachedWidth =13200
                    LayoutCachedHeight =240
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =6900
                    Top =60
                    Width =5340
                    Height =363
                    TabIndex =7
                    Name ="Comments"
                    ControlSource ="Comments"
                    StatusBarText ="Notes"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =5820
                            Top =60
                            Width =915
                            Height =240
                            FontWeight =700
                            Name ="Label32"
                            Caption ="Comments:"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =87
                    IMESentenceMode =3
                    ColumnCount =3
                    ListWidth =1650
                    Left =4500
                    Top =480
                    Width =1620
                    TabIndex =8
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"\";\"\";\"10\";\"100\""
                    Name ="cbxObserver"
                    ControlSource ="Observer"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT tlu_Contacts.Contact_ID, tlu_Contacts.Last_Name, tlu_Contacts.First_Name "
                        "FROM tlu_Contacts; "
                    ColumnWidths ="0;810;839"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =3660
                            Top =480
                            Width =840
                            Height =245
                            FontWeight =700
                            Name ="Observer_Label"
                            Caption ="Observer"
                        End
                    End
                End
                Begin Label
                    OldBorderStyle =1
                    BorderWidth =1
                    OverlapFlags =93
                    TextAlign =2
                    Left =120
                    Top =2880
                    Width =1860
                    Height =480
                    FontWeight =700
                    Name ="Label55"
                    Caption ="Microhabitat"
                    LayoutCachedLeft =120
                    LayoutCachedTop =2880
                    LayoutCachedWidth =1980
                    LayoutCachedHeight =3360
                End
                Begin Label
                    OldBorderStyle =1
                    BorderWidth =1
                    OverlapFlags =95
                    TextAlign =2
                    Left =1980
                    Top =3120
                    Width =778
                    Height =240
                    FontWeight =700
                    Name ="Label57"
                    Caption ="Q1"
                    LayoutCachedLeft =1980
                    LayoutCachedTop =3120
                    LayoutCachedWidth =2758
                    LayoutCachedHeight =3360
                End
                Begin Label
                    OldBorderStyle =1
                    BorderWidth =1
                    OverlapFlags =95
                    TextAlign =2
                    Left =2760
                    Top =3120
                    Width =780
                    Height =240
                    FontWeight =700
                    Name ="Label73"
                    Caption ="Q2"
                    LayoutCachedLeft =2760
                    LayoutCachedTop =3120
                    LayoutCachedWidth =3540
                    LayoutCachedHeight =3360
                End
                Begin Label
                    OldBorderStyle =1
                    BorderWidth =1
                    OverlapFlags =95
                    TextAlign =2
                    Left =3539
                    Top =3120
                    Width =780
                    Height =240
                    FontWeight =700
                    Name ="Label74"
                    Caption ="Q3"
                    LayoutCachedLeft =3539
                    LayoutCachedTop =3120
                    LayoutCachedWidth =4319
                    LayoutCachedHeight =3360
                End
                Begin Label
                    OldBorderStyle =1
                    BorderWidth =1
                    OverlapFlags =87
                    TextAlign =2
                    Left =1980
                    Top =2880
                    Width =2340
                    Height =240
                    FontWeight =700
                    Name ="Label76"
                    Caption ="% cover"
                    LayoutCachedLeft =1980
                    LayoutCachedTop =2880
                    LayoutCachedWidth =4320
                    LayoutCachedHeight =3120
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =95
                    IMESentenceMode =3
                    ListRows =21
                    Left =1980
                    Top =3360
                    Width =778
                    TabIndex =9
                    ColumnInfo ="\"\";\"\";\"6\";\"4\""
                    Name ="Dead_Wood_Q1"
                    RowSourceType ="Table/Query"
                    RowSource ="tlu_Cover_Code"
                    StatusBarText ="Dead wood cover percentage quadrat 1"
                    OnChange ="=UpdateMicrohabitat([Screen].[ActiveControl])"

                    LayoutCachedLeft =1980
                    LayoutCachedTop =3360
                    LayoutCachedWidth =2758
                    LayoutCachedHeight =3600
                    Begin
                        Begin Label
                            OldBorderStyle =1
                            OverlapFlags =95
                            Left =120
                            Top =3360
                            Width =1860
                            Height =240
                            Name ="Label289"
                            Caption ="Dead Wood:"
                            LayoutCachedLeft =120
                            LayoutCachedTop =3360
                            LayoutCachedWidth =1980
                            LayoutCachedHeight =3600
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =95
                    IMESentenceMode =3
                    ListRows =21
                    Left =2770
                    Top =3360
                    Width =778
                    TabIndex =10
                    ColumnInfo ="\"\";\"\";\"6\";\"4\""
                    Name ="Dead_Wood_Q2"
                    RowSourceType ="Table/Query"
                    RowSource ="tlu_Cover_Code"
                    StatusBarText ="Dead wood cover percentage quadrat 2"
                    OnChange ="=UpdateMicrohabitat([Screen].[ActiveControl])"

                    LayoutCachedLeft =2770
                    LayoutCachedTop =3360
                    LayoutCachedWidth =3548
                    LayoutCachedHeight =3600
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =95
                    IMESentenceMode =3
                    ListRows =21
                    Left =3550
                    Top =3360
                    Width =778
                    TabIndex =11
                    ColumnInfo ="\"\";\"\";\"6\";\"4\""
                    Name ="Dead_Wood_Q3"
                    RowSourceType ="Table/Query"
                    RowSource ="tlu_Cover_Code"
                    StatusBarText ="Dead wood cover percentage quadrat 3"
                    OnChange ="=UpdateMicrohabitat([Screen].[ActiveControl])"

                    LayoutCachedLeft =3550
                    LayoutCachedTop =3360
                    LayoutCachedWidth =4328
                    LayoutCachedHeight =3600
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =95
                    IMESentenceMode =3
                    ListRows =21
                    Left =1980
                    Top =3600
                    Width =778
                    TabIndex =12
                    ColumnInfo ="\"\";\"\";\"6\";\"4\""
                    Name ="Dung_Q1"
                    RowSourceType ="Table/Query"
                    RowSource ="tlu_Cover_Code"
                    StatusBarText ="Dung cover percentage quadrat 1"
                    OnChange ="=UpdateMicrohabitat([Screen].[ActiveControl])"

                    LayoutCachedLeft =1980
                    LayoutCachedTop =3600
                    LayoutCachedWidth =2758
                    LayoutCachedHeight =3840
                    Begin
                        Begin Label
                            OldBorderStyle =1
                            OverlapFlags =95
                            Left =120
                            Top =3600
                            Width =1860
                            Height =240
                            Name ="Label292"
                            Caption ="Dung"
                            LayoutCachedLeft =120
                            LayoutCachedTop =3600
                            LayoutCachedWidth =1980
                            LayoutCachedHeight =3840
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =95
                    IMESentenceMode =3
                    ListRows =21
                    Left =2770
                    Top =3600
                    Width =778
                    TabIndex =13
                    ColumnInfo ="\"\";\"\";\"6\";\"4\""
                    Name ="Dung_Q2"
                    RowSourceType ="Table/Query"
                    RowSource ="tlu_Cover_Code"
                    StatusBarText ="Dung cover percentage quadrat 2"
                    OnChange ="=UpdateMicrohabitat([Screen].[ActiveControl])"

                    LayoutCachedLeft =2770
                    LayoutCachedTop =3600
                    LayoutCachedWidth =3548
                    LayoutCachedHeight =3840
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =95
                    IMESentenceMode =3
                    ListRows =21
                    Left =3550
                    Top =3600
                    Width =778
                    TabIndex =14
                    ColumnInfo ="\"\";\"\";\"6\";\"4\""
                    Name ="Dung_Q3"
                    RowSourceType ="Table/Query"
                    RowSource ="tlu_Cover_Code"
                    StatusBarText ="Dung cover percentage quadrat 3"
                    OnChange ="=UpdateMicrohabitat([Screen].[ActiveControl])"

                    LayoutCachedLeft =3550
                    LayoutCachedTop =3600
                    LayoutCachedWidth =4328
                    LayoutCachedHeight =3840
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =95
                    IMESentenceMode =3
                    ListRows =21
                    Left =1980
                    Top =3840
                    Width =778
                    TabIndex =15
                    ColumnInfo ="\"\";\"\";\"6\";\"4\""
                    Name ="Fungus_Q1"
                    RowSourceType ="Table/Query"
                    RowSource ="tlu_Cover_Code"
                    StatusBarText ="Fungus cover percentage quadrat 1"
                    OnChange ="=UpdateMicrohabitat([Screen].[ActiveControl])"

                    LayoutCachedLeft =1980
                    LayoutCachedTop =3840
                    LayoutCachedWidth =2758
                    LayoutCachedHeight =4080
                    Begin
                        Begin Label
                            OldBorderStyle =1
                            OverlapFlags =95
                            Left =120
                            Top =3840
                            Width =1860
                            Height =240
                            Name ="Label295"
                            Caption ="Fungus"
                            LayoutCachedLeft =120
                            LayoutCachedTop =3840
                            LayoutCachedWidth =1980
                            LayoutCachedHeight =4080
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =95
                    IMESentenceMode =3
                    ListRows =21
                    Left =2770
                    Top =3840
                    Width =778
                    TabIndex =16
                    ColumnInfo ="\"\";\"\";\"6\";\"4\""
                    Name ="Fungus_Q2"
                    RowSourceType ="Table/Query"
                    RowSource ="tlu_Cover_Code"
                    StatusBarText ="Fungus cover percentage quadrat 2"
                    OnChange ="=UpdateMicrohabitat([Screen].[ActiveControl])"

                    LayoutCachedLeft =2770
                    LayoutCachedTop =3840
                    LayoutCachedWidth =3548
                    LayoutCachedHeight =4080
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =95
                    IMESentenceMode =3
                    ListRows =21
                    Left =3550
                    Top =3840
                    Width =778
                    TabIndex =17
                    ColumnInfo ="\"\";\"\";\"6\";\"4\""
                    Name ="Fungus_Q3"
                    RowSourceType ="Table/Query"
                    RowSource ="tlu_Cover_Code"
                    StatusBarText ="Fungus cover percentage quadrat 3"
                    OnChange ="=UpdateMicrohabitat([Screen].[ActiveControl])"

                    LayoutCachedLeft =3550
                    LayoutCachedTop =3840
                    LayoutCachedWidth =4328
                    LayoutCachedHeight =4080
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =95
                    IMESentenceMode =3
                    ListRows =21
                    Left =1980
                    Top =4080
                    Width =778
                    TabIndex =18
                    ColumnInfo ="\"\";\"\";\"6\";\"4\""
                    Name ="Lichen_Q1"
                    RowSourceType ="Table/Query"
                    RowSource ="tlu_Cover_Code"
                    StatusBarText ="Lichen cover percentage quadrat 1"
                    OnChange ="=UpdateMicrohabitat([Screen].[ActiveControl])"

                    LayoutCachedLeft =1980
                    LayoutCachedTop =4080
                    LayoutCachedWidth =2758
                    LayoutCachedHeight =4320
                    Begin
                        Begin Label
                            OldBorderStyle =1
                            OverlapFlags =95
                            Left =120
                            Top =4080
                            Width =1860
                            Height =240
                            Name ="Label298"
                            Caption ="Lichen"
                            LayoutCachedLeft =120
                            LayoutCachedTop =4080
                            LayoutCachedWidth =1980
                            LayoutCachedHeight =4320
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =95
                    IMESentenceMode =3
                    ListRows =21
                    Left =2770
                    Top =4080
                    Width =778
                    TabIndex =19
                    ColumnInfo ="\"\";\"\";\"6\";\"4\""
                    Name ="Lichen_Q2"
                    RowSourceType ="Table/Query"
                    RowSource ="tlu_Cover_Code"
                    StatusBarText ="Lichen cover percentage quadrat 2"
                    OnChange ="=UpdateMicrohabitat([Screen].[ActiveControl])"

                    LayoutCachedLeft =2770
                    LayoutCachedTop =4080
                    LayoutCachedWidth =3548
                    LayoutCachedHeight =4320
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =95
                    IMESentenceMode =3
                    ListRows =21
                    Left =3550
                    Top =4080
                    Width =778
                    TabIndex =20
                    ColumnInfo ="\"\";\"\";\"6\";\"4\""
                    Name ="Lichen_Q3"
                    RowSourceType ="Table/Query"
                    RowSource ="tlu_Cover_Code"
                    StatusBarText ="Lichen cover percentage quadrat 3"
                    OnChange ="=UpdateMicrohabitat([Screen].[ActiveControl])"

                    LayoutCachedLeft =3550
                    LayoutCachedTop =4080
                    LayoutCachedWidth =4328
                    LayoutCachedHeight =4320
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =95
                    IMESentenceMode =3
                    ListRows =21
                    Left =1980
                    Top =4320
                    Width =778
                    TabIndex =21
                    ColumnInfo ="\"\";\"\";\"6\";\"4\""
                    Name ="Litter_Duff_Q1"
                    RowSourceType ="Table/Query"
                    RowSource ="tlu_Cover_Code"
                    StatusBarText ="Litter/Duff cover percentage quadrat 1"
                    OnChange ="=UpdateMicrohabitat([Screen].[ActiveControl])"

                    LayoutCachedLeft =1980
                    LayoutCachedTop =4320
                    LayoutCachedWidth =2758
                    LayoutCachedHeight =4560
                    Begin
                        Begin Label
                            OldBorderStyle =1
                            OverlapFlags =95
                            Left =120
                            Top =4320
                            Width =1860
                            Height =240
                            Name ="Label301"
                            Caption ="Litter Duff"
                            LayoutCachedLeft =120
                            LayoutCachedTop =4320
                            LayoutCachedWidth =1980
                            LayoutCachedHeight =4560
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =95
                    IMESentenceMode =3
                    ListRows =21
                    Left =2770
                    Top =4320
                    Width =778
                    TabIndex =22
                    ColumnInfo ="\"\";\"\";\"6\";\"4\""
                    Name ="Litter_Duff_Q2"
                    RowSourceType ="Table/Query"
                    RowSource ="tlu_Cover_Code"
                    StatusBarText ="Litter/Duff cover percentage quadrat 2"
                    OnChange ="=UpdateMicrohabitat([Screen].[ActiveControl])"

                    LayoutCachedLeft =2770
                    LayoutCachedTop =4320
                    LayoutCachedWidth =3548
                    LayoutCachedHeight =4560
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =95
                    IMESentenceMode =3
                    ListRows =21
                    Left =3550
                    Top =4320
                    Width =778
                    TabIndex =23
                    ColumnInfo ="\"\";\"\";\"6\";\"4\""
                    Name ="Litter_Duff_Q3"
                    RowSourceType ="Table/Query"
                    RowSource ="tlu_Cover_Code"
                    StatusBarText ="Litter/Duff cover percentage quadrat 3"
                    OnChange ="=UpdateMicrohabitat([Screen].[ActiveControl])"

                    LayoutCachedLeft =3550
                    LayoutCachedTop =4320
                    LayoutCachedWidth =4328
                    LayoutCachedHeight =4560
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =95
                    IMESentenceMode =3
                    ListRows =21
                    Left =1980
                    Top =4560
                    Width =778
                    TabIndex =24
                    ColumnInfo ="\"\";\"\";\"6\";\"4\""
                    Name ="Root_Bole_Q1"
                    RowSourceType ="Table/Query"
                    RowSource ="tlu_Cover_Code"
                    StatusBarText ="Root/Bole cover percentage quadrat 1"
                    OnChange ="=UpdateMicrohabitat([Screen].[ActiveControl])"

                    LayoutCachedLeft =1980
                    LayoutCachedTop =4560
                    LayoutCachedWidth =2758
                    LayoutCachedHeight =4800
                    Begin
                        Begin Label
                            OldBorderStyle =1
                            OverlapFlags =95
                            Left =120
                            Top =4560
                            Width =1860
                            Height =240
                            Name ="Label304"
                            Caption ="Live root/Bole"
                            LayoutCachedLeft =120
                            LayoutCachedTop =4560
                            LayoutCachedWidth =1980
                            LayoutCachedHeight =4800
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =95
                    IMESentenceMode =3
                    ListRows =21
                    Left =2770
                    Top =4560
                    Width =778
                    TabIndex =25
                    ColumnInfo ="\"\";\"\";\"6\";\"4\""
                    Name ="Root_Bole_Q2"
                    RowSourceType ="Table/Query"
                    RowSource ="tlu_Cover_Code"
                    StatusBarText ="Roo/Bole cover percentage quadrat 2"
                    OnChange ="=UpdateMicrohabitat([Screen].[ActiveControl])"

                    LayoutCachedLeft =2770
                    LayoutCachedTop =4560
                    LayoutCachedWidth =3548
                    LayoutCachedHeight =4800
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =95
                    IMESentenceMode =3
                    ListRows =21
                    Left =3550
                    Top =4560
                    Width =778
                    TabIndex =26
                    ColumnInfo ="\"\";\"\";\"6\";\"4\""
                    Name ="Root_Bole_Q3"
                    RowSourceType ="Table/Query"
                    RowSource ="tlu_Cover_Code"
                    StatusBarText ="Root/Bole cover percentage quadrat 3"
                    OnChange ="=UpdateMicrohabitat([Screen].[ActiveControl])"

                    LayoutCachedLeft =3550
                    LayoutCachedTop =4560
                    LayoutCachedWidth =4328
                    LayoutCachedHeight =4800
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =95
                    IMESentenceMode =3
                    ListRows =21
                    Left =1980
                    Top =5040
                    Width =778
                    Height =300
                    TabIndex =30
                    ColumnInfo ="\"\";\"\";\"6\";\"4\""
                    Name ="Mineral_Soil_Sediment_Q1"
                    RowSourceType ="Table/Query"
                    RowSource ="tlu_Cover_Code"
                    StatusBarText ="Mineral Soil/Sediment cover percentage quadrat 1"
                    OnChange ="=UpdateMicrohabitat([Screen].[ActiveControl])"

                    LayoutCachedLeft =1980
                    LayoutCachedTop =5040
                    LayoutCachedWidth =2758
                    LayoutCachedHeight =5340
                    Begin
                        Begin Label
                            OldBorderStyle =1
                            OverlapFlags =93
                            Left =120
                            Top =5040
                            Width =1858
                            Height =240
                            Name ="Label307"
                            Caption ="Mineral Soil/Sediment"
                            LayoutCachedLeft =120
                            LayoutCachedTop =5040
                            LayoutCachedWidth =1978
                            LayoutCachedHeight =5280
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =95
                    IMESentenceMode =3
                    ListRows =21
                    Left =2770
                    Top =5040
                    Width =778
                    Height =300
                    TabIndex =31
                    ColumnInfo ="\"\";\"\";\"6\";\"4\""
                    Name ="Mineral_Soil_Sediment_Q2"
                    RowSourceType ="Table/Query"
                    RowSource ="tlu_Cover_Code"
                    StatusBarText ="Mineral Soil/Sediment cover percentage quadrat 2"
                    OnChange ="=UpdateMicrohabitat([Screen].[ActiveControl])"

                    LayoutCachedLeft =2770
                    LayoutCachedTop =5040
                    LayoutCachedWidth =3548
                    LayoutCachedHeight =5340
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =95
                    IMESentenceMode =3
                    ListRows =21
                    Left =3550
                    Top =5040
                    Width =778
                    Height =300
                    TabIndex =32
                    ColumnInfo ="\"\";\"\";\"6\";\"4\""
                    Name ="Mineral_Soil_Sediment_Q3"
                    RowSourceType ="Table/Query"
                    RowSource ="tlu_Cover_Code"
                    StatusBarText ="Mineral Soil/Sediment cover percentage quadrat 3"
                    OnChange ="=UpdateMicrohabitat([Screen].[ActiveControl])"

                    LayoutCachedLeft =3550
                    LayoutCachedTop =5040
                    LayoutCachedWidth =4328
                    LayoutCachedHeight =5340
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =255
                    IMESentenceMode =3
                    ListRows =21
                    Left =1980
                    Top =5280
                    Width =778
                    TabIndex =33
                    ColumnInfo ="\"\";\"\";\"6\";\"4\""
                    Name ="Moss_Q1"
                    RowSourceType ="Table/Query"
                    RowSource ="tlu_Cover_Code"
                    StatusBarText ="Moss cover percentage quadrat 1"
                    OnChange ="=UpdateMicrohabitat([Screen].[ActiveControl])"

                    LayoutCachedLeft =1980
                    LayoutCachedTop =5280
                    LayoutCachedWidth =2758
                    LayoutCachedHeight =5520
                    Begin
                        Begin Label
                            OldBorderStyle =1
                            OverlapFlags =223
                            Left =120
                            Top =5280
                            Width =1860
                            Height =240
                            Name ="Label310"
                            Caption ="Moss"
                            LayoutCachedLeft =120
                            LayoutCachedTop =5280
                            LayoutCachedWidth =1980
                            LayoutCachedHeight =5520
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =255
                    IMESentenceMode =3
                    ListRows =21
                    Left =2770
                    Top =5280
                    Width =778
                    TabIndex =34
                    ColumnInfo ="\"\";\"\";\"6\";\"4\""
                    Name ="Moss_Q2"
                    RowSourceType ="Table/Query"
                    RowSource ="tlu_Cover_Code"
                    StatusBarText ="Moss cover percentage quadrat 2"
                    OnChange ="=UpdateMicrohabitat([Screen].[ActiveControl])"

                    LayoutCachedLeft =2770
                    LayoutCachedTop =5280
                    LayoutCachedWidth =3548
                    LayoutCachedHeight =5520
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =255
                    IMESentenceMode =3
                    ListRows =21
                    Left =3550
                    Top =5280
                    Width =778
                    TabIndex =35
                    ColumnInfo ="\"\";\"\";\"6\";\"4\""
                    Name ="Moss_Q3"
                    RowSourceType ="Table/Query"
                    RowSource ="tlu_Cover_Code"
                    StatusBarText ="Moss cover percentage quadrat 3"
                    OnChange ="=UpdateMicrohabitat([Screen].[ActiveControl])"

                    LayoutCachedLeft =3550
                    LayoutCachedTop =5280
                    LayoutCachedWidth =4328
                    LayoutCachedHeight =5520
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =127
                    IMESentenceMode =3
                    ListRows =21
                    Left =1980
                    Top =5520
                    Width =778
                    TabIndex =36
                    ColumnInfo ="\"\";\"\";\"6\";\"4\""
                    Name ="Cryptogram_Q1"
                    RowSourceType ="Table/Query"
                    RowSource ="tlu_Cover_Code"
                    StatusBarText ="Cryptogram cover percentage quadrat 1"
                    OnChange ="=UpdateMicrohabitat([Screen].[ActiveControl])"

                    LayoutCachedLeft =1980
                    LayoutCachedTop =5520
                    LayoutCachedWidth =2758
                    LayoutCachedHeight =5760
                    Begin
                        Begin Label
                            OldBorderStyle =1
                            OverlapFlags =127
                            Left =120
                            Top =5520
                            Width =1860
                            Height =240
                            Name ="Label313"
                            Caption ="Biological Soil Crust"
                            LayoutCachedLeft =120
                            LayoutCachedTop =5520
                            LayoutCachedWidth =1980
                            LayoutCachedHeight =5760
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =127
                    IMESentenceMode =3
                    ListRows =21
                    Left =2770
                    Top =5520
                    Width =778
                    TabIndex =37
                    ColumnInfo ="\"\";\"\";\"6\";\"4\""
                    Name ="Cryptogram_Q2"
                    RowSourceType ="Table/Query"
                    RowSource ="tlu_Cover_Code"
                    StatusBarText ="Cryptogram cover percentage quadrat 2"
                    OnChange ="=UpdateMicrohabitat([Screen].[ActiveControl])"

                    LayoutCachedLeft =2770
                    LayoutCachedTop =5520
                    LayoutCachedWidth =3548
                    LayoutCachedHeight =5760
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =127
                    IMESentenceMode =3
                    ListRows =21
                    Left =3550
                    Top =5520
                    Width =778
                    TabIndex =38
                    ColumnInfo ="\"\";\"\";\"6\";\"4\""
                    Name ="Cryptogram_Q3"
                    RowSourceType ="Table/Query"
                    RowSource ="tlu_Cover_Code"
                    StatusBarText ="Cryptogram cover percentage quadrat 3"
                    OnChange ="=UpdateMicrohabitat([Screen].[ActiveControl])"

                    LayoutCachedLeft =3550
                    LayoutCachedTop =5520
                    LayoutCachedWidth =4328
                    LayoutCachedHeight =5760
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =127
                    IMESentenceMode =3
                    ListRows =21
                    Left =1980
                    Top =5760
                    Width =778
                    TabIndex =39
                    ColumnInfo ="\"\";\"\";\"6\";\"4\""
                    Name ="Road_Q1"
                    RowSourceType ="Table/Query"
                    RowSource ="tlu_Cover_Code"
                    StatusBarText ="Road cover percentage quadrat 1"
                    OnChange ="=UpdateMicrohabitat([Screen].[ActiveControl])"

                    LayoutCachedLeft =1980
                    LayoutCachedTop =5760
                    LayoutCachedWidth =2758
                    LayoutCachedHeight =6000
                    Begin
                        Begin Label
                            OldBorderStyle =1
                            OverlapFlags =127
                            Left =120
                            Top =5760
                            Width =1860
                            Height =240
                            Name ="Label316"
                            Caption ="Road"
                            LayoutCachedLeft =120
                            LayoutCachedTop =5760
                            LayoutCachedWidth =1980
                            LayoutCachedHeight =6000
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =127
                    IMESentenceMode =3
                    ListRows =21
                    Left =2770
                    Top =5760
                    Width =778
                    TabIndex =40
                    ColumnInfo ="\"\";\"\";\"6\";\"4\""
                    Name ="Road_Q2"
                    RowSourceType ="Table/Query"
                    RowSource ="tlu_Cover_Code"
                    StatusBarText ="Road cover percentage quadrat 2"
                    OnChange ="=UpdateMicrohabitat([Screen].[ActiveControl])"

                    LayoutCachedLeft =2770
                    LayoutCachedTop =5760
                    LayoutCachedWidth =3548
                    LayoutCachedHeight =6000
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =127
                    IMESentenceMode =3
                    ListRows =21
                    Left =3550
                    Top =5760
                    Width =778
                    TabIndex =41
                    ColumnInfo ="\"\";\"\";\"6\";\"4\""
                    Name ="Road_Q3"
                    RowSourceType ="Table/Query"
                    RowSource ="tlu_Cover_Code"
                    StatusBarText ="Road cover percentage quadrat 3"
                    OnChange ="=UpdateMicrohabitat([Screen].[ActiveControl])"

                    LayoutCachedLeft =3550
                    LayoutCachedTop =5760
                    LayoutCachedWidth =4328
                    LayoutCachedHeight =6000
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =127
                    IMESentenceMode =3
                    ListRows =21
                    Left =1980
                    Top =6000
                    Width =778
                    TabIndex =42
                    ColumnInfo ="\"\";\"\";\"6\";\"4\""
                    Name ="Rock_Q1"
                    RowSourceType ="Table/Query"
                    RowSource ="tlu_Cover_Code"
                    StatusBarText ="Rock cover percentage quadrat 1"
                    OnChange ="=UpdateMicrohabitat([Screen].[ActiveControl])"

                    LayoutCachedLeft =1980
                    LayoutCachedTop =6000
                    LayoutCachedWidth =2758
                    LayoutCachedHeight =6240
                    Begin
                        Begin Label
                            OldBorderStyle =1
                            OverlapFlags =127
                            Left =120
                            Top =6000
                            Width =1860
                            Height =240
                            Name ="Label319"
                            Caption ="Rock"
                            LayoutCachedLeft =120
                            LayoutCachedTop =6000
                            LayoutCachedWidth =1980
                            LayoutCachedHeight =6240
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =127
                    IMESentenceMode =3
                    ListRows =21
                    Left =2770
                    Top =6000
                    Width =778
                    TabIndex =43
                    ColumnInfo ="\"\";\"\";\"6\";\"4\""
                    Name ="Rock_Q2"
                    RowSourceType ="Table/Query"
                    RowSource ="tlu_Cover_Code"
                    StatusBarText ="Rock cover percentage quadrat 2"
                    OnChange ="=UpdateMicrohabitat([Screen].[ActiveControl])"

                    LayoutCachedLeft =2770
                    LayoutCachedTop =6000
                    LayoutCachedWidth =3548
                    LayoutCachedHeight =6240
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =127
                    IMESentenceMode =3
                    ListRows =21
                    Left =3550
                    Top =6000
                    Width =778
                    TabIndex =44
                    ColumnInfo ="\"\";\"\";\"6\";\"4\""
                    Name ="Rock_Q3"
                    RowSourceType ="Table/Query"
                    RowSource ="tlu_Cover_Code"
                    StatusBarText ="Rock cover percentage quadrat 3"
                    OnChange ="=UpdateMicrohabitat([Screen].[ActiveControl])"

                    LayoutCachedLeft =3550
                    LayoutCachedTop =6000
                    LayoutCachedWidth =4328
                    LayoutCachedHeight =6240
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =127
                    IMESentenceMode =3
                    ListRows =21
                    Left =1980
                    Top =6240
                    Width =778
                    TabIndex =45
                    ColumnInfo ="\"\";\"\";\"6\";\"4\""
                    Name ="Standing_Water_Flooded_Q1"
                    RowSourceType ="Table/Query"
                    RowSource ="tlu_Cover_Code"
                    StatusBarText ="Standing Water/Flooded cover percentage quadrat 1"
                    OnChange ="=UpdateMicrohabitat([Screen].[ActiveControl])"

                    LayoutCachedLeft =1980
                    LayoutCachedTop =6240
                    LayoutCachedWidth =2758
                    LayoutCachedHeight =6480
                    Begin
                        Begin Label
                            OldBorderStyle =1
                            OverlapFlags =127
                            Left =120
                            Top =6240
                            Width =1860
                            Height =240
                            Name ="Label322"
                            Caption ="Standing Water/Flooded"
                            LayoutCachedLeft =120
                            LayoutCachedTop =6240
                            LayoutCachedWidth =1980
                            LayoutCachedHeight =6480
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =127
                    IMESentenceMode =3
                    ListRows =21
                    Left =2770
                    Top =6240
                    Width =778
                    TabIndex =46
                    ColumnInfo ="\"\";\"\";\"6\";\"4\""
                    Name ="Standing_Water_Flooded_Q2"
                    RowSourceType ="Table/Query"
                    RowSource ="tlu_Cover_Code"
                    StatusBarText ="Standing Water/Flooded cover percentage quadrat 2"
                    OnChange ="=UpdateMicrohabitat([Screen].[ActiveControl])"

                    LayoutCachedLeft =2770
                    LayoutCachedTop =6240
                    LayoutCachedWidth =3548
                    LayoutCachedHeight =6480
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =127
                    IMESentenceMode =3
                    ListRows =21
                    Left =3550
                    Top =6240
                    Width =778
                    TabIndex =47
                    ColumnInfo ="\"\";\"\";\"6\";\"4\""
                    Name ="Standing_Water_Flooded_Q3"
                    RowSourceType ="Table/Query"
                    RowSource ="tlu_Cover_Code"
                    StatusBarText ="Standing Water/Flooded cover percentage quadrat 3"
                    OnChange ="=UpdateMicrohabitat([Screen].[ActiveControl])"

                    LayoutCachedLeft =3550
                    LayoutCachedTop =6240
                    LayoutCachedWidth =4328
                    LayoutCachedHeight =6480
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =127
                    IMESentenceMode =3
                    ListRows =21
                    Left =1980
                    Top =6480
                    Width =778
                    TabIndex =48
                    ColumnInfo ="\"\";\"\";\"6\";\"4\""
                    Name ="Stream_Q1"
                    RowSourceType ="Table/Query"
                    RowSource ="tlu_Cover_Code"
                    StatusBarText ="Stream cover percentage quadrat 1"
                    OnChange ="=UpdateMicrohabitat([Screen].[ActiveControl])"

                    LayoutCachedLeft =1980
                    LayoutCachedTop =6480
                    LayoutCachedWidth =2758
                    LayoutCachedHeight =6720
                    Begin
                        Begin Label
                            OldBorderStyle =1
                            OverlapFlags =127
                            Left =120
                            Top =6480
                            Width =1860
                            Height =240
                            Name ="Label325"
                            Caption ="Stream"
                            LayoutCachedLeft =120
                            LayoutCachedTop =6480
                            LayoutCachedWidth =1980
                            LayoutCachedHeight =6720
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =127
                    IMESentenceMode =3
                    ListRows =21
                    Left =2770
                    Top =6480
                    Width =778
                    TabIndex =49
                    ColumnInfo ="\"\";\"\";\"6\";\"4\""
                    Name ="Stream_Q2"
                    RowSourceType ="Table/Query"
                    RowSource ="tlu_Cover_Code"
                    StatusBarText ="Stream cover percentage quadrat 2"
                    OnChange ="=UpdateMicrohabitat([Screen].[ActiveControl])"

                    LayoutCachedLeft =2770
                    LayoutCachedTop =6480
                    LayoutCachedWidth =3548
                    LayoutCachedHeight =6720
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =127
                    IMESentenceMode =3
                    ListRows =21
                    Left =3550
                    Top =6480
                    Width =778
                    TabIndex =50
                    ColumnInfo ="\"\";\"\";\"6\";\"4\""
                    Name ="Stream_Q3"
                    RowSourceType ="Table/Query"
                    RowSource ="tlu_Cover_Code"
                    StatusBarText ="Stream cover percentage quadrat 3"
                    OnChange ="=UpdateMicrohabitat([Screen].[ActiveControl])"

                    LayoutCachedLeft =3550
                    LayoutCachedTop =6480
                    LayoutCachedWidth =4328
                    LayoutCachedHeight =6720
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =127
                    IMESentenceMode =3
                    ListRows =21
                    Left =1980
                    Top =6720
                    Width =778
                    TabIndex =51
                    ColumnInfo ="\"\";\"\";\"6\";\"4\""
                    Name ="Trash_Junk_Q1"
                    RowSourceType ="Table/Query"
                    RowSource ="tlu_Cover_Code"
                    StatusBarText ="Trash/Junk cover percentage quadrat 1"
                    OnChange ="=UpdateMicrohabitat([Screen].[ActiveControl])"

                    LayoutCachedLeft =1980
                    LayoutCachedTop =6720
                    LayoutCachedWidth =2758
                    LayoutCachedHeight =6960
                    Begin
                        Begin Label
                            OldBorderStyle =1
                            OverlapFlags =127
                            Left =120
                            Top =6720
                            Width =1860
                            Height =240
                            Name ="Label328"
                            Caption ="Trash/Junk"
                            LayoutCachedLeft =120
                            LayoutCachedTop =6720
                            LayoutCachedWidth =1980
                            LayoutCachedHeight =6960
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =127
                    IMESentenceMode =3
                    ListRows =21
                    Left =2770
                    Top =6720
                    Width =778
                    TabIndex =52
                    ColumnInfo ="\"\";\"\";\"6\";\"4\""
                    Name ="Trash_Junk_Q2"
                    RowSourceType ="Table/Query"
                    RowSource ="tlu_Cover_Code"
                    StatusBarText ="Trash/Junk cover percentage quadrat 2"
                    OnChange ="=UpdateMicrohabitat([Screen].[ActiveControl])"

                    LayoutCachedLeft =2770
                    LayoutCachedTop =6720
                    LayoutCachedWidth =3548
                    LayoutCachedHeight =6960
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =119
                    IMESentenceMode =3
                    ListRows =21
                    Left =3550
                    Top =6720
                    Width =778
                    TabIndex =53
                    ColumnInfo ="\"\";\"\";\"6\";\"4\""
                    Name ="Trash_Junk_Q3"
                    RowSourceType ="Table/Query"
                    RowSource ="tlu_Cover_Code"
                    StatusBarText ="Trash/Junk cover percentage quadrat 3"
                    OnChange ="=UpdateMicrohabitat([Screen].[ActiveControl])"

                    LayoutCachedLeft =3550
                    LayoutCachedTop =6720
                    LayoutCachedWidth =4328
                    LayoutCachedHeight =6960
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    SpecialEffect =0
                    OverlapFlags =215
                    IMESentenceMode =3
                    ListRows =21
                    Left =1980
                    Top =4800
                    Width =778
                    TabIndex =27
                    ColumnInfo ="\"\";\"\";\"6\";\"4\""
                    Name ="Dead_Root_Bole_Q1"
                    RowSourceType ="Table/Query"
                    RowSource ="tlu_Cover_Code"
                    StatusBarText ="Root/Bole cover percentage quadrat 1"
                    OnChange ="=UpdateMicrohabitat([Screen].[ActiveControl])"

                    LayoutCachedLeft =1980
                    LayoutCachedTop =4800
                    LayoutCachedWidth =2758
                    LayoutCachedHeight =5040
                    Begin
                        Begin Label
                            OldBorderStyle =1
                            OverlapFlags =223
                            Left =120
                            Top =4800
                            Width =1860
                            Height =240
                            Name ="Label40"
                            Caption ="Dead root/Bole"
                            LayoutCachedLeft =120
                            LayoutCachedTop =4800
                            LayoutCachedWidth =1980
                            LayoutCachedHeight =5040
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    SpecialEffect =0
                    OverlapFlags =223
                    IMESentenceMode =3
                    ListRows =21
                    Left =2770
                    Top =4800
                    Width =778
                    TabIndex =28
                    ColumnInfo ="\"\";\"\";\"6\";\"4\""
                    Name ="Dead_Root_Bole_Q2"
                    RowSourceType ="Table/Query"
                    RowSource ="tlu_Cover_Code"
                    StatusBarText ="Roo/Bole cover percentage quadrat 2"
                    OnChange ="=UpdateMicrohabitat([Screen].[ActiveControl])"

                    LayoutCachedLeft =2770
                    LayoutCachedTop =4800
                    LayoutCachedWidth =3548
                    LayoutCachedHeight =5040
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    SpecialEffect =0
                    OverlapFlags =215
                    IMESentenceMode =3
                    ListRows =21
                    Left =3550
                    Top =4800
                    Width =763
                    TabIndex =29
                    ColumnInfo ="\"\";\"\";\"6\";\"4\""
                    Name ="Dead_Root_Bole_Q3"
                    RowSourceType ="Table/Query"
                    RowSource ="tlu_Cover_Code"
                    StatusBarText ="Root/Bole cover percentage quadrat 3"
                    OnChange ="=UpdateMicrohabitat([Screen].[ActiveControl])"

                    LayoutCachedLeft =3550
                    LayoutCachedTop =4800
                    LayoutCachedWidth =4313
                    LayoutCachedHeight =5040
                End
                Begin Label
                    BackStyle =1
                    OldBorderStyle =1
                    OverlapFlags =93
                    Left =45
                    Top =1065
                    Width =4500
                    Height =1440
                    FontSize =14
                    LeftMargin =216
                    TopMargin =288
                    BackColor =14610923
                    BorderColor =12349952
                    ForeColor =0
                    Name ="lblChkboxes"
                    FontName ="Calibri"
                    GridlineColor =10921638
                    LayoutCachedLeft =45
                    LayoutCachedTop =1065
                    LayoutCachedWidth =4545
                    LayoutCachedHeight =2505
                    ThemeFontIndex =1
                    BackThemeColorIndex =6
                    BackTint =20.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin ToggleButton
                    Enabled = NotDefault
                    OverlapFlags =215
                    Left =4155
                    Top =2025
                    Width =270
                    Height =300
                    FontSize =10
                    FontWeight =700
                    TabIndex =57
                    Name ="tglNoExoticsQ3"
                    AfterUpdate ="[Event Procedure]"
                    FontName ="Calibri"
                    ControlTipText ="Q3 has no priority 1 exotics"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120
                    GridlineColor =10921638

                    LayoutCachedLeft =4155
                    LayoutCachedTop =2025
                    LayoutCachedWidth =4425
                    LayoutCachedHeight =2325
                    ForeThemeColorIndex =0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    Gradient =12
                    BackColor =12419407
                    BackThemeColorIndex =4
                    BorderColor =12419407
                    BorderThemeColorIndex =4
                    ThemeFontIndex =1
                    HoverColor =65280
                    HoverTint =80.0
                    PressedColor =10250042
                    PressedThemeColorIndex =4
                    PressedShade =80.0
                    HoverForeColor =0
                    HoverForeThemeColorIndex =0
                    PressedForeColor =0
                    PressedForeThemeColorIndex =0
                    QuickStyle =23
                    QuickStyleMask =-5
                    WebImagePaddingLeft =4
                    WebImagePaddingTop =2
                    WebImagePaddingRight =4
                    WebImagePaddingBottom =7
                    Overlaps =1
                    Begin
                        Begin Label
                            OverlapFlags =215
                            Left =4125
                            Top =1185
                            Width =330
                            Height =315
                            FontSize =11
                            FontWeight =600
                            BackColor =16777215
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="lblQ3"
                            Caption ="Q3"
                            FontName ="Calibri"
                            ControlTipText ="Q3 flags"
                            GridlineColor =10921638
                            LayoutCachedLeft =4125
                            LayoutCachedTop =1185
                            LayoutCachedWidth =4455
                            LayoutCachedHeight =1500
                            ThemeFontIndex =1
                            BackThemeColorIndex =1
                            BorderThemeColorIndex =0
                            BorderTint =50.0
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                        End
                    End
                End
                Begin ToggleButton
                    Enabled = NotDefault
                    OverlapFlags =215
                    Left =3720
                    Top =2025
                    Width =270
                    Height =299
                    FontSize =10
                    FontWeight =700
                    TabIndex =58
                    Name ="tglNoExoticsQ2"
                    AfterUpdate ="[Event Procedure]"
                    FontName ="Calibri"
                    ControlTipText ="Q2 has no priority 1 exotics"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120
                    GridlineColor =10921638

                    LayoutCachedLeft =3720
                    LayoutCachedTop =2025
                    LayoutCachedWidth =3990
                    LayoutCachedHeight =2324
                    ForeThemeColorIndex =0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    Gradient =12
                    BackColor =12419407
                    BackThemeColorIndex =4
                    BorderColor =12419407
                    BorderThemeColorIndex =4
                    ThemeFontIndex =1
                    HoverColor =65280
                    HoverTint =80.0
                    PressedColor =10250042
                    PressedThemeColorIndex =4
                    PressedShade =80.0
                    HoverForeColor =0
                    HoverForeThemeColorIndex =0
                    PressedForeColor =0
                    PressedForeThemeColorIndex =0
                    QuickStyle =23
                    QuickStyleMask =-5
                    WebImagePaddingLeft =4
                    WebImagePaddingTop =2
                    WebImagePaddingRight =4
                    WebImagePaddingBottom =7
                    Overlaps =1
                    Begin
                        Begin Label
                            OverlapFlags =215
                            Left =3690
                            Top =1185
                            Width =330
                            Height =315
                            FontSize =11
                            FontWeight =600
                            BackColor =16777215
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="lblQ2"
                            Caption ="Q2"
                            FontName ="Calibri"
                            ControlTipText ="Q2 flags"
                            GridlineColor =10921638
                            LayoutCachedLeft =3690
                            LayoutCachedTop =1185
                            LayoutCachedWidth =4020
                            LayoutCachedHeight =1500
                            ThemeFontIndex =1
                            BackThemeColorIndex =1
                            BorderThemeColorIndex =0
                            BorderTint =50.0
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                        End
                    End
                End
                Begin ToggleButton
                    Enabled = NotDefault
                    OverlapFlags =215
                    Left =3285
                    Top =2025
                    Width =270
                    Height =299
                    FontSize =10
                    FontWeight =700
                    TabIndex =59
                    Name ="tglNoExoticsQ1"
                    AfterUpdate ="[Event Procedure]"
                    FontName ="Calibri"
                    ControlTipText ="Q1 has no priority 1 exotics"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120
                    GridlineColor =10921638

                    LayoutCachedLeft =3285
                    LayoutCachedTop =2025
                    LayoutCachedWidth =3555
                    LayoutCachedHeight =2324
                    ForeThemeColorIndex =0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    Gradient =12
                    BackColor =12419407
                    BackThemeColorIndex =4
                    BorderColor =12419407
                    BorderThemeColorIndex =4
                    ThemeFontIndex =1
                    HoverColor =65280
                    HoverTint =80.0
                    PressedColor =10250042
                    PressedThemeColorIndex =4
                    PressedShade =80.0
                    HoverForeColor =0
                    HoverForeThemeColorIndex =0
                    PressedForeColor =0
                    PressedForeThemeColorIndex =0
                    QuickStyle =23
                    QuickStyleMask =-5
                    WebImagePaddingLeft =4
                    WebImagePaddingTop =2
                    WebImagePaddingRight =4
                    WebImagePaddingBottom =7
                    Overlaps =1
                    Begin
                        Begin Label
                            OverlapFlags =215
                            Left =3270
                            Top =1185
                            Width =330
                            Height =315
                            FontSize =11
                            FontWeight =600
                            BackColor =16777215
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="lblQ1"
                            Caption ="Q1"
                            FontName ="Calibri"
                            ControlTipText ="Q1 flags"
                            GridlineColor =10921638
                            LayoutCachedLeft =3270
                            LayoutCachedTop =1185
                            LayoutCachedWidth =3600
                            LayoutCachedHeight =1500
                            ThemeFontIndex =1
                            BackThemeColorIndex =1
                            BorderThemeColorIndex =0
                            BorderTint =50.0
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                        End
                    End
                End
                Begin ToggleButton
                    Enabled = NotDefault
                    OverlapFlags =215
                    Left =2520
                    Top =2040
                    Width =270
                    Height =269
                    FontSize =10
                    FontWeight =700
                    TabIndex =60
                    Name ="tglNoExoticsT"
                    AfterUpdate ="[Event Procedure]"
                    FontName ="Calibri"
                    ControlTipText ="Transect has no exotics"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120
                    GridlineColor =10921638

                    LayoutCachedLeft =2520
                    LayoutCachedTop =2040
                    LayoutCachedWidth =2790
                    LayoutCachedHeight =2309
                    ForeThemeColorIndex =0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    Gradient =12
                    BackColor =12419407
                    BackThemeColorIndex =4
                    BorderColor =12419407
                    BorderThemeColorIndex =4
                    ThemeFontIndex =1
                    HoverColor =65280
                    HoverTint =80.0
                    PressedColor =10250042
                    PressedThemeColorIndex =4
                    PressedShade =80.0
                    HoverForeColor =0
                    HoverForeThemeColorIndex =0
                    PressedForeColor =0
                    PressedForeThemeColorIndex =0
                    QuickStyle =23
                    QuickStyleMask =-5
                    WebImagePaddingLeft =4
                    WebImagePaddingTop =2
                    WebImagePaddingRight =4
                    WebImagePaddingBottom =7
                    Overlaps =1
                    Begin
                        Begin Label
                            OverlapFlags =215
                            Left =2130
                            Top =1185
                            Width =990
                            Height =315
                            FontSize =11
                            FontWeight =600
                            BackColor =16777215
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="lblTransect"
                            Caption ="Transect"
                            FontName ="Calibri"
                            ControlTipText ="Transect flags"
                            GridlineColor =10921638
                            LayoutCachedLeft =2130
                            LayoutCachedTop =1185
                            LayoutCachedWidth =3120
                            LayoutCachedHeight =1500
                            ThemeFontIndex =1
                            BackThemeColorIndex =1
                            BorderThemeColorIndex =0
                            BorderTint =50.0
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                        End
                    End
                End
                Begin Label
                    OverlapFlags =215
                    Left =165
                    Top =2025
                    Width =2025
                    Height =285
                    FontSize =11
                    BackColor =16777215
                    BorderColor =8355711
                    ForeColor =2500134
                    Name ="lblNoExotics"
                    Caption ="No Priority 1 Exotics?"
                    FontName ="Calibri"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Priority 1 exotics"
                    GridlineColor =10921638
                    LayoutCachedLeft =165
                    LayoutCachedTop =2025
                    LayoutCachedWidth =2190
                    LayoutCachedHeight =2310
                    ThemeFontIndex =1
                    BackThemeColorIndex =1
                    BorderThemeColorIndex =0
                    BorderTint =50.0
                    ForeThemeColorIndex =0
                    ForeTint =85.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin ToggleButton
                    Enabled = NotDefault
                    OverlapFlags =215
                    Left =4155
                    Top =1605
                    Width =270
                    Height =300
                    FontSize =10
                    FontWeight =700
                    TabIndex =61
                    Name ="tglNotSampledQ3"
                    AfterUpdate ="[Event Procedure]"
                    FontName ="Calibri"
                    ControlTipText ="Q3 not sampled"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120
                    GridlineColor =10921638

                    LayoutCachedLeft =4155
                    LayoutCachedTop =1605
                    LayoutCachedWidth =4425
                    LayoutCachedHeight =1905
                    ForeThemeColorIndex =0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    Gradient =12
                    BackColor =12419407
                    BackThemeColorIndex =4
                    BorderColor =12419407
                    BorderThemeColorIndex =4
                    ThemeFontIndex =1
                    HoverColor =65280
                    HoverTint =80.0
                    PressedColor =10250042
                    PressedThemeColorIndex =4
                    PressedShade =80.0
                    HoverForeColor =0
                    HoverForeThemeColorIndex =0
                    PressedForeColor =0
                    PressedForeThemeColorIndex =0
                    QuickStyle =23
                    QuickStyleMask =-5
                    WebImagePaddingLeft =4
                    WebImagePaddingTop =2
                    WebImagePaddingRight =4
                    WebImagePaddingBottom =7
                    Overlaps =1
                End
                Begin ToggleButton
                    Enabled = NotDefault
                    OverlapFlags =215
                    Left =3720
                    Top =1605
                    Width =270
                    Height =299
                    FontSize =10
                    FontWeight =700
                    TabIndex =62
                    Name ="tglNotSampledQ2"
                    AfterUpdate ="[Event Procedure]"
                    FontName ="Calibri"
                    ControlTipText ="Q2 was not sampled"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120
                    GridlineColor =10921638

                    LayoutCachedLeft =3720
                    LayoutCachedTop =1605
                    LayoutCachedWidth =3990
                    LayoutCachedHeight =1904
                    ForeThemeColorIndex =0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    Gradient =12
                    BackColor =12419407
                    BackThemeColorIndex =4
                    BorderColor =12419407
                    BorderThemeColorIndex =4
                    ThemeFontIndex =1
                    HoverColor =65280
                    HoverTint =80.0
                    PressedColor =10250042
                    PressedThemeColorIndex =4
                    PressedShade =80.0
                    HoverForeColor =0
                    HoverForeThemeColorIndex =0
                    PressedForeColor =0
                    PressedForeThemeColorIndex =0
                    QuickStyle =23
                    QuickStyleMask =-5
                    WebImagePaddingLeft =4
                    WebImagePaddingTop =2
                    WebImagePaddingRight =4
                    WebImagePaddingBottom =7
                    Overlaps =1
                End
                Begin ToggleButton
                    Enabled = NotDefault
                    OverlapFlags =215
                    Left =3285
                    Top =1605
                    Width =270
                    Height =299
                    FontSize =10
                    FontWeight =700
                    TabIndex =63
                    Name ="tglNotSampledQ1"
                    AfterUpdate ="[Event Procedure]"
                    FontName ="Calibri"
                    ControlTipText ="Q1 was not sampled"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120
                    GridlineColor =10921638

                    LayoutCachedLeft =3285
                    LayoutCachedTop =1605
                    LayoutCachedWidth =3555
                    LayoutCachedHeight =1904
                    ForeThemeColorIndex =0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    Gradient =12
                    BackColor =12419407
                    BackThemeColorIndex =4
                    BorderColor =12419407
                    BorderThemeColorIndex =4
                    ThemeFontIndex =1
                    HoverColor =65280
                    HoverTint =80.0
                    PressedColor =10250042
                    PressedThemeColorIndex =4
                    PressedShade =80.0
                    HoverForeColor =0
                    HoverForeThemeColorIndex =0
                    PressedForeColor =0
                    PressedForeThemeColorIndex =0
                    QuickStyle =23
                    QuickStyleMask =-5
                    WebImagePaddingLeft =4
                    WebImagePaddingTop =2
                    WebImagePaddingRight =4
                    WebImagePaddingBottom =7
                    Overlaps =1
                End
                Begin ToggleButton
                    Enabled = NotDefault
                    OverlapFlags =215
                    Left =2520
                    Top =1620
                    Width =270
                    Height =269
                    FontSize =10
                    FontWeight =700
                    TabIndex =64
                    Name ="tglNotSampledT"
                    AfterUpdate ="[Event Procedure]"
                    FontName ="Calibri"
                    ControlTipText ="Transect was not sampled"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120
                    GridlineColor =10921638

                    LayoutCachedLeft =2520
                    LayoutCachedTop =1620
                    LayoutCachedWidth =2790
                    LayoutCachedHeight =1889
                    ForeThemeColorIndex =0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    Gradient =12
                    BackColor =12419407
                    BackThemeColorIndex =4
                    BorderColor =12419407
                    BorderThemeColorIndex =4
                    ThemeFontIndex =1
                    HoverColor =65280
                    HoverTint =80.0
                    PressedColor =10250042
                    PressedThemeColorIndex =4
                    PressedShade =80.0
                    HoverForeColor =0
                    HoverForeThemeColorIndex =0
                    PressedForeColor =0
                    PressedForeThemeColorIndex =0
                    QuickStyle =23
                    QuickStyleMask =-5
                    WebImagePaddingLeft =4
                    WebImagePaddingTop =2
                    WebImagePaddingRight =4
                    WebImagePaddingBottom =7
                    Overlaps =1
                End
                Begin Label
                    OverlapFlags =215
                    Left =165
                    Top =1605
                    Width =2025
                    Height =285
                    FontSize =11
                    BackColor =16777215
                    BorderColor =8355711
                    ForeColor =2500134
                    Name ="lblNotSampled"
                    Caption ="Not Sampled?"
                    FontName ="Calibri"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Not sampled?"
                    GridlineColor =10921638
                    LayoutCachedLeft =165
                    LayoutCachedTop =1605
                    LayoutCachedWidth =2190
                    LayoutCachedHeight =1890
                    ThemeFontIndex =1
                    BackThemeColorIndex =1
                    BorderThemeColorIndex =0
                    BorderTint =50.0
                    ForeThemeColorIndex =0
                    ForeTint =85.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin Subform
                    OverlapFlags =255
                    Left =4620
                    Top =1020
                    Width =8580
                    Height =6718
                    TabIndex =54
                    Name ="fsub_Species_Current"
                    SourceObject ="Form.fsub_Species"
                    LinkChildFields ="Transect_ID"
                    LinkMasterFields ="Transect_ID"

                    LayoutCachedLeft =4620
                    LayoutCachedTop =1020
                    LayoutCachedWidth =13200
                    LayoutCachedHeight =7738
                End
                Begin Subform
                    Visible = NotDefault
                    OverlapFlags =247
                    SpecialEffect =0
                    Left =7080
                    Top =2592
                    Width =3300
                    Height =599
                    TabIndex =65
                    BorderColor =2366701
                    Name ="fsub_Message"
                    SourceObject ="Form.fsub_Msg"

                    LayoutCachedLeft =7080
                    LayoutCachedTop =2592
                    LayoutCachedWidth =10380
                    LayoutCachedHeight =3191
                End
                Begin TextBox
                    SpecialEffect =0
                    OldBorderStyle =0
                    OverlapFlags =85
                    BackStyle =0
                    IMESentenceMode =3
                    Left =60
                    Top =420
                    Width =3180
                    Height =255
                    TabIndex =66
                    ForeColor =8355711
                    Name ="tbxTransectID"
                    ControlSource ="Transect_ID"

                    LayoutCachedLeft =60
                    LayoutCachedTop =420
                    LayoutCachedWidth =3240
                    LayoutCachedHeight =675
                    ForeThemeColorIndex =1
                    ForeShade =50.0
                End
                Begin TextBox
                    SpecialEffect =0
                    OldBorderStyle =0
                    OverlapFlags =85
                    BackStyle =0
                    IMESentenceMode =3
                    Left =240
                    Top =7320
                    Width =1200
                    Height =255
                    TabIndex =67
                    ForeColor =8355711
                    Name ="tbxQ1"

                    LayoutCachedLeft =240
                    LayoutCachedTop =7320
                    LayoutCachedWidth =1440
                    LayoutCachedHeight =7575
                    ForeThemeColorIndex =1
                    ForeShade =50.0
                End
                Begin TextBox
                    SpecialEffect =0
                    OldBorderStyle =0
                    OverlapFlags =85
                    BackStyle =0
                    IMESentenceMode =3
                    Left =1680
                    Top =7320
                    Width =1200
                    Height =255
                    TabIndex =68
                    ForeColor =8355711
                    Name ="tbxQ2"

                    LayoutCachedLeft =1680
                    LayoutCachedTop =7320
                    LayoutCachedWidth =2880
                    LayoutCachedHeight =7575
                    ForeThemeColorIndex =1
                    ForeShade =50.0
                End
                Begin TextBox
                    SpecialEffect =0
                    OldBorderStyle =0
                    OverlapFlags =85
                    BackStyle =0
                    IMESentenceMode =3
                    Left =3120
                    Top =7320
                    Width =1200
                    Height =255
                    TabIndex =69
                    ForeColor =8355711
                    Name ="tbxQ3"

                    LayoutCachedLeft =3120
                    LayoutCachedTop =7320
                    LayoutCachedWidth =4320
                    LayoutCachedHeight =7575
                    ForeThemeColorIndex =1
                    ForeShade =50.0
                End
            End
        End
        Begin FormFooter
            Height =0
            BackColor =-2147483633
            Name ="FormFooter"
        End
    End
End
CodeBehindForm
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

' =================================
' Form:         frm_Quadrat_Transect
' Level:        Application form
' Version:      1.03
' Basis:        -
'
' Description:  Quadrat Transect form object related properties, functions & procedures for UI display
'
' Source/date:  NCPN, Unknown - for NCPN tools
' References:   -
' Revisions:    NCPN - Unknown - 1.00 - initial version
'               BLC - 3/8/2017 - 1.01 - added documentation, error handling,
'                                       added tglNotSampled (T, Q1-3) AfterUpdate events,
'                                       CheckTransectLevel(), SetQuadratToggles()
'                                       revised subform control to NOT match form name it contains
'                                       therefore fsub_Current is the container for fsub_Species subform
'                                       (handles 2010 & later species)
'               BLC - 4/21/2017 - 1.02 - added check for if species subform has records
'               BLC - 4/23/2017 - 1.03 - revised Next/Previous to cycle through transects vs. presenting
'                                        error message (Error 2105 - can't go to specified record),
'                                        pull microhabitats, species cover from respective SurfaceCover,
'                                        SpeciesCover tables
' =================================

'---------------------
' Simulated Inheritance
'---------------------

'---------------------
' Declarations
'---------------------
Dim strCheck As String

'---------------------
' Event Declarations
'---------------------

'---------------------
' Properties
'---------------------

'---------------------
' Methods
'---------------------

' ---------------------------------
' Sub:          Form_Open
' Description:  form opening actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:
'   Harvey French, July 31, 2015
'   http://stackoverflow.com/questions/31611912/how-best-to-call-a-public-sub-routine-declared-in-a-form-used-as-the-source-obje
' Source/date:  NCPN, Unknown - for NCPN tools
' Adapted:      -
' Revisions:
'   NCPN - Unknown - initial version
'   BLC - 3/8/2017 - added documentation, error handling
'   BLC - 4/23/2017 - added initialization for transect # since Next/Previous
'                     buttons now cycle through the transects vs. stopping w/ 2105 error message,
'                     added call to PopulateMicrohabitats to pull them from SurfaceCover
' ---------------------------------
Private Sub Form_Open(Cancel As Integer)
On Error GoTo Err_Handler
    
    'default
    strCheck = StringFromCodepoint(uCheck)
    
    'set hover
    tglNotSampledT.HoverColor = lngGreen
    tglNotSampledQ1.HoverColor = lngGreen
    tglNotSampledQ2.HoverColor = lngGreen
    tglNotSampledQ3.HoverColor = lngGreen
    tglNoExoticsT.HoverColor = lngGreen
    tglNoExoticsQ1.HoverColor = lngGreen
    tglNoExoticsQ2.HoverColor = lngGreen
    tglNoExoticsQ3.HoverColor = lngGreen
    
    'defaults
    Me.fsub_Message.Visible = False
    
    'initialize toggles (all toggles begin w/ tgl)
    Dim ctl As Control
    
    For Each ctl In Me.Controls
        If Left(ctl.Name, 3) = "tgl" Then
            ctl.Enabled = True
            ctl.ForeColor = lngBlack
        End If
    Next
  
    'initialize subform properties
'    If Me.fsub_Species_Current.Form.ParentForm Is Nothing Then _
'        Me.fsub_Species_Current.Form.ParentForm = Me

    'initialize Quadrat # temp vars
    Me.tbxQ1 = 0
    Me.tbxQ2 = 0
    Me.tbxQ3 = 0
  
    'set starting transect # to red
    '(1st & last transects are red for bounding since Next/Previous cycle)
    Me.Transect.ForeColor = lngRed
    
    'populate the microhabitats from SurfaceCover
    PopulateMicrohabitats
  
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Form_Open[frm_Quadrat_Transect form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          Form_Current
' Description:  form current actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:
'   Patrick Lepelletier, April 19, 2016
'   http://stackoverflow.com/questions/7000334/deleting-elements-in-an-array-if-element-is-a-certain-value-vba
' Source/date:  NCPN, Unknown - for NCPN tools
' Adapted:      -
' Revisions:
'   NCPN - Unknown - initial version
'   BLC - 3/8/2017 - added documentation, error handling
'                    revised subform control to NOT match form name it contains
'                    therefore fsub_Current is the container for fsub_Species subform
'                    (handles 2010 & later species)
'   BLC - 4/21/2017 - added check for if species subform has records
' ---------------------------------
Private Sub Form_Current()
On Error GoTo Err_Handler
              
  If Me.Parent!Start_Date < #1/1/2009# Then
    Me!fsub_Species_Current.Visible = False
    Me!fsub_Species_2008.Visible = True
    Me!fsub_Species_2009.Visible = False
  ElseIf Me.Parent!Start_Date < #1/1/2010# Then
    Me!fsub_Species_Current.Visible = False
    Me!fsub_Species_2009.Visible = True
    Me!fsub_Species_2008.Visible = False
  Else
    Me!fsub_Species_Current.Visible = True
    Me!fsub_Species_2008.Visible = False
    Me!fsub_Species_2009.Visible = False
  End If

    'update AvgCover
    'Me.fsub_Species_Current!Average_Cover = Me.fsub_Species_Current.Form.CalcAvgCover
    
    'update Quadrat IDs
'    Me.tbxQ1 = Nz(TempVars("Q1_ID"), 0)
'    Me.tbxQ2 = Nz(TempVars("Q2_ID"), 0)
'    Me.tbxQ3 = Nz(TempVars("Q3_ID"), 0)
    
    'set up toggles depending on species data
    With Me.fsub_Species_Current.Form
    
        Debug.Print .HasRecords
        'if species subform has records --> disable transect & quadrat toggles (IsSampled, NoExotics)
        If .HasRecords = True Then
            
            'check if Q1,Q2,Q3 % Cover values are set
           Debug.Print .Plant_Code
           Debug.Print .HasRecordsQ1 & "-"; .HasRecordsQ2 & "-"; .HasRecordsQ3
            'disable transect & quadrat toggles
            DisableToggles
            
            'enable select toggles depending on which quadrats have records
            Dim colToggles As New Collection
            Dim i As Integer
            Dim tgl As Variant
            
            For i = 1 To 3
                colToggles.Add i
            Next
            
            If .HasRecordsQ1 Then colToggles.Remove 1 'remove 1
            If .HasRecordsQ2 Then colToggles.Remove 2 'remove 2
            If .HasRecordsQ3 Then colToggles.Remove 3 'remove 3
            
            For Each tgl In colToggles
                EnableToggles CInt(tgl)
            Next
        Else
            
            'enable transect & quadrat toggles
            EnableToggles
        
        End If
    End With

Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Form_Current[frm_Quadrat_Transect form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          btnPrevious_KeyDown
' Description:  Previous button key down actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  NCPN, Unknown - for NCPN tools
' Adapted:      -
' Revisions:
'   NCPN - Unknown - initial version
'   BLC - 3/8/2017 - added documentation, error handling
' ---------------------------------
Private Sub btnPrevious_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo Err_Handler
  
  ' Ignore Page Down and Page Up keys for they will cycle through records
  Select Case KeyCode
    Case 33, 34
      KeyCode = 0
    End Select

Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnPrevious_KeyDown[frm_Quadrat_Transect form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          btnNext_KeyDown
' Description:  Next button key down actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  NCPN, Unknown - for NCPN tools
' Adapted:      -
' Revisions:
'   NCPN - Unknown - initial version
'   BLC - 3/8/2017 - added documentation, error handling
' ---------------------------------
Private Sub btnNext_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo Err_Handler

  ' Ignore Page Down and Page Up keys for they will cycle through records
  Select Case KeyCode
    Case 33, 34
      KeyCode = 0
    End Select
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnNext_KeyDown[frm_Quadrat_Transect form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          tbxVisitDate_KeyDown
' Description:  VisitDate textbox key down actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  NCPN, Unknown - for NCPN tools
' Adapted:      -
' Revisions:
'   NCPN - Unknown - initial version
'   BLC - 3/8/2017 - added documentation, error handling
' ---------------------------------
Private Sub tbxVisitDate_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo Err_Handler

  ' Ignore Page Down and Page Up keys for they will cycle through records
  Select Case KeyCode
    Case 33, 34
      KeyCode = 0
    End Select

Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - tbxVisitDate_KeyDown[frm_Quadrat_Transect form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          btnPrevious_Click
' Description:  Previous button click actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:
'   https://access-programmers.co.uk/forums/showthread.php?t=104478
'   Strike_Eagle, March 21, 2014
' Source/date:  NCPN, Unknown - for NCPN tools
' Adapted:      -
' Revisions:
'   NCPN - Unknown - initial version
'   BLC - 3/8/2017 - added documentation, error handling
'   BLC - 4/23/2017 - revised Next/Previous to cycle through transects vs. presenting
'                     error message (Error 2105: can't go to specified record),
'                     added calls to populate SurfaceCover & SpeciesCover
'                     from their respective tables
' ---------------------------------
Private Sub btnPrevious_Click()
On Error GoTo Err_Handler

    'determine # of records
    Dim rs As DAO.Recordset
    Dim numRecords As Integer
    
    Set rs = Me.RecordsetClone
    If Not rs.EOF And rs.BOF Then
        rs.MoveLast
    End If
    
    numRecords = rs.RecordCount
    
    'use recordset absolute position to determine
    'if on first/last record or in between
    With Me.Recordset
    
        'test for zero point (before 1st record)
        If .AbsolutePosition = 0 Then
            'go to last record if on first
            DoCmd.GoToRecord , , acLast
        Else
            'go to previous record if not on first
            DoCmd.GoToRecord , , acPrevious
        End If
        
        'identify the record as 1st or last
        'AbsolutePosition zero based, so + 1
        If .AbsolutePosition + 1 = numRecords Then
            Transect.ForeColor = lngRed
            Transect.ControlTipText = "Last Transect"
        ElseIf .AbsolutePosition = 0 Then
            Transect.ForeColor = lngRed
            Transect.ControlTipText = "First Transect"
        Else
            Transect.ForeColor = lngBlack
            Transect.ControlTipText = ""
        End If
    
    End With
    
    'populate w/ current transect's data
    PopulateMicrohabitats
      
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnPrevious_Click[frm_Quadrat_Transect form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          btnNext_Click
' Description:  Next button click actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  NCPN, Unknown - for NCPN tools
' Adapted:      -
' Revisions:
'   NCPN - Unknown - initial version
'   BLC - 3/8/2017 - added documentation, error handling
'   BLC - 4/23/2017 - revised Next/Previous to cycle through transects vs. presenting
'                     error message (Error 2105: can't go to specified record),
'                     added calls to populate SurfaceCover & SpeciesCover
'                     from their respective tables
' ---------------------------------
Private Sub btnNext_Click()
On Error GoTo Err_Handler

    'determine # of records
    Dim rs As DAO.Recordset
    Dim numRecords As Integer
    
    Set rs = Me.RecordsetClone
    If Not rs.EOF And rs.BOF Then
        rs.MoveLast
    End If
    
    numRecords = rs.RecordCount
    
    'use recordset absolute position to determine
    'if on first/last record or in between
    With Me.Recordset
    
        'test for zero point (before 1st record)
        If .AbsolutePosition + 1 = numRecords Then
            'go to first record if on last
            DoCmd.GoToRecord , , acFirst
        Else
            'go to next record if not on last
            DoCmd.GoToRecord , , acNext
        End If
        
        'identify the record as 1st or last
        'AbsolutePosition is zero based so +1
        If .AbsolutePosition + 1 = numRecords Then
            Transect.ForeColor = lngRed
            Transect.ControlTipText = "Last Transect"
        ElseIf .AbsolutePosition = 0 Then
            Transect.ForeColor = lngRed
            Transect.ControlTipText = "First Transect"
        Else
            Transect.ForeColor = lngBlack
            Transect.ControlTipText = ""
        End If
    
    End With

    'populate w/ current transect's data
    PopulateMicrohabitats
    
    'repaint?
    Me.Repaint

Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnNext_Click[frm_Quadrat_Transect form])"
    End Select
    Resume Exit_Handler
End Sub

' =================================
'   Not Sampled Flag
' =================================
' ---------------------------------
' Sub:          tglNotSampledT_AfterUpdate
' Description:  Toggle after update actions
' Assumptions:  Transect not sampled? -> no priority 1 species either
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, March 8, 2017 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 3/8/2017 - initial version
' ---------------------------------
Private Sub tglNotSampledT_AfterUpdate()
On Error GoTo Err_Handler
    
    Dim i As Integer
    Dim strControl As String

 '   strCheck = StringFromCodepoint(uCheck)

    'display as checkbox
    ToggleCaption tglNotSampledT, True
    
    SetToggles Me.tglNotSampledT
    
'    SetQuadratToggles "NotSampled"
'
'    If tglNotSampledT.Caption = strCheck Then
'        'set all no exotics as well
'        '(can't have exotics w/o sampling)
'        tglNoExoticsT.Caption = strCheck
'        tglNoExoticsT.Enabled = False
'
'        SetQuadratToggles "NoExotics"
'
'    Else
'        'enable no exotics if false
'        tglNoExoticsT.Enabled = True
'
'        'clear & enable Q1->3
'        For i = 1 To 3
'            strControl = "tglNotSampledQ" & i
'            Me.Controls(strControl).Enabled = True
'            Me.Controls(strControl).Caption = ""
'        Next
'    End If
'
'
''    'check if Transect level checked
''    If tglNotSampledT.Caption = StringFromCodepoint(uCheck) Then
''
''        'set Q1-Q3 flags & disable
''        For i = 1 To 3
''            strControl = "tglNotSampledQ" & i
''            Controls(strControl).Caption = StringFromCodepoint(uCheck)
''            Controls(strControl).Enabled = False
''        Next
''
''    Else
''
''        'ensure Q1-Q3 flags are enabled & checks are cleared
''        For i = 1 To 3
''            strControl = "tglNotSampledQ" & i
''            Controls(strControl).Caption = ""
''            Controls(strControl).Enabled = True
''        Next
''
''    End If
'
'    If tglNotSampledT Then _
'        ReadyForSave
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - tglNotSampledT_AfterUpdate[frm_Quadrat_Transect form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          tglNotSampledQ1_AfterUpdate
' Description:  Toggle after update actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, March 8, 2017 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 3/8/2017 - initial version
' ---------------------------------
Private Sub tglNotSampledQ1_AfterUpdate()
On Error GoTo Err_Handler
    
    'display as checkbox
    ToggleCaption tglNotSampledQ1, True
    
    SetToggles Me.tglNotSampledQ1
    
'    CheckTransectLevel "NotSampled"
'
'    If tglNotSampledQ1.Caption = strCheck Then
'        'not sampled? -> no priority 1 exotics either
'        tglNoExoticsQ1.Caption = strCheck
'        tglNoExoticsQ1.Enabled = False
'    Else
'        If tglNoExoticsT.Caption <> strCheck Then
'            'sampled? -> priority 1 exotics ok
'            tglNoExoticsQ1.Caption = ""
'            tglNoExoticsQ1.Enabled = True
'        End If
'    End If
'
''    If tglNotSampledQ1 Then _
''        ReadyForSave
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - tglNotSampledQ1_AfterUpdate[frm_Quadrat_Transect form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          tglNotSampledQ2_AfterUpdate
' Description:  Toggle after update actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, March 8, 2017 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 3/8/2017 - initial version
' ---------------------------------
Private Sub tglNotSampledQ2_AfterUpdate()
On Error GoTo Err_Handler
    
    'display as checkbox
    ToggleCaption tglNotSampledQ2, True
    
    CheckTransectLevel "NotSampled"
    
    SetToggles Me.tglNotSampledQ2

'    If tglNotSampledQ2.Caption = strCheck Then
'        'not sampled? -> no priority 1 exotics either
'        tglNoExoticsQ2.Caption = strCheck
'        tglNoExoticsQ2.Enabled = False
'    Else
'        If tglNoExoticsT.Caption <> strCheck Then
'            'sampled? -> priority 1 exotics ok
'            tglNoExoticsQ2.Caption = ""
'            tglNoExoticsQ2.Enabled = True
'        End If
'    End If
    
'    If tglNotSampledQ2 Then _
'        ReadyForSave
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - tglNotSampledQ2_AfterUpdate[frm_Quadrat_Transect form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          tglNotSampledQ3_AfterUpdate
' Description:  Toggle after update actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, March 8, 2017 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 3/8/2017 - initial version
' ---------------------------------
Private Sub tglNotSampledQ3_AfterUpdate()
On Error GoTo Err_Handler
    
    'display as checkbox
    ToggleCaption tglNotSampledQ3, True
    
    CheckTransectLevel "NotSampled"
    
    SetToggles Me.tglNotSampledQ3
    
'    If tglNotSampledQ3.Caption = strCheck Then
'        'not sampled? -> no priority 1 exotics either
'        tglNoExoticsQ3.Caption = strCheck
'        tglNoExoticsQ3.Enabled = False
'    Else
'        If tglNoExoticsT.Caption <> strCheck Then
'            'sampled? -> priority 1 exotics ok
'            tglNoExoticsQ3.Caption = ""
'            tglNoExoticsQ3.Enabled = True
'        End If
'    End If
'
''    If tglNotSampledQ3 Then _
''        ReadyForSave
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - tglNotSampledQ3_AfterUpdate[frm_Quadrat_Transect form])"
    End Select
    Resume Exit_Handler
End Sub

' =================================
'   No Exotics Flag
' =================================
' ---------------------------------
' Sub:          tglNoExoticsT_AfterUpdate
' Description:  Toggle after update actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, March 8, 2017 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 3/8/2017 - initial version
' ---------------------------------
Private Sub tglNoExoticsT_AfterUpdate()
On Error GoTo Err_Handler
    
    Dim i As Integer
    Dim strControl As String
    
    'display as checkbox
    ToggleCaption tglNoExoticsT, True
    
    SetToggles Me.tglNoExoticsT

'    SetQuadratToggles "NoExotics"

'    'check if Transect level checked
'    If tglNoExoticsT.Caption = StringFromCodepoint(uCheck) Then
'
'        'set Q1-Q3 flags & disable
'        For i = 1 To 3
'            strControl = "tglNoExoticsQ" & i
'            Controls(strControl).Caption = StringFromCodepoint(uCheck)
'            Controls(strControl).Enabled = False
'        Next
'
'    Else
'
'        'ensure Q1-Q3 flags are enabled & checks are cleared
'        For i = 1 To 3
'            strControl = "tglNoExoticsQ" & i
'            Controls(strControl).Caption = ""
'            Controls(strControl).Enabled = True
'        Next
'
'    End If
    
    If tglNoExoticsT Then _
        ReadyForSave
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - tglNoExoticsT_AfterUpdate[frm_Quadrat_Transect form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          tglNoExoticsQ1_AfterUpdate
' Description:  Toggle after update actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, March 8, 2017 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 3/8/2017 - initial version
' ---------------------------------
Private Sub tglNoExoticsQ1_AfterUpdate()
On Error GoTo Err_Handler
    
    'display as checkbox
    ToggleCaption tglNoExoticsQ1, True
    
    CheckTransectLevel "NoExotics"
    
    SetToggles Me.tglNoExoticsQ1
    
'    If tglNoExoticsQ1 Then _
'        ReadyForSave
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - tglNoExoticsQ1_AfterUpdate[frm_Quadrat_Transect form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          tglNoExoticsQ2_AfterUpdate
' Description:  Toggle after update actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, March 8, 2017 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 3/8/2017 - initial version
' ---------------------------------
Private Sub tglNoExoticsQ2_AfterUpdate()
On Error GoTo Err_Handler
    
    'display as checkbox
    ToggleCaption tglNoExoticsQ2, True
    
    CheckTransectLevel "NoExotics"
    
    SetToggles Me.tglNoExoticsQ2
    
'    If tglNoExoticsQ2 Then _
'        ReadyForSave
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - tglNoExoticsQ2_AfterUpdate[frm_Quadrat_Transect form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          tglNoExoticsQ3_AfterUpdate
' Description:  Toggle after update actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, March 8, 2017 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 3/8/2017 - initial version
' ---------------------------------
Private Sub tglNoExoticsQ3_AfterUpdate()
On Error GoTo Err_Handler
    
    'display as checkbox
    ToggleCaption tglNoExoticsQ3, True
    
    CheckTransectLevel "NoExotics"
    
    SetToggles Me.tglNoExoticsQ3
    
'    If tglNoExoticsQ3 Then _
'        ReadyForSave
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - tglNoExoticsQ3_AfterUpdate[frm_Quadrat_Transect form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          Form_Close
' Description:  form closing actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, March 8, 2017 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 3/8/2016 - initial version
' ---------------------------------
Private Sub Form_Close()
On Error GoTo Err_Handler

    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Form_Close[frm_Quadrat_Transect form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          SetQuadratToggles
' Description:  Set quadrat 1-3 toggles when transect toggle is set
' Assumptions:  -
' Parameters:   strToggle - name of toggle group ("NoExotics" or "NotSampled")
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, March 8, 2017 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 3/8/2017 - initial version
' ---------------------------------
Private Sub SetQuadratToggles(strToggle As String)
On Error GoTo Err_Handler
    
    Dim i As Integer
    Dim strControl As String, strLabel As String, _
        strToggle2 As String, strControl2 As String
        
    'set transect level control
    strControl = "tgl" & strToggle & "T"
'    strCheck = StringFromCodepoint(uCheck)

    'check if Transect level checked
    If Controls(strControl).Caption = strCheck Then
    
        'set Q1-Q3 flags & disable
        For i = 1 To 3
            strControl = "tgl" & strToggle & "Q" & i
            Controls(strControl).Caption = strCheck
            Controls(strControl).Enabled = False
            
            strLabel = "lblQ" & i
            Controls(strLabel).ForeColor = lngGray50
            
            Select Case i
                Case 1
                    Me.fsub_Species_Current!Q1_hm.Enabled = False
                Case 2
                    Me.fsub_Species_Current!Q2_5m.Enabled = False
                Case 3
                    Me.fsub_Species_Current!Q3_10m.Enabled = False
            End Select

        Next
            
        'when transect is either not sampled or has no exotics:
        'disable fsub since no exotic species will be recorded
        Me.fsub_Species_Current.Enabled = False
        Me.fsub_Species_Current!Plant_Code.Enabled = False
        Me.fsub_Species_Current!cbxIsDead.Enabled = False
    Else
        
        'check other toggle
        strToggle2 = IIf(strToggle <> "NoExotics", "NoExotics", "NotSampled")
        strControl2 = "tgl" & strToggle2 & "T"
        
        'ensure Q1-Q3 flags are enabled & checks are cleared
        For i = 1 To 3
            strControl = "tgl" & strToggle & "Q" & i
                        
            If Controls("tgl" & strToggle2 & "Q" & i).Caption <> strCheck Then
                Controls(strControl).Caption = ""
                Controls(strControl).Enabled = True
            End If
            
            strLabel = "lblQ" & i
            Controls(strLabel).ForeColor = lngBlack
        
            're-enable fields
            If Controls("tgl" & strToggle2 & "Q" & i).Caption <> strCheck Then _
            
                Select Case i
                    Case 1
                        Me.fsub_Species_Current!Q1_hm.Enabled = True
                    Case 2
                        Me.fsub_Species_Current!Q2_5m.Enabled = True
                    Case 3
                        Me.fsub_Species_Current!Q3_10m.Enabled = True
                End Select

            End If
        Next
        
        If Me.Controls(strControl2).Caption <> strCheck Then
            
            're-enable fields
            Me.fsub_Species_Current!Plant_Code.Enabled = True
            Me.fsub_Species_Current!cbxIsDead.Enabled = True
            
            Me.fsub_Species_Current.Enabled = True
        
        End If
    End If
    
    ToggleDisabledMessage
    
    'update AvgCover
    Me.fsub_Species_Current!Average_Cover = Me.fsub_Species_Current.Form.CalcAvgCover
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - SetQuadratToggles[frm_Quadrat_Transect form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          CheckTransectLevel
' Description:  Checks if all quadrat checkbox toggles are set
'               If so, the transect level toggle is checked & Q1-3 are disabled
'               If not, all quadrat level toggles remain active & transect level is not checked
' Assumptions:  -
' Parameters:   strToggle - name of toggle group ("NoExotics" or "NotSampled")
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, March 8, 2017 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 3/8/2016 - initial version
' ---------------------------------
Private Sub CheckTransectLevel(strToggle As String)
On Error GoTo Err_Handler

    Dim i As Integer, Count As Integer
    Dim strControl As String, strLabel As String, _
         strToggle2 As String, strControl2 As String

    'default
    Count = 0
'    strCheck = StringFromCodepoint(uCheck)

    'check @ quadrat's checkbox toggle
    For i = 1 To 3
    
        strControl = "tgl" & strToggle & "Q" & i
    
        If Controls(strControl).Caption = strCheck Then
            Count = Count + i
            
            'disable species cover field
            Select Case i
                Case 1
                    Me.fsub_Species_Current!Q1_hm.Enabled = False
                Case 2
                    Me.fsub_Species_Current!Q2_5m.Enabled = False
                Case 3
                    Me.fsub_Species_Current!Q3_10m.Enabled = False
            End Select
        End If
    
    Next
    
    'set transect control
    strControl = "tgl" & strToggle + "T"
    
    'check if all quadrats are set (if so, count = 1 + 2 + 3 = 6)
    If Count = 6 Then
        
        Controls(strControl).Caption = strCheck
        
        For i = 1 To 3
            strControl = "tgl" & strToggle & "Q" & i
            Controls(strControl).Enabled = False
            
            strLabel = "lblQ" & i
            Controls(strLabel).ForeColor = lngGray50
                
        Next
        
        'disable fsub & controls (no species should be identified)
        Me.fsub_Species_Current.Enabled = False
        
        Me.fsub_Species_Current!Q1_hm.Enabled = False
        Me.fsub_Species_Current!Q2_5m.Enabled = False
        Me.fsub_Species_Current!Q3_10m.Enabled = False
        Me.fsub_Species_Current!Plant_Code.Enabled = False
        Me.fsub_Species_Current!cbxIsDead.Enabled = False
    
    Else
    
        Controls(strControl).Caption = ""
        
        For i = 1 To 3
            strControl = "tgl" & strToggle & "Q" & i
        
            strLabel = "lblQ" & i
            Controls(strLabel).ForeColor = lngBlack
            
            'determine if species cover controls should be enabled
            strToggle2 = IIf(strToggle <> "NoExotics", "NoExotics", "NotSampled")
            strControl2 = "tgl" & strToggle2 & "Q" & i
            
            'check if NotSampled set?
            If InStr(strControl2, "NotSampled") And _
               Controls(strControl2).Caption <> strCheck Then _
                Controls(strControl).Enabled = True
            
            If Controls(strControl).Caption = strCheck Or _
               Controls(strControl2).Caption = strCheck Then
                
                Select Case i
                    Case 1
                        Me.fsub_Species_Current!Q1_hm.Enabled = False
                    Case 2
                        Me.fsub_Species_Current!Q2_5m.Enabled = False
                    Case 3
                        Me.fsub_Species_Current!Q3_10m.Enabled = False
                End Select
            
            ElseIf Controls(strControl).Caption <> strCheck And _
                Controls(strControl2).Caption <> strCheck Then
           
                Select Case i
                    Case 1
                        Me.fsub_Species_Current!Q1_hm.Enabled = True
                    Case 2
                        Me.fsub_Species_Current!Q2_5m.Enabled = True
                    Case 3
                        Me.fsub_Species_Current!Q3_10m.Enabled = True
                End Select
                
                'enable fsub & controls
                Me.fsub_Species_Current.Enabled = True
                
                Me.fsub_Species_Current!Plant_Code.Enabled = True
                Me.fsub_Species_Current!cbxIsDead.Enabled = True
                        
            End If
            
        Next
        
    End If
    
    'NotSampled @ transect level?
    '--> check & disable NoExotics @ transect level
    If Me.tglNotSampledT.Caption = strCheck Then
        Me.tglNoExoticsT.Caption = strCheck
        Me.tglNoExoticsT.Enabled = False
    End If
    
    'all disabled? --> plant_code & cbxIsDead are also disabled
    If Me.fsub_Species_Current!Q1_hm.Enabled = False And _
       Me.fsub_Species_Current!Q2_5m.Enabled = False And _
       Me.fsub_Species_Current!Q3_10m.Enabled = False Then
       
        Me.fsub_Species_Current!Plant_Code.Enabled = False
        Me.fsub_Species_Current!cbxIsDead.Enabled = False
       
    End If
    
    ToggleDisabledMessage
    
    'update AvgCover
    Me.fsub_Species_Current!Average_Cover = fsub_Species_Current.Form.CalcAvgCover
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - CheckTransectLevel[frm_Quadrat_Transect form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          ToggleDisabledMessage
' Description:  Checks if transect level flags are set
'               If both or one is set --> displays 'Disabled' message
'               If neither are is set --> hides 'Disabled' message
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, March 8, 2017 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 3/8/2016 - initial version
' ---------------------------------
Private Sub ToggleDisabledMessage()
On Error GoTo Err_Handler

'    Dim strCaption As String
'
'    strCaption = StringFromCodepoint(uCheck)

    If tglNotSampledT.Caption = strCheck Or _
       tglNoExoticsT.Caption = strCheck Then
       
        'display disabled message
        Me.fsub_Message.Visible = True
        
    Else
         
         'hide message
        Me.fsub_Message.Visible = False
       
    End If
       
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - ToggleDisabledMessage[frm_Quadrat_Transect form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          ToggleSpeciesControls
' Description:  Checks if transect or quadrat level flags are set
'               If transect set --> form disabled
'               If quadrat set --> control for quadrat disabled
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, March 8, 2017 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 3/8/2016 - initial version
' ---------------------------------
Private Sub ToggleSpeciesControls()
On Error GoTo Err_Handler

'    Dim strCaption As String
'
'    strCaption = StringFromCodepoint(uCheck)

    If tglNotSampledT.Caption = strCheck Or _
       tglNoExoticsT.Caption = strCheck Then
       
        'display disabled message
        Me.fsub_Message.Visible = True
        
    Else
         
         'hide message
        Me.fsub_Message.Visible = False
       
    End If
       
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - ToggleSpeciesControls[frm_Quadrat_Transect form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          SetToggles
' Description:  Sets toggles
' Assumptions:
'       NotSampled
'       transect ON --> quadrats & NoExotics ON
'       transect OFF --> quadrats & NoExotics OFF
'       transect ON --> quadrats ON
'       transect OFF --> quadrats OFF
'       NoExotics
'       transect ON --> quadrats ON
'       transect OFF --> quadrats OFF UNLESS NotSampled ON
' Parameters:   ToggleSet - toggle button(s) to set (ToggleButton)
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, March 8, 2017 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 3/8/2016 - initial version
' ---------------------------------
Private Sub SetToggles(ToggleSet As ToggleButton)
On Error GoTo Err_Handler
    
    Dim strToggle As String, strLabel As String
    Dim blnON As Boolean
    Dim i As Integer
    
    blnON = False
    
    If Me.Controls(ToggleSet.Name).Caption = strCheck Then _
        blnON = True
    
    'default
    strToggle = ToggleSet.Name
       
    Select Case Replace(ToggleSet.Name, "tgl", "")
    
    '------------------------------------------
    ' NotSampled
    '   transect ON --> quadrats & NoExotics ON
    '   transect OFF --> quadrats & NoExotics OFF
    '   transect ON --> quadrats ON
    '   transect OFF --> quadrats OFF
    '------------------------------------------
        Case "NotSampledT"
            For i = 1 To 3
                strToggle = "tglNotSampledQ" & i
                Controls(strToggle).Enabled = IIf(blnON, False, True)
                Controls(strToggle).Caption = IIf(blnON, strCheck, "")
                strLabel = "lblQ" & i
                Controls(strLabel).ForeColor = IIf(blnON, lngGray50, lngBlack)
            Next

            'recurse for NoExotics
            Controls("tglNoExoticsT").Caption = strCheck
            SetToggles Me.tglNoExoticsT
            
            'disable NoExotics @ Transect level IF NotSampled @ Transect level checked
            Controls("tglNoExoticsT").Enabled = IIf(blnON, False, True)
        
        Case "NotSampledQ1"
            strLabel = "lblQ1"
            Controls(strLabel).ForeColor = IIf(blnON, lngGray50, lngBlack)
            
            'recurse for NoExotics
            SetToggles Me.tglNoExoticsQ1
            
        Case "NotSampledQ2"
            strLabel = "lblQ2"
            Controls(strLabel).ForeColor = IIf(blnON, lngGray50, lngBlack)
            
            'recurse for NoExotics
            SetToggles Me.tglNoExoticsQ2
        
        Case "NotSampledQ3"
            strLabel = "lblQ3"
            Controls(strLabel).ForeColor = IIf(blnON, lngGray50, lngBlack)
            
            'recurse for NoExotics
            SetToggles Me.tglNoExoticsQ3
            
    '------------------------------------------
    ' NoExotics
    '   transect ON --> quadrats ON
    '   transect OFF --> quadrats OFF UNLESS NotSampled ON
    '------------------------------------------
        Case "NoExoticsT"
            For i = 1 To 3
                strToggle = "tglNoExoticsQ" & i
                Controls(strToggle).Enabled = IIf(blnON, False, True)
                Controls(strToggle).Caption = IIf(blnON, strCheck, "")
            Next
            
            ToggleDisabledMessage
            
        Case "NoExoticsQ1"
            If Not tglNotSampledQ1.Caption = strCheck Then
                Controls(strToggle).Enabled = IIf(blnON, False, True)
                Controls(strToggle).Caption = IIf(blnON, strCheck, "")
            Else
                
            End If
        Case "NoExoticsQ2"
            If Not tglNotSampledQ2.Caption = strCheck Then
                Controls(strToggle).Enabled = IIf(blnON, False, True)
                Controls(strToggle).Caption = IIf(blnON, strCheck, "")
            End If
        Case "NoExoticsQ3"
            If Not tglNotSampledQ3.Caption = strCheck Then
                Controls(strToggle).Enabled = IIf(blnON, False, True)
                Controls(strToggle).Caption = IIf(blnON, strCheck, "")
            End If
    End Select

    If Me.Controls(ToggleSet.Name).Caption = strCheck Then
 
        With fsub_Species_Current
            'form
            .Enabled = IIf(blnON, False, True)
            'fields
            !Plant_Code.Enabled = IIf(blnON, False, True)
            !Q1_hm.Enabled = IIf(blnON, False, True)
            !Q2_5m.Enabled = IIf(blnON, False, True)
            !Q3_10m.Enabled = IIf(blnON, False, True)
            !cbxIsDead.Enabled = IIf(blnON, False, True)
        End With
    End If
    
    ' ---------------------------------------------
    ' NotSampled Q1|Q2|Q3 --> set NoExotics Q1|Q2|Q3
    '                         Q1|Q2|Q3 label "disabled" (grayed out)
    '                         Q1_hm|Q2_5m|Q3_10m fields disabled
    '                         if all --> Species & IsDead fields disabled
    ' ---------------------------------------------
    Dim Count As Integer
    Dim strControl As String, strControlQ As String
    
    'quadrat level?
    If InStr(strToggle, "Q", "") Then
        strControl = Left(strToggle, Len(strToggle) - 1) 'remove last 1|2|3
    
        Count = 0
        
        For i = 1 To 3
            strControlQ = strControl & i
            If Controls(strControlQ).Caption = strCheck Then
                Count = Count + 1
            End If
        Next
    
        'all quadrats set? (if so, count = 1 + 2 + 3 = 6)
        '
        If Count < 6 Then
            
            Debug.Print strControlQ & "6"
        
        End If
    
    End If
    
'            If Controls(strControl).Caption = strCheck Or _
'               Controls(strControl2).Caption = strCheck Then
'
'                Select Case i
'                    Case 1
'                        Me.fsub_Species_Current!Q1_hm.Enabled = False
'                    Case 2
'                        Me.fsub_Species_Current!Q2_5m.Enabled = False
'                    Case 3
'                        Me.fsub_Species_Current!Q3_10m.Enabled = False
'                End Select
'
'            ElseIf Controls(strControl).Caption <> strCheck And _
'                Controls(strControl2).Caption <> strCheck Then
'
'                Select Case i
'                    Case 1
'                        Me.fsub_Species_Current!Q1_hm.Enabled = True
'                    Case 2
'                        Me.fsub_Species_Current!Q2_5m.Enabled = True
'                    Case 3
'                        Me.fsub_Species_Current!Q3_10m.Enabled = True
'                End Select
'
'                'enable fsub & controls
'                Me.fsub_Species_Current.Enabled = True
'
'                Me.fsub_Species_Current!Plant_Code.Enabled = True
'                Me.fsub_Species_Current!cbxIsDead.Enabled = True
'
'            End If
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - SetToggles[frm_Quadrat_Transect form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          DisableToggles
' Description:  Disables NotSampled & NoExotics toggles @ all levels (Transect & Quadrat)
' Assumptions:  When disabled, all toggle values should be cleared
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, April 21, 2017 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 4/21/2016 - initial version
' ---------------------------------
Private Sub DisableToggles()
On Error GoTo Err_Handler
    
    '------------------------------------------
    ' Species in subform ==> clear & disable toggles
    '   NotSampled
    '       transect OFF --> quadrats & NoExotics OFF
    '   NoExotics
    '       transect OFF --> quadrats OFF
    '------------------------------------------
    Dim aryToggles() As String
    Dim tgl As Variant
    Dim tglName As String, tglNameT As String
    Dim strToggle As String, strLabel As String
    Dim i As Integer
    
    'use split for string (vs. variant) array
    aryToggles = Split("NotSampled,NoExotics", ",")
    
    For Each tgl In aryToggles
        'set toggle name
        tglName = "tgl" & tgl
    
        'clear & disable transect toggles
        tglNameT = tglName & "T"
        Controls(tglNameT).Enabled = False
        Controls(tglNameT).Caption = ""
        Controls(tglNameT).ForeColor = lngGray50
        Controls("lblTransect").ForeColor = lngGray50
    
        'clear & disable quadrat toggles
        For i = 1 To 3
            'NotSampled/NoExotics
            strToggle = tglName & "Q" & i
            Controls(strToggle).Enabled = False
            Controls(strToggle).Caption = ""
            Controls(strToggle).ForeColor = lngGray50
            strLabel = "lblQ" & i
            Controls(strLabel).ForeColor = lngGray50
        Next
    Next
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - DisableToggles[frm_Quadrat_Transect form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          EnableToggles
' Description:  Enables NotSampled & NoExotics toggles @ all levels (Transect & Quadrat)
' Assumptions:  When enabled, all toggle values are left alone
' Parameters:   quadrat - number of the quadrat toggles to enable (integer, optional)
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, April 21, 2017 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 4/21/2016 - initial version
' ---------------------------------
Private Sub EnableToggles(Optional Quadrat As Integer)
On Error GoTo Err_Handler
    
    '------------------------------------------
    ' No Species in subform ==> enable toggles
    '   NotSampled
    '       transect ON --> quadrats & NoExotics ON
    '   NoExotics
    '       transect ON --> quadrats ON
    '------------------------------------------
    Dim aryToggles() As String
    Dim tgl As Variant
    Dim tglName As String, tglNameT As String
    Dim strToggle As String, strLabel As String
    Dim i As Integer
    
    'use split for string (vs. variant) array
    aryToggles = Split("NotSampled,NoExotics", ",")
    
    For Each tgl In aryToggles
        'set toggle name
        tglName = "tgl" & tgl
    
        'clear & Enable transect toggles
        tglNameT = tglName & "T"
        Controls(tglNameT).Enabled = True
        Controls(tglNameT).ForeColor = lngBlack
        Controls("lblTransect").ForeColor = lngBlack
    
        'handle individual quadrats if specified
        'or all quadrats if not
        If Quadrat > 0 Then
            
            'disable transect level (can't use if there's a quadrat w/ data)
            Controls(tglNameT).Enabled = False
            Controls(tglNameT).ForeColor = lngGray50
            
            'NotSampled/NoExotics
            strToggle = tglName & "Q" & Quadrat
            Controls(strToggle).Enabled = True
            Controls(strToggle).ForeColor = lngBlack
            strLabel = "lblQ" & Quadrat
            Controls(strLabel).ForeColor = lngBlack
        
        Else
            'enable all quadrat toggles
            For i = 1 To 3
                'NotSampled/NoExotics
                strToggle = tglName & "Q" & i
                Controls(strToggle).Enabled = True
                Controls(strToggle).ForeColor = lngBlack
                strLabel = "lblQ" & i
                Controls(strLabel).ForeColor = lngBlack
            Next
        End If
    Next
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - EnableToggles[frm_Quadrat_Transect form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          PopulateMicrohabitats
' Description:  Checks if form is ready for save
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, April 24, 2017 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 4/24/2016 - initial version
' ---------------------------------
Private Sub PopulateMicrohabitats()
On Error GoTo Err_Handler

    Dim rs As DAO.Recordset
    Dim strField As String
    
    'skip if NULL
    If IsNull(Me.tbxTransectID) Then GoTo Exit_Handler
    
    'set the transect ID
    SetTempVar "Transect_ID", CStr(Me.tbxTransectID)
    
    Set rs = GetRecords("s_surfacecover_by_transect")
    
    If Not (rs.BOF And rs.EOF) Then
        Do Until rs.EOF
        
            'set field name
            'strField = rs("ControlName")
            
            'Debug.Print strField
        
            'populate the field
            Me.Controls(rs("ControlName")) = rs("PercentCover")
            
            'set the tempvar for Quadrat ID (1,2,3)
            SetTempVar "Q" & rs("Quadrat") & "_ID", CInt(rs("Quadrat_ID"))
            
            rs.MoveNext
        Loop
    End If
        
    'populate Q1-3 IDs
    tbxQ1 = Nz(TempVars("Q1_ID"), 0)
    tbxQ2 = Nz(TempVars("Q2_ID"), 0)
    tbxQ3 = Nz(TempVars("Q3_ID"), 0)
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - PopulateMicrohabitats[frm_Quadrat_Transect form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          ReadyForSave
' Description:  Checks if form is ready for save
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, March 8, 2017 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 3/8/2016 - initial version
' ---------------------------------
Private Sub ReadyForSave()
On Error GoTo Err_Handler

    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - ReadyForSave[frm_Quadrat_Transect form])"
    End Select
    Resume Exit_Handler
End Sub
