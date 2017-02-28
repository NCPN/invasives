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
    Width =13680
    DatasheetFontHeight =9
    ItemSuffix =43
    Left =570
    Top =90
    Right =14085
    Bottom =7665
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0xb5100b474c2ee340
    End
    RecordSource ="qry_Quadrat_Transect"
    Caption ="frm_Canopy_Transect"
    OnCurrent ="[Event Procedure]"
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
        Begin FormHeader
            Height =0
            BackColor =-2147483633
            Name ="FormHeader"
        End
        Begin Section
            CanGrow = NotDefault
            Height =9360
            BackColor =-2147483633
            Name ="Detail"
            Begin
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
                    Name ="Visit_Date"
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
                    Name ="ButtonPrevious"
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
                    Name ="ButtonNext"
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
                    Left =12660
                    Top =60
                    Width =840
                    Height =180
                    TabIndex =6
                    Name ="GPS_Time"
                    ControlSource ="GPS_Time"
                    Format ="Long Time"
                    StatusBarText ="Recording time"

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
                    Name ="Observer"
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
                    Left =840
                    Top =1620
                    Width =1860
                    Height =480
                    FontWeight =700
                    Name ="Label55"
                    Caption ="Microhabitat"
                End
                Begin Label
                    OldBorderStyle =1
                    BorderWidth =1
                    OverlapFlags =95
                    TextAlign =2
                    Left =2700
                    Top =1860
                    Width =778
                    Height =240
                    FontWeight =700
                    Name ="Label57"
                    Caption ="Q1"
                    LayoutCachedLeft =2700
                    LayoutCachedTop =1860
                    LayoutCachedWidth =3478
                    LayoutCachedHeight =2100
                End
                Begin Label
                    OldBorderStyle =1
                    BorderWidth =1
                    OverlapFlags =95
                    TextAlign =2
                    Left =3480
                    Top =1860
                    Width =780
                    Height =240
                    FontWeight =700
                    Name ="Label73"
                    Caption ="Q2"
                    LayoutCachedLeft =3480
                    LayoutCachedTop =1860
                    LayoutCachedWidth =4260
                    LayoutCachedHeight =2100
                End
                Begin Label
                    OldBorderStyle =1
                    BorderWidth =1
                    OverlapFlags =95
                    TextAlign =2
                    Left =4259
                    Top =1860
                    Width =780
                    Height =240
                    FontWeight =700
                    Name ="Label74"
                    Caption ="Q3"
                    LayoutCachedLeft =4259
                    LayoutCachedTop =1860
                    LayoutCachedWidth =5039
                    LayoutCachedHeight =2100
                End
                Begin Label
                    OldBorderStyle =1
                    BorderWidth =1
                    OverlapFlags =87
                    TextAlign =2
                    Left =2700
                    Top =1620
                    Width =2340
                    Height =240
                    FontWeight =700
                    Name ="Label76"
                    Caption ="% cover"
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =95
                    IMESentenceMode =3
                    ListRows =21
                    Left =2700
                    Top =2100
                    Width =778
                    TabIndex =9
                    ColumnInfo ="\"\";\"\";\"6\";\"4\""
                    Name ="Dead_Wood_Q1"
                    ControlSource ="Dead_Wood_Q1"
                    RowSourceType ="Table/Query"
                    RowSource ="tlu_Cover_Code"
                    StatusBarText ="Dead wood cover percentage quadrat 1"

                    Begin
                        Begin Label
                            OldBorderStyle =1
                            OverlapFlags =95
                            Left =840
                            Top =2100
                            Width =1860
                            Height =240
                            Name ="Label289"
                            Caption ="Dead Wood:"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =95
                    IMESentenceMode =3
                    ListRows =21
                    Left =3490
                    Top =2100
                    Width =778
                    TabIndex =10
                    ColumnInfo ="\"\";\"\";\"6\";\"4\""
                    Name ="Dead_Wood_Q2"
                    ControlSource ="Dead_Wood_Q2"
                    RowSourceType ="Table/Query"
                    RowSource ="tlu_Cover_Code"
                    StatusBarText ="Dead wood cover percentage quadrat 2"

                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =95
                    IMESentenceMode =3
                    ListRows =21
                    Left =4270
                    Top =2100
                    Width =778
                    TabIndex =11
                    ColumnInfo ="\"\";\"\";\"6\";\"4\""
                    Name ="Dead_Wood_Q3"
                    ControlSource ="Dead_Wood_Q3"
                    RowSourceType ="Table/Query"
                    RowSource ="tlu_Cover_Code"
                    StatusBarText ="Dead wood cover percentage quadrat 3"

                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =95
                    IMESentenceMode =3
                    ListRows =21
                    Left =2700
                    Top =2340
                    Width =778
                    TabIndex =12
                    ColumnInfo ="\"\";\"\";\"6\";\"4\""
                    Name ="Dung_Q1"
                    ControlSource ="Dung_Q1"
                    RowSourceType ="Table/Query"
                    RowSource ="tlu_Cover_Code"
                    StatusBarText ="Dung cover percentage quadrat 1"

                    Begin
                        Begin Label
                            OldBorderStyle =1
                            OverlapFlags =95
                            Left =840
                            Top =2340
                            Width =1860
                            Height =240
                            Name ="Label292"
                            Caption ="Dung"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =95
                    IMESentenceMode =3
                    ListRows =21
                    Left =3490
                    Top =2340
                    Width =778
                    TabIndex =13
                    ColumnInfo ="\"\";\"\";\"6\";\"4\""
                    Name ="Dung_Q2"
                    ControlSource ="Dung_Q2"
                    RowSourceType ="Table/Query"
                    RowSource ="tlu_Cover_Code"
                    StatusBarText ="Dung cover percentage quadrat 2"

                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =95
                    IMESentenceMode =3
                    ListRows =21
                    Left =4270
                    Top =2340
                    Width =778
                    TabIndex =14
                    ColumnInfo ="\"\";\"\";\"6\";\"4\""
                    Name ="Dung_Q3"
                    ControlSource ="Dung_Q3"
                    RowSourceType ="Table/Query"
                    RowSource ="tlu_Cover_Code"
                    StatusBarText ="Dung cover percentage quadrat 3"

                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =95
                    IMESentenceMode =3
                    ListRows =21
                    Left =2700
                    Top =2580
                    Width =778
                    TabIndex =15
                    ColumnInfo ="\"\";\"\";\"6\";\"4\""
                    Name ="Fungus_Q1"
                    ControlSource ="Fungus_Q1"
                    RowSourceType ="Table/Query"
                    RowSource ="tlu_Cover_Code"
                    StatusBarText ="Fungus cover percentage quadrat 1"

                    Begin
                        Begin Label
                            OldBorderStyle =1
                            OverlapFlags =95
                            Left =840
                            Top =2580
                            Width =1860
                            Height =240
                            Name ="Label295"
                            Caption ="Fungus"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =95
                    IMESentenceMode =3
                    ListRows =21
                    Left =3490
                    Top =2580
                    Width =778
                    TabIndex =16
                    ColumnInfo ="\"\";\"\";\"6\";\"4\""
                    Name ="Fungus_Q2"
                    ControlSource ="Fungus_Q2"
                    RowSourceType ="Table/Query"
                    RowSource ="tlu_Cover_Code"
                    StatusBarText ="Fungus cover percentage quadrat 2"

                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =95
                    IMESentenceMode =3
                    ListRows =21
                    Left =4270
                    Top =2580
                    Width =778
                    TabIndex =17
                    ColumnInfo ="\"\";\"\";\"6\";\"4\""
                    Name ="Fungus_Q3"
                    ControlSource ="Fungus_Q3"
                    RowSourceType ="Table/Query"
                    RowSource ="tlu_Cover_Code"
                    StatusBarText ="Fungus cover percentage quadrat 3"

                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =95
                    IMESentenceMode =3
                    ListRows =21
                    Left =2700
                    Top =2820
                    Width =778
                    TabIndex =18
                    ColumnInfo ="\"\";\"\";\"6\";\"4\""
                    Name ="Lichen_Q1"
                    ControlSource ="Lichen_Q1"
                    RowSourceType ="Table/Query"
                    RowSource ="tlu_Cover_Code"
                    StatusBarText ="Lichen cover percentage quadrat 1"

                    Begin
                        Begin Label
                            OldBorderStyle =1
                            OverlapFlags =95
                            Left =840
                            Top =2820
                            Width =1860
                            Height =240
                            Name ="Label298"
                            Caption ="Lichen"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =95
                    IMESentenceMode =3
                    ListRows =21
                    Left =3490
                    Top =2820
                    Width =778
                    TabIndex =19
                    ColumnInfo ="\"\";\"\";\"6\";\"4\""
                    Name ="Lichen_Q2"
                    ControlSource ="Lichen_Q2"
                    RowSourceType ="Table/Query"
                    RowSource ="tlu_Cover_Code"
                    StatusBarText ="Lichen cover percentage quadrat 2"

                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =95
                    IMESentenceMode =3
                    ListRows =21
                    Left =4270
                    Top =2820
                    Width =778
                    TabIndex =20
                    ColumnInfo ="\"\";\"\";\"6\";\"4\""
                    Name ="Lichen_Q3"
                    ControlSource ="Lichen_Q3"
                    RowSourceType ="Table/Query"
                    RowSource ="tlu_Cover_Code"
                    StatusBarText ="Lichen cover percentage quadrat 3"

                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =95
                    IMESentenceMode =3
                    ListRows =21
                    Left =2700
                    Top =3060
                    Width =778
                    TabIndex =21
                    ColumnInfo ="\"\";\"\";\"6\";\"4\""
                    Name ="Litter_Duff_Q1"
                    ControlSource ="Litter_Duff_Q1"
                    RowSourceType ="Table/Query"
                    RowSource ="tlu_Cover_Code"
                    StatusBarText ="Litter/Duff cover percentage quadrat 1"

                    Begin
                        Begin Label
                            OldBorderStyle =1
                            OverlapFlags =95
                            Left =840
                            Top =3060
                            Width =1860
                            Height =240
                            Name ="Label301"
                            Caption ="Litter Duff"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =95
                    IMESentenceMode =3
                    ListRows =21
                    Left =3490
                    Top =3060
                    Width =778
                    TabIndex =22
                    ColumnInfo ="\"\";\"\";\"6\";\"4\""
                    Name ="Litter_Duff_Q2"
                    ControlSource ="Litter_Duff_Q2"
                    RowSourceType ="Table/Query"
                    RowSource ="tlu_Cover_Code"
                    StatusBarText ="Litter/Duff cover percentage quadrat 2"

                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =95
                    IMESentenceMode =3
                    ListRows =21
                    Left =4270
                    Top =3060
                    Width =778
                    TabIndex =23
                    ColumnInfo ="\"\";\"\";\"6\";\"4\""
                    Name ="Litter_Duff_Q3"
                    ControlSource ="Litter_Duff_Q3"
                    RowSourceType ="Table/Query"
                    RowSource ="tlu_Cover_Code"
                    StatusBarText ="Litter/Duff cover percentage quadrat 3"

                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =95
                    IMESentenceMode =3
                    ListRows =21
                    Left =2700
                    Top =3300
                    Width =778
                    TabIndex =24
                    ColumnInfo ="\"\";\"\";\"6\";\"4\""
                    Name ="Root_Bole_Q1"
                    ControlSource ="Root_Bole_Q1"
                    RowSourceType ="Table/Query"
                    RowSource ="tlu_Cover_Code"
                    StatusBarText ="Root/Bole cover percentage quadrat 1"

                    Begin
                        Begin Label
                            OldBorderStyle =1
                            OverlapFlags =95
                            Left =840
                            Top =3300
                            Width =1860
                            Height =240
                            Name ="Label304"
                            Caption ="Live root/Bole"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =95
                    IMESentenceMode =3
                    ListRows =21
                    Left =3490
                    Top =3300
                    Width =778
                    TabIndex =25
                    ColumnInfo ="\"\";\"\";\"6\";\"4\""
                    Name ="Root_Bole_Q2"
                    ControlSource ="Root_Bole_Q2"
                    RowSourceType ="Table/Query"
                    RowSource ="tlu_Cover_Code"
                    StatusBarText ="Roo/Bole cover percentage quadrat 2"

                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =95
                    IMESentenceMode =3
                    ListRows =21
                    Left =4270
                    Top =3300
                    Width =778
                    TabIndex =26
                    ColumnInfo ="\"\";\"\";\"6\";\"4\""
                    Name ="Root_Bole_Q3"
                    ControlSource ="Root_Bole_Q3"
                    RowSourceType ="Table/Query"
                    RowSource ="tlu_Cover_Code"
                    StatusBarText ="Root/Bole cover percentage quadrat 3"

                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =95
                    IMESentenceMode =3
                    ListRows =21
                    Left =2700
                    Top =3780
                    Width =778
                    Height =300
                    TabIndex =30
                    ColumnInfo ="\"\";\"\";\"6\";\"4\""
                    Name ="Mineral_Soil_Sediment_Q1"
                    ControlSource ="Mineral_Soil_Sediment_Q1"
                    RowSourceType ="Table/Query"
                    RowSource ="tlu_Cover_Code"
                    StatusBarText ="Mineral Soil/Sediment cover percentage quadrat 1"

                    Begin
                        Begin Label
                            OldBorderStyle =1
                            OverlapFlags =93
                            Left =840
                            Top =3780
                            Width =1858
                            Height =240
                            Name ="Label307"
                            Caption ="Mineral Soil/Sediment"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =95
                    IMESentenceMode =3
                    ListRows =21
                    Left =3490
                    Top =3780
                    Width =778
                    Height =300
                    TabIndex =31
                    ColumnInfo ="\"\";\"\";\"6\";\"4\""
                    Name ="Mineral_Soil_Sediment_Q2"
                    ControlSource ="Mineral_Soil_Sediment_Q2"
                    RowSourceType ="Table/Query"
                    RowSource ="tlu_Cover_Code"
                    StatusBarText ="Mineral Soil/Sediment cover percentage quadrat 2"

                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =95
                    IMESentenceMode =3
                    ListRows =21
                    Left =4270
                    Top =3780
                    Width =778
                    Height =300
                    TabIndex =32
                    ColumnInfo ="\"\";\"\";\"6\";\"4\""
                    Name ="Mineral_Soil_Sediment_Q3"
                    ControlSource ="Mineral_Soil_Sediment_Q3"
                    RowSourceType ="Table/Query"
                    RowSource ="tlu_Cover_Code"
                    StatusBarText ="Mineral Soil/Sediment cover percentage quadrat 3"

                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =255
                    IMESentenceMode =3
                    ListRows =21
                    Left =2700
                    Top =4020
                    Width =778
                    TabIndex =33
                    ColumnInfo ="\"\";\"\";\"6\";\"4\""
                    Name ="Moss_Q1"
                    ControlSource ="Moss_Q1"
                    RowSourceType ="Table/Query"
                    RowSource ="tlu_Cover_Code"
                    StatusBarText ="Moss cover percentage quadrat 1"

                    Begin
                        Begin Label
                            OldBorderStyle =1
                            OverlapFlags =223
                            Left =840
                            Top =4020
                            Width =1860
                            Height =240
                            Name ="Label310"
                            Caption ="Moss"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =255
                    IMESentenceMode =3
                    ListRows =21
                    Left =3490
                    Top =4020
                    Width =778
                    TabIndex =34
                    ColumnInfo ="\"\";\"\";\"6\";\"4\""
                    Name ="Moss_Q2"
                    ControlSource ="Moss_Q2"
                    RowSourceType ="Table/Query"
                    RowSource ="tlu_Cover_Code"
                    StatusBarText ="Moss cover percentage quadrat 2"

                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =255
                    IMESentenceMode =3
                    ListRows =21
                    Left =4270
                    Top =4020
                    Width =778
                    TabIndex =35
                    ColumnInfo ="\"\";\"\";\"6\";\"4\""
                    Name ="Moss_Q3"
                    ControlSource ="Moss_Q3"
                    RowSourceType ="Table/Query"
                    RowSource ="tlu_Cover_Code"
                    StatusBarText ="Moss cover percentage quadrat 3"

                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =127
                    IMESentenceMode =3
                    ListRows =21
                    Left =2700
                    Top =4260
                    Width =778
                    TabIndex =36
                    ColumnInfo ="\"\";\"\";\"6\";\"4\""
                    Name ="Cryptogram_Q1"
                    ControlSource ="Cryptogram_Q1"
                    RowSourceType ="Table/Query"
                    RowSource ="tlu_Cover_Code"
                    StatusBarText ="Cryptogram cover percentage quadrat 1"

                    Begin
                        Begin Label
                            OldBorderStyle =1
                            OverlapFlags =127
                            Left =840
                            Top =4260
                            Width =1860
                            Height =240
                            Name ="Label313"
                            Caption ="Biological Soil Crust"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =127
                    IMESentenceMode =3
                    ListRows =21
                    Left =3490
                    Top =4260
                    Width =778
                    TabIndex =37
                    ColumnInfo ="\"\";\"\";\"6\";\"4\""
                    Name ="Cryptogram_Q2"
                    ControlSource ="Cryptogram_Q2"
                    RowSourceType ="Table/Query"
                    RowSource ="tlu_Cover_Code"
                    StatusBarText ="Cryptogram cover percentage quadrat 2"

                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =127
                    IMESentenceMode =3
                    ListRows =21
                    Left =4270
                    Top =4260
                    Width =778
                    TabIndex =38
                    ColumnInfo ="\"\";\"\";\"6\";\"4\""
                    Name ="Cryptogram_Q3"
                    ControlSource ="Cryptogram_Q3"
                    RowSourceType ="Table/Query"
                    RowSource ="tlu_Cover_Code"
                    StatusBarText ="Cryptogram cover percentage quadrat 3"

                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =127
                    IMESentenceMode =3
                    ListRows =21
                    Left =2700
                    Top =4500
                    Width =778
                    TabIndex =39
                    ColumnInfo ="\"\";\"\";\"6\";\"4\""
                    Name ="Road_Q1"
                    ControlSource ="Road_Q1"
                    RowSourceType ="Table/Query"
                    RowSource ="tlu_Cover_Code"
                    StatusBarText ="Road cover percentage quadrat 1"

                    Begin
                        Begin Label
                            OldBorderStyle =1
                            OverlapFlags =127
                            Left =840
                            Top =4500
                            Width =1860
                            Height =240
                            Name ="Label316"
                            Caption ="Road"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =127
                    IMESentenceMode =3
                    ListRows =21
                    Left =3490
                    Top =4500
                    Width =778
                    TabIndex =40
                    ColumnInfo ="\"\";\"\";\"6\";\"4\""
                    Name ="Road_Q2"
                    ControlSource ="Road_Q2"
                    RowSourceType ="Table/Query"
                    RowSource ="tlu_Cover_Code"
                    StatusBarText ="Road cover percentage quadrat 2"

                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =127
                    IMESentenceMode =3
                    ListRows =21
                    Left =4270
                    Top =4500
                    Width =778
                    TabIndex =41
                    ColumnInfo ="\"\";\"\";\"6\";\"4\""
                    Name ="Road_Q3"
                    ControlSource ="Road_Q3"
                    RowSourceType ="Table/Query"
                    RowSource ="tlu_Cover_Code"
                    StatusBarText ="Road cover percentage quadrat 3"

                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =127
                    IMESentenceMode =3
                    ListRows =21
                    Left =2700
                    Top =4740
                    Width =778
                    TabIndex =42
                    ColumnInfo ="\"\";\"\";\"6\";\"4\""
                    Name ="Rock_Q1"
                    ControlSource ="Rock_Q1"
                    RowSourceType ="Table/Query"
                    RowSource ="tlu_Cover_Code"
                    StatusBarText ="Rock cover percentage quadrat 1"

                    Begin
                        Begin Label
                            OldBorderStyle =1
                            OverlapFlags =127
                            Left =840
                            Top =4740
                            Width =1860
                            Height =240
                            Name ="Label319"
                            Caption ="Rock"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =127
                    IMESentenceMode =3
                    ListRows =21
                    Left =3490
                    Top =4740
                    Width =778
                    TabIndex =43
                    ColumnInfo ="\"\";\"\";\"6\";\"4\""
                    Name ="Rock_Q2"
                    ControlSource ="Rock_Q2"
                    RowSourceType ="Table/Query"
                    RowSource ="tlu_Cover_Code"
                    StatusBarText ="Rock cover percentage quadrat 2"

                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =127
                    IMESentenceMode =3
                    ListRows =21
                    Left =4270
                    Top =4740
                    Width =778
                    TabIndex =44
                    ColumnInfo ="\"\";\"\";\"6\";\"4\""
                    Name ="Rock_Q3"
                    ControlSource ="Rock_Q3"
                    RowSourceType ="Table/Query"
                    RowSource ="tlu_Cover_Code"
                    StatusBarText ="Rock cover percentage quadrat 3"

                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =127
                    IMESentenceMode =3
                    ListRows =21
                    Left =2700
                    Top =4980
                    Width =778
                    TabIndex =45
                    ColumnInfo ="\"\";\"\";\"6\";\"4\""
                    Name ="Standing_Water_Flooded_Q1"
                    ControlSource ="Standing_Water_Flooded_Q1"
                    RowSourceType ="Table/Query"
                    RowSource ="tlu_Cover_Code"
                    StatusBarText ="Standing Water/Flooded cover percentage quadrat 1"

                    Begin
                        Begin Label
                            OldBorderStyle =1
                            OverlapFlags =127
                            Left =840
                            Top =4980
                            Width =1860
                            Height =240
                            Name ="Label322"
                            Caption ="Standing Water/Flooded"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =127
                    IMESentenceMode =3
                    ListRows =21
                    Left =3490
                    Top =4980
                    Width =778
                    TabIndex =46
                    ColumnInfo ="\"\";\"\";\"6\";\"4\""
                    Name ="Standing_Water_Flooded_Q2"
                    ControlSource ="Standing_Water_Flooded_Q2"
                    RowSourceType ="Table/Query"
                    RowSource ="tlu_Cover_Code"
                    StatusBarText ="Standing Water/Flooded cover percentage quadrat 2"

                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =127
                    IMESentenceMode =3
                    ListRows =21
                    Left =4270
                    Top =4980
                    Width =778
                    TabIndex =47
                    ColumnInfo ="\"\";\"\";\"6\";\"4\""
                    Name ="Standing_Water_Flooded_Q3"
                    ControlSource ="Standing_Water_Flooded_Q3"
                    RowSourceType ="Table/Query"
                    RowSource ="tlu_Cover_Code"
                    StatusBarText ="Standing Water/Flooded cover percentage quadrat 3"

                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =127
                    IMESentenceMode =3
                    ListRows =21
                    Left =2700
                    Top =5220
                    Width =778
                    TabIndex =48
                    ColumnInfo ="\"\";\"\";\"6\";\"4\""
                    Name ="Stream_Q1"
                    ControlSource ="Stream_Q1"
                    RowSourceType ="Table/Query"
                    RowSource ="tlu_Cover_Code"
                    StatusBarText ="Stream cover percentage quadrat 1"

                    Begin
                        Begin Label
                            OldBorderStyle =1
                            OverlapFlags =127
                            Left =840
                            Top =5220
                            Width =1860
                            Height =240
                            Name ="Label325"
                            Caption ="Stream"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =127
                    IMESentenceMode =3
                    ListRows =21
                    Left =3490
                    Top =5220
                    Width =778
                    TabIndex =49
                    ColumnInfo ="\"\";\"\";\"6\";\"4\""
                    Name ="Stream_Q2"
                    ControlSource ="Stream_Q2"
                    RowSourceType ="Table/Query"
                    RowSource ="tlu_Cover_Code"
                    StatusBarText ="Stream cover percentage quadrat 2"

                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =127
                    IMESentenceMode =3
                    ListRows =21
                    Left =4270
                    Top =5220
                    Width =778
                    TabIndex =50
                    ColumnInfo ="\"\";\"\";\"6\";\"4\""
                    Name ="Stream_Q3"
                    ControlSource ="Stream_Q3"
                    RowSourceType ="Table/Query"
                    RowSource ="tlu_Cover_Code"
                    StatusBarText ="Stream cover percentage quadrat 3"

                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =127
                    IMESentenceMode =3
                    ListRows =21
                    Left =2700
                    Top =5460
                    Width =778
                    TabIndex =51
                    ColumnInfo ="\"\";\"\";\"6\";\"4\""
                    Name ="Trash_Junk_Q1"
                    ControlSource ="Trash_Junk_Q1"
                    RowSourceType ="Table/Query"
                    RowSource ="tlu_Cover_Code"
                    StatusBarText ="Trash/Junk cover percentage quadrat 1"

                    Begin
                        Begin Label
                            OldBorderStyle =1
                            OverlapFlags =127
                            Left =840
                            Top =5460
                            Width =1860
                            Height =240
                            Name ="Label328"
                            Caption ="Trash/Junk"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =127
                    IMESentenceMode =3
                    ListRows =21
                    Left =3490
                    Top =5460
                    Width =778
                    TabIndex =52
                    ColumnInfo ="\"\";\"\";\"6\";\"4\""
                    Name ="Trash_Junk_Q2"
                    ControlSource ="Trash_Junk_Q2"
                    RowSourceType ="Table/Query"
                    RowSource ="tlu_Cover_Code"
                    StatusBarText ="Trash/Junk cover percentage quadrat 2"

                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =119
                    IMESentenceMode =3
                    ListRows =21
                    Left =4270
                    Top =5460
                    Width =778
                    TabIndex =53
                    ColumnInfo ="\"\";\"\";\"6\";\"4\""
                    Name ="Trash_Junk_Q3"
                    ControlSource ="Trash_Junk_Q3"
                    RowSourceType ="Table/Query"
                    RowSource ="tlu_Cover_Code"
                    StatusBarText ="Trash/Junk cover percentage quadrat 3"

                End
                Begin Subform
                    OverlapFlags =93
                    Left =5760
                    Top =1080
                    Width =7200
                    Height =6718
                    TabIndex =54
                    Name ="fsub_Species"
                    SourceObject ="Form.fsub_Species"
                    LinkChildFields ="Transect_ID"
                    LinkMasterFields ="Transect_ID"

                End
                Begin Subform
                    OverlapFlags =255
                    Left =5760
                    Top =1080
                    Width =7200
                    Height =6717
                    TabIndex =55
                    Name ="fsub_Species_2008"
                    SourceObject ="Form.fsub_Species_2008"
                    LinkChildFields ="Transect_ID"
                    LinkMasterFields ="Transect_ID"

                End
                Begin Subform
                    OverlapFlags =247
                    Left =5760
                    Top =1080
                    Width =7200
                    Height =6718
                    TabIndex =56
                    Name ="fsub_Species_2009"
                    SourceObject ="Form.fsub_Species_2009"
                    LinkChildFields ="Transect_ID"
                    LinkMasterFields ="Transect_ID"

                End
                Begin ComboBox
                    LimitToList = NotDefault
                    SpecialEffect =0
                    OverlapFlags =215
                    IMESentenceMode =3
                    ListRows =21
                    Left =2700
                    Top =3540
                    Width =778
                    TabIndex =27
                    ColumnInfo ="\"\";\"\";\"6\";\"4\""
                    Name ="Dead_Root_Bole_Q1"
                    ControlSource ="Dead_Root_Bole_Q1"
                    RowSourceType ="Table/Query"
                    RowSource ="tlu_Cover_Code"
                    StatusBarText ="Root/Bole cover percentage quadrat 1"

                    Begin
                        Begin Label
                            OldBorderStyle =1
                            OverlapFlags =223
                            Left =840
                            Top =3540
                            Width =1860
                            Height =240
                            Name ="Label40"
                            Caption ="Dead root/Bole"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    SpecialEffect =0
                    OverlapFlags =223
                    IMESentenceMode =3
                    ListRows =21
                    Left =3490
                    Top =3540
                    Width =778
                    TabIndex =28
                    ColumnInfo ="\"\";\"\";\"6\";\"4\""
                    Name ="Dead_Root_Bole_Q2"
                    ControlSource ="Dead_Root_Bole_Q2"
                    RowSourceType ="Table/Query"
                    RowSource ="tlu_Cover_Code"
                    StatusBarText ="Roo/Bole cover percentage quadrat 2"

                End
                Begin ComboBox
                    LimitToList = NotDefault
                    SpecialEffect =0
                    OverlapFlags =215
                    IMESentenceMode =3
                    ListRows =21
                    Left =4270
                    Top =3540
                    Width =763
                    TabIndex =29
                    ColumnInfo ="\"\";\"\";\"6\";\"4\""
                    Name ="Dead_Root_Bole_Q3"
                    ControlSource ="Dead_Root_Bole_Q3"
                    RowSourceType ="Table/Query"
                    RowSource ="tlu_Cover_Code"
                    StatusBarText ="Root/Bole cover percentage quadrat 3"

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

Private Sub ButtonNext_KeyDown(KeyCode As Integer, Shift As Integer)
  ' Ignore Page Down and Page Up keys for they will cycle through records
  Select Case KeyCode
    Case 33, 34
      KeyCode = 0
    End Select
End Sub

Private Sub ButtonPrevious_KeyDown(KeyCode As Integer, Shift As Integer)
  ' Ignore Page Down and Page Up keys for they will cycle through records
  Select Case KeyCode
    Case 33, 34
      KeyCode = 0
    End Select
End Sub



Private Sub ButtonPrevious_Click()
On Error GoTo Err_ButtonPrevious_Click

  If Me!Transect = 1 Then
    MsgBox "Already on first transect"
  Else
    DoCmd.GoToRecord , , acPrevious
  End If
  
Exit_ButtonPrevious_Click:
    Exit Sub

Err_ButtonPrevious_Click:
    MsgBox Err.Description
    Resume Exit_ButtonPrevious_Click
    
End Sub
Private Sub ButtonNext_Click()
On Error GoTo Err_ButtonNext_Click

  Dim intTransect As Byte

    DoCmd.GoToRecord , , acNext

Exit_ButtonNext_Click:
    Exit Sub

Err_ButtonNext_Click:
    MsgBox Err.Description
    Resume Exit_ButtonNext_Click
    
End Sub

Private Sub Form_Current()
  If Me.Parent!Start_Date < #1/1/2009# Then
    Me!fsub_Species.Visible = False
    Me!fsub_Species_2008.Visible = True
    Me!fsub_Species_2009.Visible = False
  ElseIf Me.Parent!Start_Date < #1/1/2010# Then
    Me!fsub_Species.Visible = False
    Me!fsub_Species_2009.Visible = True
    Me!fsub_Species_2008.Visible = False
  Else
    Me!fsub_Species.Visible = True
    Me!fsub_Species_2008.Visible = False
    Me!fsub_Species_2009.Visible = False
  End If
End Sub

Private Sub Visit_Date_KeyDown(KeyCode As Integer, Shift As Integer)
  ' Ignore Page Down and Page Up keys for they will cycle through records
  Select Case KeyCode
    Case 33, 34
      KeyCode = 0
    End Select
End Sub
