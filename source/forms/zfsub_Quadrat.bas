Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    ScrollBars =0
    TabularFamily =124
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    Cycle =1
    GridX =24
    GridY =24
    Width =13500
    DatasheetFontHeight =9
    ItemSuffix =335
    Left =810
    Top =270
    Right =14310
    Bottom =8895
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x8b370bc14b2ee340
    End
    RecordSource ="qry_Quadrat"
    Caption ="frm_Canopy_Transect"
    OnCurrent ="[Event Procedure]"
    BeforeInsert ="[Event Procedure]"
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
            TextAlign =2
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
            Height =7920
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =11460
                    Top =60
                    Width =630
                    Height =180
                    ColumnWidth =2310
                    Name ="Quadrat_ID"
                    ControlSource ="Quadrat_ID"
                    StatusBarText ="Unique record identifier - primary key"

                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =12300
                    Top =60
                    Width =630
                    Height =180
                    ColumnWidth =2310
                    TabIndex =1
                    Name ="Transect_ID"
                    ControlSource ="Transect_ID"
                    StatusBarText ="M. Link to tbl_Locations (Loc_ID)"

                End
                Begin ComboBox
                    OverlapFlags =87
                    IMESentenceMode =3
                    ColumnCount =3
                    ListWidth =1650
                    Left =1020
                    Top =60
                    Width =1620
                    TabIndex =2
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"\";\"\";\"10\";\"100\""
                    Name ="Observer"
                    ControlSource ="Observer"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT tlu_Contacts.Contact_ID, tlu_Contacts.Last_Name, tlu_Contacts.First_Name "
                        "FROM tlu_Contacts; "
                    ColumnWidths ="0;810;840"
                    OnKeyDown ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =0
                            Left =180
                            Top =60
                            Width =840
                            Height =245
                            FontWeight =700
                            Name ="Observer_Label"
                            Caption ="Observer"
                        End
                    End
                End
                Begin Subform
                    OverlapFlags =93
                    Left =5760
                    Top =480
                    Width =7260
                    Height =6720
                    TabIndex =45
                    Name ="fsub_Species"
                    SourceObject ="Form.fsub_Species"
                    LinkChildFields ="Quadrat_ID"
                    LinkMasterFields ="Quadrat_ID"

                End
                Begin Label
                    OldBorderStyle =1
                    BorderWidth =1
                    OverlapFlags =93
                    Left =660
                    Top =480
                    Width =1860
                    Height =480
                    FontWeight =700
                    Name ="Label55"
                    Caption ="Microhabitat"
                End
                Begin Label
                    OldBorderStyle =1
                    BorderWidth =1
                    OverlapFlags =223
                    Left =2505
                    Top =720
                    Width =778
                    Height =240
                    FontWeight =700
                    Name ="Label57"
                    Caption ="Q1"
                End
                Begin Label
                    OldBorderStyle =1
                    BorderWidth =1
                    OverlapFlags =93
                    Left =3300
                    Top =720
                    Width =778
                    Height =240
                    FontWeight =700
                    Name ="Label73"
                    Caption ="Q2"
                End
                Begin Label
                    OldBorderStyle =1
                    BorderWidth =1
                    OverlapFlags =95
                    Left =4080
                    Top =720
                    Width =778
                    Height =240
                    FontWeight =700
                    Name ="Label74"
                    Caption ="Q3"
                End
                Begin Label
                    OldBorderStyle =1
                    BorderWidth =1
                    OverlapFlags =87
                    Left =2520
                    Top =480
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
                    Left =2520
                    Top =960
                    Width =778
                    TabIndex =3
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
                            TextAlign =0
                            Left =660
                            Top =960
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
                    Left =3310
                    Top =960
                    Width =778
                    TabIndex =4
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
                    Left =4090
                    Top =960
                    Width =778
                    TabIndex =5
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
                    Left =2520
                    Top =1200
                    Width =778
                    TabIndex =6
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
                            TextAlign =0
                            Left =660
                            Top =1200
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
                    Left =3310
                    Top =1200
                    Width =778
                    TabIndex =7
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
                    Left =4090
                    Top =1200
                    Width =778
                    TabIndex =8
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
                    Left =2520
                    Top =1440
                    Width =778
                    TabIndex =9
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
                            TextAlign =0
                            Left =660
                            Top =1440
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
                    Left =3310
                    Top =1440
                    Width =778
                    TabIndex =10
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
                    Left =4090
                    Top =1440
                    Width =778
                    TabIndex =11
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
                    Left =2520
                    Top =1680
                    Width =778
                    TabIndex =12
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
                            TextAlign =0
                            Left =660
                            Top =1680
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
                    Left =3310
                    Top =1680
                    Width =778
                    TabIndex =13
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
                    Left =4090
                    Top =1680
                    Width =778
                    TabIndex =14
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
                    Left =2520
                    Top =1920
                    Width =778
                    TabIndex =15
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
                            TextAlign =0
                            Left =660
                            Top =1920
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
                    Left =3310
                    Top =1920
                    Width =778
                    TabIndex =16
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
                    Left =4090
                    Top =1920
                    Width =778
                    TabIndex =17
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
                    Left =2520
                    Top =2160
                    Width =778
                    TabIndex =18
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
                            TextAlign =0
                            Left =660
                            Top =2160
                            Width =1860
                            Height =240
                            Name ="Label304"
                            Caption ="Root Bole"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =95
                    IMESentenceMode =3
                    ListRows =21
                    Left =3310
                    Top =2160
                    Width =778
                    TabIndex =19
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
                    Left =4090
                    Top =2160
                    Width =778
                    TabIndex =20
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
                    Left =2520
                    Top =2400
                    Width =778
                    TabIndex =21
                    ColumnInfo ="\"\";\"\";\"6\";\"4\""
                    Name ="Mineral_Soil_Sediment_Q1"
                    ControlSource ="Mineral_Soil_Sediment_Q1"
                    RowSourceType ="Table/Query"
                    RowSource ="tlu_Cover_Code"
                    StatusBarText ="Mineral Soil/Sediment cover percentage quadrat 1"

                    Begin
                        Begin Label
                            OldBorderStyle =1
                            OverlapFlags =95
                            TextAlign =0
                            Left =660
                            Top =2400
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
                    Left =3310
                    Top =2400
                    Width =778
                    TabIndex =22
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
                    Left =4090
                    Top =2400
                    Width =778
                    TabIndex =23
                    ColumnInfo ="\"\";\"\";\"6\";\"4\""
                    Name ="Mineral_Soil_Sediment_Q3"
                    ControlSource ="Mineral_Soil_Sediment_Q3"
                    RowSourceType ="Table/Query"
                    RowSource ="tlu_Cover_Code"
                    StatusBarText ="Mineral Soil/Sediment cover percentage quadrat 3"

                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =95
                    IMESentenceMode =3
                    ListRows =21
                    Left =2520
                    Top =2640
                    Width =778
                    TabIndex =24
                    ColumnInfo ="\"\";\"\";\"6\";\"4\""
                    Name ="Moss_Q1"
                    ControlSource ="Moss_Q1"
                    RowSourceType ="Table/Query"
                    RowSource ="tlu_Cover_Code"
                    StatusBarText ="Moss cover percentage quadrat 1"

                    Begin
                        Begin Label
                            OldBorderStyle =1
                            OverlapFlags =95
                            TextAlign =0
                            Left =660
                            Top =2640
                            Width =1860
                            Height =240
                            Name ="Label310"
                            Caption ="Moss"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =95
                    IMESentenceMode =3
                    ListRows =21
                    Left =3310
                    Top =2640
                    Width =778
                    TabIndex =25
                    ColumnInfo ="\"\";\"\";\"6\";\"4\""
                    Name ="Moss_Q2"
                    ControlSource ="Moss_Q2"
                    RowSourceType ="Table/Query"
                    RowSource ="tlu_Cover_Code"
                    StatusBarText ="Moss cover percentage quadrat 2"

                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =95
                    IMESentenceMode =3
                    ListRows =21
                    Left =4090
                    Top =2640
                    Width =778
                    TabIndex =26
                    ColumnInfo ="\"\";\"\";\"6\";\"4\""
                    Name ="Moss_Q3"
                    ControlSource ="Moss_Q3"
                    RowSourceType ="Table/Query"
                    RowSource ="tlu_Cover_Code"
                    StatusBarText ="Moss cover percentage quadrat 3"

                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =95
                    IMESentenceMode =3
                    ListRows =21
                    Left =2520
                    Top =2880
                    Width =778
                    TabIndex =27
                    ColumnInfo ="\"\";\"\";\"6\";\"4\""
                    Name ="Cryptogram_Q1"
                    ControlSource ="Cryptogram_Q1"
                    RowSourceType ="Table/Query"
                    RowSource ="tlu_Cover_Code"
                    StatusBarText ="Cryptogram cover percentage quadrat 1"

                    Begin
                        Begin Label
                            OldBorderStyle =1
                            OverlapFlags =95
                            TextAlign =0
                            Left =660
                            Top =2880
                            Width =1860
                            Height =240
                            Name ="Label313"
                            Caption ="Cryptogram"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =95
                    IMESentenceMode =3
                    ListRows =21
                    Left =3310
                    Top =2880
                    Width =778
                    TabIndex =28
                    ColumnInfo ="\"\";\"\";\"6\";\"4\""
                    Name ="Cryptogram_Q2"
                    ControlSource ="Cryptogram_Q2"
                    RowSourceType ="Table/Query"
                    RowSource ="tlu_Cover_Code"
                    StatusBarText ="Cryptogram cover percentage quadrat 2"

                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =95
                    IMESentenceMode =3
                    ListRows =21
                    Left =4090
                    Top =2880
                    Width =778
                    TabIndex =29
                    ColumnInfo ="\"\";\"\";\"6\";\"4\""
                    Name ="Cryptogram_Q3"
                    ControlSource ="Cryptogram_Q3"
                    RowSourceType ="Table/Query"
                    RowSource ="tlu_Cover_Code"
                    StatusBarText ="Cryptogram cover percentage quadrat 3"

                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =95
                    IMESentenceMode =3
                    ListRows =21
                    Left =2520
                    Top =3120
                    Width =778
                    TabIndex =30
                    ColumnInfo ="\"\";\"\";\"6\";\"4\""
                    Name ="Road_Q1"
                    ControlSource ="Road_Q1"
                    RowSourceType ="Table/Query"
                    RowSource ="tlu_Cover_Code"
                    StatusBarText ="Road cover percentage quadrat 1"

                    Begin
                        Begin Label
                            OldBorderStyle =1
                            OverlapFlags =95
                            TextAlign =0
                            Left =660
                            Top =3120
                            Width =1860
                            Height =240
                            Name ="Label316"
                            Caption ="Road"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =95
                    IMESentenceMode =3
                    ListRows =21
                    Left =3310
                    Top =3120
                    Width =778
                    TabIndex =31
                    ColumnInfo ="\"\";\"\";\"6\";\"4\""
                    Name ="Road_Q2"
                    ControlSource ="Road_Q2"
                    RowSourceType ="Table/Query"
                    RowSource ="tlu_Cover_Code"
                    StatusBarText ="Road cover percentage quadrat 2"

                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =95
                    IMESentenceMode =3
                    ListRows =21
                    Left =4090
                    Top =3120
                    Width =778
                    TabIndex =32
                    ColumnInfo ="\"\";\"\";\"6\";\"4\""
                    Name ="Road_Q3"
                    ControlSource ="Road_Q3"
                    RowSourceType ="Table/Query"
                    RowSource ="tlu_Cover_Code"
                    StatusBarText ="Road cover percentage quadrat 3"

                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =95
                    IMESentenceMode =3
                    ListRows =21
                    Left =2520
                    Top =3360
                    Width =778
                    TabIndex =33
                    ColumnInfo ="\"\";\"\";\"6\";\"4\""
                    Name ="Rock_Q1"
                    ControlSource ="Rock_Q1"
                    RowSourceType ="Table/Query"
                    RowSource ="tlu_Cover_Code"
                    StatusBarText ="Rock cover percentage quadrat 1"

                    Begin
                        Begin Label
                            OldBorderStyle =1
                            OverlapFlags =95
                            TextAlign =0
                            Left =660
                            Top =3360
                            Width =1860
                            Height =240
                            Name ="Label319"
                            Caption ="Rock"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =95
                    IMESentenceMode =3
                    ListRows =21
                    Left =3310
                    Top =3360
                    Width =778
                    TabIndex =34
                    ColumnInfo ="\"\";\"\";\"6\";\"4\""
                    Name ="Rock_Q2"
                    ControlSource ="Rock_Q2"
                    RowSourceType ="Table/Query"
                    RowSource ="tlu_Cover_Code"
                    StatusBarText ="Rock cover percentage quadrat 2"

                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =95
                    IMESentenceMode =3
                    ListRows =21
                    Left =4090
                    Top =3360
                    Width =778
                    TabIndex =35
                    ColumnInfo ="\"\";\"\";\"6\";\"4\""
                    Name ="Rock_Q3"
                    ControlSource ="Rock_Q3"
                    RowSourceType ="Table/Query"
                    RowSource ="tlu_Cover_Code"
                    StatusBarText ="Rock cover percentage quadrat 3"

                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =95
                    IMESentenceMode =3
                    ListRows =21
                    Left =2520
                    Top =3600
                    Width =778
                    TabIndex =36
                    ColumnInfo ="\"\";\"\";\"6\";\"4\""
                    Name ="Standing_Water_Flooded_Q1"
                    ControlSource ="Standing_Water_Flooded_Q1"
                    RowSourceType ="Table/Query"
                    RowSource ="tlu_Cover_Code"
                    StatusBarText ="Standing Water/Flooded cover percentage quadrat 1"

                    Begin
                        Begin Label
                            OldBorderStyle =1
                            OverlapFlags =95
                            TextAlign =0
                            Left =660
                            Top =3600
                            Width =1860
                            Height =240
                            Name ="Label322"
                            Caption ="Standing Water/Flooded"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =95
                    IMESentenceMode =3
                    ListRows =21
                    Left =3310
                    Top =3600
                    Width =778
                    TabIndex =37
                    ColumnInfo ="\"\";\"\";\"6\";\"4\""
                    Name ="Standing_Water_Flooded_Q2"
                    ControlSource ="Standing_Water_Flooded_Q2"
                    RowSourceType ="Table/Query"
                    RowSource ="tlu_Cover_Code"
                    StatusBarText ="Standing Water/Flooded cover percentage quadrat 2"

                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =95
                    IMESentenceMode =3
                    ListRows =21
                    Left =4090
                    Top =3600
                    Width =778
                    TabIndex =38
                    ColumnInfo ="\"\";\"\";\"6\";\"4\""
                    Name ="Standing_Water_Flooded_Q3"
                    ControlSource ="Standing_Water_Flooded_Q3"
                    RowSourceType ="Table/Query"
                    RowSource ="tlu_Cover_Code"
                    StatusBarText ="Standing Water/Flooded cover percentage quadrat 3"

                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =95
                    IMESentenceMode =3
                    ListRows =21
                    Left =2520
                    Top =3840
                    Width =778
                    TabIndex =39
                    ColumnInfo ="\"\";\"\";\"6\";\"4\""
                    Name ="Stream_Q1"
                    ControlSource ="Stream_Q1"
                    RowSourceType ="Table/Query"
                    RowSource ="tlu_Cover_Code"
                    StatusBarText ="Stream cover percentage quadrat 1"

                    Begin
                        Begin Label
                            OldBorderStyle =1
                            OverlapFlags =95
                            TextAlign =0
                            Left =660
                            Top =3840
                            Width =1860
                            Height =240
                            Name ="Label325"
                            Caption ="Stream"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =95
                    IMESentenceMode =3
                    ListRows =21
                    Left =3310
                    Top =3840
                    Width =778
                    TabIndex =40
                    ColumnInfo ="\"\";\"\";\"6\";\"4\""
                    Name ="Stream_Q2"
                    ControlSource ="Stream_Q2"
                    RowSourceType ="Table/Query"
                    RowSource ="tlu_Cover_Code"
                    StatusBarText ="Stream cover percentage quadrat 2"

                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =95
                    IMESentenceMode =3
                    ListRows =21
                    Left =4090
                    Top =3840
                    Width =778
                    TabIndex =41
                    ColumnInfo ="\"\";\"\";\"6\";\"4\""
                    Name ="Stream_Q3"
                    ControlSource ="Stream_Q3"
                    RowSourceType ="Table/Query"
                    RowSource ="tlu_Cover_Code"
                    StatusBarText ="Stream cover percentage quadrat 3"

                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =95
                    IMESentenceMode =3
                    ListRows =21
                    Left =2520
                    Top =4080
                    Width =778
                    TabIndex =42
                    ColumnInfo ="\"\";\"\";\"6\";\"4\""
                    Name ="Trash_Junk_Q1"
                    ControlSource ="Trash_Junk_Q1"
                    RowSourceType ="Table/Query"
                    RowSource ="tlu_Cover_Code"
                    StatusBarText ="Trash/Junk cover percentage quadrat 1"

                    Begin
                        Begin Label
                            OldBorderStyle =1
                            OverlapFlags =95
                            TextAlign =0
                            Left =660
                            Top =4080
                            Width =1860
                            Height =240
                            Name ="Label328"
                            Caption ="Trash/Junk"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =95
                    IMESentenceMode =3
                    ListRows =21
                    Left =3310
                    Top =4080
                    Width =778
                    TabIndex =43
                    ColumnInfo ="\"\";\"\";\"6\";\"4\""
                    Name ="Trash_Junk_Q2"
                    ControlSource ="Trash_Junk_Q2"
                    RowSourceType ="Table/Query"
                    RowSource ="tlu_Cover_Code"
                    StatusBarText ="Trash/Junk cover percentage quadrat 2"

                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =87
                    IMESentenceMode =3
                    ListRows =21
                    Left =4090
                    Top =4080
                    Width =778
                    TabIndex =44
                    ColumnInfo ="\"\";\"\";\"6\";\"4\""
                    Name ="Trash_Junk_Q3"
                    ControlSource ="Trash_Junk_Q3"
                    RowSourceType ="Table/Query"
                    RowSource ="tlu_Cover_Code"
                    StatusBarText ="Trash/Junk cover percentage quadrat 3"

                End
                Begin Subform
                    Visible = NotDefault
                    OverlapFlags =255
                    Left =5700
                    Top =480
                    Width =7260
                    Height =6720
                    TabIndex =46
                    Name ="fsub_Species_2009"
                    SourceObject ="Form.fsub_Species_2009"
                    LinkChildFields ="Quadrat_ID"
                    LinkMasterFields ="Quadrat_ID"

                End
                Begin Subform
                    OverlapFlags =247
                    Left =5700
                    Top =480
                    Width =7260
                    Height =6719
                    TabIndex =47
                    Name ="fsub_Species_2008"
                    SourceObject ="Form.fsub_Species_2008"
                    LinkChildFields ="Quadrat_ID"
                    LinkMasterFields ="Quadrat_ID"

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

Private Sub Form_BeforeInsert(Cancel As Integer)
    On Error GoTo Err_Handler
    If IsNull(Me.Parent!Transect_ID) Then
      MsgBox "You must enter Start Time first."
      DoCmd.CancelEvent
      SendKeys "{ESC}"
      GoTo Exit_Procedure
    End If
    ' Create the GUID primary key value
    If IsNull(Me!Quadrat_ID) Then
        If GetDataType("tbl_Quadrat", "Quadrat_ID") = dbText Then
            Me.Quadrat_ID = fxnGUIDGen
        End If
    End If

Exit_Procedure:
    Exit Sub

Err_Handler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    Resume Exit_Procedure
End Sub

Private Sub Form_Current()
  If Me.Parent.Parent!Start_Date < #1/1/2009# Then
    Me!fsub_Species.Visible = False
    Me!fsub_Species_2008.Visible = True
    Me!fsub_Species_2009.Visible = False
  ElseIf Me.Parent.Parent!Start_Date < #1/1/2010# Then
    Me!fsub_Species.Visible = False
    Me!fsub_Species_2009.Visible = True
    Me!fsub_Species_2008.Visible = False
  Else
    Me!fsub_Species.Visible = True
    Me!fsub_Species_2008.Visible = False
    Me!fsub_Species_2009.Visible = False
  End If
End Sub

Private Sub Observer_KeyDown(KeyCode As Integer, Shift As Integer)
  ' Ignore Page Down and Page Up keys for they will cycle through records
  Select Case KeyCode
    Case 33, 34
      KeyCode = 0
    End Select
End Sub
