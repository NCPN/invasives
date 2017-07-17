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
    Width =13380
    DatasheetFontHeight =9
    ItemSuffix =93
    Left =720
    Top =3585
    Right =14145
    Bottom =12240
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0xc5265af90df6e440
    End
    RecordSource ="usys_temp_transect"
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
            Height =8880
            BackColor =26112
            Name ="Detail"
            AlternateBackColor =26112
            Begin
                Begin OptionGroup
                    BackStyle =1
                    OverlapFlags =93
                    Left =120
                    Top =240
                    Width =13140
                    Height =8520
                    TabIndex =80
                    BackColor =14478320
                    Name ="Frame89"

                    LayoutCachedLeft =120
                    LayoutCachedTop =240
                    LayoutCachedWidth =13260
                    LayoutCachedHeight =8760
                    Begin
                        Begin Label
                            BackStyle =1
                            OverlapFlags =215
                            TextAlign =2
                            Left =225
                            Top =60
                            Width =1755
                            Height =285
                            FontSize =10
                            FontWeight =600
                            BackColor =26112
                            ForeColor =16777181
                            Name ="lblTransectFrame"
                            Caption ="Transect Data"
                            LayoutCachedLeft =225
                            LayoutCachedTop =60
                            LayoutCachedWidth =1980
                            LayoutCachedHeight =345
                        End
                    End
                End
                Begin TextBox
                    SpecialEffect =0
                    OldBorderStyle =0
                    OverlapFlags =223
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =2280
                    Top =8295
                    Width =840
                    Height =255
                    TabIndex =77
                    ForeColor =8355711
                    Name ="tbxIsSampledSum"
                    ControlSource ="=[IsSampled_Q1]+[IsSampled_Q2]+[IsSampled_Q3]"

                    LayoutCachedLeft =2280
                    LayoutCachedTop =8295
                    LayoutCachedWidth =3120
                    LayoutCachedHeight =8550
                    ForeThemeColorIndex =1
                    ForeShade =50.0
                End
                Begin TextBox
                    SpecialEffect =0
                    OldBorderStyle =0
                    OverlapFlags =223
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =3420
                    Top =8295
                    Width =840
                    Height =255
                    TabIndex =78
                    ForeColor =8355711
                    Name ="tbxNoExoticsSum"
                    ControlSource ="=[NoExotics_Q1]+[NoExotics_Q2]+[NoExotics_Q3]"

                    LayoutCachedLeft =3420
                    LayoutCachedTop =8295
                    LayoutCachedWidth =4260
                    LayoutCachedHeight =8550
                    ForeThemeColorIndex =1
                    ForeShade =50.0
                End
                Begin Subform
                    Visible = NotDefault
                    OverlapFlags =223
                    Left =5955
                    Top =1500
                    Width =7200
                    Height =6717
                    TabIndex =55
                    Name ="fsub_Species_2008"
                    SourceObject ="Form.fsub_Species_2008"
                    LinkChildFields ="Transect_ID"
                    LinkMasterFields ="Transect_ID"

                    LayoutCachedLeft =5955
                    LayoutCachedTop =1500
                    LayoutCachedWidth =13155
                    LayoutCachedHeight =8217
                End
                Begin Subform
                    Visible = NotDefault
                    OverlapFlags =255
                    Left =5715
                    Top =1500
                    Width =7320
                    Height =6598
                    TabIndex =56
                    Name ="fsub_Species_2009"
                    SourceObject ="Form.fsub_Species_2009"
                    LinkChildFields ="Transect_ID"
                    LinkMasterFields ="Transect_ID"

                    LayoutCachedLeft =5715
                    LayoutCachedTop =1500
                    LayoutCachedWidth =13035
                    LayoutCachedHeight =8098
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =223
                    IMESentenceMode =3
                    Left =12450
                    Top =720
                    Width =630
                    Height =180
                    ColumnWidth =2310
                    Name ="Transect_ID"
                    ControlSource ="Transect_ID"
                    StatusBarText ="Unique record identifier - primary key"

                    LayoutCachedLeft =12450
                    LayoutCachedTop =720
                    LayoutCachedWidth =13080
                    LayoutCachedHeight =900
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =223
                    IMESentenceMode =3
                    Left =12450
                    Top =1020
                    Width =630
                    Height =180
                    ColumnWidth =2310
                    TabIndex =1
                    Name ="Event_ID"
                    ControlSource ="Event_ID"
                    StatusBarText ="M. Link to tbl_Locations (Loc_ID)"

                    LayoutCachedLeft =12450
                    LayoutCachedTop =1020
                    LayoutCachedWidth =13080
                    LayoutCachedHeight =1200
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    SpecialEffect =0
                    OldBorderStyle =0
                    OverlapFlags =215
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =690
                    Top =420
                    Width =1080
                    ColumnWidth =465
                    FontWeight =700
                    TabIndex =2
                    Name ="Transect"
                    ControlSource ="Transect"
                    StatusBarText ="Transect number - 1, 2, or 3"

                    LayoutCachedLeft =690
                    LayoutCachedTop =420
                    LayoutCachedWidth =1770
                    LayoutCachedHeight =660
                    Begin
                        Begin Label
                            OverlapFlags =223
                            TextAlign =1
                            Left =330
                            Top =420
                            Width =360
                            Height =240
                            FontWeight =700
                            Name ="lblTransectNumber"
                            Caption ="#"
                            LayoutCachedLeft =330
                            LayoutCachedTop =420
                            LayoutCachedWidth =690
                            LayoutCachedHeight =660
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =215
                    IMESentenceMode =3
                    Left =4920
                    Top =420
                    Width =960
                    ColumnWidth =1035
                    TabIndex =3
                    Name ="tbxStartTime"
                    ControlSource ="Start_Time"
                    Format ="Short Time"
                    StatusBarText ="Date of visit."
                    AfterUpdate ="=UpdateTransect([Screen].[ActiveControl])"
                    InputMask ="00:00;0;_"
                    OnKeyDown ="[Event Procedure]"

                    LayoutCachedLeft =4920
                    LayoutCachedTop =420
                    LayoutCachedWidth =5880
                    LayoutCachedHeight =660
                    Begin
                        Begin Label
                            OverlapFlags =215
                            Left =3060
                            Top =420
                            Width =1770
                            Height =240
                            FontWeight =700
                            Name ="lblStartTimeTransect"
                            Caption ="Transect Start Time"
                            LayoutCachedLeft =3060
                            LayoutCachedTop =420
                            LayoutCachedWidth =4830
                            LayoutCachedHeight =660
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =215
                    Left =1890
                    Top =420
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

                    LayoutCachedLeft =1890
                    LayoutCachedTop =420
                    LayoutCachedWidth =2196
                    LayoutCachedHeight =726
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =215
                    Left =2250
                    Top =420
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

                    LayoutCachedLeft =2250
                    LayoutCachedTop =420
                    LayoutCachedWidth =2556
                    LayoutCachedHeight =726
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =223
                    IMESentenceMode =3
                    Left =12240
                    Top =420
                    Width =840
                    Height =180
                    TabIndex =6
                    Name ="GPS_Time"
                    ControlSource ="GPS_Time"
                    Format ="Long Time"
                    StatusBarText ="Recording time"

                    LayoutCachedLeft =12240
                    LayoutCachedTop =420
                    LayoutCachedWidth =13080
                    LayoutCachedHeight =600
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =247
                    IMESentenceMode =3
                    Left =7290
                    Top =420
                    Width =5460
                    Height =903
                    TabIndex =7
                    Name ="tbxComments"
                    ControlSource ="Comments"
                    StatusBarText ="Notes"
                    AfterUpdate ="=UpdateTransect([Screen].[ActiveControl])"

                    LayoutCachedLeft =7290
                    LayoutCachedTop =420
                    LayoutCachedWidth =12750
                    LayoutCachedHeight =1323
                    Begin
                        Begin Label
                            OverlapFlags =215
                            Left =6270
                            Top =420
                            Width =915
                            Height =240
                            FontWeight =700
                            Name ="lblComments"
                            Caption ="Comments:"
                            LayoutCachedLeft =6270
                            LayoutCachedTop =420
                            LayoutCachedWidth =7185
                            LayoutCachedHeight =660
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =215
                    IMESentenceMode =3
                    ColumnCount =3
                    ListWidth =1650
                    Left =1260
                    Top =810
                    Width =1620
                    TabIndex =8
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"\";\"\";\"10\";\"100\""
                    Name ="cbxObserver"
                    ControlSource ="Observer"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT tlu_Contacts.Contact_ID, tlu_Contacts.Last_Name, tlu_Contacts.First_Name "
                        "FROM tlu_Contacts;"
                    ColumnWidths ="0;810;839"
                    AfterUpdate ="=UpdateTransect([Screen].[ActiveControl])"
                    LayoutCachedLeft =1260
                    LayoutCachedTop =810
                    LayoutCachedWidth =2880
                    LayoutCachedHeight =1050
                    Begin
                        Begin Label
                            OverlapFlags =215
                            Left =300
                            Top =810
                            Width =855
                            Height =240
                            FontWeight =700
                            Name ="lblObserver"
                            Caption ="Observer"
                            LayoutCachedLeft =300
                            LayoutCachedTop =810
                            LayoutCachedWidth =1155
                            LayoutCachedHeight =1050
                        End
                    End
                End
                Begin Label
                    OldBorderStyle =1
                    BorderWidth =1
                    OverlapFlags =223
                    TextAlign =2
                    Left =255
                    Top =3240
                    Width =1860
                    Height =480
                    FontWeight =700
                    Name ="lblMicrohabitat"
                    Caption ="Microhabitat"
                    LayoutCachedLeft =255
                    LayoutCachedTop =3240
                    LayoutCachedWidth =2115
                    LayoutCachedHeight =3720
                End
                Begin Label
                    OldBorderStyle =1
                    BorderWidth =1
                    OverlapFlags =223
                    TextAlign =2
                    Left =2115
                    Top =3480
                    Width =778
                    Height =240
                    FontWeight =700
                    Name ="lblMicrohabitatCoverQ1"
                    Caption ="Q1"
                    LayoutCachedLeft =2115
                    LayoutCachedTop =3480
                    LayoutCachedWidth =2893
                    LayoutCachedHeight =3720
                End
                Begin Label
                    OldBorderStyle =1
                    BorderWidth =1
                    OverlapFlags =223
                    TextAlign =2
                    Left =2895
                    Top =3480
                    Width =780
                    Height =240
                    FontWeight =700
                    Name ="lblMicrohabitatCoverQ2"
                    Caption ="Q2"
                    LayoutCachedLeft =2895
                    LayoutCachedTop =3480
                    LayoutCachedWidth =3675
                    LayoutCachedHeight =3720
                End
                Begin Label
                    OldBorderStyle =1
                    BorderWidth =1
                    OverlapFlags =223
                    TextAlign =2
                    Left =3674
                    Top =3480
                    Width =780
                    Height =240
                    FontWeight =700
                    Name ="lblMicrohabitatCoverQ3"
                    Caption ="Q3"
                    LayoutCachedLeft =3674
                    LayoutCachedTop =3480
                    LayoutCachedWidth =4454
                    LayoutCachedHeight =3720
                End
                Begin Label
                    OldBorderStyle =1
                    BorderWidth =1
                    OverlapFlags =215
                    TextAlign =2
                    Left =2115
                    Top =3240
                    Width =2340
                    Height =240
                    FontWeight =700
                    Name ="lblPctCover"
                    Caption ="% cover"
                    LayoutCachedLeft =2115
                    LayoutCachedTop =3240
                    LayoutCachedWidth =4455
                    LayoutCachedHeight =3480
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =223
                    IMESentenceMode =3
                    ListRows =21
                    Left =2115
                    Top =3720
                    Width =778
                    Height =239
                    TabIndex =9
                    ColumnInfo ="\"\";\"\";\"6\";\"4\""
                    Name ="Dead_Wood_Q1"
                    RowSourceType ="Table/Query"
                    RowSource ="tlu_Cover_Code"
                    StatusBarText ="Dead wood cover percentage quadrat 1"
                    OnChange ="=UpdateMicrohabitat([Screen].[ActiveControl])"

                    LayoutCachedLeft =2115
                    LayoutCachedTop =3720
                    LayoutCachedWidth =2893
                    LayoutCachedHeight =3959
                    Begin
                        Begin Label
                            OldBorderStyle =1
                            OverlapFlags =223
                            Left =255
                            Top =3720
                            Width =1860
                            Height =240
                            Name ="lblDeadWood"
                            Caption ="Dead Wood:"
                            LayoutCachedLeft =255
                            LayoutCachedTop =3720
                            LayoutCachedWidth =2115
                            LayoutCachedHeight =3960
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =223
                    IMESentenceMode =3
                    ListRows =21
                    Left =2905
                    Top =3720
                    Width =778
                    Height =239
                    TabIndex =10
                    ColumnInfo ="\"\";\"\";\"6\";\"4\""
                    Name ="Dead_Wood_Q2"
                    RowSourceType ="Table/Query"
                    RowSource ="tlu_Cover_Code"
                    StatusBarText ="Dead wood cover percentage quadrat 2"
                    OnChange ="=UpdateMicrohabitat([Screen].[ActiveControl])"

                    LayoutCachedLeft =2905
                    LayoutCachedTop =3720
                    LayoutCachedWidth =3683
                    LayoutCachedHeight =3959
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =255
                    IMESentenceMode =3
                    ListRows =21
                    Left =3675
                    Top =3720
                    Width =763
                    Height =239
                    TabIndex =11
                    ColumnInfo ="\"\";\"\";\"6\";\"4\""
                    Name ="Dead_Wood_Q3"
                    RowSourceType ="Table/Query"
                    RowSource ="tlu_Cover_Code"
                    StatusBarText ="Dead wood cover percentage quadrat 3"
                    OnChange ="=UpdateMicrohabitat([Screen].[ActiveControl])"

                    LayoutCachedLeft =3675
                    LayoutCachedTop =3720
                    LayoutCachedWidth =4438
                    LayoutCachedHeight =3959
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =223
                    IMESentenceMode =3
                    ListRows =21
                    Left =2115
                    Top =3960
                    Width =778
                    Height =239
                    TabIndex =12
                    ColumnInfo ="\"\";\"\";\"6\";\"4\""
                    Name ="Dung_Q1"
                    RowSourceType ="Table/Query"
                    RowSource ="tlu_Cover_Code"
                    StatusBarText ="Dung cover percentage quadrat 1"
                    OnChange ="=UpdateMicrohabitat([Screen].[ActiveControl])"

                    LayoutCachedLeft =2115
                    LayoutCachedTop =3960
                    LayoutCachedWidth =2893
                    LayoutCachedHeight =4199
                    Begin
                        Begin Label
                            OldBorderStyle =1
                            OverlapFlags =223
                            Left =255
                            Top =3960
                            Width =1860
                            Height =240
                            Name ="lblDung"
                            Caption ="Dung"
                            LayoutCachedLeft =255
                            LayoutCachedTop =3960
                            LayoutCachedWidth =2115
                            LayoutCachedHeight =4200
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =255
                    IMESentenceMode =3
                    ListRows =21
                    Left =2905
                    Top =3960
                    Width =778
                    Height =239
                    TabIndex =13
                    ColumnInfo ="\"\";\"\";\"6\";\"4\""
                    Name ="Dung_Q2"
                    RowSourceType ="Table/Query"
                    RowSource ="tlu_Cover_Code"
                    StatusBarText ="Dung cover percentage quadrat 2"
                    OnChange ="=UpdateMicrohabitat([Screen].[ActiveControl])"

                    LayoutCachedLeft =2905
                    LayoutCachedTop =3960
                    LayoutCachedWidth =3683
                    LayoutCachedHeight =4199
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =255
                    IMESentenceMode =3
                    ListRows =21
                    Left =3675
                    Top =3959
                    Width =763
                    Height =239
                    TabIndex =14
                    ColumnInfo ="\"\";\"\";\"6\";\"4\""
                    Name ="Dung_Q3"
                    RowSourceType ="Table/Query"
                    RowSource ="tlu_Cover_Code"
                    StatusBarText ="Dung cover percentage quadrat 3"
                    OnChange ="=UpdateMicrohabitat([Screen].[ActiveControl])"

                    LayoutCachedLeft =3675
                    LayoutCachedTop =3959
                    LayoutCachedWidth =4438
                    LayoutCachedHeight =4198
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =255
                    IMESentenceMode =3
                    ListRows =21
                    Left =2115
                    Top =4200
                    Width =778
                    Height =239
                    TabIndex =15
                    ColumnInfo ="\"\";\"\";\"6\";\"4\""
                    Name ="Fungus_Q1"
                    RowSourceType ="Table/Query"
                    RowSource ="tlu_Cover_Code"
                    StatusBarText ="Fungus cover percentage quadrat 1"
                    OnChange ="=UpdateMicrohabitat([Screen].[ActiveControl])"

                    LayoutCachedLeft =2115
                    LayoutCachedTop =4200
                    LayoutCachedWidth =2893
                    LayoutCachedHeight =4439
                    Begin
                        Begin Label
                            OldBorderStyle =1
                            OverlapFlags =223
                            Left =255
                            Top =4200
                            Width =1860
                            Height =240
                            Name ="lblFungus"
                            Caption ="Fungus"
                            LayoutCachedLeft =255
                            LayoutCachedTop =4200
                            LayoutCachedWidth =2115
                            LayoutCachedHeight =4440
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =255
                    IMESentenceMode =3
                    ListRows =21
                    Left =2905
                    Top =4200
                    Width =778
                    Height =239
                    TabIndex =16
                    ColumnInfo ="\"\";\"\";\"6\";\"4\""
                    Name ="Fungus_Q2"
                    RowSourceType ="Table/Query"
                    RowSource ="tlu_Cover_Code"
                    StatusBarText ="Fungus cover percentage quadrat 2"
                    OnChange ="=UpdateMicrohabitat([Screen].[ActiveControl])"

                    LayoutCachedLeft =2905
                    LayoutCachedTop =4200
                    LayoutCachedWidth =3683
                    LayoutCachedHeight =4439
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =255
                    IMESentenceMode =3
                    ListRows =21
                    Left =3675
                    Top =4198
                    Width =763
                    Height =239
                    TabIndex =17
                    ColumnInfo ="\"\";\"\";\"6\";\"4\""
                    Name ="Fungus_Q3"
                    RowSourceType ="Table/Query"
                    RowSource ="tlu_Cover_Code"
                    StatusBarText ="Fungus cover percentage quadrat 3"
                    OnChange ="=UpdateMicrohabitat([Screen].[ActiveControl])"

                    LayoutCachedLeft =3675
                    LayoutCachedTop =4198
                    LayoutCachedWidth =4438
                    LayoutCachedHeight =4437
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =255
                    IMESentenceMode =3
                    ListRows =21
                    Left =2115
                    Top =4440
                    Width =778
                    Height =239
                    TabIndex =18
                    ColumnInfo ="\"\";\"\";\"6\";\"4\""
                    Name ="Lichen_Q1"
                    RowSourceType ="Table/Query"
                    RowSource ="tlu_Cover_Code"
                    StatusBarText ="Lichen cover percentage quadrat 1"
                    OnChange ="=UpdateMicrohabitat([Screen].[ActiveControl])"

                    LayoutCachedLeft =2115
                    LayoutCachedTop =4440
                    LayoutCachedWidth =2893
                    LayoutCachedHeight =4679
                    Begin
                        Begin Label
                            OldBorderStyle =1
                            OverlapFlags =255
                            Left =255
                            Top =4440
                            Width =1860
                            Height =240
                            Name ="lblLichen"
                            Caption ="Lichen"
                            LayoutCachedLeft =255
                            LayoutCachedTop =4440
                            LayoutCachedWidth =2115
                            LayoutCachedHeight =4680
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =255
                    IMESentenceMode =3
                    ListRows =21
                    Left =2905
                    Top =4440
                    Width =778
                    Height =239
                    TabIndex =19
                    ColumnInfo ="\"\";\"\";\"6\";\"4\""
                    Name ="Lichen_Q2"
                    RowSourceType ="Table/Query"
                    RowSource ="tlu_Cover_Code"
                    StatusBarText ="Lichen cover percentage quadrat 2"
                    OnChange ="=UpdateMicrohabitat([Screen].[ActiveControl])"

                    LayoutCachedLeft =2905
                    LayoutCachedTop =4440
                    LayoutCachedWidth =3683
                    LayoutCachedHeight =4679
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =255
                    IMESentenceMode =3
                    ListRows =21
                    Left =3675
                    Top =4437
                    Width =763
                    Height =239
                    TabIndex =20
                    ColumnInfo ="\"\";\"\";\"6\";\"4\""
                    Name ="Lichen_Q3"
                    RowSourceType ="Table/Query"
                    RowSource ="tlu_Cover_Code"
                    StatusBarText ="Lichen cover percentage quadrat 3"
                    OnChange ="=UpdateMicrohabitat([Screen].[ActiveControl])"

                    LayoutCachedLeft =3675
                    LayoutCachedTop =4437
                    LayoutCachedWidth =4438
                    LayoutCachedHeight =4676
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =255
                    IMESentenceMode =3
                    ListRows =21
                    Left =2115
                    Top =4680
                    Width =778
                    Height =239
                    TabIndex =21
                    ColumnInfo ="\"\";\"\";\"6\";\"4\""
                    Name ="Litter_Duff_Q1"
                    RowSourceType ="Table/Query"
                    RowSource ="tlu_Cover_Code"
                    StatusBarText ="Litter/Duff cover percentage quadrat 1"
                    OnChange ="=UpdateMicrohabitat([Screen].[ActiveControl])"

                    LayoutCachedLeft =2115
                    LayoutCachedTop =4680
                    LayoutCachedWidth =2893
                    LayoutCachedHeight =4919
                    Begin
                        Begin Label
                            OldBorderStyle =1
                            OverlapFlags =255
                            Left =255
                            Top =4680
                            Width =1860
                            Height =240
                            Name ="lblLitterDuff"
                            Caption ="Litter Duff"
                            LayoutCachedLeft =255
                            LayoutCachedTop =4680
                            LayoutCachedWidth =2115
                            LayoutCachedHeight =4920
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =255
                    IMESentenceMode =3
                    ListRows =21
                    Left =2905
                    Top =4680
                    Width =778
                    Height =239
                    TabIndex =22
                    ColumnInfo ="\"\";\"\";\"6\";\"4\""
                    Name ="Litter_Duff_Q2"
                    RowSourceType ="Table/Query"
                    RowSource ="tlu_Cover_Code"
                    StatusBarText ="Litter/Duff cover percentage quadrat 2"
                    OnChange ="=UpdateMicrohabitat([Screen].[ActiveControl])"

                    LayoutCachedLeft =2905
                    LayoutCachedTop =4680
                    LayoutCachedWidth =3683
                    LayoutCachedHeight =4919
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =255
                    IMESentenceMode =3
                    ListRows =21
                    Left =3675
                    Top =4676
                    Width =763
                    Height =239
                    TabIndex =23
                    ColumnInfo ="\"\";\"\";\"6\";\"4\""
                    Name ="Litter_Duff_Q3"
                    RowSourceType ="Table/Query"
                    RowSource ="tlu_Cover_Code"
                    StatusBarText ="Litter/Duff cover percentage quadrat 3"
                    OnChange ="=UpdateMicrohabitat([Screen].[ActiveControl])"

                    LayoutCachedLeft =3675
                    LayoutCachedTop =4676
                    LayoutCachedWidth =4438
                    LayoutCachedHeight =4915
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =255
                    IMESentenceMode =3
                    ListRows =21
                    Left =2115
                    Top =4920
                    Width =778
                    Height =239
                    TabIndex =24
                    ColumnInfo ="\"\";\"\";\"6\";\"4\""
                    Name ="Root_Bole_Q1"
                    RowSourceType ="Table/Query"
                    RowSource ="tlu_Cover_Code"
                    StatusBarText ="Root/Bole cover percentage quadrat 1"
                    OnChange ="=UpdateMicrohabitat([Screen].[ActiveControl])"

                    LayoutCachedLeft =2115
                    LayoutCachedTop =4920
                    LayoutCachedWidth =2893
                    LayoutCachedHeight =5159
                    Begin
                        Begin Label
                            OldBorderStyle =1
                            OverlapFlags =255
                            Left =255
                            Top =4920
                            Width =1860
                            Height =240
                            Name ="lblLiveRootBole"
                            Caption ="Live root/Bole"
                            LayoutCachedLeft =255
                            LayoutCachedTop =4920
                            LayoutCachedWidth =2115
                            LayoutCachedHeight =5160
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =255
                    IMESentenceMode =3
                    ListRows =21
                    Left =2905
                    Top =4920
                    Width =778
                    Height =239
                    TabIndex =25
                    ColumnInfo ="\"\";\"\";\"6\";\"4\""
                    Name ="Root_Bole_Q2"
                    RowSourceType ="Table/Query"
                    RowSource ="tlu_Cover_Code"
                    StatusBarText ="Roo/Bole cover percentage quadrat 2"
                    OnChange ="=UpdateMicrohabitat([Screen].[ActiveControl])"

                    LayoutCachedLeft =2905
                    LayoutCachedTop =4920
                    LayoutCachedWidth =3683
                    LayoutCachedHeight =5159
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =255
                    IMESentenceMode =3
                    ListRows =21
                    Left =3675
                    Top =4915
                    Width =763
                    Height =239
                    TabIndex =26
                    ColumnInfo ="\"\";\"\";\"6\";\"4\""
                    Name ="Root_Bole_Q3"
                    RowSourceType ="Table/Query"
                    RowSource ="tlu_Cover_Code"
                    StatusBarText ="Root/Bole cover percentage quadrat 3"
                    OnChange ="=UpdateMicrohabitat([Screen].[ActiveControl])"

                    LayoutCachedLeft =3675
                    LayoutCachedTop =4915
                    LayoutCachedWidth =4438
                    LayoutCachedHeight =5154
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =223
                    IMESentenceMode =3
                    ListRows =21
                    Left =2115
                    Top =5401
                    Width =778
                    Height =239
                    TabIndex =30
                    ColumnInfo ="\"\";\"\";\"6\";\"4\""
                    Name ="Mineral_Soil_Sediment_Q1"
                    RowSourceType ="Table/Query"
                    RowSource ="tlu_Cover_Code"
                    StatusBarText ="Mineral Soil/Sediment cover percentage quadrat 1"
                    OnChange ="=UpdateMicrohabitat([Screen].[ActiveControl])"

                    LayoutCachedLeft =2115
                    LayoutCachedTop =5401
                    LayoutCachedWidth =2893
                    LayoutCachedHeight =5640
                    Begin
                        Begin Label
                            OldBorderStyle =1
                            OverlapFlags =223
                            Left =255
                            Top =5400
                            Width =1858
                            Height =240
                            Name ="lblMineralSoilSediment"
                            Caption ="Mineral Soil/Sediment"
                            LayoutCachedLeft =255
                            LayoutCachedTop =5400
                            LayoutCachedWidth =2113
                            LayoutCachedHeight =5640
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =223
                    IMESentenceMode =3
                    ListRows =21
                    Left =2905
                    Top =5401
                    Width =778
                    Height =239
                    TabIndex =31
                    ColumnInfo ="\"\";\"\";\"6\";\"4\""
                    Name ="Mineral_Soil_Sediment_Q2"
                    RowSourceType ="Table/Query"
                    RowSource ="tlu_Cover_Code"
                    StatusBarText ="Mineral Soil/Sediment cover percentage quadrat 2"
                    OnChange ="=UpdateMicrohabitat([Screen].[ActiveControl])"

                    LayoutCachedLeft =2905
                    LayoutCachedTop =5401
                    LayoutCachedWidth =3683
                    LayoutCachedHeight =5640
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =255
                    IMESentenceMode =3
                    ListRows =21
                    Left =3675
                    Top =5400
                    Width =763
                    Height =239
                    TabIndex =32
                    ColumnInfo ="\"\";\"\";\"6\";\"4\""
                    Name ="Mineral_Soil_Sediment_Q3"
                    RowSourceType ="Table/Query"
                    RowSource ="tlu_Cover_Code"
                    StatusBarText ="Mineral Soil/Sediment cover percentage quadrat 3"
                    OnChange ="=UpdateMicrohabitat([Screen].[ActiveControl])"

                    LayoutCachedLeft =3675
                    LayoutCachedTop =5400
                    LayoutCachedWidth =4438
                    LayoutCachedHeight =5639
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =223
                    IMESentenceMode =3
                    ListRows =21
                    Left =2115
                    Top =5641
                    Width =778
                    Height =239
                    TabIndex =33
                    ColumnInfo ="\"\";\"\";\"6\";\"4\""
                    Name ="Moss_Q1"
                    RowSourceType ="Table/Query"
                    RowSource ="tlu_Cover_Code"
                    StatusBarText ="Moss cover percentage quadrat 1"
                    OnChange ="=UpdateMicrohabitat([Screen].[ActiveControl])"

                    LayoutCachedLeft =2115
                    LayoutCachedTop =5641
                    LayoutCachedWidth =2893
                    LayoutCachedHeight =5880
                    Begin
                        Begin Label
                            OldBorderStyle =1
                            OverlapFlags =223
                            Left =255
                            Top =5640
                            Width =1860
                            Height =240
                            Name ="lblMoss"
                            Caption ="Moss"
                            LayoutCachedLeft =255
                            LayoutCachedTop =5640
                            LayoutCachedWidth =2115
                            LayoutCachedHeight =5880
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =255
                    IMESentenceMode =3
                    ListRows =21
                    Left =2905
                    Top =5641
                    Width =778
                    Height =239
                    TabIndex =34
                    ColumnInfo ="\"\";\"\";\"6\";\"4\""
                    Name ="Moss_Q2"
                    RowSourceType ="Table/Query"
                    RowSource ="tlu_Cover_Code"
                    StatusBarText ="Moss cover percentage quadrat 2"
                    OnChange ="=UpdateMicrohabitat([Screen].[ActiveControl])"

                    LayoutCachedLeft =2905
                    LayoutCachedTop =5641
                    LayoutCachedWidth =3683
                    LayoutCachedHeight =5880
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =255
                    IMESentenceMode =3
                    ListRows =21
                    Left =3675
                    Top =5641
                    Width =763
                    Height =239
                    TabIndex =35
                    ColumnInfo ="\"\";\"\";\"6\";\"4\""
                    Name ="Moss_Q3"
                    RowSourceType ="Table/Query"
                    RowSource ="tlu_Cover_Code"
                    StatusBarText ="Moss cover percentage quadrat 3"
                    OnChange ="=UpdateMicrohabitat([Screen].[ActiveControl])"

                    LayoutCachedLeft =3675
                    LayoutCachedTop =5641
                    LayoutCachedWidth =4438
                    LayoutCachedHeight =5880
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =255
                    IMESentenceMode =3
                    ListRows =21
                    Left =2115
                    Top =5881
                    Width =778
                    Height =239
                    TabIndex =36
                    ColumnInfo ="\"\";\"\";\"6\";\"4\""
                    Name ="Cryptogram_Q1"
                    RowSourceType ="Table/Query"
                    RowSource ="tlu_Cover_Code"
                    StatusBarText ="Cryptogram cover percentage quadrat 1"
                    OnChange ="=UpdateMicrohabitat([Screen].[ActiveControl])"

                    LayoutCachedLeft =2115
                    LayoutCachedTop =5881
                    LayoutCachedWidth =2893
                    LayoutCachedHeight =6120
                    Begin
                        Begin Label
                            OldBorderStyle =1
                            OverlapFlags =223
                            Left =255
                            Top =5880
                            Width =1860
                            Height =240
                            Name ="lblBiologicalSoilCrust"
                            Caption ="Biological Soil Crust"
                            LayoutCachedLeft =255
                            LayoutCachedTop =5880
                            LayoutCachedWidth =2115
                            LayoutCachedHeight =6120
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =255
                    IMESentenceMode =3
                    ListRows =21
                    Left =2905
                    Top =5881
                    Width =778
                    Height =239
                    TabIndex =37
                    ColumnInfo ="\"\";\"\";\"6\";\"4\""
                    Name ="Cryptogram_Q2"
                    RowSourceType ="Table/Query"
                    RowSource ="tlu_Cover_Code"
                    StatusBarText ="Cryptogram cover percentage quadrat 2"
                    OnChange ="=UpdateMicrohabitat([Screen].[ActiveControl])"

                    LayoutCachedLeft =2905
                    LayoutCachedTop =5881
                    LayoutCachedWidth =3683
                    LayoutCachedHeight =6120
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =255
                    IMESentenceMode =3
                    ListRows =21
                    Left =3675
                    Top =5881
                    Width =763
                    Height =239
                    TabIndex =38
                    ColumnInfo ="\"\";\"\";\"6\";\"4\""
                    Name ="Cryptogram_Q3"
                    RowSourceType ="Table/Query"
                    RowSource ="tlu_Cover_Code"
                    StatusBarText ="Cryptogram cover percentage quadrat 3"
                    OnChange ="=UpdateMicrohabitat([Screen].[ActiveControl])"

                    LayoutCachedLeft =3675
                    LayoutCachedTop =5881
                    LayoutCachedWidth =4438
                    LayoutCachedHeight =6120
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =255
                    IMESentenceMode =3
                    ListRows =21
                    Left =2115
                    Top =6121
                    Width =778
                    Height =239
                    TabIndex =39
                    ColumnInfo ="\"\";\"\";\"6\";\"4\""
                    Name ="Road_Q1"
                    RowSourceType ="Table/Query"
                    RowSource ="tlu_Cover_Code"
                    StatusBarText ="Road cover percentage quadrat 1"
                    OnChange ="=UpdateMicrohabitat([Screen].[ActiveControl])"

                    LayoutCachedLeft =2115
                    LayoutCachedTop =6121
                    LayoutCachedWidth =2893
                    LayoutCachedHeight =6360
                    Begin
                        Begin Label
                            OldBorderStyle =1
                            OverlapFlags =255
                            Left =255
                            Top =6120
                            Width =1860
                            Height =240
                            Name ="lblRoad"
                            Caption ="Road"
                            LayoutCachedLeft =255
                            LayoutCachedTop =6120
                            LayoutCachedWidth =2115
                            LayoutCachedHeight =6360
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =255
                    IMESentenceMode =3
                    ListRows =21
                    Left =2905
                    Top =6121
                    Width =778
                    Height =239
                    TabIndex =40
                    ColumnInfo ="\"\";\"\";\"6\";\"4\""
                    Name ="Road_Q2"
                    RowSourceType ="Table/Query"
                    RowSource ="tlu_Cover_Code"
                    StatusBarText ="Road cover percentage quadrat 2"
                    OnChange ="=UpdateMicrohabitat([Screen].[ActiveControl])"

                    LayoutCachedLeft =2905
                    LayoutCachedTop =6121
                    LayoutCachedWidth =3683
                    LayoutCachedHeight =6360
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =255
                    IMESentenceMode =3
                    ListRows =21
                    Left =3675
                    Top =6121
                    Width =763
                    Height =239
                    TabIndex =41
                    ColumnInfo ="\"\";\"\";\"6\";\"4\""
                    Name ="Road_Q3"
                    RowSourceType ="Table/Query"
                    RowSource ="tlu_Cover_Code"
                    StatusBarText ="Road cover percentage quadrat 3"
                    OnChange ="=UpdateMicrohabitat([Screen].[ActiveControl])"

                    LayoutCachedLeft =3675
                    LayoutCachedTop =6121
                    LayoutCachedWidth =4438
                    LayoutCachedHeight =6360
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =255
                    IMESentenceMode =3
                    ListRows =21
                    Left =2115
                    Top =6361
                    Width =778
                    Height =239
                    TabIndex =42
                    ColumnInfo ="\"\";\"\";\"6\";\"4\""
                    Name ="Rock_Q1"
                    RowSourceType ="Table/Query"
                    RowSource ="tlu_Cover_Code"
                    StatusBarText ="Rock cover percentage quadrat 1"
                    OnChange ="=UpdateMicrohabitat([Screen].[ActiveControl])"

                    LayoutCachedLeft =2115
                    LayoutCachedTop =6361
                    LayoutCachedWidth =2893
                    LayoutCachedHeight =6600
                    Begin
                        Begin Label
                            OldBorderStyle =1
                            OverlapFlags =255
                            Left =255
                            Top =6360
                            Width =1860
                            Height =240
                            Name ="lblRock"
                            Caption ="Rock"
                            LayoutCachedLeft =255
                            LayoutCachedTop =6360
                            LayoutCachedWidth =2115
                            LayoutCachedHeight =6600
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =255
                    IMESentenceMode =3
                    ListRows =21
                    Left =2905
                    Top =6361
                    Width =778
                    Height =239
                    TabIndex =43
                    ColumnInfo ="\"\";\"\";\"6\";\"4\""
                    Name ="Rock_Q2"
                    RowSourceType ="Table/Query"
                    RowSource ="tlu_Cover_Code"
                    StatusBarText ="Rock cover percentage quadrat 2"
                    OnChange ="=UpdateMicrohabitat([Screen].[ActiveControl])"

                    LayoutCachedLeft =2905
                    LayoutCachedTop =6361
                    LayoutCachedWidth =3683
                    LayoutCachedHeight =6600
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =255
                    IMESentenceMode =3
                    ListRows =21
                    Left =3675
                    Top =6361
                    Width =763
                    Height =239
                    TabIndex =44
                    ColumnInfo ="\"\";\"\";\"6\";\"4\""
                    Name ="Rock_Q3"
                    RowSourceType ="Table/Query"
                    RowSource ="tlu_Cover_Code"
                    StatusBarText ="Rock cover percentage quadrat 3"
                    OnChange ="=UpdateMicrohabitat([Screen].[ActiveControl])"

                    LayoutCachedLeft =3675
                    LayoutCachedTop =6361
                    LayoutCachedWidth =4438
                    LayoutCachedHeight =6600
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =255
                    IMESentenceMode =3
                    ListRows =21
                    Left =2115
                    Top =6601
                    Width =778
                    Height =239
                    TabIndex =45
                    ColumnInfo ="\"\";\"\";\"6\";\"4\""
                    Name ="Standing_Water_Flooded_Q1"
                    RowSourceType ="Table/Query"
                    RowSource ="tlu_Cover_Code"
                    StatusBarText ="Standing Water/Flooded cover percentage quadrat 1"
                    OnChange ="=UpdateMicrohabitat([Screen].[ActiveControl])"

                    LayoutCachedLeft =2115
                    LayoutCachedTop =6601
                    LayoutCachedWidth =2893
                    LayoutCachedHeight =6840
                    Begin
                        Begin Label
                            OldBorderStyle =1
                            OverlapFlags =255
                            Left =255
                            Top =6600
                            Width =1860
                            Height =240
                            Name ="lblStandingWaterFlooded"
                            Caption ="Standing Water/Flooded"
                            LayoutCachedLeft =255
                            LayoutCachedTop =6600
                            LayoutCachedWidth =2115
                            LayoutCachedHeight =6840
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =255
                    IMESentenceMode =3
                    ListRows =21
                    Left =2905
                    Top =6601
                    Width =778
                    Height =239
                    TabIndex =46
                    ColumnInfo ="\"\";\"\";\"6\";\"4\""
                    Name ="Standing_Water_Flooded_Q2"
                    RowSourceType ="Table/Query"
                    RowSource ="tlu_Cover_Code"
                    StatusBarText ="Standing Water/Flooded cover percentage quadrat 2"
                    OnChange ="=UpdateMicrohabitat([Screen].[ActiveControl])"

                    LayoutCachedLeft =2905
                    LayoutCachedTop =6601
                    LayoutCachedWidth =3683
                    LayoutCachedHeight =6840
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =255
                    IMESentenceMode =3
                    ListRows =21
                    Left =3675
                    Top =6601
                    Width =763
                    Height =239
                    TabIndex =47
                    ColumnInfo ="\"\";\"\";\"6\";\"4\""
                    Name ="Standing_Water_Flooded_Q3"
                    RowSourceType ="Table/Query"
                    RowSource ="tlu_Cover_Code"
                    StatusBarText ="Standing Water/Flooded cover percentage quadrat 3"
                    OnChange ="=UpdateMicrohabitat([Screen].[ActiveControl])"

                    LayoutCachedLeft =3675
                    LayoutCachedTop =6601
                    LayoutCachedWidth =4438
                    LayoutCachedHeight =6840
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =255
                    IMESentenceMode =3
                    ListRows =21
                    Left =2115
                    Top =6841
                    Width =778
                    Height =239
                    TabIndex =48
                    ColumnInfo ="\"\";\"\";\"6\";\"4\""
                    Name ="Stream_Q1"
                    RowSourceType ="Table/Query"
                    RowSource ="tlu_Cover_Code"
                    StatusBarText ="Stream cover percentage quadrat 1"
                    OnChange ="=UpdateMicrohabitat([Screen].[ActiveControl])"

                    LayoutCachedLeft =2115
                    LayoutCachedTop =6841
                    LayoutCachedWidth =2893
                    LayoutCachedHeight =7080
                    Begin
                        Begin Label
                            OldBorderStyle =1
                            OverlapFlags =255
                            Left =255
                            Top =6840
                            Width =1860
                            Height =240
                            Name ="lblStream"
                            Caption ="Stream"
                            LayoutCachedLeft =255
                            LayoutCachedTop =6840
                            LayoutCachedWidth =2115
                            LayoutCachedHeight =7080
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =255
                    IMESentenceMode =3
                    ListRows =21
                    Left =2905
                    Top =6841
                    Width =778
                    Height =239
                    TabIndex =49
                    ColumnInfo ="\"\";\"\";\"6\";\"4\""
                    Name ="Stream_Q2"
                    RowSourceType ="Table/Query"
                    RowSource ="tlu_Cover_Code"
                    StatusBarText ="Stream cover percentage quadrat 2"
                    OnChange ="=UpdateMicrohabitat([Screen].[ActiveControl])"

                    LayoutCachedLeft =2905
                    LayoutCachedTop =6841
                    LayoutCachedWidth =3683
                    LayoutCachedHeight =7080
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =255
                    IMESentenceMode =3
                    ListRows =21
                    Left =3675
                    Top =6841
                    Width =763
                    Height =239
                    TabIndex =50
                    ColumnInfo ="\"\";\"\";\"6\";\"4\""
                    Name ="Stream_Q3"
                    RowSourceType ="Table/Query"
                    RowSource ="tlu_Cover_Code"
                    StatusBarText ="Stream cover percentage quadrat 3"
                    OnChange ="=UpdateMicrohabitat([Screen].[ActiveControl])"

                    LayoutCachedLeft =3675
                    LayoutCachedTop =6841
                    LayoutCachedWidth =4438
                    LayoutCachedHeight =7080
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =255
                    IMESentenceMode =3
                    ListRows =21
                    Left =2115
                    Top =7081
                    Width =778
                    Height =239
                    TabIndex =51
                    ColumnInfo ="\"\";\"\";\"6\";\"4\""
                    Name ="Trash_Junk_Q1"
                    RowSourceType ="Table/Query"
                    RowSource ="tlu_Cover_Code"
                    StatusBarText ="Trash/Junk cover percentage quadrat 1"
                    OnChange ="=UpdateMicrohabitat([Screen].[ActiveControl])"

                    LayoutCachedLeft =2115
                    LayoutCachedTop =7081
                    LayoutCachedWidth =2893
                    LayoutCachedHeight =7320
                    Begin
                        Begin Label
                            OldBorderStyle =1
                            OverlapFlags =255
                            Left =255
                            Top =7080
                            Width =1860
                            Height =240
                            Name ="lblTrashJunk"
                            Caption ="Trash/Junk"
                            LayoutCachedLeft =255
                            LayoutCachedTop =7080
                            LayoutCachedWidth =2115
                            LayoutCachedHeight =7320
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =255
                    IMESentenceMode =3
                    ListRows =21
                    Left =2905
                    Top =7081
                    Width =778
                    Height =239
                    TabIndex =52
                    ColumnInfo ="\"\";\"\";\"6\";\"4\""
                    Name ="Trash_Junk_Q2"
                    RowSourceType ="Table/Query"
                    RowSource ="tlu_Cover_Code"
                    StatusBarText ="Trash/Junk cover percentage quadrat 2"
                    OnChange ="=UpdateMicrohabitat([Screen].[ActiveControl])"

                    LayoutCachedLeft =2905
                    LayoutCachedTop =7081
                    LayoutCachedWidth =3683
                    LayoutCachedHeight =7320
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =247
                    IMESentenceMode =3
                    ListRows =21
                    Left =3675
                    Top =7081
                    Width =763
                    Height =239
                    TabIndex =53
                    ColumnInfo ="\"\";\"\";\"6\";\"4\""
                    Name ="Trash_Junk_Q3"
                    RowSourceType ="Table/Query"
                    RowSource ="tlu_Cover_Code"
                    StatusBarText ="Trash/Junk cover percentage quadrat 3"
                    OnChange ="=UpdateMicrohabitat([Screen].[ActiveControl])"

                    LayoutCachedLeft =3675
                    LayoutCachedTop =7081
                    LayoutCachedWidth =4438
                    LayoutCachedHeight =7320
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =255
                    IMESentenceMode =3
                    ListRows =21
                    Left =2115
                    Top =5161
                    Width =778
                    Height =239
                    TabIndex =27
                    ColumnInfo ="\"\";\"\";\"6\";\"4\""
                    Name ="Dead_Root_Bole_Q1"
                    RowSourceType ="Table/Query"
                    RowSource ="tlu_Cover_Code"
                    StatusBarText ="Root/Bole cover percentage quadrat 1"
                    OnChange ="=UpdateMicrohabitat([Screen].[ActiveControl])"

                    LayoutCachedLeft =2115
                    LayoutCachedTop =5161
                    LayoutCachedWidth =2893
                    LayoutCachedHeight =5400
                    Begin
                        Begin Label
                            OldBorderStyle =1
                            OverlapFlags =255
                            Left =255
                            Top =5160
                            Width =1860
                            Height =240
                            Name ="lblDeadRootBole"
                            Caption ="Dead root/Bole"
                            LayoutCachedLeft =255
                            LayoutCachedTop =5160
                            LayoutCachedWidth =2115
                            LayoutCachedHeight =5400
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =255
                    IMESentenceMode =3
                    ListRows =21
                    Left =2905
                    Top =5161
                    Width =778
                    Height =239
                    TabIndex =28
                    ColumnInfo ="\"\";\"\";\"6\";\"4\""
                    Name ="Dead_Root_Bole_Q2"
                    RowSourceType ="Table/Query"
                    RowSource ="tlu_Cover_Code"
                    StatusBarText ="Roo/Bole cover percentage quadrat 2"
                    OnChange ="=UpdateMicrohabitat([Screen].[ActiveControl])"

                    LayoutCachedLeft =2905
                    LayoutCachedTop =5161
                    LayoutCachedWidth =3683
                    LayoutCachedHeight =5400
                End
                Begin Label
                    BackStyle =1
                    OldBorderStyle =1
                    OverlapFlags =223
                    Left =180
                    Top =1545
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
                    LayoutCachedLeft =180
                    LayoutCachedTop =1545
                    LayoutCachedWidth =4680
                    LayoutCachedHeight =2985
                    ThemeFontIndex =1
                    BackThemeColorIndex =6
                    BackTint =20.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin ToggleButton
                    Enabled = NotDefault
                    OverlapFlags =215
                    Left =4290
                    Top =2505
                    Width =270
                    Height =300
                    FontSize =10
                    FontWeight =700
                    TabIndex =57
                    Name ="tglNoExoticsQ3"
                    ControlSource ="NoExotics_Q3"
                    FontName ="Calibri"
                    ControlTipText ="Q3 has no priority 1 exotics"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120
                    GridlineColor =10921638

                    LayoutCachedLeft =4290
                    LayoutCachedTop =2505
                    LayoutCachedWidth =4560
                    LayoutCachedHeight =2805
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
                            OverlapFlags =223
                            Left =4260
                            Top =1665
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
                            LayoutCachedLeft =4260
                            LayoutCachedTop =1665
                            LayoutCachedWidth =4590
                            LayoutCachedHeight =1980
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
                    Left =3855
                    Top =2505
                    Width =270
                    Height =299
                    FontSize =10
                    FontWeight =700
                    TabIndex =58
                    Name ="tglNoExoticsQ2"
                    FontName ="Calibri"
                    ControlTipText ="Q2 has no priority 1 exotics"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120
                    GridlineColor =10921638

                    LayoutCachedLeft =3855
                    LayoutCachedTop =2505
                    LayoutCachedWidth =4125
                    LayoutCachedHeight =2804
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
                            OverlapFlags =223
                            Left =3825
                            Top =1665
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
                            LayoutCachedLeft =3825
                            LayoutCachedTop =1665
                            LayoutCachedWidth =4155
                            LayoutCachedHeight =1980
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
                    Left =3420
                    Top =2505
                    Width =270
                    Height =299
                    FontSize =10
                    FontWeight =700
                    TabIndex =59
                    Name ="tglNoExoticsQ1"
                    FontName ="Calibri"
                    ControlTipText ="Q1 has no priority 1 exotics"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120
                    GridlineColor =10921638

                    LayoutCachedLeft =3420
                    LayoutCachedTop =2505
                    LayoutCachedWidth =3690
                    LayoutCachedHeight =2804
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
                            OverlapFlags =223
                            Left =3405
                            Top =1665
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
                            LayoutCachedLeft =3405
                            LayoutCachedTop =1665
                            LayoutCachedWidth =3735
                            LayoutCachedHeight =1980
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
                    Left =2655
                    Top =2520
                    Width =270
                    Height =269
                    FontSize =10
                    FontWeight =700
                    TabIndex =60
                    Name ="tglNoExoticsT"
                    FontName ="Calibri"
                    ControlTipText ="Transect has no exotics"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120
                    GridlineColor =10921638

                    LayoutCachedLeft =2655
                    LayoutCachedTop =2520
                    LayoutCachedWidth =2925
                    LayoutCachedHeight =2789
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
                    Left =300
                    Top =2505
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
                    LayoutCachedLeft =300
                    LayoutCachedTop =2505
                    LayoutCachedWidth =2325
                    LayoutCachedHeight =2790
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
                    Left =4290
                    Top =2085
                    Width =270
                    Height =300
                    FontSize =10
                    FontWeight =700
                    TabIndex =61
                    Name ="tglNotSampledQ3"
                    FontName ="Calibri"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Q3 was not sampled"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120
                    GridlineColor =10921638

                    LayoutCachedLeft =4290
                    LayoutCachedTop =2085
                    LayoutCachedWidth =4560
                    LayoutCachedHeight =2385
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
                    Left =3855
                    Top =2085
                    Width =270
                    Height =299
                    FontSize =10
                    FontWeight =700
                    TabIndex =62
                    Name ="tglNotSampledQ2"
                    FontName ="Calibri"
                    ControlTipText ="Q2 was not sampled"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120
                    GridlineColor =10921638

                    LayoutCachedLeft =3855
                    LayoutCachedTop =2085
                    LayoutCachedWidth =4125
                    LayoutCachedHeight =2384
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
                    Left =3420
                    Top =2085
                    Width =270
                    Height =299
                    FontSize =10
                    FontWeight =700
                    TabIndex =63
                    Name ="tglNotSampledQ1"
                    FontName ="Calibri"
                    ControlTipText ="Q1 was not sampled"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120
                    GridlineColor =10921638

                    LayoutCachedLeft =3420
                    LayoutCachedTop =2085
                    LayoutCachedWidth =3690
                    LayoutCachedHeight =2384
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
                    Left =2655
                    Top =2100
                    Width =270
                    Height =269
                    FontSize =10
                    FontWeight =700
                    TabIndex =64
                    Name ="tglNotSampledT"
                    FontName ="Calibri"
                    ControlTipText ="Transect was not sampled"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120
                    GridlineColor =10921638

                    LayoutCachedLeft =2655
                    LayoutCachedTop =2100
                    LayoutCachedWidth =2925
                    LayoutCachedHeight =2369
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
                    Left =300
                    Top =2085
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
                    LayoutCachedLeft =300
                    LayoutCachedTop =2085
                    LayoutCachedWidth =2325
                    LayoutCachedHeight =2370
                    ThemeFontIndex =1
                    BackThemeColorIndex =1
                    BorderThemeColorIndex =0
                    BorderTint =50.0
                    ForeThemeColorIndex =0
                    ForeTint =85.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin TextBox
                    SpecialEffect =0
                    OldBorderStyle =0
                    OverlapFlags =215
                    BackStyle =0
                    IMESentenceMode =3
                    Left =300
                    Top =1140
                    Width =2820
                    Height =255
                    FontSize =7
                    TabIndex =66
                    ForeColor =8355711
                    Name ="tbxTransectID"
                    ControlSource ="t_Transect_ID"

                    LayoutCachedLeft =300
                    LayoutCachedTop =1140
                    LayoutCachedWidth =3120
                    LayoutCachedHeight =1395
                    ForeThemeColorIndex =1
                    ForeShade =50.0
                End
                Begin TextBox
                    SpecialEffect =0
                    OldBorderStyle =0
                    OverlapFlags =223
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =1980
                    Top =7380
                    Width =840
                    Height =255
                    TabIndex =67
                    ForeColor =8355711
                    Name ="tbxQ1"

                    LayoutCachedLeft =1980
                    LayoutCachedTop =7380
                    LayoutCachedWidth =2820
                    LayoutCachedHeight =7635
                    ForeThemeColorIndex =1
                    ForeShade =50.0
                End
                Begin TextBox
                    SpecialEffect =0
                    OldBorderStyle =0
                    OverlapFlags =223
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =2880
                    Top =7380
                    Width =840
                    Height =255
                    TabIndex =68
                    ForeColor =8355711
                    Name ="tbxQ2"

                    LayoutCachedLeft =2880
                    LayoutCachedTop =7380
                    LayoutCachedWidth =3720
                    LayoutCachedHeight =7635
                    ForeThemeColorIndex =1
                    ForeShade =50.0
                End
                Begin TextBox
                    SpecialEffect =0
                    OldBorderStyle =0
                    OverlapFlags =223
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =3780
                    Top =7380
                    Width =840
                    Height =255
                    TabIndex =69
                    ForeColor =8355711
                    Name ="tbxQ3"

                    LayoutCachedLeft =3780
                    LayoutCachedTop =7380
                    LayoutCachedWidth =4620
                    LayoutCachedHeight =7635
                    ForeThemeColorIndex =1
                    ForeShade =50.0
                End
                Begin TextBox
                    SpecialEffect =0
                    OldBorderStyle =0
                    OverlapFlags =223
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =1980
                    Top =7680
                    Width =840
                    Height =255
                    TabIndex =70
                    ForeColor =8355711
                    Name ="tbxQ1IS"
                    ControlSource ="IsSampled_Q1"

                    LayoutCachedLeft =1980
                    LayoutCachedTop =7680
                    LayoutCachedWidth =2820
                    LayoutCachedHeight =7935
                    ForeThemeColorIndex =1
                    ForeShade =50.0
                End
                Begin TextBox
                    SpecialEffect =0
                    OldBorderStyle =0
                    OverlapFlags =223
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =2880
                    Top =7680
                    Width =840
                    Height =255
                    TabIndex =71
                    ForeColor =8355711
                    Name ="tbxQ2IS"
                    ControlSource ="IsSampled_Q2"

                    LayoutCachedLeft =2880
                    LayoutCachedTop =7680
                    LayoutCachedWidth =3720
                    LayoutCachedHeight =7935
                    ForeThemeColorIndex =1
                    ForeShade =50.0
                End
                Begin TextBox
                    SpecialEffect =0
                    OldBorderStyle =0
                    OverlapFlags =223
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =3780
                    Top =7680
                    Width =840
                    Height =255
                    TabIndex =72
                    ForeColor =8355711
                    Name ="tbxQ3IS"
                    ControlSource ="IsSampled_Q3"

                    LayoutCachedLeft =3780
                    LayoutCachedTop =7680
                    LayoutCachedWidth =4620
                    LayoutCachedHeight =7935
                    ForeThemeColorIndex =1
                    ForeShade =50.0
                End
                Begin TextBox
                    SpecialEffect =0
                    OldBorderStyle =0
                    OverlapFlags =223
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =1980
                    Top =7980
                    Width =840
                    Height =255
                    TabIndex =73
                    ForeColor =8355711
                    Name ="tbxQ1NE"
                    ControlSource ="NoExotics_Q1"

                    LayoutCachedLeft =1980
                    LayoutCachedTop =7980
                    LayoutCachedWidth =2820
                    LayoutCachedHeight =8235
                    ForeThemeColorIndex =1
                    ForeShade =50.0
                End
                Begin TextBox
                    SpecialEffect =0
                    OldBorderStyle =0
                    OverlapFlags =223
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =2880
                    Top =7980
                    Width =840
                    Height =255
                    TabIndex =74
                    ForeColor =8355711
                    Name ="tbxQ2NE"
                    ControlSource ="NoExotics_Q2"

                    LayoutCachedLeft =2880
                    LayoutCachedTop =7980
                    LayoutCachedWidth =3720
                    LayoutCachedHeight =8235
                    ForeThemeColorIndex =1
                    ForeShade =50.0
                End
                Begin TextBox
                    SpecialEffect =0
                    OldBorderStyle =0
                    OverlapFlags =223
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =3780
                    Top =7980
                    Width =840
                    Height =255
                    TabIndex =75
                    ForeColor =8355711
                    Name ="tbxQ3NE"
                    ControlSource ="NoExotics_Q3"
                    ConditionalFormat = Begin
                        0x01000000ee000000010000000100000000000000000000004600000001000000 ,
                        0xececec00ffffff00000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x49004900660028005b0066007300750062005f00530070006500630069006500 ,
                        0x73005f00430075007200720065006e0074005d002e005b0046006f0072006d00 ,
                        0x5d002e005b0043006f006e00740072006f006c0073005d002800220074006200 ,
                        0x78004400650076004d006f0064006500220029003d00460061006c0073006500 ,
                        0x2c0031002c003000290000000000
                    End

                    LayoutCachedLeft =3780
                    LayoutCachedTop =7980
                    LayoutCachedWidth =4620
                    LayoutCachedHeight =8235
                    ForeThemeColorIndex =1
                    ForeShade =50.0
                    ConditionalFormat14 = Begin
                        0x010001000000010000000000000001000000ececec00ffffff00450000004900 ,
                        0x4900660028005b0066007300750062005f005300700065006300690065007300 ,
                        0x5f00430075007200720065006e0074005d002e005b0046006f0072006d005d00 ,
                        0x2e005b0043006f006e00740072006f006c0073005d0028002200740062007800 ,
                        0x4400650076004d006f0064006500220029003d00460061006c00730065002c00 ,
                        0x31002c0030002900000000000000000000000000000000000000000000
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =247
                    IMESentenceMode =3
                    ListRows =21
                    Left =3675
                    Top =5161
                    Width =763
                    Height =239
                    TabIndex =29
                    ColumnInfo ="\"\";\"\";\"6\";\"4\""
                    Name ="Dead_Root_Bole_Q3"
                    RowSourceType ="Table/Query"
                    RowSource ="tlu_Cover_Code"
                    StatusBarText ="Root/Bole cover percentage quadrat 3"
                    OnChange ="=UpdateMicrohabitat([Screen].[ActiveControl])"

                    LayoutCachedLeft =3675
                    LayoutCachedTop =5161
                    LayoutCachedWidth =4438
                    LayoutCachedHeight =5400
                End
                Begin TextBox
                    SpecialEffect =0
                    OldBorderStyle =0
                    OverlapFlags =215
                    BackStyle =0
                    IMESentenceMode =3
                    Left =3240
                    Top =1140
                    Width =2820
                    Height =255
                    TabIndex =76
                    ForeColor =8355711
                    Name ="tbxStart"
                    ControlSource ="Start_Time"

                    LayoutCachedLeft =3240
                    LayoutCachedTop =1140
                    LayoutCachedWidth =6060
                    LayoutCachedHeight =1395
                    ForeThemeColorIndex =1
                    ForeShade =50.0
                End
                Begin Rectangle
                    Visible = NotDefault
                    SpecialEffect =0
                    BackStyle =1
                    OldBorderStyle =0
                    OverlapFlags =247
                    Left =240
                    Top =7380
                    Width =4440
                    Height =1260
                    BackColor =14478320
                    Name ="bxHide"
                    LayoutCachedLeft =240
                    LayoutCachedTop =7380
                    LayoutCachedWidth =4680
                    LayoutCachedHeight =8640
                End
                Begin TextBox
                    SpecialEffect =0
                    OldBorderStyle =0
                    OverlapFlags =223
                    BackStyle =0
                    IMESentenceMode =3
                    Left =2265
                    Top =1665
                    Width =990
                    Height =315
                    FontSize =11
                    FontWeight =700
                    TabIndex =79
                    BackColor =16777215
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="lblTransect"
                    ControlSource ="=\"Transect\""
                    FontName ="Calibri"
                    ControlTipText ="Transect flags"
                    ConditionalFormat = Begin
                        0x01000000b4000000020000000100000000000000000000001400000001010000 ,
                        0x7f7f7f00ffffff00010000000000000015000000290000000101000092929200 ,
                        0xffffff0000000000000000000000000000000000000000000000000000000000 ,
                        0x5b0074006200780049007300530061006d0070006c0065006400530075006d00 ,
                        0x5d003d003000000000005b007400620078004e006f00450078006f0074006900 ,
                        0x63007300530075006d005d003d00330000000000
                    End
                    GridlineColor =10921638

                    LayoutCachedLeft =2265
                    LayoutCachedTop =1665
                    LayoutCachedWidth =3255
                    LayoutCachedHeight =1980
                    BackThemeColorIndex =1
                    BorderThemeColorIndex =0
                    BorderTint =50.0
                    ThemeFontIndex =1
                    ConditionalFormat14 = Begin
                        0x0100020000000100000000000000010100007f7f7f00ffffff00130000005b00 ,
                        0x74006200780049007300530061006d0070006c0065006400530075006d005d00 ,
                        0x3d00300000000000000000000000000000000000000000000001000000000000 ,
                        0x000101000092929200ffffff00130000005b007400620078004e006f00450078 ,
                        0x006f007400690063007300530075006d005d003d003300000000000000000000 ,
                        0x000000000000000000000000
                    End
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin Rectangle
                    Visible = NotDefault
                    SpecialEffect =0
                    OldBorderStyle =0
                    OverlapFlags =255
                    Left =195
                    Top =1620
                    Width =4560
                    Height =360
                    BackColor =-2147483633
                    Name ="bxCoverLabels"
                    OnClick ="[Event Procedure]"
                    LayoutCachedLeft =195
                    LayoutCachedTop =1620
                    LayoutCachedWidth =4755
                    LayoutCachedHeight =1980
                End
                Begin Subform
                    OverlapFlags =255
                    Left =4755
                    Top =1500
                    Width =8265
                    Height =6718
                    TabIndex =54
                    Name ="fsub_Species_Current"
                    SourceObject ="Form.fsub_Species"
                    LinkChildFields ="Transect_ID"
                    LinkMasterFields ="t_Transect_ID"

                    LayoutCachedLeft =4755
                    LayoutCachedTop =1500
                    LayoutCachedWidth =13020
                    LayoutCachedHeight =8218
                End
                Begin Subform
                    Visible = NotDefault
                    OverlapFlags =247
                    SpecialEffect =0
                    Left =7215
                    Top =3072
                    Width =3300
                    Height =599
                    TabIndex =65
                    BorderColor =2366701
                    Name ="fsub_Message"
                    SourceObject ="Form.fsub_Msg"

                    LayoutCachedLeft =7215
                    LayoutCachedTop =3072
                    LayoutCachedWidth =10515
                    LayoutCachedHeight =3671
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =215
                    IMESentenceMode =3
                    ColumnCount =3
                    ListWidth =1650
                    Left =4020
                    Top =810
                    Width =1620
                    TabIndex =81
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"\";\"\";\"10\";\"100\""
                    Name ="cbxRecorder"
                    ControlSource ="Recorder"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT tlu_Contacts.Contact_ID, tlu_Contacts.Last_Name, tlu_Contacts.First_Name "
                        "FROM tlu_Contacts;"
                    ColumnWidths ="0;810;839"
                    AfterUpdate ="=UpdateTransect([Screen].[ActiveControl])"
                    LayoutCachedLeft =4020
                    LayoutCachedTop =810
                    LayoutCachedWidth =5640
                    LayoutCachedHeight =1050
                    Begin
                        Begin Label
                            OverlapFlags =215
                            Left =3060
                            Top =810
                            Width =870
                            Height =240
                            FontWeight =700
                            Name ="Label92"
                            Caption ="Recorder"
                            LayoutCachedLeft =3060
                            LayoutCachedTop =810
                            LayoutCachedWidth =3930
                            LayoutCachedHeight =1050
                        End
                    End
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
' Version:      1.05
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
'               BLC - 4/25/2017 - 1.04 - revised to save quadrat flags to db
'               BLC - 7/10/2017 - 1.05 - added check for new transects, create new quadrats, quadrat surface
'                                        microhabitat records
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
' Assumptions:  Newly imported transects have no quadrats, so these must be added.
'               New quadrats also require new records for @ microhabitat surface, these
'               too must be added.
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
'   BLC - 4/25/2017 - revised to save quadrat flags to database
'   BLC - 7/10/2017 - added check for new transects, create new quadrats, quadrat surface
'                     microhabitat records, moved usys_temp_transect update to
'                     launching form (frm_Visit_Date)
'   BLC - 7/13/2017 - set development controls to show/hide based on DEV_MODE setting
' ---------------------------------
Private Sub Form_Open(Cancel As Integer)
On Error GoTo Err_Handler
    
    'show/hide dev mode controls
    bxHide.Visible = Not DEV_MODE
    
    'set form recordsource
    Me.RecordSource = "usys_temp_transect"
    
    Dim t As New VegTransect
    
    'check if transect has quadrats
    With t
        .TransectQuadratID = Me.tbxTransectID '"20170705114218-705547511.577606"
                        
        'newly imported transects have 0 quadrats --> create them & the associated
        '                                             surface microhabitat records
        If .NumQuadrats = 0 Then
            .AddQuadrats
            .AddSurfaceMicrohabitats
            
            Me.Refresh
        End If
    
    End With
            
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
        If Left(ctl.name, 3) = "tgl" Then
            ctl.Enabled = True
            ctl.ForeColor = lngBlack
        End If
    Next

    'initialize Quadrat # temp vars
    Me.tbxQ1 = 0
    Me.tbxQ2 = 0
    Me.tbxQ3 = 0
  
    'set starting transect # to red
    '(1st & last transects are red for bounding since Next/Previous cycle)
    Me.Transect.ForeColor = lngRed
    
    'populate the microhabitats from SurfaceCover
    PopulateMicrohabitats
    
    'populate the transect & quadrat flag toggles based on usys_temp_transect
    PopulateFlagToggles
  
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
'   BLC - 7/12/2017 - added PopulateFlagToggles
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
    
        'if species subform has records --> disable transect & quadrat toggles (IsSampled, NoExotics)
        If .HasRecords = True Then
            
            'check if Q1,Q2,Q3 % Cover values are set
           Debug.Print .Plant_Code
           Debug.Print .HasRecordsQ1 & "-"; .HasRecordsQ2 & "-"; .HasRecordsQ3
            'disable transect & quadrat toggles
            DisableToggles
            
            'enable select toggles depending on which quadrats have records
            Dim aryToggles() As String
            Dim strToggles As String
            Dim i As Integer
            Dim tgl As Variant
                        
            'initialize
            strToggles = "1,2,3"

            With Me.fsub_Species_Current.Form
            
                If .Controls("tbxQ1_Sampled") > 0 Then _
                    strToggles = Replace(strToggles, "1,", "")
                If .Controls("tbxQ2_Sampled") > 0 Then _
                    strToggles = Replace(strToggles, "2,", "")
                If .Controls("tbxQ3_Sampled") > 0 Then _
                    strToggles = Replace(strToggles, "3", "")
            
            End With
Debug.Print "PreTrim: " & strToggles

            If Len(strToggles) > 0 Then

                'trim any ending commas
                strToggles = IIf(Right(strToggles, 1) = ",", Left(strToggles, Len(strToggles) - 1), strToggles)
            
Debug.Print "PostTrim: " & strToggles
                aryToggles = Split(strToggles, ",")
                
                'iterate to enable IF any toggles are left
                If IsArray(aryToggles) Then
                    For Each tgl In aryToggles
                        
Debug.Print "tgl: " & tgl
                        EnableToggles CInt(tgl)
                    Next
                End If
            End If
        Else
            
            'enable transect & quadrat toggles
            EnableToggles
        
        End If
    End With

    'populate the transect & quadrat flag toggles based on usys_temp_transect
    PopulateFlagToggles
    
    'turn on/off disabled message
    ToggleDisabledMessage
    
    Me.Form.Repaint

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
' Sub:          tglNotSampledQ3_Click
' Description:  Toggles Q3 not sampled flag
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, July 17, 2017 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 7/17/2016 - initial version
' ---------------------------------
Private Sub tglNotSampledQ3_Click()
On Error GoTo Err_Handler

    UpdateFlags 3

'    Dim IsSampledQ3 As Integer
'    Dim NoExoticsQ3 As Integer
'
'    IsSampledQ3 = IIf(tglNotSampledQ3 = strCheck, 0, 1)
'    NoExoticsQ3 = IIf(tglNoExoticsQ3 = strCheck, 1, 0)
'
'    Dim q As New Quadrat
'
'    With q
'        .QuadratID = tbxQ3
'        .IsSampledQ3 = IsSampledQ3
'        .NoExoticsQ3 = NoExoticsQ3
'        .UpdateQuadratFlags
'    End With
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - tglQ3NotSampled_Click[frm_Quadrat_Transect form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          UpdateFlags
' Description:  Updates not sampled/no exotics flags
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, July 17, 2017 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 7/17/2016 - initial version
' ---------------------------------
Private Sub UpdateFlags(QuadratNumber As Integer)
On Error GoTo Err_Handler

'    Dim IsSampled, IsSampledQ1, IsSampledQ2, IsSampledQ3 As Integer
'    Dim NoExotics, NoExoticsQ1, NoExoticsQ2, NoExoticsQ3 As Integer
    Dim IsSampled As Integer
    Dim NoExotics As Integer
    Dim strNotSampled, strNoExotics, strTbx As String
    Dim tglNotSampled As ToggleButton
    Dim tglNoExotics As ToggleButton
    Dim tbx As TextBox
    
    'control names
    strNotSampled = "tglNotSampledQ" & QuadratNumber
    'strNoExotics =
    strTbx = "tbxQ" & QuadratNumber
    
    'set controls
    Set tglNotSampled = Me.Controls("tglNotSampledQ" & QuadratNumber)
    Set tglNoExotics = Me.Controls("tglNoExoticsQ" & QuadratNumber)
    Set tbx = Me.Controls("tbxQ" & QuadratNumber)
    
    'set value
    IsSampled = IIf(tglNotSampled.Caption = strCheck, 0, 1)
    NoExotics = IIf(tglNoExotics.Caption = strCheck, 1, 0)
    
    Dim q As New Quadrat
    
    With q
        .QuadratID = tbx 'Q3
        .QuadratNumber = QuadratNumber
        
        Select Case QuadratNumber
            Case 1
                .IsSampledQ1 = IsSampled
                .NoExoticsQ1 = NoExotics
            Case 2
                .IsSampledQ2 = IsSampled
                .NoExoticsQ2 = NoExotics
            Case 3
                .IsSampledQ3 = IsSampled 'Q3
                .NoExoticsQ3 = NoExotics 'Q3
        End Select
        
        .UpdateQuadratFlags
        
    End With
    
    Me.Requery
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - UpdateFlags[frm_Quadrat_Transect form])"
    End Select
    Resume Exit_Handler
End Sub

'' ---------------------------------
'' Sub:          tglNotSampledT_LostFocus
'' Description:  Toggle lost focus actions
'' Assumptions:  -
'' Parameters:   -
'' Returns:      -
'' Throws:       none
'' References:   -
'' Source/date:  Bonnie Campbell, July 14, 2017 - for NCPN tools
'' Adapted:      -
'' Revisions:
''   BLC - 7/14/2017 - initial version
'' ---------------------------------
'Private Sub tglNotSampledT_LostFocus()
'On Error GoTo Err_Handler
'
'    Debug.Print tglNotSampledT
'
'Exit_Handler:
'    Exit Sub
'Err_Handler:
'    Select Case Err.Number
'      Case Else
'        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
'            "Error encountered (#" & Err.Number & " - tglNotSampledT_LostFocus[frm_Quadrat_Transect form])"
'    End Select
'    Resume Exit_Handler
'End Sub

'' ---------------------------------
'' Sub:          tglNotSampledQ1_LostFocus
'' Description:  Toggle lost focus actions
'' Assumptions:  -
'' Parameters:   -
'' Returns:      -
'' Throws:       none
'' References:   -
'' Source/date:  Bonnie Campbell, July 14, 2017 - for NCPN tools
'' Adapted:      -
'' Revisions:
''   BLC - 7/14/2017 - initial version
'' ---------------------------------
'Private Sub tglNotSampledQ1_LostFocus()
'On Error GoTo Err_Handler
'
'    Debug.Print tglNotSampledQ1
'
'Exit_Handler:
'    Exit Sub
'Err_Handler:
'    Select Case Err.Number
'      Case Else
'        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
'            "Error encountered (#" & Err.Number & " - tglNotSampledQ1_LostFocus[frm_Quadrat_Transect form])"
'    End Select
'    Resume Exit_Handler
'End Sub

'' ---------------------------------
'' Sub:          tglNotSampledT_AfterUpdate
'' Description:  Toggle after update actions
'' Assumptions:  Transect not sampled? -> no priority 1 species either
'' Parameters:   -
'' Returns:      -
'' Throws:       none
'' References:   -
'' Source/date:  Bonnie Campbell, March 8, 2017 - for NCPN tools
'' Adapted:      -
'' Revisions:
''   BLC - 3/8/2017 - initial version
'' ---------------------------------
'Private Sub tglNotSampledT_AfterUpdate()
'On Error GoTo Err_Handler
'
'    Dim i As Integer
'    Dim strControl As String
'
' '   strCheck = StringFromCodepoint(uCheck)
'
'    'display as checkbox
'    ToggleCaption tglNotSampledT, True
'
'    SetToggles Me.tglNotSampledT
'
''    SetQuadratToggles "NotSampled"
''
''    If tglNotSampledT.Caption = strCheck Then
''        'set all no exotics as well
''        '(can't have exotics w/o sampling)
''        tglNoExoticsT.Caption = strCheck
''        tglNoExoticsT.Enabled = False
''
''        SetQuadratToggles "NoExotics"
''
''    Else
''        'enable no exotics if false
''        tglNoExoticsT.Enabled = True
''
''        'clear & enable Q1->3
''        For i = 1 To 3
''            strControl = "tglNotSampledQ" & i
''            Me.Controls(strControl).Enabled = True
''            Me.Controls(strControl).Caption = ""
''        Next
''    End If
''
''
'''    'check if Transect level checked
'''    If tglNotSampledT.Caption = StringFromCodepoint(uCheck) Then
'''
'''        'set Q1-Q3 flags & disable
'''        For i = 1 To 3
'''            strControl = "tglNotSampledQ" & i
'''            Controls(strControl).Caption = StringFromCodepoint(uCheck)
'''            Controls(strControl).Enabled = False
'''        Next
'''
'''    Else
'''
'''        'ensure Q1-Q3 flags are enabled & checks are cleared
'''        For i = 1 To 3
'''            strControl = "tglNotSampledQ" & i
'''            Controls(strControl).Caption = ""
'''            Controls(strControl).Enabled = True
'''        Next
'''
'''    End If
''
''    If tglNotSampledT Then _
''        ReadyForSave
'
'Exit_Handler:
'    Exit Sub
'Err_Handler:
'    Select Case Err.Number
'      Case Else
'        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
'            "Error encountered (#" & Err.Number & " - tglNotSampledT_AfterUpdate[frm_Quadrat_Transect form])"
'    End Select
'    Resume Exit_Handler
'End Sub

'' ---------------------------------
'' Sub:          tglNotSampledQ1_AfterUpdate
'' Description:  Toggle after update actions
'' Assumptions:  -
'' Parameters:   -
'' Returns:      -
'' Throws:       none
'' References:   -
'' Source/date:  Bonnie Campbell, March 8, 2017 - for NCPN tools
'' Adapted:      -
'' Revisions:
''   BLC - 3/8/2017 - initial version
'' ---------------------------------
'Private Sub tglNotSampledQ1_AfterUpdate()
'On Error GoTo Err_Handler
'
'    'display as checkbox
'    ToggleCaption tglNotSampledQ1, True
'
'    SetToggles Me.tglNotSampledQ1
'
''    CheckTransectLevel "NotSampled"
''
''    If tglNotSampledQ1.Caption = strCheck Then
''        'not sampled? -> no priority 1 exotics either
''        tglNoExoticsQ1.Caption = strCheck
''        tglNoExoticsQ1.Enabled = False
''    Else
''        If tglNoExoticsT.Caption <> strCheck Then
''            'sampled? -> priority 1 exotics ok
''            tglNoExoticsQ1.Caption = ""
''            tglNoExoticsQ1.Enabled = True
''        End If
''    End If
''
'''    If tglNotSampledQ1 Then _
'''        ReadyForSave
'
'Exit_Handler:
'    Exit Sub
'Err_Handler:
'    Select Case Err.Number
'      Case Else
'        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
'            "Error encountered (#" & Err.Number & " - tglNotSampledQ1_AfterUpdate[frm_Quadrat_Transect form])"
'    End Select
'    Resume Exit_Handler
'End Sub
'
'' ---------------------------------
'' Sub:          tglNotSampledQ2_AfterUpdate
'' Description:  Toggle after update actions
'' Assumptions:  -
'' Parameters:   -
'' Returns:      -
'' Throws:       none
'' References:   -
'' Source/date:  Bonnie Campbell, March 8, 2017 - for NCPN tools
'' Adapted:      -
'' Revisions:
''   BLC - 3/8/2017 - initial version
'' ---------------------------------
'Private Sub tglNotSampledQ2_AfterUpdate()
'On Error GoTo Err_Handler
'
'    'display as checkbox
'    ToggleCaption tglNotSampledQ2, True
'
'    CheckTransectLevel "NotSampled"
'
'    SetToggles Me.tglNotSampledQ2
'
''    If tglNotSampledQ2.Caption = strCheck Then
''        'not sampled? -> no priority 1 exotics either
''        tglNoExoticsQ2.Caption = strCheck
''        tglNoExoticsQ2.Enabled = False
''    Else
''        If tglNoExoticsT.Caption <> strCheck Then
''            'sampled? -> priority 1 exotics ok
''            tglNoExoticsQ2.Caption = ""
''            tglNoExoticsQ2.Enabled = True
''        End If
''    End If
'
''    If tglNotSampledQ2 Then _
''        ReadyForSave
'
'Exit_Handler:
'    Exit Sub
'Err_Handler:
'    Select Case Err.Number
'      Case Else
'        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
'            "Error encountered (#" & Err.Number & " - tglNotSampledQ2_AfterUpdate[frm_Quadrat_Transect form])"
'    End Select
'    Resume Exit_Handler
'End Sub
'
'' ---------------------------------
'' Sub:          tglNotSampledQ3_AfterUpdate
'' Description:  Toggle after update actions
'' Assumptions:  -
'' Parameters:   -
'' Returns:      -
'' Throws:       none
'' References:   -
'' Source/date:  Bonnie Campbell, March 8, 2017 - for NCPN tools
'' Adapted:      -
'' Revisions:
''   BLC - 3/8/2017 - initial version
'' ---------------------------------
'Private Sub tglNotSampledQ3_AfterUpdate()
'On Error GoTo Err_Handler
'
'    'display as checkbox
'    ToggleCaption tglNotSampledQ3, True
'
'    CheckTransectLevel "NotSampled"
'
'    SetToggles Me.tglNotSampledQ3
'
''    If tglNotSampledQ3.Caption = strCheck Then
''        'not sampled? -> no priority 1 exotics either
''        tglNoExoticsQ3.Caption = strCheck
''        tglNoExoticsQ3.Enabled = False
''    Else
''        If tglNoExoticsT.Caption <> strCheck Then
''            'sampled? -> priority 1 exotics ok
''            tglNoExoticsQ3.Caption = ""
''            tglNoExoticsQ3.Enabled = True
''        End If
''    End If
''
'''    If tglNotSampledQ3 Then _
'''        ReadyForSave
'
'Exit_Handler:
'    Exit Sub
'Err_Handler:
'    Select Case Err.Number
'      Case Else
'        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
'            "Error encountered (#" & Err.Number & " - tglNotSampledQ3_AfterUpdate[frm_Quadrat_Transect form])"
'    End Select
'    Resume Exit_Handler
'End Sub

' =================================
'   No Exotics Flag
' =================================

'' ---------------------------------
'' Sub:          tglNoExoticsT_LostFocus
'' Description:  Toggle lost focus actions
'' Assumptions:  -
'' Parameters:   -
'' Returns:      -
'' Throws:       none
'' References:   -
'' Source/date:  Bonnie Campbell, July 14, 2017 - for NCPN tools
'' Adapted:      -
'' Revisions:
''   BLC - 7/14/2017 - initial version
'' ---------------------------------
'Private Sub tglNoExoticsT_LostFocus()
'On Error GoTo Err_Handler
'
'    Debug.Print tglNoExoticsT
'
'Exit_Handler:
'    Exit Sub
'Err_Handler:
'    Select Case Err.Number
'      Case Else
'        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
'            "Error encountered (#" & Err.Number & " - tglNoExoticsT_LostFocus[frm_Quadrat_Transect form])"
'    End Select
'    Resume Exit_Handler
'End Sub
'
'' ---------------------------------
'' Sub:          tglNoExoticsQ1_LostFocus
'' Description:  Toggle lost focus actions
'' Assumptions:  -
'' Parameters:   -
'' Returns:      -
'' Throws:       none
'' References:   -
'' Source/date:  Bonnie Campbell, July 14, 2017 - for NCPN tools
'' Adapted:      -
'' Revisions:
''   BLC - 7/14/2017 - initial version
'' ---------------------------------
'Private Sub tglNoExoticsQ1_LostFocus()
'On Error GoTo Err_Handler
'
'    Debug.Print tglNoExoticsQ1
'
'Exit_Handler:
'    Exit Sub
'Err_Handler:
'    Select Case Err.Number
'      Case Else
'        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
'            "Error encountered (#" & Err.Number & " - tglNoExoticsQ1_LostFocus[frm_Quadrat_Transect form])"
'    End Select
'    Resume Exit_Handler
'End Sub
'
'
'' ---------------------------------
'' Sub:          tglNoExoticsT_AfterUpdate
'' Description:  Toggle after update actions
'' Assumptions:  -
'' Parameters:   -
'' Returns:      -
'' Throws:       none
'' References:   -
'' Source/date:  Bonnie Campbell, March 8, 2017 - for NCPN tools
'' Adapted:      -
'' Revisions:
''   BLC - 3/8/2017 - initial version
'' ---------------------------------
'Private Sub tglNoExoticsT_AfterUpdate()
'On Error GoTo Err_Handler
'
'    Dim i As Integer
'    Dim strControl As String
'
'    'display as checkbox
'    ToggleCaption tglNoExoticsT, True
'
'    SetToggles Me.tglNoExoticsT
'
''    SetQuadratToggles "NoExotics"
'
''    'check if Transect level checked
''    If tglNoExoticsT.Caption = StringFromCodepoint(uCheck) Then
''
''        'set Q1-Q3 flags & disable
''        For i = 1 To 3
''            strControl = "tglNoExoticsQ" & i
''            Controls(strControl).Caption = StringFromCodepoint(uCheck)
''            Controls(strControl).Enabled = False
''        Next
''
''    Else
''
''        'ensure Q1-Q3 flags are enabled & checks are cleared
''        For i = 1 To 3
''            strControl = "tglNoExoticsQ" & i
''            Controls(strControl).Caption = ""
''            Controls(strControl).Enabled = True
''        Next
''
''    End If
'
'    If tglNoExoticsT Then _
'        ReadyForSave
'
'Exit_Handler:
'    Exit Sub
'Err_Handler:
'    Select Case Err.Number
'      Case Else
'        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
'            "Error encountered (#" & Err.Number & " - tglNoExoticsT_AfterUpdate[frm_Quadrat_Transect form])"
'    End Select
'    Resume Exit_Handler
'End Sub
'
'' ---------------------------------
'' Sub:          tglNoExoticsQ1_AfterUpdate
'' Description:  Toggle after update actions
'' Assumptions:  -
'' Parameters:   -
'' Returns:      -
'' Throws:       none
'' References:   -
'' Source/date:  Bonnie Campbell, March 8, 2017 - for NCPN tools
'' Adapted:      -
'' Revisions:
''   BLC - 3/8/2017 - initial version
'' ---------------------------------
'Private Sub tglNoExoticsQ1_AfterUpdate()
'On Error GoTo Err_Handler
'
'    'display as checkbox
'    ToggleCaption tglNoExoticsQ1, True
'
'    CheckTransectLevel "NoExotics"
'
'    SetToggles Me.tglNoExoticsQ1
'
''    If tglNoExoticsQ1 Then _
''        ReadyForSave
'
'Exit_Handler:
'    Exit Sub
'Err_Handler:
'    Select Case Err.Number
'      Case Else
'        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
'            "Error encountered (#" & Err.Number & " - tglNoExoticsQ1_AfterUpdate[frm_Quadrat_Transect form])"
'    End Select
'    Resume Exit_Handler
'End Sub
'
'' ---------------------------------
'' Sub:          tglNoExoticsQ2_AfterUpdate
'' Description:  Toggle after update actions
'' Assumptions:  -
'' Parameters:   -
'' Returns:      -
'' Throws:       none
'' References:   -
'' Source/date:  Bonnie Campbell, March 8, 2017 - for NCPN tools
'' Adapted:      -
'' Revisions:
''   BLC - 3/8/2017 - initial version
'' ---------------------------------
'Private Sub tglNoExoticsQ2_AfterUpdate()
'On Error GoTo Err_Handler
'
'    'display as checkbox
'    ToggleCaption tglNoExoticsQ2, True
'
'    CheckTransectLevel "NoExotics"
'
'    SetToggles Me.tglNoExoticsQ2
'
''    If tglNoExoticsQ2 Then _
''        ReadyForSave
'
'Exit_Handler:
'    Exit Sub
'Err_Handler:
'    Select Case Err.Number
'      Case Else
'        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
'            "Error encountered (#" & Err.Number & " - tglNoExoticsQ2_AfterUpdate[frm_Quadrat_Transect form])"
'    End Select
'    Resume Exit_Handler
'End Sub
'
'' ---------------------------------
'' Sub:          tglNoExoticsQ3_AfterUpdate
'' Description:  Toggle after update actions
'' Assumptions:  -
'' Parameters:   -
'' Returns:      -
'' Throws:       none
'' References:   -
'' Source/date:  Bonnie Campbell, March 8, 2017 - for NCPN tools
'' Adapted:      -
'' Revisions:
''   BLC - 3/8/2017 - initial version
'' ---------------------------------
'Private Sub tglNoExoticsQ3_AfterUpdate()
'On Error GoTo Err_Handler
'
'    'display as checkbox
'    ToggleCaption tglNoExoticsQ3, True
'
'    CheckTransectLevel "NoExotics"
'
'    SetToggles Me.tglNoExoticsQ3
'
''    If tglNoExoticsQ3 Then _
''        ReadyForSave
'
'Exit_Handler:
'    Exit Sub
'Err_Handler:
'    Select Case Err.Number
'      Case Else
'        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
'            "Error encountered (#" & Err.Number & " - tglNoExoticsQ3_AfterUpdate[frm_Quadrat_Transect form])"
'    End Select
'    Resume Exit_Handler
'End Sub

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
'    Me.fsub_Species_Current!Average_Cover = Me.fsub_Species_Current.Form.CalcAvgCover
    
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
'    Me.fsub_Species_Current!Average_Cover = fsub_Species_Current.Form.CalcAvgCover
    
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
'   BLC - 7/14/2017 - revised to use IsSampled/NoExotic sums vs. toggle captions
' ---------------------------------
Private Sub ToggleDisabledMessage()
On Error GoTo Err_Handler

'    If tglNotSampledT.Caption = strCheck Or _
'       tglNoExoticsT.Caption = strCheck Then

    If Me.tbxIsSampledSum = 0 Or Me.tbxNoExoticsSum = 3 Then
       
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
    
    If Me.Controls(ToggleSet.name).Caption = strCheck Then _
        blnON = True
    
    'default
    strToggle = ToggleSet.name
       
    Select Case Replace(ToggleSet.name, "tgl", "")
    
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

    If Me.Controls(ToggleSet.name).Caption = strCheck Then
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
    If InStr(strToggle, "Q") > 0 Then
        strControl = Left(strToggle, Len(strToggle) - 1) 'remove last 1|2|3
    
        Count = 0
        
        For i = 1 To 3
            strControlQ = strControl & i
            If Controls(strControlQ).Caption = strCheck Then
                Count = Count + 1
            End If
        Next
    
        'all quadrats set? (if so, count = 1 + 2 + 3 = 6)
        
'        If Count < 6 Then
'
'            Debug.Print strControlQ & "6"
'
'        End If
    
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
' Sub:          PopulateFlagToggles
' Description:  Sets captions for transect & quadrat flag toggle buttons
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, July 12, 2017 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 7/12/2017 - initial version
' ---------------------------------
Private Sub PopulateFlagToggles()
On Error GoTo Err_Handler

    'populate toggles based on data
    tglNotSampledQ1.Caption = IIf(IsSampled_Q1 = 1, "", strCheck)
    tglNotSampledQ2.Caption = IIf(IsSampled_Q2 = 1, "", strCheck)
    tglNotSampledQ3.Caption = IIf(IsSampled_Q3 = 1, "", strCheck)
'Debug.Print "IS:" & IsSampled_Q1

    tglNoExoticsQ1.Caption = IIf(NoExotics_Q1 = 1, strCheck, "")
    tglNoExoticsQ2.Caption = IIf(NoExotics_Q2 = 1, strCheck, "")
    tglNoExoticsQ3.Caption = IIf(NoExotics_Q3 = 1, strCheck, "")

'Debug.Print "NE:" & NoExotics_Q1

    'populate transect level
    tglNotSampledT.Caption = IIf(Me.tbxIsSampledSum = 0, strCheck, "")
    tglNoExoticsT.Caption = IIf(Me.tbxNoExoticsSum = 3, strCheck, "")

Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - PopulateFlagToggles[frm_Quadrat_Transect form])"
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
' Function:     UpdateTransect
' Description:  Updates transect data (start time, observer, comments)
' Assumptions:  Control property contains the following
'                   =UpdateTransect([Screen].[ActiveControl])
'               in its on after update event property
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, July 17, 2017 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 7/17/2016 - initial version
' ---------------------------------
Private Function UpdateTransect(ctrl As Control) As Boolean
On Error GoTo Err_Handler

    Dim obs As Variant
    Dim start As Variant
    Dim cmt As Variant
    
    'set values from form
    start = Me.tbxStartTime
    obs = Me.cbxObserver
    cmt = Me.tbxComments

    Dim vt As New VegTransect
    
    With vt
        .TransectQuadratID = Me.tbxTransectID
        
        Select Case ctrl.name
            Case "tbxStartDate"
                'start time
                If Not IsNull(start) Then
                    .StartTime = start
                    .UpdateStartTime
                End If
                
            Case "cbxObserver"
                'observer
                If Not IsNull(obs) Then
                    .Observer = obs
                    .UpdateObserver
                End If
            
            Case "tbxComments"
                'comments
                If Not IsNull(cmt) Then
                    .Comments = cmt
                    .UpdateComments
                End If
        End Select
        
    End With
    
    Me.Requery
    
Exit_Handler:
    Exit Function
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - UpdateTransectData[frm_Data_Entry form])"
    End Select
    Resume Exit_Handler
End Function

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
