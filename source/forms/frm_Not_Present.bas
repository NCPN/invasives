Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    AllowDeletions = NotDefault
    AllowAdditions = NotDefault
    AllowEdits = NotDefault
    ScrollBars =2
    TabularFamily =0
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =13680
    DatasheetFontHeight =9
    ItemSuffix =40
    Left =270
    Top =600
    Right =14340
    Bottom =5145
    DatasheetGridlinesColor =12632256
    Filter ="Unit_Code = 'BRCA' AND BRCA <> 'Present'"
    RecSrcDt = Begin
        0x779484990c1fe340
    End
    RecordSource ="qry_sel_Not_Present"
    Caption ="frm_Not_Present"
    DatasheetFontName ="Arial"
    OnLoad ="[Event Procedure]"
    Begin
        Begin Label
            BackStyle =0
            TextAlign =1
            FontWeight =700
            BackColor =-2147483633
            ForeColor =-2147483630
        End
        Begin Rectangle
            SpecialEffect =3
            BackStyle =0
        End
        Begin Image
            BackStyle =0
            OldBorderStyle =0
            PictureAlignment =2
        End
        Begin CommandButton
            FontSize =8
            FontWeight =400
            FontName ="MS Sans Serif"
        End
        Begin OptionButton
            SpecialEffect =2
            LabelX =230
            LabelY =-30
        End
        Begin CheckBox
            SpecialEffect =2
            LabelX =230
            LabelY =-30
        End
        Begin OptionGroup
            SpecialEffect =3
        End
        Begin BoundObjectFrame
            SpecialEffect =2
            OldBorderStyle =0
            BackStyle =0
        End
        Begin TextBox
            FELineBreak = NotDefault
            SpecialEffect =2
            BackColor =-2147483643
            ForeColor =-2147483640
            AsianLineBreak =255
        End
        Begin ListBox
            SpecialEffect =2
            BackColor =-2147483643
            ForeColor =-2147483640
        End
        Begin ComboBox
            SpecialEffect =2
            BackColor =-2147483643
            ForeColor =-2147483640
        End
        Begin Subform
            SpecialEffect =2
        End
        Begin UnboundObjectFrame
            SpecialEffect =2
            OldBorderStyle =1
        End
        Begin ToggleButton
            FontSize =8
            FontWeight =400
            FontName ="MS Sans Serif"
        End
        Begin Tab
            BackStyle =0
        End
        Begin FormHeader
            Height =1260
            BackColor =-2147483633
            Name ="FormHeader"
            Begin
                Begin Label
                    OverlapFlags =93
                    TextAlign =2
                    Left =120
                    Top =960
                    Width =2775
                    Height =240
                    FontWeight =600
                    Name ="Plant_Code_Label"
                    Caption ="Plant Code - Family - Species"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =10980
                    Top =960
                    Width =2505
                    Height =240
                    FontWeight =600
                    Name ="State_Heading"
                    Caption ="Utah PLANT Code - Species"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =5640
                    Top =960
                    Width =2460
                    Height =240
                    FontWeight =600
                    Name ="Add_Synonyms_Label"
                    Caption ="Additional Synonyms"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =8265
                    Top =960
                    Width =2535
                    Height =240
                    FontWeight =600
                    Name ="Taxonomic_Notes_Label"
                    Caption ="Taxonomic Notes"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =3045
                    Top =960
                    Width =2415
                    Height =240
                    FontWeight =600
                    Name ="Master_Common_Name_Label"
                    Caption ="Master Common Name"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    OverlapFlags =87
                    TextAlign =2
                    Left =120
                    Top =720
                    Width =2775
                    Height =240
                    Name ="Label34"
                    Caption ="Master"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =4320
                    Top =60
                    Width =4560
                    Height =420
                    FontSize =14
                    Name ="Label36"
                    Caption ="Species not Present in Park"
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    SpecialEffect =0
                    OldBorderStyle =0
                    OverlapFlags =85
                    BackStyle =0
                    IMESentenceMode =3
                    Left =180
                    Top =180
                    Width =660
                    Name ="Unit_Code"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =0
                            Left =900
                            Top =180
                            Width =1800
                            Height =240
                            FontWeight =400
                            Name ="Unit_Desc"
                            Caption ="Text37:"
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =12000
                    Top =240
                    Width =1020
                    Height =300
                    TabIndex =1
                    Name ="ButtonClose"
                    Caption ="Close Form"
                    OnClick ="[Event Procedure]"
                End
            End
        End
        Begin Section
            Height =720
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =120
                    Top =60
                    Width =960
                    Height =255
                    ColumnWidth =2310
                    Name ="Plant_Code"
                    ControlSource ="Plant_Code"
                    StatusBarText ="Query all species by park"
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1260
                    Top =60
                    Width =1140
                    Height =255
                    ColumnWidth =2310
                    TabIndex =1
                    Name ="Master_Family"
                    ControlSource ="Master_Family"
                    StatusBarText ="Master_Family"
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =120
                    Top =360
                    Width =2760
                    Height =255
                    ColumnWidth =2310
                    TabIndex =2
                    Name ="Master_Species"
                    ControlSource ="Master_Species"
                    StatusBarText ="Master Species (ITIS)"
                End
                Begin TextBox
                    Visible = NotDefault
                    Enabled = NotDefault
                    Locked = NotDefault
                    OverlapFlags =93
                    IMESentenceMode =3
                    Left =10980
                    Top =60
                    Width =959
                    Height =255
                    ColumnWidth =900
                    TabIndex =3
                    Name ="Utah_PLANT_Code"
                    ControlSource ="Utah_PLANT_Code"
                    StatusBarText ="Utah Species PLANTS Code"
                End
                Begin TextBox
                    Visible = NotDefault
                    Enabled = NotDefault
                    Locked = NotDefault
                    OverlapFlags =93
                    IMESentenceMode =3
                    Left =10980
                    Top =360
                    Width =2448
                    Height =255
                    ColumnWidth =2310
                    TabIndex =4
                    Name ="Utah_Species"
                    ControlSource ="Utah_Species"
                    StatusBarText ="Utah Species (Welsh et al 2003)"
                End
                Begin TextBox
                    Visible = NotDefault
                    Enabled = NotDefault
                    Locked = NotDefault
                    OverlapFlags =255
                    IMESentenceMode =3
                    Left =10980
                    Top =60
                    Width =959
                    Height =255
                    ColumnWidth =900
                    TabIndex =5
                    Name ="Co_PLANT_Code"
                    ControlSource ="Co_PLANT_Code"
                    StatusBarText ="Colorado Species PLANTS Code"
                End
                Begin TextBox
                    Visible = NotDefault
                    Enabled = NotDefault
                    Locked = NotDefault
                    OverlapFlags =255
                    IMESentenceMode =3
                    Left =10980
                    Top =360
                    Width =2448
                    Height =255
                    ColumnWidth =2310
                    TabIndex =6
                    Name ="Co_Species"
                    ControlSource ="Co_Species"
                    StatusBarText ="Colorado Species (Weber & Wittmann 2001)"
                End
                Begin TextBox
                    Visible = NotDefault
                    Enabled = NotDefault
                    Locked = NotDefault
                    OverlapFlags =247
                    IMESentenceMode =3
                    Left =10980
                    Top =60
                    Width =959
                    Height =255
                    ColumnWidth =900
                    TabIndex =7
                    Name ="Wy_PLANT_code"
                    ControlSource ="Wy_PLANT_code"
                    StatusBarText ="Wyoming species PLANTS code"
                End
                Begin TextBox
                    Visible = NotDefault
                    Enabled = NotDefault
                    Locked = NotDefault
                    OverlapFlags =247
                    IMESentenceMode =3
                    Left =10980
                    Top =360
                    Width =2448
                    Height =255
                    ColumnWidth =2310
                    TabIndex =8
                    Name ="Wy_Species"
                    ControlSource ="Wy_Species"
                    StatusBarText ="Wyoming Species (Dorn 2001)"
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =5655
                    Top =60
                    Width =2445
                    Height =600
                    ColumnWidth =3000
                    TabIndex =9
                    Name ="Add_Synonyms"
                    ControlSource ="Add_Synonyms"
                    StatusBarText ="Additional Synonyms"
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =8280
                    Top =60
                    Width =2520
                    Height =600
                    ColumnWidth =3000
                    TabIndex =10
                    Name ="Taxonomic_Notes"
                    ControlSource ="Taxonomic_Notes"
                    StatusBarText ="Taxonomic Notes"
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =3060
                    Top =60
                    Width =2400
                    Height =255
                    ColumnWidth =2310
                    TabIndex =11
                    Name ="Master_Common_Name"
                    ControlSource ="Master_Common_Name"
                    StatusBarText ="Master Common Name"
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

Private Sub ButtonClose_Click()
On Error GoTo Err_ButtonClose_Click


    DoCmd.Close

Exit_ButtonClose_Click:
    Exit Sub

Err_ButtonClose_Click:
    MsgBox Err.Description
    Resume Exit_ButtonClose_Click
    
End Sub

Private Sub Form_Load()

  Me!Unit_Code = Me.OpenArgs
  Me!Unit_Desc.Caption = DLookup("[ParkName]", "tlu_Parks", "[ParkCode] = '" & Me.OpenArgs & "'")
  If DLookup("[ParkState]", "tlu_Parks", "[ParkCode] = '" & Me.OpenArgs & "'") = "WY" Then
    Me!Wy_PLANT_code.Visible = True
    Me!Wy_Species.Visible = True
    Me!State_Heading.Caption = "WY PLANT Code - Species"
  ElseIf DLookup("[ParkState]", "tlu_Parks", "[ParkCode] = '" & Me.OpenArgs & "'") = "CO" Then
    Me!Co_PLANT_Code.Visible = True
    Me!Co_Species.Visible = True
    Me!State_Heading.Caption = "CO PLANT Code - Species"
  Else
    Me!State_Heading.Caption = "Utah PLANT Code - Species"
    Me!Utah_PLANT_Code.Visible = True
    Me!Utah_Species.Visible = True
  End If
End Sub
