Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    AllowDeletions = NotDefault
    AllowAdditions = NotDefault
    ScrollBars =2
    TabularFamily =0
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =11520
    DatasheetFontHeight =10
    ItemSuffix =16
    Left =855
    Top =1125
    Right =12375
    Bottom =4590
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0xa7df49ecebfee240
    End
    RecordSource ="qry_List_Versions"
    Caption ="frm_Version_List"
    DatasheetFontName ="Arial"
    Begin
        Begin Label
            BackStyle =0
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
            Height =1140
            BackColor =-2147483633
            Name ="FormHeader"
            Begin
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =300
                    Top =660
                    Width =1320
                    Height =420
                    FontWeight =700
                    Name ="version_key_number_Label"
                    Caption ="Version Key Number"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =1800
                    Top =660
                    Width =1140
                    Height =420
                    FontWeight =700
                    Name ="version_key_date_Label"
                    Caption ="Version Key Date"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =3120
                    Top =660
                    Width =1080
                    Height =420
                    FontWeight =700
                    Name ="narrative_version_Label"
                    Caption ="Narrative Version"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =4380
                    Top =660
                    Width =1020
                    Height =420
                    FontWeight =700
                    Name ="version_comments_Label"
                    Caption ="Version Comments"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    OverlapFlags =85
                    Left =4320
                    Top =60
                    Width =4200
                    Height =450
                    FontSize =18
                    FontWeight =700
                    Name ="Label12"
                    Caption ="Master Version Key List"
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =9120
                    Top =300
                    Width =1020
                    Height =405
                    Name ="ButtonClose"
                    Caption ="Close Form"
                    OnClick ="[Event Procedure]"
                End
            End
        End
        Begin Section
            Height =465
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =93
                    IMESentenceMode =3
                    Left =60
                    Top =60
                    Width =660
                    Height =255
                    ColumnWidth =900
                    Name ="project_ID"
                    ControlSource ="project_ID"
                    StatusBarText ="Project ID number to ensure uniqueness across all projects"
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    OverlapFlags =247
                    IMESentenceMode =3
                    Left =660
                    Top =60
                    Width =600
                    Height =255
                    ColumnWidth =900
                    TabIndex =1
                    Name ="version_key_number"
                    ControlSource ="version_key_number"
                    StatusBarText ="Protocol version key number (maintained in SOP #10)"
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1860
                    Top =60
                    Width =960
                    Height =255
                    ColumnWidth =1035
                    TabIndex =2
                    Name ="version_key_date"
                    ControlSource ="version_key_date"
                    Format ="Short Date"
                    StatusBarText ="Date of protocol version key number"
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    DecimalPlaces =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =3360
                    Top =60
                    Width =540
                    Height =255
                    ColumnWidth =2310
                    TabIndex =3
                    Name ="narrative_version"
                    ControlSource ="narrative_version"
                    Format ="Fixed"
                    StatusBarText ="Version of protocol narrative"
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =93
                    IMESentenceMode =3
                    Left =4380
                    Top =60
                    Width =3240
                    Height =255
                    ColumnWidth =3000
                    TabIndex =4
                    Name ="version_comments"
                    ControlSource ="version_comments"
                    StatusBarText ="Comments regarding version, if any"
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =8160
                    Top =60
                    Width =1410
                    Height =300
                    TabIndex =5
                    Name ="ButtonDetail"
                    Caption ="View SOP Detail"
                    OnClick ="[Event Procedure]"
                End
                Begin CommandButton
                    OverlapFlags =215
                    Left =7620
                    Top =60
                    Width =306
                    Height =306
                    TabIndex =6
                    Name ="ButtonZoom"
                    Caption ="Command13"
                    OnClick ="[Event Procedure]"
                    PictureData = Begin
                        0x280000001000000010000000010004000000000080000000c40e0000c40e0000 ,
                        0x1000000000000000000000000000800000800000008080008000000080008000 ,
                        0x80800000c0c0c000808080000000ff0000ff000000ffff00ff000000ff00ff00 ,
                        0xffff0000ffffff00666666666666446666666666666474466666666666474446 ,
                        0x666666666474446666660000474446666600777f8444666660877777f8086666 ,
                        0x607770777f066666077770777770666607777077777066660700000007706666 ,
                        0x07f770777770666660ff707777066666608ff077780666666600777700666666 ,
                        0x6666000066666666
                    End
                    ObjectPalette = Begin
                        0x0003100000000000800000000080000080800000000080008000800000808000 ,
                        0xc0c0c00080808000ff00000000ff0000ffff00000000ff00ff00ff0000ffff00 ,
                        0xffffff0000000000
                    End
                    ControlTipText ="Close Form"
                    Picture ="C:\\arcgis\\arcexe9x\\odetools\\Bitmaps\\zoomin.bmp"
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =9840
                    Top =60
                    Width =1410
                    Height =300
                    TabIndex =7
                    Name ="ButtonPrint"
                    Caption ="Print Listing"
                    OnClick ="[Event Procedure]"
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

Private Sub ButtonDetail_Click()
On Error GoTo Err_ButtonDetail_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "frm_Show_All_Versions"
    
    stLinkCriteria = "[version_key_number]=" & Me![version_key_number]
    DoCmd.OpenForm stDocName, , , stLinkCriteria

Exit_ButtonDetail_Click:
    Exit Sub

Err_ButtonDetail_Click:
    MsgBox Err.Description
    Resume Exit_ButtonDetail_Click
    
End Sub
Private Sub ButtonZoom_Click()
  Me!version_comments.SetFocus
  SendKeys ("+{F2}")
End Sub
Private Sub ButtonPrint_Click()
On Error GoTo Err_ButtonPrint_Click

    Dim stDocName As String
    Dim strWhere As String

    strWhere = "version_key_number = " & Me![version_key_number]
    stDocName = "rpt_Show_All_Versions"
    DoCmd.OpenReport stDocName, acPreview, , strWhere

Exit_ButtonPrint_Click:
    Exit Sub

Err_ButtonPrint_Click:
    MsgBox Err.Description
    Resume Exit_ButtonPrint_Click
    
End Sub
Private Sub ButtonClose_Click()
On Error GoTo Err_ButtonClose_Click


    DoCmd.Close

Exit_ButtonClose_Click:
    Exit Sub

Err_ButtonClose_Click:
    MsgBox Err.Description
    Resume Exit_ButtonClose_Click
    
End Sub
