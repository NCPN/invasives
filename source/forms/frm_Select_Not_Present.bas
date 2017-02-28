Version =20
VersionRequired =20
Begin Form
    PopUp = NotDefault
    RecordSelectors = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    DefaultView =0
    ScrollBars =0
    TabularFamily =0
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =4320
    DatasheetFontHeight =9
    ItemSuffix =4
    Left =4080
    Top =4515
    Right =8190
    Bottom =5550
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x9d2210c6b41ee340
    End
    Caption ="Select for Not in Park Report"
    DatasheetFontName ="Arial"
    Begin
        Begin Label
            BackStyle =0
            FontName ="Tahoma"
        End
        Begin CommandButton
            FontSize =8
            FontWeight =400
            ForeColor =-2147483630
            FontName ="Tahoma"
        End
        Begin ComboBox
            SpecialEffect =2
            FontName ="Tahoma"
        End
        Begin Section
            Height =2880
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin ComboBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =2
                    ListWidth =3600
                    Left =2280
                    Top =1320
                    Width =840
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"10\";\"10\""
                    Name ="Park_Code"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT tlu_Parks.ParkCode, tlu_Parks.ParkName FROM tlu_Parks; "
                    ColumnWidths ="720;2880"
                    AfterUpdate ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =960
                            Top =1320
                            Width =1260
                            Height =245
                            FontWeight =700
                            Name ="Select a Park_Label"
                            Caption ="Select a Park"
                            EventProcPrefix ="Select_a_Park_Label"
                        End
                    End
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =180
                    Top =420
                    Width =3540
                    Height =360
                    FontSize =12
                    FontWeight =700
                    Name ="Label2"
                    Caption ="Not Present in Park Report"
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =1560
                    Top =1980
                    Width =1035
                    Height =300
                    TabIndex =1
                    Name ="ButtonClose"
                    Caption ="Close Form"
                    OnClick ="[Event Procedure]"
                End
            End
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

Private Sub Park_Code_AfterUpdate()
  Dim strfrmName As String
  Dim strWhere As String
  
  strfrmName = "frm_Not_Present"
  strWhere = "Unit_Code = '" & Me!Park_Code & "' AND " & Me!Park_Code & " <> 'Present'"
  DoCmd.OpenForm strfrmName, , , strWhere, , , Me!Park_Code
  DoCmd.Close acForm, "frm_Select_Not_Present"
  
End Sub
