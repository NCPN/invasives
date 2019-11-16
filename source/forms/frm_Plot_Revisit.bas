Version =20
VersionRequired =20
Begin Form
    PopUp = NotDefault
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    AllowDeletions = NotDefault
    DividingLines = NotDefault
    AllowAdditions = NotDefault
    DefaultView =0
    ScrollBars =0
    TabularFamily =127
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =10800
    DatasheetFontHeight =9
    ItemSuffix =52
    Left =2550
    Top =3420
    Right =13410
    Bottom =8385
    DatasheetGridlinesColor =12632256
    Filter ="[Location_ID]='20081016093629-468700110.912323'"
    RecSrcDt = Begin
        0x9becc7edac0fe340
    End
    RecordSource ="tbl_Locations"
    DatasheetFontName ="Arial"
    PrtMip = Begin
        0xa0050000a0050000a0050000a005000000000000201c0000e010000001000000 ,
        0x010000006801000000000000a10700000100000001000000
    End
    OnLoad ="[Event Procedure]"
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
        Begin CheckBox
            SpecialEffect =2
            LabelX =230
            LabelY =-30
        End
        Begin TextBox
            FELineBreak = NotDefault
            SpecialEffect =2
            OldBorderStyle =0
            FontName ="Tahoma"
            AsianLineBreak =255
        End
        Begin ComboBox
            SpecialEffect =2
            FontName ="Tahoma"
        End
        Begin Subform
            SpecialEffect =2
        End
        Begin Section
            CanGrow = NotDefault
            Height =5760
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin Label
                    OverlapFlags =93
                    TextAlign =2
                    Left =3600
                    Top =120
                    Width =3435
                    Height =480
                    FontSize =18
                    FontWeight =700
                    Name ="Label0"
                    Caption ="Route Revisit"
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    SpecialEffect =0
                    OverlapFlags =93
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =720
                    Top =360
                    Width =540
                    TabIndex =1
                    Name ="Unit_Code"
                    ControlSource ="Unit_Code"
                    StatusBarText ="Park Code."
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =1
                            Left =180
                            Top =360
                            Width =480
                            Height =240
                            FontWeight =700
                            Name ="Label1"
                            Caption ="Park"
                        End
                    End
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    SpecialEffect =0
                    OverlapFlags =87
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =2400
                    Top =360
                    Width =1200
                    TabIndex =2
                    Name ="Plot_ID"
                    ControlSource ="Plot_ID"
                    StatusBarText ="Plot identifier"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =1560
                            Top =360
                            Width =825
                            Height =240
                            FontWeight =700
                            Name ="Label2"
                            Caption ="Route ID"
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =2940
                    Top =3180
                    Width =1455
                    Height =300
                    TabIndex =3
                    Name ="ButtonClose"
                    Caption ="Cancel New Visit"
                    OnClick ="[Event Procedure]"
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =87
                    IMESentenceMode =3
                    Left =180
                    Top =120
                    Width =720
                    TabIndex =4
                    Name ="Location_ID"
                    ControlSource ="Location_ID"
                    StatusBarText ="M. Location identifier (Loc_ID)"
                End
                Begin Subform
                    OverlapFlags =85
                    Left =1380
                    Top =780
                    Width =7410
                    Height =1560
                    Name ="fsub_Revisit"
                    SourceObject ="Form.fsub_Revisit"
                    LinkChildFields ="Location_ID"
                    LinkMasterFields ="Location_ID"
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =5820
                    Top =3180
                    Width =1454
                    Height =299
                    TabIndex =5
                    Name ="ButtonContinue"
                    Caption ="Continue"
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

    Dim stDocName As String
    Dim db As Database
    Dim Events As DAO.Recordset
    Dim strSQL As String
  If Not IsNull(Me!fsub_Revisit.Form!Event_ID) Then
    strSQL = "Select * FROM tbl_events WHERE event_ID = '" & Me!fsub_Revisit.Form!Event_ID & "'"
    Set db = CurrentDb
  ' Get the added events record
    Set Events = db.OpenRecordset(strSQL)
    If Not Events.EOF Then
      Events.Delete
'      Delete cancelled events record
      Events.Close
    End If
  End If
    DoCmd.Close

Exit_ButtonClose_Click:
    Exit Sub

Err_ButtonClose_Click:
    MsgBox Err.Description
    Resume Exit_ButtonClose_Click
    
End Sub

Private Sub Form_Load()
  
    DoCmd.Restore

End Sub
Private Sub ButtonContinue_Click()
On Error GoTo Err_ButtonContinue_Click

    Dim stDocName As String
    Dim stLinkCriteria As String
    
    If IsNull(Me!fsub_Revisit.Form!Start_Date) Then
      MsgBox "Revisit date required"
      GoTo Exit_ButtonContinue_Click
    Else
    DoCmd.RunCommand acCmdSaveRecord  ' Save the new event record
    End If
    stDocName = "frm_Data_Entry"
    stLinkCriteria = "[Event_ID]=" & "'" & Me!fsub_Revisit.Form!Event_ID & "'"
    DoCmd.OpenForm stDocName, , , stLinkCriteria
        If Not IsNull(Me!Location_ID) Then
            ' Fill in Location
            Forms!frm_Data_Entry!cboLocation_ID = Me!Location_ID
            Forms!frm_Data_Entry.Update_Loc_Info
        End If
    DoCmd.Close acForm, "frm_Plot_Revisit"
Exit_ButtonContinue_Click:
    Exit Sub

Err_ButtonContinue_Click:
    MsgBox Err.Description
    Resume Exit_ButtonContinue_Click
    
End Sub
