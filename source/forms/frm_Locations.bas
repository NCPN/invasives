Version =20
VersionRequired =20
Begin Form
    PopUp = NotDefault
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    DividingLines = NotDefault
    FilterOn = NotDefault
    DefaultView =0
    ScrollBars =0
    TabularFamily =0
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    Cycle =1
    GridX =24
    GridY =24
    Width =6480
    DatasheetFontHeight =9
    ItemSuffix =128
    Left =4830
    Top =3105
    Right =11070
    Bottom =6540
    DatasheetGridlinesColor =12632256
    Filter ="[Location_ID]='20101019114804-533424019.813538'"
    RecSrcDt = Begin
        0xdca6db037508e340
    End
    RecordSource ="tbl_Locations"
    Caption =" Locations"
    OnCurrent ="[Event Procedure]"
    BeforeUpdate ="[Event Procedure]"
    OnClose ="[Event Procedure]"
    DatasheetFontName ="Arial"
    PrtMip = Begin
        0xa0050000a0050000a0050000a005000000000000201c0000e010000001000000 ,
        0x010000006801000000000000a10700000100000001000000
    End
    OnLoad ="[Event Procedure]"
    FilterOnLoad =255
    AllowLayoutView =0
    DatasheetGridlinesColor12 =12632256
    Begin
        Begin Label
            BackStyle =0
            FontSize =18
        End
        Begin CommandButton
            FontSize =8
            FontWeight =400
            ForeColor =-2147483630
            FontName ="Tahoma"
            BorderLineStyle =0
        End
        Begin CheckBox
            SpecialEffect =2
            BorderLineStyle =0
            LabelX =230
            LabelY =-30
        End
        Begin TextBox
            FELineBreak = NotDefault
            SpecialEffect =2
            OldBorderStyle =0
            BorderLineStyle =0
            FontName ="Tahoma"
            AsianLineBreak =255
        End
        Begin ComboBox
            SpecialEffect =2
            BorderLineStyle =0
            FontName ="Tahoma"
        End
        Begin Subform
            SpecialEffect =2
            BorderLineStyle =0
        End
        Begin UnboundObjectFrame
            SpecialEffect =2
            OldBorderStyle =1
        End
        Begin Section
            CanGrow = NotDefault
            Height =3600
            BackColor =12574431
            Name ="Detail"
            Begin
                Begin TextBox
                    Visible = NotDefault
                    Locked = NotDefault
                    TabStop = NotDefault
                    SpecialEffect =3
                    OldBorderStyle =1
                    OverlapFlags =85
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =60
                    Top =60
                    Width =948
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="txtLocation_ID"
                    ControlSource ="Location_ID"
                    StatusBarText ="Unique identifier for each sample location"

                End
                Begin ComboBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =2
                    ListWidth =3150
                    Left =1140
                    Top =720
                    Width =960
                    TabIndex =1
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"10\";\"10\""
                    Name ="Unit_Code"
                    ControlSource ="Unit_Code"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT tlu_Parks.ParkCode, tlu_Parks.ParkName FROM tlu_Parks; "
                    ColumnWidths ="585;2565"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =180
                            Top =720
                            Width =900
                            Height =240
                            FontSize =8
                            FontWeight =700
                            Name ="ParkCode_Label"
                            Caption ="Park Code"
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =1800
                    Top =120
                    Width =2880
                    Height =420
                    FontSize =16
                    FontWeight =700
                    Name ="Label11"
                    Caption ="Edit/Add Routes"
                    FontName ="Tahoma"
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =3000
                    Top =720
                    Width =2400
                    TabIndex =2
                    Name ="PlotID"
                    ControlSource ="Plot_ID"
                    StatusBarText ="Plot identifier"
                    AfterUpdate ="[Event Procedure]"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2280
                            Top =720
                            Width =660
                            Height =240
                            FontSize =8
                            FontWeight =700
                            Name ="Label12"
                            Caption ="Route"
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin ComboBox
                    OverlapFlags =87
                    IMESentenceMode =3
                    ColumnCount =2
                    ListWidth =1440
                    Left =1020
                    Top =1200
                    Width =1980
                    TabIndex =3
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"10\";\"100\""
                    Name ="Recorder"
                    ControlSource ="Observer"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT tlu_Contacts.Contact_ID, tlu_Contacts.Last_Name FROM tlu_Contacts; "
                    ColumnWidths ="0;1440"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =180
                            Top =1200
                            Width =840
                            Height =245
                            FontSize =8
                            FontWeight =700
                            Name ="Recorder_Label"
                            Caption ="Recorder"
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =2520
                    Top =2280
                    Width =1035
                    Height =300
                    TabIndex =4
                    Name ="ButtonClose"
                    Caption ="Close Form"
                    OnClick ="[Event Procedure]"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =120
                    Top =3240
                    TabIndex =5
                    Name ="txtUpdated_Date"
                    ControlSource ="Updated_Date"
                    Format ="Short Date"
                    StatusBarText ="MA. Date of entry or last change (Upd_Date)"

                End
                Begin TextBox
                    OverlapFlags =87
                    IMESentenceMode =3
                    Left =720
                    Top =1680
                    Width =3600
                    TabIndex =6
                    Name ="area"
                    ControlSource ="area"

                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =180
                            Top =1680
                            Width =540
                            Height =240
                            FontSize =8
                            FontWeight =700
                            Name ="Label127"
                            Caption ="Area"
                            FontName ="Tahoma"
                        End
                    End
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
Option Explicit


' =================================
' Description:  Locations entry form
' Data source:  tbl_Locations
' Data access:  edit, add, delete
' Pages:        none
' Functions:    none
' References:   fxnGUIDGen
' Source/date:  Rescued from Simon Kingston, Sept. 2006 by Russ DenBleyker
' Revisions:    <name, date, desc - add lines as you go>
' =================================


Private Sub Form_BeforeUpdate(Cancel As Integer)
'check to see if a primary key is needed and add it (used for string GUIDs)

    Me!txtUpdated_Date = Now()
    If IsNull(Me!txtLocation_ID) Then
        If GetDataType("tbl_Locations", "Location_ID") = dbText Then
            Me!txtLocation_ID = fxnGUIDGen
        End If
    End If

End Sub

Private Sub Form_Close()
'update control as necessary on calling form to reflect new location values
UpdateControl Me.OpenArgs
End Sub

Private Sub Form_Current()
'check to see if a primary key is needed and add it (used for string GUIDs)
  If Me.NewRecord Then
    If GetDataType("tbl_Locations", "Location_ID") = dbText Then
        Me!txtLocation_ID = fxnGUIDGen
    End If
  End If
End Sub


Private Sub Form_Load()
  Dim strCriteria As String
  If Me.OpenArgs = "New Record" Then
    strCriteria = "Project = 'NCPN Invasive Monitoring'"
    MsgBox strCriteria
    MsgBox DLookup("Park", "tsys_App_Defaults", strCriteria)
  '  Me!Unit_Code = DLookup("Park", "tsys_App_Defaults", strCriteria)
  End If

End Sub

Private Sub PlotID_AfterUpdate()
  Dim strCriteria As String
  Dim strMessage As String
  
  strCriteria = "Unit_Code = '" & Me!Unit_Code & "' And Plot_Id = '" & Me!Plot_ID & "'"
  If Not IsNull(DLookup("Location_ID", "tbl_Locations", strCriteria)) Then
    strMessage = "Plot " & Me!Plot_ID & " already exists for " & Me!Unit_Code & "."
    MsgBox strMessage
    DoCmd.CancelEvent
    SendKeys "{ESC}"
    Me!Unit_Code.SetFocus
    If Me.OpenArgs = "New Record" Then
      strCriteria = "Project = 'NCPN Invasive Monitoring'"
      Me!Unit_Code = DLookup("Park", "tsys_App_Defaults", strCriteria)
    End If
  End If

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
