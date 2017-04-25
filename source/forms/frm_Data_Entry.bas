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
    FilterOn = NotDefault
    DefaultView =0
    ScrollBars =0
    ViewsAllowed =1
    TabularFamily =0
    BorderStyle =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    Cycle =1
    GridX =24
    GridY =24
    Width =14280
    DatasheetFontHeight =10
    ItemSuffix =200
    Left =4080
    Top =2610
    Right =18360
    Bottom =14355
    DatasheetGridlinesColor =12632256
    Filter ="[Location_ID]='20121121125102-961953163.146973' AND [Event_ID]='20151107154044-7"
        "60723590.85083'"
    RecSrcDt = Begin
        0xaf0254819108e340
    End
    RecordSource ="qfrm_DataEntry"
    Caption =" Data Entry Form - Filter by sampling event"
    OnCurrent ="[Event Procedure]"
    BeforeInsert ="[Event Procedure]"
    BeforeUpdate ="[Event Procedure]"
    OnOpen ="[Event Procedure]"
    OnClose ="[Event Procedure]"
    DatasheetFontName ="Arial"
    PrtMip = Begin
        0xa0050000a0050000a0050000a005000000000000201c0000e010000001000000 ,
        0x010000006801000000000000a10700000100000001000000
    End
    AllowDatasheetView =0
    AllowPivotTableView =0
    AllowPivotChartView =0
    AllowPivotChartView =0
    FilterOnLoad =255
    AllowLayoutView =0
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
        Begin Section
            CanGrow = NotDefault
            Height =11160
            BackColor =12574431
            Name ="Detail"
            Begin
                Begin CommandButton
                    TabStop = NotDefault
                    OverlapFlags =85
                    Left =6240
                    Top =480
                    Width =1080
                    Height =479
                    FontSize =9
                    FontWeight =700
                    TabIndex =11
                    Name ="cmdClose"
                    Caption ="Close"
                    OnClick ="[Event Procedure]"
                    FontName ="Arial"
                    ControlTipText ="Close the data entry form"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin ComboBox
                    ColumnHeads = NotDefault
                    LimitToList = NotDefault
                    Visible = NotDefault
                    Locked = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =5
                    ListWidth =7488
                    Left =7980
                    Top =900
                    Width =768
                    ColumnWidth =1440
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"\";\"\";\"\";\"\";\"\";\"\";\"10\";\"510\""
                    Name ="cboLocation_ID"
                    ControlSource ="Location_ID"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT tbl_Locations.Location_ID, tbl_Locations.Plot_ID, tbl_Locations.Unit_Code"
                        ", tbl_Locations.E_Coord, tbl_Locations.N_Coord FROM tbl_Locations ORDER BY tbl_L"
                        "ocations.Unit_Code, tbl_Locations.Plot_ID, tbl_Locations.E_Coord, tbl_Locations."
                        "N_Coord; "
                    ColumnWidths ="0;3456;1152;1440;1440"
                    StatusBarText ="Unique identifier for each sample location"
                    AfterUpdate ="[Event Procedure]"
                    FontName ="Arial"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =8760
                            Top =840
                            Width =780
                            Height =255
                            FontWeight =700
                            Name ="labLocation_ID"
                            Caption ="Site"
                            FontName ="Arial"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =93
                    TextAlign =1
                    IMESentenceMode =3
                    Left =1140
                    Top =900
                    Width =1080
                    TabIndex =3
                    Name ="txtStart_Date"
                    ControlSource ="Start_Date"
                    Format ="Short Date"
                    StatusBarText ="M. Starting date for the event (Start_Date)"
                    AfterUpdate ="[Event Procedure]"
                    InputMask ="99/99/0000;0;_"

                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =120
                            Top =900
                            Width =960
                            Height =240
                            FontWeight =700
                            Name ="Label55"
                            Caption ="Start Date"
                        End
                    End
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    TabStop = NotDefault
                    SpecialEffect =0
                    OldBorderStyle =0
                    OverlapFlags =93
                    BackStyle =0
                    IMESentenceMode =3
                    Left =1140
                    Top =120
                    Width =840
                    TabIndex =5
                    Name ="txtUnit_Code"

                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =120
                            Top =120
                            Width =990
                            Height =240
                            FontWeight =700
                            Name ="Label60"
                            Caption ="Park"
                        End
                    End
                End
                Begin Rectangle
                    OverlapFlags =255
                    Left =60
                    Top =60
                    Width =5760
                    Name ="Box62"
                End
                Begin Tab
                    MultiRow = NotDefault
                    OverlapFlags =85
                    Left =45
                    Top =1260
                    Width =14130
                    Height =9795
                    TabIndex =6
                    Name ="pgTabs"

                    LayoutCachedLeft =45
                    LayoutCachedTop =1260
                    LayoutCachedWidth =14175
                    LayoutCachedHeight =11055
                    Begin
                        Begin Page
                            OverlapFlags =87
                            Left =180
                            Top =1665
                            Width =13860
                            Height =9255
                            Name ="pgCoords_and_loc_details"
                            Caption ="Monitoring Transects"
                            LayoutCachedLeft =180
                            LayoutCachedTop =1665
                            LayoutCachedWidth =14040
                            LayoutCachedHeight =10920
                            WebImagePaddingLeft =2
                            WebImagePaddingTop =2
                            WebImagePaddingRight =2
                            WebImagePaddingBottom =2
                            Begin
                                Begin Subform
                                    OverlapFlags =215
                                    Left =180
                                    Top =1665
                                    Width =13710
                                    Height =9075
                                    Name ="frm_Quadrat_Transect"
                                    SourceObject ="Form.frm_Quadrat_Transect"
                                    LinkChildFields ="Event_ID"
                                    LinkMasterFields ="Event_ID"

                                    LayoutCachedLeft =180
                                    LayoutCachedTop =1665
                                    LayoutCachedWidth =13890
                                    LayoutCachedHeight =10740
                                End
                            End
                        End
                    End
                End
                Begin Rectangle
                    OverlapFlags =255
                    Left =60
                    Top =780
                    Width =5760
                    Height =480
                    Name ="Box65"
                End
                Begin TextBox
                    Visible = NotDefault
                    Enabled = NotDefault
                    Locked = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =7560
                    Top =540
                    Width =300
                    TabIndex =7
                    Name ="txtLocation_ID"
                    ControlSource ="Location_ID"
                    StatusBarText ="M. Link to tbl_Locations (Loc_ID)"

                End
                Begin CheckBox
                    Visible = NotDefault
                    OverlapFlags =85
                    Left =7560
                    Top =900
                    TabIndex =8
                    Name ="Site_Selection"
                    ControlSource ="Site_Selection"
                    StatusBarText ="Site accepted or rejected"

                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =8100
                    Top =540
                    Width =540
                    TabIndex =9
                    Name ="version_key_number"
                    ControlSource ="version_key_number"
                    StatusBarText ="Master protocol version key"

                End
                Begin Label
                    OverlapFlags =247
                    Left =2220
                    Top =120
                    Width =600
                    Height =240
                    FontWeight =700
                    Name ="Label91"
                    Caption ="Route"
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    SpecialEffect =0
                    OldBorderStyle =0
                    OverlapFlags =247
                    BackStyle =0
                    IMESentenceMode =3
                    Left =2880
                    Top =120
                    Width =2820
                    TabIndex =10
                    Name ="SiteDisplay"
                    ControlSource ="Plot_ID"

                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =215
                    IMESentenceMode =3
                    Left =8880
                    Top =180
                    Width =4800
                    Height =899
                    TabIndex =1
                    Name ="Comments"
                    ControlSource ="Comments"
                    StatusBarText ="Plot revisit comments."

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =7440
                            Top =180
                            Width =1380
                            Height =240
                            FontWeight =700
                            Name ="Label94"
                            Caption ="Visit Comments:"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =247
                    IMESentenceMode =3
                    ColumnCount =3
                    ListWidth =1650
                    Left =960
                    Top =480
                    Width =2100
                    TabIndex =2
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"\";\"\";\"10\";\"100\""
                    Name ="Observer"
                    ControlSource ="Observer"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT tlu_Contacts.Contact_ID, tlu_Contacts.Last_Name, tlu_Contacts.First_Name "
                        "FROM tlu_Contacts; "
                    ColumnWidths ="0;810;839"
                    Begin
                        Begin Label
                            OverlapFlags =255
                            Left =120
                            Top =480
                            Width =840
                            Height =245
                            FontWeight =700
                            Name ="Observer_Label"
                            Caption ="Observer"
                        End
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =247
                    IMESentenceMode =3
                    Left =4980
                    Top =960
                    Width =360
                    TabIndex =12
                    Name ="Unit_Code"
                    ControlSource ="Unit_Code"
                    StatusBarText ="Park Code."

                End
                Begin TextBox
                    OverlapFlags =247
                    TextAlign =1
                    IMESentenceMode =3
                    Left =3480
                    Top =900
                    Width =1080
                    TabIndex =4
                    Name ="Start_Time"
                    ControlSource ="Start_Time"
                    Format ="Short Time"
                    StatusBarText ="M. Starting date for the event (Start_Date)"
                    InputMask ="00:00;0;_"

                    Begin
                        Begin Label
                            OverlapFlags =247
                            Left =2460
                            Top =900
                            Width =960
                            Height =240
                            FontWeight =700
                            Name ="Label199"
                            Caption ="GPS Time"
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
' FORM NAME:    frm_Data_Entry
' Description:  Primary field data entry form
' Data source:  tbl_Locations
' Data access:  edit; allow additions off except for new records
' Pages:        none
' Functions:    none
' References:   fxnSwitchboardIsOpen, fxnGUIDGen
' Source/date:  John R. Boetsch, June 2006
' Revisions:    <name, date, desc - add lines as you go>
' =================================

Private Sub cboLocation_ID_AfterUpdate()
' Update_Loc_Info
End Sub

Private Sub Form_BeforeUpdate(Cancel As Integer)
  If IsNull(Me!Start_Date) Then
        ' ask user if (s)he wants to enter data or cancel and close form
        If MsgBox("Visit date is missing - do you want to enter the missing data?", vbYesNo, "Date missing") = vbNo Then
            Me.Undo
        End If
  End If
End Sub

Private Sub Form_Close()
If IsLoaded("frm_Data_Gateway") Then
    Forms("frm_Data_Gateway").Requery
End If
End Sub

Private Sub Form_Current()

Update_Loc_Info

End Sub

Private Sub Form_Open(Cancel As Integer)
    On Error GoTo Err_Handler

    Dim strCaptionSuffix As String

    ' Set the opening parameters depending on the arguments passed from the previous form
    If Me.OpenArgs = "New record" Or Me.OpenArgs = "Filter by location" Then
        strCaptionSuffix = " - " & Me.OpenArgs
    ElseIf Me.OpenArgs = "New event" Then
        strCaptionSuffix = " - " & Me.OpenArgs
    ElseIf Me.OpenArgs <> "" Then
        strCaptionSuffix = " - " & "Filter by sampling event"
    End If
    Me.Caption = Me.Caption & strCaptionSuffix
    Me!txtStart_date.SetFocus
Exit_Procedure:
    Exit Sub

Err_Handler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    Resume Exit_Procedure

End Sub

Private Sub Form_BeforeInsert(Cancel As Integer)

        Dim db As DAO.Database
        Dim Versions As DAO.Recordset
        Dim strSQL As String
        
    On Error GoTo Err_Handler
    
    ' Set master version number on event record
    Set db = CurrentDb
    strSQL = "SELECT [version_key_number] FROM [tbl_master_version] ORDER BY [version_key_number] DESC"
    Set Versions = db.OpenRecordset(strSQL)
    Versions.MoveFirst
    Me![version_key_number] = Versions![version_key_number]
    Versions.Close

    ' Create the GUID primary key value
    If IsNull(Me!Event_ID) Then
        If GetDataType("tbl_Events", "Event_ID") = dbText Then
            Me.Event_ID = fxnGUIDGen
        End If
    End If

Exit_Procedure:
    Exit Sub

Err_Handler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    Resume Exit_Procedure

End Sub

Private Sub cmdClose_Click()
    On Error GoTo Err_Handler

    DoCmd.RunCommand acCmdSaveRecord
    DoCmd.Close , , acSaveNo

Exit_Procedure:
    Exit Sub

Err_Handler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    Resume Exit_Procedure

End Sub


Public Sub Update_Loc_Info()
' Description:  Updates associated location information when Location_ID is updated
' References:   GetCriteriaString
' Source/date:  Simon Kingston, Sept. 2006
' Revisions:    <name, date, desc - add lines as you go>

Dim strXY As Variant
Dim strCriteria As String

If IsNull(Me!txtLocation_ID) Then
    Me!txtUnit_Code = Null
Else
    strCriteria = GetCriteriaString("Location_ID=", "tbl_Locations", "Location_ID", Me.Name, "txtLocation_ID")
    Me!txtUnit_Code = DLookup("Unit_Code", "tbl_Locations", strCriteria)
End If
End Sub


Private Sub txtStart_Date_AfterUpdate()
        Dim db As DAO.Database
        Dim Events As DAO.Recordset
        Dim strSQL As String
        
    On Error GoTo Err_Handler
    
    ' Check for duplicate date
    strSQL = "SELECT Event_ID FROM tbl_Events WHERE [Location_ID] = '" & Me!cboLocation_ID & "' AND [Start_Date] = #" & Me!Start_Date & "#"
'    MsgBox strSQL
    Set db = CurrentDb
    Set Events = db.OpenRecordset(strSQL)
    If Not Events.EOF Then
      MsgBox " Duplicate visit date - update cancelled."
      Me.Undo
      Events.Close
      DoCmd.Close
      GoTo Exit_Procedure
    End If
    Events.Close
Exit_Procedure:
    Exit Sub

Err_Handler:
    MsgBox Err.Description
    Resume Exit_Procedure
End Sub
