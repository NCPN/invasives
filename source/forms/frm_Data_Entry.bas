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
    Width =14160
    DatasheetFontHeight =10
    ItemSuffix =206
    Left =2895
    Top =915
    Right =17055
    Bottom =12360
    DatasheetGridlinesColor =12632256
    Filter ="[Location_ID]='20101021094935-986093163.490295' AND [Event_ID]='20170701144612-1"
        "28249883.651733'"
    RecSrcDt = Begin
        0xc0d562cf56f6e440
    End
    RecordSource ="qfrm_Data_Entry"
    Caption =" Data Entry Form - Filter by sampling event - Filter by sampling event - Filter "
        "by sampling event - Filter by sampling event - Filter by sampling event - Filter"
        " by sampling event - Filter by sampling event"
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
            Height =11460
            BackColor =12574431
            Name ="Detail"
            Begin
                Begin CommandButton
                    TabStop = NotDefault
                    OverlapFlags =85
                    Left =6300
                    Top =240
                    Width =1080
                    Height =479
                    FontSize =9
                    FontWeight =700
                    TabIndex =9
                    Name ="btnClose"
                    Caption ="Close"
                    OnClick ="[Event Procedure]"
                    FontName ="Arial"
                    ControlTipText ="Close the data entry form"
                    LeftPadding =60
                    TopPadding =45
                    RightPadding =150
                    BottomPadding =150

                    LayoutCachedLeft =6300
                    LayoutCachedTop =240
                    LayoutCachedWidth =7380
                    LayoutCachedHeight =719
                    ForeThemeColorIndex =0
                    UseTheme =1
                    BackColor =15921906
                    BackThemeColorIndex =1
                    BackShade =95.0
                    BorderThemeColorIndex =0
                    HoverColor =9434577
                    PressedColor =0
                    PressedThemeColorIndex =0
                    PressedShade =80.0
                    HoverForeColor =16711680
                    PressedForeColor =0
                    PressedForeThemeColorIndex =0
                    QuickStyle =22
                    QuickStyleMask =-53
                    WebImagePaddingLeft =4
                    WebImagePaddingTop =3
                    WebImagePaddingRight =9
                    WebImagePaddingBottom =9
                    Overlaps =1
                End
                Begin ComboBox
                    ColumnHeads = NotDefault
                    LimitToList = NotDefault
                    Visible = NotDefault
                    Locked = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =93
                    IMESentenceMode =3
                    ColumnCount =5
                    ListWidth =7488
                    Left =7980
                    Top =900
                    Width =768
                    ColumnWidth =1440
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"\";\"\";\"\";\"\";\"\";\"\";\"10\";\"510\""
                    Name ="cbxLocationID"
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
                    LayoutCachedLeft =7980
                    LayoutCachedTop =900
                    LayoutCachedWidth =8748
                    LayoutCachedHeight =1140
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =8820
                            Top =900
                            Width =780
                            Height =255
                            FontWeight =700
                            Name ="lblLocation_ID"
                            Caption ="Site"
                            FontName ="Arial"
                            LayoutCachedLeft =8820
                            LayoutCachedTop =900
                            LayoutCachedWidth =9600
                            LayoutCachedHeight =1155
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
                    Left =1440
                    Top =360
                    Width =840
                    TabIndex =3
                    Name ="txtUnit_Code"

                    LayoutCachedLeft =1440
                    LayoutCachedTop =360
                    LayoutCachedWidth =2280
                    LayoutCachedHeight =600
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =420
                            Top =360
                            Width =990
                            Height =240
                            FontWeight =700
                            Name ="lblPark"
                            Caption ="Park"
                            LayoutCachedLeft =420
                            LayoutCachedTop =360
                            LayoutCachedWidth =1410
                            LayoutCachedHeight =600
                        End
                    End
                End
                Begin Tab
                    MultiRow = NotDefault
                    OverlapFlags =85
                    BackStyle =1
                    Left =60
                    Top =1560
                    Width =14040
                    Height =9795
                    FontWeight =600
                    TabIndex =4
                    Name ="pgTabs"

                    LayoutCachedLeft =60
                    LayoutCachedTop =1560
                    LayoutCachedWidth =14100
                    LayoutCachedHeight =11355
                    UseTheme =255
                    BackColor =26112
                    OldBorderStyle =0
                    HoverColor =65280
                    PressedColor =3114797
                    HoverForeColor =16777181
                    PressedForeColor =16777181
                    ForeColor =16777215
                    Begin
                        Begin Page
                            OverlapFlags =87
                            Left =135
                            Top =1965
                            Width =13890
                            Height =9315
                            Name ="pgCoords_and_loc_details"
                            Caption ="Monitoring Transects"
                            LayoutCachedLeft =135
                            LayoutCachedTop =1965
                            LayoutCachedWidth =14025
                            LayoutCachedHeight =11280
                            WebImagePaddingLeft =2
                            WebImagePaddingTop =2
                            WebImagePaddingRight =2
                            WebImagePaddingBottom =2
                            Begin
                                Begin Subform
                                    OverlapFlags =215
                                    Left =225
                                    Top =2100
                                    Width =13710
                                    Height =9075
                                    Name ="frm_Quadrat_Transect"
                                    SourceObject ="Form.frm_Quadrat_Transect"
                                    LinkChildFields ="Event_ID"
                                    LinkMasterFields ="Event_ID"

                                    LayoutCachedLeft =225
                                    LayoutCachedTop =2100
                                    LayoutCachedWidth =13935
                                    LayoutCachedHeight =11175
                                End
                            End
                        End
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    Enabled = NotDefault
                    Locked = NotDefault
                    OverlapFlags =93
                    IMESentenceMode =3
                    Left =7560
                    Top =540
                    Width =300
                    TabIndex =5
                    Name ="txtLocation_ID"
                    ControlSource ="Location_ID"
                    StatusBarText ="M. Link to tbl_Locations (Loc_ID)"

                    LayoutCachedLeft =7560
                    LayoutCachedTop =540
                    LayoutCachedWidth =7860
                    LayoutCachedHeight =780
                End
                Begin CheckBox
                    Visible = NotDefault
                    OverlapFlags =85
                    Left =7560
                    Top =900
                    TabIndex =6
                    Name ="Site_Selection"
                    ControlSource ="Site_Selection"
                    StatusBarText ="Site accepted or rejected"

                    LayoutCachedLeft =7560
                    LayoutCachedTop =900
                    LayoutCachedWidth =7820
                    LayoutCachedHeight =1140
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =93
                    IMESentenceMode =3
                    Left =8100
                    Top =540
                    Width =540
                    TabIndex =7
                    Name ="version_key_number"
                    ControlSource ="version_key_number"
                    StatusBarText ="Master protocol version key"

                    LayoutCachedLeft =8100
                    LayoutCachedTop =540
                    LayoutCachedWidth =8640
                    LayoutCachedHeight =780
                End
                Begin Label
                    OverlapFlags =93
                    Left =2520
                    Top =360
                    Width =600
                    Height =240
                    FontWeight =700
                    Name ="lblRoute"
                    Caption ="Route"
                    LayoutCachedLeft =2520
                    LayoutCachedTop =360
                    LayoutCachedWidth =3120
                    LayoutCachedHeight =600
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    SpecialEffect =0
                    OldBorderStyle =0
                    OverlapFlags =93
                    BackStyle =0
                    IMESentenceMode =3
                    Left =3180
                    Top =360
                    Width =2820
                    TabIndex =8
                    Name ="SiteDisplay"
                    ControlSource ="Plot_ID"

                    LayoutCachedLeft =3180
                    LayoutCachedTop =360
                    LayoutCachedWidth =6000
                    LayoutCachedHeight =600
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =247
                    IMESentenceMode =3
                    Left =7860
                    Top =420
                    Width =5100
                    Height =1079
                    TabIndex =1
                    Name ="Comments"
                    ControlSource ="Comments"
                    StatusBarText ="Plot revisit comments."

                    LayoutCachedLeft =7860
                    LayoutCachedTop =420
                    LayoutCachedWidth =12960
                    LayoutCachedHeight =1499
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =7560
                            Top =120
                            Width =1380
                            Height =240
                            FontWeight =700
                            Name ="lblVisitComments"
                            Caption ="Visit Comments:"
                            LayoutCachedLeft =7560
                            LayoutCachedTop =120
                            LayoutCachedWidth =8940
                            LayoutCachedHeight =360
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =95
                    IMESentenceMode =3
                    ColumnCount =3
                    ListWidth =1650
                    Left =4800
                    Top =1020
                    Width =2100
                    TabIndex =2
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"\";\"\";\"10\";\"100\""
                    Name ="Observer"
                    ControlSource ="Observer"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT tlu_Contacts.Contact_ID, tlu_Contacts.Last_Name, tlu_Contacts.First_Name "
                        "FROM tlu_Contacts; "
                    ColumnWidths ="0;810;839"
                    ControlTipText ="Sampling visit observer"
                    LayoutCachedLeft =4800
                    LayoutCachedTop =1020
                    LayoutCachedWidth =6900
                    LayoutCachedHeight =1260
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =3960
                            Top =1020
                            Width =840
                            Height =245
                            FontWeight =700
                            Name ="lblObserver"
                            Caption ="Observer"
                            LayoutCachedLeft =3960
                            LayoutCachedTop =1020
                            LayoutCachedWidth =4800
                            LayoutCachedHeight =1265
                        End
                    End
                End
                Begin TextBox
                    SpecialEffect =0
                    OldBorderStyle =0
                    OverlapFlags =215
                    BackStyle =0
                    IMESentenceMode =3
                    Left =2160
                    Top =1560
                    Width =2820
                    TabIndex =10
                    ForeColor =5855577
                    Name ="tbxEventID"
                    ControlSource ="Event_ID"
                    ConditionalFormat = Begin
                        0x0100000012010000010000000100000000000000000000005800000001000000 ,
                        0xececec00ffffff00000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x49004900660028005b00660072006d005f005100750061006400720061007400 ,
                        0x5f005400720061006e0073006500630074005d002e0046006f0072006d002100 ,
                        0x5b0066007300750062005f0053007000650063006900650073005f0043007500 ,
                        0x7200720065006e0074005d002e0046006f0072006d002e0043006f006e007400 ,
                        0x72006f006c007300280022007400620078004400650076004d006f0064006500 ,
                        0x220029002c0030002c003100290000000000
                    End

                    LayoutCachedLeft =2160
                    LayoutCachedTop =1560
                    LayoutCachedWidth =4980
                    LayoutCachedHeight =1800
                    ForeThemeColorIndex =0
                    ForeTint =65.0
                    ConditionalFormat14 = Begin
                        0x010001000000010000000000000001000000ececec00ffffff00570000004900 ,
                        0x4900660028005b00660072006d005f0051007500610064007200610074005f00 ,
                        0x5400720061006e0073006500630074005d002e0046006f0072006d0021005b00 ,
                        0x66007300750062005f0053007000650063006900650073005f00430075007200 ,
                        0x720065006e0074005d002e0046006f0072006d002e0043006f006e0074007200 ,
                        0x6f006c007300280022007400620078004400650076004d006f00640065002200 ,
                        0x29002c0030002c00310029000000000000000000000000000000000000000000 ,
                        0x00
                    End
                End
                Begin OptionGroup
                    OverlapFlags =255
                    Left =120
                    Top =840
                    Width =7320
                    Height =660
                    TabIndex =11
                    Name ="frmSamplingVisit"

                    LayoutCachedLeft =120
                    LayoutCachedTop =840
                    LayoutCachedWidth =7440
                    LayoutCachedHeight =1500
                    Begin
                        Begin Label
                            BackStyle =1
                            OverlapFlags =247
                            TextAlign =2
                            Left =240
                            Top =720
                            Width =1380
                            Height =240
                            FontWeight =700
                            BackColor =12574431
                            Name ="lblSamplingVisit"
                            Caption ="Sampling Visit"
                            LayoutCachedLeft =240
                            LayoutCachedTop =720
                            LayoutCachedWidth =1620
                            LayoutCachedHeight =960
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =247
                    TextAlign =1
                    IMESentenceMode =3
                    Left =960
                    Top =1020
                    Width =1080
                    TabIndex =12
                    Name ="tbxStartDate"
                    ControlSource ="Start_Date"
                    Format ="Short Date"
                    StatusBarText ="M. Starting date for the event (Start_Date)"
                    InputMask ="99/99/0000;0;_"
                    ControlTipText ="Sampling visit start date"

                    LayoutCachedLeft =960
                    LayoutCachedTop =1020
                    LayoutCachedWidth =2040
                    LayoutCachedHeight =1260
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =247
                    IMESentenceMode =3
                    Left =7080
                    Top =1020
                    Width =360
                    TabIndex =13
                    Name ="Unit_Code"
                    ControlSource ="Unit_Code"
                    StatusBarText ="Park Code."

                    LayoutCachedLeft =7080
                    LayoutCachedTop =1020
                    LayoutCachedWidth =7440
                    LayoutCachedHeight =1260
                End
                Begin TextBox
                    Enabled = NotDefault
                    OverlapFlags =247
                    TextAlign =1
                    IMESentenceMode =3
                    Left =2760
                    Top =1020
                    Width =1080
                    TabIndex =14
                    Name ="Start_Time"
                    ControlSource ="Start_Time"
                    Format ="Short Time"
                    StatusBarText ="M. Starting time for the event (Start_Date)"
                    InputMask ="00:00;0;_"
                    ControlTipText ="Sampling visit start time (based on GPS time)"

                    LayoutCachedLeft =2760
                    LayoutCachedTop =1020
                    LayoutCachedWidth =3840
                    LayoutCachedHeight =1260
                End
                Begin Label
                    OverlapFlags =247
                    Left =360
                    Top =1020
                    Width =495
                    Height =240
                    FontWeight =700
                    Name ="lblStartDate"
                    Caption ="Date"
                    LayoutCachedLeft =360
                    LayoutCachedTop =1020
                    LayoutCachedWidth =855
                    LayoutCachedHeight =1260
                End
                Begin OptionGroup
                    OverlapFlags =247
                    Left =120
                    Top =180
                    Width =5940
                    Height =480
                    TabIndex =15
                    Name ="frmLocation"

                    LayoutCachedLeft =120
                    LayoutCachedTop =180
                    LayoutCachedWidth =6060
                    LayoutCachedHeight =660
                    Begin
                        Begin Label
                            BackStyle =1
                            OverlapFlags =247
                            TextAlign =2
                            Left =240
                            Top =60
                            Width =900
                            Height =240
                            FontWeight =700
                            BackColor =12574431
                            Name ="lblLocation"
                            Caption ="Location"
                            LayoutCachedLeft =240
                            LayoutCachedTop =60
                            LayoutCachedWidth =1140
                            LayoutCachedHeight =300
                        End
                    End
                End
                Begin Label
                    OverlapFlags =247
                    Left =2160
                    Top =1020
                    Width =495
                    Height =240
                    FontWeight =700
                    Name ="lblLocalGPSTime"
                    Caption ="Time"
                    LayoutCachedLeft =2160
                    LayoutCachedTop =1020
                    LayoutCachedWidth =2655
                    LayoutCachedHeight =1260
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
' Form:         frm_Data_Entry
' Level:        Application form
' Version:      1.02
' Basis:        -
'
' Description:  Data entry form object related properties, functions & procedures for UI display
'               Primary field data entry form
' Data source:  tbl_Locations
' Data access:  edit; allow additions off except for new records
' Pages:        none
' Functions:    none
' Source/date:  John R. Boetsch, June 2006
' References:   fxnSwitchboardIsOpen, fxnGUIDGen
' Revisions:    JRB - 6/x/2006  - 1.00 - initial version
'               BLC - 7/12/2017 - 1.01 - added documentation, error handling
'               BLC - 7/26/2017 - 1.02 - add CancelOpen() for canceling Form_Open from subform
' =================================

'---------------------
' Simulated Inheritance
'---------------------

'---------------------
' Declarations
'---------------------

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
' References:   -
' Adapted:      -
' Revisions:
'   JRB - 6/x/2006 - initial version
'   BLC - 7/12/2017 - added documentation, error handling
'   BLC - 7/26/2017 - added check for frm_Quadrat_Transect closed
' ---------------------------------
Private Sub Form_Open(Cancel As Integer)
On Error GoTo Err_Handler

    'check if frm_Quadrat_Transect was closed (as it is after new records are added)
    ' --> if so, close this form
    If GetTempVarIndex("CloseForm") > 1 Then
        
        'remove the CloseForm temp var & close this form
        If TempVars("CloseForm") = True Then
            
            TempVars.Remove "CloseForm"
            Cancel = True
            
            GoTo Exit_Handler
            
        End If
        
    End If

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
    Me!tbxStartDate.SetFocus

Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Form_Open[frm_Data_Entry form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          Form_Current
' Description:  form current record actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Adapted:      -
' Revisions:
'   NCPN - Unknown - initial version
'   BLC - 7/12/2017 - added documentation, error handling
' ---------------------------------
Private Sub Form_Current()
On Error GoTo Err_Handler

    Update_Loc_Info

Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Form_Current[frm_Data_Entry form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          Form_BeforeUpdate
' Description:  form actions before update
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Adapted:      -
' Revisions:
'   NCPN - Unknown - initial version
'   BLC - 7/12/2017 - added documentation, error handling
' ---------------------------------
Private Sub Form_BeforeUpdate(Cancel As Integer)
On Error GoTo Err_Handler
  
  If IsNull(Me!Start_Date) Then
        ' ask user if (s)he wants to enter data or cancel and close form
        If MsgBox("Visit date is missing - do you want to enter the missing data?", vbYesNo, "Date missing") = vbNo Then
            Me.Undo
        End If
  End If

Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Form_BeforeUpdate[frm_Data_Entry form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          Form_BeforeInsert
' Description:  form actions before a new record is inserted
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Adapted:      -
' Revisions:
'   NCPN - Unknown - initial version
'   BLC - 7/12/2017 - added documentation, error handling
' ---------------------------------
Private Sub Form_BeforeInsert(Cancel As Integer)
On Error GoTo Err_Handler

        Dim Db As DAO.Database
        Dim Versions As DAO.Recordset
        Dim strSQL As String
        
    
    ' Set master version number on event record
    Set Db = CurrentDb
    strSQL = "SELECT [version_key_number] FROM [tbl_master_version] ORDER BY [version_key_number] DESC"
    Set Versions = Db.OpenRecordset(strSQL)
    Versions.MoveFirst
    Me![version_key_number] = Versions![version_key_number]
    Versions.Close

    ' Create the GUID primary key value
    If IsNull(Me!Event_ID) Then
        If GetDataType("tbl_Events", "Event_ID") = dbText Then
            Me.Event_ID = fxnGUIDGen
        End If
    End If

Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Form_BeforeInsert[frm_Data_Entry form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          tbxStartDate_AfterUpdate
' Description:  tbxStartDate actions after update
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Adapted:      -
' Revisions:
'   NCPN - Unknown - initial version
'   BLC - 7/12/2017 - added documentation, error handling
' ---------------------------------
Private Sub tbxStartDate_AfterUpdate()
On Error GoTo Err_Handler
        
        Dim Db As DAO.Database
        Dim Events As DAO.Recordset
        Dim strSQL As String
        
    On Error GoTo Err_Handler
    
    ' Check for duplicate date
    strSQL = "SELECT Event_ID FROM tbl_Events WHERE [Location_ID] = '" & Me!cboLocation_ID & "' AND [Start_Date] = #" & Me!Start_Date & "#"

    Set Db = CurrentDb
    Set Events = Db.OpenRecordset(strSQL)
    
    If Not Events.EOF Then
      MsgBox " Duplicate visit date - update cancelled."
      Me.Undo
      Events.Close
      DoCmd.Close
      GoTo Exit_Handler
    End If
    Events.Close

Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - tbxStartDate_AfterUpdate[frm_Data_Entry form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          cbxLocationID_AfterUpdate
' Description:  cbxLocationID actions after update
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Adapted:      -
' Revisions:
'   NCPN - Unknown - initial version
'   BLC - 7/12/2017 - added documentation, error handling
' ---------------------------------
Private Sub cbxLocationID_AfterUpdate()
On Error GoTo Err_Handler

    ' Update_Loc_Info

Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - cbxLocationID_AfterUpdate[frm_Data_Entry form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          btnClose_Click
' Description:  Close button click actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Adapted:      -
' Revisions:
'   NCPN - Unknown - initial version
'   BLC - 7/12/2017 - added documentation, error handling
' ---------------------------------
Private Sub btnClose_Click()
On Error GoTo Err_Handler
    
    DoCmd.RunCommand acCmdSaveRecord
    DoCmd.Close , , acSaveNo

Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnClose_Click[frm_Data_Entry form])"
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
' Adapted:      -
' Revisions:
'   NCPN - Unknown - initial version
'   BLC - 7/12/2017 - added documentation, error handling
' ---------------------------------
Private Sub Form_Close()
On Error GoTo Err_Handler

    If IsLoaded("frm_Data_Gateway") Then
        Forms("frm_Data_Gateway").Requery
    End If

Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Form_Close[frm_Data_Entry form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          Update_Loc_Info
' Description:  Updates associated location information when Location_ID is updated
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   GetCriteriaString
' Adapted:      -
' Revisions:
'   SK  - 9/x/2006 - initial version
'   BLC - 7/12/2017 - added documentation, error handling
' ---------------------------------
Public Sub Update_Loc_Info()
On Error GoTo Err_Handler

    Dim strXY As Variant
    Dim strCriteria As String
    
    If IsNull(Me!txtLocation_ID) Then
        Me!txtUnit_Code = Null
    Else
        strCriteria = GetCriteriaString("Location_ID=", "tbl_Locations", "Location_ID", Me.Name, "txtLocation_ID")
        Me!txtUnit_Code = DLookup("Unit_Code", "tbl_Locations", strCriteria)
    End If

Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Update_Loc_Info[frm_Data_Entry form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Function:     UpdateEvent
' Description:  Updates sampling event data (start date, observer, comments)
' Assumptions:  Control property contains the following
'                   =UpdateEvent([Screen].[ActiveControl])
'               in its on change event property
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, July 17, 2017 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 7/17/2016 - initial version
' ---------------------------------
Private Function UpdateEvent(ctrl As Control) As Boolean
On Error GoTo Err_Handler

    Dim obs As Variant
    Dim start As Variant
    Dim cmt As Variant
    
    'set values from form
    start = Me.tbxStartDate
    obs = Me.Observer
    cmt = Me.Comments

    Dim se As New SamplingEvent
    
    With se
        .EventID = Me.tbxEventID
        
        Select Case ctrl.Name
            Case "tbxStartDate"
                'start date
                If Not IsNull(start) Then
                    .StartDate = start
                    .UpdateStartDate
                End If
                
            Case "Observer"
                'observer
                If Not IsNull(obs) Then
                    .Observer = obs
                    .UpdateObserver
                End If
            
            Case "Comments"
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
            "Error encountered (#" & Err.Number & " - UpdateEvent[frm_Data_Entry form])"
    End Select
    Resume Exit_Handler
End Function
