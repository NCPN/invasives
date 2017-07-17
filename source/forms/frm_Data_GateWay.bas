Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    MaxButton = NotDefault
    MinButton = NotDefault
    ControlBox = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    AllowDeletions = NotDefault
    DividingLines = NotDefault
    AllowAdditions = NotDefault
    OrderByOn = NotDefault
    ScrollBars =2
    TabularFamily =0
    BorderStyle =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =9420
    DatasheetFontHeight =10
    ItemSuffix =38
    Left =630
    Top =4350
    Right =15600
    Bottom =9780
    DatasheetGridlinesColor =12632256
    Filter ="Unit_code = 'ZION'"
    OrderBy ="Plot_ID DESC, Unit_Code"
    RecSrcDt = Begin
        0x29b5dcdf75fbe240
    End
    RecordSource ="qfrm_Data_Gateway"
    Caption ="Data Gateway"
    OnOpen ="[Event Procedure]"
    DatasheetFontName ="Arial"
    PrtMip = Begin
        0xa0050000a0050000a0050000a005000000000000201c0000e010000001000000 ,
        0x010000006801000000000000a10700000100000001000000
    End
    AllowPivotTableView =0
    AllowPivotChartView =0
    AllowPivotChartView =0
    FilterOnLoad =0
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
        Begin FormHeader
            Height =1248
            BackColor =11056034
            Name ="FormHeader"
            Begin
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =4380
                    Top =1020
                    Width =1680
                    Height =228
                    Name ="lblUpdated_Date"
                    Caption ="Entered/updated*"
                    FontName ="Arial"
                    OnDblClick ="[Event Procedure]"
                    Tag ="DetachedLabel"
                    ControlTipText ="Double click to sort by date updated"
                    LayoutCachedLeft =4380
                    LayoutCachedTop =1020
                    LayoutCachedWidth =6060
                    LayoutCachedHeight =1248
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =105
                    Top =1020
                    Width =795
                    Height =225
                    Name ="lblUnit_code"
                    Caption ="Unit*"
                    FontName ="Arial"
                    OnDblClick ="[Event Procedure]"
                    Tag ="DetachedLabel"
                    ControlTipText ="Double click to sort by park (unit code)"
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =8460
                    Top =120
                    Width =720
                    Height =354
                    FontSize =9
                    FontWeight =700
                    TabIndex =3
                    Name ="btnClose"
                    Caption ="Close"
                    OnClick ="[Event Procedure]"
                    FontName ="Arial"
                    ControlTipText ="Close the data entry form"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120

                    LayoutCachedLeft =8460
                    LayoutCachedTop =120
                    LayoutCachedWidth =9180
                    LayoutCachedHeight =474
                    ForeThemeColorIndex =0
                    UseTheme =1
                    BackColor =15921906
                    BackThemeColorIndex =1
                    BackShade =95.0
                    BorderThemeColorIndex =0
                    HoverColor =9434577
                    PressedColor =15921906
                    PressedThemeColorIndex =1
                    PressedShade =95.0
                    HoverForeColor =16711680
                    PressedForeColor =0
                    PressedForeThemeColorIndex =0
                    QuickStyle =22
                    QuickStyleMask =-53
                    WebImagePaddingLeft =4
                    WebImagePaddingTop =2
                    WebImagePaddingRight =4
                    WebImagePaddingBottom =7
                End
                Begin Label
                    OverlapFlags =85
                    Left =120
                    Top =120
                    Width =4860
                    Height =408
                    BackColor =16777215
                    ForeColor =0
                    Name ="lablbl"
                    Caption ="* Double-click on the field label to change sort order.  Double-click on a Plot "
                        "ID to open the Site form for that record."
                    FontName ="Arial"
                    ControlTipText ="View mode"
                End
                Begin ComboBox
                    TabStop = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1500
                    Top =660
                    Width =960
                    ColumnOrder =1
                    TabIndex =1
                    ColumnInfo ="\"\";\"\";\"10\";\"8\""
                    Name ="cbxPark"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT DISTINCT tbl_Locations.Unit_Code FROM tbl_Locations ORDER BY tbl_Location"
                        "s.Unit_Code; "
                    StatusBarText ="Park code"
                    AfterUpdate ="[Event Procedure]"
                    FontName ="Arial"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =120
                            Top =660
                            Width =1320
                            Height =228
                            Name ="lblPark"
                            Caption ="Filter by:  Park"
                            FontName ="Arial"
                        End
                    End
                End
                Begin ToggleButton
                    TabStop = NotDefault
                    OverlapFlags =85
                    Left =2760
                    Top =660
                    Width =480
                    Height =300
                    ColumnOrder =0
                    Name ="tglFilterByPark"
                    Caption ="Filter on"
                    PictureData = Begin
                        0x2800000010000000100000000100040000000000800000000000000000000000 ,
                        0x0000000000000000000000000000800000800000008080008000000080008000 ,
                        0x8080000080808000c0c0c0000000ff00c0c0c00000ffff00ff000000c0c0c000 ,
                        0xffff0000ffffff00dadadadadadadadaadadadadadadadaddadadadadadadada ,
                        0xadadad0000adadaddadada0660dadadaadadad0660adadaddadada0f80dadada ,
                        0xadadad0f80adadaddadad088860adadaadad06888660adaddad068f888660ada ,
                        0xad068f88888660add068fff88886660aa00000000000000ddadadadadadadada ,
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
                    FontName ="Arial"
                    OnClick ="[Event Procedure]"
                    ObjectPalette = Begin
                        0x000301000000000000000000
                    End
                    ControlTipText ="Turn the park filter on or off"

                    LayoutCachedLeft =2760
                    LayoutCachedTop =660
                    LayoutCachedWidth =3240
                    LayoutCachedHeight =960
                    UseTheme =1
                    HoverColor =9434577
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =1005
                    Top =1020
                    Width =675
                    Height =225
                    Name ="lblPlot_ID"
                    Caption ="Route*"
                    FontName ="Arial"
                    OnDblClick ="[Event Procedure]"
                    ControlTipText ="Double click to sort by route"
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =6540
                    Top =120
                    Width =1680
                    FontSize =9
                    FontWeight =700
                    TabIndex =4
                    Name ="btnNewSite"
                    Caption ="Add New Route"
                    OnClick ="[Event Procedure]"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120

                    LayoutCachedLeft =6540
                    LayoutCachedTop =120
                    LayoutCachedWidth =8220
                    LayoutCachedHeight =480
                    ForeThemeColorIndex =0
                    UseTheme =1
                    BackColor =15921906
                    BackThemeColorIndex =1
                    BackShade =95.0
                    BorderThemeColorIndex =0
                    HoverColor =9434577
                    PressedColor =15921906
                    PressedThemeColorIndex =1
                    PressedShade =95.0
                    HoverForeColor =16711680
                    PressedForeColor =0
                    PressedForeThemeColorIndex =0
                    QuickStyle =22
                    QuickStyleMask =-53
                    WebImagePaddingLeft =4
                    WebImagePaddingTop =2
                    WebImagePaddingRight =4
                    WebImagePaddingBottom =7
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =5340
                    Top =120
                    Width =960
                    FontSize =9
                    FontWeight =700
                    TabIndex =2
                    Name ="btnRefresh"
                    Caption ="Refresh"
                    OnClick ="[Event Procedure]"
                    FontName ="Arial"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120

                    LayoutCachedLeft =5340
                    LayoutCachedTop =120
                    LayoutCachedWidth =6300
                    LayoutCachedHeight =480
                    ForeThemeColorIndex =0
                    UseTheme =1
                    BackColor =15921906
                    BackThemeColorIndex =1
                    BackShade =95.0
                    BorderThemeColorIndex =0
                    HoverColor =9434577
                    PressedColor =15921906
                    PressedThemeColorIndex =1
                    PressedShade =95.0
                    HoverForeColor =16711680
                    PressedForeColor =0
                    PressedForeThemeColorIndex =0
                    QuickStyle =22
                    QuickStyleMask =-53
                    WebImagePaddingLeft =4
                    WebImagePaddingTop =2
                    WebImagePaddingRight =4
                    WebImagePaddingBottom =7
                End
            End
        End
        Begin Section
            Height =420
            BackColor =11056034
            Name ="Detail"
            Begin
                Begin CommandButton
                    OverlapFlags =93
                    Left =6360
                    Top =60
                    Width =1320
                    Height =300
                    TabIndex =6
                    Name ="btnNewVisit"
                    Caption ="Add New Visit"
                    OnClick ="[Event Procedure]"
                    FontName ="Arial"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120

                    LayoutCachedLeft =6360
                    LayoutCachedTop =60
                    LayoutCachedWidth =7680
                    LayoutCachedHeight =360
                    ForeThemeColorIndex =0
                    UseTheme =1
                    BackColor =15921906
                    BackThemeColorIndex =1
                    BackShade =95.0
                    BorderThemeColorIndex =0
                    HoverColor =9434577
                    PressedColor =15921906
                    PressedThemeColorIndex =1
                    PressedShade =95.0
                    HoverForeColor =16711680
                    PressedForeColor =0
                    PressedForeThemeColorIndex =0
                    QuickStyle =22
                    QuickStyleMask =-53
                    WebImagePaddingLeft =4
                    WebImagePaddingTop =2
                    WebImagePaddingRight =4
                    WebImagePaddingBottom =7
                    Overlaps =1
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =4380
                    Top =60
                    Width =1680
                    Height =239
                    ColumnWidth =1710
                    TabIndex =1
                    Name ="tbxUpdated"
                    ControlSource ="Updated_Date"
                    Format ="yyyy mmm dd hh:nn"
                    StatusBarText ="Date on which data entry occurred"
                    FontName ="Arial"

                    LayoutCachedLeft =4380
                    LayoutCachedTop =60
                    LayoutCachedWidth =6060
                    LayoutCachedHeight =299
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =120
                    Top =60
                    Width =780
                    ColumnWidth =2310
                    Name ="tbxPark"
                    ControlSource ="Unit_Code"
                    StatusBarText ="Unit code"
                    FontName ="Arial"

                End
                Begin TextBox
                    Visible = NotDefault
                    Enabled = NotDefault
                    Locked = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =93
                    IMESentenceMode =3
                    Left =1920
                    Top =60
                    Width =420
                    TabIndex =2
                    Name ="txtLocation_ID"
                    ControlSource ="Location_ID"
                    StatusBarText ="Name of the location"
                    FontName ="Arial"

                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    OverlapFlags =247
                    IMESentenceMode =3
                    Left =1050
                    Top =60
                    Width =3240
                    TabIndex =3
                    Name ="tbxPlot_ID"
                    ControlSource ="Plot_ID"
                    StatusBarText ="Plot identifier"

                    LayoutCachedLeft =1050
                    LayoutCachedTop =60
                    LayoutCachedWidth =4290
                    LayoutCachedHeight =300
                End
                Begin CommandButton
                    OverlapFlags =215
                    Left =6360
                    Top =60
                    Width =1320
                    Height =300
                    TabIndex =4
                    Name ="btnVisitList"
                    Caption ="View Visits"
                    OnClick ="[Event Procedure]"
                    FontName ="Arial"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120

                    LayoutCachedLeft =6360
                    LayoutCachedTop =60
                    LayoutCachedWidth =7680
                    LayoutCachedHeight =360
                    ForeThemeColorIndex =0
                    UseTheme =1
                    BackColor =15921906
                    BackThemeColorIndex =1
                    BackShade =95.0
                    BorderThemeColorIndex =0
                    HoverColor =9434577
                    PressedColor =15921906
                    PressedThemeColorIndex =1
                    PressedShade =95.0
                    HoverForeColor =16711680
                    PressedForeColor =0
                    PressedForeThemeColorIndex =0
                    QuickStyle =22
                    QuickStyleMask =-53
                    WebImagePaddingLeft =4
                    WebImagePaddingTop =2
                    WebImagePaddingRight =4
                    WebImagePaddingBottom =7
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =7800
                    Top =60
                    Width =1410
                    Height =300
                    TabIndex =5
                    Name ="btnSiteChar"
                    Caption ="Edit/Add Routes"
                    OnClick ="[Event Procedure]"
                    FontName ="Arial"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120

                    LayoutCachedLeft =7800
                    LayoutCachedTop =60
                    LayoutCachedWidth =9210
                    LayoutCachedHeight =360
                    ForeThemeColorIndex =0
                    UseTheme =1
                    BackColor =15921906
                    BackThemeColorIndex =1
                    BackShade =95.0
                    BorderThemeColorIndex =0
                    HoverColor =9434577
                    PressedColor =15921906
                    PressedThemeColorIndex =1
                    PressedShade =95.0
                    HoverForeColor =16711680
                    PressedForeColor =0
                    PressedForeThemeColorIndex =0
                    QuickStyle =22
                    QuickStyleMask =-53
                    WebImagePaddingLeft =4
                    WebImagePaddingTop =2
                    WebImagePaddingRight =4
                    WebImagePaddingBottom =7
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
' Form:         frm_Data_Gateway
' Level:        Application form
' Version:      1.01
' Basis:        -
'
' Description:  Data Gateway form object related properties, functions & procedures for UI display
'
' Data source:  qfrm_Data_Gateway
' Data access:  view and delete records (delete by cmdDeleteRec)
' Pages:        none
' Functions:    fxnSortRecords
' Source/date:  John R. Boetsch, June 7, 2006
' References:   -
' Revisions:    JRB - 6/7/2006  - 1.00 - initial version
'               SK - 9/x/2006   - 1.01 - added CorrectText calls where strings were being used in criteria
'                                        - updated cmdDeleteRec_Click() event to use appropriate criteria depending on primary key
'               BLC - 7/14/2017 - 1.02 - added documentation, error handling, renamed buttons
' =================================

'---------------------
' Simulated Inheritance
'---------------------

'---------------------
' Declarations
'---------------------
Dim strSortField As String    ' Keeps track of current sort settings
Dim strSortOrder As String
Dim strSortFieldLabel As String

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
' Source/date:  NCPN, Unknown - for NCPN tools
' Adapted:      -
' Revisions:
'   JRB - 6/7/2006  - initial version
'   BLC - 7/14/2017 - added documentation, error handling
' ---------------------------------
Private Sub Form_Open(Cancel As Integer)
On Error GoTo Err_Handler
    
    Dim varReturn As Variant

    ' On opening the form, set the initial sort order
    strSortFieldLabel = "lblRoute"
    varReturn = SortRecords("Unit_Code", "Plot_ID")
    
    ' Set the filter
    If fxnSwitchboardIsOpen Then
        Me.cbxPark = Forms!frm_Switchboard.cPark
        Me.Filter = "Unit_code = " & CorrectText(Me.cbxPark)
        Me.FilterOn = True
        Me.lblPark.FontBold = True
        Me.tglFilterByPark = True
    End If
  
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Form_Open[frm_Data_Gateway form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          cbxPark_AfterUpdate
' Description:  Park selection actions after update
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  NCPN, Unknown - for NCPN tools
' Adapted:      -
' Revisions:
'   JRB - 6/7/2006  - initial version
'   BLC - 7/12/2017 - added documentation, error handling
' ---------------------------------
Private Sub cbxPark_AfterUpdate()
On Error GoTo Err_Handler
    
    Me.Filter = "Unit_code = " & CorrectText(Me.cbxPark)
    If tglFilterByPark Then
      Me.Filter = "Unit_code = " & CorrectText(Me.cbxPark)
      Me.FilterOn = True
      Me.lblPark.FontBold = True
    End If
  
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - cbxPark_AfterUpdate[frm_Data_Gateway form])"
    End Select
    Resume Exit_Handler
End Sub

'---------------------
' Button/Toggle Click Events
'---------------------

' ---------------------------------
' Sub:          tglFilterByPark_AfterUpdate
' Description:  Filter by park toggle actions after update
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  NCPN, Unknown - for NCPN tools
' Adapted:      -
' Revisions:
'   JRB - 6/7/2006  - initial version
'   BLC - 7/12/2017 - added documentation, error handling
' ---------------------------------
Private Sub tglFilterByPark_AfterUpdate()
On Error GoTo Err_Handler
    
    If Me.ActiveControl Then
      If Not IsNull(Me!cbxPark) Then
        Me.Filter = "Unit_code = " & CorrectText(Me.cbxPark)
        Me.FilterOn = True
        Me.lblPark.FontBold = True
      End If
    Else
        Me.FilterOn = False
        Me.lblPark.FontBold = False
    End If
  
  
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - tglFilterByPark_AfterUpdate[frm_Data_Gateway form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          btnRefresh_Click
' Description:  Refresh button click actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  NCPN, Unknown - for NCPN tools
' Adapted:      -
' Revisions:
'   JRB - 6/7/2006  - initial version
'   BLC - 7/12/2017 - added documentation, error handling
' ---------------------------------
Private Sub btnRefresh_Click()
On Error GoTo Err_Handler
    
    Me.Requery
  
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnRefresh_Click[frm_Data_Gateway form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          btnNewSite_Click
' Description:  New site button click actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  NCPN, Unknown - for NCPN tools
' Adapted:      -
' Revisions:
'   JRB - 6/7/2006  - initial version
'   BLC - 7/12/2017 - added documentation, error handling
' ---------------------------------
Private Sub btnNewSite_Click()
On Error GoTo Err_Handler
    
    DoCmd.OpenForm "frm_Locations", , , , acFormAdd, , "New record"
      
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnNewSite_Click[frm_Data_Gateway form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          btnVisitList_Click
' Description:  Visit list button click actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  NCPN, Unknown - for NCPN tools
' Adapted:      -
' Revisions:
'   JRB - 6/7/2006  - initial version
'   BLC - 7/12/2017 - added documentation, error handling
' ---------------------------------
Private Sub btnVisitList_Click()
On Error GoTo Err_Handler
    
    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "frm_Visit_Date"
    
    stLinkCriteria = "[Location_ID]=" & "'" & Me![txtLocation_ID] & "'"
    DoCmd.OpenForm stDocName, , , stLinkCriteria
  
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnVisitList_Click[frm_Data_Gateway form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          btnNewVisit_Click
' Description:  New visit button click actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  NCPN, Unknown - for NCPN tools
' Adapted:      -
' Revisions:
'   JRB - 6/7/2006  - initial version
'   BLC - 7/12/2017 - added documentation, error handling
'                     NOTE: control is currently hidden under View Visits
'                           form design must be changed to allow users to use it
' ---------------------------------
Private Sub btnNewVisit_Click()
On Error GoTo Err_Handler
    
    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "frm_Plot_Revisit"
    
    stLinkCriteria = "[Location_ID]=" & "'" & Me![txtLocation_ID] & "'"
    DoCmd.OpenForm stDocName, , , stLinkCriteria
'    DoCmd.Close acForm, "frm_Data_Gateway"
  
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnNewVisit_Click[frm_Data_Gateway form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          btnSiteChar_Click
' Description:  Site characteristic click actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  NCPN, Unknown - for NCPN tools
' Adapted:      -
' Revisions:
'   JRB - 6/7/2006  - initial version
'   BLC - 7/12/2017 - added documentation, error handling
' ---------------------------------
Private Sub btnSiteChar_Click()
On Error GoTo Err_Handler
    
    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "frm_Locations"
    
    stLinkCriteria = "[Location_ID]=" & "'" & Me![txtLocation_ID] & "'"
    DoCmd.OpenForm stDocName, , , stLinkCriteria
  
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnSiteChar_Click[frm_Data_Gateway form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          btnClose_Click
' Description:  form closing actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  NCPN, Unknown - for NCPN tools
' Adapted:      -
' Revisions:
'   JRB - 6/7/2006  - initial version
'   BLC - 7/12/2017 - added documentation, error handling
' ---------------------------------
Private Sub btnClose_Click()
On Error GoTo Err_Handler
    
    DoCmd.Close , , acSaveNo
  
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnClose_Click[frm_Data_Gateway form])"
    End Select
    Resume Exit_Handler
End Sub

'---------------------
' Sorting Methods
'---------------------
'  Procedures that re-sort the records if the user
'  double-clicks on a field label

' ---------------------------------
' Sub:          lblUnit_code_DblClick
' Description:  Unit code label double click actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  NCPN, Unknown - for NCPN tools
' Adapted:      -
' Revisions:
'   JRB - 6/7/2006  - initial version
'   BLC - 7/12/2017 - added documentation, error handling
' ---------------------------------
Private Sub lblUnit_code_DblClick(Cancel As Integer)
On Error GoTo Err_Handler
    
    SortRecords ("Unit_code")
  
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - lblUnit_code_DblClick[frm_Data_Gateway form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          lblPlot_ID_DblClick
' Description:  Plot ID (route) label double click actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  NCPN, Unknown - for NCPN tools
' Adapted:      -
' Revisions:
'   JRB - 6/7/2006  - initial version
'   BLC - 7/12/2017 - added documentation, error handling
' ---------------------------------
Private Sub lblPlot_ID_DblClick(Cancel As Integer)
On Error GoTo Err_Handler
    
    SortRecords ("Plot_ID") 'Plot_ID = Route
  
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - lblPlot_ID_DblClick[frm_Data_Gateway form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          lblUpdated_Date_DblClick
' Description:  Updated date double click actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  NCPN, Unknown - for NCPN tools
' Adapted:      -
' Revisions:
'   JRB - 6/7/2006  - initial version
'   BLC - 7/12/2017 - added documentation, error handling
' ---------------------------------
Private Sub lblUpdated_Date_DblClick(Cancel As Integer)
On Error GoTo Err_Handler
    
    SortRecords ("Updated_Date")
  
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - lblUpdated_Date_DblClick[frm_Data_Gateway form])"
    End Select
    Resume Exit_Handler
End Sub

'---------------------
' Custom Functions
'---------------------

' ---------------------------------
' Function:     SortRecords
' Description:  Sorts form records by field indicated
' Assumptions:  -
' Parameters:   strFieldName - first field to sort by (string)
'               strField2Name - second field to sort by (optional, string)
' Returns:      -
' Throws:       none
' References:   strFieldName, strSortOrder, strSortFieldLabel
'               (form-level variables)
' Source/date:  JRB, May 5, 2006 - for NCPN tools
' Adapted:      -
' Revisions:
'   JRB - 5/5/2006  - initial version
'   BLC - 7/12/2017 - added documentation, error handling
' ---------------------------------
Private Function SortRecords(ByVal strFieldName As String, _
                        Optional ByVal strField2Name As String)
On Error GoTo Err_Handler
      
    Dim strOrderBy As String

    ' Toggle sort ASC/DESC
    If strFieldName = strSortField And strSortOrder = "" Then
        strSortOrder = " DESC"
    Else
        strSortOrder = ""
    End If
    
    ' Create ORDER BY string & activate filter
    strOrderBy = strFieldName & strSortOrder
    
    If strField2Name <> "" Then
        strOrderBy = strField2Name & " DESC, " & strOrderBy
    End If
    
    strSortField = strFieldName
    
    Me.Form.OrderBy = strOrderBy
    Me.Form.OrderByOn = True

    ' Set label format to indicate sorted field
'    Me.Controls.item(strSortFieldLabel).FontItalic = False
'    Me.Controls.item(strSortFieldLabel).fontBold = False
    
    strSortFieldLabel = "lbl" & strFieldName
    Me.Controls.item(strSortFieldLabel).FontItalic = True
    Me.Controls.item(strSortFieldLabel).FontBold = True
  
Exit_Handler:
    Exit Function
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - SortRecords[frm_Data_Gateway form])"
    End Select
    Resume Exit_Handler
End Function
