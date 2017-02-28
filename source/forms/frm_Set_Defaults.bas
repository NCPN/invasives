Version =20
VersionRequired =20
Begin Form
    PopUp = NotDefault
    Modal = NotDefault
    RecordSelectors = NotDefault
    MaxButton = NotDefault
    MinButton = NotDefault
    ControlBox = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    AllowDeletions = NotDefault
    CloseButton = NotDefault
    DividingLines = NotDefault
    AllowAdditions = NotDefault
    DefaultView =0
    ScrollBars =0
    ViewsAllowed =1
    TabularFamily =0
    BorderStyle =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =4380
    DatasheetFontHeight =10
    ItemSuffix =9
    Left =6885
    Top =2100
    Right =11265
    Bottom =5475
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0xa3c57b9aedcee240
    End
    RecordSource ="tsys_App_Defaults"
    Caption =" Set application default values"
    DatasheetFontName ="Arial"
    PrtMip = Begin
        0xa0050000a0050000a0050000a005000000000000201c0000e010000001000000 ,
        0x010000006801000000000000a10700000100000001000000
    End
    AllowDatasheetView =0
    AllowPivotTableView =0
    AllowPivotChartView =0
    AllowPivotChartView =0
    Begin
        Begin Label
            BackStyle =0
            FontName ="Tahoma"
        End
        Begin Rectangle
            SpecialEffect =3
            BackStyle =0
        End
        Begin CommandButton
            FontSize =8
            FontWeight =400
            ForeColor =-2147483630
            FontName ="Tahoma"
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
        Begin Section
            Height =3180
            BackColor =11056034
            Name ="Detail"
            Begin
                Begin ComboBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =2
                    ListWidth =3600
                    Left =972
                    Top =960
                    Width =1245
                    Height =252
                    FontSize =9
                    TabIndex =1
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"10\";\"10\""
                    Name ="cmbPark"
                    ControlSource ="Park"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT tlu_Parks.ParkCode, tlu_Parks.ParkName FROM tlu_Parks ORDER BY tlu_Parks."
                        "ParkCode; "
                    ColumnWidths ="720;2880"
                    AfterUpdate ="[Event Procedure]"
                    FontName ="Arial"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =180
                            Top =960
                            Width =480
                            Height =255
                            FontSize =9
                            FontWeight =700
                            BackColor =11056034
                            Name ="labPark"
                            Caption ="Park"
                            FontName ="Arial"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =972
                    Top =600
                    Width =3165
                    Height =252
                    FontSize =9
                    ColumnInfo ="\"\";\"\";\"10\";\"0\""
                    Name ="cmbUser"
                    ControlSource ="User_name"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT Last_Name & \"_\" & First_Name FROM tlu_Contacts ORDER BY Last_Name, Firs"
                        "t_Name; "
                    FontName ="Arial"
                    OnNotInList ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =180
                            Top =600
                            Width =468
                            Height =252
                            FontSize =9
                            FontWeight =700
                            BackColor =11056034
                            Name ="labUser"
                            Caption ="User"
                            FontName ="Arial"
                        End
                    End
                End
                Begin TextBox
                    TabStop = NotDefault
                    SpecialEffect =0
                    OldBorderStyle =1
                    OverlapFlags =215
                    IMESentenceMode =3
                    Left =180
                    Top =2280
                    Width =3963
                    Height =417
                    FontSize =9
                    TabIndex =4
                    Name ="Soil_Survey_Area"
                    ControlSource ="Soil_Survey_Area"
                    FontName ="Arial"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =180
                            Top =2040
                            Width =1617
                            Height =252
                            FontSize =9
                            FontWeight =700
                            BackColor =11056034
                            Name ="labActivity"
                            Caption ="Soil Survey Area"
                            FontName ="Arial"
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =3480
                    Top =120
                    Width =720
                    Height =354
                    FontSize =9
                    FontWeight =700
                    TabIndex =6
                    ForeColor =0
                    Name ="cmdOK"
                    Caption ="OK"
                    OnClick ="[Event Procedure]"
                    FontName ="Arial"
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =3120
                    Top =1020
                    Width =1035
                    FontSize =9
                    FontWeight =700
                    TabIndex =5
                    ForeColor =0
                    Name ="cmdNewUser"
                    Caption ="New user"
                    OnClick ="[Event Procedure]"
                    FontName ="Arial"
                    ControlTipText ="Add a new user"
                End
                Begin ComboBox
                    RowSourceTypeInt =1
                    OldBorderStyle =0
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =972
                    Top =1320
                    Width =1245
                    FontSize =9
                    TabIndex =2
                    Name ="cmbDatum"
                    ControlSource ="Datum"
                    RowSourceType ="Value List"
                    RowSource ="\"NAD27\";\"NAD83\";\"WGS84\""
                    FontName ="Arial"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =180
                            Top =1320
                            Width =672
                            Height =252
                            FontSize =9
                            FontWeight =700
                            BackColor =11056034
                            Name ="labDatum"
                            Caption ="Datum"
                            FontName ="Arial"
                        End
                    End
                End
                Begin ComboBox
                    RowSourceTypeInt =1
                    OverlapFlags =85
                    IMESentenceMode =3
                    ListWidth =1440
                    Left =960
                    Top =1680
                    Width =1260
                    TabIndex =3
                    Name ="Zone"
                    ControlSource ="Zone"
                    RowSourceType ="Value List"
                    RowSource ="12;13"
                    ColumnWidths ="1440"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =180
                            Top =1680
                            Width =660
                            Height =245
                            FontSize =9
                            FontWeight =700
                            Name ="UTM Zone_Label"
                            Caption ="Zone"
                            FontName ="Arial"
                            EventProcPrefix ="UTM_Zone_Label"
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
' FORM NAME:    frm_Set_Defaults
' Description:  Standard module for setting application defaults
' Data source:  tsys_App_Defaults
' Data access:  edit only, no deletions
' Pages:        none
' Functions:    none
' References:   none
' Source/date:  John R. Boetsch, May 16, 2006
' Revisions:    <name, date, desc - add lines as you go>
' =================================

Private Sub cmbUser_NotInList(NewData As String, Response As Integer)
    On Error GoTo Err_Handler

    Me.ActiveControl.Undo

Exit_Procedure:
    Exit Sub

Err_Handler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    Resume Exit_Procedure

End Sub

Private Sub cmdNewUser_Click()
    On Error GoTo Err_Handler
    
    ' Open the contacts form
    DoCmd.OpenForm "frm_Contacts"

Exit_Procedure:
    Exit Sub

Err_Handler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    Resume Exit_Procedure

End Sub

Private Sub cmbPark_AfterUpdate()
    On Error GoTo Err_Handler

    Dim strMsg As String
    Dim strZone As String
    Dim strDatum As String
    Dim strNetwork As String

    If Not IsNull(Me!cmbDatum) Or Not IsNull(Me!Zone) Then
    ' On changing the park, prompt for resetting the datum and declination
        strZone = Nz(Me.Zone, "---")
        strDatum = Nz(Me.cmbDatum, "---")
        strMsg = "Changing parks requires verification of other settings." & vbCrLf & vbCrLf
        strMsg = strMsg & "Datum: " & strDatum & "  Zone: " & strZone & vbCrLf & vbCrLf
        strMsg = strMsg & "Would you like to keep these settings?"
        If MsgBox(strMsg, vbYesNo, "Verify park info") = vbNo Then
            Me.cmbDatum = Null
            Me.Zone = Null
        End If
    End If

Exit_Procedure:
    Exit Sub

Err_Handler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    Resume Exit_Procedure

End Sub

Private Sub cmdOK_Click()
    On Error GoTo Err_Handler

    Dim varOpenArgs As Variant
    
    varOpenArgs = Me.OpenArgs
    
    ' Make sure the information is valid before updating the record
    If varOpenArgs <> 0 Then
        '  Verify that the critical data elements have been completed before saving
        If IsNull(Me.User_name) Then
            MsgBox "Please indicate the user name", vbOKOnly, "Validation error"
            Me.cmbUser.SetFocus
            GoTo Exit_Procedure
        ElseIf IsNull(Me.Park) Then
            MsgBox "Please indicate the park", vbOKOnly, "Validation error"
            Me.cmbPark.SetFocus
            GoTo Exit_Procedure
        End If
    End If

    DoCmd.Close acForm, Me.name, acSaveNo
    DoCmd.OpenForm "frm_Switchboard"
    Select Case varOpenArgs
        Case 1
            DoCmd.OpenForm "frm_Data_Gateway", , , , , , varOpenArgs
        Case 2
            DoCmd.OpenForm "frm_Browser", , , , , , varOpenArgs
        Case 3
            DoCmd.OpenForm "frm_Select_Not_Present", , , , , , varOpenArgs
        Case 4
            ' opened by switchboard only ... do nothing
        Case Else
            MsgBox "Error: OpenArgs property out of range", vbCritical
    End Select

Exit_Procedure:
    Exit Sub

Err_Handler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    Resume Exit_Procedure

End Sub
