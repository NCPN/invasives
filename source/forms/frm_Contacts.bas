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
    AllowDeletions = NotDefault
    DividingLines = NotDefault
    DefaultView =0
    ScrollBars =0
    ViewsAllowed =1
    BorderStyle =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    Cycle =1
    GridX =24
    GridY =24
    Width =10260
    DatasheetFontHeight =10
    ItemSuffix =135
    Left =4425
    Top =3000
    Right =14685
    Bottom =8025
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0xd44c5a625608e340
    End
    RecordSource ="tlu_Contacts"
    Caption =" View and edit contact information"
    OnCurrent ="[Event Procedure]"
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
        Begin Line
            SpecialEffect =3
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
            SpecialEffect =2
            BackColor =-2147483643
            ForeColor =-2147483640
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
        Begin Section
            CanGrow = NotDefault
            Height =5040
            BackColor =12574431
            Name ="Detail"
            Begin
                Begin CommandButton
                    OverlapFlags =85
                    Left =9180
                    Top =120
                    Width =780
                    Height =414
                    FontSize =9
                    FontWeight =700
                    TabIndex =1
                    Name ="cmdClose"
                    Caption ="Close"
                    OnClick ="[Event Procedure]"
                    FontName ="Arial"
                    ControlTipText ="Return to the previous screen"
                End
                Begin Line
                    OverlapFlags =85
                    Left =3420
                    Top =720
                    Width =5400
                    Name ="line1"
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    Enabled = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =4
                    ListRows =20
                    ListWidth =7920
                    Left =4380
                    Top =300
                    Width =4392
                    Height =252
                    FontSize =9
                    TabIndex =2
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"\";\"\";\"\";\"\";\"10\";\"0\""
                    Name ="cboContact"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT tlu_Contacts.Contact_ID, [Last_Name] & (\", \"+[First_Name]) & (\" \"+[Mi"
                        "ddle_Init]), tlu_Contacts.Organization, tlu_Contacts.Position_title FROM tlu_Con"
                        "tacts ORDER BY Last_Name, First_Name; "
                    ColumnWidths ="0;2160;2880;2880"
                    AfterUpdate ="[Event Procedure]"
                    FontName ="Arial"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3480
                            Top =300
                            Width =708
                            Height =252
                            FontSize =9
                            FontWeight =700
                            Name ="labContact"
                            Caption ="Search:"
                            FontName ="Arial"
                        End
                    End
                End
                Begin OptionGroup
                    Locked = NotDefault
                    OverlapFlags =85
                    Left =1020
                    Top =120
                    Width =1980
                    Height =720
                    Name ="optgFilterContacts"
                    AfterUpdate ="[Event Procedure]"
                    DefaultValue ="0"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =300
                            Top =120
                            Width =552
                            Height =252
                            FontSize =9
                            FontWeight =700
                            BackColor =16777215
                            ForeColor =0
                            Name ="labFilterContacts"
                            Caption ="Filter:"
                            FontName ="Arial"
                        End
                        Begin OptionButton
                            OverlapFlags =87
                            Left =1140
                            Top =239
                            OptionValue =0
                            Name ="optFilterOff"
                            Begin
                                Begin Label
                                    OverlapFlags =215
                                    TextAlign =2
                                    Left =1380
                                    Top =180
                                    Width =1500
                                    Height =252
                                    FontSize =9
                                    BackColor =16777215
                                    ForeColor =0
                                    Name ="labFilterOff"
                                    Caption ="View all contacts"
                                    FontName ="Arial"
                                End
                            End
                        End
                        Begin OptionButton
                            OverlapFlags =87
                            Left =1140
                            Top =588
                            OptionValue =1
                            Name ="optFilterOn"
                            Begin
                                Begin Label
                                    OverlapFlags =215
                                    TextAlign =2
                                    Left =1380
                                    Top =528
                                    Width =1368
                                    Height =252
                                    FontSize =9
                                    BackColor =16777215
                                    ForeColor =0
                                    Name ="labFilterOn"
                                    Caption ="Filter by search"
                                    FontName ="Arial"
                                End
                            End
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    Left =1716
                    Top =3600
                    Width =7800
                    Height =864
                    FontSize =9
                    TabIndex =22
                    Name ="txtNotes"
                    ControlSource ="Contact_notes"
                    FontName ="Arial"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =588
                            Top =3600
                            Width =960
                            Height =252
                            FontSize =9
                            Name ="labNotes"
                            Caption ="Comments"
                            FontName ="Arial"
                        End
                    End
                End
                Begin ComboBox
                    OverlapFlags =85
                    Left =1716
                    Top =2160
                    Width =3180
                    Height =252
                    FontSize =9
                    TabIndex =10
                    ColumnInfo ="\"\";\"\";\"10\";\"100\""
                    Name ="cboOrganization"
                    ControlSource ="Organization"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT DISTINCT tlu_Contacts.Organization FROM tlu_Contacts ORDER BY tlu_Contact"
                        "s.Organization; "
                    FontName ="Arial"
                    OnGotFocus ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =396
                            Top =2160
                            Width =1152
                            Height =252
                            FontSize =9
                            FontWeight =700
                            Name ="labOrg"
                            Caption ="Organization"
                            FontName ="Arial"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    Left =1716
                    Top =1800
                    Width =2940
                    Height =252
                    FontSize =9
                    TabIndex =9
                    Name ="txtLastName"
                    ControlSource ="Last_name"
                    FontName ="Arial"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =540
                            Top =1800
                            Width =984
                            Height =252
                            FontSize =9
                            FontWeight =700
                            Name ="labLastName"
                            Caption ="Last name"
                            FontName ="Arial"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    SpecialEffect =0
                    OverlapFlags =85
                    TextAlign =1
                    Left =1740
                    Top =4680
                    Width =7800
                    Height =252
                    FontSize =9
                    TabIndex =23
                    BackColor =16777215
                    Name ="txtContactID"
                    ControlSource ="Contact_ID"
                    FontName ="Arial"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =660
                            Top =4680
                            Width =948
                            Height =252
                            FontSize =9
                            FontWeight =700
                            Name ="labContactID"
                            Caption ="Contact ID"
                            FontName ="Arial"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    Left =1716
                    Top =1080
                    Width =2940
                    Height =252
                    FontSize =9
                    TabIndex =7
                    Name ="txtFirstName"
                    ControlSource ="First_name"
                    FontName ="Arial"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =600
                            Top =1080
                            Width =948
                            Height =252
                            FontSize =9
                            FontWeight =700
                            Name ="labFirstName"
                            Caption ="First name"
                            FontName ="Arial"
                        End
                    End
                End
                Begin ComboBox
                    OverlapFlags =85
                    Left =1716
                    Top =2520
                    Width =3180
                    Height =252
                    FontSize =9
                    TabIndex =11
                    ColumnInfo ="\"\";\"\";\"10\";\"100\""
                    Name ="cboPosition"
                    ControlSource ="Position_title"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT DISTINCT tlu_Contacts.Position_title FROM tlu_Contacts ORDER BY tlu_Conta"
                        "cts.Position_title; "
                    FontName ="Arial"
                    OnGotFocus ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =420
                            Top =2520
                            Width =1128
                            Height =252
                            FontSize =9
                            Name ="labPosition"
                            Caption ="Position/title"
                            FontName ="Arial"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    Left =1692
                    Top =2880
                    Width =2088
                    Height =252
                    FontSize =9
                    TabIndex =12
                    Name ="txtWorkPhone"
                    ControlSource ="Work_Phone"
                    FontName ="Arial"
                    InputMask ="!\\(999\") \"000\\-0000;0;_"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =360
                            Top =2880
                            Width =1200
                            Height =252
                            FontSize =9
                            FontWeight =700
                            Name ="labWorkVoice"
                            Caption ="Work phone"
                            FontName ="Arial"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    Left =1692
                    Top =3240
                    Width =3288
                    Height =252
                    FontSize =9
                    TabIndex =14
                    Name ="txtEmail"
                    ControlSource ="Email_Address"
                    Format ="<"
                    FontName ="Arial"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =780
                            Top =3240
                            Width =756
                            Height =252
                            FontSize =9
                            Name ="labEmail"
                            Caption ="Email"
                            FontName ="Arial"
                        End
                    End
                End
                Begin Line
                    OverlapFlags =85
                    Left =360
                    Top =4560
                    Width =9300
                    Name ="line124"
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =7680
                    Top =960
                    Width =648
                    Height =300
                    FontSize =9
                    FontWeight =700
                    TabIndex =5
                    Name ="cmdUndo"
                    Caption ="Undo"
                    OnClick ="[Event Procedure]"
                    FontName ="Arial"
                    ControlTipText ="Undo all edits to this record"
                End
                Begin CommandButton
                    Enabled = NotDefault
                    OverlapFlags =85
                    Left =6300
                    Top =960
                    Width =1176
                    Height =300
                    FontSize =9
                    FontWeight =700
                    TabIndex =4
                    Name ="cmdNew"
                    Caption ="New record"
                    OnClick ="[Event Procedure]"
                    FontName ="Arial"
                    ControlTipText ="Create a new program record"
                End
                Begin CommandButton
                    Enabled = NotDefault
                    OverlapFlags =85
                    Left =4980
                    Top =960
                    Width =1140
                    Height =300
                    FontSize =9
                    FontWeight =700
                    TabIndex =3
                    Name ="cmdEdit"
                    Caption ="Edit record"
                    OnClick ="[Event Procedure]"
                    FontName ="Arial"
                    ControlTipText ="Edit the information for the selected program"
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =8640
                    Top =960
                    Width =780
                    Height =300
                    FontSize =9
                    FontWeight =700
                    TabIndex =6
                    Name ="cmdSubmit"
                    Caption ="Done"
                    OnClick ="[Event Procedure]"
                    FontName ="Arial"
                    ControlTipText ="Submit edits to this record"
                End
                Begin TextBox
                    OverlapFlags =85
                    Left =1716
                    Top =1440
                    Width =1104
                    Height =252
                    FontSize =9
                    TabIndex =8
                    Name ="txtMiddleInit"
                    ControlSource ="Middle_init"
                    FontName ="Arial"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =360
                            Top =1440
                            Width =1164
                            Height =252
                            FontSize =9
                            Name ="labMiddleInit"
                            Caption ="Middle initial"
                            FontName ="Arial"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    Left =4296
                    Top =2880
                    Width =684
                    Height =252
                    FontSize =9
                    TabIndex =13
                    Name ="txtWorkExt"
                    ControlSource ="Work_Extension"
                    FontName ="Arial"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =3840
                            Top =2880
                            Width =360
                            Height =252
                            FontSize =9
                            Name ="labWorkExt"
                            Caption ="ext"
                            FontName ="Arial"
                        End
                    End
                End
                Begin ComboBox
                    RowSourceTypeInt =1
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =6255
                    Top =1440
                    Width =3225
                    TabIndex =15
                    Name ="cboAddressType"
                    ControlSource ="Address_Type"
                    RowSourceType ="Value List"
                    RowSource ="\"Mailing\";\"Physical\";\"Physical and Mailing\""
                    StatusBarText ="M. Address (mailing, physical, both) type (addrtype)"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =4980
                            Top =1440
                            Width =1095
                            Height =240
                            Name ="Label127"
                            Caption ="Address Type"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =6240
                    Top =2160
                    Width =3240
                    TabIndex =17
                    Name ="txtAddress2"
                    ControlSource ="Address2"
                    StatusBarText ="M. Street address (cntaddr)"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =5220
                            Top =2160
                            Width =840
                            Height =240
                            Name ="Label130"
                            Caption ="Address 2"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =6240
                    Top =2520
                    Width =3240
                    TabIndex =18
                    Name ="txtCity"
                    ControlSource ="City"
                    StatusBarText ="M. City or town (city)"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =5640
                            Top =2520
                            Width =420
                            Height =240
                            Name ="lblCity"
                            Caption ="City"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =6240
                    Top =2880
                    Width =720
                    TabIndex =19
                    Name ="txtStateCode"
                    ControlSource ="State_Code"
                    StatusBarText ="M. State or province (state)"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =5100
                            Top =2880
                            Width =960
                            Height =240
                            Name ="Label132"
                            Caption ="State Code"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =7980
                    Top =2880
                    Width =1500
                    TabIndex =20
                    Name ="txtZipCode"
                    ControlSource ="Zip_Code"
                    StatusBarText ="M. Zip code (postal)"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =7200
                            Top =2880
                            Width =720
                            Height =240
                            Name ="Label133"
                            Caption ="Zip Code"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =6240
                    Top =3240
                    Width =3240
                    TabIndex =21
                    Name ="txtCountry"
                    ControlSource ="Country"
                    StatusBarText ="M. Country (country)"
                    DefaultValue ="\"USA\""
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =5400
                            Top =3240
                            Width =660
                            Height =240
                            Name ="Label134"
                            Caption ="Country"
                        End
                    End
                End
                Begin ComboBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =2
                    ListWidth =5760
                    Left =6240
                    Top =1800
                    Width =3240
                    TabIndex =16
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"10\";\"100\""
                    Name ="cboAddress"
                    ControlSource ="Address"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT DISTINCT tlu_Contacts.Address, [Address] & (\" \"+[Address2]) & (\", \"+["
                        "City]) & (\", \"+[State_Code]) & (\"  \"+[Zip_Code]) & (\", \"+[Country]) AS Ful"
                        "lAddress FROM tlu_Contacts WHERE Address IS NOT NULL; "
                    ColumnWidths ="144;5760"
                    StatusBarText ="M. Street address (cntaddr)"
                    AfterUpdate ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =5220
                            Top =1800
                            Width =840
                            Height =240
                            Name ="Label128"
                            Caption ="Address 1"
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
' FORM NAME:    frm_Contacts
' Description:  Standard module to view and edit contact information
' Data source:  tlu_Project_Crew
' Data access:  no edits, additions, or deletions unless properties are changed
'               (see fxnFormDefaults)
' Pages:        none
' Functions:    fxnFormDefaults, fxnBuildID, fxnValidate
' References:   fxnSwitchboardIsOpen, fxnChangeDelimiter, fxnTrimSpaces
' Source/date:  John R. Boetsch, 2002
' Revisions:    JRB, May 25, 2006 - documentation, changed validation, and combined
'                   what was previously in a subform into a single form
'               Simon D. Kingston, 9/18/2006 - removed home phone, mobile phone, and audit info.; added address type,
'                   address1, address 2, city, state, zip, country
'               SDK, 9/22/2006 - added Close event code to update contact drop-down lists on various forms
'               SDK, 9/27/2006 - replaced form level variable to check if no records with dynamic checks when needed
'               SDK, 9/28/2006 - removed fxnBuildID since I'm not using natural key for Contact_ID
' ================================

Private Sub cboAddress_AfterUpdate()
' Description:  Allows addresses to be selected from previous entries instead of entering by hand repeatedly
' Parameters:   none
' Returns:      none
' Throws:       none
' References:   none
' Source/date:  Simon Kingston, Sept. 2006
' Revisions:
Dim strSQL As String
Dim rst As DAO.Recordset

strSQL = "SELECT Address, Address2, City, State_Code, Zip_Code, country FROM tlu_Contacts "
strSQL = strSQL & "WHERE Address=" & CorrectText(Me.cboAddress) & ";"

Set rst = CurrentDb.OpenRecordset(strSQL, dbOpenForwardOnly)
If Not (rst.EOF And rst.BOF) And IsNull(Me.txtAddress2) And IsNull(Me.txtCity) And IsNull(Me.txtStateCode) And IsNull(Me.Zip_Code) Then
    Me.txtAddress2 = rst!Address2
    Me.txtCity = rst!City
    Me.State_Code = rst!State_Code
    Me.Zip_Code = rst!Zip_Code
    Me.Country = rst!Country
    Me.txtAddress2.Requery
    Me.txtCity.Requery
    Me.txtStateCode.Requery
    Me.txtZipCode.Requery
    Me.txtCountry.Requery
End If

rst.Close
Set rst = Nothing
End Sub

Private Sub Form_Close()
' Description:  update all the contact drop-down lists that may be open, so that new contacts are available to choose
' Parameters:   none
' Returns:      none
' Throws:       none
' References:   none
' Source/date:  Simon Kingston, Sept. 2006
' Revisions:
Dim strFormName As String

strFormName = "frm_Data_Entry"
If IsLoaded(strFormName) Then
    Forms(strFormName)!subObservers.Form!cmbContact_ID.Requery
End If

strFormName = "frm_Set_Defaults"
If IsLoaded(strFormName) Then
    Forms(strFormName)!cmbUser.Requery
End If
    
End Sub

Private Sub Form_Open(Cancel As Integer)
    On Error GoTo Err_Handler

    ' Count the records in the recordset to determine the form settings
    If DCount("*", "tlu_Contacts") = 0 Then
        ' If no records, set view to new
        fxnFormDefaults ("new")
    Else
        ' Set to filter view depending on the opening arguments
        Select Case Me.OpenArgs
            Case "new"
                fxnFormDefaults ("new")
            Case ""
                fxnFormDefaults ("view")
            Case Is <> ""
                Me.cboContact = Me.OpenArgs
                Me.optgFilterContacts = 1
                optgFilterContacts_AfterUpdate
            Case Else
                fxnFormDefaults ("view")
        End Select
    End If

Exit_Procedure:
    Exit Sub

Err_Handler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    Resume Exit_Procedure

End Sub

Private Sub Form_Current()
    On Error GoTo Err_Handler

If Me.NewRecord Then
    If GetDataType("tlu_Contacts", "Contact_ID") = dbText Then
        Me!Contact_ID = fxnGUIDGen
    End If
End If

    ' On moving to a different record, turn off the filter and update the
    '   contact selector
    If Me.optgFilterContacts = 0 Then
        Me.FilterOn = False
        Me.cboContact.Enabled = False
        Me.cboContact = Me.txtContactID
    ElseIf Me.optgFilterContacts = 1 And Me.cboContact.Enabled Then
        Me.cboContact.SetFocus
    End If

Exit_Procedure:
    Exit Sub

Err_Handler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    Resume Exit_Procedure

End Sub

Private Sub Form_BeforeUpdate(Cancel As Integer)
    On Error GoTo Err_Handler

    ' Validate the record and cancel updates if not valid
    If fxnValidate Then
        If IsNothing(Me!Contact_ID) Then
            If GetDataType("tlu_Contacts", "Contact_ID") = dbText Then
                Me.Contact_ID = fxnGUIDGen
            End If
        End If
    Else
        DoCmd.CancelEvent
    End If

Exit_Procedure:
    Exit Sub

Err_Handler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    Resume Exit_Procedure

End Sub

Private Sub cmdClose_Click()
    On Error GoTo Err_Handler

    If fxnValidate Then
        ' Close the form and requery the contact list in the referring form
        DoCmd.Close , , acSaveNo
    End If

Exit_Procedure:
    Exit Sub

Err_Handler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    Resume Exit_Procedure

End Sub

Private Sub cboContact_GotFocus()
    On Error GoTo Err_Handler

    ' Requery the control once it gets the focus
    Me.cboContact.Requery

Exit_Procedure:
    Exit Sub

Err_Handler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    Resume Exit_Procedure

End Sub

Private Sub cboContact_NotInList(NewData As String, Response As Integer)
    On Error GoTo Err_Handler

    Me.ActiveControl.Undo

Exit_Procedure:
    Exit Sub

Err_Handler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    Resume Exit_Procedure

End Sub

Private Sub cboContact_AfterUpdate()
    On Error GoTo Err_Handler

    ' If a name has been selected, filter the form to the selected ID
    If IsNull(Me.cboContact) = False Then
        Me.optgFilterContacts = 1
        SetFilter
    End If

Exit_Procedure:
    Exit Sub

Err_Handler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    Resume Exit_Procedure

End Sub

Private Sub optgFilterContacts_AfterUpdate()
    On Error GoTo Err_Handler

    If Me.optgFilterContacts = 0 Then
        Me.cboContact.Enabled = False
        Me.FilterOn = False
    ' Or connect the subform to view only the record related to the selected contact
    ElseIf Me.optgFilterContacts = 1 Then
        Me.cboContact.Enabled = True
        If IsNull(Me.cboContact) = False Then
            SetFilter
        End If
        Me.cboContact.SetFocus
    End If

Exit_Procedure:
    Exit Sub

Err_Handler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    Resume Exit_Procedure

End Sub

Private Sub cmdEdit_Click()
    On Error GoTo Err_Handler

    ' Set the current data mode to edit and reset the form settings accordingly
    fxnFormDefaults ("edit")

Exit_Procedure:
    Exit Sub

Err_Handler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    Resume Exit_Procedure

End Sub

Private Sub cmdNew_Click()
    On Error GoTo Err_Handler

    ' Set the current data mode to new and reset the form settings accordingly
    fxnFormDefaults ("new")

Exit_Procedure:
    Exit Sub

Err_Handler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    Resume Exit_Procedure

End Sub

Private Sub cmdUndo_Click()
    On Error GoTo Err_Handler

    ' Undo changes to the current record and restore the form settings
    '   for the current data mode
    Me.Undo
    ' Switch back to view mode
    fxnFormDefaults ("view")
    Me.optgFilterContacts.SetFocus

Exit_Procedure:
    Exit Sub

Err_Handler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    Resume Exit_Procedure

End Sub

Private Sub cmdSubmit_Click()
    On Error GoTo Err_Handler

    If fxnValidate Then
        ' Save edits
        DoCmd.RunCommand acCmdSaveRecord
        ' Reset form to view mode
        Me.cboContact.Requery
        fxnFormDefaults ("view")
        Me.optgFilterContacts.SetFocus
    End If

Exit_Procedure:
    Exit Sub

Err_Handler:
    Select Case Err.Number
        Case 2046
            Resume Next
        Case Else
            MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
            Resume Exit_Procedure
    End Select

End Sub

Private Sub cboOrganization_GotFocus()
    On Error GoTo Err_Handler

    ' Requery the recursive lookup combo box
    Me.ActiveControl.Requery

Exit_Procedure:
    Exit Sub

Err_Handler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    Resume Exit_Procedure

End Sub

Private Sub cboPosition_GotFocus()
    On Error GoTo Err_Handler

    ' Requery the recursive lookup combo box
    Me.ActiveControl.Requery

Exit_Procedure:
    Exit Sub

Err_Handler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    Resume Exit_Procedure

End Sub

' =================================
' FUNCTION:     fxnFormDefaults
' Description:  Sets properties of the form depending on the form mode
' Parameters:   strFormMode - form mode (view, edit, new)
' Returns:      none
' Throws:       none
' References:   none
' Source/date:  John R. Boetsch, 2002
' Revisions:    JRB, May 25, 2006 - documentation, updated code for enabling/disabling
'               controls
'               SDK, 9/27/2006 - added Me.DataEntry = False line to "view" to prevent errors when undo clicked on new record
'                              - removed line to lock txtContactID since I am not using natural key there is no need to ever unlock it
'                              - added dynamic record count to replace form level variable record count
' =================================

Function fxnFormDefaults(strFormMode As String)
    On Error GoTo Err_Handler

    Dim bLockState As Boolean

    bLockState = True   ' Default is to lock controls

    ' The form should not be reset to view or edit mode if there are no records
    If DCount("*", "tlu_Contacts") = 0 Then strFormMode = "new"

    ' Change the form properties depending on the mode specified by the user
    Select Case strFormMode
    Case "new"
    ' Modify the form properties to allow new records
        Me.cmdClose.SetFocus    ' Must do this before turning off new button
        Me.cmdUndo.Visible = True
        Me.cmdSubmit.Visible = True
        Me.cmdEdit.Enabled = False
        Me.cmdNew.Enabled = False
        Me.AllowAdditions = True
        Me.Detail.BackColor = 12574431 ' haystack
        DoCmd.GoToRecord , , acNewRec
        Me.txtFirstName.SetFocus    ' Needed on new record before disabling ctls
    ' Unlock fields
        bLockState = False
        GoTo Change_Ctl_State

    Case "edit"
    ' Modify the form properties to allow edits
        Me.cmdClose.SetFocus    ' Must do this before turning off edit button
        Me.cmdUndo.Visible = True
        Me.cmdSubmit.Visible = True
        Me.cmdEdit.Enabled = False
        Me.cmdNew.Enabled = False
        Me.AllowAdditions = True
        Me.Detail.BackColor = 12574431 ' haystack
    ' Unlock fields
        bLockState = False
        GoTo Change_Ctl_State

    Case "view"
    ' Set the form to the default form view
        Me.cmdClose.SetFocus    ' Must do this before disabling ctls
        Me.cmdUndo.Visible = False
        Me.cmdSubmit.Visible = False
        Me.cmdEdit.Enabled = True
        Me.cmdNew.Enabled = True
        Me.DataEntry = False
        Me.AllowAdditions = False
        Me.Detail.BackColor = 14541277 ' light blue (default)
    ' Lock fields
        bLockState = True
        GoTo Change_Ctl_State

    End Select

Change_Ctl_State:
    Me.optgFilterContacts.Locked = Not bLockState
    Me.cboContact.Enabled = bLockState
    Me.txtFirstName.Locked = bLockState
    Me.txtLastName.Locked = bLockState
    Me.txtMiddleInit.Locked = bLockState
    Me.cboOrganization.Locked = bLockState
    Me.cboPosition.Locked = bLockState
    Me.txtWorkPhone.Locked = bLockState
    Me.txtWorkExt.Locked = bLockState
    Me.txtEmail.Locked = bLockState
    Me.cboAddressType.Locked = bLockState
    Me.cboAddress.Locked = bLockState
    Me.txtAddress2.Locked = bLockState
    Me.txtCity.Locked = bLockState
    Me.txtStateCode.Locked = bLockState
    Me.txtZipCode.Locked = bLockState
    Me.txtCountry.Locked = bLockState
    Me.txtNotes.Locked = bLockState

Exit_Procedure:
    Exit Function

Err_Handler:
    Select Case Err.Number
        Case Else
            MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
                "Error encountered (fxnFormDefaults)"
    End Select
    Resume Exit_Procedure

End Function

' =================================
' FUNCTION:     fxnValidate
' Description:  Validate the record prior to saving, closing or moving to another record
' Parameters:   none
' Returns:      True if the record passes validation rules, or False
' Throws:       none
' References:   none
' Source/date:  John R. Boetsch, May 25, 2006
' Revisions:    <name, date, desc - add lines as you go>
' =================================

Function fxnValidate() As Boolean
    On Error GoTo Err_Handler

    Dim strErrorMsg As String

    strErrorMsg = "There is already a contact with the same ID." & _
        "If this record is for a different person with a similar name, " & _
        "either add a middle initial if there is none, or " & _
        "make a minor modification to the ID below.  Otherwise UNDO edits."

    ' Make sure the information is valid before updating the record
    If Me.Dirty = True Then
        ' If information for a new contact has been entered,
        '  verify that the critical data elements have been completed before saving
        If IsNull(Me!Last_Name) Then
            MsgBox "Fill in the last name", vbOKOnly, "Validation error"
            Me!txtLastName.SetFocus
            GoTo Exit_Procedure
        ElseIf IsNull(Me!First_Name) Then
            MsgBox "Fill in the first name", vbOKOnly, "Validation error"
            Me!txtFirstName.SetFocus
            GoTo Exit_Procedure
        ElseIf IsNull(Me!Organization) Then
            MsgBox "Fill in the employer/organization of the contact", vbOKOnly, _
                "Validation error"
            Me!cboOrganization.SetFocus
            GoTo Exit_Procedure
        End If
    End If

    fxnValidate = True

Exit_Procedure:
    Exit Function

Err_Handler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
                "Error encountered (fxnValidate)"
    Resume Exit_Procedure

End Function

Public Sub SetFilter()
Dim strCriteria As String

strCriteria = GetCriteriaString("[Contact_ID]=", "tlu_Contacts", "Contact_ID", Me.name, "cboContact")
Me.Filter = strCriteria
Me.FilterOn = True
End Sub
