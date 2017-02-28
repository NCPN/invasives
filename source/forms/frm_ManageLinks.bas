Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    ViewsAllowed =1
    TabularFamily =127
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =11700
    DatasheetFontHeight =9
    ItemSuffix =9
    Left =1215
    Top =285
    Right =12915
    Bottom =7935
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0xdb9d0340740be340
    End
    RecordSource ="tsys_Link_Files"
    Caption ="Manage Linked Tables"
    OnCurrent ="[Event Procedure]"
    OnClose ="[Event Procedure]"
    DatasheetFontName ="Arial"
    AllowDatasheetView =0
    AllowPivotTableView =0
    AllowPivotChartView =0
    AllowPivotChartView =0
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
        Begin TextBox
            FELineBreak = NotDefault
            SpecialEffect =2
            OldBorderStyle =0
            FontName ="Tahoma"
            AsianLineBreak =255
        End
        Begin Subform
            SpecialEffect =2
        End
        Begin Section
            CanGrow = NotDefault
            Height =7860
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin Label
                    OverlapFlags =85
                    Left =60
                    Top =60
                    Width =2235
                    Height =285
                    FontSize =10
                    FontWeight =700
                    Name ="Label0"
                    Caption ="Manage Linked Tables"
                    FontName ="Arial"
                End
                Begin TextBox
                    OverlapFlags =215
                    IMESentenceMode =3
                    Left =840
                    Top =1140
                    ColumnWidth =1365
                    Name ="txtLink_type"
                    ControlSource ="Link_type"
                    StatusBarText ="Back-end database type"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =60
                            Top =1140
                            Width =795
                            Height =270
                            FontSize =9
                            Name ="Label1"
                            Caption ="Link type"
                            FontName ="Arial"
                        End
                    End
                End
                Begin CommandButton
                    TabStop = NotDefault
                    OverlapFlags =85
                    Left =7200
                    Top =1740
                    Width =888
                    Height =300
                    FontSize =9
                    FontWeight =700
                    TabIndex =1
                    Name ="cmdBrowse"
                    Caption ="Browse"
                    OnClick ="[Event Procedure]"
                    FontName ="Arial"
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    CanShrink = NotDefault
                    TabStop = NotDefault
                    FELineBreak = NotDefault
                    SpecialEffect =0
                    OverlapFlags =85
                    Left =840
                    Top =1740
                    Width =6123
                    FontSize =9
                    TabIndex =2
                    Name ="txtCurrentName"
                    ControlSource ="Link_file_name"
                    FontName ="Arial"
                    AsianLineBreak =0
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =60
                            Top =1560
                            Width =696
                            Height =444
                            FontSize =9
                            Name ="labCurrentName"
                            Caption ="Current name:"
                            FontName ="Arial"
                        End
                    End
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    TabStop = NotDefault
                    FELineBreak = NotDefault
                    SpecialEffect =0
                    OverlapFlags =85
                    Left =840
                    Top =2100
                    Width =7977
                    FontSize =9
                    TabIndex =3
                    Name ="txtCurrentPath"
                    ControlSource ="Link_file_path"
                    StatusBarText ="Current linked file path"
                    FontName ="Arial"
                    AsianLineBreak =0
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =240
                            Top =2100
                            Width =540
                            Height =240
                            FontSize =9
                            Name ="labCurrentPath"
                            Caption ="Path:"
                            FontName ="Arial"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1500
                    Top =2460
                    Width =7320
                    ColumnWidth =2610
                    TabIndex =4
                    Name ="txtLink_description"
                    ControlSource ="Link_description"
                    StatusBarText ="Describes the types of data tables included in the link"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =60
                            Top =2460
                            Width =1380
                            Height =270
                            FontSize =9
                            Name ="Label2"
                            Caption ="Link description"
                            FontName ="Arial"
                        End
                    End
                End
                Begin Label
                    OverlapFlags =85
                    Left =60
                    Top =420
                    Width =8880
                    Height =600
                    Name ="Label3"
                    Caption ="Use this form to manage links to tables in other databases.  You can determine w"
                        "hich database file to link to and which tables in that database to link.  You ca"
                        "n add new links and delete existing links also.  To simply re-link to an existin"
                        "g database, use the same general table linking form as the users."
                End
                Begin Subform
                    OverlapFlags =87
                    Left =120
                    Top =3120
                    Width =11445
                    Height =4365
                    TabIndex =5
                    Name ="subLinkedTables"
                    SourceObject ="Form.fsub_LinkedTables"
                    LinkChildFields ="Link_type"
                    LinkMasterFields ="Link_type"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =120
                            Top =2880
                            Width =1035
                            Height =240
                            Name ="Linked Tables Label"
                            Caption ="Linked Tables"
                            EventProcPrefix ="Linked_Tables_Label"
                        End
                    End
                End
                Begin CommandButton
                    TabStop = NotDefault
                    OverlapFlags =85
                    Left =9360
                    Top =2460
                    Width =1083
                    Height =300
                    FontSize =9
                    FontWeight =700
                    TabIndex =6
                    Name ="cmdAddTables"
                    Caption ="Add Tables"
                    OnClick ="[Event Procedure]"
                    FontName ="Arial"
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

Private Sub cmdAddTables_Click()

If IsNothing(Me!txtLink_type) Then
    MsgBox "You must enter a link type before you can add tables!", vbCritical + vbOKOnly, "Enter Link Type"
Else
    If IsNothing(Me!txtCurrentPath) Then
        MsgBox "You must select a file before you can add tables!", vbCritical + vbOKOnly, "Select File"
    Else
        DoCmd.RunCommand acCmdSaveRecord
        DoCmd.OpenForm "frm_SelectTables", , , , , acDialog, XML_Tag("FileName", Me!txtCurrentPath) & XML_Tag("LinkType", Me!txtLink_type)
        Me!subLinkedTables.Requery
    End If
End If
End Sub

Private Sub cmdBrowse_Click()
Dim varFileName As Variant
Dim strfilter As String

strfilter = adhAddFilterItem(strfilter, "MS Access databases", "*.mdb")

varFileName = adhCommonFileOpenSave(, GetPath(CurrentDb.name), strfilter, , , , "Select MS Access Database", True)

Me!txtCurrentPath = varFileName
Me!txtCurrentName = GetFileName(Nz(varFileName, ""))

Me!cmdAddTables.Enabled = Not IsNothing(Me!txtCurrentPath)
End Sub

Private Sub Form_Close()
Dim strSQL As String
Dim rst As DAO.Recordset
Dim strMessage As String

On Error GoTo Err_Handler

strSQL = "SELECT tsys_Link_Files.Link_type, Count(tsys_Link_Tables.Link_table) AS CountOfLink_table "
strSQL = strSQL & "FROM tsys_Link_Files LEFT OUTER JOIN tsys_Link_Tables ON tsys_Link_Files.Link_type = tsys_Link_Tables.Link_type "
strSQL = strSQL & "GROUP BY tsys_Link_Files.Link_type;"

Set rst = CurrentDb.OpenRecordset(strSQL, dbOpenSnapshot)

With rst
    Do Until .EOF
        If !CountOfLink_table = 0 Then
            strMessage = !Link_type & " link file record to "
            strMessage = strMessage & DLookup("[Link_file_name]", "tsys_Link_Files", "[Link_type]=" & CorrectText(!Link_type))
            strMessage = strMessage & " will be deleted since there are no longer any tables linked from that file."
            MsgBox strMessage, vbOKOnly + vbInformation, "Deleting Link File Record"
            CurrentDb.Execute "DELETE * FROM tsys_Link_Files WHERE Link_type=" & CorrectText(!Link_type) & ";"
        End If
        .MoveNext
    Loop
End With

Exit_Handler:
    On Error Resume Next
    rst.Close
    Set rst = Nothing
    Exit Sub

Err_Handler:
    MsgBox Err.Number & " - " & Err.Description
    Resume Exit_Handler

End Sub

Private Sub Form_Current()
Me!cmdAddTables.Enabled = Not IsNothing(Me!txtCurrentPath)
End Sub
