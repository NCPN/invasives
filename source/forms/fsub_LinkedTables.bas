Version =20
VersionRequired =20
Begin Form
    AutoCenter = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =2
    TabularFamily =0
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =6300
    DatasheetFontHeight =9
    ItemSuffix =20
    Left =3645
    Top =5790
    Right =14820
    Bottom =9885
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0xb70bcdfd760be340
    End
    RecordSource ="tsys_Link_Tables"
    Caption ="Linked Tables"
    OnDelete ="[Event Procedure]"
    AfterUpdate ="[Event Procedure]"
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
            Height =0
            BackColor =-2147483633
            Name ="FormHeader"
        End
        Begin Section
            Height =4095
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1680
                    Top =120
                    Width =4560
                    Height =450
                    ColumnWidth =1860
                    Name ="txtLink_table"
                    ControlSource ="Link_table"
                    StatusBarText ="Table name"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =60
                            Top =120
                            Width =1560
                            Height =255
                            Name ="Link_table_Label"
                            Caption ="Link_table"
                        End
                    End
                End
                Begin ComboBox
                    RowSourceTypeInt =1
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1680
                    Top =660
                    Width =2310
                    Height =255
                    ColumnWidth =1080
                    TabIndex =1
                    Name ="cboTable_type"
                    ControlSource ="Table_type"
                    RowSourceType ="Value List"
                    RowSource ="project;standard;template"
                    StatusBarText ="Type of table (e.g., project table, standard table, etc.)"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =60
                            Top =660
                            Width =1560
                            Height =255
                            Name ="Table_type_Label"
                            Caption ="Table_type"
                        End
                    End
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1680
                    Top =1020
                    Width =4560
                    Height =450
                    ColumnWidth =3000
                    TabIndex =2
                    Name ="txtDescription_text"
                    ControlSource ="Description_text"
                    StatusBarText ="Table description"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =60
                            Top =1020
                            Width =1560
                            Height =255
                            Name ="Description_text_Label"
                            Caption ="Description_text"
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =85
                    Left =1680
                    Top =1560
                    ColumnWidth =996
                    TabIndex =3
                    Name ="chkIs_hidden"
                    ControlSource ="Is_hidden"
                    StatusBarText ="Indicates whether the table should be hidden in database view"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =60
                            Top =1560
                            Width =1560
                            Height =255
                            Name ="Is_hidden_Label"
                            Caption ="Is_hidden"
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =85
                    Left =1680
                    Top =1920
                    ColumnWidth =1785
                    TabIndex =4
                    Name ="chkAllow_edits_lookup"
                    ControlSource ="Allow_edits_lookup"
                    StatusBarText ="Indicates whether the table should be available for user edits in the lookup bro"
                        "wser"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =60
                            Top =1920
                            Width =1560
                            Height =255
                            Name ="Allow_edits_lookup_Label"
                            Caption ="Allow_edits_lookup"
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =85
                    Left =1680
                    Top =2280
                    ColumnWidth =1344
                    TabIndex =5
                    Name ="chkBrowser_view"
                    ControlSource ="Browser_view"
                    StatusBarText ="Indicates whether the table is available for viewing with the data browser"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =60
                            Top =2280
                            Width =1560
                            Height =255
                            Name ="Browser_view_Label"
                            Caption ="Browser_view"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1680
                    Top =3720
                    Width =465
                    Height =255
                    ColumnWidth =1050
                    TabIndex =6
                    Name ="Sort_order"
                    ControlSource ="Sort_order"
                    StatusBarText ="Sort order for table listings"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =60
                            Top =3720
                            Width =1560
                            Height =255
                            Name ="Sort_order_Label"
                            Caption ="Sort_order"
                        End
                    End
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

Private Sub Form_AfterUpdate()
Application.SetHiddenAttribute acTable, Me!Link_table, Me.chkIs_hidden
End Sub

Private Sub Form_Delete(Cancel As Integer)
Dim strSQL As String

If MsgBox("This action will delete the table link from the database! Are you sure you wish to continue?", vbYesNo + vbExclamation, "Delete Table Link?") = vbYes Then
    CurrentDb.TableDefs.Delete Me.Link_table
    CurrentDb.TableDefs.Refresh
    strSQL = "DELETE * FROM tsys_Link_Tables WHERE Link_Table=" & CorrectText(Me!Link_table) & " AND Link_Type=" & CorrectText(Me!Link_type) & ";"
    CurrentDb.Execute strSQL
    Cancel = True
Else
    Cancel = True
End If

End Sub
