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
    Width =5040
    DatasheetFontHeight =9
    ItemSuffix =7
    Left =5625
    Top =2940
    Right =11010
    Bottom =7140
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x9d2210c6b41ee340
    End
    Caption ="Select for Plot Revisit Data Sheet"
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
            Height =3600
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
                    Left =705
                    Top =540
                    Width =3585
                    Height =360
                    FontSize =12
                    FontWeight =700
                    Name ="Label2"
                    Caption ="Plot Revisit Data Sheet"
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =600
                    Top =2580
                    Width =1395
                    Height =300
                    TabIndex =1
                    Name ="ButtonClose"
                    Caption ="Close Form"
                    OnClick ="[Event Procedure]"
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =2880
                    Top =2580
                    Height =300
                    TabIndex =2
                    Name ="Button Print"
                    Caption ="Generate Report"
                    OnClick ="[Event Procedure]"
                    EventProcPrefix ="Button_Print"
                End
                Begin ComboBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    ListWidth =540
                    Left =2280
                    Top =1800
                    Width =840
                    TabIndex =3
                    ColumnInfo ="\"\";\"\";\"10\";\"100\""
                    Name ="Plot_ID"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT tbl_Locations.Location_ID, tbl_Locations.Plot_ID FROM tbl_Locations; "
                    ColumnWidths ="540"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =1560
                            Top =1800
                            Width =645
                            Height =245
                            FontWeight =700
                            Name ="Plot ID_Label"
                            Caption ="Plot ID"
                            EventProcPrefix ="Plot_ID_Label"
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

  Me!Plot_ID = Null
  If Not IsNull(Me!Park_Code) Then
    Me!Plot_ID.RowSource = "SELECT Plot_ID FROM tbl_locations WHERE [Unit_Code] = '" & Me!Park_Code & "' ORDER BY Plot_ID"
    Me!Plot_ID.Requery
  Else
    MsgBox "You must select a park!"
  End If
    
End Sub
Private Sub Button_Print_Click()
On Error GoTo Err_Button_Print
' Generate Plot Establishment sheets for selected park.
' Russ DenBleyker - Northern Colorado Plateau Network April, 2008.
' Added Forest Woodlands June, 2009.
' Added revisit comments, March, 2010.

    Dim objWord As Word.Application
    Dim fld As Field
    Dim rst As Recordset
    Dim Mrst As Recordset
    Dim cat As ADOX.Catalog
    Dim tbl_Work_Surface_Type As ADOX.table
    Dim RevisitComments As Recordset
    Dim strSQL As String
    Dim MArray(12) As String
    Dim intIndex As Integer
    Dim strFieldName As String
    Dim intResponse As Integer
    Dim strPrompt As String
    
    ' Initialize monument tree array
    MArray(0) = "T1-Oa"
    MArray(1) = "T1-Ob"
    MArray(2) = "T1-Oc"
    MArray(3) = "T1-Ea"
    MArray(4) = "T1-Eb"
    MArray(5) = "T1-Ec"
    MArray(6) = "T3-Oa"
    MArray(7) = "T3-Ob"
    MArray(8) = "T3-Oc"
    MArray(9) = "T3-Ea"
    MArray(10) = "T3-Eb"
    MArray(11) = "T3-Ec"
    
    If IsNull(Me!Park_Code) Then
      MsgBox "You must select a park!"
      GoTo Exit_Button_Print
    End If
    If IsNull(Me!Plot_ID) Then
      MsgBox "You must select a plot!"
      GoTo Exit_Button_Print
    End If
    
    DoCmd.Hourglass True
    
    ' Launch Word and load the report template
    Set objWord = New Word.Application
    objWord.Documents.Add _
     Application.CurrentProject.Path & "\Plot_Establishment.dot"
    objWord.Visible = True
    
    ' Build main SQL string for details
     
     strSQL = "SELECT * FROM tbl_locations WHERE [Unit_Code] = '" & Me!Park_Code & "' AND [Plot_ID] = " & Me!Plot_ID
    ' Get the database record.
    Set rst = New ADODB.Recordset
    rst.Open strSQL, CurrentProject.Connection
    If rst.EOF Then
      MsgBox "Record not found."
      GoTo Exit_Button_Print
    End If

    rst.MoveFirst

 ' Park_Name = DLookup("[ParkName]", "tlu_Parks", "[ParkCode] = '" & rst.Fields("Button_Print") & "'")

    ' Update Word template
    With objWord.ActiveDocument.Bookmarks
      If Not IsNull(rst.Fields("Unit_Code")) Then
        .item("Unit_Code").Range.text = rst.Fields("Unit_Code")
      End If
      If Not IsNull(rst.Fields("Plot_ID")) Then
        .item("Plot_ID").Range.text = rst.Fields("Plot_ID")
      End If
      If Not IsNull(rst.Fields("E_Coord")) Then
        .item("E_Coord").Range.text = rst.Fields("E_Coord")
      End If
      If Not IsNull(rst.Fields("N_Coord")) Then
        .item("N_Coord").Range.text = rst.Fields("N_Coord")
      End If
      If Not IsNull(rst.Fields("UTM_Zone")) Then
        .item("UTM_Zone").Range.text = rst.Fields("UTM_Zone")
      End If
      If Not IsNull(rst.Fields("Plot_Slope")) Then
        .item("Plot_Slope").Range.text = " " & rst.Fields("Plot_Slope")
      End If
      If Not IsNull(rst.Fields("Plot_Aspect")) Then
        .item("Plot_Aspect").Range.text = " " & rst.Fields("Plot_Aspect")
      End If
      If Not IsNull(rst.Fields("Datum")) Then
        .item("Datum").Range.text = rst.Fields("Datum")
      End If
      If Not IsNull(rst.Fields("Azimuth")) Then
        .item("Azimuth").Range.text = " " & rst.Fields("Azimuth")
      End If
      If Not IsNull(rst.Fields("T1O_UTME")) Then
        .item("T1O_UTME").Range.text = " " & rst.Fields("T1O_UTME")
      End If
      If Not IsNull(rst.Fields("T1O_UTMN")) Then
        .item("T1O_UTMN").Range.text = " " & rst.Fields("T1O_UTMN")
      End If
      If Not IsNull(rst.Fields("T2O_UTME")) Then
        .item("T2O_UTME").Range.text = " " & rst.Fields("T2O_UTME")
      End If
      If Not IsNull(rst.Fields("T2O_UTMN")) Then
        .item("T2O_UTMN").Range.text = " " & rst.Fields("T2O_UTMN")
      End If
      If Not IsNull(rst.Fields("T3O_UTME")) Then
        .item("T3O_UTME").Range.text = " " & rst.Fields("T3O_UTME")
      End If
      If Not IsNull(rst.Fields("T3O_UTMN")) Then
        .item("T3O_UTMN").Range.text = " " & rst.Fields("T3O_UTMN")
      End If
      If Not IsNull(rst.Fields("T1E_UTME")) Then
        .item("T1E_UTME").Range.text = " " & rst.Fields("T1E_UTME")
      End If
      If Not IsNull(rst.Fields("T1E_UTMN")) Then
        .item("T1E_UTMN").Range.text = " " & rst.Fields("T1E_UTMN")
      End If
      If Not IsNull(rst.Fields("T2E_UTME")) Then
        .item("T2E_UTME").Range.text = " " & rst.Fields("T2E_UTME")
      End If
      If Not IsNull(rst.Fields("T2E_UTMN")) Then
        .item("T2E_UTMN").Range.text = " " & rst.Fields("T2E_UTMN")
      End If
      If Not IsNull(rst.Fields("T3E_UTME")) Then
        .item("T3E_UTME").Range.text = " " & rst.Fields("T3E_UTME")
      End If
      If Not IsNull(rst.Fields("T3E_UTMN")) Then
        .item("T3E_UTMN").Range.text = " " & rst.Fields("T3E_UTMN")
      End If
      If Not IsNull(rst.Fields("T1O_Rebar")) Then
        .item("T1O_Rebar").Range.text = " " & rst.Fields("T1O_Rebar")
      End If
      If Not IsNull(rst.Fields("T1E_Rebar")) Then
        .item("T1E_Rebar").Range.text = " " & rst.Fields("T1E_Rebar")
      End If
      If Not IsNull(rst.Fields("T2O_Rebar")) Then
        .item("T2O_Rebar").Range.text = " " & rst.Fields("T2O_Rebar")
      End If
      If Not IsNull(rst.Fields("T2E_Rebar")) Then
        .item("T2E_Rebar").Range.text = " " & rst.Fields("T2E_Rebar")
      End If
      If Not IsNull(rst.Fields("T3O_Rebar")) Then
        .item("T3O_Rebar").Range.text = " " & rst.Fields("T3O_Rebar")
      End If
      If Not IsNull(rst.Fields("T3E_Rebar")) Then
        .item("T3E_Rebar").Range.text = " " & rst.Fields("T3E_Rebar")
      End If
      If Not IsNull(rst.Fields("T1_Elevation")) Then
        .item("T1_Elevation").Range.text = " " & rst.Fields("T1_Elevation")
      End If
      If Not IsNull(rst.Fields("T2_Elevation")) Then
        .item("T2_Elevation").Range.text = " " & rst.Fields("T2_Elevation")
      End If
      If Not IsNull(rst.Fields("T3_Elevation")) Then
        .item("T3_Elevation").Range.text = " " & rst.Fields("T3_Elevation")
      End If
      If Not IsNull(rst.Fields("Plot_Directions")) Then
        .item("Plot_Directions").Range.text = " " & rst.Fields("Plot_Directions")
      End If
      If rst.Fields("Vegetation_Type") <> "grassland/shrubland" Then
        If Not IsNull(rst.Fields("SlopeA")) Then
          .item("SlopeA").Range.text = " " & rst.Fields("SlopeA")
        End If
        If Not IsNull(rst.Fields("SlopeAUD")) Then
          .item("SlopeAUD").Range.text = " " & rst.Fields("SlopeAUD")
        End If
        If Not IsNull(rst.Fields("SlopeB")) Then
          .item("SlopeB").Range.text = " " & rst.Fields("SlopeB")
        End If
        If Not IsNull(rst.Fields("SlopeBUD")) Then
          .item("SlopeBUD").Range.text = " " & rst.Fields("SlopeBUD")
        End If
        If Not IsNull(rst.Fields("SlopeC")) Then
          .item("SlopeC").Range.text = " " & rst.Fields("SlopeC")
        End If
        If Not IsNull(rst.Fields("SlopeCUD")) Then
          .item("SlopeCUD").Range.text = " " & rst.Fields("SlopeCUD")
        End If
        If Not IsNull(rst.Fields("SlopeD")) Then
          .item("SlopeD").Range.text = " " & rst.Fields("SlopeD")
        End If
        If Not IsNull(rst.Fields("SlopeDUD")) Then
          .item("SlopeDUD").Range.text = " " & rst.Fields("SlopeDUD")
        End If
        strSQL = "Select * from tbl_Monument WHERE Location_ID = '" & rst.Fields("Location_ID") & "'"
        ' Get the monument tree records.
        Set Mrst = New ADODB.Recordset
        Mrst.Open strSQL, CurrentProject.Connection
        If Not Mrst.EOF Then
          Mrst.MoveFirst
          Do Until Mrst.EOF
            intIndex = 0   ' Initialize index
            Do Until intIndex > 11
              If MArray(intIndex) = Mrst.Fields("Monument_Code") Then
                Exit Do
              End If
              intIndex = intIndex + 1
            Loop
            If intIndex > 11 Then
              Exit Do  ' Code not found - skip it
            End If
            intIndex = intIndex + 1  ' set index to correct subscript for bookmarks
            If Not IsNull(Mrst.Fields("Tag_No")) Then
              strFieldName = "Tag_No" & intIndex  ' set bookmark name
              .item(strFieldName).Range.text = " " & Mrst.Fields("Tag_No")
            End If
            If Not IsNull(Mrst.Fields("Species")) Then
              strFieldName = "Species" & intIndex  ' set bookmark name
              .item(strFieldName).Range.text = " " & DLookup("[LU_Code]", "tlu_NCPN_Plants", "[Master_Plant_Code] = '" & Mrst.Fields("Species") & "'")
            End If
            If Not IsNull(Mrst.Fields("DBH")) Then
              strFieldName = "DBH" & intIndex  ' set bookmark name
              .item(strFieldName).Range.text = " " & Mrst.Fields("DBH")
            End If
            If Not IsNull(Mrst.Fields("Bearing")) Then
              strFieldName = "Bearing" & intIndex  ' set bookmark name
              .item(strFieldName).Range.text = " " & Mrst.Fields("Bearing")
            End If
            If Not IsNull(Mrst.Fields("Rebar_Distance")) Then
              strFieldName = "Rebar_Distance" & intIndex  ' set bookmark name
              .item(strFieldName).Range.text = " " & Mrst.Fields("Rebar_Distance")
            End If
            Mrst.MoveNext
          Loop
        End If
      End If  ' End if for forest/woodland compare
      
      ' Check for revisit comments and print if necessary
      If Not IsNull(Me!Park_Code) And Not IsNull(Me!Plot_ID) Then
        strSQL = "SELECT * FROM qry_Visit_Comments WHERE [Unit_Code] = '" & Me!Park_Code & "' AND [Plot_ID] = " & Me!Plot_ID & " ORDER BY [Start_Date] DESC"
        ' Get the database record.
        Set RevisitComments = New ADODB.Recordset
        RevisitComments.Open strSQL, CurrentProject.Connection
        If Not RevisitComments.EOF Then
          RevisitComments.MoveFirst
          If Not IsNull(RevisitComments!Comments) Then
            strPrompt = "There are revisit comments dated " & RevisitComments.Fields("Start_Date") & " for this plot." & Chr(13) & Chr(10) & "Do you want to print them?"
            intReply = MsgBox(strPrompt, vbYesNo, "Want revisit comments?")
            If intReply = vbYes Then
              .item("Revisit_Comments").Range.text = " " & RevisitComments.Fields("Comments")
            End If
          End If
        End If  ' End if for comments eof test
      End If  ' End if for null unit or plot test
      End With

    rst.Close    ' Close for main query

Exit_Button_Print:
    DoCmd.Hourglass False
    Exit Sub

Err_Button_Print:
    MsgBox Err.Description
    Resume Exit_Button_Print
   
    
End Sub
