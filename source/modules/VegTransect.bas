Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

' =================================
' CLASS:        VegTransect
' Level:        Framework class
' Version:      1.01
'
' Description:  VegTransect object related properties, events, functions & procedures
'
' Source/date:  Bonnie Campbell, 10/28/2015
' References:   -
' Revisions:    BLC - 10/28/2015 - 1.00 - initial version
'               BLC - 8/8/2016   - 1.01 - SaveToDb() added update parameter to identify if
'                                        this is an update vs. an insert
'               BLC - 7/5/2017   - 1.02 - AddQuadrats() & AddSurfaceMicrohabitats() to
'                                         initialize records tied to new transect
' =================================

'---------------------
' Declarations
'---------------------
Private m_ID As Long
Private m_LocationID As Long
Private m_EventID As Long
Private m_TransectQuadratID As String
Private m_TransectNumber As Integer
Private m_SampleDate As Date

Private m_Park As String
Private m_ObserverID As Integer
Private m_RecorderID As Integer
Private m_Observer As String
Private m_Recorder As String

Private m_HasQuadrats As Boolean
Private m_TransectQuadrats As Variant
Private m_NumQuadrats As Integer

'---------------------
' Events
'---------------------
Public Event InvalidTransectNumber(Value As Integer)
Public Event InvalidTransectQuadratID(Value As String)

'---------------------
' Properties
'---------------------
Public Property Let ID(Value As Long)
    m_ID = Value
End Property

Public Property Get ID() As Long
    ID = m_ID
End Property

Public Property Let LocationID(Value As Long)
    m_LocationID = Value
    'set the appropriate park value
'    Me.Park = GetParkCode(Value)
End Property

Public Property Get LocationID() As Long
    LocationID = m_LocationID
End Property

Public Property Let EventID(Value As Long)
    m_EventID = Value
End Property

Public Property Get EventID() As Long
    EventID = m_EventID
End Property

Public Property Let TransectQuadratID(Value As String)
    m_TransectQuadratID = Value
    
    'set the tempvar also
    SetTempVar "TransectQuadratID", Value
    
    'populate related properties
    GetTransectQuadrats
End Property

Public Property Get TransectQuadratID() As String
    TransectQuadratID = m_TransectQuadratID
End Property


Public Property Let TransectNumber(Value As Integer)
    If IsNull(Me.Park) Then
        MsgBox "Park must be set before setting transect number.", vbCritical, "Missing Park"
        
    End If
    'validate park (BLCA & CANY only)
'    Select Case Me.Park
'        Case "BLCA", "CANY"
'            'check value
'            'validate transect #
'            Dim aryTransectNum() As String
'            aryTransectNum = Split(TRANSECT_NUMBERS, ",")
'            If IsInArray(CStr(value), aryTransectNum) Then
'                m_TransectNumber = value
'            Else
'                RaiseEvent InvalidTransectNumber(value)
'            End If
'        Case "DINO"
'            'invalid
'            RaiseEvent InvalidTransectNumber(value)
'        Case Else
'            'invalid
'            RaiseEvent InvalidTransectNumber(value)
'    End Select
End Property

Public Property Get TransectNumber() As Integer
    TransectNumber = m_TransectNumber
End Property

Public Property Let SampleDate(Value As Date)
    m_SampleDate = Value
End Property

Public Property Get SampleDate() As Date
    SampleDate = m_SampleDate
End Property

Public Property Let Park(Value As String)
    m_Park = Value
End Property

Public Property Get Park() As String
    Park = m_Park
End Property

Public Property Let ObserverID(Value As Integer)
    m_ObserverID = Value
End Property

Public Property Get ObserverID() As Integer
    ObserverID = m_ObserverID
End Property

Public Property Let Observer(Value As String)
    m_Observer = Value
End Property

Public Property Get Observer() As String
    Observer = m_Observer
End Property

Public Property Let RecorderID(Value As Integer)
    m_RecorderID = Value
End Property

Public Property Get RecorderID() As Integer
    RecorderID = m_RecorderID
End Property

Public Property Let Recorder(Value As String)
    m_Recorder = Value
End Property

Public Property Get Recorder() As String
    Recorder = m_Recorder
End Property

Public Property Let HasQuadrats(Value As Boolean)
    m_HasQuadrats = Value
End Property

Public Property Get HasQuadrats() As Boolean
    HasQuadrats = m_HasQuadrats
End Property

Public Property Let TransectQuadrats(Value As Variant)
    m_TransectQuadrats = Value
End Property

Public Property Get TransectQuadrats() As Variant
    TransectQuadrats = m_TransectQuadrats
End Property

Public Property Let NumQuadrats(Value As Integer)
    m_NumQuadrats = Value
End Property

Public Property Get NumQuadrats() As Integer
    NumQuadrats = m_NumQuadrats
End Property

'---------------------
' Methods
'---------------------

'======== Standard Methods ===========

' ---------------------------------
' SUB:          Class_Initialize
' Description:  Initialize the class
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  -
' Adapted:      Bonnie Campbell, April 4, 2016 - for NCPN tools
' Revisions:
'   BLC - 4/4/2016 - initial version
' ---------------------------------
Private Sub Class_Initialize()
On Error GoTo Err_Handler

Exit_Handler:
    Exit Sub

Err_Handler:
    Select Case Err.Number
        Case Else
            MsgBox "Error #" & Err.Description, vbCritical, _
                "Error encounter (#" & Err.Number & " - Class_Initialize[cls_VegPlot])"
    End Select
    Resume Exit_Handler
End Sub

'---------------------------------------------------------------------------------------
' SUB:          Class_Terminate
' Description:  -
' Parameters:   -
' Returns:      -
' Throws:       -
' References:   -
' Source/Date:  Bonnie Campbell
' Adapted:      Bonnie Campbell, 4/4/2016 - for NCPN tools
' Revisions:
'   BLC, 4/4/2016 - initial version
'---------------------------------------------------------------------------------------
Private Sub Class_Terminate()
On Error GoTo Err_Handler

    'Set m_ID = 0

Exit_Handler:
    Exit Sub

Err_Handler:
    Select Case Err.Number
        Case Else
            MsgBox "Error #" & Err.Description, vbCritical, _
                "Error encounter (#" & Err.Number & " - Class_Terminate[cls_VegPlot])"
    End Select
    Resume Exit_Handler
End Sub

'======== Custom Methods ===========
'---------------------------------------------------------------------------------------
' SUB:          SaveToDb
' Description:  -
' Parameters:   -
' Returns:      -
' Throws:       -
' References:
'   Fionnuala, February 2, 2009
'   David W. Fenton, October 27, 2009
'   http://stackoverflow.com/questions/595132/how-to-get-id-of-newly-inserted-record-using-excel-vba
' Source/Date:  Bonnie Campbell
' Adapted:      Bonnie Campbell, 4/4/2016 - for NCPN tools
' Revisions:
'   BLC, 4/4/2016 - initial version
'   BLC, 8/8/2016 - added update parameter to identify if this is an update vs. an insert
'   BLC, 9/8/2016 - code cleanup
'---------------------------------------------------------------------------------------
Public Sub SaveToDb(Optional isUpdate As Boolean = False)
On Error GoTo Err_Handler
    
    Dim Template As String
    
    Template = "i_vegtransect"
    
    Dim params(0 To 6) As Variant

    With Me
        params(0) = "VegTransect"
        params(1) = .LocationID
        params(2) = .EventID
        params(3) = .TransectNumber
        params(4) = .SampleDate
        
        If isUpdate Then
            Template = "u_vegtransect"
            params(5) = .ID
        End If
        
        .ID = SetRecord(Template, params)
    End With
    
    'SetObserverRecorder Me, "VegTransect"

Exit_Handler:
    Exit Sub

Err_Handler:
    Select Case Err.Number
        Case Else
            MsgBox "Error #" & Err.Description, vbCritical, _
                "Error encounter (#" & Err.Number & " - SaveToDb[cls_VegPlot])"
    End Select
    Resume Exit_Handler
End Sub

'---------------------------------------------------------------------------------------
' Function:     GetTransectQuadrats
' Description:  Fetch the quadrats for the transect (if any)
' Parameters:   -
' Returns:      Array of quadrats
' Throws:       -
' References:   -
' Source/Date:  Bonnie Campbell
' Adapted:      Bonnie Campbell, 7/10/2017 - for NCPN tools
' Revisions:
'   BLC, 7/10/2017 - initial version
'---------------------------------------------------------------------------------------
Public Function GetTransectQuadrats() As Variant
On Error GoTo Err_Handler
    
    Dim rs As DAO.Recordset
    
    'retrieve array of all quadrats associated w/ this transect
    Set rs = GetRecords("s_transect_quadrat_IDs")
        
    rs.MoveLast
    rs.MoveFirst
    
    With Me
        'defaults
        .HasQuadrats = False
        .NumQuadrats = 0
        
        'set had quadrats
        If rs.RecordCount > 0 Then
            .HasQuadrats = True
            .NumQuadrats = rs.RecordCount
        End If
    
        'return the 2-dimensional array (1-columns, 2-rows)
        .TransectQuadrats = rs.GetRows(rs.RecordCount)
    
        GetTransectQuadrats = .TransectQuadrats
    
    End With
    
Exit_Handler:
    rs.Close
    Set rs = Nothing
    Exit Function

Err_Handler:
    Select Case Err.Number
        Case Else
            MsgBox "Error #" & Err.Description, vbCritical, _
                "Error encounter (#" & Err.Number & " - GetTransectQuadrats[cls_VegPlot])"
    End Select
    Resume Exit_Handler
End Function

'---------------------------------------------------------------------------------------
' SUB:          AddQuadrats
' Description:  Adds quadrats to a new transect which has no quadrat records
'               QUADRATS_PER_TRANSECT (see mod_App_Settings) is the current # of quadrats
'               existing along a transect

' Parameters:   QuadratNum - number of quadrat to add (integer, optional)
'                   0 - add all quadrats 1 - QUADRATS_PER_TRANSECT value (currently 3)
'                   1, 2, or 3 - add quadrat 1,2, or 3
' Returns:      -
' Throws:       -
' References:   -
' Source/Date:  Bonnie Campbell
' Adapted:      Bonnie Campbell, 7/5/2017 - for NCPN tools
' Revisions:
'   BLC, 7/5/2017 - initial version
'---------------------------------------------------------------------------------------
Public Sub AddQuadrats(Optional QuadratNum As Integer = 0)
On Error GoTo Err_Handler
    
    Dim Template As String
    Dim i As Integer
    
    Template = "i_new_transect_quadrat"
    
    Dim params(0 To 3) As Variant

    'if QuadratNum is set
    If QuadratNum <> 0 Then
        With Me
            params(0) = "Transect"
            params(1) = .TransectQuadratID
            params(2) = i                   'quadrat number
            
            .ID = SetRecord(Template, params)
        End With
    
        'exit
        GoTo Exit_Handler
    End If

    'if QuadratNum = 0 then assume add all quadrats
    'iterate once per quadrat
    For i = 1 To QUADRATS_PER_TRANSECT

        With Me
            params(0) = "Transect"
            params(1) = .TransectQuadratID
            params(2) = i                   'quadrat number
            
            .ID = SetRecord(Template, params)
        End With
        
        'SetObserverRecorder Me, "VegTransect"
    
    Next
    
Exit_Handler:
    Exit Sub

Err_Handler:
    Select Case Err.Number
        Case Else
            MsgBox "Error #" & Err.Description, vbCritical, _
                "Error encounter (#" & Err.Number & " - AddQuadrats[cls_VegPlot])"
    End Select
    Resume Exit_Handler
End Sub

'---------------------------------------------------------------------------------------
' SUB:          AddSurfaceMicrohabitats
' Description:  Adds quadrats to a new transect which has no quadrat records
'               QUADRATS_PER_TRANSECT (see mod_App_Settings) is the current # of quadrats
'               existing along a transect

' Parameters:   QuadratNum - number of quadrat to add (integer, optional)
'                   0 - add all quadrats 1 - QUADRATS_PER_TRANSECT value (currently 3)
'                   1, 2, or 3 - add quadrat 1,2, or 3
' Returns:      -
' Throws:       -
' References:   -
' Source/Date:  Bonnie Campbell
' Adapted:      Bonnie Campbell, 7/5/2017 - for NCPN tools
' Revisions:
'   BLC, 7/5/2017 - initial version
'---------------------------------------------------------------------------------------
Public Sub AddSurfaceMicrohabitats(Optional SfcMicrohabitat As Integer = 0)
On Error GoTo Err_Handler
    
    Dim Template As String
    Dim arySurfaces As Variant
    Dim aryQuadrats As Variant
    Dim rs As DAO.Recordset
    Dim sfc_id As Variant
    Dim QuadratID As Variant
    
    'retrieve array of all surface IDs
    Set rs = GetRecords("s_surface_IDs")
    rs.MoveLast
    rs.MoveFirst
    arySurfaces = rs.GetRows(rs.RecordCount)
    
    'retrieve array of all quadrats associated w/ this transect
    Set rs = GetRecords("s_transect_quadrat_IDs")
    rs.MoveLast
    rs.MoveFirst
    aryQuadrats = rs.GetRows(rs.RecordCount)
    
    Template = "i_new_transect_quadrat_sfccover"
    
    Dim params(0 To 3) As Variant

    'iterate once per quadrat
    For Each QuadratID In aryQuadrats
    
        'iterate once per surface
        For Each sfc_id In arySurfaces
    
            With Me
                params(0) = "Transect"
                params(1) = QuadratID       'quadrat ID
                params(2) = sfc_id          'surface microhabitat ID
                
                .ID = SetRecord(Template, params)
            End With
        
            'SetObserverRecorder Me, "VegTransect"
                
        Next
        
    Next
    
Exit_Handler:
    Exit Sub

Err_Handler:
    Select Case Err.Number
        Case Else
            MsgBox "Error #" & Err.Description, vbCritical, _
                "Error encounter (#" & Err.Number & " - AddSurfaceMicrohabitats[cls_VegPlot])"
    End Select
    Resume Exit_Handler
End Sub