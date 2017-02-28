' =================================
' MODULE:       basUtilities
' Description:  Standard module containing commonly-used utility functions
' Source/date:  John R. Boetsch, May 17, 2006
' Revisions:    <name, date, desc - add lines as you go>

Option Compare Database
Option Explicit

' =================================
' FUNCTION:     fxnReplaceString
' Description:  Replaces a substring in a string with another
' Parameters:   strTextIn - string to work on
'               strFind - string to find
'               strReplace - string to replace with
'               fCaseSensitive - True for case sensitive search, False otherwise
' Returns:      modified string
' Throws:       none
' References:   none
' Source/date:  Simon Kingston, date unknown
' Revisions:    John R. Boetsch, May 17, 2006 - error trapping, documentation
' =================================

Function fxnReplaceString(strTextIn As String, strFind As String, _
    strReplace As String, fCaseSensitive As Boolean) As String

    On Error GoTo Err_Handler

    Dim strTemp As String
    Dim intPos As Integer
    Dim intCaseSensitive As Integer

    ' Convert the case-sensitive boolean to the comparison constant (1=binary, 2=textual)
    intCaseSensitive = fCaseSensitive + 1

    strTemp = strTextIn
    intPos = InStr(1, strTemp, strFind, intCaseSensitive)

    Do While intPos > 0
        strTemp = Left$(strTemp, intPos - 1) & strReplace & Mid$(strTemp, intPos + Len(strFind))
        intPos = InStr(intPos + Len(strReplace), strTemp, strFind, intCaseSensitive)
    Loop

    fxnReplaceString = strTemp

Exit_Procedure:
    Exit Function

Err_Handler:
    Select Case Err.Number
        Case Else
            MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
                "Error encountered (fxnReplaceString)"
            Resume Exit_Procedure
    End Select

End Function

' =================================
' FUNCTION:     fxnChangeDelimiter
' Description:  Replaces delimiters in an input string; default is to change double-quotes
'               to single quotes
' Parameters:   strInputText - string to work on
'               strCurrDelimiter - current delimiter in the string (default: double-quote)
'               strNewDelimiter - desired replacement delimiter (default: single-quote)
' Returns:      modified string
' Throws:       none
' References:   fxnReplaceString
' Source/date:  John R. Boetsch, May 17, 2006
' Revisions:    <name, date, desc - add lines as you go>
' =================================

Public Function fxnChangeDelimiter(strInputText As String, _
    Optional strCurrDelimiter As String = """", _
    Optional strNewDelimiter As String = "'") As String

    On Error GoTo Err_Handler

    Dim strTemp As String
    
    ' Call the replace string function, specifying the delimiter and no case-sensitive search
    strTemp = fxnReplaceString(strInputText, strCurrDelimiter, strNewDelimiter, False)
    fxnChangeDelimiter = strTemp

Exit_Procedure:
    Exit Function

Err_Handler:
    Select Case Err.Number
        Case Else
            MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
                "Error encountered (fxnChangeDelimiter)"
            Resume Exit_Procedure
    End Select

End Function

' =================================
' FUNCTION:     fxnTrimSpaces
' Description:  Removes leading and trailing space characters from a string
' Parameters:   strInputText - string to work on
' Returns:      modified string
' Throws:       none
' References:   none
' Source/date:  John R. Boetsch, May 25, 2006
' Revisions:    <name, date, desc - add lines as you go>
' =================================

Public Function fxnTrimSpaces(strInputText As String) As String
    On Error GoTo Err_Handler

    Dim strTemp As String

    strTemp = strInputText

    ' First trim leading spaces
    Do While Left(strTemp, 1) = " "
        strTemp = Right(strTemp, Len(strTemp) - 1)
    Loop
    ' Then trim trailing spaces
    Do While Right(strTemp, 1) = " "
        strTemp = Left(strTemp, Len(strTemp) - 1)
    Loop

    fxnTrimSpaces = strTemp

Exit_Procedure:
    Exit Function

Err_Handler:
    Select Case Err.Number
        Case Else
            MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
                "Error encountered (fxnTrimSpaces)"
            Resume Exit_Procedure
    End Select

End Function

Public Function IsNetwork(varUnitCode As Variant) As Boolean
Select Case varUnitCode
    Case "ARCN", "CAKN", "CHDN", "CUPN", "ERMN", "GLKN", "GRYN", "GULN", "HTLN", "KLMN", "MEDN", "MIDN", "MOJN", "NCBN", "NCCN", "NCPN", "NCRN", "NETN", "NGPN", "PACN", "ROMN", "SCPN", "SEAN", "SECN", "SFAN", "SFCN", "SIEN", "SODN", "SOPN", "SWAN", "UCBN"
        IsNetwork = True
End Select
End Function