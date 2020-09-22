Attribute VB_Name = "Function"
Option Explicit
Public cn As ADODB.Connection

Public Function OpenDatabase() As Boolean
On Error GoTo checkErr
    Dim cmd As String

    Set cn = New ADODB.Connection
    
    cn.Provider = "Microsoft.Jet.OLEDB.4.0"
    cn.ConnectionString = "Data Source=" & App.Path & "\" & gstrDatabasePath
    cn.Properties("Jet OLEDB:Database Password") = gstrPassword
    cn.Open
    
    OpenDatabase = True
    Exit Function
checkErr:
    OpenDatabase = False
    WriteText "Error"
End Function

Public Function CloseDatabase()
On Error GoTo checkErr
    If cn Is Nothing Then
    Else
        If cn.State = adStateOpen Then
            cn.Close
        End If
        Set cn = Nothing
    End If
    Exit Function
checkErr:
    WriteText "Error"
End Function

Public Function CheckString(strInput As String) As String
    CheckString = Replace(strInput, "'", "''")
End Function

'Update Birthday for specified DOB Day and Month, Year is optional
'Public Sub UpdateBirthday()
'Dim strSQL As String
''On Error GoTo WriteErrorLog
'    strSQL = "UPDATE Contact SET Birthday = DateSerial(Year(Date()),DOBMonth,DOBDay)" & _
'            " WHERE DOBMonth IS NOT NULL AND DOBDay IS NOT NULL AND NOT (DOBMonth=2 AND DOBDay=29);"
'    cn.Execute strSQL
'    Exit Sub
'WriteErrorLog:
'    WriteText "Error"
'End Sub

'Update Birthday for specified DOB Day and Month, Year is optional
Public Sub UpdateBirthday()
Dim strSQL As String
On Error GoTo checkErr
    'If Current year is not leap year
    If CheckLeapYear(Year(Date)) = False Then
        'Birthday not fall in 29 Feb
        strSQL = "UPDATE Contact SET Birthday = DateSerial(Year(Date()), DOBMonth, DOBDay)" & _
                " WHERE DOBMonth IS NOT NULL AND DOBDay IS NOT NULL AND NOT (DOBMonth=2 AND DOBDay=29);"
        cn.Execute strSQL
        'Birthday fall on 29th Feb
        strSQL = "UPDATE Contact SET Birthday = #28 Feb " & Year(Date) & _
        "# WHERE DOBMonth = 2 AND DOBDay = 29;"
        cn.Execute strSQL
        'Birthday not complete (Has DOBMonth only)
        strSQL = "UPDATE Contact SET Birthday = DateSerial(Year(Date()), DOBMonth, 1)" & _
                " WHERE DOBMonth IS NOT NULL AND DOBDay IS NULL;"
        cn.Execute strSQL
    Else 'Current year is a leap year
        'Birthday fall on 29th Feb
        strSQL = "UPDATE Contact SET Birthday = DateSerial(Year(Date()), DOBMonth, DOBDay)" & _
                " WHERE DOBMonth = 2 AND DOBDay = 29;"
    End If
    Exit Sub
checkErr:
    WriteText "Error"
End Sub

Public Function ComputeAge(pDay As Integer, pMonth As Integer, pYear As Integer) As Integer
Dim blnYear As Boolean
Dim blnMonth As Boolean
Dim blnDay As Boolean

If pYear > 1581 And pYear < 10000 Then blnYear = True
If pMonth > 0 And pMonth < 13 Then blnMonth = True
If pDay > 0 And pDay < 32 Then blnDay = True
    
If blnYear = True Then
    If blnMonth = True Then
        If blnDay = True Then
            ComputeAge = DateDiff("yyyy", DateSerial(pYear, pMonth, pDay), Date)
        Else
            ComputeAge = DateDiff("yyyy", DateSerial(pYear, pMonth, 1), Date)
        End If
    Else
        ComputeAge = DateDiff("yyyy", DateSerial(pYear, 1, 1), Date)
    End If
Else
    ComputeAge = 0
End If
End Function

Public Function ExecuteSelectSQL(strSQL As String, Optional aiCursorType As Integer = adOpenDynamic) As ADODB.Recordset
Dim rstTemp As ADODB.Recordset
On Error GoTo checkErr
    Set rstTemp = New ADODB.Recordset
  
    rstTemp.Open strSQL, cn, adOpenForwardOnly, adLockOptimistic
    Set ExecuteSelectSQL = rstTemp
    Exit Function
checkErr:
    WriteText "Error"
End Function

Public Function GenWord()
    Dim intArray(7) As Integer
    Dim l As Integer
    
    intArray(0) = 51
    intArray(1) = 50
    intArray(2) = 51
    intArray(3) = 48
    intArray(4) = 50
    intArray(5) = 49
    intArray(6) = 56
    
    gstrPassword = ""
    
    For l = 0 To 6
        gstrPassword = gstrPassword & Chr(intArray(l))
    Next
End Function

Public Sub WriteText(FileName As String, Optional sNote As String)
    Open App.Path & "\" & FileName & ".txt" For Append As #1
    If sNote = "" Then
        Write #1, Now, Error
    Else
        Write #1, Now, Error & " @ " & sNote
    End If
    Close #1
End Sub

Public Sub ReadText(ByVal FileName As String, ByVal LineNo As Integer, ByRef sOutput As String)
    On Error GoTo newfile
    Dim I As Integer
    Open App.Path & "\" & FileName & ".txt" For Input As #2
        If LineNo < 0 Then
            Do Until EOF(2) = True
                Input #2, sOutput
            Loop
        ElseIf LineNo > 0 Then
            For I = 0 To LineNo
                If Not EOF(2) Then
                    Input #2, sOutput
                Else
                    Exit For
                End If
            Next
        Else
            Input #2, sOutput
        End If
    Close
    Exit Sub
newfile:
    WriteText "Error", "ReadText(" & FileName & ".txt)"
End Sub

Public Function FileExists(strPath As String) As Boolean
Dim lngRetVal As Long

On Error Resume Next
    lngRetVal = Len(Dir$(strPath))
    If Err Or lngRetVal = 0 Then
        FileExists = False
    Else
        FileExists = True
    End If
End Function

Public Function ReadSettings()
Dim gstrTemp As String
On Error GoTo checkErr
    GenWord
    ReadText "Data", 1, gstrDatabasePath '"MyDatabase.mdb"
    ReadText "Data", 3, gstrTemp
    ReadText "Data", 5, gstrFontName
    If gstrPassword = gstrTemp Then
        gblnWithPassword = True
    Else
        gblnWithPassword = False
    End If
    Exit Function
checkErr:
    WriteText "Error", "ReadSettings(Data.txt)"
End Function

Public Function CheckLeapYear(pintYear As Integer) As Boolean
If pintYear < 1582 Or pintYear > 9999 Then Exit Function
If pintYear Mod 400 = 0 Then
    CheckLeapYear = True
Else
    If pintYear Mod 100 = 0 Then
        CheckLeapYear = False
    Else
        If pintYear Mod 4 = 0 Then
            CheckLeapYear = True
        End If
    End If
End If
End Function

Public Function CheckValidDay(pintDay As Integer, pintMonth As Integer, pintYear As Integer) As Boolean
    If pintDay < 1 Or pintDay > 31 Then Exit Function
    If pintMonth < 1 Or pintMonth > 12 Then Exit Function
    If pintYear < 1582 Or pintYear > 9999 Then Exit Function
        
    If pintMonth = 2 Then
        If CheckLeapYear(pintYear) = True Then
            If pintDay < 30 Then CheckValidDay = True
        Else
            If pintDay < 29 Then CheckValidDay = True
        End If
    ElseIf (pintMonth = 4 Or pintMonth = 6 Or pintMonth = 9 Or pintMonth = 11) Then
        If pintDay < 31 Then CheckValidDay = True
    Else
        CheckValidDay = True
    End If
End Function

Public Function CorrectDay(pintDay As Integer, pintMonth As Integer, pintYear As Integer) As Integer
    CorrectDay = pintDay
    
    If pintYear < 1582 Or pintYear > 9999 Then Exit Function
    If pintMonth < 1 Or pintMonth > 12 Then Exit Function
    If pintDay < 1 Or pintDay > 31 Then Exit Function
    
    If CheckValidDay(pintDay, pintMonth, pintYear) = False Then
        CorrectDay = pintDay - 1
        CorrectDay = CorrectDay(CorrectDay, pintMonth, pintYear)
    End If
End Function
