Attribute VB_Name = "modDatabase"
Public Sub GetRates()
    'Sub to load the telephone rates from the database
    
    On Error GoTo ErrHandler
    
    Set conn = New ADODB.Connection
    conn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=CostTracker.mdb;Jet OLEDB:Engine Type=5;"
    Set rs = New ADODB.Recordset
    rs.Open "SELECT * FROM rates", conn
      
    Do Until rs.EOF
        modMain.Rates(rs("day"), rs("hour")) = rs("rate")
        rs.MoveNext
    Loop
    
    rs.Close
    conn.Close
    Set rs = Nothing
    Set conn = Nothing
    
    Exit Sub
ErrHandler:
    MsgBox modErrHandling.ErrHandler(Err), vbCritical, "CostTracker"
    Resume Next
End Sub
Public Sub SaveRates()
    'Sub to write new rates to the database
    
    On Error GoTo ErrHandler
    
    Set conn = New ADODB.Connection
    conn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=CostTracker.mdb;Jet OLEDB:Engine Type=5;"
    Set rs = New ADODB.Recordset
    rs.Open "SELECT * FROM rates", conn, 3, 3
      
    Do Until rs.EOF
        rs.Delete
        rs.MoveNext
    Loop
    
    For i% = 1 To 8
        For i2% = 0 To 23
            rs.AddNew
            rs("day") = i%
            rs("hour") = i2%
            rs("rate") = Rates(i%, i2%)
            rs.Update
        Next i2%
    Next i%
    
    rs.Close
    conn.Close
    Set rs = Nothing
    Set conn = Nothing
    
    Exit Sub
ErrHandler:
    MsgBox modErrHandling.ErrHandler(Err), vbCritical, "CostTracker"
    Resume Next
End Sub
Public Sub AddSession(UserName As String, StartTime As String, StartDate As String, Claimable As Boolean, Costs As Currency)
    'Sub to add the session to the database
    
    On Error GoTo ErrHandler
    
    Set conn = New ADODB.Connection
    conn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=CostTracker.mdb;Jet OLEDB:Engine Type=5;"
    Set rs = New ADODB.Recordset
    rs.Open "SELECT * FROM sessions", conn, 3, 3
    
    'This sub is called every minute so I need to delete any sessions that have been stored earlier, before we can add the new one. If I wouldn't the database would get filled with a lot of one-minute sessions
    Do Until rs.EOF
        If rs("startdate") = StartDate And rs("enddate") = Date And rs("starttime") = StartTime Then
            rs("user") = UserName
            rs("endtime") = Time
            rs("claimable") = Claimable
            rs("costs") = CInt(Costs * 100) / 100
            rs.Update
            
            rs.Close
            conn.Close
            Set rs = Nothing
            Set conn = Nothing
            
            Exit Sub
        End If
        rs.MoveNext
    Loop
    
    rs.AddNew
    rs("user") = UserName
    rs("startdate") = StartDate
    rs("enddate") = Date
    rs("starttime") = StartTime
    rs("endtime") = Time
    rs("claimable") = Claimable
    rs("costs") = CInt(Costs * 100) / 100
    rs.Update
    
    rs.Close
    conn.Close
    Set rs = Nothing
    Set conn = Nothing
    
    Exit Sub
ErrHandler:
    MsgBox modErrHandling.ErrHandler(Err), vbCritical, "CostTracker"
    Resume Next
End Sub
Public Function GetMonthCosts(UserName As String, SpecifiedMonth As Integer, Claimable As Boolean) As Currency
    'Sub to check how much a user has already spent this month. It can check all the costs or just the claimable sessions, depending on the value of the Claimable boolean
    
    On Error GoTo ErrHandler
    
    Set conn = New ADODB.Connection
    conn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=CostTracker.mdb;Jet OLEDB:Engine Type=5;"
    Set rs = New ADODB.Recordset
    rs.Open "SELECT * FROM sessions", conn
    
    Do Until rs.EOF
        If rs("user") = UserName And Month(rs("startdate")) = SpecifiedMonth Then
            If (Claimable = rs("claimable")) Or Claimable = False Then GetMonthCosts = GetMonthCosts + rs("costs")
        End If
        rs.MoveNext
    Loop
    
    rs.Close
    conn.Close
    Set rs = Nothing
    Set conn = Nothing
    
    Exit Function
ErrHandler:
    MsgBox modErrHandling.ErrHandler(Err), vbCritical, "CostTracker"
    Resume Next

End Function
Public Function DeleteMonthCosts(UserName As String, SpecifiedMonth As Integer)
    'Sub to remove the sessions of a specific user and month
    
    On Error GoTo ErrHandler
    
    Set conn = New ADODB.Connection
    conn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=CostTracker.mdb;Jet OLEDB:Engine Type=5;"
    Set rs = New ADODB.Recordset
    rs.Open "SELECT * FROM sessions", conn, 3, 3
    
    Do Until rs.EOF
        If rs("user") = UserName And Month(rs("startdate")) = SpecifiedMonth Then
            rs.Delete
        End If
        rs.MoveNext
    Loop
    
    rs.Close
    conn.Close
    Set rs = Nothing
    Set conn = Nothing
    
    Exit Function
ErrHandler:
    MsgBox modErrHandling.ErrHandler(Err), vbCritical, "CostTracker"
    Resume Next

End Function
Public Sub GetUsers(Administrator As Boolean, DestForm As Form)
    'Sub to get all the usernames and password from the database
    
    On Error GoTo ErrHandler
    
    DestForm.lstUsers.Clear
    DestForm.lstPasswords.Clear
    
    Set conn = New ADODB.Connection
    conn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=CostTracker.mdb;Jet OLEDB:Engine Type=5;"
    Set rs = New ADODB.Recordset
    rs.Open "SELECT * FROM users", conn
     
    Do Until rs.EOF
        If Administrator = False Or (Administrator = True And rs("administrator") = True) Then
            DestForm.lstUsers.AddItem rs("name")
            DestForm.lstUsers.ItemData(DestForm.lstUsers.NewIndex) = rs("administrator")
            'The passwords must first be decrypted
            DestForm.lstPasswords.AddItem modEncryption.Decrypt(CStr(rs("password"))), DestForm.lstUsers.NewIndex
        End If
        rs.MoveNext
    Loop
     
    rs.Close
    conn.Close
    Set rs = Nothing
    Set conn = Nothing
    
    Exit Sub
ErrHandler:
    MsgBox modErrHandling.ErrHandler(Err), vbCritical, "CostTracker"
    Resume Next
End Sub

Public Sub SaveUsers(SourceForm As Form)
    'Sub to add a new user to the database
    
    On Error GoTo ErrHandler
    
    
    Set conn = New ADODB.Connection
    conn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=CostTracker.mdb;Jet OLEDB:Engine Type=5;"
    Set rs = New ADODB.Recordset
    rs.Open "SELECT * FROM users", conn, 3, 3
     
    Do Until rs.EOF
        rs.Delete
        rs.MoveNext
    Loop
    
    For i% = 0 To SourceForm.lstUsers.ListCount - 1
        rs.AddNew
        rs("name") = SourceForm.lstUsers.List(i%)
        'The password has to be encrypted first
        rs("password") = modEncryption.Encrypt(SourceForm.lstPasswords.List(i%))
        rs("administrator") = SourceForm.lstUsers.ItemData(i%)
        rs.Update
    Next i%
    
    rs.Close
    conn.Close
    Set rs = Nothing
    Set conn = Nothing
    
    Exit Sub
ErrHandler:
    MsgBox modErrHandling.ErrHandler(Err), vbCritical, "CostTracker"
    Resume Next

End Sub

Public Sub GetSessionMonths(DestForm As Form)
    'Sub to get all the months of which data is available
    
    On Error GoTo ErrHandler
    
    Dim AlreadyAdded As String
    
    DestForm.lstMonths.Clear
    
    Set conn = New ADODB.Connection
    conn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=CostTracker.mdb;Jet OLEDB:Engine Type=5;"
    Set rs = New ADODB.Recordset
    rs.Open "SELECT * FROM sessions", conn, 3, 3
        
    Do Until rs.EOF
        If InStr(1, AlreadyAdded, Month(rs("startdate")) & "*") = 0 Then
            DestForm.lstMonths.AddItem Month(rs("startdate"))
            AlreadyAdded = AlreadyAdded & Month(rs("startdate")) & "*"
        End If
        rs.MoveNext
    Loop
        
    rs.Close
    conn.Close
    Set rs = Nothing
    Set conn = Nothing
    
    Exit Sub
ErrHandler:
    MsgBox modErrHandling.ErrHandler(Err), vbCritical, "CostTracker"
    Resume Next

End Sub
