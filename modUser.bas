Attribute VB_Name = "modUser"
Public UserName As String

'This function is used when a user has to identify himself.
Public Function ValidateUser(TimeLimit As Integer, Administrator As Boolean)
    
    On Error GoTo ErrHandler
    
    Dim LoadTimer As Single
    
    'Clear the user and password lists
    UserName = ""
    frmChooseUser.lstUsers.Clear
    frmChooseUser.lstPasswords.Clear
    'Load the user and password lists. If Administrator is true, we only look for administrators
    modDatabase.GetUsers Administrator, frmChooseUser
    
    'If there are no users: exit
    If frmChooseUser.lstUsers.ListCount = 0 Then Unload frmChooseUser: Exit Function
    
    'If a user has already identified hisself (when he connected), then check if this user is on the list and if so: exit. A user can also not be on the list, because for some functions only administrators may log in
    If Len(modMain.UserName) > 0 Then
        For i% = 0 To frmChooseUser.lstUsers.ListCount - 1
            If frmChooseUser.lstUsers.List(i%) = modMain.UserName Then Unload frmChooseUser: Exit Function
        Next i%
    End If
    
    'Show the logon window. The logon form takes care of the rest of the procedure and returns the username if somebody has succesfully logged in.
    Load frmChooseUser
    
    'Start the timer
    LoadTimer = Timer
    Do
        DoEvents
        'Wait until the user has identified hisself
        If UserName <> "" Then Exit Do
        'Check if a possible timelimit has exceeded and if so: exit
        If TimeLimit > 0 Then If Timer > LoadTimer + TimeLimit Then UserName = "Unknown"
    Loop
    
    'Hide the logon window
    Unload frmChooseUser
    
    'The user has identified hisself: return the username
    ValidateUser = UserName
    
    Exit Function
ErrHandler:
    MsgBox modErrHandling.ErrHandler(Err), vbCritical, "CostTracker"
    Resume Next
End Function

