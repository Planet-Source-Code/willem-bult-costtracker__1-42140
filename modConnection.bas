Attribute VB_Name = "modConnection"
Private Declare Function InternetGetConnectedState Lib "wininet" (lpdwFlags As Long, ByVal dwReserved As Long) As Boolean
Private Declare Function InternetAutodialHangup Lib "wininet.dll" (ByVal dwReserved As Long) As Long
Public Function CheckConnection() As Boolean
    Dim Flags As Long
    On Error GoTo ErrHandler
    
    'Check if the user is connected to the internet by modem:
    Temp = InternetGetConnectedState(Flags, 0)
    If Flags And 1 Then
        CheckConnection = True
    Else
        CheckConnection = False
    End If
    
    'Return value and exit function
    Exit Function
    
ErrHandler:
    MsgBox modErrHandling.ErrHandler(Err), vbCritical, "CostTracker"
    Resume Next
        
End Function
Public Sub Disconnect()
    On Error GoTo ErrHandler
    
    'Terminate the connection
    frmConnection.mnuClaimable.Checked = 0
    InternetAutodialHangup (0)
    
    Exit Sub
ErrHandler:
    MsgBox modErrHandling.ErrHandler(Err), vbCritical, "CostTracker"
    Resume Next
End Sub
