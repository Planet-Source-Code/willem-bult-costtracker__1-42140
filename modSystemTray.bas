Attribute VB_Name = "modSystemTray"
Option Explicit

Private Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean

Private Const NIM_ADD = &H0
Private Const NIM_MODIFY = &H1
Private Const NIM_DELETE = &H2
Private Const NIF_MESSAGE = &H1
Private Const NIF_ICON = &H2
Private Const NIF_TIP = &H4
Private Const WM_MOUSEMOVE = &H200
Private Const WM_LBUTTONDOWN = &H201
Private Const WM_LBUTTONDBLCLK = &H203

Private Type NOTIFYICONDATA
    cbSize As Long
    hwnd As Long
    uId As Long
    uFlags As Long
    uCallBackMessage As Long
    hIcon As Long
    szTip As String * 64
End Type

Private IconData As NOTIFYICONDATA
Public Sub PlaceIcon()
    'This sub places an icon in the taskbar. I use this sub when te program loads
    On Error GoTo ErrHandler
    
    With IconData
        .cbSize = Len(IconData)
        .hwnd = frmConnection.hwnd
        .uId = vbNull
        .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
        .uCallBackMessage = WM_MOUSEMOVE
        .hIcon = frmConnection.imgIcon.Picture
        .szTip = "CostTracker" & vbNullChar
    End With
    Shell_NotifyIcon NIM_ADD, IconData
    
    Exit Sub
ErrHandler:
    MsgBox modErrHandling.ErrHandler(Err), vbCritical, "CostTracker"
    Resume Next
End Sub
Public Sub ModifyIcon(IconImage)
    'This sub modifies the icon. I use different icons for connected and disconnected state.
    On Error GoTo ErrHandler
    
    IconData.hIcon = IconImage
    IconData.uFlags = NIF_ICON
    Shell_NotifyIcon NIM_MODIFY, IconData
    
    Exit Sub
ErrHandler:
    MsgBox modErrHandling.ErrHandler(Err), vbCritical, "CostTracker"
    Resume Next
End Sub
Public Sub ModifyTip(Tip)
    'This sub modifies the so called tip. When the user points his mouse to the icon, it shows the current session's costs. I use this sub to change the tip at that time
    On Error GoTo ErrHandler
    
    IconData.uFlags = NIF_TIP
    IconData.szTip = Tip
    Shell_NotifyIcon NIM_MODIFY, IconData
    
    Exit Sub
ErrHandler:
    MsgBox modErrHandling.ErrHandler(Err), vbCritical, "CostTracker"
    Resume Next
End Sub
Public Sub RemoveIcon()
    'This sub removes the icon from the taskbar. I use this sub when the program ends
    On Error GoTo ErrHandler
    
    Shell_NotifyIcon NIM_DELETE, IconData
    
    Exit Sub
ErrHandler:
    MsgBox modErrHandling.ErrHandler(Err), vbCritical, "CostTracker"
    Resume Next
End Sub
