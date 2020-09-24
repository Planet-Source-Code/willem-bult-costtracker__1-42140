VERSION 5.00
Begin VB.Form frmConnection 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Costchecker - Connection details"
   ClientHeight    =   1185
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   5610
   Icon            =   "frmConnection.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1185
   ScaleWidth      =   5610
   StartUpPosition =   3  'Windows Default
   Begin VB.Label lblUser 
      Height          =   255
      Left            =   1320
      TabIndex        =   11
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label Label4 
      Caption         =   "User:"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   120
      Width           =   855
   End
   Begin VB.Label lblClaimable 
      Height          =   255
      Left            =   4080
      TabIndex        =   9
      Top             =   840
      Width           =   1455
   End
   Begin VB.Label Label3 
      Caption         =   "Claimable"
      Height          =   255
      Left            =   2880
      TabIndex        =   8
      Top             =   840
      Width           =   1095
   End
   Begin VB.Image imgIcon 
      Height          =   480
      Left            =   6120
      Picture         =   "frmConnection.frx":0442
      Top             =   240
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label lblMonthCosts 
      Height          =   315
      Left            =   4080
      TabIndex        =   7
      Top             =   480
      Width           =   1485
   End
   Begin VB.Label lblCosts 
      Height          =   315
      Left            =   4080
      TabIndex        =   6
      Top             =   120
      Width           =   1485
   End
   Begin VB.Label lblSessionTime 
      Height          =   315
      Left            =   1320
      TabIndex        =   5
      Top             =   840
      Width           =   1335
   End
   Begin VB.Label lblBegin 
      Height          =   315
      Left            =   1320
      TabIndex        =   4
      Top             =   480
      Width           =   1365
   End
   Begin VB.Label Label6 
      Caption         =   "Months costs"
      Height          =   255
      Left            =   2880
      TabIndex        =   3
      Top             =   480
      Width           =   1095
   End
   Begin VB.Label Label5 
      Caption         =   "Session costs:"
      Height          =   255
      Left            =   2880
      TabIndex        =   2
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "Duration:"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Start:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   1095
   End
   Begin VB.Menu mnuSystemTray 
      Caption         =   "System Tray"
      Visible         =   0   'False
      Begin VB.Menu mnuDetails 
         Caption         =   "Users information"
      End
      Begin VB.Menu mnuClaimable 
         Caption         =   "Claimable"
      End
      Begin VB.Menu mnuSpacer 
         Caption         =   "-"
      End
      Begin VB.Menu mnuConfiguration 
         Caption         =   "Configuration panel"
      End
      Begin VB.Menu mnuSpacer2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "frmConnection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'This sub handles the functions of the systemtray icon
    On Error GoTo ErrHandler
    
    Const WM_MOUSEMOVE = &H200
    Const WM_LBUTTONDOWN = &H201
    Const WM_LBUTTONDBLCLK = &H203
    Const WM_RBUTTONDOWN = &H204
    
    Select Case X / Screen.TwipsPerPixelX
        'If the left button is clicked: display the menu
        Case WM_LBUTTONDOWN
            Me.PopupMenu mnuSystemTray, , , , mnuDetails
        'If the right button is clicked: display the menu
        Case WM_RBUTTONDOWN
            Me.PopupMenu mnuSystemTray, , , , mnuDetails
        'If the left button is clicked twice: perform default action
        Case WM_LBUTTONDBLCLK
            mnuDetails_Click
        'If the mouse is only positioned on the icon: display the tip, which includes the current costs
        Case WM_MOUSEMOVE
            If modMain.Connected = True Then modSystemTray.ModifyTip "CostTracker. Sessie: " & Int(Costs * 100) / 100 & " €"
    End Select
    
    Exit Sub
ErrHandler:
    MsgBox modErrHandling.ErrHandler(Err), vbCritical, "CostTracker"
    Resume Next
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo ErrHandler
    'Hide the window
    
    Me.WindowState = vbMinimized
    Me.Hide
    Cancel = 1
    
    Exit Sub
ErrHandler:
    MsgBox modErrHandling.ErrHandler(Err), vbCritical, "CostTracker"
    Resume Next
End Sub

Private Sub mnuClaimable_Click()
    'This sub is called if a user wants to change the claimable state of this session
    On Error GoTo ErrHandler
        
    'Check if the computer is connected to the internet
    If modMain.Connected = False Then
        MsgBox "Not connected to the internet", vbInformation, "CostTracker"
        Exit Sub
    End If
    
    If mnuClaimable.Checked = True Then
        'If the session was claimable upuntil now: make it non-claimable
        mnuClaimable.Checked = False
    Else
        'If the user wants to make it claimable: let's get authorization from an administrator.
        'The user doesn't have to get authorization if he is an administrator himself. That's handled in the modUser module
        
        If modUser.ValidateUser(0, True) = "Unknown" Then
            MsgBox "You are not allowed to start a claimable session", vbExclamation, "CostTracker"
        Else
            mnuClaimable.Checked = True
        End If
    End If
    
    Exit Sub
ErrHandler:
    MsgBox modErrHandling.ErrHandler(Err), vbCritical, "CostTracker"
    Resume Next
End Sub

Private Sub mnuConfiguration_Click()
    'This sub is called when a user wants to open the configuration panel
    On Error GoTo ErrHandler
    
    'Let's get administrator authorization
    If modUser.ValidateUser(0, True) = "Unknown" Then
        MsgBox "You are not allowed to change the configuration", vbExclamation, "CostTracker"
        Exit Sub
    End If
    
    'Load the configuration panel
    Load frmConfiguration
    
    Exit Sub
ErrHandler:
    MsgBox modErrHandling.ErrHandler(Err), vbCritical, "CostTracker"
    Resume Next
End Sub

Private Sub mnuDetails_Click()
    'If the user is connected to the internet: show this form with information about the current session.
    'If the user is disconnected: show frmDetails which displays information about the history sessions.
    On Error GoTo ErrHandler
    
    If modMain.Connected = True Then
        Me.Show
        Me.WindowState = vbNormal
    Else
        Load frmDetails
    End If
    
    Exit Sub
ErrHandler:
    MsgBox modErrHandling.ErrHandler(Err), vbCritical, "CostTracker"
    Resume Next
End Sub

Private Sub mnuExit_Click()
    'If a user tries to exit the program, this sub is called
    
    On Error GoTo ErrHandler
    
    'Only administrators may exit the program, so lets get authorization
    If modUser.ValidateUser(0, True) = "Unknown" Then
        MsgBox "You are not allowed to exit this program", vbExclamation, "CostTracker"
        Exit Sub
    End If
    
    'Remove the systemtray icon and end the program
    modSystemTray.RemoveIcon
    End
    
    Exit Sub
ErrHandler:
    MsgBox modErrHandling.ErrHandler(Err), vbCritical, "CostTracker"
    Resume Next
End Sub
