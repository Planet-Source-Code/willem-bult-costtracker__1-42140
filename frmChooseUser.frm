VERSION 5.00
Begin VB.Form frmChooseUser 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Select your username"
   ClientHeight    =   3450
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   2820
   Icon            =   "frmChooseUser.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3450
   ScaleWidth      =   2820
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Appearance      =   0  'Flat
      Caption         =   "Cancel"
      Height          =   375
      Left            =   1440
      TabIndex        =   3
      Top             =   3000
      Width           =   1215
   End
   Begin VB.ListBox lstPasswords 
      Height          =   450
      ItemData        =   "frmChooseUser.frx":0442
      Left            =   2160
      List            =   "frmChooseUser.frx":0444
      TabIndex        =   5
      Top             =   1680
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton cmdOK 
      Appearance      =   0  'Flat
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3000
      Width           =   1215
   End
   Begin VB.TextBox txtPassword 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   120
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   2520
      Width           =   2535
   End
   Begin VB.ListBox lstUsers 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      Height          =   1980
      ItemData        =   "frmChooseUser.frx":0446
      Left            =   120
      List            =   "frmChooseUser.frx":0448
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   120
      Width           =   2535
   End
   Begin VB.Label lblPassword 
      Caption         =   "Password:"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   2280
      Width           =   1335
   End
End
Attribute VB_Name = "frmChooseUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This form contains a list with users and a list with the password of these users.
'The list of passwords is not visible to the user and is used to check if the entered password is correct.
'The passwords in the list have already been decrypted, so that doesn't have to be done anymore.

Private Sub cmdCancel_Click()
    'If the user pressed cancel: username is unknown, which means nobody has succesfuly identified himself
    On Error GoTo ErrHandler
    
    modUser.UserName = "Unknown"
    
    Exit Sub
ErrHandler:
    MsgBox modErrHandling.ErrHandler(Err), vbCritical, "CostTracker"
    Resume Next
End Sub
Private Sub cmdOK_Click()
    'If a user clicked OK, we have to check if a username is selected and if so, if the password matches. The passwords are not case-sensitive
    On Error GoTo ErrHandler
    
    If lstUsers.ListIndex < 0 Then
        MsgBox "Please select your username", vbExclamation, "CostTracker"
        Exit Sub
    ElseIf UCase(Replace(txtPassword.Text, " ", "")) <> UCase(lstPasswords.List(lstUsers.ListIndex)) Then
        MsgBox "Password is incorrect", vbExclamation, "CostTracker"
        Exit Sub
    End If
            
    modUser.UserName = lstUsers.Text
    
    Exit Sub
ErrHandler:
    MsgBox modErrHandling.ErrHandler(Err), vbCritical, "CostTracker"
    Resume Next
End Sub

Private Sub Form_Load()
    'When the form loads, show it and give focus to the userlist
    On Error GoTo ErrHandler
    
    Me.Show
    lstUsers.SetFocus
    
    Exit Sub
ErrHandler:
    MsgBox modErrHandling.ErrHandler(Err), vbCritical, "CostTracker"
    Resume Next
End Sub

Private Sub lstUsers_KeyPress(KeyAscii As Integer)
    'If the focus is still at the userlist, and the user presses a key, I assume that the user wants to enter his password.
    'So: I give the focus to the password textbox, insert the caracter that matches the pressed key and set the cursor at the end of the textbox.
    'After that we set the KeyAscii to zero again so that the userlist doesn't receive it
    'This whole procedure makes it a lot easier to log in. Otherwise you would have to press tab first, and I experienced that a lot of beginning computer users (my family) grabs the mouse to give the textbox the focus.
    
    On Error GoTo ErrHandler
    
    If KeyAscii > 48 And KeyAscii < 122 Then
        txtPassword.SetFocus
        txtPassword.Text = txtPassword.Text & Chr(KeyAscii)
        txtPassword.SelStart = Len(txtPassword.Text)
        KeyAscii = 0
    End If
    
    Exit Sub
ErrHandler:
    MsgBox modErrHandling.ErrHandler(Err), vbCritical, "CostTracker"
    Resume Next
End Sub
