VERSION 5.00
Begin VB.Form frmConfiguration 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Configuration Panel"
   ClientHeight    =   8040
   ClientLeft      =   1155
   ClientTop       =   1815
   ClientWidth     =   12735
   Icon            =   "frmConfiguration.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8040
   ScaleWidth      =   12735
   Begin VB.CommandButton cmdDelete 
      Appearance      =   0  'Flat
      Caption         =   "Delete"
      Height          =   495
      Left            =   4440
      TabIndex        =   5
      Top             =   6360
      Width           =   1575
   End
   Begin VB.CommandButton cmdAdd 
      Appearance      =   0  'Flat
      Caption         =   "Add / Edit"
      Height          =   495
      Left            =   2760
      TabIndex        =   4
      Top             =   6360
      Width           =   1575
   End
   Begin VB.TextBox txtUserName 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      Height          =   315
      Left            =   2760
      TabIndex        =   1
      Top             =   4680
      Width           =   3255
   End
   Begin VB.CommandButton cmdCancel 
      Appearance      =   0  'Flat
      Caption         =   "Cancel"
      Height          =   495
      Left            =   4200
      TabIndex        =   8
      Top             =   7440
      Width           =   1935
   End
   Begin VB.CommandButton cmdUndo 
      Appearance      =   0  'Flat
      Caption         =   "Undo everything"
      Height          =   495
      Left            =   2160
      TabIndex        =   7
      Top             =   7440
      Width           =   1935
   End
   Begin VB.CommandButton cmdSave 
      Appearance      =   0  'Flat
      Caption         =   "Save changes"
      Height          =   495
      Left            =   120
      TabIndex        =   6
      Top             =   7440
      Width           =   1935
   End
   Begin VB.ListBox lstPasswords 
      Height          =   1425
      Left            =   6360
      TabIndex        =   23
      Top             =   4560
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox txtPassword 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   2760
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   5400
      Width           =   3255
   End
   Begin VB.CheckBox chkAdministrator 
      Appearance      =   0  'Flat
      Caption         =   "Administrator"
      Height          =   375
      Left            =   2760
      MaskColor       =   &H8000000F&
      TabIndex        =   3
      Top             =   5880
      UseMaskColor    =   -1  'True
      Width           =   2295
   End
   Begin VB.ListBox lstUsers 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      Height          =   2370
      Left            =   120
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   4440
      Width           =   2415
   End
   Begin VB.Frame fraRates 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3375
      Left            =   120
      TabIndex        =   17
      Top             =   480
      Width           =   12495
      Begin VB.TextBox txtTarif 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         Height          =   375
         Index           =   0
         Left            =   960
         TabIndex        =   16
         Top             =   360
         Width           =   375
      End
      Begin VB.Line Line2 
         BorderWidth     =   2
         X1              =   960
         X2              =   960
         Y1              =   0
         Y2              =   3360
      End
      Begin VB.Label lblDay 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Conn. rate"
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   7
         Left            =   0
         TabIndex        =   26
         Top             =   3000
         Width           =   975
      End
      Begin VB.Line Line3 
         BorderWidth     =   2
         X1              =   0
         X2              =   12480
         Y1              =   360
         Y2              =   360
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Day / Hour"
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   0
         TabIndex        =   19
         Top             =   0
         Width           =   975
      End
      Begin VB.Label lblHour 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   0
         Left            =   960
         TabIndex        =   18
         Top             =   0
         Width           =   375
      End
      Begin VB.Label lblDay 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Sunday"
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   6
         Left            =   0
         TabIndex        =   15
         Top             =   2520
         Width           =   975
      End
      Begin VB.Label lblDay 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Saturday"
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   5
         Left            =   0
         TabIndex        =   14
         Top             =   2160
         Width           =   975
      End
      Begin VB.Label lblDay 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Friday"
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   4
         Left            =   0
         TabIndex        =   13
         Top             =   1800
         Width           =   975
      End
      Begin VB.Label lblDay 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Thursday"
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   3
         Left            =   0
         TabIndex        =   12
         Top             =   1440
         Width           =   975
      End
      Begin VB.Label lblDay 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Wednesday"
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   2
         Left            =   0
         TabIndex        =   11
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label lblDay 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Monday"
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   0
         Left            =   0
         TabIndex        =   9
         Top             =   360
         Width           =   975
      End
      Begin VB.Label lblDay 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Tuesday"
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   1
         Left            =   0
         TabIndex        =   10
         Top             =   720
         Width           =   975
      End
   End
   Begin VB.Label Label6 
      Caption         =   "Username:"
      Height          =   255
      Left            =   2760
      TabIndex        =   25
      Top             =   4440
      Width           =   1455
   End
   Begin VB.Line Line5 
      X1              =   120
      X2              =   720
      Y1              =   7200
      Y2              =   7200
   End
   Begin VB.Label Label4 
      Caption         =   "Password:"
      Height          =   255
      Left            =   2760
      TabIndex        =   22
      Top             =   5160
      Width           =   1575
   End
   Begin VB.Line Line4 
      X1              =   120
      X2              =   960
      Y1              =   4320
      Y2              =   4320
   End
   Begin VB.Label Label3 
      Caption         =   "Users:"
      Height          =   255
      Left            =   120
      TabIndex        =   21
      Top             =   4080
      Width           =   1095
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   2400
      Y1              =   360
      Y2              =   360
   End
   Begin VB.Label Label2 
      Caption         =   "Telephone rates: (per minute)"
      Height          =   255
      Left            =   120
      TabIndex        =   20
      Top             =   120
      Width           =   2415
   End
   Begin VB.Label Label5 
      Caption         =   "Save:"
      Height          =   255
      Left            =   120
      TabIndex        =   24
      Top             =   6960
      Width           =   855
   End
End
Attribute VB_Name = "frmConfiguration"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAdd_Click()
    'When the Add/Edit button is pressed, this sub is called
    
    On Error GoTo ErrHandler
    
    'Check if there is a password and a username
    If Len(Replace(txtPassword.Text, " ", "")) = 0 Then
        MsgBox "Invalid password", vbExclamation, "CostTracker"
        Exit Sub
    ElseIf Len(Replace(txtPassword.Text, " ", "")) = 0 Then
        MsgBox "Invalid username", vbExclamation, "CostTracker"
        Exit Sub
    End If
    
    'If the user already exists: only alter the values and exit instead of adding a new user
    For i% = 0 To lstUsers.ListCount - 1
        If lstUsers.List(i%) = txtUserName.Text Then
            lstUsers.ItemData(i%) = chkAdministrator.Value
            lstPasswords.List(i%) = txtPassword
            Exit Sub
        End If
    Next i%
            
    'If the user didn't already exist: add the user to the list
    lstUsers.AddItem txtUserName.Text
    lstUsers.ItemData(lstUsers.NewIndex) = chkAdministrator.Value
    lstPasswords.AddItem txtPassword.Text, lstUsers.NewIndex
    
    Exit Sub
ErrHandler:
    MsgBox modErrHandling.ErrHandler(Err), vbCritical, "CostTracker"
    Resume Next
       
End Sub

Private Sub cmdCancel_Click()
    'If cancel is pressed: unload the configuration panel
    Unload Me
End Sub

Private Sub cmdDelete_Click()
    On Error GoTo ErrHandler
    
    'Delete a user
    For i% = 0 To lstUsers.ListCount - 1
        If lstUsers.List(i%) = txtUserName Then
            lstUsers.RemoveItem i%
            lstPasswords.RemoveItem i%
            Exit For
        End If
    Next i%
    
    Exit Sub
ErrHandler:
    MsgBox modErrHandling.ErrHandler(Err), vbCritical, "CostTracker"
    Resume Next
End Sub

Private Sub cmdSave_Click()
    On Error GoTo ErrHandler
    
    'Load the new rates into the array
    For i% = 0 To 23
        For i2% = 0 To 7
            Rates(i2% + 1, i%) = txtTarif(i2% * 24 + i%).Text
        Next i2%
    Next i%
    'Save the rates in the database
    modDatabase.SaveRates
    'Save the users in the database
    modDatabase.SaveUsers frmConfiguration
    
    MsgBox "The new configuration will be used immediately", vbInformation, "CostTracker"
    Unload Me
    
    Exit Sub
ErrHandler:
    MsgBox modErrHandling.ErrHandler(Err), vbCritical, "CostTracker"
    Resume Next
End Sub

Private Sub cmdUndo_Click()
    On Error GoTo ErrHandler
    
    'If undo is pressed: reload the current configuration
    
    For i% = 0 To 23
        For i2% = 0 To 7
            txtTarif(i2% * 24 + i%).Text = Rates(i2% + 1, i%)
        Next i2%
    Next i%
    
    modDatabase.GetUsers False, frmConfiguration
    
    Exit Sub
ErrHandler:
    MsgBox modErrHandling.ErrHandler(Err), vbCritical, "CostTracker"
    Resume Next
            
End Sub

Private Sub Form_Load()
    
    On Error GoTo ErrHandler
    
    'Whe need to build up the field of textboxes in which the user can specify the telephone rates.
    For i% = 0 To 23
        If i% > 0 Then
            Load lblHour(i%)
            With lblHour(i%)
                .Left = i% * 360 + 940
                .Caption = i%
                .Visible = True
            End With
        End If
        For i2% = 0 To 7
            If i2% * 24 + i% <> 0 Then Load txtTarif(i2% * 24 + i%)
            With txtTarif(i2% * 24 + i%)
                .Left = i% * 360 + 940
                If i% = 0 Then .Left = .Left + 20
                .Top = i2% * 360 + 360
                If i2% = 7 Then .Top = .Top + 120
                .Visible = True
                .Text = Rates(i2% + 1, i%)
                .TabIndex = i2% * 24 + i%
            End With
        Next i2%
    Next i%
    
    'Adjust the size of the form and some contents
    fraRates.Width = 360 * 23 + 1450
    Line3.X2 = fraRates.Width - 150
    frmConfiguration.Width = fraRates.Width + 200
    
    'Get the users and passwords. We don't want just the administrators, but all the users, so administrators=false
    modDatabase.GetUsers False, frmConfiguration
    
    'Show the form
    Me.Show
    
    Exit Sub
ErrHandler:
    MsgBox modErrHandling.ErrHandler(Err), vbCritical, "CostTracker"
    Resume Next
    
End Sub


Private Sub lstUsers_Click()
    On Error GoTo ErrHandler
    
    'When a username is selected: update the username,password,administrator fields
    If lstUsers.ListIndex >= 0 Then
        txtUserName.Text = lstUsers.Text
        txtPassword.Text = lstPasswords.List(lstUsers.ListIndex)
        chkAdministrator = Abs(lstUsers.ItemData(lstUsers.ListIndex))
    End If
    
    Exit Sub
ErrHandler:
    MsgBox modErrHandling.ErrHandler(Err), vbCritical, "CostTracker"
    Resume Next
End Sub


Private Sub txtTarif_KeyPress(Index As Integer, KeyAscii As Integer)
    On Error GoTo ErrHandler
    
    'Make sure that only numbers and the comma and backspace can be entered in these textboxes
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 44 And KeyAscii <> 8 Then KeyAscii = 0
    
    Exit Sub
ErrHandler:
    MsgBox modErrHandling.ErrHandler(Err), vbCritical, "CostTracker"
    Resume Next
End Sub


Private Sub txtUserName_Change()
    On Error GoTo ErrHandler
    
    'If the username in the textbox changes: see if this username is in the list and if so, select it so that the other fields will be updated (see lstUsers_Click)
    For i% = 0 To lstUsers.ListCount - 1
        If lstUsers.List(i%) = txtUserName Then lstUsers.ListIndex = i%
    Next i%
    
    Exit Sub
ErrHandler:
    MsgBox modErrHandling.ErrHandler(Err), vbCritical, "CostTracker"
    Resume Next
End Sub

Private Sub txtUserName_KeyPress(KeyAscii As Integer)
    On Error GoTo ErrHandler
    
    'If the username changes: reset the other fields
    txtPassword.Text = ""
    chkAdministrator = 0
    lstUsers.ListIndex = -1
    
    Exit Sub
ErrHandler:
    MsgBox modErrHandling.ErrHandler(Err), vbCritical, "CostTracker"
    Resume Next

End Sub
