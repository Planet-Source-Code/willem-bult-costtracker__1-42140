VERSION 5.00
Begin VB.Form frmDetails 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Users information"
   ClientHeight    =   3585
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   4455
   Icon            =   "frmDetails.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3585
   ScaleWidth      =   4455
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete this data"
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   3120
      Width           =   4215
   End
   Begin VB.ListBox lstMonths 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      Height          =   1785
      Left            =   2280
      TabIndex        =   2
      Top             =   360
      Width           =   2055
   End
   Begin VB.ListBox lstPasswords 
      Appearance      =   0  'Flat
      Height          =   420
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.ListBox lstUsers 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      Height          =   1785
      Left            =   120
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   360
      Width           =   2055
   End
   Begin VB.Label Label4 
      Caption         =   "Month:"
      Height          =   255
      Left            =   2280
      TabIndex        =   9
      Top             =   100
      Width           =   2055
   End
   Begin VB.Label Label3 
      Caption         =   "User:"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   100
      Width           =   2055
   End
   Begin VB.Label lblClaimable 
      Height          =   255
      Left            =   1560
      TabIndex        =   6
      Top             =   2760
      Width           =   2775
   End
   Begin VB.Label lblCosts 
      Height          =   255
      Left            =   1560
      TabIndex        =   5
      Top             =   2400
      Width           =   2775
   End
   Begin VB.Label Label1 
      Caption         =   "Claimable costs:"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   2760
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "Total costs:"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   2400
      Width           =   1335
   End
End
Attribute VB_Name = "frmDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdDelete_Click()
    On Error GoTo ErrHandler
    
    'Only administrators can delete connection history so let's get authorization
    If modUser.ValidateUser(0, True) = "Unknown" Then
        MsgBox "You are not alowed to delete this data", vbExclamation, "CostTracker"
        Exit Sub
    End If
    
    'See if a user and month is selected
    If lstUsers.ListIndex < 0 Or lstMonths.ListIndex < 0 Then Exit Sub
    
    'Delete the history of this user in this particular month from the database
    modDatabase.DeleteMonthCosts lstUsers.Text, lstMonths.Text
    
    'Reset the costs fields
    lblCosts.Caption = "0"
    lblClaimable.Caption = "0"
    
    'Reload the fields by running the load sub
    Form_Load
    
    Exit Sub
ErrHandler:
    MsgBox modErrHandling.ErrHandler(Err), vbCritical, "CostTracker"
    Resume Next
End Sub

Private Sub Form_Load()

    On Error GoTo ErrHandler
    
    'Get the users from the database
    'There is a password list on this form that is not used, but it's there to be able to use the same sub as the one that's used to identify the user
    modDatabase.GetUsers False, frmDetails
    'Add an unknown user, for the costs that are made when no user has identified
    lstUsers.AddItem "Unknown"
    'Get the costs of this month
    modDatabase.GetSessionMonths frmDetails
    
    Me.Show
    
    Exit Sub
ErrHandler:
    MsgBox modErrHandling.ErrHandler(Err), vbCritical, "CostTracker"
    Resume Next
End Sub

Private Sub lstMonths_Click()
    On Error GoTo ErrHandler
    
    'If a month is selected: see if a user is also selected
    If lstUsers.ListIndex < 0 Or lstMonths.ListIndex < 0 Then Exit Sub
    
    'Get the overall costs and the claimable costs from the database and display them
    lblCosts.Caption = modDatabase.GetMonthCosts(lstUsers.Text, lstMonths.Text, False)
    lblClaimable.Caption = modDatabase.GetMonthCosts(lstUsers.Text, lstMonths.Text, True)
    
    Exit Sub
ErrHandler:
    MsgBox modErrHandling.ErrHandler(Err), vbCritical, "CostTracker"
    Resume Next

End Sub

Private Sub lstUsers_Click()
    On Error GoTo ErrHandler
    
    'If a user is selected: see if a month is also selected
    If lstUsers.ListIndex < 0 Or lstMonths.ListIndex < 0 Then Exit Sub
    
    'Get the overall costs and the claimable costs from the database and display them
    lblCosts.Caption = modDatabase.GetMonthCosts(lstUsers.Text, lstMonths.Text, False)
    lblClaimable.Caption = modDatabase.GetMonthCosts(lstUsers.Text, lstMonths.Text, True)
    
    Exit Sub
ErrHandler:
    MsgBox modErrHandling.ErrHandler(Err), vbCritical, "CostTracker"
    Resume Next
    
End Sub
