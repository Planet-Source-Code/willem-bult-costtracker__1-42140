Attribute VB_Name = "modMain"
' CostTracker is a program to keep track of the telephonecosts that you get when you have a dial-up internet connection.
' With this program you can add multiple users and administre for each user how much money they spent.
' This program has a configuration panel in which you can add users, make users adminsistrators and enter the telephone rates that your provider uses.
' It's also possible for a user, if he has the user rights to do so, to mark a session as claimable.
' For example: my parents pay my internetcosts for school, but not for other things. So, when I connect to the internet for school I can mark that session as claimable.
'
' PLEASE VOTE FOR THIS CODE ON PLANETSOURCECODE.COM IF YOU LIKE IT, I WOULD HIGHLY APPRECIATE IT

Public Connected As Boolean
Private LastUpdate As Single

' The connectionrates are stored in a 2-dimensional array of 8 x 24 (7 weeks and 24 hours a day)
' The extra 'week' is for the costs that some providers use for making the connection.
Public Rates(1 To 8, 0 To 23)

Private StartTime As String
Private StartDate As String
Public Costs As Currency
Private MonthCosts As Currency
Private Claimable As Boolean
Private Starttimer As Single
Public UserName As String

' This module contains the main sub of the program. From this sub the program checks whether or not you are connected to the internet and if so, calculates the period and cost of the current connection.

Private Sub Main()
    
    On Error GoTo ErrHandler
    
    Dim RestTimer As Single

    'First, the program loads the telephone rates with a sub in the database module
    modDatabase.GetRates
    
    'Now we add an icon to the systemtray.
    modSystemTray.PlaceIcon
    
    Do
        'A little pause
        RestTimer = Timer: While Timer < RestTimer + 1: DoEvents: Wend
        
        'If the user is connected to the internet we need to add the costs every minute
        If Connected = True And Timer >= LastUpdate + 60 Then
            Costs = Costs + ((Timer - LastUpdate) / 60) * Rates(Weekday(Now, vbMonday), Hour(Now))
            LastUpdate = Timer
            modDatabase.AddSession UserName, StartTime, StartDate, Claimable, Costs
        End If
       
        'If the connection state changes: the user connects or disconnects then some things have to be done
        If modConnection.CheckConnection <> Connected Then
            Connected = modConnection.CheckConnection
            'If the user connected:
            If Connected = True Then
                'Set some default values
                frmConnection.mnuClaimable.Checked = False
                Claimable = False
                LastUpdate = Timer
                Starttimer = Timer
                StartTime = Time
                StartDate = Date
                'Add the costs of connecting
                Costs = Rates(8, Hour(Now))
                
                'Now we ask the user to identify himself:
                UserName = modUser.ValidateUser(60, False)
                
                'If the user couldn't identify: terminate the connection and add the current costs to Unknown's account
                If UserName = "Unknown" Or UserName = "" Then
                    UserName = ""
                    Connected = False
                    modConnection.Disconnect
                    modDatabase.AddSession "Unknown", StartTime, StartDate, False, Costs + ((Timer - Starttimer) / 60) * Rates(Weekday(Now, vbMonday), Hour(Now))
                Else
                'If he could identify:
                    'let's see how much money he has already spent this month
                    MonthCosts = modDatabase.GetMonthCosts(UserName, Month(Now), False)
                    'adjust the taskbar icon to the 'connected' icon
                    modSystemTray.ModifyIcon frmConnection.Icon
                    frmConnection.Show
                    'show a windows with the connection information
                    frmConnection.WindowState = vbNormal
                End If
            Else
            'if the user disconnected:
                'Add the current session to the database
                modDatabase.AddSession UserName, StartTime, StartDate, Claimable, Costs
                'reset the username
                UserName = ""
                'unload the connection information window and change the icon to the 'disconnected' icon
                Unload frmConnection
                frmConnection.mnuClaimable.Checked = 0
                modSystemTray.ModifyIcon frmConnection.imgIcon.Picture
            End If
            
        End If
        
        'If the claimable state changes and we're still connected:
        If frmConnection.mnuClaimable.Checked <> Claimable And Connected = True Then
            Claimable = frmConnection.mnuClaimable.Checked
            
            ' in case the user marks the session non-claimable from this point on or if the user has connected more than three minutes ago and marks the session claimable, add the session to the database and reset all timers and counters.
            If (Claimable = True And Timer - Starttimer > 180) Or Claimable = False Then
                modDatabase.AddSession UserName, StartTime, StartDate, Claimable, Costs
                Costs = 0
                StartDate = Date
                StartTime = Time
                Starttimer = Timer
            End If
        End If
        
        'If the window with connection information is being shown: update information
        If Connected = True And frmConnection.WindowState = vbNormal Then
            frmConnection.lblUser.Caption = UserName
            frmConnection.lblBegin.Caption = StartTime
            frmConnection.lblSessionTime.Caption = Int(CInt(Timer - Starttimer) / 3600) & ":" & Int((CInt(Timer - Starttimer) Mod 3600) / 60) & ":" & (CInt(Timer - Starttimer) Mod 3600) Mod 60
            frmConnection.lblCosts.Caption = CInt(Costs * 100) / 100
            frmConnection.lblMonthCosts.Caption = MonthCosts + CInt(Costs * 100) / 100
            frmConnection.lblClaimable.Caption = Claimable
        End If
    
    Loop
    
    Exit Sub
ErrHandler:
    'Use the ErrHandling module to generate an error description and display it.
    MsgBox modErrHandling.ErrHandler(Err), vbCritical, "CostTracker"
    Resume Next
End Sub
