Attribute VB_Name = "modErrHandling"
Private PreviousError As Integer
'This function is used to create error messages. I could have used err.description, but I wanted to be able to easily adjust the program to another language.

Public Function ErrHandler(ErrCode As ErrObject)
    
    'If the same error occurs two times in a row, the program should better end, in most cases
    If PreviousError = ErrCode.Number Then
        MsgBox "The program has stopped responding. Will now stop", vbCritical, "CostTracker"
        modSystemTray.RemoveIcon
        End
    End If
    
    'Don't remember errors higher than 32767, otherwise I would have to use a bigger variable
    If Abs(ErrCode.Number) <= 32767 Then PreviousError = ErrCode.Number
    
    Select Case ErrCode
        Case 7: ErrHandler = "An error has occurred. No memory available"
        Case 13: ErrHandler = "An error has occurred. Invalid input"
        Case 18: ErrHandler = "An error has occurred. Interrupted by user"
        Case 20: ErrHandler = "An error has occurred. Resume next without error"
        Case 52: ErrHandler = "An error has occurred. Invalid filename"
        Case 53: ErrHandler = "An error has occurred. File not found"
        Case 55: ErrHandler = "An error has occurred. File is already in use"
        Case 57: ErrHandler = "An error has occurred. Device failure"
        Case 58: ErrHandler = "An error has occurred. File already exists"
        Case 61: ErrHandler = "An error has occurred. Disk is full"
        Case 62: ErrHandler = "An error has occurred. Invalid file"
        Case 63: ErrHandler = "An error has occurred. Invalid file"
        Case 67: ErrHandler = "An error has occurred. Too many files"
        Case 68: ErrHandler = "An error has occurred. Device not available"
        Case 70: ErrHandler = "An error has occurred. Access denied"
        Case 71: ErrHandler = "An error has occurred. Disk not ready"
        Case 75: ErrHandler = "An error has occurred. Fileaccess denied"
        Case 76: ErrHandler = "An error has occurred. Path not found"
        Case 94: ErrHandler = "An error has occurred. Invalid input"
        Case -2147467259: ErrHandler = "An error has occurred. Can't locate database"
        Case Else: ErrHandler = "An unknown error has occurred. Code " & ErrCode & ": " & Err.Description
    End Select

    ErrHandler = ErrHandler & ". Will try to continue. The program may stop responding."
End Function
