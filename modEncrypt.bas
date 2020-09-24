Attribute VB_Name = "modEncryption"
'This module contains a simple algorithm I wrote to encrypt and decrypt the passwords.

Public Function Encrypt(Text)
   
    On Error GoTo ErrHandler
    
    For i% = 1 To Len(Text)
        Encrypt = Encrypt & Replace(Space(3 - Len(CStr(Val("&H" & Right(Replace(Space(2 - Len(Hex(255 - Asc(Mid$(Text, i%, 1))))), " ", "0") & Hex(255 - Asc(Mid$(Text, i%, 1))), 1) & Left(Replace(Space(2 - Len(Hex(255 - Asc(Mid$(Text, i%, 1))))), " ", "0") & Hex(255 - Asc(Mid$(Text, i%, 1))), 1))))), " ", "0") & CStr(Val("&H" & Right(Replace(Space(2 - Len(Hex(255 - Asc(Mid$(Text, i%, 1))))), " ", "0") & Hex(255 - Asc(Mid$(Text, i%, 1))), 1) & Left(Replace(Space(2 - Len(Hex(255 - Asc(Mid$(Text, i%, 1))))), " ", "0") & Hex(255 - Asc(Mid$(Text, i%, 1))), 1)))
    Next i%
    
    Exit Function
ErrHandler:
    MsgBox modErrHandling.ErrHandler(Err), vbCritical, "CostTracker"
    Resume Next
End Function
Public Function Decrypt(Crypt)
    On Error GoTo ErrHandler
    
    For i% = 1 To Len(Crypt) Step 3
        Decrypt = Decrypt & Chr(255 - Val("&H" & Right(Replace(Space(2 - Len(Hex(Val(Mid$(Crypt, i%, 3))))), " ", "0") & Hex(Val(Mid$(Crypt, i%, 3))), 1) & Left(Replace(Space(2 - Len(Hex(Val(Mid$(Crypt, i%, 3))))), " ", "0") & Hex(Val(Mid$(Crypt, i%, 3))), 1)))
    Next i%
    
    Exit Function
ErrHandler:
    MsgBox modErrHandling.ErrHandler(Err), vbCritical, "CostTracker"
    Resume Next
End Function
