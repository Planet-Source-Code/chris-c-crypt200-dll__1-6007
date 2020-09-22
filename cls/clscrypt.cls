VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clscrypt"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Public CancelFlag%
Public Event PercentDone(ByVal Percent As Integer, ByVal TotalFileSize As Long, ProcessFileSize As Long)
Public Function Get_Output(ByVal lpTempFile As String, ByVal lpData As String) As Integer
Dim hOutput, OLF&
Get_Output = 0
hOutput = FreeFile
Open lpTempFile For Binary Access Write As hOutput
If Err Then
    Get_Output = 1
Else
    OLF& = LOF(hOutput)
    If OLF& = 0 Then
        Put hOutput, 1, lpData
    Else
        Put hOutput, OLF& + 1, lpData
    End If
End If

Close hOutput
End Function
Public Function Get_File(ByVal lpSourceFile As String, ByVal lpTempFile As String, ByVal lpSeed As Integer, ByVal lpKey As String, ByVal lpMode As Integer) As Integer
    ' Open the log file.
    Dim Temp$, hData, BSize, LF&, CFp&
    CancelFlag% = 0
    Get_File = 0
    hData = FreeFile
    Open lpSourceFile For Binary Access Read As hData
    If Err Then
        MsgBox Error$, 48, "Data Security"
    Else
        'Read the file in blocks the size of the transmit buffer.
        BSize = 1024
        LF& = LOF(hData)
        Do Until EOF(hData)
            DoEvents
            'Check for user cancelation
            If CancelFlag% = 1 Then
                Kill lpTempFile
                Exit Do
            End If
            
            ' Don't read too much at the end.
            CFp& = Loc(hData)
            If LF& - Loc(hData) <= BSize Then
               BSize = LF& - Loc(hData) + 1
            End If
      
            ' Read a block of data.
            Temp$ = Space$(BSize)
            Get hData, , Temp$
            If EOF(hData) = True Then Temp$ = Mid(Temp$, 1, Len(Temp$) - 1)
            
            Select Case lpMode
            Case 0 'Encrypt
                'Write Output File
                If Get_Output(lpTempFile, Get_Data(Temp, lpSeed, lpKey, 0)) = 1 Or Err Then
                    MsgBox Error$, 48, "Data Security"
                    Get_File = 1
                    Exit Do
                End If
            Case 1 'Decrypt
                If Get_Output(lpTempFile, Get_Data(Temp, lpSeed, lpKey, 1)) = 1 Or Err Then
                    MsgBox Error$, 48, "Data Security"
                    Get_File = 1
                    Exit Do
                End If
            End Select
            
            RaiseEvent PercentDone(CInt(((CFp& + BSize) / LF&) * 100), LF&, CFp& + BSize)
        Loop
    End If
   
    Close hData

End Function




Public Function Get_Data(ByVal lpData As String, ByVal lpSeed As Integer, ByVal lpKey As String, ByVal lpMode As Integer) As String
Dim L&, K&, Cnt&, Mix$, Key$, Ret$, Tmp%, Final$
If lpData = "" Then
    Get_Data = ""
    Exit Function
Else
    '[START] Generate the consistant Random ASCII code.
    Rnd (-1)
    Randomize (lpSeed)
    L& = Len(lpData)
    For Cnt& = 1 To L&: Mix$ = Mix$ & Chr(Int((255 * Rnd) + 0)): Next
    '[END] Generate the consistant Random ASCII code.
    
    '[START] Duplicat the Ref. Key value to the lenght of lpData
    K& = L& Mod Len(lpKey)
    K& = (L& - K&) / Len(lpKey)
    If K& < 0 Then
        For Cnt& = 1 To L&: Key$ = Key$ & Mid(lpKey, Cnt&, 1): Next
    Else
        For Cnt& = 1 To K&: Key$ = Key$ & lpKey: Next
        K& = L& Mod Len(lpKey)
        For Cnt& = 1 To K&: Key$ = Key$ & Mid(lpKey, Cnt&, 1): Next
    End If
    '[END] Duplicat the Ref. Key value to the lenght of lpData
    
    Select Case lpMode
    Case 0 'Encrypt
        '[START] Sum the Data value with Key value and Random ASCII str value.
        For Cnt& = 1 To L&: Ret$ = Ret$ & Chr(0) & Asc(Mid(lpData, Cnt&, 1)) + Asc(Mid(Key$, Cnt&, 1)) + Asc(Mid(Mix$, Cnt&, 1)): Next
        '[END] Sum the Data value with Key value and Random ASCII str value.
    
        'Remove the 1st NULL Character
        Ret$ = Mid(Ret$, 2)
        '[START] Check for value > 256
        Do While Len(Ret$) <> 0
            If InStr(1, Ret$, Chr(0), vbBinaryCompare) <> 0 Then
                Tmp% = CInt(Left(Ret$, InStr(1, Ret$, Chr(0), vbBinaryCompare) - 1))
                Ret$ = Mid(Ret$, InStr(1, Ret$, Chr(0), vbBinaryCompare) + 1)
            Else
                Tmp% = CInt(Trim(Ret$))
                Ret$ = ""
            End If
            Do While Tmp% > 255
                Tmp% = Tmp% - 255
            Loop
            Final$ = Final$ & Chr(Tmp%)
        Loop
        '[END] Check for value > 256
    Case 1 'Decrypt
        '[START] Minus the Key value and Random ASCII str value from Data Value.
        For Cnt& = 1 To L&: Ret$ = Ret$ & Chr(0) & Asc(Mid(lpData, Cnt&, 1)) - Asc(Mid(Key$, Cnt&, 1)) - Asc(Mid(Mix$, Cnt&, 1)): Next
        '[END] Minus the Key value and Random ASCII str value from Data Value.
    
        'Remove the 1st NULL Character
        Ret$ = Mid(Ret$, 2)
        '[START] Check for value > 256
        Do While Len(Ret$) <> 0
            If InStr(1, Ret$, Chr(0), vbBinaryCompare) <> 0 Then
                Tmp% = CInt(Left(Ret$, InStr(1, Ret$, Chr(0), vbBinaryCompare) - 1))
                Ret$ = Mid(Ret$, InStr(1, Ret$, Chr(0), vbBinaryCompare) + 1)
            Else
                Tmp% = CInt(Trim(Ret$))
                Ret$ = ""
            End If
            
            Do While Tmp% < 0
                Tmp% = Tmp% + 255
            Loop
            Final$ = Final$ & Chr(Tmp%)
        Loop
        '[END] Check for value > 256
    End Select
    Get_Data = Final$
End If



End Function

