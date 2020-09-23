Attribute VB_Name = "Module1"
Public Unlocked As Boolean

Public Const APP_NAME = "MyApp"
Public Const SECTION_NAME = "Locking"
Public Const KEY_NAME = "Locked status"

Public Sub CheckLockedStatus()
    
    Dim temp As String
    temp = GetSetting(APP_NAME, SECTION_NAME, KEY_NAME, "")
    If temp = "unlocked" Then
        Unlocked = True
    Else
        Unlocked = False
    End If

End Sub

Public Function GenerateKeyNumber(Username As String) As Long

    Dim i As Integer
    Dim s As String * 1
    Dim key As Long
    
    key = 0
    For i = 1 To Len(Username)
        s = Mid(Username, i, 1)
        key = key + Asc(s)
    Next i
    
    GenerateKeyNumber = Int(key * 12345.67)

End Function

Public Sub UnlockProgram(Username As String, key As Long)

    Dim CorrectKey As Long

    CorrectKey = GenerateKeyNumber(Username)

    If key = CorrectKey Then
        SaveSetting APP_NAME, SECTION_NAME, KEY_NAME, "unlocked"
        MsgBox "Program successfully unlocked"
        Unlocked = True
    Else
        MsgBox "Invalid user name or key"
    End If

End Sub

Public Function RunningInIDE() As Boolean

    On Error Resume Next
    Debug.Print 1 / 0
    If Err.Number = 0 Then
        RunningInIDE = False
    Else
        RunningInIDE = True
    End If

End Function

Sub resetKey()

    On Error Resume Next
    
    DeleteSetting APP_NAME, SECTION_NAME, KEY_NAME
    Unlocked = False

End Sub
