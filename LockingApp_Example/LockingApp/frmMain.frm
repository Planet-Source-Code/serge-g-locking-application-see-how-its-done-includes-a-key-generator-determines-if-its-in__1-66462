VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00FF0000&
   BorderStyle     =   0  'None
   Caption         =   "Unlock Application"
   ClientHeight    =   2895
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4695
   ClipControls    =   0   'False
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2895
   ScaleWidth      =   4695
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
      Caption         =   "x"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4230
      TabIndex        =   6
      Top             =   150
      Width           =   255
   End
   Begin VB.CommandButton cmdGetKey 
      Caption         =   "Get Key"
      Enabled         =   0   'False
      Height          =   375
      Left            =   1800
      TabIndex        =   5
      Top             =   2385
      Width           =   1095
   End
   Begin VB.CommandButton cmdReset 
      Caption         =   "Reset"
      Height          =   375
      Left            =   1800
      TabIndex        =   4
      Top             =   2400
      Width           =   1095
   End
   Begin VB.CommandButton cmdSubmit 
      Caption         =   "Submit"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   300
      Left            =   3255
      TabIndex        =   2
      Top             =   660
      Width           =   975
   End
   Begin VB.TextBox txtKey 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   90
      TabIndex        =   1
      Top             =   690
      Width           =   2850
   End
   Begin VB.TextBox txtUN 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   90
      TabIndex        =   0
      Top             =   255
      Width           =   2850
   End
   Begin VB.Image Image2 
      Height          =   240
      Left            =   3900
      Picture         =   "frmMain.frx":014A
      Top             =   2415
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image Image1 
      Height          =   240
      Left            =   3555
      Picture         =   "frmMain.frx":0294
      Top             =   2415
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Label lblLocked 
      Alignment       =   2  'Center
      Caption         =   "This Application Is Locked"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   510
      Left            =   240
      TabIndex        =   3
      Top             =   1560
      Width           =   4215
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H8000000F&
      FillStyle       =   0  'Solid
      Height          =   2895
      Left            =   0
      Shape           =   4  'Rounded Rectangle
      Top             =   0
      Width           =   4695
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const LWA_COLORKEY = &H1
Const GWL_EXSTYLE = (-20)
Const WS_EX_LAYERED = &H80000
Const BM_SETSTATE = &HF3

Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32.dll" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long

Dim isEnd As Boolean
Dim is1 As Boolean
Dim is2 As Boolean

Private Sub cmdGetKey_Click()

    txtKey.Text = GenerateKeyNumber(LCase(Trim(txtUN.Text)))

End Sub

Private Sub cmdReset_Click()

    resetKey
    setVisibles
    txtUN.Text = ""
    txtKey.Text = ""
    Load Me
    Me.Show
    Load frmEntered
    frmEntered.Timer1.Enabled = True
    
End Sub

Private Sub cmdSubmit_Click()

    UnlockProgram LCase(Trim(txtUN.Text)), CLng(txtKey.Text)
    
    setVisibles
    
    If Unlocked = True Then
        isEnd = True
        Unload Me
    End If

End Sub

Private Sub Command1_Click()

    isEnd = True
    Unload frmEntered

End Sub

Private Sub dfg_Click()

End Sub

Private Sub Form_Load()

    CLR = RGB(0, 0, 255)
    Ret = GetWindowLong(Me.hwnd, GWL_EXSTYLE)
    Ret = Ret Or WS_EX_LAYERED
    SetWindowLong Me.hwnd, GWL_EXSTYLE, Ret
    SetLayeredWindowAttributes Me.hwnd, CLR, 0, LWA_COLORKEY
  
    CheckLockedStatus
    isEnd = False
    is1 = False
    is2 = False
    
    setVisibles
    
'    frmEntered.Show
'    Unload Me
    
'    txtUN.SetFocus

End Sub

Private Sub Form_Unload(Cancel As Integer)

    If isEnd = False Then Cancel = -1
'    CheckLockedStatus
'    If Unlocked = False Then End
'    MsgBox (Cancel)

End Sub

Private Sub txtKey_Change()

    If Len(txtKey.Text) > 0 Then
        is2 = True
    Else
        is2 = False
    End If
    
    enableCmd

End Sub

Private Sub txtKey_KeyPress(KeyAscii As Integer)

    
    If KeyAscii = 8 Then Exit Sub
    
    If KeyAscii < Asc(0) Or KeyAscii > Asc(9) Then
        KeyAscii = 0
    End If

End Sub

Private Sub txtUN_Change()

    If Len(txtUN) > 0 Then
        is1 = True
        cmdGetKey.Enabled = True
    Else
        is1 = False
        cmdGetKey.Enabled = False
    End If
    
    enableCmd

End Sub

Sub enableCmd()

    If is1 = True And is2 = True Then
        cmdSubmit.Enabled = True
    Else
        cmdSubmit.Enabled = False
    End If

End Sub

Sub setVisibles()

    lblLocked.Visible = Not Unlocked
    txtUN.Visible = Not Unlocked
    txtKey.Visible = Not Unlocked
    cmdSubmit.Visible = Not Unlocked
    cmdGetKey.Visible = Not Unlocked
    cmdReset.Visible = Unlocked
    If Unlocked Then
        Me.Caption = "Program Unlocked"
        Me.Icon = Image1.Picture
    Else
        Me.Caption = "Unlock Application"
        Me.Icon = Image2.Picture
    End If
        
    If RunningInIDE = False Then
        cmdGetKey.Visible = False
    ElseIf RunningInIDE And Unlocked = False Then
        cmdGetKey.Visible = True
    End If
    
End Sub
