VERSION 5.00
Begin VB.Form frmKeyGen 
   BackColor       =   &H00808080&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "KeyGen"
   ClientHeight    =   1320
   ClientLeft      =   3870
   ClientTop       =   3765
   ClientWidth     =   2865
   Icon            =   "frmKeyGen.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1320
   ScaleWidth      =   2865
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdGen 
      Caption         =   "Generate"
      Default         =   -1  'True
      Height          =   300
      Left            =   960
      TabIndex        =   2
      Top             =   915
      Width           =   915
   End
   Begin VB.TextBox txtKey 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   285
      Left            =   600
      TabIndex        =   1
      Top             =   420
      Width           =   2100
   End
   Begin VB.TextBox txtName 
      Appearance      =   0  'Flat
      Height          =   255
      Left            =   585
      TabIndex        =   0
      Top             =   45
      Width           =   2115
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Key"
      ForeColor       =   &H8000000E&
      Height          =   240
      Left            =   60
      TabIndex        =   4
      Top             =   465
      Width           =   375
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Name"
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Left            =   45
      TabIndex        =   3
      Top             =   60
      Width           =   555
   End
End
Attribute VB_Name = "frmKeyGen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''API Start'''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Declare Function SetWindowPos Lib "user32" _
    (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, _
    ByVal X As Long, ByVal Y As Long, ByVal cx As Long, _
    ByVal cy As Long, ByVal wFlags As Long) As Long
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Const HWND_TOPMOST = -1        '''''''''''''''''''''''''''''
Const HWND_NOTOPMOST = -2      '''''''''''''''''''''''''''''
Const SWP_NOSIZE = &H1         '''''''''''''''''''''''''''''
Const SWP_NOMOVE = &H2         '''''''''''''''''''''''''''''
Const SWP_NOACTIVATE = &H10    '''''''''''''''''''''''''''''
Const SWP_SHOWWINDOW = &H40    '''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''API End'''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim username As String

Private Sub cmdGen_Click()
    
    If txtName.Text <> "" Then
        txtKey.Enabled = True
        txtKey.BackColor = vbWhite
        username = LCase(Trim(txtName.Text))
        getKey
    Else
        MsgBox ("Please enter your name"), vbOKOnly, "KeyGen"
    End If

End Sub
Sub getKey()

    Dim i As Integer
    Dim s As String * 1
    Dim key As Long
    
    key = 0
    For i = 1 To Len(username)
        s = Mid(username, i, 1)
        key = key + Asc(s)
    Next i
    
    txtKey.Text = Int(key * 12345.67)
    
End Sub

Private Sub setTopMost()

    SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE
    
End Sub

Private Sub Form_Load()

    setTopMost
    Me.SetFocus

End Sub
