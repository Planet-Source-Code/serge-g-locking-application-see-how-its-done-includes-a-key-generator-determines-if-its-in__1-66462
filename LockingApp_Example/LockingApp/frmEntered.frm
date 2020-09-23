VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmEntered 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Untitled 1"
   ClientHeight    =   6450
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   7590
   Icon            =   "frmEntered.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6450
   ScaleWidth      =   7590
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000009&
      BorderStyle     =   0  'None
      Height          =   2580
      Left            =   3285
      ScaleHeight     =   2580
      ScaleWidth      =   1425
      TabIndex        =   6
      Top             =   -15
      Visible         =   0   'False
      Width           =   1425
      Begin VB.ListBox List1 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000A&
         Height          =   2565
         Left            =   0
         TabIndex        =   7
         Top             =   15
         Width           =   1410
      End
   End
   Begin VB.PictureBox pic1 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000A&
      ForeColor       =   &H80000008&
      Height          =   2490
      Left            =   15
      ScaleHeight     =   2460
      ScaleWidth      =   2745
      TabIndex        =   3
      Top             =   480
      Visible         =   0   'False
      Width           =   2775
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   2385
         Left            =   15
         MultiLine       =   -1  'True
         TabIndex        =   5
         Text            =   "frmEntered.frx":014A
         Top             =   -30
         Visible         =   0   'False
         Width           =   2670
      End
      Begin MSComCtl2.MonthView cal 
         Height          =   2370
         Left            =   0
         TabIndex        =   4
         Top             =   30
         Visible         =   0   'False
         Width           =   2700
         _ExtentX        =   4763
         _ExtentY        =   4180
         _Version        =   393216
         ForeColor       =   -2147483630
         BackColor       =   -2147483633
         Appearance      =   1
         StartOfWeek     =   22806529
         CurrentDate     =   38919
      End
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   25
      Left            =   3360
      Top             =   3240
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   5805
      Left            =   -15
      TabIndex        =   2
      Top             =   435
      Width           =   7590
      _ExtentX        =   13388
      _ExtentY        =   10239
      _Version        =   393217
      TextRTF         =   $"frmEntered.frx":0290
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   225
      Left            =   0
      TabIndex        =   1
      Top             =   6225
      Width           =   7590
      _ExtentX        =   13388
      _ExtentY        =   397
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   7
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            Enabled         =   0   'False
            Object.Width           =   882
            MinWidth        =   882
            TextSave        =   "CAPS"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            Object.Width           =   882
            MinWidth        =   882
            TextSave        =   "NUM"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   3
            Alignment       =   1
            AutoSize        =   1
            Enabled         =   0   'False
            Object.Width           =   2258
            MinWidth        =   882
            TextSave        =   "INS"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   4
            Alignment       =   1
            AutoSize        =   1
            Enabled         =   0   'False
            Object.Width           =   2258
            MinWidth        =   882
            TextSave        =   "SCRL"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            AutoSize        =   1
            Object.Width           =   2258
            MinWidth        =   882
            TextSave        =   "8:55 PM"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            AutoSize        =   1
            Object.Width           =   2434
            MinWidth        =   1058
            TextSave        =   "9/4/2006"
         EndProperty
         BeginProperty Panel7 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   7
            Alignment       =   1
            AutoSize        =   1
            Enabled         =   0   'False
            Object.Width           =   2258
            MinWidth        =   882
            TextSave        =   "KANA"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   3450
      Top             =   2940
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   46
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEntered.frx":0312
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEntered.frx":046C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEntered.frx":0A06
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEntered.frx":0FA0
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEntered.frx":10FA
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEntered.frx":1694
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEntered.frx":17EE
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEntered.frx":1948
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEntered.frx":1AA2
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEntered.frx":1BFC
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEntered.frx":2196
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEntered.frx":2730
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEntered.frx":288A
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEntered.frx":2E24
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEntered.frx":2F7E
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEntered.frx":30D8
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEntered.frx":3232
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEntered.frx":338C
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEntered.frx":34E6
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEntered.frx":3A80
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEntered.frx":401A
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEntered.frx":45B4
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEntered.frx":470E
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEntered.frx":4868
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEntered.frx":49C2
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEntered.frx":4B1C
            Key             =   ""
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEntered.frx":50B6
            Key             =   ""
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEntered.frx":5650
            Key             =   ""
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEntered.frx":57AA
            Key             =   ""
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEntered.frx":5D44
            Key             =   ""
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEntered.frx":5E9E
            Key             =   ""
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEntered.frx":5F80
            Key             =   ""
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEntered.frx":60DA
            Key             =   ""
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEntered.frx":6234
            Key             =   ""
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEntered.frx":638E
            Key             =   ""
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEntered.frx":6928
            Key             =   ""
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEntered.frx":7202
            Key             =   ""
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEntered.frx":735C
            Key             =   ""
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEntered.frx":7C36
            Key             =   ""
         EndProperty
         BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEntered.frx":8510
            Key             =   ""
         EndProperty
         BeginProperty ListImage41 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEntered.frx":8DEA
            Key             =   ""
         EndProperty
         BeginProperty ListImage42 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEntered.frx":96C4
            Key             =   ""
         EndProperty
         BeginProperty ListImage43 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEntered.frx":9F9E
            Key             =   ""
         EndProperty
         BeginProperty ListImage44 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEntered.frx":A878
            Key             =   ""
         EndProperty
         BeginProperty ListImage45 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEntered.frx":A9D2
            Key             =   ""
         EndProperty
         BeginProperty ListImage46 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEntered.frx":AB2C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7590
      _ExtentX        =   13388
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   22
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Open a blank Document"
            Object.Tag             =   "blank"
            ImageIndex      =   39
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Tag             =   "end"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Tag             =   "txtUp"
            ImageIndex      =   3
            Style           =   1
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Tag             =   "calc"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Tag             =   "calUp"
            ImageIndex      =   6
            Style           =   1
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   7
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   8
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   9
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   10
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   11
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   12
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   13
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   14
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   15
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   16
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   17
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   18
         EndProperty
         BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   19
         EndProperty
         BeginProperty Button20 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   20
         EndProperty
         BeginProperty Button21 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   21
         EndProperty
         BeginProperty Button22 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   22
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu sf1 
         Caption         =   "Submenu1"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "E&dit"
      Begin VB.Menu se1 
         Caption         =   "Submenu1"
      End
   End
   Begin VB.Menu mnuOpts 
      Caption         =   "&Options"
      Begin VB.Menu mnuReset 
         Caption         =   "Reset Registration"
      End
   End
   Begin VB.Menu mnuEditor 
      Caption         =   "&Editor"
      Begin VB.Menu ssm1 
         Caption         =   "Submenu1"
      End
   End
   Begin VB.Menu mnuWindows 
      Caption         =   "&Windows"
      Begin VB.Menu sw1 
         Caption         =   "Submenu1"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu sh1 
         Caption         =   "Submenu1"
      End
   End
   Begin VB.Menu mnuCustom 
      Caption         =   "Custom Menu"
   End
End
Attribute VB_Name = "frmEntered"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SendMessage Lib "user32" Alias _
        "SendMessageA" (ByVal hwnd As Long, ByVal wMsg _
        As Long, ByVal wParam As Long, lParam As Any) As Long
 
Const LB_ITEMFROMPOINT = &H1A9


Private Sub Form_Click()

    Picture1.Visible = False

End Sub

Private Sub Form_Load()
    
    Me.Show
    CheckLockedStatus
    If Unlocked = False Then
        frmMain.Show vbModeless, Me
'    Else
'        Unload frmMain
    End If
    
    For i = 1 To 20
        List1.AddItem "Menu " & i
    Next i

End Sub

Private Sub Form_Unload(Cancel As Integer)

    End

End Sub

Private Sub List1_DblClick()
    
    Picture1.Visible = False
    MsgBox ("You selected : " & List1.Text)
    
End Sub

Private Sub List1_LostFocus()

    Picture1.Visible = False

End Sub

Private Sub mnuCustom_Click()

    Picture1.Visible = Not Picture1.Visible
    If Picture1.Visible Then
        List1.SetFocus
    End If

End Sub

Private Sub mnuEdit_Click()

    Picture1.Visible = False

End Sub

Private Sub mnuEditor_Click()

    Picture1.Visible = False

End Sub

Private Sub mnuFile_Click()

    Picture1.Visible = False

End Sub

Private Sub mnuHelp_Click()

    Picture1.Visible = False

End Sub

Private Sub mnuOpts_Click()

    Picture1.Visible = False

End Sub

Private Sub mnuReset_Click()

    Picture1.Visible = False
    frmMain.Show

End Sub

Private Sub mnuWindows_Click()

    Picture1.Visible = False

End Sub

Private Sub RichTextBox1_Change()

    Picture1.Visible = False

End Sub

Private Sub RichTextBox1_Click()

    Picture1.Visible = False

End Sub

Private Sub StatusBar1_PanelClick(ByVal Panel As MSComctlLib.Panel)

    Picture1.Visible = False

End Sub

Private Sub Timer1_Timer()

    frmMain.Hide
    frmMain.Show vbModeless, Me
    Timer1.Enabled = False

End Sub



Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    
    Picture1.Visible = False
    
    Select Case Button.Tag
    Case "end"
        End
    Case "calUp"
        Button.Tag = "calDown"
        pic1.Visible = True
        pic1.Height = cal.Height
        pic1.Width = cal.Width
        cal.Visible = True
    Case "calDown"
        cal.Visible = False
        pic1.Visible = False
        Button.Tag = "calUp"
    Case "calc"
        Button.Tag = "calcDown"
    Case "calcDown"
        Button.Tag = "calc"
    Case "txtUp"
        pic1.Height = Text1.Height
        pic1.Width = Text1.Width
        pic1.Visible = True
        Text1.Visible = True
        Button.Tag = "txtDown"
    Case "txtDown"
        pic1.Visible = False
        Text1.Visible = False
        Button.Tag = "txtUp"
    End Select

End Sub

Private Sub List1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

    List1.SetFocus

    Dim lX As Long
    Dim lY As Long
    Dim lIdx As Long

    lX = CLng(x / Screen.TwipsPerPixelX)
    lY = CLng(y / Screen.TwipsPerPixelY)
    
    With List1
        'get selected item from list using an API
        'call to tell you which listindex a pixel falls
        'within
        lIdx = SendMessage(.hwnd, _
          LB_ITEMFROMPOINT, _
          0, _
          ByVal ((lY * 65536) + lX))
        ' show tip or clear last one
        If (lIdx >= 0) And (lIdx <= .ListCount) Then
            .ToolTipText = .List(lIdx)
            .Text = .List(lIdx)
        Else
            .ToolTipText = ""
        End If
    End With

End Sub


