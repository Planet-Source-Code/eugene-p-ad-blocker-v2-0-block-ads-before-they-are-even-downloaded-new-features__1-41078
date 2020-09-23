VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form Form1 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "AD Blocker 2.0 "
   ClientHeight    =   1020
   ClientLeft      =   150
   ClientTop       =   795
   ClientWidth     =   4005
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1020
   ScaleWidth      =   4005
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Inject !"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   105
      TabIndex        =   1
      Top             =   135
      Width           =   1110
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   210
      Left            =   1350
      TabIndex        =   0
      Top             =   330
      Width           =   2520
      _ExtentX        =   4445
      _ExtentY        =   370
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "bY: Eugene"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   195
      Left            =   1515
      TabIndex        =   3
      Top             =   750
      Width           =   915
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   120
      X2              =   3870
      Y1              =   660
      Y2              =   660
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Progress:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1350
      TabIndex        =   2
      Top             =   105
      Width           =   795
   End
   Begin VB.Menu mnuPop 
      Caption         =   "X"
      Begin VB.Menu mnuRestore 
         Caption         =   "Display &AD Blocker"
      End
      Begin VB.Menu mnuLine1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuStat 
         Caption         =   "0 AD(s) Blocked !"
      End
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "Options"
      Begin VB.Menu mnuEditFake 
         Caption         =   "Edit FakeAD"
      End
      Begin VB.Menu mnuViewList 
         Caption         =   "View Blocked AD Server(s)"
      End
   End
   Begin VB.Menu mnuStatistics 
      Caption         =   "Statistics"
      Begin VB.Menu mnuStartServer 
         Caption         =   "Start Statistics Server"
      End
      Begin VB.Menu mnuLine2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewStats 
         Caption         =   "View Statistics"
      End
   End
   Begin VB.Menu mnuTray 
      Caption         =   "Send to Tray !"
   End
   Begin VB.Menu mnuAbout 
      Caption         =   "About !"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim nHost As String
Private Sub Command1_Click()
Y = MsgBox("Are you sure you want to INJECT ?? You only need to do this once on each computer !", vbQuestion + vbYesNo, "Are you sure?")
If Y = vbNo Then
    Exit Sub
End If

Dim X As Integer
X = 0
ProgressBar1.Max = Form3.List1.ListCount

If nHost <> "" Then  'IF TEXTBOX2 doesnt have nothing THEN
    Open nHost & "hosts" For Append As #1 'OPEN the HOSTS File For Appending Write
    Print #1, vbCrLf 'Print a Carrige Return
    Print #1, "# AD Blocker List By - Eugene"
    
    Do Until X = Form3.List1.ListCount 'Begin a Do Until Loop
    DoEvents 'Do events
        Print #1, Form3.List1.List(X) 'Append the Server from the Listbox
        X = X + 1 'add 1 to X so next leep we will get onto the NEXT ITEM in the List
        ProgressBar1.Value = X 'increase the Progressbar value by 1
    Loop 'loop (for the new to VB people, it goes back to "Print #1, List1.List(X)" 3 lines up
    Close #1 'Close the file, save the writing
    
    MsgBox "AD Blocker ! Injected into the system, process complete ! Thank you for using AD Blocker !  !", vbInformation, "Complete !"
End If

End Sub

Private Sub Command2_Click()
    Winsock1.Close 'Close Winsock or RESET
    Winsock1.Listen 'Listen for new Connections
End Sub

Private Sub Command3_Click()

End Sub

Private Sub Form_Load()
If FExist(App.Path & "\database.txt") = True Then 'if the database.txt FIle Exists then
    Call Loadlistbox(App.Path & "\database.txt", Form3.List1) 'Load the contents of the file to the listbox1
End If

If Command <> "" Then 'IF Command Syntax Doesnt Equel Something "" then
    If InStr(1, Command, "/startserver") > 0 Then   'IF /startserver found in SYNTAX then
        mnuStartServer_Click
    End If
    If InStr(1, Command, "/sendtotray") > 0 Then    'IF /sendtotray found in SYNTAX  then
        mnuTray_Click
    End If
End If

    mnuPop.Visible = False  'Hide the PopUP Menu for the SySTray
    Form3.Label1.Caption = "AD Server Database: " & Form3.List1.ListCount & " blocked" 'Display the current # ad servers !
    Form2.txtFake = GetSetting("adblock", "conf", "fake", Form2.txtFake) 'Get the saved system setting for the Fake AD on FOrm2 txtFake
    
'///// FIND the HOST FILE
If FExist("c:\windows\system32\drivers\etc\hosts") = True Then
    nHost = "c:\windows\system32\drivers\etc\"
ElseIf FExist("c:\windows\hosts") = True Then
    nHost = "c:\windows\"
ElseIf FExist("c:\winnt\system32\drivers\etc\hosts") = True Then
    nHost = "c:\winnt\system32\drivers\etc\"
End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Dim Message As Long
   On Error Resume Next
    Message = X / Screen.TwipsPerPixelX
    Select Case Message
    
        Case WM_RBUTTONUP    'When RIGHT mouse button is UP
            PopupMenu mnuPop 'Pop UP the hidden MenuPop
        Case WM_RBUTTONDOWN
        
        Case WM_LBUTTONUP
            PopupMenu mnuPop 'Pop UP the hidden MenuPop
        Case WM_LBUTTONDOWN
        
    End Select
End Sub

Private Sub Form_Terminate()
    Call RemoveFromTray  'Remove the tray icon if its still alive
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call RemoveFromTray 'Remove the tray icon if its still alive
    Unload Form5
    Unload Form4
    Unload Form3
    Unload Form2
End Sub

Private Sub mnuAbout_Click()
    Form5.Show
End Sub

Private Sub mnuEditFake_Click()
    Form2.Show
End Sub

Private Sub mnuRestore_Click()
    Call RemoveFromTray   'remove tray icon
    Me.Show               'show me !
End Sub

Private Sub mnuStartServer_Click()
'///// INIT STATISTICS SERVER
On Error GoTo Problem
    Form4.Winsock1.Close 'Close the COnnection
    Form4.Winsock1.Listen '(LISTEN) Wait for another connection
    Exit Sub

Problem:
    MsgBox "Failed to LOAD WinSock on PORT 80, the STATISTICS SERVER WILL BE OFFLINE !!" & vbCrLf & vbCrLf & "How To FIX: Check and make sure nothing is running on Port 80 !", vbCritical, "Error !"
End Sub

Private Sub mnuTray_Click()
    Me.Visible = False 'hide me
    Call AddToTray(Me, "[eugenius] AD Blocker !", Me.Icon) 'add to system tray
End Sub

Private Sub mnuViewList_Click()
    Form3.Show
End Sub

Private Sub mnuViewStats_Click()
    Form4.Show
End Sub
