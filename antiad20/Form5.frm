VERSION 5.00
Begin VB.Form Form5 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "About \\ AD Blocker !"
   ClientHeight    =   3105
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   7290
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3105
   ScaleWidth      =   7290
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   315
      ScaleHeight     =   825
      ScaleWidth      =   6645
      TabIndex        =   0
      Top             =   165
      Width           =   6675
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "created"
         Height          =   1650
         Left            =   870
         TabIndex        =   1
         Top             =   735
         Width           =   4995
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   3420
      Top             =   1995
   End
   Begin VB.Image Image1 
      Height          =   3000
      Left            =   45
      Picture         =   "Form5.frx":0000
      Top             =   60
      Width           =   7200
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Label2.Caption = " •••• created by •••• " & vbCrLf & "eugene p." & vbCrLf & vbCrLf & "•••• thanks to ••••" & vbCrLf & "Contributors from Planet Source Code !" & vbCrLf & vbCrLf & "Look for more updates soon !" & vbCrLf & "http://www.eugenius.tk"
    
End Sub

Private Sub Timer1_Timer()
Label2.Top = Label2.Top - 5

If Label2.Top < Picture1.Top - Label2.Height Then Label2.Top = 735
End Sub
