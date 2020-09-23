VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mad Taz Clicker v1.1"
   ClientHeight    =   1350
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   4605
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1350
   ScaleWidth      =   4605
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   720
      Top             =   1440
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   120
      Top             =   1440
   End
   Begin VB.Frame Frame2 
      Height          =   1335
      Left            =   2640
      TabIndex        =   1
      Top             =   0
      Width           =   1935
      Begin VB.CommandButton Command6 
         Caption         =   "About"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   960
         Width           =   1695
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Help"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   720
         Width           =   1695
      End
      Begin VB.CommandButton Command4 
         Caption         =   "&Stop"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   480
         Width           =   1695
      End
      Begin VB.CommandButton Command3 
         Caption         =   "&Begin"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1335
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2655
      Begin VB.CommandButton Command2 
         Caption         =   "&Unlock"
         Height          =   315
         Left            =   1320
         TabIndex        =   7
         Top             =   960
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Lock"
         Height          =   315
         Left            =   240
         TabIndex        =   6
         Top             =   960
         Width           =   1095
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   1560
         TabIndex        =   5
         Top             =   480
         Width           =   855
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   240
         TabIndex        =   4
         Top             =   480
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "Y Pos:"
         Height          =   255
         Left            =   1560
         TabIndex        =   3
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "X Pos:"
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuMode 
         Caption         =   "Mode"
         Begin VB.Menu mnuModeStandard 
            Caption         =   "Standard"
         End
         Begin VB.Menu mnuMode4x 
            Caption         =   "4x Mode"
         End
      End
      Begin VB.Menu mnuHelp 
         Caption         =   "Help"
      End
      Begin VB.Menu mnuDash 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Timer1.Enabled = False
End Sub

Private Sub Command2_Click()
Timer1.Enabled = True
End Sub

Private Sub Command3_Click()
Timer2.Enabled = True
End Sub

Private Sub Command4_Click()
Timer2.Enabled = False
End Sub

Private Sub Command5_Click()
Form2.Visible = True
End Sub

Private Sub Command6_Click()
MsgBox "This program was created by Tazrockon in June 2003. The creator can not and will not be held responsible for any problems or damages that may occur on your computer before, during, or after this proram has been run. You may not buy or sell this program for any amount of money without the owners permission. You may not edit or change this program or its' files in any way.", vbInformation + vbOKOnly, "Mad Taz Clicker - About"
End Sub

Private Sub mnuExit_Click()
End
End Sub

Private Sub mnuHelp_Click()
Form2.Visible = True
End Sub

Private Sub mnuMode4x_Click()
Form3.Visible = True
Form1.Visible = False
End Sub

Private Sub mnuModeStandard_Click()
Form1.Visible = True
Form3.Visible = False
End Sub

Private Sub Timer1_Timer()
Dim pos
Dim pt As PointAPI
pos = GetCursorPos(pt)
Text1.Text = pt.x
Text2.Text = pt.y
End Sub

Private Sub Timer2_Timer()
Dim xP As Long
Dim yP As Long
Dim move As Long
xP = Text1.Text
yP = Text2.Text
MouseMove (xP), (yP)
LeftClick (xP), (yP)
End Sub
