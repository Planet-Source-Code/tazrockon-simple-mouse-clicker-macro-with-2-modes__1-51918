VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form3 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mad Taz Clicker v1.1"
   ClientHeight    =   3660
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   4605
   Icon            =   "Form1_adv.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3660
   ScaleWidth      =   4605
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      Caption         =   "Click Interval"
      Height          =   1095
      Left            =   0
      TabIndex        =   26
      Top             =   2520
      Width           =   4575
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   3480
         Locked          =   -1  'True
         MaxLength       =   1
         TabIndex        =   29
         Text            =   "1"
         Top             =   720
         Width           =   255
      End
      Begin MSComctlLib.Slider Slider1 
         Height          =   495
         Left            =   480
         TabIndex        =   27
         Top             =   240
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   873
         _Version        =   393216
         Min             =   1
         Max             =   6
         SelStart        =   1
         Value           =   1
      End
      Begin VB.Label Label16 
         Caption         =   "seconds"
         Height          =   255
         Left            =   3720
         TabIndex        =   30
         Top             =   720
         Width           =   615
      End
      Begin VB.Label Label15 
         Caption         =   "Click every"
         Height          =   255
         Left            =   2640
         TabIndex        =   28
         Top             =   720
         Width           =   855
      End
   End
   Begin VB.Timer Timer5 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   2520
      Top             =   3720
   End
   Begin VB.Timer Timer4 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   1920
      Top             =   3720
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   1320
      Top             =   3720
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   720
      Top             =   3720
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   120
      Top             =   3720
   End
   Begin VB.Frame Frame2 
      Height          =   2535
      Left            =   2640
      TabIndex        =   1
      Top             =   0
      Width           =   1935
      Begin VB.CommandButton Command6 
         Caption         =   "About"
         Height          =   375
         Left            =   120
         TabIndex        =   9
         Top             =   1680
         Width           =   1695
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Help"
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Top             =   1320
         Width           =   1695
      End
      Begin VB.CommandButton Command4 
         Caption         =   "&End"
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   960
         Width           =   1695
      End
      Begin VB.CommandButton Command3 
         Caption         =   "&Begin"
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   600
         Width           =   1695
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000000&
      Height          =   2535
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2655
      Begin VB.CommandButton Command8 
         Caption         =   "&Clear"
         Height          =   315
         Left            =   1320
         TabIndex        =   25
         Top             =   2160
         Width           =   1095
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Unlock"
         Height          =   315
         Left            =   1320
         TabIndex        =   24
         Top             =   1800
         Width           =   1095
      End
      Begin VB.CommandButton Command7 
         Caption         =   "&Set"
         Height          =   315
         Left            =   240
         TabIndex        =   23
         Top             =   2160
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Lock"
         Height          =   315
         Left            =   240
         TabIndex        =   22
         Top             =   1800
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
      Begin VB.Label Label14 
         Height          =   255
         Left            =   1800
         TabIndex        =   21
         Top             =   1560
         Width           =   615
      End
      Begin VB.Label Label13 
         Height          =   255
         Left            =   960
         TabIndex        =   20
         Top             =   1560
         Width           =   615
      End
      Begin VB.Label Label12 
         Height          =   255
         Left            =   1800
         TabIndex        =   19
         Top             =   1320
         Width           =   615
      End
      Begin VB.Label Label11 
         Height          =   255
         Left            =   960
         TabIndex        =   18
         Top             =   1320
         Width           =   615
      End
      Begin VB.Label Label10 
         Height          =   255
         Left            =   1800
         TabIndex        =   17
         Top             =   1080
         Width           =   615
      End
      Begin VB.Label Label9 
         Height          =   255
         Left            =   960
         TabIndex        =   16
         Top             =   1080
         Width           =   615
      End
      Begin VB.Label Label8 
         Height          =   255
         Left            =   1800
         TabIndex        =   15
         Top             =   840
         Width           =   615
      End
      Begin VB.Label Label7 
         Height          =   255
         Left            =   960
         TabIndex        =   14
         Top             =   840
         Width           =   615
      End
      Begin VB.Label Label6 
         Caption         =   "Click 4:"
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   1560
         Width           =   615
      End
      Begin VB.Label Label5 
         Caption         =   "Click 3:"
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   1320
         Width           =   615
      End
      Begin VB.Label Label4 
         Caption         =   "Click 2:"
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   1080
         Width           =   615
      End
      Begin VB.Label Label3 
         Caption         =   "Click 1:"
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   840
         Width           =   615
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
Attribute VB_Name = "Form3"
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
Timer2.Interval = Text3.Text * 1000
Timer3.Interval = Text3.Text * 1000
Timer4.Interval = Text3.Text * 1000
Timer5.Interval = Text3.Text * 1000
Timer2.Enabled = True
End Sub

Private Sub Command4_Click()
Timer2.Enabled = False
Timer3.Enabled = False
Timer4.Enabled = False
Timer5.Enabled = False
End Sub

Private Sub Command5_Click()
Form2.Visible = True
End Sub

Private Sub Command6_Click()
MsgBox "This program was created by Tazrockon in June 2003. The creator can not and will not be held responsible for any problems or damages that may occur on your computer before, during, or after this proram has been run. You may not buy or sell this program for any amount of money without the owners permission. You may not edit or change this program or its' files in any way.", vbInformation + vbOKOnly, "Mad Taz Clicker - About"
End Sub

Private Sub Command7_Click()
If Label7.Caption = "" Then
    Label7.Caption = Text1.Text
    Label8.Caption = Text2.Text
ElseIf Label7.Caption <> "" And Label9.Caption = "" Then
    Label9.Caption = Text1.Text
    Label10.Caption = Text2.Text
ElseIf Label7.Caption <> "" And Label9.Caption <> "" And Label11.Caption = "" Then
    Label11.Caption = Text1.Text
    Label12.Caption = Text2.Text
ElseIf Label7.Caption <> "" And Label9.Caption <> "" And Label11.Caption <> "" And Label13.Caption = "" Then
    Label13.Caption = Text1.Text
    Label14.Caption = Text2.Text
ElseIf Label7.Caption <> "" And Label9.Caption <> "" And Label11.Caption <> "" And Label13.Caption <> "" Then
    MsgBox "All of the click position spaces are full. You must clear them if you want to add that click.", vbCritical + vbOKOnly, "Spaces Full"
End If
End Sub

Private Sub Command8_Click()
Label7.Caption = ""
Label8.Caption = ""
Label9.Caption = ""
Label10.Caption = ""
Label11.Caption = ""
Label12.Caption = ""
Label13.Caption = ""
Label14.Caption = ""
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

Private Sub Slider1_Click()
Text3.Text = Slider1.Value
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
If Label7.Caption = "" Then
    MsgBox "You must first set the coordinates to be clicked.", vbExclamation + vbOKOnly, "Set Coords!"
    Exit Sub
ElseIf Label7.Caption <> "" Then
    xP = Label7.Caption
    yP = Label8.Caption
    MouseMove (xP), (yP)
    LeftClick (xP), (yP)
    Timer3.Enabled = True
    Timer2.Enabled = False
End If
End Sub

Private Sub Timer3_Timer()
Dim xP As Long
Dim yP As Long
If Label9.Caption = "" Then
    Timer2.Enabled = True
    Timer3.Enabled = False
ElseIf Label7.Caption <> "" Then
    xP = Label9.Caption
    yP = Label10.Caption
    MouseMove (xP), (yP)
    LeftClick (xP), (yP)
    Timer4.Enabled = True
    Timer3.Enabled = False
End If
End Sub

Private Sub Timer4_Timer()
Dim xP As Long
Dim yP As Long
If Label11.Caption = "" Then
    Timer2.Enabled = True
    Timer4.Enabled = False
ElseIf Label7.Caption <> "" Then
    xP = Label11.Caption
    yP = Label12.Caption
    MouseMove (xP), (yP)
    LeftClick (xP), (yP)
    Timer5.Enabled = True
    Timer4.Enabled = False
End If
End Sub

Private Sub Timer5_Timer()
Dim xP As Long
Dim yP As Long
If Label11.Caption = "" Then
    Timer2.Enabled = True
    Timer5.Enabled = False
ElseIf Label7.Caption <> "" Then
    xP = Label13.Caption
    yP = Label14.Caption
    MouseMove (xP), (yP)
    LeftClick (xP), (yP)
    Timer2.Enabled = True
    Timer5.Enabled = False
End If
End Sub
