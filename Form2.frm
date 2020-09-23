VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mad Taz Clicker - Help"
   ClientHeight    =   1845
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5640
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1845
   ScaleWidth      =   5640
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text2 
      Height          =   1215
      Left            =   2400
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   480
      Width           =   3135
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   2400
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   120
      Width           =   1935
   End
   Begin VB.ListBox List1 
      Height          =   1620
      Left            =   120
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   120
      Width           =   2175
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
List1.AddItem "Locking Mouse Position"
List1.AddItem "Unlocking Mouse Position"
List1.AddItem "Begin Clicking"
List1.AddItem "Stop Clicking"
List1.AddItem "X and Y Position"
List1.AddItem "Purpose of Program"
List1.AddItem "Contact"
List1.AddItem "Standard Mode"
List1.AddItem "4x Mode"
List1.AddItem "Set"
List1.AddItem "Clear"
End Sub

Private Sub List1_Click()
If List1.Text = "Locking Mouse Position" Then
    Text1.Text = "Locking Mouse Position"
    Text2.Text = "Before Mad Taz Clicker can start clicking on the coordinates you desire, you have to lock the mouse coordinates of the point you want. To do this, position you mouse over the point you want to be clicked and press Alt and L at the same time. This should lock the current mouse coordinate. You know you have done this successfully when you can move the mouse around without the positions shown in the text boxes changing."
ElseIf List1.Text = "Unlocking Mouse Position" Then
    Text1.Text = "Unlocking Mouse Position"
    Text2.Text = "After you have locked coordinates in the text boxes you can click either Unlock or press Alt and U on your keyboard if you wish to change them."
ElseIf List1.Text = "Begin Clicking" Then
    Text1.Text = "Begin Clicking"
    Text2.Text = "After you have locked coordinates in the text boxes you can have Mad Taz Clicker Standard Mode start clicking on them by clicking the Begin button or by pressing Atl and B on your keyboard. If using 4x mode you must first set the coords before clicking begin. This will make Mad Taz Clicker click on the locked coordinates once per every two seconds in standard mode or once every interval in 4x mode until you tell it to stop."
ElseIf List1.Text = "Stop Clicking" Then
    Text1.Text = "Stop Clicking"
    Text2.Text = "Once Mad Taz Clicker has began clicking on the points you locked, you can stop it by clicking Stop or by pressing Alt and S on your keyboard."
ElseIf List1.Text = "X and Y Position" Then
    Text1.Text = "X and Y Position"
    Text2.Text = "For every point on the screen that you move your mouse there is an X and Y Position, also known as the mouse's coordinates. You can view your mouse's coordinates in the X and Y position text boxes."
ElseIf List1.Text = "Purpose of Program" Then
    Text1.Text = "Purpose of Program"
    Text2.Text = "The purpose of Mad Taz Clicker is to aid in the monotonous task of clicking in the same spot repeatedly, required by some computer programs. This program replicates the user's job of clicking on that spot."
ElseIf List1.Text = "Contact" Then
    Text1.Text = "Contact"
    Text2.Text = "If you have any questions, comments, or any problems with this program then send an email to tazrockon@msn.com. In the subject put Mad Taz Clicker and in the body of the message be sure to clearly and quickly say what you want to."
ElseIf List1.Text = "Standard Mode" Then
    Text1.Text = "Standard Mode"
    Text2.Text = "When you run Mad Taz Clicker, the mode you will first see it in is Standard Mode. This mode allows you to click one spot many times."
ElseIf List1.Text = "4x Mode" Then
    Text1.Text = "4x Mode"
    Text2.Text = "If you want to use 4x Mode click file and under mode click 4x Mode. This mode allows you to set up to four spots for Mad Taz Clicker to click. In this mode you can also set the mouse to click anywhere from every one second to every six seconds."
ElseIf List1.Text = "Set" Then
    Text1.Text = "Set"
    Text2.Text = "In 4x mode, after you have loacked the coordinate that you want to be clicked, you have to press set to set it. You can set up to four different coords and then have Mad Taz Clicker click them all."
ElseIf List1.Text = "Clear" Then
    Text1.Text = "Clear"
    Text2.Text = "In 4x Mode, if you want to clear all of the stored coordinates press the Clear button."
End If
End Sub
