VERSION 5.00
Begin VB.Form Signon 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "Mo's PC Security"
   ClientHeight    =   2160
   ClientLeft      =   0
   ClientTop       =   75
   ClientWidth     =   9510
   Icon            =   "SignIn.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2160
   ScaleWidth      =   9510
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox PicEnd 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   6600
      ScaleHeight     =   240
      ScaleWidth      =   1440
      TabIndex        =   12
      Top             =   1560
      Width           =   1440
      Begin VB.Label exiter 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Exit"
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   0
         TabIndex        =   13
         Top             =   0
         Width           =   1455
      End
   End
   Begin VB.PictureBox PicAbout 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   4200
      ScaleHeight     =   240
      ScaleWidth      =   1440
      TabIndex        =   10
      Top             =   1560
      Width           =   1440
      Begin VB.Label pichelp 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Help And About"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   0
         TabIndex        =   11
         Top             =   0
         Width           =   1455
      End
   End
   Begin VB.PictureBox buttup 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   2520
      Picture         =   "SignIn.frx":0442
      ScaleHeight     =   240
      ScaleWidth      =   1440
      TabIndex        =   9
      Top             =   480
      Visible         =   0   'False
      Width           =   1440
   End
   Begin VB.PictureBox buttdown 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   2520
      Picture         =   "SignIn.frx":082C
      ScaleHeight     =   240
      ScaleWidth      =   1440
      TabIndex        =   8
      Top             =   240
      Visible         =   0   'False
      Width           =   1440
   End
   Begin VB.PictureBox picSignIn 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   1680
      ScaleHeight     =   240
      ScaleWidth      =   1440
      TabIndex        =   6
      Top             =   1560
      Width           =   1440
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Sign - In"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   0
         TabIndex        =   7
         Top             =   0
         Width           =   1455
      End
   End
   Begin VB.Timer tmrset 
      Interval        =   1
      Left            =   1800
      Top             =   120
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      ForeColor       =   &H000000FF&
      Height          =   285
      Left            =   1560
      TabIndex        =   0
      Text            =   "Enter Name Here"
      Top             =   1200
      Width           =   1695
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "http://Dictator.50Megs.Com"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   7080
      TabIndex        =   15
      Top             =   1985
      Width           =   2415
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Dictator@MyPlace.com"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   0
      TabIndex        =   14
      Top             =   1985
      Width           =   2415
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Date:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   4080
      TabIndex        =   5
      Top             =   960
      Width           =   1575
   End
   Begin VB.Label lblDate 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00C000C0&
      Height          =   255
      Left            =   4080
      TabIndex        =   4
      Top             =   1200
      Width           =   1695
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Sign-In Name:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1560
      TabIndex        =   3
      Top             =   960
      Width           =   1695
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Time:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   6600
      TabIndex        =   2
      Top             =   960
      Width           =   1575
   End
   Begin VB.Label lblDisplay 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   6480
      TabIndex        =   1
      Top             =   1200
      Width           =   1695
   End
End
Attribute VB_Name = "Signon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub exiter_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Change exiter's forecolor to Yellow when the user moves the
    'mouse over it
    exiter.ForeColor = vbYellow
End Sub

Private Sub exiter_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Change the picture when the user releases the mouse button.
    PicEnd.Picture = buttup.Picture
End Sub

Private Sub Form_Load()
    'Set the stuff for loading...
    'Make sure of the following:
    
    'Label3's text color is white
    Label3.ForeColor = vbWhite
    'pichelp's text color is white
    pichelp.ForeColor = vbWhite
    'exiter's text color is white
    exiter.ForeColor = vbWhite
    'signin's pic is buttup (which is the button that looks
    'normal and not pushed down)
    picSignIn.Picture = buttup.Picture
    'picabout's pic is buttup
    PicAbout.Picture = buttup.Picture
    'picend's pic is buttup
    PicEnd.Picture = buttup.Picture
    
    'Load the background picture from the Data/ file.
    Me.Picture = LoadPicture(App.Path & "/Data/MyBack.jpg")
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Change label3's forecolor to white,
    Label3.ForeColor = vbWhite
    'pichelp's to white,
    pichelp.ForeColor = vbWhite
    'and exiter, also.
    exiter.ForeColor = vbWhite
End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Change Label3's text color to white,
    Label3.ForeColor = vbWhite
    'pichelp's text color to white,
    pichelp.ForeColor = vbWhite
    'and exiter's text color to white.
    exiter.ForeColor = vbWhite
End Sub

Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Change Label3's text color to white,
    Label3.ForeColor = vbWhite
    'and pichelp's
    pichelp.ForeColor = vbWhite
    'and exiter's
    exiter.ForeColor = vbWhite
End Sub

Private Sub Label3_Click()
    'Error handler for incorrect accounts
   If Text1.Text = "" Or Text1.Text = " " Or Text1.Text = "Enter Name Here" Then
    'messages the user to fix it
    MsgBox "You must fill in valid information!"
    'exits the sub
    Exit Sub
   'if everything is alright, then
   Else
    'open curty.dat for some editing
    Open "C:\Windows\System\curty.dat" For Append As #1
    'print the line of dashes to separate from the previous log
    Print #1, "- - - - - - - - - - - - - - - - - -"
    'print the users name, date and time when logged on
    Print #1, Text1.Text; " is logged in at:"; Date; ", "; Time; " ."
    'print a dash
    Print #1, "-"
    'print name
    Print #1, Text1.Text
    'print date
    Print #1, Date
    'print time
    Print #1, Time
    'close the file
    Close #1
    'end the program
    End
    'close the program
    Close
   'end of the IF statement
   End If
End Sub

Private Sub Label3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Change the button's pic to the button that is pushed down
    'when the user holds down the mouse button.
    picSignIn.Picture = buttdown.Picture
End Sub

Private Sub Label3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Change the text color of Label3 to yellow when the user
    'moves the mouse over it
    Label3.ForeColor = vbYellow
End Sub

Private Sub Label3_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Change back the pic when the user releases the mouse button
    picSignIn.Picture = buttup.Picture
End Sub

Private Sub exiter_Click()
    'End the program
    End
    'Close it
    Close
End Sub

Private Sub exiter_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Change the pic when the user holds down the mouse button
    PicEnd.Picture = buttdown.Picture
End Sub

Private Sub Label4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Change the text color of Label3 to white when the mouse is
    'moved over label4
    Label3.ForeColor = vbWhite
    'and change pichelp's text color to white
    pichelp.ForeColor = vbWhite
    'and exiter's text color to white, also.
    exiter.ForeColor = vbWhite
End Sub

Private Sub pichelp_Click()
    'Show the About form
    frmAbout.Show
    'de-activate the timer on this form
    tmrset.Enabled = False
    'activate the timer on the About form
    frmAbout.TmrOn.Enabled = True
End Sub

Private Sub pichelp_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Change the picture of the button when the user holds down
    'the mouse button
    PicAbout.Picture = buttdown.Picture
End Sub

Private Sub pichelp_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'When the mouse is moved over the button, change the text to
    'yellow
    pichelp.ForeColor = vbYellow
End Sub

Private Sub pichelp_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'When the mouse button is released, change the pic back
    PicAbout.Picture = buttup.Picture
End Sub

Private Sub Text1_Change()
 'if the user enters a space then change the text to blank.
 If Text1.Text = " " Then Text1.Text = ""
 'if the user types 'enter name here',then empty the text box.
 If Text1.Text = "Enter Name Here" Then Text1.Text = ""
 'if the user presses backspace when Enter Name Here is in the
 'sign in box, then empty the text box.
 If Text1.Text = "Enter Name Her" Then Text1.Text = ""
 'if the user presses a space, then empty the text box.
 If Text1.Text = "Enter Name Here " Then Text1.Text = ""
End Sub

Private Sub Text1_Click()
 'make text1's text empty if the user clicks the text box
 Text1.Text = ""
End Sub

Private Sub Text1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Change the text color of label3 to white when the mouse is
    'moved over Text1
    Label3.ForeColor = vbWhite
    'and pichelp's text color to white
    pichelp.ForeColor = vbWhite
    'and exiter's text color to white, also.
    exiter.ForeColor = vbWhite
End Sub

Private Sub tmrset_Timer()
    'Keep showing this form every millisecond
     Signon.Show
     'make lblDisplay show the time
     lblDisplay.Caption = Time
     'and lblDate show the date
     lblDate.Caption = Date
End Sub
