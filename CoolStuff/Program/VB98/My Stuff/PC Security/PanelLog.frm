VERSION 5.00
Begin VB.Form Signon 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "Mo's PC Security"
   ClientHeight    =   2160
   ClientLeft      =   0
   ClientTop       =   75
   ClientWidth     =   9510
   Icon            =   "PanelLog.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "PanelLog.frx":0442
   ScaleHeight     =   2160
   ScaleWidth      =   9510
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox buttup 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   5280
      Picture         =   "PanelLog.frx":5611
      ScaleHeight     =   240
      ScaleWidth      =   1440
      TabIndex        =   8
      Top             =   480
      Visible         =   0   'False
      Width           =   1440
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H0000FF00&
      Height          =   975
      Left            =   2760
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   5
      Text            =   "PanelLog.frx":59FB
      Top             =   960
      Width           =   4215
   End
   Begin VB.PictureBox PicEnd 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   6960
      Picture         =   "PanelLog.frx":5A0F
      ScaleHeight     =   240
      ScaleWidth      =   1440
      TabIndex        =   3
      Top             =   720
      Width           =   1440
      Begin VB.Label exiter 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Exit"
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   0
         TabIndex        =   4
         Top             =   0
         Width           =   1455
      End
   End
   Begin VB.PictureBox PicAbout 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   1320
      Picture         =   "PanelLog.frx":5DC8
      ScaleHeight     =   240
      ScaleWidth      =   1440
      TabIndex        =   1
      Top             =   720
      Width           =   1440
      Begin VB.Label pichelp 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Help And About"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   0
         TabIndex        =   2
         Top             =   0
         Width           =   1455
      End
   End
   Begin VB.PictureBox buttdown 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   3840
      Picture         =   "PanelLog.frx":6181
      ScaleHeight     =   240
      ScaleWidth      =   1440
      TabIndex        =   0
      Top             =   480
      Visible         =   0   'False
      Width           =   1440
   End
   Begin VB.Timer tmrset 
      Interval        =   1
      Left            =   1800
      Top             =   120
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
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
      Left            =   6600
      TabIndex        =   7
      Top             =   1985
      Width           =   2895
   End
   Begin VB.Label Label6 
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
      Left            =   120
      TabIndex        =   6
      Top             =   1985
      Width           =   2055
   End
End
Attribute VB_Name = "Signon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
'Tells the program to load the error handler if it there is a bug
On Error GoTo FileError
    
    'Open's our hidden log file for viewing
    Open "C:\Windows\System\curty.dat" For Input As #1
    'reads it into text1's textbox
    Text1.Text = Input(LOF(1), #1)    'Read File
    'closes it
    Close #1                                'Close File
    'exits this the Form_Load Sub
    Exit Sub
    
'Our error handler
FileError:
    'Make Text1's text notify the user of a problem if the file
    'doesnt exist. This can happen if that file is being used for
    'something else, or the user never signed in on this computer.
    Text1.Text = "There Was A File Error Or File Does Not Exist!!!"
    'resume next
    Resume Next

    'Change pichelp's text color to white
    pichelp.ForeColor = vbWhite
    'and exiter's text color to white.
    exiter.ForeColor = vbWhite
 End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'when the mouse is moved over the form, it changes pichelp's
    'text color to white
    pichelp.ForeColor = vbWhite
    'and exiter's text color to white.
    exiter.ForeColor = vbWhite
End Sub

Private Sub exiter_Click()
    'End the program
    End
    'and close it.
    Close
End Sub

Private Sub exiter_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Change the picture of the button when the mouse is pressed
    ' down
    PicEnd.Picture = buttdown.Picture
End Sub

Private Sub pichelp_Click()
    'Show the About section.
    frmAbout.Show
    'Disable this timer
    tmrset.Enabled = False
    'and enable that one.
    frmAbout.TmrOn.Enabled = True
End Sub

Private Sub pichelp_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Change the button's picture when the mouse button is pressed
    'down
    PicAbout.Picture = buttdown.Picture
End Sub

Private Sub pichelp_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Change the text color of pichelp to yellow if the mosue is
    'moved over it
    pichelp.ForeColor = vbYellow
End Sub

Private Sub pichelp_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Show the other pic for the button if the user releases the
    'mouse button
    PicAbout.Picture = buttup.Picture
End Sub

Private Sub Text1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Change the textcolor of pichelp to white if the user moves
    'mouse over text1
    pichelp.ForeColor = vbWhite
    'and change exiter's text color to white, also.
    exiter.ForeColor = vbWhite
End Sub

Private Sub tmrset_Timer()
    'Show SignOn (the form)
     Signon.Show
End Sub
