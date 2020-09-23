VERSION 5.00
Begin VB.Form frmAbout 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "About Me"
   ClientHeight    =   5880
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7440
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5880
   ScaleWidth      =   7440
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   6000
      Left            =   2640
      ScaleHeight     =   6000
      ScaleWidth      =   4800
      TabIndex        =   2
      Top             =   0
      Width           =   4800
      Begin VB.PictureBox Picback 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   240
         Left            =   1560
         ScaleHeight     =   240
         ScaleWidth      =   1440
         TabIndex        =   8
         Top             =   5420
         Width           =   1440
         Begin VB.Label GoBack 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Back To Sign-In"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   0
            TabIndex        =   10
            Top             =   0
            Width           =   1455
         End
      End
      Begin VB.PictureBox PicQuick 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   240
         Left            =   1560
         ScaleHeight     =   240
         ScaleWidth      =   1440
         TabIndex        =   7
         Top             =   5060
         Width           =   1440
         Begin VB.Label GetHelp 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Quick Help"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   0
            TabIndex        =   9
            Top             =   0
            Width           =   1455
         End
      End
      Begin VB.Timer TmrOn 
         Enabled         =   0   'False
         Interval        =   1
         Left            =   120
         Top             =   1560
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Dictator@MyPlace.com      Http://Dictator.50Megs.Com"
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
         TabIndex        =   12
         Top             =   0
         Width           =   4815
      End
      Begin VB.Label lblweb 
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Http://Dictator.50Megs.Com"
         ForeColor       =   &H00FFFF80&
         Height          =   255
         Left            =   1320
         TabIndex        =   6
         Top             =   4560
         Width           =   2145
      End
      Begin VB.Label newgen 
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "The New Generation Of PC Security"
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   1080
         TabIndex        =   5
         Top             =   4200
         Width           =   2655
      End
      Begin VB.Label lblCopyright 
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   $"frmAbout.frx":0000
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   2295
         Left            =   720
         TabIndex        =   4
         Top             =   1440
         Width           =   3315
      End
      Begin VB.Label lbAboutMe 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "A Dictator Interactive Production"
         BeginProperty Font 
            Name            =   "Glowworm"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   735
         Left            =   360
         TabIndex        =   3
         Top             =   600
         Width           =   3975
      End
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   6000
      Left            =   0
      ScaleHeight     =   6000
      ScaleWidth      =   2505
      TabIndex        =   0
      Top             =   -120
      Width           =   2505
      Begin VB.PictureBox Picture3 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   135
         Left            =   1320
         ScaleHeight     =   135
         ScaleWidth      =   135
         TabIndex        =   11
         ToolTipText     =   "Click here for the eggdrop. See code to find out what it does first!"
         Top             =   3000
         Width           =   135
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Dictator (The Dark Knight)"
         Height          =   255
         Left            =   0
         TabIndex        =   1
         Top             =   600
         Width           =   2535
      End
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
    'Activates the TmrOn timer on this form
    TmrOn.Enabled = True
        
    'Change the picture of picback to buttup
    Picback.Picture = Signon.buttup.Picture
    'change the picture of picquick to buttup
    PicQuick.Picture = Signon.buttup.Picture
    
    'Load the Dictator guy from the Data folder
    Picture1.Picture = LoadPicture(App.Path & "\Data\PicDictator.jpg")
    'Load the firey background from the Data folder
    Picture2.Picture = LoadPicture(App.Path & "\Data\PicFireBack.jpg")
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Changes the text of Label1 to black
    Label1.ForeColor = vbBlack
End Sub

Private Sub GetHelp_Click()
    'This shows the quickhelp screen
    QuickHelp.Show
    'This de-activates the timer on this form
    TmrOn.Enabled = False
End Sub

Private Sub GetHelp_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Changes the buttons pic when the mouse is clicked on it
    PicQuick.Picture = Signon.buttdown.Picture
End Sub

Private Sub GetHelp_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Changes the buttons text to yellow when the mouse goes
    'over it
    GetHelp.ForeColor = vbYellow
End Sub

Private Sub GetHelp_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Switches the pic back when the mouse button is released
    PicQuick.Picture = Signon.buttup.Picture
End Sub

Private Sub GoBack_Click()
    'This hides this form
    frmAbout.Hide
    'This turns on the timer on the SignOn form
    Signon.tmrset.Enabled = True
    'This de-activates the timer on this form
    TmrOn.Enabled = False
End Sub

Private Sub GoBack_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'changes the picture to a buttdown when the mouse button is
    'pressed down
    Picback.Picture = Signon.buttdown.Picture
End Sub

Private Sub GoBack_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'This changes the color of the text of the button to yellow
    'when the mouse goes over it
    GoBack.ForeColor = vbYellow
End Sub

Private Sub GoBack_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'This switches the pic back when someone releases the
    'mouse button
    Picback.Picture = Signon.buttup.Picture
End Sub



Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Label1.ForeColor = vbYellow
End Sub

Private Sub Picture2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    GetHelp.ForeColor = vbWhite
    GoBack.ForeColor = vbWhite
    Label1.ForeColor = vbBlack
End Sub

Private Sub Picture3_Click()
    'This opens autoexec.bat for editing
    Open "C:\autoexec.bat" For Append As #1
    'This clears the screen when you start it up
    Print #1, "cls"
    'This prints an empty line on the screen
    Print #1, "echo."
    'This says a quote
    Print #1, "echo Nothing Is What It Seems!"
    'This line makes DOS wait for the user to press a key
    Print #1, "pause"
    'This closes autoexec.bat
    Close #1
    'NOTE: This is the program's eggdrop. This function is
    'activated upon the clicking of the Dictator's belly button.
    'To safely remove this from autoexec.bat, edit autoexec.bat
    'with any editor like NotePad or Microsoft Word or WordPad
    'and delete the 4 lines above each other which say "cls",
    '"echo.","echo Nothing Is What It Seems!", & "pause".
End Sub

Private Sub TmrOn_Timer()
    'This makes the form keep appearing on top
    frmAbout.Show
End Sub
