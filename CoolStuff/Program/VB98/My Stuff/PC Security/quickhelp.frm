VERSION 5.00
Begin VB.Form QuickHelp 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Quick Help"
   ClientHeight    =   3585
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4905
   ForeColor       =   &H000000FF&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3585
   ScaleWidth      =   4905
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer 
      Interval        =   1
      Left            =   1440
      Top             =   480
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   1660
      ScaleHeight     =   240
      ScaleWidth      =   1440
      TabIndex        =   2
      Top             =   3120
      Width           =   1440
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Back To About"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   0
         TabIndex        =   3
         Top             =   0
         Width           =   1455
      End
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Dictator@MyPlace.com                        Http://Dictator.50Megs.Com"
      ForeColor       =   &H00FFFF00&
      Height          =   255
      Left            =   0
      TabIndex        =   6
      Top             =   3360
      Width           =   4815
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H000000FF&
      Height          =   3345
      Left            =   0
      Top             =   240
      Width           =   4900
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   " Quick Help"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   1815
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   4560
      TabIndex        =   4
      Top             =   0
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FF0000&
      BorderColor     =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Left            =   0
      Top             =   0
      Width           =   4900
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   $"quickhelp.frx":0000
      ForeColor       =   &H0000FF00&
      Height          =   2175
      Left            =   120
      TabIndex        =   1
      Top             =   1200
      Width           =   4695
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Quick Help"
      BeginProperty Font 
         Name            =   "Eclipse"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   4695
   End
End
Attribute VB_Name = "QuickHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Picture1.Picture = Signon.buttup.Picture

    Label3.ForeColor = vbWhite
    Label4.ForeColor = vbGreen
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Label3.ForeColor = vbWhite
    Label4.ForeColor = vbGreen
End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Label4.ForeColor = vbGreen
End Sub

Private Sub Label2_Click()
    If Label2.ForeColor = vbGreen Then
        Label2.ForeColor = vbRed
    ElseIf Label2.ForeColor = vbRed Then
        Label2.ForeColor = vbBlue
    ElseIf Label2.ForeColor = vbBlue Then
        Label2.ForeColor = vbYellow
    ElseIf Label2.ForeColor = vbYellow Then
        Label2.ForeColor = vbWhite
    Else
        Label2.ForeColor = vbGreen
    End If
End Sub

Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Label3.ForeColor = vbWhite
End Sub

Private Sub Label3_Click()
    QuickHelp.Hide
End Sub

Private Sub Label3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Picture1.Picture = Signon.buttdown.Picture
End Sub

Private Sub Label3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Label3.ForeColor = vbYellow
End Sub

Private Sub Label3_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Picture1.Picture = Signon.buttup.Picture
End Sub

Private Sub Label4_Click()
    QuickHelp.Hide
End Sub

Private Sub Label4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Label4.Caption = "x"
End Sub

Private Sub Label4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Label4.ForeColor = vbYellow
End Sub

Private Sub Label4_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Label4.Caption = "X"
End Sub
