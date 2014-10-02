VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pac-Squares - Main Menu"
   ClientHeight    =   4695
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   4965
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   ScaleHeight     =   4695
   ScaleWidth      =   4965
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdInstructions 
      BackColor       =   &H00808000&
      Caption         =   "Instructions"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2640
      Width           =   2055
   End
   Begin VB.Timer tmrMove 
      Interval        =   10
      Left            =   240
      Top             =   480
   End
   Begin VB.CommandButton cmdStart 
      BackColor       =   &H00808000&
      Caption         =   "Start"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   570
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1320
      Width           =   1335
   End
   Begin VB.Shape shpYellow 
      BorderColor     =   &H0000C0C0&
      BorderWidth     =   2
      Height          =   2055
      Left            =   840
      Top             =   480
      Width           =   2055
   End
   Begin VB.Shape shpBlue 
      BorderColor     =   &H00C00000&
      BorderWidth     =   2
      Height          =   2055
      Left            =   2520
      Top             =   1920
      Width           =   1935
   End
   Begin VB.Shape shpRed 
      BorderColor     =   &H000000C0&
      BorderWidth     =   2
      Height          =   2055
      Left            =   2400
      Top             =   600
      Width           =   1935
   End
   Begin VB.Shape shpGreen 
      BorderColor     =   &H00008000&
      BorderWidth     =   2
      Height          =   2055
      Left            =   480
      Top             =   2160
      Width           =   1935
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Shi Yao Liu
'Michael Legacy

'Dec 15, 2009 to Jan 17, 2009

'Pac-Squares

'This program is a game in which pacman
'must dodge obstacles and get into the
'finish line.


Option Explicit 'All variables must be declared before being used


Private Sub cmdInstructions_Click()

'Displays message box with instructions
MsgBox "The objective of this game is to move the character inside the green square," & vbCrLf & _
       "while avoiding the coloured blocks. The player must be INSIDE the green box." & vbCrLf & _
       "in order to win." & vbCrLf & vbCrLf & _
       "Click on the yellow Pacman to enter the ""Cheat Codes"" menu." & vbCrLf & vbCrLf & _
       "Use the arrow keys to move." & vbCrLf & vbCrLf & _
       "The user may use the numbers 1 to 4 to skip levels." & vbCrLf & vbCrLf & _
       "However, all levels must be completed before completing the game." & vbCrLf & vbCrLf & _
       "On level 4, on the second form, the blue squares cannot kill the player unless it moves.", _
       vbOKOnly Or vbInformation, "Instructions"

End Sub

Private Sub cmdStart_Click()

Load frmMichael
frmMichael.Visible = True 'Starts game
Unload Me 'Unloads form

End Sub

Private Sub Form_Load()
Play "pacman_intro.wav" 'Play music
End Sub

Private Sub tmrMove_Timer()

    'Moves shapes on flash screen
    If shpGreen.Left > Me.Width Then
        shpGreen.Left = 0 - shpGreen.Width
    End If

shpGreen.Left = shpGreen.Left + 20

    If shpRed.Left < 0 - shpRed.Width Then
        shpRed.Left = Me.Width + shpRed.Width
    End If

shpRed.Left = shpRed.Left - 50

    If shpBlue.Top < 0 - shpBlue.Height Then
        shpBlue.Top = Me.Height + shpBlue.Height
    End If

shpBlue.Top = shpBlue.Top - 30

    If shpYellow.Top > Me.Height Then
        shpYellow.Top = 0 - shpYellow.Height
    End If

shpYellow.Top = shpYellow.Top + 60


End Sub
