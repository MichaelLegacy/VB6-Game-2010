VERSION 5.00
Begin VB.Form frmEnd 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pac-Squares - End"
   ClientHeight    =   4695
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4995
   Icon            =   "frmEnd.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4695
   ScaleWidth      =   4995
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrReturn 
      Enabled         =   0   'False
      Interval        =   15
      Left            =   4560
      Top             =   1680
   End
   Begin VB.Timer tmrCredits 
      Interval        =   20
      Left            =   4560
      Top             =   1080
   End
   Begin VB.Label lblReturn 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   495
      Left            =   840
      TabIndex        =   12
      Top             =   3000
      Width           =   3255
   End
   Begin VB.Image imgMacrosoft 
      Height          =   495
      Left            =   840
      Picture         =   "frmEnd.frx":5D2FA
      Stretch         =   -1  'True
      Top             =   2280
      Width           =   3255
   End
   Begin VB.Label lblFinalScore 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   375
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   4335
   End
   Begin VB.Shape shpFinalScore 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H000040C0&
      BorderWidth     =   3
      Height          =   615
      Left            =   240
      Top             =   240
      Width           =   4575
   End
   Begin VB.Label lblCredits 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Reflection Report - Michael Legacy and Shi Yao Liu"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   10
      Left            =   240
      TabIndex        =   11
      Top             =   16080
      Width           =   4575
   End
   Begin VB.Label lblCredits 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1695
      Index           =   9
      Left            =   240
      TabIndex        =   10
      Top             =   13680
      Width           =   4575
   End
   Begin VB.Label lblCredits 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "User Guide - Michael Legacy and Shi Yao Liu"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   8
      Left            =   480
      TabIndex        =   9
      Top             =   12680
      Width           =   4095
   End
   Begin VB.Label lblCredits 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Encapsulation - Shi Yao Liu"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   7
      Left            =   360
      TabIndex        =   8
      Top             =   11680
      Width           =   4335
   End
   Begin VB.Label lblCredits 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "End Screen - Shi Yao Liu"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   6
      Left            =   720
      TabIndex        =   7
      Top             =   10680
      Width           =   3615
   End
   Begin VB.Label lblCredits 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Game Levels - Michael Legacy and Shi Yao Liu"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   5
      Left            =   360
      TabIndex        =   6
      Top             =   9680
      Width           =   4335
   End
   Begin VB.Label lblCredits 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Game Intro Screen - Shi Yao Liu"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   4
      Left            =   720
      TabIndex        =   5
      Top             =   8680
      Width           =   3615
   End
   Begin VB.Label lblCredits 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Feasibility Study - Michael Legacy"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   3
      Left            =   720
      TabIndex        =   4
      Top             =   7680
      Width           =   3615
   End
   Begin VB.Label lblCredits 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Pseudocode - Shi Yao Liu"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   2
      Left            =   720
      TabIndex        =   3
      Top             =   6680
      Width           =   3615
   End
   Begin VB.Label lblCredits 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Scope Document - Shi Yao Liu"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   720
      TabIndex        =   2
      Top             =   5680
      Width           =   3615
   End
   Begin VB.Label lblCredits 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Project Manager and Lead Programmer - Shi Yao Liu"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   4680
      Width           =   4815
   End
End
Attribute VB_Name = "frmEnd"
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

Private Sub Form_Load()

lblFinalScore.Caption = "You finished with " & intDies & " dies." 'Displays number of times died

Play "pacman_dies_y" 'Plays music

lblCredits(9).Caption = "Testing - Michael Legacy and Shi Yao Liu" & vbCrLf & _
"Chris Goguen, Vanyel Tjia" & vbCrLf & "Brandon Vandermeer, Nathan Lynds" & vbCrLf & _
"Ryan Lilley, Jordan Tandan" & vbCrLf & "Adam Macdonald, SpencerTimms" & vbCrLf & _
"Kieren McCowell, Steven Pretty" & vbCrLf & "Michael Vanderweide, Zac Goddard" 'Displays the people who tested (played) this game in a label

lblReturn.Caption = "Click on the Logo" & vbCrLf & "to Return to the Main Menu" 'Displays instructions in a label

imgMacrosoft.Top = lblCredits(10).Top + lblCredits(10).Height + 2000 'Moves "Macrosoft" logo so that is is below the visible parts of the form

lblReturn.Top = Me.Height 'Moves label so that is is below the visible parts of the form
End Sub

Private Sub imgMacrosoft_Click()
Load frmMain
frmMain.Show 'Displays Splash Screen
intDies = 0 'Resets number of deaths
Unload Me 'Unloads current form

End Sub

Private Sub tmrCredits_Timer()

Dim i As Integer 'Declare local variable

    'Moves credits
    For i = 0 To 10
        lblCredits(i).Top = lblCredits(i).Top - 20
    Next i


imgMacrosoft.Top = imgMacrosoft.Top - 20 'Moves "Macrosoft" logo up

    'Stops "Macrosoft" logo once it reaches a certain spot on the form
    'and disables the timer responsible for moving the logo
    If imgMacrosoft.Top <= 2000 Then
        tmrCredits.Enabled = False
        tmrReturn.Enabled = True
    End If

End Sub

Private Sub tmrReturn_Timer()

lblReturn.Top = lblReturn.Top - 20 'Moves label with instructions up
    
    'Stops label with instructions once it reaches a certain spot on the form
    'and disables the timer responsible for moving the label
    If lblReturn.Top < imgMacrosoft.Top + imgMacrosoft.Height + 200 Then
        tmrReturn.Enabled = False
    End If

End Sub
