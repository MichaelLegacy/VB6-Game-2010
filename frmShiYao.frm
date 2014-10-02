VERSION 5.00
Begin VB.Form frmShiYao 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Level ShiYao"
   ClientHeight    =   5475
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   11445
   Icon            =   "frmShiYao.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5475
   ScaleWidth      =   11445
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrPacman 
      Interval        =   150
      Left            =   5160
      Top             =   2520
   End
   Begin VB.Timer tmrLevel 
      Interval        =   1
      Left            =   10800
      Top             =   2880
   End
   Begin VB.Timer tmrMove 
      Interval        =   1
      Left            =   0
      Top             =   3240
   End
   Begin VB.Timer Timer7 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   11040
      Top             =   3720
   End
   Begin VB.Timer tmrCollision 
      Interval        =   1
      Left            =   0
      Top             =   1800
   End
   Begin VB.Timer Timer6 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   10920
      Top             =   4920
   End
   Begin VB.Timer Timer5 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   10920
      Top             =   120
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Caption         =   "Levels Completed"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000F&
      Height          =   735
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   11175
      Begin VB.Label lblCompleted 
         BackColor       =   &H00000000&
         Caption         =   "Level 4 - Incomplete"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Index           =   3
         Left            =   9120
         TabIndex        =   5
         Top             =   360
         Width           =   1815
      End
      Begin VB.Label lblCompleted 
         BackColor       =   &H00000000&
         Caption         =   "Level 3 - Incomplete"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Index           =   2
         Left            =   6240
         TabIndex        =   4
         Top             =   360
         Width           =   1815
      End
      Begin VB.Label lblCompleted 
         BackColor       =   &H00000000&
         Caption         =   "Level 2 - Incomplete"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Index           =   1
         Left            =   3240
         TabIndex        =   3
         Top             =   360
         Width           =   1815
      End
      Begin VB.Label lblCompleted 
         BackColor       =   &H00000000&
         Caption         =   "Level 1 - Incomplete"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   2
         Top             =   360
         Width           =   1815
      End
   End
   Begin VB.Image imgPacman 
      Height          =   495
      Left            =   240
      Picture         =   "frmShiYao.frx":5D2FA
      Stretch         =   -1  'True
      Top             =   2460
      Width           =   495
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H000000C0&
      BorderWidth     =   2
      Height          =   1215
      Index           =   2
      Left            =   12600
      Top             =   4320
      Width           =   1215
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H000000C0&
      BorderWidth     =   2
      Height          =   1215
      Index           =   1
      Left            =   12600
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H000000C0&
      BorderWidth     =   2
      Height          =   1215
      Index           =   0
      Left            =   12600
      Top             =   0
      Width           =   1215
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H00008000&
      Height          =   2055
      Left            =   10200
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H000000C0&
      BorderWidth     =   2
      Height          =   1215
      Index           =   2
      Left            =   7200
      Top             =   7000
      Width           =   1215
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H000000C0&
      BorderWidth     =   2
      Height          =   1215
      Index           =   1
      Left            =   4800
      Top             =   7000
      Width           =   1215
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H000000C0&
      BorderWidth     =   2
      Height          =   1215
      Index           =   0
      Left            =   2400
      Top             =   7000
      Width           =   1215
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H000000C0&
      BorderWidth     =   2
      Height          =   1215
      Index           =   3
      Left            =   8400
      Top             =   -2000
      Width           =   1215
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H000000C0&
      BorderWidth     =   2
      Height          =   1215
      Index           =   2
      Left            =   6000
      Top             =   -2000
      Width           =   1215
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H000000C0&
      BorderWidth     =   2
      Height          =   1215
      Index           =   1
      Left            =   3600
      Top             =   -2000
      Width           =   1215
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H000000C0&
      BorderWidth     =   2
      Height          =   1215
      Index           =   0
      Left            =   1200
      Top             =   -2000
      Width           =   1215
   End
   Begin VB.Shape Shape5 
      BorderColor     =   &H0000FFFF&
      Height          =   735
      Left            =   120
      Top             =   4680
      Width           =   11175
   End
   Begin VB.Label lblLevel 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Level 1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   4800
      Width           =   11175
   End
   Begin VB.Label lblDies 
      BackColor       =   &H00000000&
      Caption         =   "Dies: 0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   375
      Left            =   600
      TabIndex        =   6
      Top             =   1080
      Width           =   10695
   End
   Begin VB.Label lblEnd 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Click Here to Enter End Screen"
      ForeColor       =   &H00000000&
      Height          =   615
      Left            =   10320
      TabIndex        =   7
      Top             =   2400
      Width           =   975
   End
End
Attribute VB_Name = "frmShiYao"
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

Dim intLevel, dir, pacmanspeed, completed As Integer 'Declares global variables
Dim mouth As Boolean

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    'Move player around
    Select Case KeyCode
        
        Case vbKeyRight
            dir = 4
            tmrMove.Enabled = True
    
        Case vbKeyLeft
            dir = 3
            tmrMove.Enabled = True
    
        Case vbKeyUp
            dir = 1
            tmrMove.Enabled = True
            
        Case vbKeyDown
            dir = 2
            tmrMove.Enabled = True
    
    End Select

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

    'Level select codes
    Select Case KeyAscii

        Case 49
            intLevel = 1
            Call resetImgPos 'Procedure call
            Call stopImage
            Call resetBoxes
            lblLevel.Caption = "Level " & intLevel
        
        Case 50
            intLevel = 2
            Call resetImgPos
            Call stopImage
            Call resetBoxes
            lblLevel.Caption = "Level " & intLevel
        
        Case 51
            intLevel = 3
            Call resetImgPos
            Call stopImage
            Call resetBoxes
            lblLevel.Caption = "Level " & intLevel
        
        Case 52
            intLevel = 4
            Call resetImgPos
            Call stopImage
            Call resetBoxes
            lblLevel.Caption = "Level " & intLevel

    End Select

End Sub

Private Sub resetBoxes()
'Procedure - resets obstacles

Dim counter As Integer 'Declare local variable
    
    For counter = 0 To 3
        
        Shape1(counter).Top = -2000
    
    Next counter

    For counter = 0 To 2
        
        Shape2(counter).Top = 7000
        Shape4(counter).Left = 12600
    
    Next counter

End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
'tmrMove.Enabled = False
End Sub

Private Sub Form_Load()
Dim i As Integer 'Declare local variable

tmrPacman.Enabled = True 'Enables timer responsible for the opening and closing of Pacman's mouth

pacmanspeed = 50 'Sets Pacman's speed

'Sets and displays level
intLevel = 1
lblLevel.Caption = "Level " & intLevel
lblDies.Caption = "Dies: " & intDies

    'Sets obstacle colors
    For i = 0 To 2
    
        Shape4(i).BorderColor = vbBlue
        Shape2(i).BorderColor = vbYellow
    
    Next i

End Sub

Private Sub imgPacman_Click()
Dim cheat As String 'Declare local variable
Dim i As Integer

'Opens input box
cheat = InputBox("Enter a cheat code." & vbCrLf & vbCrLf & _
                 "superspeed" & vbCrLf & _
                 "gothrubloks" & vbCrLf & _
                 "fullpwrups" & vbCrLf & _
                 "cancel" & vbCrLf & _
                 "complete" & vbCrLf & vbCrLf _
                 , "Cheat Codes", "CHEATER!!!")

    'Cheat codes
    Select Case cheat

        Case "superspeed"
            pacmanspeed = 100
            
        Case "gothrubloks"
            tmrCollision.Enabled = False
            
        Case "fullpwrups"
            pacmanspeed = 100
            tmrCollision.Enabled = False
        
        Case "cancel"
            tmrCollision.Enabled = True
            pacmanspeed = 50
        
        Case "complete"
            
            For i = 0 To 3
                    
                lblCompleted(i).Caption = "Level 3 - Complete"
                lblCompleted(i).ForeColor = vbGreen
            
            Next i
            
            Call resetBoxes 'Procedure call
            Call stopImage
            Call resetImgPos
            intLevel = 5
            
    End Select

End Sub

Private Sub lblEnd_Click()

    'If game completed, advance to next form
    If lblEnd.ForeColor = vbGreen Then
        Load frmEnd
        frmEnd.Visible = True
        Unload Me
    End If
End Sub

Private Sub Timer5_Timer()
Dim counter As Integer 'Declare local variable
    
    'Moves obstacles
    For counter = 0 To 3
        
        If Shape1(counter).Top > Me.Height Then
            
            Shape1(counter).Top = 0 - Shape1(counter).Height
        
        End If

        Shape1(counter).Top = Shape1(counter).Top + 60

    Next counter

End Sub

Private Sub Timer6_Timer()
Dim counter As Integer 'Declare local variable

    'Moves obstacles
    For counter = 0 To 2
    
        If Shape2(counter).Top < 0 - Shape2(counter).Height Then
        
            Shape2(counter).Top = Me.Height + Shape2(counter).Height
    
        End If
    
        Shape2(counter).Top = Shape2(counter).Top - 50

    Next counter

End Sub

Private Sub Timer7_Timer()

Dim counter As Integer 'Declare local variable
    
    'Moves obstacles
    For counter = 0 To 2
        If Shape4(counter).Left + Shape4(counter).Width < 0 Then
            Shape4(counter).Left = Me.Width + Shape4(counter).Width
        End If

        Shape4(counter).Left = Shape4(counter).Left - 25

    Next counter

End Sub

Private Sub tmrCollision_Timer()
Call checkShape1Collision  'Procedure call
Call checkShape2Collision
Call checkShape4Collision
End Sub

Private Sub Dies()
intDies = intDies + 1 'Increase and display number of deaths
lblDies.Caption = "Dies: " & intDies
End Sub

Private Sub checkShape4Collision()
Dim counter As Integer 'Declare local variable



    'Check collision with obstacle
    For counter = 0 To 2
        
        If imgPacman.Top <> 2460 Or imgPacman.Left <> 240 Then
            
            If imgPacman.Top <= Shape4(counter).Top + Shape4(counter).Height And imgPacman.Left + imgPacman.Width > Shape4(counter).Left _
            And imgPacman.Left < Shape4(counter).Left + Shape4(counter).Width And imgPacman.Top + imgPacman.Height > Shape4(counter).Top Then
                
                Call Dies 'Procedure call
                Play "splat.wav" 'Play sound
                Call stopImage
                Call resetImgPos
            
            End If
        
        End If

    Next counter

End Sub

Private Sub disableAllTimers()
'Disable all timers
tmrCollision.Enabled = False
tmrMove.Enabled = False
Timer5.Enabled = False
Timer6.Enabled = False
Timer7.Enabled = False
tmrLevel.Enabled = False
End Sub

Private Sub Level()

Dim i As Integer 'Declare local variable
Dim displayLevel As Integer

    'Checks level
    If intLevel = 5 And lblCompleted(0).ForeColor = vbGreen And _
    lblCompleted(1).ForeColor = vbGreen And lblCompleted(2).ForeColor = vbGreen _
    And lblCompleted(3).ForeColor = vbGreen Then
        
        lblLevel.Caption = String(10, "*") & " GAME COMPLETED! " & String(10, "*")
        Call resetBoxes 'Procedure call
        Call stopImage
        Call resetImgPos
        Call disableAllTimers
        
        lblEnd.ForeColor = vbGreen
        Play "applause-2.wav" 'Plays sound
        MsgBox "Game Completed!" & vbCrLf & vbCrLf & _
         "Click the text in the green rectangle to continue!", vbOKOnly Or vbInformation, "Game Completed!"
        
    ElseIf intLevel > 4 And lblCompleted(3).ForeColor = vbGreen Then
        
        intLevel = 1
        Timer5.Enabled = False
        Timer6.Enabled = False
        Timer7.Enabled = False
        Call stopImage
        Call resetImgPos
        Call resetBoxes
        lblLevel.Caption = "Level " & intLevel

    End If

    'Checks goal collision
    If imgPacman.Left > Shape3.Left And imgPacman.Top > Shape3.Top And _
    imgPacman.Top + imgPacman.Height < Shape3.Top + Shape3.Height Then
        
        Play "applause.wav" 'Plays sound
        lblCompleted(intLevel - 1).Caption = "Level " & intLevel & "- Completed"
        lblCompleted(intLevel - 1).ForeColor = vbGreen
        Call resetBoxes
        Call stopImage
        Call resetImgPos
        intLevel = intLevel + 1
         
        If intLevel = 5 Then
            lblLevel.Caption = String(10, "*") & " GAME COMPLETED! " & String(10, "*")
        Else
        
        lblLevel.Caption = "Level " & intLevel
        displayLevel = intLevel - 1
        MsgBox "Level " & displayLevel & " Completed!", vbOKOnly Or vbInformation, "Level Completed!"
        
        End If

    End If

    'Checks level and moves obstacles and enables appropriate timers
    Select Case intLevel

        Case 1
            Timer5.Enabled = False
            Timer6.Enabled = False
            Timer7.Enabled = False
        Case 2
            Timer5.Enabled = True
            Timer6.Enabled = False
            Timer7.Enabled = False
        Case 3
            Timer5.Enabled = True
            Timer6.Enabled = True
            Timer7.Enabled = False
        
        Case 4
            Timer5.Enabled = True
            Timer6.Enabled = True
            Timer7.Enabled = True
            
    End Select

End Sub

Private Sub stopImage()
tmrMove.Enabled = False 'Stops player from moving
End Sub

Private Sub resetImgPos()
imgPacman.Top = 2460 'Resets Pacman
imgPacman.Left = 240
End Sub


Private Sub checkShape1Collision()
Dim counter As Integer 'Declare local variable


    
    'Check collision with obstacle
    For counter = 0 To 3

        If imgPacman.Top <= Shape1(counter).Top + Shape1(counter).Height And imgPacman.Left + imgPacman.Width > Shape1(counter).Left _
        And imgPacman.Left < Shape1(counter).Left + Shape1(counter).Width And imgPacman.Top + imgPacman.Height > Shape1(counter).Top Then
            
            Play "splat.wav" 'Plays sound
            Call Dies 'Procedure call
            Call stopImage
            Call resetImgPos
        
        End If

    Next counter

End Sub

Private Sub checkShape2Collision()
Dim counter As Integer 'Declare local variable



    'Check collision with obstacle
    For counter = 0 To 2

        If imgPacman.Top <= Shape2(counter).Top + Shape2(counter).Height And imgPacman.Left + imgPacman.Width > Shape2(counter).Left _
        And imgPacman.Left < Shape2(counter).Left + Shape2(counter).Width And imgPacman.Top + imgPacman.Height > Shape2(counter).Top Then
            
            Play "splat.wav"
            Call Dies 'Procedure call
            Call stopImage
            Call resetImgPos
        
        End If

    Next counter
End Sub

Private Sub tmrLevel_Timer()
Call Level 'Procedure call
End Sub

Private Sub tmrMove_Timer()

    'Checks direction
    Select Case dir

        Case 1
            imgPacman.Top = imgPacman.Top - pacmanspeed

        If imgPacman.Top < 0 Then
            imgPacman.Top = Me.Height
        End If
    
        Case 2
            imgPacman.Top = imgPacman.Top + pacmanspeed
            
            If imgPacman.Top > Me.Height Then
                imgPacman.Top = 0 - imgPacman.Height
            End If
    
        Case 3
            imgPacman.Left = imgPacman.Left - pacmanspeed
            If imgPacman.Left < 0 Then
                imgPacman.Left = 0
            End If
    
        Case 4
            imgPacman.Left = imgPacman.Left + pacmanspeed
            If imgPacman.Left > Me.Width Then
                imgPacman.Left = 0 - imgPacman.Width
            End If
    
    End Select

End Sub

Private Sub tmrPacman_Timer()

'Opens and closes mouth
    If mouth Then
        imgPacman.Picture = LoadPicture(App.Path & "\Pacman copy2.gif")
        mouth = False
    Else
        imgPacman.Picture = LoadPicture(App.Path & "\Pacman_Right.gif")
        mouth = True
    End If
End Sub
