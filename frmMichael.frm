VERSION 5.00
Begin VB.Form frmMichael 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Level Mike"
   ClientHeight    =   5655
   ClientLeft      =   45
   ClientTop       =   495
   ClientWidth     =   11670
   DrawMode        =   12  'Nop
   Icon            =   "frmMichael.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5655
   ScaleWidth      =   11670
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrPacman 
      Interval        =   150
      Left            =   5280
      Top             =   2640
   End
   Begin VB.Timer tmrObstacleMove2 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   3960
      Top             =   4800
   End
   Begin VB.Timer tmrLevelSelect 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   6360
      Top             =   4800
   End
   Begin VB.Timer tmrWalls 
      Interval        =   1
      Left            =   5880
      Top             =   4800
   End
   Begin VB.Timer tmrCollision 
      Interval        =   1
      Left            =   5400
      Top             =   4800
   End
   Begin VB.Timer tmrObstacleMove 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   4440
      Top             =   4800
   End
   Begin VB.Timer tmrObstacleMove3 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   3480
      Top             =   4800
   End
   Begin VB.Timer tmrPlayerMove 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   4920
      Top             =   4800
   End
   Begin VB.Image imgPlayer 
      Height          =   495
      Left            =   240
      Picture         =   "frmMichael.frx":5D2FA
      Stretch         =   -1  'True
      Top             =   2160
      Width           =   495
   End
   Begin VB.Shape Shape5 
      BorderColor     =   &H0000FFFF&
      Height          =   735
      Left            =   240
      Top             =   4560
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
      Left            =   360
      TabIndex        =   5
      Top             =   4680
      Width           =   11175
   End
   Begin VB.Label lblDies 
      BackColor       =   &H00000000&
      Caption         =   "Dies: 0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   9615
   End
   Begin VB.Label lblpart4 
      BackColor       =   &H00000000&
      Caption         =   "Level 4  InComplete"
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   10080
      TabIndex        =   3
      Top             =   2520
      Width           =   1455
   End
   Begin VB.Label lblpart3 
      BackColor       =   &H00000000&
      Caption         =   "Level 3  InComplete"
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   10080
      TabIndex        =   2
      Top             =   1800
      Width           =   1455
   End
   Begin VB.Label lblpart2 
      BackColor       =   &H00000000&
      Caption         =   "Level 2  InComplete"
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   10080
      TabIndex        =   1
      Top             =   1080
      Width           =   1455
   End
   Begin VB.Label lblpart1 
      BackColor       =   &H00000000&
      Caption         =   "Level 1  InComplete"
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   10080
      TabIndex        =   0
      Top             =   360
      Width           =   1575
   End
   Begin VB.Shape shpBottomLeft 
      BorderColor     =   &H00008000&
      Height          =   1215
      Left            =   1320
      Top             =   3000
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Shape shpTopRight 
      BorderColor     =   &H00008000&
      Height          =   1215
      Left            =   7440
      Top             =   720
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Shape shpBottomRight 
      BorderColor     =   &H00008000&
      Height          =   1215
      Left            =   7440
      Top             =   3000
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Shape shpTopLeft 
      BorderColor     =   &H00008000&
      Height          =   1215
      Left            =   1320
      Top             =   720
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00008000&
      X1              =   8640
      X2              =   9120
      Y1              =   1920
      Y2              =   1920
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00008000&
      X1              =   8640
      X2              =   9120
      Y1              =   3000
      Y2              =   3000
   End
   Begin VB.Shape shpEndwalls 
      BorderColor     =   &H00008000&
      Height          =   1095
      Left            =   8640
      Top             =   1920
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Shape shpStartWalls 
      BorderColor     =   &H00008000&
      Height          =   1095
      Left            =   0
      Top             =   1920
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Shape shpOuterWallsTopBottom 
      BorderColor     =   &H00008000&
      FillColor       =   &H00FFFFFF&
      Height          =   3495
      Left            =   2520
      Top             =   720
      Visible         =   0   'False
      Width           =   4935
   End
   Begin VB.Shape shpFinish 
      BorderColor     =   &H00008000&
      Height          =   1095
      Left            =   9120
      Top             =   1920
      Width           =   735
   End
   Begin VB.Shape shpObstacle 
      BorderColor     =   &H00008000&
      Height          =   735
      Left            =   7680
      Shape           =   5  'Rounded Square
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Shape shpObstacle2 
      BorderColor     =   &H00008000&
      Height          =   735
      Left            =   1560
      Shape           =   5  'Rounded Square
      Top             =   3240
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00008000&
      Height          =   4215
      Left            =   9960
      Top             =   0
      Width           =   1695
   End
   Begin VB.Line Line10 
      BorderColor     =   &H00008000&
      X1              =   1320
      X2              =   8640
      Y1              =   4200
      Y2              =   4200
   End
   Begin VB.Line Line9 
      BorderColor     =   &H00008000&
      X1              =   1320
      X2              =   8640
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Line Line8 
      BorderColor     =   &H00008000&
      X1              =   8640
      X2              =   8640
      Y1              =   720
      Y2              =   1920
   End
   Begin VB.Line Line7 
      BorderColor     =   &H00008000&
      X1              =   8640
      X2              =   8640
      Y1              =   3000
      Y2              =   4200
   End
   Begin VB.Line Line6 
      BorderColor     =   &H00008000&
      X1              =   1320
      X2              =   1320
      Y1              =   720
      Y2              =   1920
   End
   Begin VB.Line Line5 
      BorderColor     =   &H00008000&
      X1              =   1320
      X2              =   1320
      Y1              =   3000
      Y2              =   4200
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00008000&
      X1              =   1320
      X2              =   0
      Y1              =   3000
      Y2              =   3000
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00008000&
      X1              =   1320
      X2              =   0
      Y1              =   1920
      Y2              =   1920
   End
   Begin VB.Shape shpCentre 
      BorderColor     =   &H00008000&
      Height          =   1335
      Left            =   2520
      Top             =   1800
      Width           =   4935
   End
End
Attribute VB_Name = "frmMichael"
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
Dim PlayerSpeed As Integer 'Declares global variables
Dim ObstacleSpeed As Integer
Dim ObstacleDir As Integer
Dim Obstacle2Dir As Integer
Dim PlayerDir As Integer
Dim Level As Integer
Dim mouth As Boolean
Dim Code As String

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

'Moves player around
Select Case KeyCode

    Case vbKeyUp
        PlayerDir = 1
        tmrPlayerMove.Enabled = True
    Case vbKeyDown
        PlayerDir = 2
        tmrPlayerMove.Enabled = True
    Case vbKeyLeft
        PlayerDir = 3
        tmrPlayerMove.Enabled = True
    Case vbKeyRight
        PlayerDir = 4
        tmrPlayerMove.Enabled = True
        
End Select
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

    'Level select
    Select Case KeyAscii

        Case 49
            tmrLevelSelect.Enabled = True
            Level = 1
            shpObstacle.Top = 960
            shpObstacle.Left = 7680
            shpObstacle2.Top = 3240
            shpObstacle2.Left = 1560
            imgPlayer.Left = 240
            imgPlayer.Top = 2160
        Case 50
            tmrLevelSelect.Enabled = True
            Level = 2
            shpObstacle.Top = 960
            shpObstacle.Left = 7680
            shpObstacle2.Top = 3240
            shpObstacle2.Left = 1560
            imgPlayer.Left = 240
            imgPlayer.Top = 2160
        Case 51
            tmrLevelSelect.Enabled = True
            Level = 3
            shpObstacle.Top = 960
            shpObstacle.Left = 7680
            shpObstacle2.Top = 3240
            shpObstacle2.Left = 1560
            imgPlayer.Left = 240
            imgPlayer.Top = 2160
        Case 52
            tmrLevelSelect.Enabled = True
            Level = 4
            shpObstacle.Top = 960
            shpObstacle.Left = 7680
            shpObstacle2.Top = 3240
            shpObstacle2.Left = 1560
            imgPlayer.Left = 240
            imgPlayer.Top = 2160
    End Select

End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

'Disable move timer when key goes up
Select Case KeyCode

    Case vbKeyUp
        tmrPlayerMove.Enabled = False
    Case vbKeyDown
        tmrPlayerMove.Enabled = False
    Case vbKeyLeft
        tmrPlayerMove.Enabled = False
    Case vbKeyRight
        tmrPlayerMove.Enabled = False
        
End Select

End Sub

Private Sub Form_Load()
mouth = True 'Makes Pacman's mouth move
ObstacleDir = 1 'Sets direction
Obstacle2Dir = 2
ObstacleSpeed = 75 'Sets speed
PlayerSpeed = 55
Level = 1 'Sets level

End Sub


Private Sub imgplayer_Click()


'Opens input box
Code = InputBox("Insert a Cheat Code" & vbCrLf & vbCrLf & _
                "gothruwalls" & vbCrLf & "superspeed" & vbCrLf & _
                "supereasy" & vbCrLf & "fullpwrups" & vbCrLf _
                , "Cheat Codes")

'Cheat codes
If Code = "gothruwalls" Then
    tmrWalls.Enabled = False
    PlayerSpeed = 50
    ObstacleSpeed = 75
ElseIf Code = "superspeed" Then
    PlayerSpeed = 100
    tmrWalls.Enabled = True
    ObstacleSpeed = 75
ElseIf Code = "supereasy" Then
    ObstacleSpeed = 20
    tmrWalls.Enabled = True
    PlayerSpeed = 50
ElseIf Code = "fullpwrups" Then
    ObstacleSpeed = 20
    tmrWalls.Enabled = False
    PlayerSpeed = 100
Else
   If tmrObstacleMove3.Enabled = True Then
    tmrWalls.Enabled = True
   PlayerSpeed = 65
   ObstacleSpeed = 62
   Else
    tmrWalls.Enabled = True
    PlayerSpeed = 55
    ObstacleSpeed = 75
    End If
End If

End Sub


Private Sub tmrCollision_Timer()

'Check wall collisions
If imgPlayer.Left < 0 Then
    imgPlayer.Left = 0
ElseIf imgPlayer.Top < 0 Then
    imgPlayer.Top = 0
ElseIf imgPlayer.Top + imgPlayer.Height > Me.Height Then
    imgPlayer.Top = Me.Height - imgPlayer.Height
ElseIf imgPlayer.Left + imgPlayer.Width > Shape2.Left Then
    imgPlayer.Left = Shape2.Left - imgPlayer.Width
End If

'Check obstacle collisions
If imgPlayer.Top < shpObstacle.Top + shpObstacle.Height And imgPlayer.Left > shpObstacle.Left - imgPlayer.Width _
And imgPlayer.Left < shpObstacle.Left + shpObstacle.Width And imgPlayer.Top + imgPlayer.Height > shpObstacle.Top And shpObstacle.Visible = True Then
    imgPlayer.Left = 240
    imgPlayer.Top = 2160
    Play "splat.wav"
    intDies = intDies + 1
    lblDies.Caption = "Dies:" & intDies
End If

'Check obstacle collisions
If imgPlayer.Top < shpObstacle2.Top + shpObstacle2.Height And imgPlayer.Left > shpObstacle2.Left - imgPlayer.Width _
And imgPlayer.Left < shpObstacle2.Left + shpObstacle2.Width And imgPlayer.Top + imgPlayer.Height > shpObstacle2.Top And shpObstacle2.Visible = True Then
    imgPlayer.Left = 240
    imgPlayer.Top = 2160
    Play "splat.wav"
    intDies = intDies + 1
    lblDies.Caption = "Dies:" & intDies
End If

'Finish Spot'
If imgPlayer.Left + imgPlayer.Width >= shpFinish.Left + shpFinish.Width And imgPlayer.Top >= shpFinish.Top And imgPlayer.Top + imgPlayer.Height <= shpFinish.Top + shpFinish.Height Then
    
    Play "applause.wav"
    tmrPlayerMove.Enabled = False
    
    'Resets all obstacle and player locations
    imgPlayer.Left = 240
    imgPlayer.Top = 2160
    shpObstacle.Top = 960
    shpObstacle.Left = 7680
    shpObstacle2.Top = 3240
    shpObstacle2.Left = 1560
    MsgBox "You beat Level " & Level & " congratulations", vbInformation, "Level " & Level & " complete"

        If Level = 1 Then
            lblpart1.ForeColor = vbGreen
            lblpart1.Caption = "Level 1 Complete"
        ElseIf Level = 2 Then
            lblpart2.ForeColor = vbGreen
            lblpart2.Caption = "Level 2 Complete"
        ElseIf Level = 3 Then
            lblpart3.ForeColor = vbGreen
            lblpart3.Caption = "Level 3 Complete"
        ElseIf Level = 4 Then
            lblpart4.ForeColor = vbGreen
            lblpart4.Caption = "Level 4 Complete"
        End If
        
        Level = Level + 1 'Increment and display level
        lblLevel.Caption = "Level " & Level

        'Checks if all levels are complete
        If Level = 5 And lblpart1.ForeColor = vbGreen And lblpart2.ForeColor = vbGreen And lblpart3.ForeColor = vbGreen And lblpart4.ForeColor = vbGreen Then
            tmrPacman.Enabled = False
            Load frmShiYao
            frmShiYao.Visible = True
            Unload Me
        
        ElseIf Level = 5 Then
            Level = 1
        
        End If
                
        'Resets player position
        imgPlayer.Left = 240
        imgPlayer.Top = 2160
        tmrPlayerMove.Enabled = False
        tmrLevelSelect.Enabled = True
        End If

End Sub

Private Sub tmrLevelSelect_Timer()

'Checks level and activates / deactivate obstacle movements
If Level = 1 Then
    shpObstacle.Visible = False
    shpObstacle2.Visible = False
    tmrObstacleMove.Enabled = False
    tmrObstacleMove2.Enabled = False
    tmrObstacleMove3.Enabled = False

ElseIf Level = 2 Then
    tmrObstacleMove.Enabled = True
    tmrObstacleMove2.Enabled = False
    tmrObstacleMove3.Enabled = False

ElseIf Level = 3 Then

    tmrObstacleMove.Enabled = False
    tmrObstacleMove2.Enabled = True
    tmrObstacleMove3.Enabled = False

ElseIf Level = 4 Then
    If Code = "superspeed" Then
        PlayerSpeed = 100
        ObstacleSpeed = 62
    ElseIf Code = "fullpwrups" Then
        PlayerSpeed = 100
        ObstacleSpeed = 20
    ElseIf Code = "supereasy" Then
        PlayerSpeed = 65
        ObstacleSpeed = 20
    Else
        ObstacleSpeed = 62
        PlayerSpeed = 65
    End If
    tmrObstacleMove.Enabled = False
    tmrObstacleMove2.Enabled = False
    tmrObstacleMove3.Enabled = True

End If

lblLevel.Caption = "Level " & Level 'Displays current level
tmrLevelSelect.Enabled = False 'Disable current timer

End Sub

Private Sub tmrObstacleMove_Timer()

shpObstacle.Visible = True
shpObstacle2.Visible = True

    'Checks obstacle direction and when to change direction
    If ObstacleDir = 1 Then
        If shpObstacle.Top <= 960 Then
        ObstacleDir = 4
        End If
    shpObstacle.Top = shpObstacle.Top - ObstacleSpeed
    End If

    If Obstacle2Dir = 1 Then
        If shpObstacle2.Top <= 960 Then
        Obstacle2Dir = 4
        End If
    shpObstacle2.Top = shpObstacle2.Top - ObstacleSpeed
    End If


    If ObstacleDir = 2 Then
        If shpObstacle.Top >= 3240 Then
        ObstacleDir = 3
        End If
    shpObstacle.Top = shpObstacle.Top + ObstacleSpeed
    End If

    If Obstacle2Dir = 2 Then
        If shpObstacle2.Top >= 3240 Then
        Obstacle2Dir = 3
        End If
    shpObstacle2.Top = shpObstacle2.Top + ObstacleSpeed
    End If


    If ObstacleDir = 3 Then
        If shpObstacle.Left <= 1560 Then
        ObstacleDir = 1
        End If
    shpObstacle.Left = shpObstacle.Left - ObstacleSpeed
    End If

    If Obstacle2Dir = 3 Then
        If shpObstacle2.Left <= 1560 Then
        Obstacle2Dir = 1
        End If
    shpObstacle2.Left = shpObstacle2.Left - ObstacleSpeed
    End If


    If ObstacleDir = 4 Then
        If shpObstacle.Left >= 7680 Then
        ObstacleDir = 2
        End If
    shpObstacle.Left = shpObstacle.Left + ObstacleSpeed
    End If

    If Obstacle2Dir = 4 Then
        If shpObstacle2.Left >= 7680 Then
        Obstacle2Dir = 2
        End If
    shpObstacle2.Left = shpObstacle2.Left + ObstacleSpeed
    End If

End Sub

Private Sub tmrObstacleMove2_Timer()

shpObstacle.Visible = True 'Show obstacles
shpObstacle2.Visible = True

'Checks obstacle direction and when to change direction
If ObstacleDir = 1 Then
        If shpObstacle.Top <= 960 Then
        ObstacleDir = 3
        End If
    shpObstacle.Top = shpObstacle.Top - ObstacleSpeed
    End If

    If Obstacle2Dir = 1 Then
        If shpObstacle2.Top <= 960 Then
        Obstacle2Dir = 3
        End If
    shpObstacle2.Top = shpObstacle2.Top - ObstacleSpeed
    End If


    If ObstacleDir = 2 Then
        If shpObstacle.Top >= 3240 Then
        ObstacleDir = 4
        End If
    shpObstacle.Top = shpObstacle.Top + ObstacleSpeed
    End If

    If Obstacle2Dir = 2 Then
        If shpObstacle2.Top >= 3240 Then
        Obstacle2Dir = 4
        End If
    shpObstacle2.Top = shpObstacle2.Top + ObstacleSpeed
    End If


    If ObstacleDir = 3 Then
        If shpObstacle.Left <= 1560 Then
        ObstacleDir = 2
        End If
    shpObstacle.Left = shpObstacle.Left - ObstacleSpeed
    End If

    If Obstacle2Dir = 3 Then
        If shpObstacle2.Left <= 1560 Then
        Obstacle2Dir = 2
        End If
    shpObstacle2.Left = shpObstacle2.Left - ObstacleSpeed
    End If


    If ObstacleDir = 4 Then
        If shpObstacle.Left >= 7680 Then
        ObstacleDir = 1
        End If
    shpObstacle.Left = shpObstacle.Left + ObstacleSpeed
    End If

    If Obstacle2Dir = 4 Then
        If shpObstacle2.Left >= 7680 Then
        Obstacle2Dir = 1
        End If
    shpObstacle2.Left = shpObstacle2.Left + ObstacleSpeed
    End If

End Sub

Private Sub tmrObstacleMove3_Timer()

shpObstacle.Visible = True 'Show obstacles
shpObstacle2.Visible = True

'Checks obstacle direction and when to change direction
If ObstacleDir = 1 Then
        If shpObstacle.Top <= 960 Then
        ObstacleDir = 3
        End If
    shpObstacle.Top = shpObstacle.Top - ObstacleSpeed
    End If

    If Obstacle2Dir = 1 Then
        If shpObstacle2.Top <= 960 Then
        Obstacle2Dir = 2
        End If
    shpObstacle2.Top = shpObstacle2.Top - ObstacleSpeed
    End If


    If ObstacleDir = 2 Then
        If shpObstacle.Top >= 3240 Then
        ObstacleDir = 1
        End If
    shpObstacle.Top = shpObstacle.Top + ObstacleSpeed
    End If

    If Obstacle2Dir = 2 Then
        If shpObstacle2.Top >= 3240 Then
        Obstacle2Dir = 4
        End If
    shpObstacle2.Top = shpObstacle2.Top + ObstacleSpeed
    End If


    If ObstacleDir = 3 Then
        If shpObstacle.Left <= 1560 Then
        ObstacleDir = 4
        End If
    shpObstacle.Left = shpObstacle.Left - ObstacleSpeed
    End If

    If Obstacle2Dir = 3 Then
        If shpObstacle2.Left <= 1560 Then
        Obstacle2Dir = 1
        End If
    shpObstacle2.Left = shpObstacle2.Left - ObstacleSpeed
    End If


    If ObstacleDir = 4 Then
        If shpObstacle.Left >= 7680 Then
            ObstacleDir = 2
        End If
        shpObstacle.Left = shpObstacle.Left + ObstacleSpeed
    End If

    If Obstacle2Dir = 4 Then
        If shpObstacle2.Left >= 7680 Then
        Obstacle2Dir = 3
        End If
    shpObstacle2.Left = shpObstacle2.Left + ObstacleSpeed
    End If

End Sub

Private Sub tmrPacman_Timer()
    'Loads picture
    If mouth Then
        imgPlayer.Picture = LoadPicture(App.Path & "\Pacman copy2.gif")
        mouth = False
    Else
        imgPlayer.Picture = LoadPicture(App.Path & "\Pacman_Right.gif")
        mouth = True
    End If
End Sub

Private Sub tmrPlayerMove_Timer()

    'Checks player directions
    If PlayerDir = 1 Then
        imgPlayer.Top = imgPlayer.Top - PlayerSpeed
    End If

    If PlayerDir = 2 Then
        imgPlayer.Top = imgPlayer.Top + PlayerSpeed
    End If

    If PlayerDir = 3 Then
        imgPlayer.Left = imgPlayer.Left - PlayerSpeed
    End If

    If PlayerDir = 4 Then
        imgPlayer.Left = imgPlayer.Left + PlayerSpeed
    End If

End Sub

Private Sub tmrWalls_Timer()

'top of centre'
If imgPlayer.Top + imgPlayer.Height >= shpCentre.Top And imgPlayer.Left + imgPlayer.Width >= shpCentre.Left And imgPlayer.Left <= shpCentre.Left + shpCentre.Width And imgPlayer.Top < shpCentre.Top Then
imgPlayer.Top = shpCentre.Top - imgPlayer.Height - 1
End If

'bottom of centre'
If imgPlayer.Top <= shpCentre.Top + shpCentre.Height And imgPlayer.Left + imgPlayer.Width >= shpCentre.Left And imgPlayer.Left <= shpCentre.Left + shpCentre.Width And imgPlayer.Top + imgPlayer.Height > shpCentre.Top + shpCentre.Height Then
imgPlayer.Top = shpCentre.Top + shpCentre.Height + 1
End If

'left of centre'
If imgPlayer.Left + imgPlayer.Width >= shpCentre.Left And imgPlayer.Top <= shpCentre.Top + shpCentre.Height And imgPlayer.Top + imgPlayer.Height >= shpCentre.Top And imgPlayer.Left < shpCentre.Left Then
imgPlayer.Left = shpCentre.Left - imgPlayer.Width - 1
End If

'right of centre'
If imgPlayer.Left <= shpCentre.Left + shpCentre.Width And imgPlayer.Top <= shpCentre.Top + shpCentre.Height And imgPlayer.Top + imgPlayer.Height >= shpCentre.Top And imgPlayer.Left + imgPlayer.Width > shpCentre.Left + shpCentre.Width Then
imgPlayer.Left = shpCentre.Left + shpCentre.Width + 1
End If

'outside top wall'
If imgPlayer.Top <= shpOuterWallsTopBottom.Top Then
imgPlayer.Top = shpOuterWallsTopBottom.Top + 1
End If

'outside bottom wall'
If imgPlayer.Top + imgPlayer.Height >= shpOuterWallsTopBottom.Top + shpOuterWallsTopBottom.Height Then
imgPlayer.Top = shpOuterWallsTopBottom.Top + shpOuterWallsTopBottom.Height - imgPlayer.Height - 1
End If

'start/finish area top walls'
If imgPlayer.Top <= shpStartWalls.Top And imgPlayer.Left <= shpStartWalls.Left + shpStartWalls.Width And imgPlayer.Top + imgPlayer.Height > shpStartWalls.Top Or imgPlayer.Top <= shpStartWalls.Top And imgPlayer.Left + imgPlayer.Width >= shpEndwalls.Left And imgPlayer.Top + imgPlayer.Height > shpStartWalls.Top Then
imgPlayer.Top = shpStartWalls.Top + 1
End If

'start/finish area bottom walls'
If imgPlayer.Top + imgPlayer.Height >= shpStartWalls.Top + shpStartWalls.Height And imgPlayer.Left <= shpStartWalls.Left + shpStartWalls.Width And imgPlayer.Top < shpStartWalls.Top + shpStartWalls.Height Or imgPlayer.Top + imgPlayer.Height >= shpStartWalls.Top + shpStartWalls.Height And imgPlayer.Left + imgPlayer.Width >= shpEndwalls.Left And imgPlayer.Top < shpStartWalls.Top + shpStartWalls.Height Then
imgPlayer.Top = shpStartWalls.Top + shpStartWalls.Height - imgPlayer.Height - 1
End If

'Wall to the left side of top'
If imgPlayer.Left <= shpTopLeft.Left And imgPlayer.Top < shpTopLeft.Top + shpTopLeft.Height And imgPlayer.Left + imgPlayer.Width >= shpTopLeft.Left Then
imgPlayer.Left = shpTopLeft.Left + 1
End If

'wall to the left side of the bottom'
If imgPlayer.Left <= shpBottomLeft.Left And imgPlayer.Top + imgPlayer.Height > shpBottomLeft.Top And imgPlayer.Left + imgPlayer.Width >= shpTopLeft.Left Then
imgPlayer.Left = shpBottomLeft.Left + 1
End If

'wall at the right of the top'
If imgPlayer.Left + imgPlayer.Width >= shpTopRight.Left + shpTopRight.Width And imgPlayer.Top < shpTopRight.Top + shpTopRight.Height And imgPlayer.Left + imgPlayer.Width >= shpTopLeft.Left And imgPlayer.Left < shpEndwalls.Left Then
imgPlayer.Left = shpTopRight.Left + shpTopRight.Width - imgPlayer.Width - 1
End If

'wall at the right of the bottom'
If imgPlayer.Left + imgPlayer.Width >= shpBottomRight.Left + shpBottomRight.Width And imgPlayer.Top + imgPlayer.Height > shpBottomRight.Top And imgPlayer.Left + imgPlayer.Width >= shpTopLeft.Left And imgPlayer.Left < shpEndwalls.Left Then
imgPlayer.Left = shpBottomRight.Left + shpBottomRight.Width - imgPlayer.Width - 1
End If

End Sub
