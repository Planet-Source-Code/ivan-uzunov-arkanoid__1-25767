VERSION 5.00
Begin VB.Form frmArquanoid 
   BackColor       =   &H00404040&
   BorderStyle     =   0  'None
   ClientHeight    =   7125
   ClientLeft      =   75
   ClientTop       =   -495
   ClientWidth     =   11655
   ControlBox      =   0   'False
   ForeColor       =   &H00000000&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7125
   ScaleWidth      =   11655
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picEnemyGrey 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      BorderStyle     =   0  'None
      Height          =   405
      Left            =   11280
      Picture         =   "frmArquanoid.frx":0000
      ScaleHeight     =   27
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   42
      TabIndex        =   20
      Top             =   1920
      Visible         =   0   'False
      Width           =   630
   End
   Begin VB.PictureBox picEnemyYellow 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      BorderStyle     =   0  'None
      Height          =   405
      Left            =   11280
      Picture         =   "frmArquanoid.frx":0DC2
      ScaleHeight     =   27
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   42
      TabIndex        =   19
      Top             =   2400
      Visible         =   0   'False
      Width           =   630
   End
   Begin VB.PictureBox picEnemyBlue 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      BorderStyle     =   0  'None
      Height          =   405
      Left            =   10560
      Picture         =   "frmArquanoid.frx":1B84
      ScaleHeight     =   27
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   42
      TabIndex        =   18
      Top             =   2400
      Visible         =   0   'False
      Width           =   630
   End
   Begin VB.PictureBox picBack 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   1500
      Left            =   10680
      Picture         =   "frmArquanoid.frx":2946
      ScaleHeight     =   96
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   96
      TabIndex        =   14
      Top             =   3000
      Width           =   1500
   End
   Begin VB.PictureBox picEnemyMask 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      BorderStyle     =   0  'None
      Height          =   405
      Left            =   10560
      Picture         =   "frmArquanoid.frx":3DA8
      ScaleHeight     =   27
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   42
      TabIndex        =   13
      Top             =   1920
      Visible         =   0   'False
      Width           =   630
   End
   Begin VB.PictureBox picEnemy 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      BorderStyle     =   0  'None
      Height          =   405
      Left            =   10560
      Picture         =   "frmArquanoid.frx":4B6A
      ScaleHeight     =   27
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   42
      TabIndex        =   12
      Top             =   1440
      Visible         =   0   'False
      Width           =   630
   End
   Begin VB.PictureBox picBonus 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   11160
      Picture         =   "frmArquanoid.frx":592C
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   11
      Top             =   1080
      Width           =   240
   End
   Begin VB.PictureBox picBonusMask 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   11520
      Picture         =   "frmArquanoid.frx":5C3E
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   10
      Top             =   1080
      Width           =   240
   End
   Begin VB.PictureBox picBackGround 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   8040
      Left            =   10800
      ScaleHeight     =   532
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   564
      TabIndex        =   9
      Top             =   5280
      Width           =   8520
   End
   Begin VB.PictureBox picPlayerMask 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H000000FF&
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   10560
      Picture         =   "frmArquanoid.frx":5F50
      ScaleHeight     =   17
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   57
      TabIndex        =   4
      Top             =   720
      Width           =   855
   End
   Begin VB.PictureBox picPlayer 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H000000FF&
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   10560
      Picture         =   "frmArquanoid.frx":6AFE
      ScaleHeight     =   17
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   57
      TabIndex        =   3
      Top             =   360
      Width           =   855
   End
   Begin VB.PictureBox picBallMask 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   10800
      Picture         =   "frmArquanoid.frx":76AC
      ScaleHeight     =   12
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   12
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   1080
      Width           =   180
   End
   Begin VB.PictureBox picBall 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   10560
      Picture         =   "frmArquanoid.frx":789E
      ScaleHeight     =   12
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   12
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   1080
      Width           =   180
   End
   Begin VB.PictureBox picGame 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H008080FF&
      BorderStyle     =   0  'None
      Height          =   6960
      Left            =   0
      ScaleHeight     =   464
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   544
      TabIndex        =   0
      Top             =   0
      Width           =   8160
   End
   Begin VB.Timer tmrMove 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   8760
      Top             =   1800
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "A R K A N O I D"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   8280
      TabIndex        =   24
      Top             =   165
      Width           =   1455
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      X1              =   8280
      X2              =   10120
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H80000009&
      BorderWidth     =   2
      Height          =   255
      Left            =   9840
      Shape           =   4  'Rounded Rectangle
      Top             =   120
      Width           =   255
   End
   Begin VB.Label lblExit 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   9900
      TabIndex        =   23
      Top             =   120
      Width           =   150
   End
   Begin VB.Label lblPause 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Pause"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   495
      Left            =   8400
      TabIndex        =   22
      Top             =   4560
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label lblLevelText 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "L e v e l"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   8640
      TabIndex        =   21
      Top             =   3840
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label lblLevel 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Haettenschweiler"
         Size            =   72
         Charset         =   204
         Weight          =   800
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1575
      Left            =   8640
      TabIndex        =   17
      Top             =   2400
      Width           =   1215
   End
   Begin VB.Label lblLevelEditor 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Level Edit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   8640
      TabIndex        =   16
      Top             =   6360
      Width           =   1215
   End
   Begin VB.Shape Shape3 
      BorderWidth     =   2
      Height          =   495
      Left            =   8520
      Shape           =   4  'Rounded Rectangle
      Top             =   6240
      Width           =   1455
   End
   Begin VB.Label lblNewGame 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "New Game"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   8520
      TabIndex        =   15
      Top             =   5640
      Width           =   1455
   End
   Begin VB.Shape Shape2 
      BorderWidth     =   2
      Height          =   495
      Left            =   8520
      Shape           =   4  'Rounded Rectangle
      Top             =   5520
      Width           =   1455
   End
   Begin VB.Label lblLives 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   9720
      TabIndex        =   8
      Top             =   1200
      Width           =   375
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Lives :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   8400
      TabIndex        =   7
      Top             =   1200
      Width           =   855
   End
   Begin VB.Label lblScore 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "00000"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   9240
      TabIndex        =   6
      Top             =   840
      Width           =   855
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Score:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   8400
      TabIndex        =   5
      Top             =   840
      Width           =   855
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00CA8826&
      BackStyle       =   1  'Opaque
      BorderWidth     =   4
      FillColor       =   &H00C0FFFF&
      Height          =   6975
      Left            =   8190
      Shape           =   4  'Rounded Rectangle
      Top             =   0
      Width           =   2055
   End
End
Attribute VB_Name = "frmArquanoid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***********************************************************************************
' Author    : Ivan Uzunov
' E- mail   : kicheto@goatrance.com
' Made with : Visual Basic 6.00
' Date      : 07-30-2001
'***********************************************************************************
Option Explicit

Dim FormWidth As Single
Dim FormHeight As Single

Dim BallVelocityX As Long
Dim BallVelocityY As Long

Dim ENEMY_WIDTH As Single
Dim ENEMY_HEIGHT As Single

Dim StartGame As Boolean
Dim LiveNumber As Integer
Dim LevelNumber As Integer
Dim Score As Long

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyLeft Then
       PlayerDirection.MoveLeft = True
    ElseIf KeyCode = vbKeyRight Then
       PlayerDirection.MoveRight = True
    ElseIf KeyCode = vbKeyN Then
       Call mnuNewGame_Click
    ElseIf KeyCode = vbKeyEscape Then
       Unload Me
    ElseIf KeyCode = vbKeyP Then
       If StartGame = True Then
          tmrMove.Enabled = Not tmrMove.Enabled
          lblPause.Visible = Not tmrMove.Enabled
       End If
    End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyLeft Then
       PlayerDirection.MoveLeft = False
    ElseIf KeyCode = vbKeyRight Then
       PlayerDirection.MoveRight = False
    ElseIf KeyCode = vbKeySpace Then
       StartGame = True
       'tmrMove.Enabled = True
    End If
End Sub

Private Sub Form_Load()
Dim I As Integer
Dim J As Integer
    frmArquanoid.Width = picGame.Width + Shape1.Width + 40
    frmArquanoid.Height = picGame.Height - 40
    
    Shape1.Top = 0
    Shape1.Height = frmArquanoid.Height
    
    FormWidth = picGame.ScaleWidth
    FormHeight = picGame.ScaleHeight
    
    BallRect.Width = picBall.ScaleWidth
    BallRect.Height = picBall.ScaleHeight
    
    ENEMY_WIDTH = picEnemy.ScaleWidth
    ENEMY_HEIGHT = picEnemy.ScaleHeight
    
    For I = 0 To 5
        For J = 0 To 4
            BitBlt picBackGround.hDC, picBack.ScaleWidth * I, picBack.ScaleHeight * J, picBack.ScaleWidth, picBack.ScaleHeight, picBack.hDC, 0, 0, vbSrcCopy
        Next J
    Next I
    picBackGround.Refresh
    Call GameBackGroungRefresh
End Sub

Private Sub PlayerMove()
    If PlayerDirection.MoveLeft = True Then
       If PlayerRect.Left - PlayerStep >= 0 Then
          PlayerRect.Left = PlayerRect.Left - PlayerStep
       Else
          PlayerRect.Left = 0
       End If
    ElseIf PlayerDirection.MoveRight = True Then
       If PlayerRect.Left + PlayerRect.Width + PlayerStep <= FormWidth Then
          PlayerRect.Left = PlayerRect.Left + PlayerStep
       Else
          PlayerRect.Left = FormWidth - PlayerRect.Width
       End If
    End If
    Call DrawBitmap(picPlayer, picPlayerMask.hDC, picGame.hDC, PlayerRect.Left, PlayerRect.Top)
End Sub

Private Sub BallMove()
Dim I As Long
    BallRect.Left = BallRect.Left + BallVelocityX
    If BallRect.Left <= 0 Then
       BallRect.Left = 0
       BallVelocityX = -BallVelocityX
    ElseIf BallRect.Left + BallRect.Width >= FormWidth Then
       BallRect.Left = FormWidth - BallRect.Width
       BallVelocityX = -BallVelocityX
    End If
    
    BallRect.Top = BallRect.Top + BallVelocityY
    If BallRect.Top <= 0 Then
       BallRect.Top = 0
       BallVelocityY = -BallVelocityY
    ElseIf BallRect.Top + BallRect.Height >= FormHeight Then
       Beep
       PlayerDirection.MoveLeft = False
       PlayerDirection.MoveRight = False
       LiveNumber = LiveNumber - 1
       lblLives.Caption = LiveNumber
       If LiveNumber = 0 Then
          MsgBox "Game Over", , "Arquanoid"
          Call mnuNewGame_Click
       Else
          Call SetStartPosition
       End If
    End If
    
    BallRect.Top = BallRect.Top + Abs(BallVelocityY) 'Çà äà íå çàñè÷à îáðàçà
    If BallPlayerCollision(BallRect, PlayerRect) = True Then
       BallRect.Top = PlayerRect.Top - BallRect.Height - 1
       BallVelocityY = -BallVelocityY
       Call PlaySound(2)
    Else
       BallRect.Top = BallRect.Top - Abs(BallVelocityY)
    End If
    
    For I = 0 To EnemyCount - 1
        If Enemy(I).Alive = True Then
           If Collision(Enemy(I).EnemyRect, BallRect) = True Then
              If Enemy(I).Color <> Grey Then
                 If picBall.Tag <> "fast" Then
                    Call ChangeVellosity(Enemy(I).EnemyRect)
                 End If
                 Score = Score + 20
                 lblScore.Caption = Score
              
                 Enemy(I).Alive = False
                 If Enemy(I).Bonus <> None Then
                    BonusCount = BonusCount + 1
                    ReDim Preserve Bonuses(BonusCount)
                    Bonuses(BonusCount).BonusRect.Width = picBonus.ScaleWidth
                    Bonuses(BonusCount).BonusRect.Height = picBonus.ScaleHeight
                    Bonuses(BonusCount).BonusRect.Top = Enemy(I).EnemyRect.Top + Enemy(I).EnemyRect.Height + 1
                    Bonuses(BonusCount).BonusRect.Left = Enemy(I).EnemyRect.Left + Enemy(I).EnemyRect.Width / 2 - Bonuses(BonusCount).BonusRect.Width / 2
                    Bonuses(BonusCount).BonusNotExist = False
                    Randomize
                    Bonuses(BonusCount).BonusType = Enemy(I).Bonus
                  End If
                  Call PlaySound(1)
              Else 'is gray
                 Call ChangeVellosity(Enemy(I).EnemyRect)
              End If
              'íÿìà ñìèñúë äà ïðîâåðÿâà äðóãèòå
              Exit For
           End If 'Collision(Enemy(I).EnemyRect, BallRect) = True
      End If
    Next I
    Call DrawBitmap(picBall, picBallMask.hDC, picGame.hDC, BallRect.Left, BallRect.Top)
End Sub

Private Sub mnuNewGame_Click()
    Call SetStartPosition
    
    Score = 0
    lblScore.Caption = Score
    LiveNumber = 50
    lblLives.Caption = LiveNumber
    lblLevelText.Visible = True
    
    tmrMove.Enabled = True

    picGame.SetFocus
    LevelNumber = 1
    Call LoadLevels(LevelNumber)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload frmEditor
End Sub

Private Sub lblExit_Click()
    Unload Me
End Sub

Private Sub lblLevelEditor_Click()
    frmEditor.Show
End Sub

Private Sub lblNewGame_Click()
    Call mnuNewGame_Click
End Sub

Private Sub tmrMove_Timer()
    If StartGame = True Then
        Call GameBackGroungRefresh
        Call PlayerMove
        Call BallMove
        Call DrawEnemis
        If BonusCount > 0 Then
            Call BonusMove
        End If
        picGame.Refresh
    Else
        Call GameBackGroungRefresh
        Call PlayerMove
        Call MoveBallWithPlayer
        Call DrawEnemis
        picGame.Refresh
    End If
End Sub

Private Sub MoveBallWithPlayer()
    BallRect.Left = PlayerRect.Left + PlayerRect.Width / 2 - BallRect.Width / 2
    BallRect.Top = PlayerRect.Top - BallRect.Height - 1
    Call DrawBitmap(picBall, picBallMask.hDC, picGame.hDC, BallRect.Left, BallRect.Top)
End Sub

Private Sub LoadLevels(LevelNum As Integer)
Dim intFileNum As Integer
Dim I As Integer
Dim sPath As String
Dim s As String

    On Error GoTo LoadLevelsError
    sPath = App.Path
    If Right(sPath, 1) <> "\" Then
       sPath = sPath & "\"
    End If
    
    sPath = sPath & "Levels\Level" & LevelNum & ".txt"
    intFileNum = FreeFile
    
    Open sPath For Input As #intFileNum
        Line Input #intFileNum, s
        EnemyCount = Val(s)
        ReDim Enemy(EnemyCount - 1)
        I = 0
        Do Until EOF(intFileNum)
           Line Input #intFileNum, s
           Enemy(I).EnemyRect.Left = Val(Mid(s, 1, 4))
           Enemy(I).EnemyRect.Top = Val(Mid(s, 6, 4))
           Enemy(I).EnemyRect.Width = ENEMY_WIDTH
           Enemy(I).EnemyRect.Height = ENEMY_HEIGHT
           Enemy(I).Alive = True
           Enemy(I).Bonus = Mid(s, 11, 1)
           Enemy(I).Color = Mid(s, 13, 1)
           I = I + 1
        Loop
    Close #intFileNum
    lblLevel.Caption = LevelNum
    Exit Sub
LoadLevelsError:
    If Err.Number = 53 And LevelNum > 1 Then 'File not found
       MsgBox "Level " & LevelNum & " not found", vbCritical, "Arkanoid"
       LevelNumber = 1
       Call LoadLevels(LevelNumber)
    Else
       MsgBox Err.Description, vbCritical, "Arkanoid"
       Close #intFileNum
       If LevelNum = 1 Then Unload Me
    End If
End Sub

Public Sub GameBackGroungRefresh()
Dim lngResponce As Long
    lngResponce = BitBlt(picGame.hDC, 0, 0, FormWidth, FormHeight, picBackGround.hDC, 0, 0, vbSrcCopy)
End Sub

Private Sub SetStartPosition()
    Set picPlayer.Picture = LoadResPicture("Player", 0)
    Set picPlayerMask.Picture = LoadResPicture("PlayerMask", 0)
    
    PlayerRect.Width = picPlayer.ScaleWidth
    PlayerRect.Height = picPlayer.ScaleHeight
    PlayerRect.Left = picGame.ScaleWidth / 2 - PlayerRect.Width / 2
    PlayerRect.Top = picGame.ScaleHeight - PlayerRect.Height - 5
    
    BallRect.Left = PlayerRect.Left + PlayerRect.Width / 2 - BallRect.Width / 2
    BallRect.Top = PlayerRect.Top - BallRect.Height
    BallVelocityX = 10
    BallVelocityY = 10
    
    BonusCount = 0
    picBall.Tag = ""
    StartGame = False
End Sub

Private Sub DrawEnemis()
Dim I As Integer
Dim MisionComplite As Boolean
    MisionComplite = True
    For I = 0 To EnemyCount - 1
        If Enemy(I).Alive = True Then
           Select Case Enemy(I).Color
                  Case Red: Call DrawBitmap(picEnemy, picEnemyMask.hDC, picGame.hDC, Enemy(I).EnemyRect.Left, Enemy(I).EnemyRect.Top)
                            MisionComplite = False
                  Case Blue: Call DrawBitmap(picEnemyBlue, picEnemyMask.hDC, picGame.hDC, Enemy(I).EnemyRect.Left, Enemy(I).EnemyRect.Top)
                             MisionComplite = False
                  Case Yellow: Call DrawBitmap(picEnemyYellow, picEnemyMask.hDC, picGame.hDC, Enemy(I).EnemyRect.Left, Enemy(I).EnemyRect.Top)
                               MisionComplite = False
                  Case Grey: Call DrawBitmap(picEnemyGrey, picEnemyMask.hDC, picGame.hDC, Enemy(I).EnemyRect.Left, Enemy(I).EnemyRect.Top)
           End Select
        End If
    Next I
    If MisionComplite = True Then
       PlayerDirection.MoveLeft = False
       PlayerDirection.MoveRight = False
       Call SetStartPosition
       LevelNumber = LevelNumber + 1
       Call LoadLevels(LevelNumber)
    End If
End Sub

Private Sub BonusMove()
Dim I As Integer
Dim MoreBonuses As Boolean
    MoreBonuses = False
    For I = 1 To BonusCount
        If Bonuses(I).BonusNotExist = False Then
           Bonuses(I).BonusRect.Top = Bonuses(I).BonusRect.Top + BonusStep
           If Collision(Bonuses(I).BonusRect, PlayerRect) = True Then
              Bonuses(I).BonusNotExist = True
              Call CheckBonusType(Bonuses(I).BonusType)
           Else
              MoreBonuses = True
              Call DrawBitmap(picBonus, picBonusMask.hDC, picGame.hDC, Bonuses(I).BonusRect.Left, Bonuses(I).BonusRect.Top)
           End If
        End If
    Next I
    If MoreBonuses = False Then
       BonusCount = 0
    End If
End Sub

Private Sub CheckBonusType(ByVal BonusType As BonusEnum)
    'Ïúðâî âðúùàìå êúì íà÷àëíîòî ñúñòîÿíèå
    Set picPlayer.Picture = LoadResPicture("Player", 0)
    Set picPlayerMask.Picture = LoadResPicture("PlayerMask", 0)
    PlayerRect.Width = picPlayer.ScaleWidth
    PlayerRect.Height = picPlayer.ScaleHeight
    
    If picBall.Tag = "fast" Then
       BallVelocityX = BallVelocityX / 2
       BallVelocityY = BallVelocityY / 2
    End If
    picBall.Tag = ""
    
    Select Case BonusType
        Case Large: Set picPlayer.Picture = LoadResPicture("PlayerLarge", 0)
                    Set picPlayerMask.Picture = LoadResPicture("PlayerlargeMask", 0)
                    PlayerRect.Width = picPlayer.ScaleWidth
                    PlayerRect.Height = picPlayer.ScaleHeight
        Case Live: LiveNumber = LiveNumber + 1
                   lblLives.Caption = LiveNumber
        Case FireBall: BallVelocityX = BallVelocityX * 2
                       BallVelocityY = BallVelocityY * 2
                       picBall.Tag = "fast"
    End Select
End Sub

Private Sub ChangeVellosity(EnemyRect As Rect)
Dim OldX As Long
Dim OldY As Long
Dim J As Long
   'íÿìà çíà÷åíèå îò êúäå èäâà òîï÷åòî
   OldY = BallRect.Top - BallVelocityY
   OldX = BallRect.Left - BallVelocityX
   
    If BallVelocityY > 0 Then
       If BallVelocityX > 0 Then
          If EnemyRect.Left - OldX <= EnemyRect.Top - OldY Then 'y>0 x>0
             BallVelocityY = -BallVelocityY
             BallRect.Top = EnemyRect.Top - BallRect.Height - 1
             For J = 1 To EnemyCount - 1
                 If Enemy(J).Alive = True Then
                    If Collision(BallRect, Enemy(J).EnemyRect) = True Then
                       BallRect.Left = EnemyRect.Left - BallRect.Width - 1
                       BallRect.Top = EnemyRect.Top
                       BallVelocityX = -BallVelocityX
                       BallVelocityY = -BallVelocityY
                       Exit For
                    End If
                 End If
             Next J
          Else
             BallVelocityX = -BallVelocityX
             BallRect.Left = EnemyRect.Left - BallRect.Width - 1
             For J = 1 To EnemyCount - 1
                 If Enemy(J).Alive = True Then
                    If Collision(BallRect, Enemy(J).EnemyRect) = True Then
                       BallRect.Left = EnemyRect.Left
                       BallRect.Top = EnemyRect.Top - BallRect.Height - 1
                       BallVelocityX = -BallVelocityX
                       BallVelocityY = -BallVelocityY
                       Exit For
                    End If
                 End If
             Next J
          End If
       Else
          If OldX - (EnemyRect.Left + EnemyRect.Width) <= EnemyRect.Top - OldY Then 'y>0 x<0
             BallVelocityY = -BallVelocityY
             BallRect.Top = EnemyRect.Top - BallRect.Height - 1
             For J = 1 To EnemyCount - 1
                 If Enemy(J).Alive = True Then
                    If Collision(BallRect, Enemy(J).EnemyRect) = True Then
                       BallRect.Left = EnemyRect.Left + EnemyRect.Width + 1
                       BallRect.Top = EnemyRect.Top
                       BallVelocityX = -BallVelocityX
                       BallVelocityY = -BallVelocityY
                       Exit For
                    End If
                 End If
             Next J
          Else
             BallVelocityX = -BallVelocityX
             BallRect.Left = EnemyRect.Left + EnemyRect.Width + 1
             For J = 1 To EnemyCount - 1
                 If Enemy(J).Alive = True Then
                    If Collision(BallRect, Enemy(J).EnemyRect) = True Then
                       BallRect.Left = EnemyRect.Left + EnemyRect.Width - BallRect.Width
                       BallRect.Top = EnemyRect.Top - BallRect.Height - 1
                       BallVelocityX = -BallVelocityX
                       BallVelocityY = -BallVelocityY
                       Exit For
                    End If
                 End If
             Next J
          End If
       End If 'BallVelocityX>0
    Else
       If BallVelocityX > 0 Then
          If EnemyRect.Left - OldX <= OldY - (EnemyRect.Top + EnemyRect.Height) Then 'y<0 x>0
             BallVelocityY = -BallVelocityY
             BallRect.Top = EnemyRect.Top + EnemyRect.Height + 1
             For J = 1 To EnemyCount - 1
                 If Enemy(J).Alive = True Then
                    If Collision(BallRect, Enemy(J).EnemyRect) = True Then
                       BallRect.Left = EnemyRect.Left - BallRect.Width - 1
                       BallRect.Top = EnemyRect.Top + EnemyRect.Height - BallRect.Height
                       BallVelocityX = -BallVelocityX
                       BallVelocityY = -BallVelocityY
                       Exit For
                    End If
                 End If
             Next J
          Else
             BallVelocityX = -BallVelocityX
             BallRect.Left = EnemyRect.Left - BallRect.Width - 1
             For J = 1 To EnemyCount - 1
                 If Enemy(J).Alive = True Then
                    If Collision(BallRect, Enemy(J).EnemyRect) = True Then
                        BallRect.Left = EnemyRect.Left
                        BallRect.Top = EnemyRect.Top + EnemyRect.Height + 1
                        BallVelocityX = -BallVelocityX
                        BallVelocityY = -BallVelocityY
                        Exit For
                    End If
                 End If
             Next J
          End If
       Else
          If OldX - (EnemyRect.Left + EnemyRect.Width) <= OldY - (EnemyRect.Top + EnemyRect.Height) Then 'y<0 x<0
             BallVelocityY = -BallVelocityY
             BallRect.Top = EnemyRect.Top + EnemyRect.Height + 1
             For J = 1 To EnemyCount - 1
                 If Enemy(J).Alive = True Then
                    If Collision(BallRect, Enemy(J).EnemyRect) = True Then
                       BallRect.Left = EnemyRect.Left + EnemyRect.Width + 1
                       BallRect.Top = EnemyRect.Top + EnemyRect.Height - BallRect.Height
                       BallVelocityX = -BallVelocityX
                       BallVelocityY = -BallVelocityY
                       Exit For
                    End If
                 End If
             Next J
          Else
             BallVelocityX = -BallVelocityX
             BallRect.Left = EnemyRect.Left + EnemyRect.Width + 1
             For J = 1 To EnemyCount - 1
                 If Enemy(J).Alive = True Then
                    If Collision(BallRect, Enemy(J).EnemyRect) = True Then
                       BallRect.Left = EnemyRect.Left + EnemyRect.Width - BallRect.Width
                       BallRect.Top = EnemyRect.Top + EnemyRect.Height + 1
                       BallVelocityX = -BallVelocityX
                       BallVelocityY = -BallVelocityY
                       Exit For
                    End If
                 End If
             Next J
          End If
       End If 'BallVelocityX>0
    End If 'BallVelocityY>0
End Sub
