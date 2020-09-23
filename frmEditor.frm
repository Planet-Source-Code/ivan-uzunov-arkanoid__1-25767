VERSION 5.00
Begin VB.Form frmEditor 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Level Editor"
   ClientHeight    =   4425
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9870
   Icon            =   "frmEditor.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4425
   ScaleWidth      =   9870
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdClear 
      Caption         =   "&Clear"
      Height          =   375
      Left            =   4200
      TabIndex        =   6
      Top             =   120
      Width           =   855
   End
   Begin VB.CommandButton cmdOpen 
      Caption         =   "&Open"
      Height          =   375
      Left            =   2280
      TabIndex        =   5
      Top             =   120
      Width           =   855
   End
   Begin VB.TextBox txtLevelNum 
      Height          =   285
      Left            =   1560
      TabIndex        =   4
      Top             =   160
      Width           =   615
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Height          =   375
      Left            =   3240
      TabIndex        =   3
      Top             =   120
      Width           =   855
   End
   Begin VB.CommandButton cmdNewEnemy 
      Caption         =   "&New Enemy"
      Default         =   -1  'True
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1335
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
      Top             =   600
      Width           =   8160
      Begin VB.PictureBox picEnemy 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H0000FF00&
         BorderStyle     =   0  'None
         Height          =   405
         Index           =   0
         Left            =   0
         Picture         =   "frmEditor.frx":0442
         ScaleHeight     =   27
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   42
         TabIndex        =   2
         Tag             =   "0"
         Top             =   0
         Width           =   630
      End
   End
   Begin VB.Image imgColor 
      Height          =   405
      Index           =   3
      Left            =   7320
      Picture         =   "frmEditor.frx":1204
      Top             =   120
      Width           =   630
   End
   Begin VB.Image imgColor 
      Height          =   405
      Index           =   2
      Left            =   6600
      Picture         =   "frmEditor.frx":1FC6
      Top             =   120
      Width           =   630
   End
   Begin VB.Image imgColor 
      Height          =   405
      Index           =   0
      Left            =   5160
      Picture         =   "frmEditor.frx":2D88
      Top             =   120
      Width           =   630
   End
   Begin VB.Image imgColor 
      Height          =   405
      Index           =   1
      Left            =   5880
      Picture         =   "frmEditor.frx":3B4A
      Top             =   120
      Width           =   630
   End
   Begin VB.Menu mnuPopUp 
      Caption         =   "PopUp"
      Visible         =   0   'False
      Begin VB.Menu mnuBonus 
         Caption         =   "Bonus"
         Begin VB.Menu mnuBonusType 
            Caption         =   "None"
            Index           =   0
         End
         Begin VB.Menu mnuBonusType 
            Caption         =   "Large"
            Index           =   1
         End
         Begin VB.Menu mnuBonusType 
            Caption         =   "Live"
            Index           =   2
         End
         Begin VB.Menu mnuBonusType 
            Caption         =   "FireBall"
            Index           =   3
         End
      End
      Begin VB.Menu mnuColors 
         Caption         =   "Color"
         Begin VB.Menu mnuColor 
            Caption         =   "Red"
            Index           =   0
         End
         Begin VB.Menu mnuColor 
            Caption         =   "Blue"
            Index           =   1
         End
         Begin VB.Menu mnuColor 
            Caption         =   "Yellow"
            Index           =   2
         End
         Begin VB.Menu mnuColor 
            Caption         =   "Grey"
            Index           =   3
         End
      End
   End
End
Attribute VB_Name = "frmEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim EnemyMove As Boolean
Dim OldX As Single
Dim OldY As Single
Dim CurrentEnemy As Integer
Dim CurrColor As Integer

Private Sub cmdClear_Click()
Dim I As Integer
    For I = 1 To picEnemy.UBound
        Unload picEnemy(I)
    Next I
    BitBlt picGame.hDC, 0, 0, frmArquanoid.picGame.ScaleWidth, frmArquanoid.picGame.ScaleHeight, frmArquanoid.picBackGround.hDC, 0, 0, vbSrcCopy
    picGame.Refresh
End Sub

Private Sub cmdNewEnemy_Click()
    Load picEnemy(picEnemy.Count)
    picEnemy(picEnemy.UBound).Move 0, 0
    picEnemy(picEnemy.UBound).Tag = 0
    picEnemy(picEnemy.UBound).Visible = True
End Sub

Private Sub cmdOpen_Click()
Dim intFileNum As Integer
Dim I As Integer
Dim sPath As String
Dim s As String
    On Error GoTo cmdOpenError
    For I = 1 To picEnemy.UBound
        Unload picEnemy(I)
    Next I
    sPath = App.Path
    If Right(sPath, 1) <> "\" Then
       sPath = sPath & "\"
    End If
    
    sPath = sPath & "Levels\Level" & txtLevelNum.Text & ".txt"
    intFileNum = FreeFile
    
    Open sPath For Input As #intFileNum
        Line Input #intFileNum, s
        I = 0
        Do Until EOF(intFileNum)
           Line Input #intFileNum, s
           picEnemy(I).Left = Mid(s, 1, 4)
           picEnemy(I).Top = Mid(s, 6, 4)
           picEnemy(I).Tag = Mid(s, 11, 1)
           Set picEnemy(I).Picture = imgColor(Val(Mid(s, 13, 1))).Picture
           picEnemy(I).Visible = True
           I = I + 1
           Load picEnemy(I)
        Loop
    Close #intFileNum
    Exit Sub
cmdOpenError:
    MsgBox Err.Description, vbCritical, "Arkanoid"
    Close #intFileNum
End Sub

Private Sub cmdSave_Click()
Dim intFileNum As Integer
Dim I As Integer
Dim sPath As String
Dim s As String
    If IsNumeric(txtLevelNum.Text) = False Then
       MsgBox "Incorrect level number", vbCritical, "Level Editor"
       Exit Sub
    End If
    
    sPath = App.Path
    If Right(sPath, 1) <> "\" Then
       sPath = sPath & "\"
    End If
    sPath = sPath & "Levels\Level" & txtLevelNum.Text & ".txt"
    intFileNum = FreeFile
    Open sPath For Output As #intFileNum
        Print #intFileNum, picEnemy.Count
        For I = 0 To picEnemy.UBound
            Print #intFileNum, WriteLine(picEnemy(I))
        Next I
    Close #intFileNum
    
    txtLevelNum.Text = CStr(Val(txtLevelNum.Text) + 1)
End Sub

Private Sub Form_Load()
    BitBlt picGame.hDC, 0, 0, frmArquanoid.picGame.ScaleWidth, frmArquanoid.picGame.ScaleHeight, frmArquanoid.picBackGround.hDC, 0, 0, vbSrcCopy
    picGame.Refresh
    Call Form_Resize
    txtLevelNum.Text = "1"
End Sub

Private Sub Form_Resize()
    Me.Width = picGame.Width + 40
    Me.Height = picGame.Top + picGame.Height + 400
End Sub

Private Sub imgColor_Click(Index As Integer)
    CurrColor = Index
End Sub

Private Sub mnuBonusType_Click(Index As Integer)
    picEnemy(CurrentEnemy).Tag = Index
End Sub

Private Sub mnuColor_Click(Index As Integer)
    Set picEnemy(CurrentEnemy).Picture = imgColor(Index).Picture
End Sub

Private Sub picEnemy_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim I As Integer
    If vbLeftButton = Button Then
        EnemyMove = True
        OldX = X
        OldY = Y
    ElseIf Button = vbRightButton Then
        For I = 0 To mnuBonusType.UBound
            mnuBonusType(I).Checked = False
        Next I
        CurrentEnemy = Index
        mnuBonusType(Val(picEnemy(Index).Tag)).Checked = True
        PopupMenu mnuPopUp
    End If
End Sub

Private Sub picEnemy_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If vbLeftButton = Button And EnemyMove Then
        If Abs(X - OldX) >= 7 Then
           If X - OldX > 0 Then
              picEnemy(Index).Left = picEnemy(Index).Left + 7
           Else
              picEnemy(Index).Left = picEnemy(Index).Left - 7
           End If
        ElseIf Abs(Y - OldY) >= 9 Then
           If Y - OldY > 0 Then
              picEnemy(Index).Top = picEnemy(Index).Top + 9
           Else
              picEnemy(Index).Top = picEnemy(Index).Top - 9
           End If
        End If
    End If
End Sub

Private Sub picEnemy_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    EnemyMove = False
End Sub

Public Function WriteLine(Source As Control) As String
Dim s As String
    s = Source.Left
    WriteLine = CorrectLenght(s) & ","
    s = Source.Top
    WriteLine = WriteLine & CorrectLenght(s) & ","
    s = Source.Tag
    WriteLine = WriteLine & s & ","
    If Source.Picture = imgColor(0).Picture Then
       s = 0
    ElseIf Source.Picture = imgColor(1).Picture Then
       s = 1
    ElseIf Source.Picture = imgColor(2).Picture Then
       s = 2
    ElseIf Source.Picture = imgColor(3).Picture Then
       s = 3
    End If
    WriteLine = WriteLine & s
End Function

Private Function CorrectLenght(ByRef strS As String) As String
    Select Case Len(strS)
           Case 0: CorrectLenght = "0"
           Case 1: CorrectLenght = "000" & strS
           Case 2: CorrectLenght = "00" & strS
           Case 3: CorrectLenght = "0" & strS
           Case 4: CorrectLenght = strS
    End Select
End Function

Private Sub picGame_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim L As Single
Dim T As Single
    Load picEnemy(picEnemy.Count)
    picEnemy(picEnemy.UBound).Move (X \ 7) * 7, (Y \ 9) * 9
    picEnemy(picEnemy.UBound).Tag = 0
    picEnemy(picEnemy.UBound).Visible = True
    Set picEnemy(picEnemy.UBound).Picture = imgColor(CurrColor).Picture
End Sub
