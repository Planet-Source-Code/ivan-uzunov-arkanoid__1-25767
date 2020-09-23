Attribute VB_Name = "Module1"
Option Explicit

Public Type PlayerMoveDirection
    MoveLeft As Boolean
    MoveRight As Boolean
End Type

Public Type Rect
    Left As Long
    Top As Long
    Width As Long
    Height As Long
End Type

Public Type EnemyType
    Bonus As BonusEnum
    Alive As Boolean
    EnemyRect As Rect
    Color As EnemyColor
End Type

Public Enum EnemyColor
    Red = 0
    Blue = 1
    Yellow = 2
    Grey = 3
End Enum

Public Enum BonusEnum
    None = 0
    Large = 1
    Live = 2
    FireBall = 3
    'Glue = 2
End Enum

Public Type TypeBonus
    BonusRect As Rect
    BonusType As BonusEnum
    BonusNotExist As Boolean
End Type

Public Const PlayerStep = 15
Public Const BonusStep = 5

Public PlayerDirection As PlayerMoveDirection
Public PlayerRect As Rect
Public BallRect As Rect
Public Enemy() As EnemyType
Public EnemyCount As Integer
Public Bonuses() As TypeBonus
Public BonusCount As Integer

Public Function Collision(Rect1 As Rect, Rect2 As Rect) As Boolean
    Collision = False
    If Rect1.Left + Rect1.Width > Rect2.Left Then
        If Rect1.Left < Rect2.Left + Rect2.Width Then
            If Rect1.Top + Rect1.Height > Rect2.Top Then
                If Rect1.Top < Rect2.Top + Rect2.Height Then
                    Collision = True
                End If
            End If
        End If
    End If
End Function

Public Function BallPlayerCollision(BallRect1 As Rect, PlayerRect2 As Rect) As Boolean
    BallPlayerCollision = False
    If BallRect1.Top + BallRect1.Height + 1 > PlayerRect2.Top Then
        If BallRect1.Left + BallRect1.Width > PlayerRect2.Left Then
            If BallRect1.Left < PlayerRect2.Left + PlayerRect2.Width Then
               BallPlayerCollision = True
            End If
        End If
    End If
End Function

