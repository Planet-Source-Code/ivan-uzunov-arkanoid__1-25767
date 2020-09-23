Attribute VB_Name = "Module2"
Option Explicit

Const SND_ASYNC = &H1            'Play asynchronously
Const SND_NODEFAULT = &H2        'Don't use default sound
Const SND_MEMORY = &H4           'lpszSoundName points to a memory file
Const SND_LOOP = &H8
Const SND_NOSTOP = &H10
 
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long

Public Sub DrawBitmap(PrimaryPicture As PictureBox, MaskPictureHDC As Long, BackGroungPictureHDC As Long, ByVal LeftPos As Long, ByVal TopPos As Long)
'Draws a transparent bitmap
Dim lngReturn As Long
    'Ïðèëàãàíå íà ìàñêàòà
    lngReturn = BitBlt(BackGroungPictureHDC, LeftPos, TopPos, PrimaryPicture.ScaleWidth, PrimaryPicture.ScaleHeight, MaskPictureHDC, 0, 0, vbSrcAnd)
    'Èç÷åðòàâàíå íà ÍËÎ
    lngReturn = BitBlt(BackGroungPictureHDC, LeftPos, TopPos, PrimaryPicture.ScaleWidth, PrimaryPicture.ScaleHeight, PrimaryPicture.hDC, 0, 0, vbSrcPaint)
End Sub

Public Sub PlaySound(ByVal Index As Long)
'Play wave file from Resource
Dim sSoundBuffer As String
Dim Ret As Long
    sSoundBuffer = StrConv(LoadResData(Index, "SOUND"), vbUnicode)
    Ret = sndPlaySound(sSoundBuffer, SND_ASYNC Or SND_NODEFAULT Or SND_MEMORY Or SND_NOSTOP)
End Sub
