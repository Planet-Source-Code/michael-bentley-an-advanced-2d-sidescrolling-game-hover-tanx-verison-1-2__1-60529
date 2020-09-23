VERSION 5.00
Begin VB.Form frmEditor 
   BackColor       =   &H00FFFF00&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Hover Tanx Editor"
   ClientHeight    =   6225
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10500
   Icon            =   "frmEditor.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   415
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   700
   StartUpPosition =   3  'Windows Default
   Begin VB.HScrollBar hScroll 
      Height          =   255
      LargeChange     =   3
      Left            =   0
      Max             =   10
      TabIndex        =   0
      Top             =   6000
      Width           =   10575
   End
   Begin VB.Image imgSprite 
      Height          =   375
      Index           =   0
      Left            =   480
      Top             =   480
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Shape shpGround 
      FillColor       =   &H0080FFFF&
      FillStyle       =   0  'Solid
      Height          =   1815
      Left            =   -75
      Top             =   4800
      Width           =   10650
   End
End
Attribute VB_Name = "frmEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
frmOptions.Show
frmTank.Show
For i = 1 To 100
    Load imgSprite(i)
    imgSprite(i).Visible = False
    SpriteVisible(i) = False
Next 'i
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'User has clicked the screen indicating they want to place a sprite at the location

Dim curSprite As Long
If Y >= 340 Then 'You're trying to place a sprite above the bottom limit of the map
    Exit Sub
End If

curSprite = FindImage 'Get an empty sprite

If curSprite = -1 Then Exit Sub 'No sprite found, don't make a new one

'Set the new sprite to being visible, set the picture to being the current picture
imgSprite(curSprite).Visible = True
imgSprite(curSprite).Picture = frmOptions.imgSprite(frmOptions.hSprite.Value).Picture
SpriteVisible(curSprite) = True
SpriteType(curSprite) = frmOptions.hSprite.Value
SpriteX(curSprite) = (Me.ScaleWidth * hScroll.Value) + X 'Actual X is the mouse X plus the width of screens before it

'The "floor" button is not pressed
If frmOptions.chkBottom.Value = 0 And frmOptions.chkWater.Value = 0 And frmOptions.chkTop.Value = 0 Then
    SpriteY(curSprite) = Y
    imgSprite(curSprite).Top = Y
ElseIf frmOptions.chkBottom.Value = 1 Then 'Align sprite with bottom of screen
    SpriteY(curSprite) = shpGround.Top - imgSprite(curSprite).Height
    imgSprite(curSprite).Top = SpriteY(curSprite)
ElseIf frmOptions.chkWater.Value = 1 Then 'align sprite with water level
    SpriteY(curSprite) = 300 - imgSprite(curSprite).Height
    imgSprite(curSprite).Top = SpriteY(curSprite)
ElseIf frmOptions.chkTop.Value = 1 Then 'Align sprite with top of the screen
    SpriteY(curSprite) = 0
    imgSprite(curSprite).Top = SpriteY(curSprite)
End If
imgSprite(curSprite).Left = X
End Sub
Private Function FindImage() As Long
'Finds an empty sprite, or -1 if none exists
For i = 1 To 100
    If SpriteVisible(i) = False Then
        FindImage = i
        Exit Function
    End If
Next 'i
'Nothing has been found
FindImage = -1
End Function

Private Sub hScroll_Change()
'Figure out which sprites appear on this particular screen
For i = 1 To 100
    imgSprite(i).Visible = False
    If SpriteVisible(i) = True Then 'The sprite has already been placed
        'If the sprite is on this screen, set it to visible
        If SpriteX(i) >= (Me.ScaleWidth * hScroll.Value) And SpriteX(i) <= (Me.ScaleWidth * (hScroll.Value + 1)) Then
            imgSprite(i).Visible = True
            imgSprite(i).Left = SpriteX(i) - (Me.ScaleWidth * hScroll.Value) 'Set the correct x position
        End If
    End If
Next 'i
End Sub

Private Sub hScroll_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeySpace Then 'Debug feature that gets rid of the foreground temporarily
    If shpGround.Visible = True Then
        shpGround.Visible = False
        For i = 1 To 100
            If imgSprite(i).Visible = True Then
                If imgSprite(i).Top > shpGround.Top + 20 Then
                    imgSprite(i).Top = 200
                End If
                If imgSprite(i).Top + imgSprite(i).Height < -10 Then
                    imgSprite(i).Top = 50
                End If
            End If
        Next 'i
    Else
        shpGround.Visible = True
    End If
End If
End Sub

Private Sub imgSprite_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then 'Deletes the current sprite with a right click of it
    imgSprite(Index).Visible = False
    SpriteVisible(Index) = False
End If
End Sub
