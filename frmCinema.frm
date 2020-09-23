VERSION 5.00
Begin VB.Form frmCinema 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Hover Tanx"
   ClientHeight    =   4500
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7500
   Icon            =   "frmCinema.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   300
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   500
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picCinema 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   3750
      Left            =   0
      Picture         =   "frmCinema.frx":0D2A
      ScaleHeight     =   250
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   500
      TabIndex        =   0
      Top             =   375
      Width           =   7500
   End
End
Attribute VB_Name = "frmCinema"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim bRunning As Boolean 'If the loop is running
Dim TickDif As Double 'How fast the loop runs
Dim LastTick As Double 'Last tick to be executed
Dim intDelay As Long 'Delay between frames
Dim intAnim As Long 'Current frame
Private Sub Image1_Click()

End Sub

Private Sub Form_Activate()
On Error Resume Next
If intAnim = 0 Then intAnim = 1 'Set the frame to 1
bRunning = True 'Set the loop to running
Call GameLoop 'Start the loop
End Sub

Private Sub Form_Load()
On Error Resume Next
intDelay = 0 'Set cycles through the loop to 0
TickDif = 0.05 'Set speed of loop
intAnim = 1 'Set animation frame
End Sub

Private Sub Form_Unload(Cancel As Integer)
'Unload this form and load the briefing screen depending on what's next
On Error Resume Next
'Reset loop values
bRunning = False
intAnim = 1
intDelay = 0
Select Case anType
    Case 0 'Load the movie screen again with the desert boss cinemeatic
        If firstCinematic = False Then
            anType = 1
            intAnim = 1
            Call Form_Load
            bRunning = True
            Call GameLoop
        Else
            frmIntro.Show
        End If
    Case 1 'Otherwise, just load the next briefing screen depending on level
        NextLevel = "level1.ini"
        frmBriefing.Show
    Case 2
        NextLevel = "level3.ini"
        frmBriefing.Show
    Case 3
        NextLevel = "level5.ini"
        frmBriefing.Show
    Case 4
        NextLevel = "level7.ini"
        frmBriefing.Show
    Case 5
        NextLevel = "level9.ini"
        frmBriefing.Show
    Case 6
        NextLevel = "level11.ini"
        frmBriefing.Show
    Case 7
        frmCredits.Show
End Select
End Sub

Private Sub Animation0()
'Draws the animation
On Error Resume Next
intDelay = intDelay + 1
Select Case intDelay
Case 2
    curMidi = "prologue.mid"
    PlayMidi (curMidi)
Case 60
    intAnim = 2
Case 62
    intAnim = 3
Case 64
    intAnim = 4
Case 126
    intAnim = 5
Case 128
    intAnim = 6
Case 130
    intAnim = 7
Case 132
    intAnim = 8
Case 140
    intAnim = 9
Case 200
    intAnim = 10
Case 201
    intAnim = 11
Case 202
    intAnim = 12
Case 203
    intAnim = 13
Case 204
    intAnim = 14
Case 205
    intAnim = 15
Case 206
    intAnim = 16
Case 207
    intAnim = 17
Case 208
    intAnim = 18
Case 209
    intAnim = 19
Case 210
    intAnim = 20
Case 211
    intAnim = 21
Case 250
    intAnim = 22
Case 254
    intAnim = 23
Case 258
    intAnim = 24
Case 262
    intAnim = 25
Case 280
    intAnim = 26
Case 420
    intAnim = 27
Case 550
    intAnim = 100
End Select
If intAnim = 100 Then
    If firstCinematic = False Then
        anType = 1
        intAnim = 1
        intDelay = 1
    Else
        Unload Me
    End If
    intAnim = 1
    Exit Sub
End If
If intAnim < 10 Then
    picCinema.Picture = LoadPicture(App.Path & "\Cinema\10" & CStr(intAnim) & ".jpg")
Else
    picCinema.Picture = LoadPicture(App.Path & "\Cinema\1" & CStr(intAnim) & ".jpg")
End If
End Sub
Private Sub GameLoop()
On Error Resume Next
Dim curFreq As Currency
Dim curStart As Currency
Dim curEnd As Currency
Dim dblResult As Double

QueryPerformanceFrequency curFreq       'Get the timer frequency
Do While bRunning = True
    QueryPerformanceCounter curStart        'Get the start time
    dblResult = (curStart - LastTick) / curFreq   'Calculate the duration (in seconds)
    If dblResult > TickDif Then
        If intAnim = 0 Then intAnim = 1
        LastTick = curStart
        'Call the vairous cinematic method
        Select Case anType
            Case 0
                Call Animation0
            Case 1
                Call Animation1
            Case 2
                Call Animation2
            Case 3
                Call Animation3
            Case 4
                Call Animation4
            Case 5
                Call Animation5
            Case 6
                Call Animation6
            Case 7
                Call Animation7
        End Select
    End If
    DoEvents
Loop
End Sub
Private Sub Animation1()
On Error Resume Next
intDelay = intDelay + 1
Select Case intDelay
Case 2
    curMidi = "mov01.mid"
    PlayMidi (curMidi)
Case 60
    intAnim = 2
Case 150
    intAnim = 3
Case 250
    intAnim = 4
Case 350
    intAnim = 5
Case 425
    intAnim = 6
Case 429
    intAnim = 7
Case 455
    intAnim = 8
Case 457
    intAnim = 9
Case 520
    intAnim = 10
Case 620
    intAnim = 11
Case 700
    intAnim = 12
Case 800
    intAnim = 13
Case 900
    intAnim = 14
Case 950
    intAnim = 15
Case 954
    intAnim = 16
Case 958
    intAnim = 17
Case 1100
    intAnim = 100
End Select
If intAnim = 100 Then
    bRunning = False
    Unload Me
    Exit Sub
End If
If intAnim < 10 Then
    picCinema.Picture = LoadPicture(App.Path & "\Cinema\20" & CStr(intAnim) & ".jpg")
Else
    picCinema.Picture = LoadPicture(App.Path & "\Cinema\2" & CStr(intAnim) & ".jpg")
End If
End Sub
Private Sub Animation2()
On Error Resume Next
intDelay = intDelay + 1
Select Case intDelay
Case 2
    curMidi = "mov02.mid"
    PlayMidi (curMidi)
Case 80
    intAnim = 2
Case 85
    intAnim = 3
Case 90
    intAnim = 4
Case 95
    intAnim = 5
Case 100
    intAnim = 6
Case 105
    intAnim = 7
Case 110
    intAnim = 8
Case 115
    intAnim = 9
Case 120
    intAnim = 10
Case 140
    intAnim = 11
Case 145
    intAnim = 12
Case 150
    intAnim = 13
Case 155
    intAnim = 14
Case 160
    intAnim = 15
Case 280
    intAnim = 16
Case 284
    intAnim = 17
Case 288
    intAnim = 18
Case 292
    intAnim = 19
Case 296
    intAnim = 20
Case 300
    intAnim = 21
Case 350
    intAnim = 100
End Select
If intAnim = 100 Then
    bRunning = False
    Unload Me
    Exit Sub
End If
If intAnim < 10 Then
    picCinema.Picture = LoadPicture(App.Path & "\Cinema\30" & CStr(intAnim) & ".jpg")
Else
    picCinema.Picture = LoadPicture(App.Path & "\Cinema\3" & CStr(intAnim) & ".jpg")
End If
End Sub
Private Sub Animation3()
On Error Resume Next
intDelay = intDelay + 1
Select Case intDelay
Case 2
    curMidi = "mov03.mid"
    PlayMidi (curMidi)
Case 60
    intAnim = 2
Case 160
    intAnim = 3
Case 165
    intAnim = 4
Case 170
    intAnim = 5
Case 175
    intAnim = 6
Case 210
    intAnim = 7
Case 213
    intAnim = 8
Case 216
    intAnim = 9
Case 219
    intAnim = 10
Case 222
    intAnim = 11
Case 225
    intAnim = 12
Case 227
    intAnim = 13
Case 230
    intAnim = 14
Case 233
    intAnim = 15
Case 236
    intAnim = 16
Case 239
    intAnim = 17
Case 242
    intAnim = 18
Case 270
    intAnim = 19
Case 274
    intAnim = 20
Case 275
    intAnim = 21
Case 300
    intAnim = 100
End Select
If intAnim = 100 Then
    bRunning = False
    Unload Me
    Exit Sub
End If
If intAnim < 10 Then
    picCinema.Picture = LoadPicture(App.Path & "\Cinema\40" & CStr(intAnim) & ".jpg")
Else
    picCinema.Picture = LoadPicture(App.Path & "\Cinema\4" & CStr(intAnim) & ".jpg")
End If
End Sub
Private Sub Animation4()
On Error Resume Next
intDelay = intDelay + 1
Select Case intDelay
Case 2
    curMidi = "mov04.mid"
    PlayMidi (curMidi)
Case 100
    intAnim = 2
Case 105
    intAnim = 3
Case 105
    intAnim = 4
Case 110
    intAnim = 5
Case 115
    intAnim = 6
Case 120
    intAnim = 7
Case 125
    intAnim = 8
Case 130
    intAnim = 9
Case 135
    intAnim = 10
Case 140
    intAnim = 11
Case 145
    intAnim = 12
Case 180
    intAnim = 13
Case 183
    intAnim = 14
Case 183
    intAnim = 15
Case 188
    intAnim = 16
Case 189
    intAnim = 17
Case 192
    intAnim = 18
Case 195
    intAnim = 19
Case 198
    intAnim = 20
Case 201
    intAnim = 21
Case 204
    intAnim = 22
Case 207
    intAnim = 23
Case 210
    intAnim = 24
Case 213
    intAnim = 25
Case 250
    intAnim = 100
End Select
If intAnim = 100 Then
    bRunning = False
    Unload Me
    Exit Sub
End If
If intAnim < 10 Then
    picCinema.Picture = LoadPicture(App.Path & "\Cinema\50" & CStr(intAnim) & ".jpg")
Else
    picCinema.Picture = LoadPicture(App.Path & "\Cinema\5" & CStr(intAnim) & ".jpg")
End If

End Sub
Private Sub Animation5()
On Error Resume Next
intDelay = intDelay + 1
Select Case intDelay
Case 2
    curMidi = "mov05.mid"
    PlayMidi (curMidi)
Case 50
    intAnim = 2
Case 54
    intAnim = 3
Case 58
    intAnim = 4
Case 62
    intAnim = 5
Case 66
    intAnim = 6
Case 70
    intAnim = 7
Case 74
    intAnim = 8
Case 78
    intAnim = 9
Case 82
    intAnim = 10
Case 86
    intAnim = 11
Case 90
    intAnim = 12
Case 94
    intAnim = 13
Case 98
    intAnim = 14
Case 102
    intAnim = 15
Case 106
    intAnim = 16
Case 110
    intAnim = 17
Case 114
    intAnim = 18
Case 118
    intAnim = 19
Case 122
    intAnim = 20
Case 126
    intAnim = 21
Case 130
    intAnim = 22
Case 134
    intAnim = 23
Case 180
    intAnim = 100
End Select
If intAnim = 100 Then
    bRunning = False
    Unload Me
    Exit Sub
End If
If intAnim < 10 Then
    picCinema.Picture = LoadPicture(App.Path & "\Cinema\60" & CStr(intAnim) & ".jpg")
Else
    picCinema.Picture = LoadPicture(App.Path & "\Cinema\6" & CStr(intAnim) & ".jpg")
End If

End Sub
Private Sub Animation6()
On Error Resume Next
intDelay = intDelay + 1
Select Case intDelay
Case 2
    curMidi = "mov06.mid"
    PlayMidi (curMidi)
Case 50
    intAnim = 2
Case 54
    intAnim = 3
Case 58
    intAnim = 4
Case 62
    intAnim = 5
Case 66
    intAnim = 6
Case 70
    intAnim = 7
Case 74
    intAnim = 8
Case 78
    intAnim = 9
Case 82
    intAnim = 10
Case 86
    intAnim = 11
Case 115
    intAnim = 12
Case 120
    intAnim = 13
Case 150
    intAnim = 14
Case 152
    intAnim = 15
Case 154
    intAnim = 16
Case 175
    intAnim = 17
Case 200
    intAnim = 18
Case 205
    intAnim = 19
Case 210
    intAnim = 20
Case 240
    intAnim = 21
Case 243
    intAnim = 22
Case 246
    intAnim = 23
Case 249
    intAnim = 24
Case 252
    intAnim = 25
Case 255
    intAnim = 26
Case 258
    intAnim = 27
Case 261
    intAnim = 28
Case 264
    intAnim = 29
Case 267
    intAnim = 30
Case 270
    intAnim = 31
Case 273
    intAnim = 32
Case 375
    intAnim = 100
End Select
If intAnim = 100 Then
    bRunning = False
    Unload Me
    Exit Sub
End If
If intAnim < 10 Then
    picCinema.Picture = LoadPicture(App.Path & "\Cinema\70" & CStr(intAnim) & ".jpg")
Else
    picCinema.Picture = LoadPicture(App.Path & "\Cinema\7" & CStr(intAnim) & ".jpg")
End If
End Sub
Private Sub Animation7()
On Error Resume Next
intDelay = intDelay + 1
Select Case intDelay
Case 2
    curMidi = "ending.mid"
    PlayMidi (curMidi)
Case 50
    intAnim = 2
Case 54
    intAnim = 3
Case 58
    intAnim = 4
Case 62
    intAnim = 5
Case 66
    intAnim = 6
Case 90
    intAnim = 7
Case 93
    intAnim = 8
Case 96
    intAnim = 9
Case 99
    intAnim = 10
Case 130
    intAnim = 11
Case 150
    intAnim = 12
Case 154
    intAnim = 13
Case 158
    intAnim = 14
Case 162
    intAnim = 15
Case 166
    intAnim = 16
Case 170
    intAnim = 17
Case 174
    intAnim = 18
Case 178
    intAnim = 19
Case 200
    intAnim = 20
Case 450
    intAnim = 21
Case 454
    intAnim = 22
Case 458
    intAnim = 23
Case 462
    intAnim = 24
Case 466
    intAnim = 25
Case 468
    intAnim = 26
Case 472
    intAnim = 27
Case 476
    intAnim = 28
Case 480
    intAnim = 29
Case 484
    intAnim = 30
Case 488
    intAnim = 31
Case 492
    intAnim = 32
Case 496
    intAnim = 33
Case 510
    intAnim = 34
Case 520
    intAnim = 35
Case 630
    intAnim = 100
End Select
If intAnim = 100 Then
    bRunning = False
    Unload Me
    Exit Sub
End If
If intAnim < 10 Then
    picCinema.Picture = LoadPicture(App.Path & "\Cinema\80" & CStr(intAnim) & ".jpg")
Else
    picCinema.Picture = LoadPicture(App.Path & "\Cinema\8" & CStr(intAnim) & ".jpg")
End If

End Sub

