VERSION 5.00
Begin VB.Form frmIntro 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Hover Tanx - Version 1.2"
   ClientHeight    =   4620
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8040
   Icon            =   "frmIntro.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   308
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   536
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer timeMidi 
      Interval        =   35
      Left            =   6840
      Top             =   4320
   End
   Begin VB.Timer timeDemo 
      Interval        =   40000
      Left            =   6720
      Top             =   0
   End
   Begin VB.Frame framProlouge 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   4695
      Left            =   0
      TabIndex        =   5
      Top             =   4500
      Width           =   8055
      Begin VB.Timer timePlot 
         Interval        =   50
         Left            =   480
         Top             =   1200
      End
      Begin VB.Label lblSkip 
         BackStyle       =   0  'Transparent
         Caption         =   "Skip"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   3720
         TabIndex        =   7
         Top             =   4200
         Width           =   375
      End
      Begin VB.Shape shpBar 
         FillStyle       =   0  'Solid
         Height          =   735
         Index           =   1
         Left            =   0
         Top             =   3960
         Width           =   8055
      End
      Begin VB.Shape shpBar 
         FillStyle       =   0  'Solid
         Height          =   735
         Index           =   0
         Left            =   0
         Top             =   120
         Width           =   8055
      End
      Begin VB.Label lblPlot 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmIntro.frx":0D2A
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2415
         Index           =   3
         Left            =   600
         TabIndex        =   10
         Top             =   960
         Width           =   7095
      End
      Begin VB.Label lblPlot 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmIntro.frx":0E93
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2415
         Index           =   2
         Left            =   600
         TabIndex        =   9
         Top             =   960
         Width           =   7095
      End
      Begin VB.Label lblPlot 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmIntro.frx":10A2
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2295
         Index           =   1
         Left            =   600
         TabIndex        =   8
         Top             =   960
         Width           =   7095
      End
      Begin VB.Label lblPlot 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmIntro.frx":1279
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2295
         Index           =   0
         Left            =   600
         TabIndex        =   6
         Top             =   960
         Width           =   7095
      End
   End
   Begin VB.CommandButton cmdHowToPlay 
      Caption         =   "&How To Play"
      Height          =   495
      Left            =   6840
      TabIndex        =   4
      Top             =   1080
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Timer timeIntro 
      Interval        =   35
      Left            =   5280
      Top             =   240
   End
   Begin VB.CommandButton cmdEditor 
      Caption         =   "&Editor"
      Height          =   495
      Left            =   5280
      TabIndex        =   3
      Top             =   3720
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "&Quit"
      Height          =   495
      Left            =   6840
      TabIndex        =   2
      Top             =   1560
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton cmdLoad 
      Caption         =   "&Load"
      Height          =   495
      Left            =   6840
      TabIndex        =   1
      Top             =   2040
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton cmdPlay 
      Caption         =   "&Camapaign"
      Height          =   495
      Left            =   6840
      TabIndex        =   0
      Top             =   600
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label lblQuit 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   375
      Left            =   1440
      MouseIcon       =   "frmIntro.frx":1424
      MousePointer    =   99  'Custom
      TabIndex        =   14
      Top             =   3840
      Visible         =   0   'False
      Width           =   4935
   End
   Begin VB.Label lblOptions 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Game Options"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   375
      Left            =   1440
      MouseIcon       =   "frmIntro.frx":1CEE
      MousePointer    =   99  'Custom
      TabIndex        =   13
      Top             =   3480
      Visible         =   0   'False
      Width           =   4935
   End
   Begin VB.Label lblHowToPlay 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "How To Play"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   375
      Left            =   1440
      MouseIcon       =   "frmIntro.frx":25B8
      MousePointer    =   99  'Custom
      TabIndex        =   12
      Top             =   3120
      Visible         =   0   'False
      Width           =   4935
   End
   Begin VB.Label lblCampaign 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Campaign"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   375
      Left            =   1440
      MouseIcon       =   "frmIntro.frx":2E82
      MousePointer    =   99  'Custom
      TabIndex        =   11
      Top             =   2760
      Visible         =   0   'False
      Width           =   4935
   End
   Begin VB.Image imgDocInc 
      Height          =   1500
      Left            =   0
      Picture         =   "frmIntro.frx":374C
      Top             =   1680
      Width           =   3750
   End
   Begin VB.Image imgPresents 
      Height          =   750
      Index           =   4
      Left            =   120
      Picture         =   "frmIntro.frx":3A6B
      Top             =   3600
      Visible         =   0   'False
      Width           =   3000
   End
   Begin VB.Image imgPresents 
      Height          =   750
      Index           =   3
      Left            =   4200
      Picture         =   "frmIntro.frx":3E52
      Top             =   2760
      Visible         =   0   'False
      Width           =   3750
   End
   Begin VB.Image imgPresents 
      Height          =   1125
      Index           =   2
      Left            =   0
      Picture         =   "frmIntro.frx":4429
      Top             =   1680
      Visible         =   0   'False
      Width           =   4500
   End
   Begin VB.Image imgPresents 
      Height          =   750
      Index           =   1
      Left            =   4320
      Picture         =   "frmIntro.frx":4A54
      Top             =   720
      Visible         =   0   'False
      Width           =   3000
   End
   Begin VB.Image imgPresents 
      Height          =   750
      Index           =   0
      Left            =   0
      Picture         =   "frmIntro.frx":4E3B
      Top             =   0
      Visible         =   0   'False
      Width           =   3750
   End
   Begin VB.Image imgLogo 
      Height          =   1635
      Left            =   1440
      Picture         =   "frmIntro.frx":5412
      Top             =   1080
      Visible         =   0   'False
      Width           =   4890
   End
End
Attribute VB_Name = "frmIntro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim introLoaded As Boolean
Dim intPlot As Long
Dim intMode As Long
Private Sub cmdEditor_Click()
timeDemo.Enabled = False
Demo = False
Call PlySound("Transmission")
frmEditor.Show
End Sub

Private Sub cmdHowToPlay_Click()
frmHowToPlay.Show
End Sub

Private Sub cmdLoad_Click()
Call PlySound("Transmission")
Dim strLoad As String
strLoad = GetFromIni("GEN", "SAVE", App.Path & "\save.ini")
If strLoad = "" Then
    MsgBox "No level to load!"
Else
    NextLevel = strLoad
    frmBriefing.Show
End If
End Sub

Private Sub cmdPlay_Click()
Demo = False
timeDemo.Enabled = False
curCampaign = ""
Call PlySound("Transmission")
StopMidi (curMidi)
frmCampaign.Show
End Sub

Private Sub cmdPlay_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF1 Then
    timeDemo.Enabled = False
    Demo = False
    NextLevel = InputBox("Input level name (level1.ini)")
    StopMidi (curMidi)
    frmBriefing.Show
End If
If KeyCode = vbKeyF2 Then
    Call timeDemo_Timer
End If
End Sub

Private Sub cmdQuit_Click()
Call PlySound("Transmission")
Unload Me
End Sub

Private Sub Form_Activate()
If introLoaded = False Then
    Call Form_Load
End If
End Sub

Private Sub Form_Load()
On Error Resume Next
Dim strSave As String
Dim strFirstLoad As String
Dim strTemp As String
strSave = App.Path & "\settings.dat"
strFirstLoad = GetFromIni("GEN", "FIRSTLOAD", strSave)
strTemp = GetFromIni("GEN", "MIDI", strSave)
If strTemp = "1" Or strTemp = "" Then
    ENABLE_MIDI = True
Else
    ENABLE_MIDI = False
End If
strTemp = GetFromIni("GEN", "WAV", strSave)
If strTemp = "1" Or strTemp = "" Then
    ENABLE_WAV = True
Else
    ENABLE_WAV = False
End If
strTemp = GetFromIni("GEN", "DEMO", strSave)
If strTemp = "1" Or strTemp = "" Then
    ENABLE_DEMO = True
Else
    ENABLE_DEMO = False
End If

If strFirstLoad = "" Then
    anType = 0
    frmCinema.Show
    timePlot.Enabled = False
    timeIntro.Enabled = False
    Call WriteIni("GEN", "FIRSTLOAD", "1", App.Path & "\settings.dat")
    firstCinematic = True
    timeDemo.Enabled = False
    introLoaded = False
    Me.Hide
Else
    introLoaded = True
    firstCinematic = False
    framProlouge.Visible = False
    timePlot.Enabled = False
    timeIntro.Enabled = True
    intMode = 0
    curMidi = "intro.mid"
    PlayMidi (curMidi)
End If

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblCampaign.ForeColor = COLINTROIN
lblHowToPlay.ForeColor = COLINTROIN
lblOptions.ForeColor = COLINTROIN
lblQuit.ForeColor = COLINTROIN
End Sub

Private Sub Form_Unload(Cancel As Integer)
StopMidi (curMidi)
End
End Sub

Private Sub imgLogo_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 And Shift = 1 Then
    cmdEditor.Visible = True
End If
End Sub

Private Sub imgLogo_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblCampaign.ForeColor = COLINTROIN
lblHowToPlay.ForeColor = COLINTROIN
lblOptions.ForeColor = COLINTROIN
lblQuit.ForeColor = COLINTROIN
End Sub

Private Sub lblCampaign_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
    Demo = False
    timeDemo.Enabled = False
    curCampaign = ""
    Call PlySound("Transmission")
    StopMidi (curMidi)
    frmCampaign.Show
    frmIntro.Hide
ElseIf Button = 2 And Shift = 1 Then
    Demo = False
    timeDemo.Enabled = False
    curCampaign = ""
    StopMidi (curMidi)
    NextLevel = InputBox("Next level")
    frmBriefing.Show
    frmIntro.Hide
End If
End Sub

Private Sub lblCampaign_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblCampaign.ForeColor = COLINTROIN
lblHowToPlay.ForeColor = COLINTROIN
lblOptions.ForeColor = COLINTROIN
lblQuit.ForeColor = COLINTROIN

lblCampaign.ForeColor = COLGREEN
End Sub

Private Sub lblHowToPlay_Click()
frmHowToPlay.Show
End Sub

Private Sub lblHowToPlay_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblCampaign.ForeColor = COLINTROIN
lblHowToPlay.ForeColor = COLINTROIN
lblOptions.ForeColor = COLINTROIN
lblQuit.ForeColor = COLINTROIN

lblHowToPlay.ForeColor = COLGREEN
End Sub

Private Sub lblOptions_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
frmGameOptions.Show
End Sub

Private Sub lblOptions_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblCampaign.ForeColor = COLINTROIN
lblHowToPlay.ForeColor = COLINTROIN
lblOptions.ForeColor = COLINTROIN
lblQuit.ForeColor = COLINTROIN

lblOptions.ForeColor = COLGREEN
End Sub

Private Sub lblQuit_Click()
Unload Me
End Sub

Private Sub lblQuit_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblCampaign.ForeColor = COLINTROIN
lblHowToPlay.ForeColor = COLINTROIN
lblOptions.ForeColor = COLINTROIN
lblQuit.ForeColor = COLINTROIN

lblQuit.ForeColor = COLGREEN
End Sub

Private Sub lblSkip_Click()
framProlouge.Visible = False
timeIntro.Enabled = True
intMode = 0
StopMidi (curMidi)
curMidi = "intro.mid"
PlayMidi (curMidi)
timePlot.Enabled = False
End Sub

Private Sub timeDemo_Timer()
frmIntro.Hide
LevelLoaded = False
Dim intRand As Long
Randomize
intRand = Int(Rnd * 2)
If intRand = 1 Then
    NextLevel = "level3.ini"
Else
    NextLevel = "level1.ini"
End If
Demo = True
frmGame.Show
timeDemo.Enabled = False
End Sub

Private Sub timeIntro_Timer()
If intMode = 0 Then
    imgDocInc.Left = imgDocInc.Left + 15
    If imgDocInc.Left >= (Me.ScaleWidth / 2) - 20 Then
        PlySound ("skid")
        intMode = 1
    End If
ElseIf intMode = 1 Then
    imgDocInc.Left = imgDocInc.Left - 10
    If imgDocInc.Left <= (Me.ScaleWidth / 2) - 100 Then
        intMode = 2
    End If
ElseIf intMode < 20 Then
    intMode = intMode + 1
ElseIf intMode = 20 Then
    imgDocInc.Visible = False
    For i = 0 To 4
        imgPresents(i).Visible = True
    Next 'i
    PlySound ("docent")
    intMode = 21
ElseIf intMode = 21 Then
    imgPresents(0).Left = imgPresents(0).Left + 16
    imgPresents(1).Left = imgPresents(1).Left - 14
    imgPresents(2).Left = imgPresents(2).Left + 18
    imgPresents(3).Left = imgPresents(3).Left - 15
    imgPresents(4).Left = imgPresents(4).Left + 17
    If imgPresents(1).Left + imgPresents(1).Width <= 0 Then
        For i = 0 To 4
            imgPresents(i).Visible = False
        Next 'i
        intMode = 22
    End If
ElseIf intMode < 185 Then
    intMode = intMode + 1
Else
    timeIntro.Enabled = False
    imgLogo.Visible = True
    lblCampaign.Visible = True
    lblHowToPlay.Visible = True
    lblOptions.Visible = True
    lblQuit.Visible = True
    'cmdPlay.Visible = True
    'cmdEditor.Visible = True
    'cmdQuit.Visible = True
    'cmdLoad.Visible = True
    'cmdHowToPlay.Visible = True
End If
End Sub



Private Sub timeMidi_Timer()
Call RepeatMidi
End Sub

Private Sub timePlot_Timer()
For i = 0 To 3
    lblPlot(i).Top = lblPlot(i).Top - 7
    If lblPlot(i).Top <= shpBar(1).Top Then
        If lblPlot(i).Top + lblPlot(i).Height <= shpBar(0).Top + shpBar(0).Height Then
            If i = 3 Then
                framProlouge.Visible = False
                timeIntro.Enabled = True
                intMode = 0
                StopMidi (curMidi)
                curMidi = "intro.mid"
                PlayMidi (curMidi)
                timePlot.Enabled = False
            End If
            lblPlot(i).Visible = False
        Else
            lblPlot(i).Visible = True
        End If
    End If
Next 'i
End Sub

Private Sub Timer1_Timer()
Call RepeatMidi
End Sub
