VERSION 5.00
Begin VB.Form frmBriefing 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Hover Tanx Briefing Room"
   ClientHeight    =   4650
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7500
   Icon            =   "frmBriefing.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   310
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   500
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame framMap 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   4695
      Left            =   -120
      TabIndex        =   12
      Top             =   4440
      Width           =   7575
      Begin VB.Timer timeMap 
         Interval        =   2000
         Left            =   4560
         Top             =   120
      End
      Begin VB.Image imgMap 
         Height          =   3750
         Left            =   0
         Picture         =   "frmBriefing.frx":0D2A
         Top             =   480
         Width           =   7500
      End
   End
   Begin VB.CommandButton cmdInfo 
      BackColor       =   &H00404080&
      Caption         =   "&Tank Info"
      Height          =   615
      Left            =   5640
      MaskColor       =   &H00000000&
      TabIndex        =   6
      Top             =   1320
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Timer timeBars 
      Interval        =   700
      Left            =   5520
      Top             =   2160
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Height          =   495
      Left            =   5640
      TabIndex        =   3
      Top             =   3480
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton cmdForward 
      Caption         =   "&Forward"
      Enabled         =   0   'False
      Height          =   495
      Left            =   5640
      TabIndex        =   0
      Top             =   3000
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "&Back"
      Enabled         =   0   'False
      Height          =   495
      Left            =   5640
      TabIndex        =   1
      Top             =   2520
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "&Quit"
      Height          =   615
      Left            =   5640
      TabIndex        =   4
      Top             =   720
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton cmdPlay 
      Caption         =   "&Play"
      Enabled         =   0   'False
      Height          =   615
      Left            =   5640
      TabIndex        =   2
      Top             =   1920
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label lblForward 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Forward"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   255
      Left            =   1380
      MouseIcon       =   "frmBriefing.frx":95E2
      MousePointer    =   99  'Custom
      TabIndex        =   19
      Top             =   4140
      Width           =   855
   End
   Begin VB.Shape shpButton 
      BorderColor     =   &H0000FF00&
      Height          =   615
      Index           =   4
      Left            =   1320
      Shape           =   4  'Rounded Rectangle
      Top             =   3960
      Width           =   975
   End
   Begin VB.Label lblBack 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Back"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   255
      Left            =   300
      MouseIcon       =   "frmBriefing.frx":9EAC
      MousePointer    =   99  'Custom
      TabIndex        =   18
      Top             =   4140
      Width           =   855
   End
   Begin VB.Shape shpButton 
      BorderColor     =   &H0000FF00&
      Height          =   615
      Index           =   3
      Left            =   240
      Shape           =   4  'Rounded Rectangle
      Top             =   3960
      Width           =   975
   End
   Begin VB.Label lblQuit 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   2460
      MouseIcon       =   "frmBriefing.frx":A776
      MousePointer    =   99  'Custom
      TabIndex        =   17
      Top             =   420
      Width           =   855
   End
   Begin VB.Shape shpButton 
      BorderColor     =   &H0000FF00&
      Height          =   615
      Index           =   2
      Left            =   2400
      Shape           =   4  'Rounded Rectangle
      Top             =   240
      Width           =   975
   End
   Begin VB.Label lblPlay 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Play"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   255
      Left            =   1380
      MouseIcon       =   "frmBriefing.frx":B040
      MousePointer    =   99  'Custom
      TabIndex        =   16
      Top             =   420
      Width           =   855
   End
   Begin VB.Shape shpButton 
      BorderColor     =   &H0000FF00&
      Height          =   615
      Index           =   1
      Left            =   1320
      Shape           =   4  'Rounded Rectangle
      Top             =   240
      Width           =   975
   End
   Begin VB.Label lblTankInfo 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Tank Info"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   300
      MouseIcon       =   "frmBriefing.frx":B90A
      MousePointer    =   99  'Custom
      TabIndex        =   15
      Top             =   420
      Width           =   855
   End
   Begin VB.Shape shpButton 
      BorderColor     =   &H0000FF00&
      Height          =   615
      Index           =   0
      Left            =   240
      Shape           =   4  'Rounded Rectangle
      Top             =   240
      Width           =   975
   End
   Begin VB.Label lblName 
      BackStyle       =   0  'Transparent
      Caption         =   "Desolate Desert"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   360
      TabIndex        =   14
      Top             =   1515
      Width           =   4695
   End
   Begin VB.Shape shpName 
      BorderColor     =   &H0000FF00&
      FillColor       =   &H0000FF00&
      Height          =   375
      Left            =   240
      Shape           =   4  'Rounded Rectangle
      Top             =   1440
      Width           =   5055
   End
   Begin VB.Label lblMessage 
      BackStyle       =   0  'Transparent
      Caption         =   "Mesasge from H.Q."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   360
      TabIndex        =   13
      Top             =   1155
      Width           =   4695
   End
   Begin VB.Shape shpMessage 
      BorderColor     =   &H0000FF00&
      FillColor       =   &H0000FF00&
      Height          =   375
      Left            =   240
      Shape           =   4  'Rounded Rectangle
      Top             =   1080
      Width           =   5055
   End
   Begin VB.Label lblCloak 
      BackStyle       =   0  'Transparent
      Caption         =   "Cloak"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   3240
      TabIndex        =   11
      Top             =   2640
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label lblDoubleHover 
      BackStyle       =   0  'Transparent
      Caption         =   "Double Hover"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   3240
      TabIndex        =   10
      Top             =   2400
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label lblBoost 
      BackStyle       =   0  'Transparent
      Caption         =   "Boost"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   3240
      TabIndex        =   9
      Top             =   2880
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label lblSFeatures 
      BackStyle       =   0  'Transparent
      Caption         =   "Special Features:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   3120
      TabIndex        =   8
      Top             =   2160
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Label lblTank 
      BackStyle       =   0  'Transparent
      Caption         =   "D1 Hover Tank - Light Armor and Fire Power.  Good Mobility and Hovering."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   615
      Left            =   360
      TabIndex        =   7
      Top             =   3600
      Visible         =   0   'False
      Width           =   4815
   End
   Begin VB.Image imgTank 
      Height          =   1125
      Left            =   240
      Picture         =   "frmBriefing.frx":C1D4
      Top             =   2040
      Visible         =   0   'False
      Width           =   2205
   End
   Begin VB.Image imgBars 
      Height          =   735
      Index           =   5
      Left            =   4080
      Picture         =   "frmBriefing.frx":C8F3
      Top             =   240
      Visible         =   0   'False
      Width           =   3150
   End
   Begin VB.Image imgBars 
      Height          =   735
      Index           =   4
      Left            =   4080
      Picture         =   "frmBriefing.frx":D140
      Top             =   240
      Visible         =   0   'False
      Width           =   3150
   End
   Begin VB.Image imgBars 
      Height          =   735
      Index           =   3
      Left            =   4080
      Picture         =   "frmBriefing.frx":D97B
      Top             =   240
      Visible         =   0   'False
      Width           =   3150
   End
   Begin VB.Image imgBars 
      Height          =   735
      Index           =   2
      Left            =   4080
      Picture         =   "frmBriefing.frx":E187
      Top             =   240
      Visible         =   0   'False
      Width           =   3150
   End
   Begin VB.Image imgBars 
      Height          =   735
      Index           =   1
      Left            =   4080
      Picture         =   "frmBriefing.frx":E9B9
      Top             =   240
      Visible         =   0   'False
      Width           =   3150
   End
   Begin VB.Image imgCircle 
      Height          =   2205
      Index           =   3
      Left            =   6240
      Picture         =   "frmBriefing.frx":F206
      Top             =   2280
      Visible         =   0   'False
      Width           =   1050
   End
   Begin VB.Image imgCircle 
      Height          =   2205
      Index           =   2
      Left            =   6240
      Picture         =   "frmBriefing.frx":FA5E
      Top             =   2280
      Visible         =   0   'False
      Width           =   1050
   End
   Begin VB.Image imgCircle 
      Height          =   2205
      Index           =   1
      Left            =   6240
      Picture         =   "frmBriefing.frx":102DF
      Top             =   2280
      Visible         =   0   'False
      Width           =   1050
   End
   Begin VB.Label lblDesc 
      BackStyle       =   0  'Transparent
      Caption         =   "No briefing."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   2055
      Left            =   240
      TabIndex        =   5
      Top             =   1920
      Width           =   5895
   End
   Begin VB.Image imgBars 
      Height          =   735
      Index           =   0
      Left            =   4080
      Picture         =   "frmBriefing.frx":10B5C
      Top             =   240
      Width           =   3150
   End
   Begin VB.Image imgCircle 
      Height          =   2205
      Index           =   0
      Left            =   6240
      Picture         =   "frmBriefing.frx":1138A
      Top             =   2280
      Width           =   1050
   End
End
Attribute VB_Name = "frmBriefing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim bLoaded As Boolean
Dim BarFrame As Long 'The frame of the green bars that go left and right
Dim CircleFrame As Long 'The frame of the green bars over the circle
Dim curBriefing As Long 'The part of the briefing you're at
Dim strBriefing(0 To 2) As String 'The actual text of each briefing part

Private Sub cmdBack_Click()
On Error Resume Next
Call PlySound("Transmission")
curBriefing = curBriefing - 1 'Decrement the briefing you're at
lblDesc.Caption = strBriefing(curBriefing) 'Load the previous briefing text
If curBriefing = 0 Then 'Disable the back button, enable the forward button if you're at the first briefing
    cmdBack.Enabled = False
    lblBack.ForeColor = COLGRAY
    cmdPlay.Enabled = False
    lblPlay.ForeColor = COLGRAY
    cmdForward.Enabled = True
    lblForward.ForeColor = COLGREEN
End If
End Sub

Private Sub cmdForward_Click()
On Error Resume Next
Call PlySound("Transmission")
curBriefing = curBriefing + 1 'Increment the briefing you're at
lblDesc.Caption = strBriefing(curBriefing) 'Load the next briefing text
cmdBack.Enabled = True 'Enable the back button
lblBack.ForeColor = COLGREEN
If curBriefing = 2 Then 'Disable the forward button since you're at the end
    cmdForward.Enabled = False
    lblForward.ForeColor = COLGRAY
    cmdPlay.Enabled = True
    lblPlay.ForeColor = COLGREEN
Else
    If strBriefing(curBriefing + 1) = "" Then 'Disable the forward button because the next text is blank
        cmdForward.Enabled = False
        lblForward.ForeColor = COLGRAY
        cmdPlay.Enabled = True
        lblPlay.ForeColor = COLGREEN
    End If
End If
End Sub

Private Sub cmdInfo_Click()

On Error Resume Next
If lblTankInfo.Caption = "Tank Info" Then 'Load the information about the tank for the current mission
    lblTankInfo.Caption = "Briefing"
    cmdForward.Visible = False
    lblPlay.Visible = False
    lblQuit.Visible = False
    lblBack.Visible = False
    lblForward.Visible = False
    lblTankInfo.Caption = "Briefing"
    For i = 1 To shpButton.UBound
        shpButton(i).Visible = False
    Next 'i
    lblDesc.Visible = False
    imgTank.Visible = True
    lblTank.Visible = True
    lblSFeatures.Visible = True
    lblBoost.Visible = True
    lblDoubleHover.Visible = True
    lblCloak.Visible = True
Else 'Load the regular briefing
    lblDesc.Visible = True
    lblPlay.Visible = True
    lblQuit.Visible = True
    lblBack.Visible = True
    lblForward.Visible = True
    For i = 1 To shpButton.UBound
        shpButton(i).Visible = True
    Next 'i
    lblDesc.Visible = True
    imgTank.Visible = False
    lblTank.Visible = False
    lblSFeatures.Visible = False
    lblBoost.Visible = False
    lblDoubleHover.Visible = False
    lblCloak.Visible = False
    lblTankInfo.Caption = "Tank Info"
End If
End Sub

Private Sub cmdPlay_Click()
'Start the map
On Error Resume Next
Call PlySound("Transmission")
bLoaded = False 'Not loaded yet
LevelLoaded = False 'Level not loaded yet
Me.Hide
StopMidi (curMidi)
curMidi = GetFromIni("GEN", "MIDI", App.Path & "\" & NextLevel)
PlayMidi (curMidi)
frmGame.Show 'Load the game screen
End Sub

Private Sub cmdQuit_Click()
'Unload the briefing screen
Call PlySound("Transmission")
Unload Me
End Sub

Private Sub cmdSave_Click()
'Save progress (no longer used)
On Error Resume Next
Call PlySound("Transmission")
Call WriteIni("GEN", "SAVE", NextLevel, App.Path & "\save.ini")
End Sub

Private Sub Form_Activate()
Call Form_Load
End Sub

Private Sub Form_Load()
'Load the briefing
On Error Resume Next
Dim strSave As String
Dim strTemp As String
If bLoaded = True Then Exit Sub 'You've already loaded the briefing, quit
bLoaded = True 'You have now loaded the briefing screen
curMidi = "briefing.mid"
PlayMidi (curMidi) 'Play the briefing music
Call SaveCampaign
strSave = App.Path & "\" & NextLevel 'Set the loading file
For i = 0 To 2 'Load the briefing text
    strBriefing(i) = GetFromIni("GEN", "BRIEFING" & i, strSave)
Next 'i
lblName.Caption = GetFromIni("GEN", "NAME", strSave) 'Load level name
strTemp = GetFromIni("GEN", "TANKTYPE", strSave) 'Load the tank type
If strTemp = "" Then strTemp = "0" 'Set the tank type to default if it's nothing
imgTank.Picture = frmTank.imgTank(CInt(strTemp)).Picture
'Load the tank information
Select Case strTemp
Case "0"
    lblTank.Caption = "D1 Hover Tank - Light Armor and Firepower.  Good Mobility and Hovering."
Case "1"
    lblTank.Caption = "W1 Heavy Hover Tank - Special treds for snow travel.  Good Armor and Fire Power.  Light Hovering."
Case "2"
    lblTank.Caption = "C11-S Special Hover Tank - Great Hovering, Mobility and Firepower.  Average Armor."
Case "3"
    lblTank.Caption = "Hover Truck - Lacks any weapons.  Good Armor.  Average Hovering and Mobility."
Case "4"
    lblTank.Caption = "H20 Hover Boat - Capabale of water travel.  Average Firepower, Hovering, Armor."
Case "5"
    lblTank.Caption = "Hover Tanx MRA - Extremely Light Weight.  Low Firepower, Armor and Hover Capabilites."
Case "6"
    lblTank.Caption = "SX Hover Tanx Submarine - Emergency vehicle able to travel underwater for short distances.  Average Firepower and Armor."
Case "7"
    lblTank.Caption = "MXT-2000 Space Tank - Advanced Hover Tanx able to travel through space.  Low Firepower. Medium Armor, Hovering and Mobility."
End Select
'Load the tank attributes
strTemp = GetFromIni("GEN", "DOUBLEHOVER", strSave)
If strTemp = "1" Then
    lblDoubleHover.ForeColor = COLGREEN
Else
    lblDoubleHover.ForeColor = COLSNOW
End If
strTemp = GetFromIni("GEN", "BOOST", strSave)
If strTemp = "1" Then
    lblBoost.ForeColor = COLGREEN
Else
    lblBoost.ForeColor = COLSNOW
End If
strTemp = GetFromIni("GEN", "CLOAK", strSave)
If strTemp = "1" Then
    lblCloak.ForeColor = COLGREEN
Else
    lblCloak.ForeColor = COLSNOW
End If

lblDesc.Caption = strBriefing(0)
cmdBack.Enabled = False
lblBack.ForeColor = COLGRAY
curBriefing = 0
If strBriefing(1) = "" Then
    cmdPlay.Enabled = True
    lblPlay.ForeColor = COLGREEN
    cmdForward.Enabled = False
    lblForward.ForeColor = COLGRAY
Else
    cmdPlay.Enabled = False
    lblPlay.ForeColor = COLGRAY
    cmdForward.Enabled = True
    lblForward.ForeColor = COLGREEN
End If
framMap.Top = 0
framMap.Left = 0
framMap.Visible = True
timeMap.Enabled = True
Dim strLevelName() As String
strLevelName = Split(NextLevel, ".", -1, vbTextCompare)
imgMap.Picture = LoadPicture(App.Path & "\" & strLevelName(0) & ".jpg")


End Sub

Private Sub Form_Unload(Cancel As Integer)
Call StopMidi(curMidi) 'Stop the briefing music
bLoaded = False 'the briefing has not been loaded
frmCampaign.Show 'Load the campaign screen
End Sub

Private Sub Label1_Click()

End Sub

Private Sub lblBack_Click()
cmdBack_Click 'Call the back button
End Sub

Private Sub lblBack_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Change mouse icon to hand if applicable
If lblBack.ForeColor = COLGRAY Then
    lblBack.MousePointer = 0
Else
    lblBack.MousePointer = 99
End If
End Sub

Private Sub lblForward_Click()
cmdForward_Click 'Call the forward button
End Sub

Private Sub lblForward_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Change mouse icon to hand if applicable
If lblForward.ForeColor = COLGRAY Then
    lblForward.MousePointer = 0
Else
    lblForward.MousePointer = 99
End If
End Sub

Private Sub lblPlay_Click()
cmdPlay_Click 'Call the play button
End Sub

Private Sub lblPlay_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Change mouse icon to hand if applicable
If lblPlay.ForeColor = COLGRAY Then
    lblPlay.MousePointer = 0
Else
    lblPlay.MousePointer = 99
End If
End Sub

Private Sub lblQuit_Click()
cmdQuit_Click 'Call the quit button
End Sub

Private Sub lblTankInfo_Click()
cmdInfo_Click 'Call the tank info button
End Sub

Private Sub timeBars_Timer()
'Animate the bars
For i = 0 To imgBars.UBound
    imgBars(i).Visible = False
Next 'i
BarFrame = BarFrame + 1
If BarFrame > imgBars.UBound Then BarFrame = 0
imgBars(BarFrame).Visible = True

For i = 0 To imgCircle.UBound
    imgCircle(i).Visible = False
Next 'i
CircleFrame = CircleFrame + 1
If CircleFrame > imgCircle.UBound Then CircleFrame = 0
imgCircle(CircleFrame).Visible = True

End Sub

Private Sub timeMap_Timer()
framMap.Visible = False
timeMap.Enabled = False
End Sub
Private Sub SaveCampaign()
'Determines if this level is a later level in the campaign than previously and if so, saves this info
On Error Resume Next
If NextLevel = "" Then Exit Sub 'next level is nothing, so you're automatically not further
Dim intCampaign As Long
Dim intNewLevel As Long
Dim strCampaign As String
strCampaign = GetFromIni("GEN", "LEVEL", App.Path & "\" & curCampaign)
intCampaign = LevelRank(strCampaign) 'Get the level value of the latest campaign level
intNewLevel = LevelRank(NextLevel) 'Get the level value of the current level
If intNewLevel >= intCampaign Then 'You're further than the campaign level, update
    Call WriteIni("GEN", "LEVEL", NextLevel, App.Path & "\" & curCampaign)
End If
End Sub
Private Function LevelRank(strLevel As String) As Long
'Assign a value to each level
On Error Resume Next
Dim intNewLevel As Long
Select Case strLevel
Case "level1.ini"
    intNewLevel = 0
Case "level2.ini"
    intNewLevel = 1
Case "dboss.ini"
    intNewLevel = 2
Case "level3.ini"
    intNewLevel = 3
Case "level4.ini"
    intNewLevel = 4
Case "iceboss.ini"
    intNewLevel = 5
Case "level5.ini"
    intNewLevel = 6
Case "level6.ini"
    intNewLevel = 7
Case "vboss.ini"
    intNewLevel = 8
Case "level7.ini"
    intNewLevel = 9
Case "level8.ini"
    intNewLevel = 10
Case "oboss.ini"
    intNewLevel = 11
Case "level9.ini"
    intNewLevel = 12
Case "level10.ini"
    intNewLevel = 13
Case "uboss.ini"
    intNewLevel = 14
Case "level11.ini"
    intNewLevel = 15
Case "level12.ini"
    intNewLevel = 16
Case "cboss.ini"
    intNewLevel = 17
Case "level13.ini"
    intNewLevel = 18
Case "level14.ini"
    intNewLevel = 19
Case "level15.ini"
    intNewLevel = 20
Case Else
    intNewLevel = 0
End Select
LevelRank = intNewLevel
End Function
