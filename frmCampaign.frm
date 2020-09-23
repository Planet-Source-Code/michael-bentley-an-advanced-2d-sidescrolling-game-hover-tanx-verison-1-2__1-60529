VERSION 5.00
Begin VB.Form frmCampaign 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Hover Tanx - Campaign"
   ClientHeight    =   4425
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10875
   Icon            =   "frmCampaign.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   295
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   725
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox lstLevels 
      BackColor       =   &H00000000&
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
      Height          =   3210
      ItemData        =   "frmCampaign.frx":0D2A
      Left            =   8160
      List            =   "frmCampaign.frx":0D31
      TabIndex        =   7
      Top             =   480
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.TextBox txtNewName 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      ForeColor       =   &H0000FF00&
      Height          =   285
      Left            =   4440
      MaxLength       =   20
      TabIndex        =   2
      Text            =   "New Campaign Name"
      Top             =   2640
      Width           =   1815
   End
   Begin VB.FileListBox filCampaign 
      BackColor       =   &H00404040&
      ForeColor       =   &H00FFFFFF&
      Height          =   2430
      Left            =   2280
      Pattern         =   "*.hvx"
      TabIndex        =   0
      Top             =   600
      Width           =   2055
   End
   Begin VB.Image imgCinema 
      Height          =   390
      Left            =   5880
      Picture         =   "frmCampaign.frx":0D46
      Top             =   2040
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.Shape shpLocation 
      BorderColor     =   &H000000FF&
      BorderWidth     =   3
      Height          =   255
      Left            =   4560
      Shape           =   3  'Circle
      Top             =   1560
      Visible         =   0   'False
      Width           =   255
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
      Left            =   6420
      TabIndex        =   10
      Top             =   3420
      Width           =   615
   End
   Begin VB.Shape shpQuit 
      BorderColor     =   &H0000FF00&
      BorderWidth     =   3
      Height          =   375
      Left            =   6360
      Shape           =   4  'Rounded Rectangle
      Top             =   3360
      Width           =   735
   End
   Begin VB.Shape shpMap 
      BorderColor     =   &H0000FF00&
      BorderWidth     =   3
      Height          =   3375
      Left            =   600
      Top             =   480
      Visible         =   0   'False
      Width           =   7185
   End
   Begin VB.Image imgMars 
      Height          =   3390
      Left            =   600
      Picture         =   "frmCampaign.frx":10EE
      Top             =   480
      Visible         =   0   'False
      Width           =   7170
   End
   Begin VB.Label lblBack 
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
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   9690
      TabIndex        =   9
      Top             =   4095
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label lblLoadMission 
      BackStyle       =   0  'Transparent
      Caption         =   "Load"
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
      Left            =   8370
      TabIndex        =   8
      Top             =   4095
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Shape shpBack 
      BorderColor     =   &H0000FF00&
      BorderWidth     =   3
      Height          =   255
      Left            =   9480
      Shape           =   4  'Rounded Rectangle
      Top             =   4080
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Shape shpLoadLvl 
      BorderColor     =   &H0000FF00&
      BorderWidth     =   3
      Height          =   255
      Left            =   8160
      Shape           =   4  'Rounded Rectangle
      Top             =   4080
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Shape shpLevels 
      BorderColor     =   &H0000FF00&
      BorderWidth     =   3
      Height          =   3735
      Left            =   7920
      Shape           =   4  'Rounded Rectangle
      Top             =   240
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.Image imgMap 
      Height          =   3390
      Left            =   600
      Picture         =   "frmCampaign.frx":7ADA
      Top             =   480
      Visible         =   0   'False
      Width           =   7170
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H0000FF00&
      BorderWidth     =   3
      Height          =   375
      Left            =   4200
      Shape           =   4  'Rounded Rectangle
      Top             =   3120
      Width           =   735
   End
   Begin VB.Label lblLoad 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Load"
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
      Left            =   4260
      TabIndex        =   6
      Top             =   3180
      Width           =   615
   End
   Begin VB.Label lblLoadLvl 
      Alignment       =   2  'Center
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
      Left            =   2280
      TabIndex        =   5
      Top             =   3360
      Width           =   2055
   End
   Begin VB.Label lblcLvl 
      BackStyle       =   0  'Transparent
      Caption         =   "Current Level:"
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
      Left            =   2640
      TabIndex        =   4
      Top             =   3120
      Width           =   1335
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H0000FF00&
      BorderWidth     =   3
      Height          =   375
      Left            =   6360
      Shape           =   4  'Rounded Rectangle
      Top             =   2640
      Width           =   735
   End
   Begin VB.Label lblNewCampaign 
      BackStyle       =   0  'Transparent
      Caption         =   "Create"
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
      Left            =   6450
      TabIndex        =   3
      Top             =   2700
      Width           =   615
   End
   Begin VB.Label lblLoadDesc 
      BackStyle       =   0  'Transparent
      Caption         =   "Please select a Campaign File to load, or create a new file by entering your name in the textbox below."
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
      Height          =   855
      Left            =   4440
      TabIndex        =   1
      Top             =   600
      Width           =   2895
   End
   Begin VB.Shape shpLoad 
      BorderColor     =   &H0000FF00&
      BorderWidth     =   3
      Height          =   3615
      Left            =   1800
      Shape           =   4  'Rounded Rectangle
      Top             =   360
      Width           =   5895
   End
End
Attribute VB_Name = "frmCampaign"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private loadLevel As Boolean
Private Sub filCampaign_Click()
'When you click on the campaign file list, display the current level the campaign is at
On Error Resume Next
If filCampaign.ListCount = 0 Then 'No campaigns in directory
    Exit Sub
ElseIf filCampaign.List(filCampaign.ListIndex) = "" Then 'An empty file
    Exit Sub
End If

Dim strLevel As String 'Get the last level from the campaign file and assign a name corresponding to the level
strLevel = GetFromIni("GEN", "LEVEL", App.Path & "\" & filCampaign.List(filCampaign.ListIndex))
Select Case strLevel
    Case "level1.ini"
        lblLoadLvl.Caption = "Desolate Desert"
    Case "level2.ini"
        lblLoadLvl.Caption = "Desert Defense"
    Case "dboss.ini"
        lblLoadLvl.Caption = "Desert Storm"
    Case "level3.ini"
        lblLoadLvl.Caption = "Steep Snow Drifts"
    Case "level4.ini"
        lblLoadLvl.Caption = "Mountain Pass"
    Case "iceboss.ini"
        lblLoadLvl.Caption = "Blizzard Battle"
    Case "level5.ini"
        lblLoadLvl.Caption = "Eruptive Vacation"
    Case "level6.ini"
        lblLoadLvl.Caption = "Tiny Trouble"
    Case "vboss.ini"
        lblLoadLvl.Caption = "High Protector"
    Case "level7.ini"
        lblLoadLvl.Caption = "Stranded!"
    Case "level8.ini"
        lblLoadLvl.Caption = "Cruisin'"
    Case "oboss.ini"
        lblLoadLvl.Caption = "Battleship Battle"
    Case "level9.ini"
        lblLoadLvl.Caption = "The Bottom of the Sea"
    Case "level10.ini"
        lblLoadLvl.Caption = "Trench Trouble"
    Case "uboss.ini"
        lblLoadLvl.Caption = "Base of the Abyss"
    Case "level11.ini"
        lblLoadLvl.Caption = "Metropolis Madness"
    Case "level12.ini"
        lblLoadLvl.Caption = "Rooftop Rumble"
    Case "cboss.ini"
        lblLoadLvl.Caption = "The Final Encounter"
    Case "level13.ini"
        lblLoadLvl.Caption = "Mars Patrol"
    Case "level14.ini"
        lblLoadLvl.Caption = "Red Planet"
    Case "level15.ini"
        lblLoadLvl.Caption = "Secrets of Space"
    Case Else
        lblLoadLvl.Caption = "Desolate Desert"
End Select
End Sub

Private Sub filCampaign_DblClick()
Call lblLoad_Click
End Sub

Private Sub Form_Activate()
loadLevel = False 'Level has not been loaded
End Sub

Private Sub Form_Load()
'Load the campaign screen
On Error Resume Next
If curCampaign <> "" Then 'You've already selected a campaign, move to the second part of the screen
    Call LoadSecondScreen(curCampaign)
End If
Dim strSave As String
Dim activeCampaign As String
Dim campaignNum As Long 'The last campaign to be loaded
strSave = App.Path & "\settings.ini"
activeCampaign = GetFromIni("GEN", "CAMPAIGN", strSave)
campaignNum = -1
If filCampaign.ListCount > 0 Then 'There are actually campaigns in the list
    For i = 0 To filCampaign.ListCount - 1
        If filCampaign.List(i) = activeCampaign Then 'This is the campaign you're looking for
            campaignNum = i
        End If
    Next 'i
End If
If activeCampaign <> "" And campaignNum > -1 Then 'You found active campaign, change the highlighted campaign to it
    filCampaign.ListIndex = campaignNum
    filCampaign_Click
End If
filCampaign.Path = App.Path
PlayMidi ("campaign.mid")

End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
StopMidi ("campaign.mid") 'Stop the campaign music
If loadLevel = False Then
    frmIntro.Show 'Load the intro screen
End If
End Sub

Private Sub PopulateList(strMap As String)
'Adds level names to the level select list depending on how far you've gotten
On Error Resume Next
lstLevels.Clear
Select Case strMap
Case "intro"
    lstLevels.AddItem "Introduction"
Case "mov01"
    lstLevels.AddItem "Introduction"
    lstLevels.AddItem "Desert Cinematic"
Case "level1.ini"
    lstLevels.AddItem "Introduction"
    lstLevels.AddItem "Desert Cinematic"
    lstLevels.AddItem "Desolate Desert"
Case "level2.ini"
    lstLevels.AddItem "Introduction"
    lstLevels.AddItem "Desert Cinematic"
    lstLevels.AddItem "Desolate Desert"
    lstLevels.AddItem "Desert Defense"
Case "dboss.ini"
    lstLevels.AddItem "Introduction"
    lstLevels.AddItem "Desert Cinematic"
    lstLevels.AddItem "Desolate Desert"
    lstLevels.AddItem "Desert Defense"
    lstLevels.AddItem "Desert Storm"
Case "mov02"
    lstLevels.AddItem "Introduction"
    lstLevels.AddItem "Desert Cinematic"
    lstLevels.AddItem "Desolate Desert"
    lstLevels.AddItem "Desert Defense"
    lstLevels.AddItem "Desert Storm"
    lstLevels.AddItem "Snow Cinematic"
Case "level3.ini"
    lstLevels.AddItem "Introduction"
    lstLevels.AddItem "Desert Cinematic"
    lstLevels.AddItem "Desolate Desert"
    lstLevels.AddItem "Desert Defense"
    lstLevels.AddItem "Desert Storm"
    lstLevels.AddItem "Snow Cinematic"
    lstLevels.AddItem "Steep Snow Drifts"
Case "level4.ini"
    lstLevels.AddItem "Introduction"
    lstLevels.AddItem "Desert Cinematic"
    lstLevels.AddItem "Desolate Desert"
    lstLevels.AddItem "Desert Defense"
    lstLevels.AddItem "Desert Storm"
    lstLevels.AddItem "Snow Cinematic"
    lstLevels.AddItem "Steep Snow Drifts"
    lstLevels.AddItem "Mountain Pass"
Case "iceboss.ini"
    lstLevels.AddItem "Introduction"
    lstLevels.AddItem "Desert Cinematic"
    lstLevels.AddItem "Desolate Desert"
    lstLevels.AddItem "Desert Defense"
    lstLevels.AddItem "Desert Storm"
    lstLevels.AddItem "Snow Cinematic"
    lstLevels.AddItem "Steep Snow Drifts"
    lstLevels.AddItem "Mountain Pass"
    lstLevels.AddItem "Blizzard Battle"
Case "mov03"
    lstLevels.AddItem "Introduction"
    lstLevels.AddItem "Desert Cinematic"
    lstLevels.AddItem "Desolate Desert"
    lstLevels.AddItem "Desert Defense"
    lstLevels.AddItem "Desert Storm"
    lstLevels.AddItem "Snow Cinematic"
    lstLevels.AddItem "Steep Snow Drifts"
    lstLevels.AddItem "Mountain Pass"
    lstLevels.AddItem "Blizzard Battle"
    lstLevels.AddItem "Vacation Cinematic"
Case "level5.ini"
    lstLevels.AddItem "Introduction"
    lstLevels.AddItem "Desert Cinematic"
    lstLevels.AddItem "Desolate Desert"
    lstLevels.AddItem "Desert Defense"
    lstLevels.AddItem "Desert Storm"
    lstLevels.AddItem "Snow Cinematic"
    lstLevels.AddItem "Steep Snow Drifts"
    lstLevels.AddItem "Mountain Pass"
    lstLevels.AddItem "Blizzard Battle"
    lstLevels.AddItem "Vacation Cinematic"
    lstLevels.AddItem "Eruptive Vacation"
Case "level6.ini"
    lstLevels.AddItem "Introduction"
    lstLevels.AddItem "Desert Cinematic"
    lstLevels.AddItem "Desolate Desert"
    lstLevels.AddItem "Desert Defense"
    lstLevels.AddItem "Desert Storm"
    lstLevels.AddItem "Snow Cinematic"
    lstLevels.AddItem "Steep Snow Drifts"
    lstLevels.AddItem "Mountain Pass"
    lstLevels.AddItem "Blizzard Battle"
    lstLevels.AddItem "Vacation Cinematic"
    lstLevels.AddItem "Eruptive Vacation"
    lstLevels.AddItem "Tiny Trouble"
Case "vboss.ini"
    lstLevels.AddItem "Introduction"
    lstLevels.AddItem "Desert Cinematic"
    lstLevels.AddItem "Desolate Desert"
    lstLevels.AddItem "Desert Defense"
    lstLevels.AddItem "Desert Storm"
    lstLevels.AddItem "Snow Cinematic"
    lstLevels.AddItem "Steep Snow Drifts"
    lstLevels.AddItem "Mountain Pass"
    lstLevels.AddItem "Blizzard Battle"
    lstLevels.AddItem "Vacation Cinematic"
    lstLevels.AddItem "Eruptive Vacation"
    lstLevels.AddItem "Tiny Trouble"
    lstLevels.AddItem "High Protector"
Case "mov04"
    lstLevels.AddItem "Introduction"
    lstLevels.AddItem "Desert Cinematic"
    lstLevels.AddItem "Desolate Desert"
    lstLevels.AddItem "Desert Defense"
    lstLevels.AddItem "Desert Storm"
    lstLevels.AddItem "Snow Cinematic"
    lstLevels.AddItem "Steep Snow Drifts"
    lstLevels.AddItem "Mountain Pass"
    lstLevels.AddItem "Blizzard Battle"
    lstLevels.AddItem "Vacation Cinematic"
    lstLevels.AddItem "Eruptive Vacation"
    lstLevels.AddItem "Tiny Trouble"
    lstLevels.AddItem "High Protector"
    lstLevels.AddItem "Ocean Cinematic"
Case "level7.ini"
    lstLevels.AddItem "Introduction"
    lstLevels.AddItem "Desert Cinematic"
    lstLevels.AddItem "Desolate Desert"
    lstLevels.AddItem "Desert Defense"
    lstLevels.AddItem "Desert Storm"
    lstLevels.AddItem "Snow Cinematic"
    lstLevels.AddItem "Steep Snow Drifts"
    lstLevels.AddItem "Mountain Pass"
    lstLevels.AddItem "Blizzard Battle"
    lstLevels.AddItem "Vacation Cinematic"
    lstLevels.AddItem "Eruptive Vacation"
    lstLevels.AddItem "Tiny Trouble"
    lstLevels.AddItem "High Protector"
    lstLevels.AddItem "Ocean Cinematic"
    lstLevels.AddItem "Stranded!"
Case "level8.ini"
    lstLevels.AddItem "Introduction"
    lstLevels.AddItem "Desert Cinematic"
    lstLevels.AddItem "Desolate Desert"
    lstLevels.AddItem "Desert Defense"
    lstLevels.AddItem "Desert Storm"
    lstLevels.AddItem "Snow Cinematic"
    lstLevels.AddItem "Steep Snow Drifts"
    lstLevels.AddItem "Mountain Pass"
    lstLevels.AddItem "Blizzard Battle"
    lstLevels.AddItem "Vacation Cinematic"
    lstLevels.AddItem "Eruptive Vacation"
    lstLevels.AddItem "Tiny Trouble"
    lstLevels.AddItem "High Protector"
    lstLevels.AddItem "Ocean Cinematic"
    lstLevels.AddItem "Stranded!"
    lstLevels.AddItem "Cruisn'"
Case "oboss.ini"
    lstLevels.AddItem "Introduction"
    lstLevels.AddItem "Desert Cinematic"
    lstLevels.AddItem "Desolate Desert"
    lstLevels.AddItem "Desert Defense"
    lstLevels.AddItem "Desert Storm"
    lstLevels.AddItem "Snow Cinematic"
    lstLevels.AddItem "Steep Snow Drifts"
    lstLevels.AddItem "Mountain Pass"
    lstLevels.AddItem "Blizzard Battle"
    lstLevels.AddItem "Vacation Cinematic"
    lstLevels.AddItem "Eruptive Vacation"
    lstLevels.AddItem "Tiny Trouble"
    lstLevels.AddItem "High Protector"
    lstLevels.AddItem "Ocean Cinematic"
    lstLevels.AddItem "Stranded!"
    lstLevels.AddItem "Cruisn'"
    lstLevels.AddItem "Battleship Battle"
Case "mov05"
    lstLevels.AddItem "Introduction"
    lstLevels.AddItem "Desert Cinematic"
    lstLevels.AddItem "Desolate Desert"
    lstLevels.AddItem "Desert Defense"
    lstLevels.AddItem "Desert Storm"
    lstLevels.AddItem "Snow Cinematic"
    lstLevels.AddItem "Steep Snow Drifts"
    lstLevels.AddItem "Mountain Pass"
    lstLevels.AddItem "Blizzard Battle"
    lstLevels.AddItem "Vacation Cinematic"
    lstLevels.AddItem "Eruptive Vacation"
    lstLevels.AddItem "Tiny Trouble"
    lstLevels.AddItem "High Protector"
    lstLevels.AddItem "Ocean Cinematic"
    lstLevels.AddItem "Stranded!"
    lstLevels.AddItem "Cruisn'"
    lstLevels.AddItem "Battleship Battle"
    lstLevels.AddItem "Underwater Cinematic"
Case "level9.ini"
    lstLevels.AddItem "Introduction"
    lstLevels.AddItem "Desert Cinematic"
    lstLevels.AddItem "Desolate Desert"
    lstLevels.AddItem "Desert Defense"
    lstLevels.AddItem "Desert Storm"
    lstLevels.AddItem "Snow Cinematic"
    lstLevels.AddItem "Steep Snow Drifts"
    lstLevels.AddItem "Mountain Pass"
    lstLevels.AddItem "Blizzard Battle"
    lstLevels.AddItem "Vacation Cinematic"
    lstLevels.AddItem "Eruptive Vacation"
    lstLevels.AddItem "Tiny Trouble"
    lstLevels.AddItem "High Protector"
    lstLevels.AddItem "Ocean Cinematic"
    lstLevels.AddItem "Stranded!"
    lstLevels.AddItem "Cruisn'"
    lstLevels.AddItem "Battleship Battle"
    lstLevels.AddItem "Underwater Cinematic"
    lstLevels.AddItem "The Bottom of the Sea"
Case "level10.ini"
    lstLevels.AddItem "Introduction"
    lstLevels.AddItem "Desert Cinematic"
    lstLevels.AddItem "Desolate Desert"
    lstLevels.AddItem "Desert Defense"
    lstLevels.AddItem "Desert Storm"
    lstLevels.AddItem "Snow Cinematic"
    lstLevels.AddItem "Steep Snow Drifts"
    lstLevels.AddItem "Mountain Pass"
    lstLevels.AddItem "Blizzard Battle"
    lstLevels.AddItem "Vacation Cinematic"
    lstLevels.AddItem "Eruptive Vacation"
    lstLevels.AddItem "Tiny Trouble"
    lstLevels.AddItem "High Protector"
    lstLevels.AddItem "Ocean Cinematic"
    lstLevels.AddItem "Stranded!"
    lstLevels.AddItem "Cruisn'"
    lstLevels.AddItem "Battleship Battle"
    lstLevels.AddItem "Underwater Cinematic"
    lstLevels.AddItem "The Bottom of the Sea"
    lstLevels.AddItem "Trench Trouble"
Case "uboss.ini"
    lstLevels.AddItem "Introduction"
    lstLevels.AddItem "Desert Cinematic"
    lstLevels.AddItem "Desolate Desert"
    lstLevels.AddItem "Desert Defense"
    lstLevels.AddItem "Desert Storm"
    lstLevels.AddItem "Snow Cinematic"
    lstLevels.AddItem "Steep Snow Drifts"
    lstLevels.AddItem "Mountain Pass"
    lstLevels.AddItem "Blizzard Battle"
    lstLevels.AddItem "Vacation Cinematic"
    lstLevels.AddItem "Eruptive Vacation"
    lstLevels.AddItem "Tiny Trouble"
    lstLevels.AddItem "High Protector"
    lstLevels.AddItem "Ocean Cinematic"
    lstLevels.AddItem "Stranded!"
    lstLevels.AddItem "Cruisn'"
    lstLevels.AddItem "Battleship Battle"
    lstLevels.AddItem "Underwater Cinematic"
    lstLevels.AddItem "The Bottom of the Sea"
    lstLevels.AddItem "Trench Trouble"
    lstLevels.AddItem "Base of the Abyss"
Case "mov06"
    lstLevels.AddItem "Introduction"
    lstLevels.AddItem "Desert Cinematic"
    lstLevels.AddItem "Desolate Desert"
    lstLevels.AddItem "Desert Defense"
    lstLevels.AddItem "Desert Storm"
    lstLevels.AddItem "Snow Cinematic"
    lstLevels.AddItem "Steep Snow Drifts"
    lstLevels.AddItem "Mountain Pass"
    lstLevels.AddItem "Blizzard Battle"
    lstLevels.AddItem "Vacation Cinematic"
    lstLevels.AddItem "Eruptive Vacation"
    lstLevels.AddItem "Tiny Trouble"
    lstLevels.AddItem "High Protector"
    lstLevels.AddItem "Ocean Cinematic"
    lstLevels.AddItem "Stranded!"
    lstLevels.AddItem "Cruisn'"
    lstLevels.AddItem "Battleship Battle"
    lstLevels.AddItem "Underwater Cinematic"
    lstLevels.AddItem "The Bottom of the Sea"
    lstLevels.AddItem "Trench Trouble"
    lstLevels.AddItem "Base of the Abyss"
    lstLevels.AddItem "City Cinematic"
Case "level11.ini"
    lstLevels.AddItem "Introduction"
    lstLevels.AddItem "Desert Cinematic"
    lstLevels.AddItem "Desolate Desert"
    lstLevels.AddItem "Desert Defense"
    lstLevels.AddItem "Desert Storm"
    lstLevels.AddItem "Snow Cinematic"
    lstLevels.AddItem "Steep Snow Drifts"
    lstLevels.AddItem "Mountain Pass"
    lstLevels.AddItem "Blizzard Battle"
    lstLevels.AddItem "Vacation Cinematic"
    lstLevels.AddItem "Eruptive Vacation"
    lstLevels.AddItem "Tiny Trouble"
    lstLevels.AddItem "High Protector"
    lstLevels.AddItem "Ocean Cinematic"
    lstLevels.AddItem "Stranded!"
    lstLevels.AddItem "Cruisn'"
    lstLevels.AddItem "Battleship Battle"
    lstLevels.AddItem "Underwater Cinematic"
    lstLevels.AddItem "The Bottom of the Sea"
    lstLevels.AddItem "Trench Trouble"
    lstLevels.AddItem "Base of the Abyss"
    lstLevels.AddItem "City Cinematic"
    lstLevels.AddItem "Metropolis Madness"
Case "level12.ini"
    lstLevels.AddItem "Introduction"
    lstLevels.AddItem "Desert Cinematic"
    lstLevels.AddItem "Desolate Desert"
    lstLevels.AddItem "Desert Defense"
    lstLevels.AddItem "Desert Storm"
    lstLevels.AddItem "Snow Cinematic"
    lstLevels.AddItem "Steep Snow Drifts"
    lstLevels.AddItem "Mountain Pass"
    lstLevels.AddItem "Blizzard Battle"
    lstLevels.AddItem "Vacation Cinematic"
    lstLevels.AddItem "Eruptive Vacation"
    lstLevels.AddItem "Tiny Trouble"
    lstLevels.AddItem "High Protector"
    lstLevels.AddItem "Ocean Cinematic"
    lstLevels.AddItem "Stranded!"
    lstLevels.AddItem "Cruisn'"
    lstLevels.AddItem "Battleship Battle"
    lstLevels.AddItem "Underwater Cinematic"
    lstLevels.AddItem "The Bottom of the Sea"
    lstLevels.AddItem "Trench Trouble"
    lstLevels.AddItem "Base of the Abyss"
    lstLevels.AddItem "City Cinematic"
    lstLevels.AddItem "Metropolis Madness"
    lstLevels.AddItem "Rooftop Rumble"
Case "cboss.ini"
    lstLevels.AddItem "Introduction"
    lstLevels.AddItem "Desert Cinematic"
    lstLevels.AddItem "Desolate Desert"
    lstLevels.AddItem "Desert Defense"
    lstLevels.AddItem "Desert Storm"
    lstLevels.AddItem "Snow Cinematic"
    lstLevels.AddItem "Steep Snow Drifts"
    lstLevels.AddItem "Mountain Pass"
    lstLevels.AddItem "Blizzard Battle"
    lstLevels.AddItem "Vacation Cinematic"
    lstLevels.AddItem "Eruptive Vacation"
    lstLevels.AddItem "Tiny Trouble"
    lstLevels.AddItem "High Protector"
    lstLevels.AddItem "Ocean Cinematic"
    lstLevels.AddItem "Stranded!"
    lstLevels.AddItem "Cruisn'"
    lstLevels.AddItem "Battleship Battle"
    lstLevels.AddItem "Underwater Cinematic"
    lstLevels.AddItem "The Bottom of the Sea"
    lstLevels.AddItem "Trench Trouble"
    lstLevels.AddItem "Base of the Abyss"
    lstLevels.AddItem "City Cinematic"
    lstLevels.AddItem "Metropolis Madness"
    lstLevels.AddItem "Rooftop Rumble"
    lstLevels.AddItem "The Final Encounter"
Case "final"
    lstLevels.AddItem "Introduction"
    lstLevels.AddItem "Desert Cinematic"
    lstLevels.AddItem "Desolate Desert"
    lstLevels.AddItem "Desert Defense"
    lstLevels.AddItem "Desert Storm"
    lstLevels.AddItem "Snow Cinematic"
    lstLevels.AddItem "Steep Snow Drifts"
    lstLevels.AddItem "Mountain Pass"
    lstLevels.AddItem "Blizzard Battle"
    lstLevels.AddItem "Vacation Cinematic"
    lstLevels.AddItem "Eruptive Vacation"
    lstLevels.AddItem "Tiny Trouble"
    lstLevels.AddItem "High Protector"
    lstLevels.AddItem "Ocean Cinematic"
    lstLevels.AddItem "Stranded!"
    lstLevels.AddItem "Cruisn'"
    lstLevels.AddItem "Battleship Battle"
    lstLevels.AddItem "Underwater Cinematic"
    lstLevels.AddItem "The Bottom of the Sea"
    lstLevels.AddItem "Trench Trouble"
    lstLevels.AddItem "Base of the Abyss"
    lstLevels.AddItem "City Cinematic"
    lstLevels.AddItem "Metropolis Madness"
    lstLevels.AddItem "Rooftop Rumble"
    lstLevels.AddItem "The Final Encounter"
    lstLevels.AddItem "Ending"
Case "level13.ini"
    lstLevels.AddItem "Introduction"
    lstLevels.AddItem "Desert Cinematic"
    lstLevels.AddItem "Desolate Desert"
    lstLevels.AddItem "Desert Defense"
    lstLevels.AddItem "Desert Storm"
    lstLevels.AddItem "Snow Cinematic"
    lstLevels.AddItem "Steep Snow Drifts"
    lstLevels.AddItem "Mountain Pass"
    lstLevels.AddItem "Blizzard Battle"
    lstLevels.AddItem "Vacation Cinematic"
    lstLevels.AddItem "Eruptive Vacation"
    lstLevels.AddItem "Tiny Trouble"
    lstLevels.AddItem "High Protector"
    lstLevels.AddItem "Ocean Cinematic"
    lstLevels.AddItem "Stranded!"
    lstLevels.AddItem "Cruisn'"
    lstLevels.AddItem "Battleship Battle"
    lstLevels.AddItem "Underwater Cinematic"
    lstLevels.AddItem "The Bottom of the Sea"
    lstLevels.AddItem "Trench Trouble"
    lstLevels.AddItem "Base of the Abyss"
    lstLevels.AddItem "City Cinematic"
    lstLevels.AddItem "Metropolis Madness"
    lstLevels.AddItem "Rooftop Rumble"
    lstLevels.AddItem "The Final Encounter"
    lstLevels.AddItem "Ending"
    lstLevels.AddItem "Mars Patrol"
Case "level14.ini"
    lstLevels.AddItem "Introduction"
    lstLevels.AddItem "Desert Cinematic"
    lstLevels.AddItem "Desolate Desert"
    lstLevels.AddItem "Desert Defense"
    lstLevels.AddItem "Desert Storm"
    lstLevels.AddItem "Snow Cinematic"
    lstLevels.AddItem "Steep Snow Drifts"
    lstLevels.AddItem "Mountain Pass"
    lstLevels.AddItem "Blizzard Battle"
    lstLevels.AddItem "Vacation Cinematic"
    lstLevels.AddItem "Eruptive Vacation"
    lstLevels.AddItem "Tiny Trouble"
    lstLevels.AddItem "High Protector"
    lstLevels.AddItem "Ocean Cinematic"
    lstLevels.AddItem "Stranded!"
    lstLevels.AddItem "Cruisn'"
    lstLevels.AddItem "Battleship Battle"
    lstLevels.AddItem "Underwater Cinematic"
    lstLevels.AddItem "The Bottom of the Sea"
    lstLevels.AddItem "Trench Trouble"
    lstLevels.AddItem "Base of the Abyss"
    lstLevels.AddItem "City Cinematic"
    lstLevels.AddItem "Metropolis Madness"
    lstLevels.AddItem "Rooftop Rumble"
    lstLevels.AddItem "The Final Encounter"
    lstLevels.AddItem "Ending"
    lstLevels.AddItem "Mars Patrol"
    lstLevels.AddItem "Red Planet"
Case "level15.ini"
    lstLevels.AddItem "Introduction"
    lstLevels.AddItem "Desert Cinematic"
    lstLevels.AddItem "Desolate Desert"
    lstLevels.AddItem "Desert Defense"
    lstLevels.AddItem "Desert Storm"
    lstLevels.AddItem "Snow Cinematic"
    lstLevels.AddItem "Steep Snow Drifts"
    lstLevels.AddItem "Mountain Pass"
    lstLevels.AddItem "Blizzard Battle"
    lstLevels.AddItem "Vacation Cinematic"
    lstLevels.AddItem "Eruptive Vacation"
    lstLevels.AddItem "Tiny Trouble"
    lstLevels.AddItem "High Protector"
    lstLevels.AddItem "Ocean Cinematic"
    lstLevels.AddItem "Stranded!"
    lstLevels.AddItem "Cruisn'"
    lstLevels.AddItem "Battleship Battle"
    lstLevels.AddItem "Underwater Cinematic"
    lstLevels.AddItem "The Bottom of the Sea"
    lstLevels.AddItem "Trench Trouble"
    lstLevels.AddItem "Base of the Abyss"
    lstLevels.AddItem "City Cinematic"
    lstLevels.AddItem "Metropolis Madness"
    lstLevels.AddItem "Rooftop Rumble"
    lstLevels.AddItem "The Final Encounter"
    lstLevels.AddItem "Ending"
    lstLevels.AddItem "Mars Patrol"
    lstLevels.AddItem "Red Planet"
    lstLevels.AddItem "Secrets of Space"
Case Else
    lstLevels.AddItem "Introduction"
End Select
lstLevels.ListIndex = lstLevels.ListCount
lstLevels.Text = lstLevels.List(lstLevels.ListCount - 1)
lstLevels_Click
End Sub

Private Sub lblBack_Click()
'Goes back to the campaign select escreen
On Error Resume Next
'Load all of the campaign select widgets
filCampaign.Visible = True
txtNewName.Visible = True
lblLoad.Visible = True
Me.lblLoadDesc.Visible = True
Me.Shape1.Visible = True
Me.Shape2.Visible = True
Me.lblLoadLvl.Visible = True
Me.lblNewCampaign.Visible = True
lblcLvl.Visible = True
shpLoad.Visible = True
shpQuit.Visible = True
lblQuit.Visible = True

'Hide all of the level select widgets
shpLevels.Visible = False
imgMap.Visible = False
shpMap.Visible = False
lstLevels.Visible = False
shpLocation.Visible = False
shpLoadLvl.Visible = False
shpBack.Visible = False
lblLoadMission.Visible = False
lblBack.Visible = False
imgMars.Visible = False
imgCinema.Visible = False

PlySound ("Transmission")
End Sub

Private Sub lblLoad_Click()
'Load the campaign
On Error Resume Next
If filCampaign.ListCount = 0 Then 'No campaigns to load
    Exit Sub
ElseIf filCampaign.List(filCampaign.ListIndex) = "" Then 'Campaign file does not exist
    Exit Sub
End If

Dim strSave As String 'Set the last campaign loaded to this one
strSave = App.Path & "\settings.ini"
Call WriteIni("GEN", "CAMPAIGN", filCampaign.List(filCampaign.ListIndex), strSave)
PlySound ("Transmission")

Call LoadSecondScreen(filCampaign.List(filCampaign.ListIndex)) 'Load the widgets on the level select screen
End Sub

Private Sub lblLoadMission_Click()
Call lstLevels_DblClick
End Sub

Private Sub lblNewCampaign_Click()
'Create a new campaign
On Error Resume Next
Dim strCheck As String
strCheck = GetFromIni("GEN", "LEVEL", App.Path & "\" & txtNewName.Text & ".hvx")
If strCheck <> "" Or txtNewName.Text = "" Then 'Check if the campaign already exists
    PlySound ("buzzer") 'Yell at the user for being so stupid
    Exit Sub
End If
'The campaign doesn't already exist, make one
Dim strSave As String
strSave = App.Path & "\" & txtNewName.Text & ".hvx"
Call WriteIni("GEN", "LEVEL", "mov01", strSave) 'Set first level
filCampaign.Refresh
PlySound ("Transmission")
End Sub

Private Sub lblQuit_Click()
'Unload the campaign screen
On Error Resume Next
Unload Me
frmIntro.Show
frmIntro.timeDemo.Enabled = True
Demo = True
End Sub

Private Sub lstLevels_Click()
'Set the position of the circle on the map
On Error Resume Next
shpLocation.Left = 0
Select Case lstLevels.Text
Case "Desolate Desert"
    shpLocation.Top = imgMap.Top + 76
    shpLocation.Left = imgMap.Left + 290
Case "Desert Defense"
    shpLocation.Top = imgMap.Top + 76
    shpLocation.Left = imgMap.Left + 305
Case "Desert Storm"
    shpLocation.Top = imgMap.Top + 82
    shpLocation.Left = imgMap.Left + 305
Case "Steep Snow Drifts"
    shpLocation.Top = imgMap.Top + 16
    shpLocation.Left = imgMap.Left + 374
Case "Mountain Pass"
    shpLocation.Top = imgMap.Top + 16
    shpLocation.Left = imgMap.Left + 359
Case "Blizzard Battle"
    shpLocation.Top = imgMap.Top + 12
    shpLocation.Left = imgMap.Left + 349
Case "Eruptive Vacation"
    shpLocation.Top = imgMap.Top + 89
    shpLocation.Left = imgMap.Left + 12
Case "Tiny Trouble"
    shpLocation.Top = imgMap.Top + 89
    shpLocation.Left = imgMap.Left + 14
Case "High Protector"
    shpLocation.Top = imgMap.Top + 91
    shpLocation.Left = imgMap.Left + 15
Case "Stranded!"
    shpLocation.Top = imgMap.Top + 91
    shpLocation.Left = imgMap.Left + 25
Case "Cruisn'"
    shpLocation.Top = imgMap.Top + 115
    shpLocation.Left = imgMap.Left + 19
Case "Battleship Battle"
    shpLocation.Top = imgMap.Top + 120
    shpLocation.Left = imgMap.Left + 16
Case "The Bottom of the Sea"
    shpLocation.Top = imgMap.Top + 146
    shpLocation.Left = imgMap.Left + 30
Case "Trench Trouble"
    shpLocation.Top = imgMap.Top + 113
    shpLocation.Left = imgMap.Left + 60
Case "Base of the Abyss"
    shpLocation.Top = imgMap.Top + 115
    shpLocation.Left = imgMap.Left + 70
Case "Metropolis Madness"
    shpLocation.Top = imgMap.Top + 50
    shpLocation.Left = imgMap.Left + 250
Case "Rooftop Rumble"
    shpLocation.Top = imgMap.Top + 50
    shpLocation.Left = imgMap.Left + 254
Case "The Final Encounter"
    shpLocation.Top = imgMap.Top + 52
    shpLocation.Left = imgMap.Left + 252
End Select
If shpLocation.Left > 0 Then 'A location was found on the earth map
    shpLocation.Visible = True
    imgCinema.Visible = False
    imgMap.Visible = True
    imgMars.Visible = False
Else 'Check if it's a cinematic
    imgCinema.Left = 0
    Select Case lstLevels.Text
        Case "Introduction"
            imgCinema.Left = imgMap.Left + 175
            imgCinema.Top = imgMap.Top + 75
        Case "Desert Cinematic"
            imgCinema.Left = imgMap.Left + 305
            imgCinema.Top = imgMap.Top + 76
        Case "Snow Cinematic"
            imgCinema.Left = imgMap.Left + 364
            imgCinema.Top = imgMap.Top + 16
        Case "Vacation Cinematic"
            imgCinema.Left = imgMap.Left + 12
            imgCinema.Top = imgCinema.Top + 89
        Case "Ocean Cinematic"
            imgCinema.Left = imgMap.Left + 23
            imgCinema.Top = imgMap.Top + 90
        Case "Underwater Cinematic"
            imgCinema.Left = imgMap.Left + 25
            imgCinema.Top = imgMap.Top + 91
        Case "City Cinematic"
            imgCinema.Left = imgMap.Left + 250
            imgCinema.Top = imgCinema.Top + 50
        Case "Ending"
            imgCinema.Left = imgMap.Left + 250
            imgCinema.Top = imgCinema.Top + 52
    End Select
    imgCinema.Visible = True
    shpLocation.Visible = False
End If
If imgCinema.Left <> 0 Then 'Cinema found on the earth map
    imgMap.Visible = True
    imgMars.Visible = False
Else 'Check the mars map
    imgMap.Visible = False
    imgMars.Visible = True
    shpLocation.Visible = True
    imgCinema.Visible = False
    shpLocation.Visible = True
    imgCinema.Visible = False
    Select Case lstLevels.Text
        Case "Mars Patrol"
            shpLocation.Left = imgMars.Left + 130
            shpLocation.Top = imgMars.Top + 100
        Case "Red Planet"
            shpLocation.Left = imgMars.Left + 150
            shpLocation.Top = imgMars.Top + 105
        Case "Secrets of Space"
            shpLocation.Left = imgMars.Left + 170
            shpLocation.Top = imgMars.Top + 103
    End Select
End If
End Sub

Private Sub lstLevels_DblClick()
'Get the next level file from the name
On Error Resume Next
NextLevel = ""
Select Case lstLevels.Text
Case "Desolate Desert"
    NextLevel = "level1.ini"
Case "Desert Defense"
    NextLevel = "level2.ini"
Case "Desert Storm"
    NextLevel = "dboss.ini"
Case "Steep Snow Drifts"
    NextLevel = "level3.ini"
Case "Mountain Pass"
    NextLevel = "level4.ini"
Case "Blizzard Battle"
    NextLevel = "iceboss.ini"
Case "Eruptive Vacation"
    NextLevel = "level5.ini"
Case "Tiny Trouble"
    NextLevel = "level6.ini"
Case "High Protector"
    NextLevel = "vboss.ini"
Case "Stranded!"
    NextLevel = "level7.ini"
Case "Cruisin'"
    NextLevel = "level8.ini"
Case "Battleship Battle"
    NextLevel = "oboss.ini"
Case "The Bottom of the Sea"
    NextLevel = "level9.ini"
Case "Trench Trouble"
    NextLevel = "level10.ini"
Case "Base of the Abyss"
    NextLevel = "uboss.ini"
Case "Metropolis Madness"
    NextLevel = "level11.ini"
Case "Rooftop Rumble"
    NextLevel = "level12.ini"
Case "The Final Encounter"
    NextLevel = "cboss.ini"
Case "Mars Patrol"
    NextLevel = "level13.ini"
Case "Red Planet"
    NextLevel = "level14.ini"
Case "Secrets of Space"
    NextLevel = "level15.ini"
End Select
If NextLevel <> "" Then
    PlySound ("Transmission")
    loadLevel = True
    Unload Me
    frmBriefing.Show
Else 'Cinematic
    Select Case lstLevels.Text
        Case "Introduction"
            anType = 0
        Case "Desert Cinematic"
            anType = 1
        Case "Snow Cinematic"
            anType = 2
        Case "Vacation Cinematic"
            anType = 3
        Case "Ocean Cinematic"
            anType = 4
        Case "Underwater Cinematic"
            anType = 5
        Case "City Cinematic"
            anType = 6
        Case "Ending"
            anType = 7
    End Select
    loadLevel = True
    PlySound ("transmission")
    Unload Me
    frmCinema.Show
End If
End Sub

Private Sub txtNewName_Click()
On Error Resume Next
If txtNewName.Text = "New Campaign Name" Then 'Get rid of "New Campaign Name" on click
    txtNewName.Text = ""
End If
End Sub

Private Sub txtNewName_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then 'Call the label
    lblNewCampaign_Click
End If
End Sub
Private Sub LoadSecondScreen(strFile As String)
'Hides the widgets of the select campaign screen, shows the widgets of the select level screen
On Error Resume Next
Dim strSave As String
strSave = App.Path & "\" & strFile
curCampaign = strFile
Call PopulateList(GetFromIni("GEN", "LEVEL", strSave))

filCampaign.Visible = False
txtNewName.Visible = False
lblLoad.Visible = False
Me.lblLoadDesc.Visible = False
Me.Shape1.Visible = False
Me.Shape2.Visible = False
Me.lblLoadLvl.Visible = False
Me.lblNewCampaign.Visible = False
lblcLvl.Visible = False
shpLoad.Visible = False
shpQuit.Visible = False
lblQuit.Visible = False

shpLevels.Visible = True
imgMap.Visible = True
shpMap.Visible = True
lstLevels.Visible = True
shpLoadLvl.Visible = True
shpBack.Visible = True
lblLoadMission.Visible = True
lblBack.Visible = True
End Sub
