VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmOptions 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Hover Tanx Options"
   ClientHeight    =   6495
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   9120
   Icon            =   "frmOptions.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   433
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   608
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkWater 
      Caption         =   "Force To Water"
      Height          =   255
      Left            =   3120
      TabIndex        =   27
      Top             =   240
      Width           =   1575
   End
   Begin VB.CheckBox chkTop 
      Caption         =   "Force To Top"
      Height          =   255
      Left            =   4800
      TabIndex        =   28
      Top             =   240
      Width           =   1575
   End
   Begin VB.TextBox txtMidi 
      Height          =   285
      Left            =   2760
      MaxLength       =   100
      TabIndex        =   13
      Text            =   "nomidi.mid"
      Top             =   6120
      Width           =   1215
   End
   Begin VB.TextBox txtBriefing 
      Height          =   285
      Index           =   2
      Left            =   480
      MaxLength       =   250
      TabIndex        =   11
      Top             =   5520
      Width           =   8295
   End
   Begin VB.TextBox txtBriefing 
      Height          =   285
      Index           =   1
      Left            =   480
      MaxLength       =   250
      TabIndex        =   10
      Top             =   5160
      Width           =   8295
   End
   Begin VB.TextBox txtNextLevel 
      Height          =   285
      Left            =   480
      MaxLength       =   100
      TabIndex        =   12
      Text            =   "nolevel.ini"
      Top             =   6120
      Width           =   1215
   End
   Begin MSComDlg.CommonDialog comDiag 
      Left            =   8040
      Top             =   4080
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Filter          =   "Hover Tanx Level Files | *.ini"
   End
   Begin VB.CheckBox chkBottom 
      Caption         =   "Force To Bottom"
      Height          =   255
      Left            =   1560
      TabIndex        =   21
      Top             =   240
      Width           =   1575
   End
   Begin VB.CheckBox chkHard 
      Caption         =   "Hard Mode"
      Height          =   255
      Left            =   5520
      TabIndex        =   20
      Top             =   4080
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.ComboBox cmbBoss 
      Height          =   315
      ItemData        =   "frmOptions.frx":0D2A
      Left            =   3480
      List            =   "frmOptions.frx":0D43
      TabIndex        =   19
      Text            =   "Tank"
      Top             =   4080
      Width           =   1935
   End
   Begin VB.CheckBox chkBoss 
      Caption         =   "Boss At End of Level"
      Height          =   255
      Left            =   3000
      TabIndex        =   17
      Top             =   3720
      Width           =   1815
   End
   Begin VB.CommandButton cmdTypeUpdate 
      Caption         =   "Update &Type"
      Height          =   255
      Left            =   5880
      TabIndex        =   16
      Top             =   3240
      Width           =   1335
   End
   Begin VB.ComboBox cmbLevelType 
      Height          =   315
      ItemData        =   "frmOptions.frx":0D74
      Left            =   3960
      List            =   "frmOptions.frx":0D8D
      TabIndex        =   15
      Text            =   "Desert"
      Top             =   3240
      Width           =   1935
   End
   Begin VB.TextBox txtBriefing 
      Height          =   285
      Index           =   0
      Left            =   480
      MaxLength       =   250
      TabIndex        =   9
      Top             =   4800
      Width           =   8295
   End
   Begin VB.TextBox txtName 
      Height          =   285
      Left            =   480
      MaxLength       =   100
      TabIndex        =   7
      Text            =   "Untitled"
      Top             =   4200
      Width           =   1215
   End
   Begin VB.CommandButton cmdWidthUpdate 
      Caption         =   "&Update"
      Height          =   255
      Left            =   960
      TabIndex        =   5
      Top             =   3480
      Width           =   975
   End
   Begin VB.TextBox txtWidth 
      Height          =   285
      Left            =   480
      MaxLength       =   2
      TabIndex        =   4
      Text            =   "10"
      Top             =   3480
      Width           =   375
   End
   Begin VB.HScrollBar hSprite 
      Height          =   255
      Left            =   360
      Max             =   2
      TabIndex        =   0
      Top             =   2400
      Width           =   1935
   End
   Begin VB.Frame framSprites 
      Caption         =   "Sprites 3"
      Height          =   2895
      Index           =   2
      Left            =   4320
      TabIndex        =   26
      Top             =   0
      Visible         =   0   'False
      Width           =   4695
      Begin VB.Image imgSprite 
         Height          =   570
         Index           =   55
         Left            =   3960
         Picture         =   "frmOptions.frx":0DC7
         Top             =   2280
         Visible         =   0   'False
         Width           =   1305
      End
      Begin VB.Image imgSprite 
         Height          =   570
         Index           =   54
         Left            =   3840
         Picture         =   "frmOptions.frx":130B
         Top             =   1680
         Visible         =   0   'False
         Width           =   1305
      End
      Begin VB.Image imgSprite 
         Height          =   210
         Index           =   53
         Left            =   4080
         Picture         =   "frmOptions.frx":184E
         Top             =   1320
         Visible         =   0   'False
         Width           =   330
      End
      Begin VB.Image imgSprite 
         Height          =   105
         Index           =   52
         Left            =   4080
         Picture         =   "frmOptions.frx":1BD2
         Top             =   1200
         Visible         =   0   'False
         Width           =   345
      End
      Begin VB.Image imgSprite 
         Height          =   210
         Index           =   51
         Left            =   4080
         Picture         =   "frmOptions.frx":1F38
         Top             =   960
         Visible         =   0   'False
         Width           =   330
      End
      Begin VB.Image imgSprite 
         Height          =   285
         Index           =   50
         Left            =   0
         Picture         =   "frmOptions.frx":22BD
         Top             =   2040
         Visible         =   0   'False
         Width           =   8970
      End
      Begin VB.Image imgSprite 
         Height          =   300
         Index           =   49
         Left            =   0
         Picture         =   "frmOptions.frx":26C2
         Top             =   2160
         Visible         =   0   'False
         Width           =   4860
      End
      Begin VB.Image imgSprite 
         Height          =   300
         Index           =   48
         Left            =   0
         Picture         =   "frmOptions.frx":2A9E
         Top             =   2520
         Visible         =   0   'False
         Width           =   3015
      End
      Begin VB.Image imgSprite 
         Height          =   45
         Index           =   47
         Left            =   3360
         Picture         =   "frmOptions.frx":2E5D
         Top             =   2520
         Visible         =   0   'False
         Width           =   300
      End
      Begin VB.Image imgSprite 
         Height          =   210
         Index           =   46
         Left            =   2760
         Picture         =   "frmOptions.frx":3193
         Top             =   2520
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.Image imgSprite 
         Height          =   1050
         Index           =   45
         Left            =   1920
         Picture         =   "frmOptions.frx":3518
         Top             =   1680
         Visible         =   0   'False
         Width           =   570
      End
      Begin VB.Image imgSprite 
         Height          =   195
         Index           =   44
         Left            =   1680
         Picture         =   "frmOptions.frx":39BE
         Top             =   1920
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.Image imgSprite 
         Height          =   315
         Index           =   43
         Left            =   1320
         Picture         =   "frmOptions.frx":3D39
         Top             =   1920
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.Image imgSprite 
         Height          =   435
         Index           =   42
         Left            =   3480
         Picture         =   "frmOptions.frx":4126
         Top             =   1200
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.Image imgSprite 
         Height          =   900
         Index           =   41
         Left            =   3120
         Picture         =   "frmOptions.frx":420A
         Top             =   1200
         Visible         =   0   'False
         Width           =   285
      End
      Begin VB.Image imgSprite 
         Height          =   900
         Index           =   40
         Left            =   2760
         Picture         =   "frmOptions.frx":466C
         Top             =   1200
         Visible         =   0   'False
         Width           =   285
      End
      Begin VB.Image imgSprite 
         Height          =   180
         Index           =   39
         Left            =   2520
         Picture         =   "frmOptions.frx":4ACE
         Top             =   1320
         Visible         =   0   'False
         Width           =   165
      End
      Begin VB.Image imgSprite 
         Height          =   1470
         Index           =   38
         Left            =   480
         Picture         =   "frmOptions.frx":4E42
         Top             =   1080
         Visible         =   0   'False
         Width           =   765
      End
      Begin VB.Image imgSprite 
         Height          =   225
         Index           =   37
         Left            =   0
         Picture         =   "frmOptions.frx":5486
         Top             =   2040
         Visible         =   0   'False
         Width           =   705
      End
      Begin VB.Image imgSprite 
         Height          =   285
         Index           =   36
         Left            =   2400
         Picture         =   "frmOptions.frx":581C
         Top             =   840
         Visible         =   0   'False
         Width           =   1065
      End
      Begin VB.Image imgSprite 
         Height          =   285
         Index           =   35
         Left            =   3240
         Picture         =   "frmOptions.frx":5C6F
         Top             =   360
         Visible         =   0   'False
         Width           =   870
      End
      Begin VB.Image imgSprite 
         Height          =   360
         Index           =   34
         Left            =   2520
         Picture         =   "frmOptions.frx":6050
         Top             =   360
         Visible         =   0   'False
         Width           =   645
      End
      Begin VB.Image imgSprite 
         Height          =   750
         Index           =   33
         Left            =   240
         Picture         =   "frmOptions.frx":649D
         Top             =   240
         Visible         =   0   'False
         Width           =   2235
      End
   End
   Begin VB.Frame framSprites 
      Caption         =   "Sprites 4"
      Height          =   2895
      Index           =   3
      Left            =   4320
      TabIndex        =   29
      Top             =   0
      Visible         =   0   'False
      Width           =   4695
      Begin VB.Image imgSprite 
         Height          =   300
         Index           =   71
         Left            =   120
         Picture         =   "frmOptions.frx":6A80
         Top             =   2520
         Visible         =   0   'False
         Width           =   1380
      End
      Begin VB.Image imgSprite 
         Height          =   1515
         Index           =   70
         Left            =   3840
         Picture         =   "frmOptions.frx":6F61
         Top             =   2400
         Visible         =   0   'False
         Width           =   795
      End
      Begin VB.Image imgSprite 
         Height          =   1890
         Index           =   69
         Left            =   3240
         Picture         =   "frmOptions.frx":741A
         Top             =   2400
         Visible         =   0   'False
         Width           =   600
      End
      Begin VB.Image imgSprite 
         Height          =   210
         Index           =   68
         Left            =   3480
         Picture         =   "frmOptions.frx":7CD0
         Top             =   240
         Visible         =   0   'False
         Width           =   600
      End
      Begin VB.Image imgSprite 
         Height          =   1890
         Index           =   67
         Left            =   4080
         Picture         =   "frmOptions.frx":8062
         Top             =   240
         Visible         =   0   'False
         Width           =   600
      End
      Begin VB.Image imgSprite 
         Height          =   780
         Index           =   66
         Left            =   1920
         Picture         =   "frmOptions.frx":891F
         Top             =   2040
         Visible         =   0   'False
         Width           =   1005
      End
      Begin VB.Image imgSprite 
         Height          =   1020
         Index           =   65
         Left            =   3000
         Picture         =   "frmOptions.frx":8E24
         Top             =   1200
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Image imgSprite 
         Height          =   765
         Index           =   64
         Left            =   2160
         Picture         =   "frmOptions.frx":9888
         Top             =   1200
         Visible         =   0   'False
         Width           =   765
      End
      Begin VB.Image imgSprite 
         Height          =   345
         Index           =   63
         Left            =   1800
         Picture         =   "frmOptions.frx":9E1E
         Top             =   1440
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Image imgSprite 
         Height          =   660
         Index           =   62
         Left            =   1320
         Picture         =   "frmOptions.frx":A1ED
         Top             =   1440
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Image imgSprite 
         Height          =   285
         Index           =   61
         Left            =   2640
         Picture         =   "frmOptions.frx":A6CA
         Top             =   840
         Visible         =   0   'False
         Width           =   915
      End
      Begin VB.Image imgSprite 
         Height          =   255
         Index           =   60
         Left            =   1320
         Picture         =   "frmOptions.frx":AB12
         Top             =   1080
         Visible         =   0   'False
         Width           =   45
      End
      Begin VB.Image imgSprite 
         Height          =   375
         Index           =   59
         Left            =   2640
         Picture         =   "frmOptions.frx":AE47
         Top             =   240
         Visible         =   0   'False
         Width           =   645
      End
      Begin VB.Image imgSprite 
         Height          =   810
         Index           =   58
         Left            =   1200
         Picture         =   "frmOptions.frx":B2B0
         Top             =   120
         Visible         =   0   'False
         Width           =   1260
      End
      Begin VB.Image imgSprite 
         Height          =   2400
         Index           =   57
         Left            =   600
         Picture         =   "frmOptions.frx":B78A
         Top             =   120
         Visible         =   0   'False
         Width           =   420
      End
      Begin VB.Image imgSprite 
         Height          =   2400
         Index           =   56
         Left            =   120
         Picture         =   "frmOptions.frx":BCF9
         Top             =   120
         Visible         =   0   'False
         Width           =   420
      End
   End
   Begin VB.Frame framSprites 
      Caption         =   "Sprites 2"
      Height          =   2895
      Index           =   1
      Left            =   4320
      TabIndex        =   25
      Top             =   0
      Visible         =   0   'False
      Width           =   4695
      Begin VB.Image imgSprite 
         Height          =   285
         Index           =   32
         Left            =   2880
         Picture         =   "frmOptions.frx":C262
         Top             =   2400
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Image imgSprite 
         Height          =   450
         Index           =   31
         Left            =   3600
         Picture         =   "frmOptions.frx":C71E
         Top             =   1800
         Visible         =   0   'False
         Width           =   525
      End
      Begin VB.Image imgSprite 
         Height          =   450
         Index           =   30
         Left            =   3600
         Picture         =   "frmOptions.frx":C88F
         Top             =   1320
         Visible         =   0   'False
         Width           =   525
      End
      Begin VB.Image imgSprite 
         Height          =   975
         Index           =   29
         Left            =   3000
         Picture         =   "frmOptions.frx":CA00
         Top             =   1320
         Visible         =   0   'False
         Width           =   570
      End
      Begin VB.Image imgSprite 
         Height          =   210
         Index           =   28
         Left            =   2400
         Picture         =   "frmOptions.frx":CC16
         Top             =   2520
         Visible         =   0   'False
         Width           =   225
      End
      Begin VB.Image imgSprite 
         Height          =   240
         Index           =   27
         Left            =   1920
         Picture         =   "frmOptions.frx":CF88
         Top             =   2520
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Image imgSprite 
         Height          =   225
         Index           =   26
         Left            =   1680
         Picture         =   "frmOptions.frx":D2F0
         Top             =   2520
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.Image imgSprite 
         Height          =   1605
         Index           =   25
         Left            =   4320
         Picture         =   "frmOptions.frx":D677
         Top             =   120
         Visible         =   0   'False
         Width           =   360
      End
      Begin VB.Image imgSprite 
         Height          =   420
         Index           =   24
         Left            =   4320
         Picture         =   "frmOptions.frx":DB76
         Top             =   1800
         Visible         =   0   'False
         Width           =   360
      End
      Begin VB.Image imgSprite 
         Height          =   300
         Index           =   23
         Left            =   240
         Picture         =   "frmOptions.frx":DF12
         Top             =   2400
         Visible         =   0   'False
         Width           =   1380
      End
      Begin VB.Image imgSprite 
         Height          =   1050
         Index           =   22
         Left            =   3000
         Picture         =   "frmOptions.frx":E3F3
         Top             =   240
         Visible         =   0   'False
         Width           =   1155
      End
      Begin VB.Image imgSprite 
         Height          =   2130
         Index           =   21
         Left            =   2400
         Picture         =   "frmOptions.frx":E947
         Top             =   240
         Visible         =   0   'False
         Width           =   585
      End
      Begin VB.Image imgSprite 
         Height          =   2130
         Index           =   20
         Left            =   1800
         Picture         =   "frmOptions.frx":EE7E
         Top             =   240
         Visible         =   0   'False
         Width           =   585
      End
      Begin VB.Image imgSprite 
         Height          =   2130
         Index           =   19
         Left            =   1440
         Picture         =   "frmOptions.frx":F389
         Top             =   240
         Visible         =   0   'False
         Width           =   345
      End
      Begin VB.Image imgSprite 
         Height          =   2025
         Index           =   16
         Left            =   240
         Picture         =   "frmOptions.frx":F548
         Top             =   240
         Visible         =   0   'False
         Width           =   465
      End
      Begin VB.Image imgSprite 
         Height          =   1770
         Index           =   17
         Left            =   600
         Picture         =   "frmOptions.frx":FB76
         Top             =   240
         Visible         =   0   'False
         Width           =   465
      End
      Begin VB.Image imgSprite 
         Height          =   675
         Index           =   18
         Left            =   960
         Picture         =   "frmOptions.frx":101A5
         Top             =   240
         Visible         =   0   'False
         Width           =   450
      End
   End
   Begin VB.Frame framSprites 
      Caption         =   "Frame1"
      Height          =   2895
      Index           =   0
      Left            =   4320
      TabIndex        =   24
      Top             =   0
      Visible         =   0   'False
      Width           =   4695
      Begin VB.Image imgSprite 
         Height          =   1800
         Index           =   0
         Left            =   0
         Picture         =   "frmOptions.frx":103D2
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Image imgSprite 
         Height          =   300
         Index           =   1
         Left            =   960
         Picture         =   "frmOptions.frx":108B9
         Top             =   1560
         Visible         =   0   'False
         Width           =   1185
      End
      Begin VB.Image imgSprite 
         Height          =   255
         Index           =   2
         Left            =   480
         Picture         =   "frmOptions.frx":10D95
         Top             =   1560
         Visible         =   0   'False
         Width           =   420
      End
      Begin VB.Image imgSprite 
         Height          =   795
         Index           =   3
         Left            =   360
         Picture         =   "frmOptions.frx":11135
         Top             =   0
         Visible         =   0   'False
         Width           =   585
      End
      Begin VB.Image imgSprite 
         Height          =   435
         Index           =   4
         Left            =   960
         Picture         =   "frmOptions.frx":11582
         Top             =   120
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.Image imgSprite 
         Height          =   435
         Index           =   5
         Left            =   1440
         Picture         =   "frmOptions.frx":11666
         Top             =   120
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.Image imgSprite 
         Height          =   435
         Index           =   6
         Left            =   960
         Picture         =   "frmOptions.frx":1174A
         Top             =   600
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.Image imgSprite 
         Height          =   600
         Index           =   7
         Left            =   1680
         Picture         =   "frmOptions.frx":1182E
         Top             =   600
         Visible         =   0   'False
         Width           =   690
      End
      Begin VB.Image imgSprite 
         Height          =   375
         Index           =   8
         Left            =   2400
         Picture         =   "frmOptions.frx":11C92
         Top             =   0
         Visible         =   0   'False
         Width           =   285
      End
      Begin VB.Image imgSprite 
         Height          =   495
         Index           =   9
         Left            =   2760
         Picture         =   "frmOptions.frx":12042
         Top             =   0
         Visible         =   0   'False
         Width           =   810
      End
      Begin VB.Image imgSprite 
         Height          =   450
         Index           =   10
         Left            =   2400
         Picture         =   "frmOptions.frx":1248F
         Top             =   720
         Visible         =   0   'False
         Width           =   285
      End
      Begin VB.Image imgSprite 
         Height          =   300
         Index           =   11
         Left            =   3720
         Picture         =   "frmOptions.frx":1285D
         Top             =   0
         Visible         =   0   'False
         Width           =   1185
      End
      Begin VB.Image imgSprite 
         Height          =   450
         Index           =   12
         Left            =   2880
         Picture         =   "frmOptions.frx":12D39
         Top             =   600
         Visible         =   0   'False
         Width           =   525
      End
      Begin VB.Image imgSprite 
         Height          =   1395
         Index           =   13
         Left            =   3360
         Picture         =   "frmOptions.frx":12EAA
         Top             =   360
         Visible         =   0   'False
         Width           =   2970
      End
      Begin VB.Image imgSprite 
         Height          =   750
         Index           =   14
         Left            =   0
         Picture         =   "frmOptions.frx":13EC9
         Top             =   2040
         Visible         =   0   'False
         Width           =   2235
      End
      Begin VB.Image imgSprite 
         Height          =   750
         Index           =   15
         Left            =   2280
         Picture         =   "frmOptions.frx":144AC
         Top             =   2040
         Visible         =   0   'False
         Width           =   2235
      End
   End
   Begin VB.Label lblGen 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Midi file (song.mid):"
      Height          =   195
      Index           =   8
      Left            =   2760
      TabIndex        =   23
      Top             =   5880
      Width           =   1335
   End
   Begin VB.Label lblGen 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Next Level To Load (level.ini):"
      Height          =   195
      Index           =   6
      Left            =   480
      TabIndex        =   22
      Top             =   5880
      Width           =   2115
   End
   Begin VB.Label lblGen 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Boss:"
      Height          =   195
      Index           =   5
      Left            =   3000
      TabIndex        =   18
      Top             =   4080
      Width           =   390
   End
   Begin VB.Label lblGen 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Level Type:"
      Height          =   195
      Index           =   4
      Left            =   3000
      TabIndex        =   14
      Top             =   3240
      Width           =   840
   End
   Begin VB.Label lblGen 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Briefing:"
      Height          =   195
      Index           =   3
      Left            =   480
      TabIndex        =   8
      Top             =   4560
      Width           =   570
   End
   Begin VB.Label lblGen 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Level Name:"
      Height          =   195
      Index           =   2
      Left            =   480
      TabIndex        =   6
      Top             =   3960
      Width           =   900
   End
   Begin VB.Label lblGen 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Level Width In Screens:"
      Height          =   195
      Index           =   1
      Left            =   480
      TabIndex        =   3
      Top             =   3240
      Width           =   1710
   End
   Begin VB.Label lblSDesc 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "One Hit Wall"
      Height          =   435
      Left            =   360
      TabIndex        =   2
      Top             =   2760
      Width           =   1845
   End
   Begin VB.Label lblGen 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Sprites:"
      Height          =   195
      Index           =   0
      Left            =   960
      TabIndex        =   1
      Top             =   240
      Width           =   525
   End
   Begin VB.Image imgSpriteShow 
      Height          =   1800
      Left            =   1080
      Picture         =   "frmOptions.frx":14A8F
      Top             =   480
      Width           =   375
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuNew 
         Caption         =   "&New"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuOpen 
         Caption         =   "&Open"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuSave 
         Caption         =   "&Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuPlay 
         Caption         =   "&Play Level"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu mnuOther 
      Caption         =   "&Other Options"
      Begin VB.Menu mnuTankEditor 
         Caption         =   "&Tank Editor"
      End
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmbLevelType_Change()
'Change the color of the foreground to reflect different level types
If cmbLevelType.Text = "Desert" Then
    frmEditor.shpGround.FillColor = COLDESERT
ElseIf cmbLevelType.Text = "Snow" Then
    frmEditor.shpGround.FillColor = COLSNOW
ElseIf cmbLevelType.Text = "Volcano" Then
    frmEditor.shpGround.FillColor = COLSOIL
ElseIf cmbLevelType.Text = "Ocean" Then
    frmEditor.shpGround.FillColor = COLWATERB
ElseIf cmbLevelType.Text = "Underwater" Then
    frmEditor.shpGround.FillColor = COLWATER
ElseIf cmbLevelType.Text = "City" Then
    frmEditor.shpGround.FillColor = COLCITY
ElseIf cmbLevelType.Text = "Mars" Then
    frmEditor.shpGround.FillColor = COLMARS
End If
End Sub

Private Sub cmdWidthUpdate_Click()
'Change the width of the editor's scroll bar depending on the txtWidth field
On Error Resume Next
Dim intLevelWidth As Long
intLevelWidth = CLng(txtWidth.Text)
If frmEditor.hScroll.Value > intLevelWidth Then
    frmEditor.hScroll.Value = 0
End If
frmEditor.hScroll.Max = intLevelWidth

End Sub

Private Sub Form_Load()
'Set the scroll bar to scroll through the sprite to the maximum number possible
hSprite.Max = imgSprite.UBound
End Sub

Private Sub hSprite_Change()
On Error Resume Next
imgSpriteShow.Picture = imgSprite(hSprite.Value).Picture 'Change the current picture
'Update the label with the description of the sprite
Select Case hSprite.Value
Case 0
    lblSDesc.Caption = "One Hit Wall"
Case 1
    lblSDesc.Caption = "Ground Missile"
Case 2
    lblSDesc.Caption = "Land-Mine Rock"
Case 3
    lblSDesc.Caption = "Drop Bomb"
Case 4
    lblSDesc.Caption = "Air to High Air Rock"
Case 5
    lblSDesc.Caption = "Air To Mid Air Rock"
Case 6
    lblSDesc.Caption = "Air To Ground Rock"
Case 7
    lblSDesc.Caption = "Exploding Arcing Rock"
Case 8
    lblSDesc.Caption = "Health Powerup"
Case 9
    lblSDesc.Caption = "Arcing Bullet"
Case 10
    lblSDesc.Caption = "Flame"
Case 11
    lblSDesc.Caption = "Air To Ground Missile"
Case 12
    lblSDesc.Caption = "Short Range Volcanic Arcing Rock"
Case 13
    lblSDesc.Caption = "Nuclear Truck"
Case 14
    lblSDesc.Caption = "Bomber Coptor"
Case 15
    lblSDesc.Caption = "Suicide Coptor"
Case 16
    lblSDesc.Caption = "Ice Wall"
Case 17
    lblSDesc.Caption = "Damaged Ice Wall"
Case 18
    lblSDesc.Caption = "Falling Ice"
Case 19
    lblSDesc.Caption = "Window"
Case 20
    lblSDesc.Caption = "Strong Pillar"
Case 21
    lblSDesc.Caption = "Damaged Pillar"
Case 22
    lblSDesc.Caption = "Door"
Case 23
    lblSDesc.Caption = "UFO"
Case 24
    lblSDesc.Caption = "Firing Cannon (Rest State)"
Case 25
    lblSDesc.Caption = "Firing Cannon (Fire State)"
Case 26
    lblSDesc.Caption = "Form Shake"
Case 27
    lblSDesc.Caption = "UFO Laser"
Case 28
    lblSDesc.Caption = "Engage/Disengage All Range Mode"
Case 29
    lblSDesc.Caption = "Ice Drill"
Case 30
    lblSDesc.Caption = "Medium Range Arcing Rock"
Case 31
    lblSDesc.Caption = "Long Range Arcing Rock"
Case 32
    lblSDesc.Caption = "Air To Ground Skimming Missile"
Case 33
    lblSDesc.Caption = "Helicoptor Firing Skimming Missile"
Case 34
    lblSDesc.Caption = "Torpedo"
Case 35
    lblSDesc.Caption = "Rocket Torpedo (Off)"
Case 36
    lblSDesc.Caption = "Rocket Torpedo (On)"
Case 37
    lblSDesc.Caption = "Island"
Case 38
    lblSDesc.Caption = "Palm Tree"
Case 39
    lblSDesc.Caption = "Rapid Fire Power-Up"
Case 40
    lblSDesc.Caption = "Ground To Air To Ground Missile"
Case 41
    lblSDesc.Caption = "Ground To High Air To Ground Missile"
Case 42
    lblSDesc.Caption = "Bouncing Rock"
Case 43
    lblSDesc.Caption = "Plasma Ball Large"
Case 44
    lblSDesc.Caption = "Plasma Ball Small"
Case 45
    lblSDesc.Caption = "Turret"
Case 46
    lblSDesc.Caption = "Destroyed Turret"
Case 47
    lblSDesc.Caption = "Laser"
Case 48
    lblSDesc.Caption = "Small Gap"
Case 49
    lblSDesc.Caption = "Medium Gap"
Case 50
    lblSDesc.Caption = "Large Gap"
Case 51
    lblSDesc.Caption = "Arcing Bullet Down"
Case 52
    lblSDesc.Caption = "Arcing Bullet Level"
Case 53
    lblSDesc.Caption = "Arcing Bullet Up"
Case 54
    lblSDesc.Caption = "Plane Left"
Case 55
    lblSDesc.Caption = "Plane Right"
Case 56
    lblSDesc.Caption = "Celing Turet"
Case 57
    lblSDesc.Caption = "Floor Turet"
Case 58
    lblSDesc.Caption = "Submarine"
Case 59
    lblSDesc.Caption = "Back and Forth Dive Bomb"
Case 60
    lblSDesc.Caption = "Vertical Laser"
Case 61
    lblSDesc.Caption = "Heat Seeking Missile"
End Select
End Sub

Private Sub mnuExit_Click()
'Quit the editor
Unload Me
frmEditor.Hide
frmIntro.Show
End Sub

Private Sub mnuNew_Click()
'Create a new level
Dim intChoice As Long
intChoice = MsgBox("Are you sure you want to make a new level?", vbYesNo, "Are you sure?")
If intChoice = vbYes Then 'User clicked yes, make a new level
    txtWidth.Text = "10" 'Set default width to 10
    frmEditor.hScroll.Max = 10
    frmEditor.hScroll.Value = 0
    cmbLevelType.Text = "Desert" 'Set the level type of desert
    chkBoss.Value = 0 'No boss
    txtBriefing(0).Text = "" 'No briefing
    txtNextLevel.Text = "nolevel.ini" 'No next level
    txtMidi.Text = "nomidi.mid" 'No music
    For i = 1 To 100 'Set all sprites to invisible
        SpriteVisible(i) = False
        frmEditor.imgSprite(i).Visible = False
        If i <= 2 Then
            txtBriefing(i).Text = ""
        End If
    Next 'i
End If
End Sub

Private Sub mnuOpen_Click()
'Open a level
On Error Resume Next
comDiag.ShowOpen
Dim strSave As String
Dim strTemp As String
strSave = comDiag.FileName 'Open a dialog window to determine which file to open
txtName.Text = GetFromIni("GEN", "NAME", strSave) 'Get the name of the level
For i = 0 To 2 'Load breifing options
    txtBriefing(i).Text = GetFromIni("GEN", "BRIEFING" & i, strSave)
Next 'i
cmbLevelType.Text = GetFromIni("GEN", "TYPE", strSave) 'Load level type
'Set foreground color
If cmbLevelType.Text = "Desert" Then
    frmEditor.shpGround.FillColor = COLDESERT
ElseIf cmbLevelType.Text = "Snow" Then
    frmEditor.shpGround.FillColor = COLSNOW
ElseIf cmbLevelType.Text = "Volcano" Then
    frmEditor.shpGround.FillColor = COLSOIL
ElseIf cmbLevelType.Text = "Ocean" Then
    frmEditor.shpGround.FillColor = COLWATERB
ElseIf cmbLevelType.Text = "Underwater" Then
    frmEditor.shpGround.FillColor = COLWATER
ElseIf cmbLevelType.Text = "City" Then
    frmEditor.shpGround.FillColor = COLCITY
ElseIf cmbLevelType.Text = "Mars" Then
    frmEditor.shpGround.FillColor = COLMARS
End If
'Load various level information
cmbBoss.Text = GetFromIni("GEN", "BOSSNAME", strSave)
txtWidth.Text = GetFromIni("GEN", "WIDTH", strSave)
txtNextLevel.Text = GetFromIni("GEN", "NEXTLVL", strSave)
txtMidi.Text = GetFromIni("GEN", "MIDI", strSave)

strTemp = GetFromIni("GEN", "BOSS", strSave)
If strTemp = "1" Then
    chkBoss.Value = 1
Else
    chkBoss.Value = 0
End If
strTemp = GetFromIni("GEN", "HARDBOSS", strSave)
If strTemp = "1" Then
    chkHard.Value = 1
Else
    chkHard.Value = 0
End If

'Load tank options
strTemp = GetFromIni("GEN", "DOUBLEHOVER", strSave)
If strTemp = "1" Then 'Double Hover is enabled
    frmTank.chkDoubleHover.Value = 1
Else
    frmTank.chkDoubleHover.Value = 0
End If
strTemp = GetFromIni("GEN", "BOOST", strSave)
If strTemp = "1" Then
    frmTank.chkBoost.Value = 1
Else
    frmTank.chkBoost.Value = 0
End If
strTemp = GetFromIni("GEN", "CLOAK", strSave)
If strTemp = "1" Then
    frmTank.chkCloak.Value = 1
Else
    frmTank.chkCloak.Value = 0
End If

frmTank.txtBullets.Text = GetFromIni("GEN", "FIREDELAY", strSave)
strTemp = GetFromIni("GEN", "TANKTYPE", strSave)
If strTemp = "" Then strTemp = "0"
frmTank.hTank.Value = CInt(strTemp)

If txtWidth.Text = "" Then txtWidth.Text = "10"
Dim intWidth As Long
intWidth = CLng(txtWidth.Text)
frmEditor.hScroll.Max = intWidth
frmEditor.hScroll.Value = 0
For i = 1 To 100 'Load the enemy information
    strTemp = GetFromIni("GEN", "SV" & i, strSave)
    If strTemp = "1" Then 'Enemy exists, load stats
        SpriteVisible(i) = True
        SpriteType(i) = GetFromIni("GEN", "ST" & i, strSave)
        SpriteX(i) = GetFromIni("GEN", "SX" & i, strSave)
        SpriteY(i) = GetFromIni("GEN", "SY" & i, strSave)
        frmEditor.imgSprite(i).Picture = imgSprite(SpriteType(i)).Picture
        frmEditor.imgSprite(i).Top = SpriteY(i)
        If SpriteX(i) < Me.ScaleWidth Then
            frmEditor.imgSprite(i).Left = SpriteX(i)
            frmEditor.imgSprite(i).Visible = True
        Else
            frmEditor.imgSprite(i).Visible = False
        End If
    Else
        SpriteVisible(i) = False
        imgSprite(i).Visible = False
    End If
Next 'i

End Sub

Private Sub mnuPlay_Click()
'Loads a level to play test
comDiag.ShowOpen
NextLevel = comDiag.FileTitle
frmBriefing.Show

End Sub

Private Sub mnuSave_Click()
'Save the map information to an ini file
On Error Resume Next
comDiag.ShowSave 'Load the save dialog box
Dim strSave As String
strSave = comDiag.FileName 'Set the save file

'Save level options
Call WriteIni("GEN", "NAME", txtName.Text, strSave)
For i = 0 To 2
    Call WriteIni("GEN", "BRIEFING" & i, txtBriefing(i).Text, strSave)
Next 'i
Call WriteIni("GEN", "TYPE", cmbLevelType.Text, strSave)
Call WriteIni("GEN", "BOSSNAME", cmbBoss.Text, strSave)
Call WriteIni("GEN", "WIDTH", txtWidth.Text, strSave)
Call WriteIni("GEN", "NEXTLVL", txtNextLevel.Text, strSave)
Call WriteIni("GEN", "MIDI", txtMidi.Text, strSave)

If chkHard.Value = 0 Then
    Call WriteIni("GEN", "HARDBOSS", "0", strSave)
Else
    Call WriteIni("GEN", "HARDBOSS", "1", strSave)
End If
If chkBoss.Value = 0 Then
    Call WriteIni("GEN", "BOSS", "0", strSave)
Else
    Call WriteIni("GEN", "BOSS", "1", strSave)
End If
'Save tank options
Call WriteIni("GEN", "DOUBLEHOVER", frmTank.chkDoubleHover.Value, strSave)
Call WriteIni("GEN", "BOOST", frmTank.chkBoost.Value, strSave)
Call WriteIni("GEN", "CLOAK", frmTank.chkCloak.Value, strSave)
Call WriteIni("GEN", "FIREDELAY", frmTank.txtBullets.Text, strSave)
Call WriteIni("GEN", "TANKTYPE", frmTank.hTank.Value, strSave)


For i = 1 To 100 'Save enemy data
    If SpriteVisible(i) = True Then
        Call WriteIni("GEN", "SV" & i, "1", strSave)
        Call WriteIni("GEN", "ST" & i, SpriteType(i), strSave)
        Call WriteIni("GEN", "SX" & i, SpriteX(i), strSave)
        Call WriteIni("GEN", "SY" & i, SpriteY(i), strSave)
    Else
        Call WriteIni("GEN", "SV" & i, "0", strSave)
    End If
Next 'i
End Sub

Private Sub mnuTankEditor_Click()
'Load the tank editor screen
frmTank.Show
End Sub
