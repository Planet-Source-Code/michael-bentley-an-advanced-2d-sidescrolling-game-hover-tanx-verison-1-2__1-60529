VERSION 5.00
Begin VB.Form frmTank 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Tank Properties"
   ClientHeight    =   3345
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   8415
   ForeColor       =   &H00C00000&
   Icon            =   "frmTank.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   223
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   561
   StartUpPosition =   3  'Windows Default
   Begin VB.HScrollBar hTank 
      Height          =   255
      Left            =   2760
      Max             =   2
      TabIndex        =   5
      Top             =   1560
      Width           =   1935
   End
   Begin VB.TextBox txtBullets 
      Height          =   285
      Left            =   240
      MaxLength       =   3
      TabIndex        =   4
      Text            =   "20"
      Top             =   1800
      Width           =   375
   End
   Begin VB.CheckBox chkCloak 
      Caption         =   "Cloak"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   1200
      Width           =   855
   End
   Begin VB.CheckBox chkDoubleHover 
      Caption         =   "Double Hover"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   1575
   End
   Begin VB.CheckBox chkBoost 
      Caption         =   "Boost"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   855
   End
   Begin VB.Image imgTank 
      Height          =   1410
      Index           =   7
      Left            =   480
      Picture         =   "frmTank.frx":0D2A
      Top             =   1920
      Visible         =   0   'False
      Width           =   1905
   End
   Begin VB.Image imgTank 
      Height          =   870
      Index           =   6
      Left            =   4080
      Picture         =   "frmTank.frx":14AE
      Top             =   2400
      Visible         =   0   'False
      Width           =   2640
   End
   Begin VB.Image imgTank 
      Height          =   540
      Index           =   5
      Left            =   5040
      Picture         =   "frmTank.frx":1B15
      Top             =   1320
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Image imgTank 
      Height          =   1140
      Index           =   4
      Left            =   5400
      Picture         =   "frmTank.frx":1FBB
      Top             =   120
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Image imgTank 
      Height          =   1395
      Index           =   3
      Left            =   6720
      Picture         =   "frmTank.frx":278C
      Top             =   360
      Visible         =   0   'False
      Width           =   2970
   End
   Begin VB.Image imgTank 
      Height          =   1125
      Index           =   0
      Left            =   5160
      Picture         =   "frmTank.frx":382A
      Top             =   360
      Visible         =   0   'False
      Width           =   2205
   End
   Begin VB.Image imgTank 
      Height          =   945
      Index           =   1
      Left            =   2880
      Picture         =   "frmTank.frx":3F49
      Top             =   2160
      Visible         =   0   'False
      Width           =   2250
   End
   Begin VB.Image imgTank 
      Height          =   1095
      Index           =   2
      Left            =   5280
      Picture         =   "frmTank.frx":4728
      Top             =   1680
      Visible         =   0   'False
      Width           =   2250
   End
   Begin VB.Label lblGen 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tank:"
      Height          =   195
      Index           =   7
      Left            =   3120
      TabIndex        =   6
      Top             =   120
      Width           =   420
   End
   Begin VB.Image imgTankSelect 
      Height          =   1125
      Left            =   2520
      Picture         =   "frmTank.frx":4AD3
      Top             =   360
      Width           =   2205
   End
   Begin VB.Label lblGen 
      AutoSize        =   -1  'True
      Caption         =   "Cool Down Time In Frames:"
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   3
      Top             =   1560
      Width           =   1950
   End
End
Attribute VB_Name = "frmTank"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
'Set the maximum number of tanks to the scroll through to the number of actual tanks ther are
hTank.Max = imgTank.UBound
End Sub

Private Sub hTank_Change()
'Update the picture of the "current tank selected"
On Error Resume Next
imgTankSelect.Picture = imgTank(hTank.Value).Picture
End Sub
