VERSION 5.00
Begin VB.Form frmFinish 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Congratulations!"
   ClientHeight    =   3690
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6885
   Icon            =   "frmFinish.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3690
   ScaleWidth      =   6885
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   255
      Left            =   2760
      TabIndex        =   1
      Top             =   3360
      Width           =   1335
   End
   Begin VB.Image imgLogo 
      Height          =   1635
      Left            =   840
      Picture         =   "frmFinish.frx":0D2A
      Top             =   360
      Width           =   4890
   End
   Begin VB.Label lblCongrats 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmFinish.frx":208A
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   240
      TabIndex        =   0
      Top             =   2160
      Width           =   6375
   End
End
Attribute VB_Name = "frmFinish"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub Form_Load()
curMidi = "finish.mid"
PlayMidi ("finish.mid")
End Sub

Private Sub Form_Unload(Cancel As Integer)
StopMidi (curMidi)
Me.Hide
frmIntro.Show
End Sub
