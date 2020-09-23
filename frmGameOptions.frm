VERSION 5.00
Begin VB.Form frmGameOptions 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Hover Tanx Options"
   ClientHeight    =   2910
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7125
   Icon            =   "frmGameOptions.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2910
   ScaleWidth      =   7125
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   495
      Left            =   4320
      TabIndex        =   5
      Top             =   2280
      Width           =   1215
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Height          =   495
      Left            =   1200
      TabIndex        =   4
      Top             =   2280
      Width           =   1215
   End
   Begin VB.CheckBox chkDemo 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Enable Demo"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   1680
      Value           =   1  'Checked
      Width           =   1815
   End
   Begin VB.CheckBox chkWav 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Sound Effects"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   1320
      Value           =   1  'Checked
      Width           =   1815
   End
   Begin VB.CheckBox chkMidi 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Background Music"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   960
      Value           =   1  'Checked
      Width           =   1815
   End
   Begin VB.Image imgTank 
      Height          =   945
      Left            =   3240
      Picture         =   "frmGameOptions.frx":0D2A
      Top             =   960
      Width           =   2250
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Hover Tanx Options"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   720
      TabIndex        =   3
      Top             =   120
      Width           =   5295
   End
End
Attribute VB_Name = "frmGameOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
Me.Hide
End Sub

Private Sub cmdSave_Click()
'Save the values to settings.dat and update the global vairables associated with them
Dim strSave As String
strSave = App.Path & "\settings.dat"
Call WriteIni("GEN", "MIDI", CStr(chkMidi.Value), strSave)
Call WriteIni("GEN", "WAV", CStr(chkMidi.Value), strSave)
Call WriteIni("GEN", "DEMO", CStr(chkMidi.Value), strSave)
If chkMidi.Value = 1 Then
    ENABLE_MIDI = True
Else
    ENABLE_MIDI = False
End If
If chkWav.Value = 1 Then
    ENABLE_WAV = True
Else
    ENABLE_WAV = False
End If
If chkDemo.Value = 1 Then
    ENABLE_DEMO = True
Else
    ENABLE_DEMO = False
End If
End Sub

Private Sub Form_Load()
'Set the checkboxes to the actual values of the settings
If ENABLE_MIDI = True Then
    chkMidi.Value = 1
Else
    chkMidi.Value = 0
End If
If ENABLE_WAV = True Then
    chkWav.Value = 1
Else
    chkWav.Value = 0
End If
If ENABLE_DEMO = True Then
    chkDemo.Value = 1
Else
    chkDemo.Value = 0
End If
End Sub
