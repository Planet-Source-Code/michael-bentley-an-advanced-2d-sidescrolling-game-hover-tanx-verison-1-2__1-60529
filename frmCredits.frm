VERSION 5.00
Begin VB.Form frmCredits 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Hover Tanx Credits"
   ClientHeight    =   4875
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9270
   Icon            =   "frmCredits.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   325
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   618
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "frmCredits"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim intCredit As Long 'Current name that you're on
Dim bCreateCredit As Boolean
Dim Credits(1 To 10) As typeText 'Actual text
Dim TickDif As Double 'how fast to run the credits loop
Dim LastTick As Double 'what the last CPU tick was
Dim bRunning As Boolean 'If the credits loop is running

Private Sub cmdDone_Click()
frmCredits.Hide ' Get rid of this form
frmIntro.Show 'Load the intro
End Sub

Private Sub Form_Activate()
'Start the gameloop, and cal lthe function
bRunning = True
Call GameLoop
End Sub

Private Sub Form_Load()
On Error Resume Next
bCreateCredit = True
TickDif = 0.035 'The loop runs ever 35 ms
StopMidi (curMidi) 'Stop the current midi
curMidi = "credits.mid" 'Set the credits song
PlayMidi (curMidi) 'Play the credits song
For i = 1 To 10 'Set the credit text to invisible and set attributes
    Credits(i).Visible = False
    Credits(i).Size = 28
    Credits(i).Font = "Arial Black"
Next 'i
If curCampaign <> "" Then 'You're in campaign mode, set the next available level to level13 (mars)
    Dim strCampaign As String
    strCampaign = GetFromIni("GEN", "LEVEL", App.Path & "\" & curCampaign)
    If strCampaign <> "level13.ini" And strCampaign <> "level14.ini" And strCampaign <> "level15.ini" Then
        Call WriteIni("GEN", "LEVEL", "level13.ini", App.Path & "\" & curCampaign)
    End If
End If
End Sub
Private Sub GameLoop()
'Performs an action every tickDif interval
Dim curFreq As Currency 'How often this computer runs a loop
Dim curStart As Currency 'Start time
Dim curEnd As Currency 'End time
Dim dblResult As Double 'Difference between lastTick and this tick

QueryPerformanceFrequency curFreq       'Get the timer frequency
Do While bRunning = True
    QueryPerformanceCounter curStart        'Get the start time
    dblResult = (curStart - LastTick) / curFreq   'Calculate the duration (in seconds)
    If dblResult > TickDif Then 'It's time for the loop to run
        LastTick = curStart 'Set the last tick to the current tick
        Me.Cls 'Clear the screen
        GetCredits 'Move the credits
        DrawText 'Draw the credit text
        Me.Refresh 'Refresh
    End If
    DoEvents
Loop

End Sub
Private Sub DrawText()
'Draws the credit text
Line (0, 0)-Step(Me.ScaleWidth, 50), COLBLACK, BF
Line (0, Me.ScaleHeight - 50)-Step(Me.ScaleWidth, 50), COLBLACK, BF

For i = 1 To 10 'Loop through each credit
    If Credits(i).Visible = True Then
        Credits(i).X = Credits(i).X + Credits(i).Speed 'Move the credits according ot hte speed
        If Credits(i).Speed < 0 And Credits(i).X <= -400 Then 'The credit is now off the screen
            Credits(i).Visible = False 'Set to invisible
        ElseIf Credits(i).Speed > 0 And Credits(i).X >= Me.ScaleWidth Then 'the credit is off the screen the other way
            Credits(i).Visible = False 'Set to invisible
        End If
        Me.ForeColor = Credits(i).Color
        Me.FontBold = Credits(i).Bold
        Me.FontItalic = Credits(i).Italics
        Me.Font = Credits(i).Font
        Me.FontSize = Credits(i).Size
        Call TextOut(Me.hdc, Credits(i).X, Credits(i).Y, Credits(i).Text, Len(Credits(i).Text))
    End If
Next 'i
End Sub
Private Sub GetCredits()
'Updates the text in each credit
If bCreateCredit = True Then 'If it's time to make a new credit
    bCreateCredit = False
    Select Case intCredit
    Case 0
        Call MakeCredit("Hover Tanx", 50, 6, COLGRAY, False)
        Call MakeCredit("Hover Tanx", 125, -7, COLDGRAY, False)
        Call MakeCredit("Hover Tanx", 200, 8, COLLGRAY, False)
    Case 1
        Call MakeCredit("Designed By", 50, 8, COLDGRAY, False)
        Call MakeCredit("Designed By", 125, -7, COLLGRAY, False)
        Call MakeCredit("Designed By", 200, 6, COLGRAY, False)
    Case 2
        Call MakeCredit("Michael Bentley", 125, -7, COLDGRAY, True)
    Case 3
        Call MakeCredit("Programmed By", 50, 8, COLDGRAY, False)
        Call MakeCredit("Programmed By", 125, -7, COLLGRAY, False)
        Call MakeCredit("Programmed By", 200, 6, COLGRAY, False)
    Case 4
        Call MakeCredit("Michael Bentley", 125, -7, COLDGRAY, False)
    Case 5
        Call MakeCredit("Graphics By", 50, 8, COLDGRAY, False)
        Call MakeCredit("Graphics By", 125, -7, COLLGRAY, False)
        Call MakeCredit("Graphics By", 200, 6, COLGRAY, False)
    Case 6
        Call MakeCredit("Paul Thurlow", 75, -7, COLDGRAY, True)
        Call MakeCredit("Michael Bentley", 150, 6, COLGRAY, True)
    Case 7
        Call MakeCredit("Music From", 50, 8, COLDGRAY, False)
        Call MakeCredit("Music From", 125, -7, COLLGRAY, False)
        Call MakeCredit("Music From", 200, 6, COLGRAY, False)
    Case 8
        Call MakeCredit("The Good People At", 75, -7, COLDGRAY, True)
        Call MakeCredit("www.vgmusic.com", 150, 6, COLGRAY, True)
    Case 9
        Call MakeCredit("Sound Effects From", 50, 8, COLDGRAY, False)
        Call MakeCredit("Sound Effects From", 125, -7, COLLGRAY, False)
        Call MakeCredit("Sound Effects From", 200, 6, COLGRAY, False)
    Case 10
        Call MakeCredit("www.a1freesoundeffects.com", 125, 6, COLGRAY, True)
    Case 11
        Call MakeCredit("Developed By", 50, 8, COLDGRAY, False)
        Call MakeCredit("Developed By", 125, -7, COLLGRAY, False)
        Call MakeCredit("Developed By", 200, 6, COLGRAY, False)
    Case 12
        Call MakeCredit("Doc Entertainment", 125, -7, COLDGRAY, True)
    Case 13
        Call MakeCredit("Published By", 50, 8, COLDGRAY, False)
        Call MakeCredit("Published By", 125, -7, COLLGRAY, False)
        Call MakeCredit("Published By", 200, 6, COLGRAY, False)
    Case 14
        Call MakeCredit("Doc Inc.", 125, -7, COLDGRAY, True)
    Case 15
        Call MakeCredit("Thanks For Playing!", 50, 8, COLDGRAY, False)
        Call MakeCredit("Thanks For Playing!", 125, -7, COLLGRAY, False)
        Call MakeCredit("Thanks For Playing!", 200, 6, COLGRAY, False)
    Case 16
        Call MakeCredit("But Wait....", 125, -7, COLDGRAY, True)
    Case 17
        Call MakeCredit("Check Your Campaign Screen!", 50, 8, COLDGRAY, False)
        Call MakeCredit("Check Your Campaign Screen!", 125, -7, COLLGRAY, False)
        Call MakeCredit("Check Your Campaign Screen!", 200, 6, COLGRAY, False)
    Case 18
        Unload Me
        frmIntro.Show
    End Select
Else 'Not time to update the credits
    For i = 1 To 10
        If Credits(i).Visible = True Then 'If a credit is still on the screen, do nothing
            Exit Sub
        End If
    Next 'i
    intCredit = intCredit + 1 'Increment the credits count
    bCreateCredit = True 'Update the credits next time
End If
End Sub
Private Sub MakeCredit(Text As String, Y As Long, Speed As Long, Color As Long, Italics As Boolean)
'Creates a new credit from an empty one
For i = 1 To 10
    If Credits(i).Visible = False Then
        Credits(i).Text = Text
        Credits(i).Speed = Speed
        Credits(i).Y = Y
        If Speed < 0 Then
            Credits(i).X = Me.ScaleWidth + 25
        Else
            Credits(i).X = -(15 * Len(Text))
        End If
        Credits(i).Color = Color
        Credits(i).Italics = Italics
        Credits(i).Visible = True
        Exit Sub
    End If
Next 'i
End Sub

Private Sub Form_Unload(Cancel As Integer)
bRunning = False 'Stop the credit loop from running
StopMidi (curMidi) 'Stop the credits music
PlayMidi ("intro.mid") 'Start the intro music
End Sub
