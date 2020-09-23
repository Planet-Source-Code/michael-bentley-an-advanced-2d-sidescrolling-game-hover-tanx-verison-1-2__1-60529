Attribute VB_Name = "modHoverTanx"
' Most code, except for INI and some Midi Functions
' ©2005 Michael Bentley (ikillkenny@comcast.net)
' Feel free to use code from this project, but it would nice
' if you gave me a little acknowledgement in the credits
' E-mail me any questions, and if you like the code please vote
' for me at www.planetsourcecode.com


Public ENABLE_DEMO As Boolean 'Whether the demo has been enabled
Public ENABLE_MIDI As Boolean 'Whether or not to play music files
Public ENABLE_WAV As Boolean 'Whehter or not to play sounds
Public firstCinematic As Boolean 'If you haven't seen the first cinematic yet

'Variables for map editor
Public SpriteType(1 To 100) As String 'what picture the sprite is
Public SpriteX(1 To 100) As String 'X location of sprite
Public SpriteY(1 To 100) As String 'Y location of sprite
Public SpriteVisible(1 To 100) As Boolean 'If the sprite has been placed on the map yet

'Various color constants
Public Const COLDESERT As Long = &HC0E0FF
Public Const COLSNOW As Long = &HE0E0E0
Public Const COLSOIL As Long = &H4080&
Public Const COLCLEARSKY As Long = &HFFFF80
Public Const COLGREEN As Long = &HFF00&
Public Const COLRED As Long = &HFF&
Public Const COLCITY As Long = &H808080
Public Const COLMARS As Long = &H80&
Public Const COLVOLCANO As Long = &H4080&
Public Const COLWHITE As Long = &HFFFFFF
Public Const COLWATER As Long = &HC00000
Public Const COLWATERB As Long = &H80C0FF
Public Const COLBLACK As Long = &H0&
Public Const COLGRAY As Long = &H404040
Public Const COLLGRAY As Long = &HE0E0E0
Public Const COLDGRAY As Long = &H808080
Public Const COLINTROIN As Long = &H404040


Public LevelLoaded As Boolean 'If the level has been loaded from the ini files yet
Public Demo As Boolean 'If you're in demo mode or not

Public curMidi As String 'Current midi being played

Public NextLevel As String 'The next ini file to load
Public curCampaign As String 'The current campaign file you're in
Public anType As Long

'Declare the BitBlt function, used for drawing graphics
Declare Function BitBlt Lib "gdi32" ( _
        ByVal hDestDC As Long, _
        ByVal X As Long, _
        ByVal Y As Long, _
        ByVal nWidth As Long, _
        ByVal nHeight As Long, _
        ByVal hSrcDC As Long, _
        ByVal xSrc As Long, _
        ByVal ySrc As Long, _
        ByVal dwRop As Long) As Long

'Declare the Sprite type, encompassing tanks, enemies, some other things
Public Type Sprite
    X As Long 'Left position on form
    Y As Long 'Top position on form
    Width As Long 'Width of picture
    Height As Long 'Height of picture
    Visible As Boolean 'Whehter or not to draw the sprite
    Picture As Long 'Corresponds to index of picSprite
    Speed As Long 'How fast sprite moves in X direction
    SpeedY As Long 'How fast sprite moves in Y direction
    Damage As Long 'How much damage sprite does
    Draw As Boolean 'Sprite exists, but may not be drawn in draw=false
    Mode As Long 'What operation the sprite is in
    SecondPic As Long 'Denotes a second picture frame, such as with the missile
    HP As Long 'How much health the sprite has
    Invincible As Boolean 'If the sprite cannot be killed
    Blink As Long 'Grace window before being hit again
End Type

'Object representing text drawn to the screen
Public Type typeText
    Text As String 'Actual contents of text
    X As Long 'Left position
    Y As Long 'Top position
    Speed As Long 'How fast it moves in X direction
    Color As Long 'color of text
    Font As String 'Font of text
    Size As Long 'Size of text
    Italics As Boolean 'Italics of not
    Bold As Boolean 'Bold or not
    Visible As Boolean 'If the text exists or not
End Type

Private curSound As Long 'Index of last sound to be played (from 1-10)

'Declarations for printing text
Public Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal lpString As String, ByVal nCount As Long) As Long
'Declarations for gameloop
Public Declare Function GetTickCount Lib "kernel32" () As Long

'High resolution timing
Public Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Long
Public Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long

'Wave Functions
Private Declare Function sndPlaySound Lib "WINMM.DLL" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long

'Midi functions
Declare Function mciSendString Lib "WINMM.DLL" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As Any, ByVal uReturnLength As Long, ByVal hwndcallback As Long) As Long
Declare Function mciGetErrorString Lib "WINMM.DLL" Alias "mciGetErrorStringA" (ByVal dwError As Long, ByVal lpstrBuffer As String, ByVal uLength As Long) As Long
Declare Function mciSetRepeat Lib "WINMM.DLL" Alias "mciSetRepeatA" ()
Private Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal lBuffer As Long) As Long

'Ini Functions
Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
'Get values back from an INI file
Function GetFromIni(strSectionHeader As String, strVariableName As String, strFileName As String) As String
    Dim strReturn As String
    strReturn = String(255, Chr(0))
    GetFromIni = Left$(strReturn, GetPrivateProfileString(strSectionHeader, ByVal strVariableName, "", strReturn, Len(strReturn), strFileName))
End Function
'Write values to an INI file
Function WriteIni(strSectionHeader As String, strVariableName As String, strValue As String, strFileName As String) As Integer
    WriteIni = WritePrivateProfileString(strSectionHeader, strVariableName, strValue, strFileName)
End Function

Sub PlySound(strSound As String)
'Plays a wave file
If ENABLE_WAV = False Then Exit Sub 'Sounds not enabled, so just exit

'Get the short (DOS friendly) location of the wave file, since mci needs it
strSound = GetShortPath(App.Path) & "\" & strSound & ".wav"
If strSound = "" Then Exit Sub 'No path found, exit

Call mciSendString("close wav" & CStr(curSound), 0&, 0, 0) 'Gets rid of the previous file existing on this alias
Call mciSendString("open " & strSound & " Alias wav" & CStr(curSound), 0&, 0, 0) 'Open the new wav file
Call mciSendString("play wav" & CStr(curSound), 0&, 0, 0) 'Play the wave file

curSound = curSound + 1 'This is the number of the sound to play
If curSound > 10 Then 'Loop back to the 0th sound
    curSound = 0
End If

End Sub
Sub PlayMidi(strMidi As String)
'Plays background music files
If ENABLE_MIDI = False Then Exit Sub 'Music not enabled, exit
strMidi = GetShortPath(App.Path) & "\" & strMidi 'Get the short location of the midi file
If strMidi = "" Then Exit Sub 'No midi exists, exit
Call mciSendString("close all", 0&, 0, 0) 'Stop all previous music
Call mciSendString("open " & strMidi$ & " Alias midi", 0&, 0, 0) 'Open the new file
Call mciSendString("play midi", 0&, 0, 0) 'Play the new file

End Sub
Sub StopMidi(strMidi As String)
On Error Resume Next
'Below no longer used
'strMidi = GetShortPath(App.Path) & "\" & strMidi
'If strMidi = "" Then Exit Sub
'Close and stop the current midi being played
Call mciSendString("stop midi", 0&, 0, 0)
Call mciSendString("close all", 0&, 0, 0)
End Sub
Public Function GetShortPath(strFileName As String) As String
    Dim lngRes As Long, strPath As String
    'Create a buffer
    strPath = String$(165, 0)
    'retrieve the short pathname
    lngRes = GetShortPathName(strFileName, strPath, 164)
    'remove all unnecessary chr$(0)'s
    GetShortPath = Left$(strPath, lngRes)
End Function
Public Sub PauseMidi(strMidi As String)
On Error Resume Next
If ENABLE_MIDI = False Then Exit Sub 'Music not enabled, exit
Dim intReturn As Long
strMidi = GetShortPath(App.Path) & "\" & strMidi 'Get the short path of the midi
If strMidi = "" Then Exit Sub 'No midi exists
intReturn = mciSendString("Pause midi", 0&, 0, 0) 'Send the pause command
If intReturn <> 0 Then 'Error pausing
    Debug.Print "pause error"
End If
End Sub
Public Sub ResumeMidi(strMidi As String)
'Resumes a midi from the position when it was paused
On Error Resume Next
If ENABLE_MIDI = False Then Exit Sub 'Music not enabled, exit
Dim dwReturn As Long
Dim pos As String * 128
strMidi = GetShortPath(App.Path) & "\" & strMidi 'Get short path of midi
If strMidi = "" Then Exit Sub 'No midi exists

dwReturn = mciSendString("status midi position", pos, 128, 0&) 'Get the midi's position
Call mciSendString("play midi from " & pos, 0&, 0&, 0&) 'Play the midi from the position you got

End Sub
Public Sub RepeatMidi()
'Repeats a midi file
On Error Resume Next
If ENABLE_MIDI = False Then Exit Sub 'Music not enabled, exit
Dim dwReturn As Long
Dim Total As String * 128

dwReturn = mciSendString("set midi time format frames", Total, 128, 0&) 'Set the midi to the proper format
dwReturn = mciSendString("status midi length", Total, 128, 0&) 'Get the total length of the midi

Dim pos As String * 128

dwReturn = mciSendString("status midi position", pos, 128, 0&) 'Get the current position of the midi

If pos = Total Then 'If the position is that the end, repeat
    dwReturn = mciSendString("seek midi to 0", 0&, 0&, 0&) 'Seek the midi to 0 position (the start)
    mciSendString "Play midi", 0&, 0&, 0& 'Play the midi again
End If

End Sub
