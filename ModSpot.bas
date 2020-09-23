Attribute VB_Name = "ModSpot"
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Const SRCPAINT = &HEE0086        ' (DWORD) dest = source OR dest
Public Const SRCAND = &H8800C6  ' (DWORD) dest = source AND dest

Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Public Const SND_ASYNC = &H1         '  play asynchronously
Public Const SND_SYNC = &H0         '  play synchronously (default)

Enum eAction
None = 0
Sleep = 1
eat = 2
Play = 3
End Enum

Type Spot
DoLose As Long

Action As eAction
LpAction As Long
Frame As Long

SleepTimer As Long

Sleep As Long
Stomach As Long
Brain As Long


LoseActivity As Long
Activity As Long
TimeHungry As Long
TimeTired As Long

Alive As Boolean
End Type

Public Spot As Spot

Function DoSpot()
Select Case Spot.Action
Case 2
Spot.Stomach = Spot.Stomach + 1: If Spot.Stomach > 50 Then Spot.Stomach = 50

Spot.TimeHungry = 0

Spot.Frame = Spot.Frame + 1
If Spot.Frame > 2 Then
DoSound "eat.wav"
Spot.Frame = 0
Spot.LpAction = Spot.LpAction - 1
If Spot.LpAction <= 0 Then
Spot.LpAction = 0
Spot.Action = 0
End If
End If

Case 1
Spot.Sleep = Spot.Sleep + 1: If Spot.Sleep > 50 Then Spot.Sleep = 50
Spot.TimeTired = 0

Spot.Frame = Spot.Frame + 1
If Spot.Frame > 2 Then
Spot.Frame = 0
Spot.LpAction = Spot.LpAction - 1
If Spot.LpAction <= 0 Then
Spot.LpAction = 0
Spot.Action = 0
End If
End If

Case 0
If Spot.DoLose > 5 Then
Spot.DoLose = 0

If Spot.SleepTimer > Spot.Activity Then
Spot.SleepTimer = 0
Spot.Sleep = Spot.Sleep - 1
If Spot.Sleep <= 0 Then
Spot.Sleep = 0
Spot.TimeTired = Spot.TimeTired + 1
If Spot.TimeTired > 5 Then Spot.Action = eAction.Sleep: Spot.LpAction = 8
End If
Else
Spot.SleepTimer = Spot.SleepTimer + 1
End If

Spot.Stomach = Spot.Stomach - 1: If Spot.Stomach < 0 Then Spot.Stomach = 0
If Spot.Stomach <= 0 Then
Spot.Stomach = 0
Spot.TimeHungry = Spot.TimeHungry + 1
If Spot.TimeHungry > 10 Then Spot.Alive = False
End If

Spot.Brain = Spot.Brain - 1: If Spot.Brain < 0 Then Spot.Brain = 0: If Spot.Brain <= 0 Then Spot.Brain = 0

If Spot.LoseActivity > (Spot.Brain \ 3) Then
Spot.LoseActivity = 0
Spot.Activity = Spot.Activity - 1: If Spot.Activity < 0 Then Spot.Activity = 0: If Spot.Activity <= 0 Then Spot.Activity = 0
Else
Spot.LoseActivity = Spot.LoseActivity + 1
End If

Else
Spot.DoLose = Spot.DoLose + 1
End If

Spot.Frame = Spot.Frame + 1: If Spot.Frame > 2 Then Spot.Frame = 0
Case 3
Spot.Brain = Spot.Brain + 1: If Spot.Brain > 50 Then Spot.Brain = 50

Spot.Frame = Spot.Frame + 1
If Spot.Frame > 2 Then
DoSound "play.wav"
Spot.Frame = 0
Spot.LpAction = Spot.LpAction - 1
Spot.Activity = Spot.Activity + 3: If Spot.Activity > 50 Then Spot.Activity = 50

Spot.Stomach = Spot.Stomach - 3: If Spot.Stomach < 0 Then Spot.Stomach = 0
If Spot.Stomach <= 0 Then
Spot.Stomach = 0
Spot.TimeHungry = Spot.TimeHungry + 1
If Spot.TimeHungry > 10 Then Spot.Alive = False
End If

If Spot.LpAction <= 0 Then
Spot.LpAction = 0
Spot.Action = 0
End If
End If
End Select
End Function

Function DoSound(file As String)
sndPlaySound "C:\WINDOWS\SPOT\" & file, SND_ASYNC
End Function
