VERSION 5.00
Begin VB.Form frmSpot 
   BorderStyle     =   0  'None
   Caption         =   "Spot"
   ClientHeight    =   2265
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3540
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2265
   ScaleWidth      =   3540
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrSpot 
      Interval        =   300
      Left            =   2760
      Top             =   3240
   End
   Begin VB.PictureBox board 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00800000&
      BorderStyle     =   0  'None
      Height          =   2295
      Left            =   0
      ScaleHeight     =   153
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   241
      TabIndex        =   1
      Top             =   0
      Width           =   3615
      Begin VB.CommandButton cmdReset 
         BackColor       =   &H00C00000&
         Caption         =   "Reset Spot"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2040
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   1560
         Width           =   1455
      End
      Begin VB.CommandButton cmdHelp 
         BackColor       =   &H00C00000&
         Caption         =   "Help"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2040
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   1200
         Width           =   1455
      End
      Begin VB.CommandButton cmdPlay 
         BackColor       =   &H00C00000&
         Caption         =   "Play"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2040
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   840
         Width           =   1455
      End
      Begin VB.CommandButton cmdExit 
         BackColor       =   &H00C00000&
         Caption         =   "Exit"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2040
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   1920
         Width           =   1455
      End
      Begin VB.CommandButton cmdSleep 
         BackColor       =   &H00C00000&
         Caption         =   "Sleep"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2040
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   480
         Width           =   1455
      End
      Begin VB.CommandButton cmdEat 
         BackColor       =   &H00C00000&
         Caption         =   "Eat"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2040
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   120
         Width           =   1455
      End
      Begin VB.PictureBox pic 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FF0000&
         BorderStyle     =   0  'None
         Height          =   435
         Left            =   1200
         ScaleHeight     =   29
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   30
         TabIndex        =   2
         Top             =   120
         Width           =   450
      End
   End
   Begin VB.Timer tmrRun 
      Interval        =   300
      Left            =   2760
      Top             =   2760
   End
   Begin VB.PictureBox spotsrc 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1815
      Left            =   4200
      Picture         =   "frmSpot.frx":0000
      ScaleHeight     =   121
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   90
      TabIndex        =   0
      Top             =   2280
      Visible         =   0   'False
      Width           =   1350
   End
End
Attribute VB_Name = "frmSpot"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdEat_Click()
If Spot.Action <> 0 Then Exit Sub
Spot.Action = 2
Spot.LpAction = 3
End Sub

Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub cmdHelp_Click()
Load frmHelp
frmHelp.Visible = True
End Sub

Private Sub cmdPlay_Click()
If Spot.Action <> 0 Then Exit Sub
Spot.Action = 3
Spot.LpAction = 3
End Sub

Private Sub cmdSleep_Click()
Select Case cmdSleep.Caption
Case "Sleep"
If Spot.Action <> 0 Then Exit Sub
Spot.Action = 1
Spot.LpAction = 999999999
cmdSleep.Caption = "Wake-up"
Case "Wake-up"
Spot.Action = 0
Spot.LpAction = 0
cmdSleep.Caption = "Sleep"
End Select
End Sub

Function FileExist(path As String) As Boolean
On Error GoTo oops
FileExist = True
Open path For Input As #1
Close #1
Exit Function
oops:
FileExist = False
End Function

Function ResetSpot()
Spot.Action = 0
Spot.Activity = 10
Spot.Alive = True
Spot.Brain = 50
Spot.DoLose = 0
Spot.Frame = 0
Spot.LoseActivity = 0
Spot.LpAction = 0
Spot.Sleep = 50
Spot.SleepTimer = 0
Spot.Stomach = 50
Spot.TimeHungry = 0
Spot.TimeTired = 0
End Function

Function LoadSound(file As String)
Dim B() As Byte
B() = LoadResData(file, "CUSTOM")
Open "C:\WINDOWS\SPOT\" & file For Binary As #1
Put #1, , B()
Close #1
End Function

Private Sub Form_Load()
If UCase(Dir("C:\WINDOWS\SPOT", vbDirectory)) <> "SPOT" Then MkDir ("C:\WINDOWS\SPOT\")

If FileExist("C:\WINDOWS\SPOT\play.wav") = False Then LoadSound "play.wav"
If FileExist("C:\WINDOWS\SPOT\eat.wav") = False Then LoadSound "eat.wav"
If FileExist("C:\WINDOWS\SPOT\spot.dat") = False Then
ResetSpot
Open "C:\WINDOWS\SPOT\spot.dat" For Output As #1
Write #1, Spot.Action, Spot.Activity, Spot.Alive, Spot.Brain, Spot.DoLose, Spot.Frame, Spot.LoseActivity, Spot.LpAction, Spot.LpAction, Spot.Sleep, Spot.SleepTimer, Spot.Stomach, Spot.TimeHungry
Close #1
End If

Open "C:\WINDOWS\SPOT\spot.dat" For Input As #1
Input #1, Spot.Action, Spot.Activity, Spot.Alive, Spot.Brain, Spot.DoLose, Spot.Frame, Spot.LoseActivity, Spot.LpAction, Spot.LpAction, Spot.Sleep, Spot.SleepTimer, Spot.Stomach, Spot.TimeHungry
Close #1

tmrRun_Timer
End Sub

Private Sub Form_Unload(Cancel As Integer)
Open "C:\WINDOWS\SPOT\spot.dat" For Output As #1
Write #1, Spot.Action, Spot.Activity, Spot.Alive, Spot.Brain, Spot.DoLose, Spot.Frame, Spot.LoseActivity, Spot.LpAction, Spot.LpAction, Spot.Sleep, Spot.SleepTimer, Spot.Stomach, Spot.TimeHungry
Close #1
End Sub

Private Sub tmrRun_Timer()
board.Cls
board.ForeColor = vbBlack
If Spot.Alive = True Then board.ForeColor = vbBlue
board.FontBold = True
board.FontSize = 18
board.CurrentX = 10
board.CurrentY = 10
board.Print "Spot"

DrawBar "Sleep:", Spot.Sleep, 0
DrawBar "Stomach:", Spot.Stomach, 1
DrawBar "Brain:", Spot.Brain, 2
DrawBar "Activity:", Spot.Activity, 3

pic.Cls
Select Case Spot.Alive
Case True
BitBlt pic.hDC, 0, 0, 30, 30, spotsrc.hDC, Spot.Frame * 30, Spot.Action * 30, SRCPAINT
BitBlt pic.hDC, 0, 0, 30, 30, spotsrc.hDC, Spot.Frame * 30, Spot.Action * 30, SRCAND
Case False
pic.DrawWidth = 10
pic.Line (0, 0)-(pic.ScaleWidth, pic.ScaleHeight)
pic.Line (pic.ScaleWidth, 0)-(0, pic.ScaleHeight)
pic.DrawWidth = 1
End Select
End Sub

Function GetY(spt As Long) As Long
Dim bigheight
board.FontSize = 18
bigheight = board.TextHeight("|")
board.FontSize = 8

GetY = 10 + bigheight + 10 + ((5 + board.TextHeight("|")) * spt)
End Function

Function DrawBar(Text As String, Value As Long, spt As Long)
board.CurrentX = 10
board.CurrentY = GetY(spt)
board.Print Text
board.Line (10 + board.TextWidth(Text) + 5, GetY(spt))-(10 + board.TextWidth(Text) + 5 + 50, GetY(spt) + board.TextHeight("|")), vbRed, BF
board.Line (10 + board.TextWidth(Text) + 5, GetY(spt))-(10 + board.TextWidth(Text) + 5 + Value, GetY(spt) + board.TextHeight("|")), vbGreen, BF
End Function

Private Sub tmrSpot_Timer()
If Spot.Alive = True Then DoSpot
End Sub
