VERSION 5.00
Begin VB.Form frmHelp 
   BackColor       =   &H000000C0&
   Caption         =   "Help"
   ClientHeight    =   3960
   ClientLeft      =   60
   ClientTop       =   390
   ClientWidth     =   5790
   Icon            =   "frmHelp.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3960
   ScaleWidth      =   5790
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdHelp 
      BackColor       =   &H000000C0&
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3960
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3360
      Width           =   1695
   End
   Begin VB.TextBox txtHelp 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3135
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Text            =   "frmHelp.frx":0442
      Top             =   120
      Width           =   5535
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "By Kevin Fleet -               Copyright(R) 2002 KevCom"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   3360
      Width           =   3255
   End
End
Attribute VB_Name = "frmHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdHelp_Click()
Unload Me
End Sub
