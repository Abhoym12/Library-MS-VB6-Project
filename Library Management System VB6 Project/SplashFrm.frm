VERSION 5.00
Begin VB.Form SplashFrm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Library"
   ClientHeight    =   5850
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9120
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "SplashFrm.frx":0000
   ScaleHeight     =   10.319
   ScaleMode       =   0  'User
   ScaleWidth      =   16.087
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "Press Any Key To Continue"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   435
      Left            =   4080
      TabIndex        =   0
      Top             =   5280
      Width           =   4980
   End
End
Attribute VB_Name = "SplashFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_KeyPress(KeyAscii As Integer)
Unload Me
Loginchoicefrm.Show
End Sub
