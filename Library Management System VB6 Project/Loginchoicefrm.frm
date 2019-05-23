VERSION 5.00
Begin VB.Form Loginchoicefrm 
   Caption         =   "Login Choice"
   ClientHeight    =   6300
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   11835
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6300
   ScaleWidth      =   11835
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      BackColor       =   &H80000018&
      Caption         =   "Admin Login"
      BeginProperty Font 
         Name            =   "Perpetua Titling MT"
         Size            =   15.75
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   6840
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2400
      Width           =   3495
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H80000018&
      Caption         =   "Student Login"
      BeginProperty Font 
         Name            =   "Perpetua Titling MT"
         Size            =   15.75
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2400
      Width           =   3735
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H0080C0FF&
      Caption         =   "Choose Your Preference"
      BeginProperty Font 
         Name            =   "Felix Titling"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   540
      Left            =   3480
      TabIndex        =   2
      Top             =   840
      Width           =   5685
   End
   Begin VB.Image Image1 
      Height          =   6255
      Left            =   -240
      Picture         =   "Loginchoicefrm.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   12120
   End
End
Attribute VB_Name = "Loginchoicefrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
Unload Me
LoginfrmStd.Show
End Sub

Private Sub Command2_Click()
Unload Me
LoginFrmAdmin.Show
End Sub

