VERSION 5.00
Begin VB.Form LoginFrmAdmin 
   Caption         =   "Admin Login"
   ClientHeight    =   5355
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   8865
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5355
   ScaleMode       =   0  'User
   ScaleWidth      =   14434.04
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton backcmd 
      BackColor       =   &H0080FFFF&
      Height          =   375
      Left            =   240
      Picture         =   "LoginFrmAdmin.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   360
      Width           =   615
   End
   Begin VB.TextBox txtPass 
      BeginProperty Font 
         Name            =   "Goudy Old Style"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   5400
      PasswordChar    =   "*"
      TabIndex        =   6
      Top             =   2640
      Width           =   1935
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5280
      MaskColor       =   &H00404040&
      Picture         =   "LoginFrmAdmin.frx":6988A
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3240
      Width           =   615
   End
   Begin VB.TextBox txtUsername 
      BeginProperty Font 
         Name            =   "Goudy Old Style"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   5400
      TabIndex        =   2
      Top             =   2040
      Width           =   1935
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00400000&
      Caption         =   "Admin Login "
      BeginProperty Font 
         Name            =   "Felix Titling"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000018&
      Height          =   480
      Left            =   2640
      TabIndex        =   5
      Top             =   840
      Width           =   2970
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "Use a valid username and password to gain access to the System."
      BeginProperty Font 
         Name            =   "Modern No. 20"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0FF&
      Height          =   615
      Left            =   3000
      TabIndex        =   4
      Top             =   4320
      Width           =   5280
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H80000018&
      Caption         =   "Password :"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   1
      Left            =   3840
      TabIndex        =   1
      Top             =   2640
      Width           =   1230
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H80000018&
      Caption         =   "Username :"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   0
      Left            =   3840
      TabIndex        =   0
      Top             =   2040
      Width           =   1230
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FF0000&
      Height          =   2535
      Left            =   3360
      Shape           =   4  'Rounded Rectangle
      Top             =   1560
      Width           =   4695
   End
   Begin VB.Image Image1 
      Height          =   5325
      Left            =   0
      Picture         =   "LoginFrmAdmin.frx":917D2
      Stretch         =   -1  'True
      Top             =   0
      Width           =   8820
   End
End
Attribute VB_Name = "LoginFrmAdmin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CboUserAdmin_Change()

End Sub

Private Sub Backcmd_Click()
Unload Me
Loginchoicefrm.Show
End Sub

Private Sub cmdOK_Click()
Dim username As String
Dim password As String
username = "admin"
password = "nsecadmin"
If (username = txtUsername.Text And password = txtPass.Text) Then
MsgBox "Login Successful"
mainfrm.Show
Unload Me
Else
MsgBox "Login Failed. Authentication Error. Try Again."
txtUsername.Text = ""
txtPass.Text = ""
End If
End Sub

Private Sub txtUsername_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
txtPass.SetFocus
End If
End Sub

Private Sub txtPass_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
cmdOK_Click
End If
End Sub

Private Sub Form_Resize()
Image1.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight
End Sub

