VERSION 5.00
Begin VB.Form LoginfrmStd 
   Caption         =   "Student Login"
   ClientHeight    =   5775
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   10380
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5775
   ScaleWidth      =   10380
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton backcmd 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      Height          =   495
      Left            =   240
      Picture         =   "LoginStdfrm.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   240
      Width           =   615
   End
   Begin VB.CommandButton Clrcmd 
      BackColor       =   &H0080FFFF&
      Caption         =   "Clear"
      BeginProperty Font 
         Name            =   "Perpetua Titling MT"
         Size            =   15.75
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3360
      Width           =   1455
   End
   Begin VB.CommandButton Logincmd 
      BackColor       =   &H0080FFFF&
      Caption         =   "Login"
      BeginProperty Font 
         Name            =   "Perpetua Titling MT"
         Size            =   15.75
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6120
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3360
      Width           =   1455
   End
   Begin VB.TextBox IDtxt 
      BeginProperty Font 
         Name            =   "Goudy Old Style"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4920
      TabIndex        =   2
      Top             =   2160
      Width           =   2535
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H0080FFFF&
      Height          =   495
      Left            =   240
      Shape           =   4  'Rounded Rectangle
      Top             =   240
      Width           =   615
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H0080C0FF&
      Caption         =   "Student ID: "
      BeginProperty Font 
         Name            =   "Goudy Old Style"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   375
      Left            =   2880
      TabIndex        =   1
      Top             =   2160
      Width           =   1590
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00C0FFFF&
      Height          =   3135
      Left            =   2520
      Shape           =   4  'Rounded Rectangle
      Top             =   1560
      Width           =   5295
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00400000&
      Caption         =   "Student Login"
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
      Left            =   3510
      TabIndex        =   0
      Top             =   360
      Width           =   3555
   End
   Begin VB.Image Image1 
      Height          =   5880
      Left            =   0
      Picture         =   "LoginStdfrm.frx":6988A
      Stretch         =   -1  'True
      Top             =   0
      Width           =   10440
   End
End
Attribute VB_Name = "LoginfrmStd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim str As String
Dim db As Database
Dim rs As Recordset
Public tempUname As String

Private Sub Backcmd_Click()
Loginchoicefrm.Show
Unload Me

End Sub

Private Sub clrcmd_Click()
IDtxt.Text = ""
IDtxt.SetFocus
End Sub

Private Sub Logincmd_Click()
Set db = OpenDatabase("F:\own vb6\StdReg.mdb")
str = "select * from Regfrm where StudentID='" & IDtxt.Text & "'"
Set rs = db.OpenRecordset(str)
If rs.EOF Then
MsgBox ("Record Not Found")
IDtxt.Text = ""
IDtxt.SetFocus
Else
If rs.Fields(0) = IDtxt.Text Then
tempUname = IDtxt.Text
Studentviewfrm.Text1.Text = tempUname
Unload Me
Studentviewfrm.Show
End If
End If
End Sub
