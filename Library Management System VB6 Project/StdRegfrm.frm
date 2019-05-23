VERSION 5.00
Begin VB.Form StdRegfrm 
   Caption         =   "Student Registration"
   ClientHeight    =   6555
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   12645
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6555
   ScaleWidth      =   12645
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      BackColor       =   &H80000018&
      Height          =   375
      Left            =   240
      Picture         =   "StdRegfrm.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   360
      Width           =   615
   End
   Begin VB.ComboBox Streamcmb 
      BackColor       =   &H80000018&
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   450
      ItemData        =   "StdRegfrm.frx":6988A
      Left            =   4440
      List            =   "StdRegfrm.frx":698A6
      TabIndex        =   10
      Text            =   "Choose Stream"
      Top             =   4440
      Width           =   2895
   End
   Begin VB.ComboBox Yearcmb 
      BackColor       =   &H80000018&
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   450
      ItemData        =   "StdRegfrm.frx":698CF
      Left            =   4440
      List            =   "StdRegfrm.frx":698DF
      TabIndex        =   9
      Text            =   "Choose Year"
      Top             =   3600
      Width           =   2895
   End
   Begin VB.CommandButton clrcmd 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Clear"
      BeginProperty Font 
         Name            =   "Modern No. 20"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4200
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   5280
      Width           =   1095
   End
   Begin VB.CommandButton Addcmd 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Add"
      BeginProperty Font 
         Name            =   "Modern No. 20"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5640
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   5280
      Width           =   1695
   End
   Begin VB.TextBox Nametxt 
      BackColor       =   &H80000018&
      BeginProperty Font 
         Name            =   "Goudy Old Style"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   450
      Left            =   4440
      TabIndex        =   6
      Top             =   2880
      Width           =   2895
   End
   Begin VB.TextBox Idtxt 
      BackColor       =   &H80000018&
      BeginProperty Font 
         Name            =   "Goudy Old Style"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   450
      Left            =   4440
      TabIndex        =   5
      Top             =   2160
      Width           =   2895
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00400000&
      Height          =   4095
      Left            =   1920
      Shape           =   4  'Rounded Rectangle
      Top             =   1920
      Width           =   6135
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackColor       =   &H0080C0FF&
      Caption         =   "Stream : "
      BeginProperty Font 
         Name            =   "Goudy Old Style"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   345
      Left            =   2160
      TabIndex        =   4
      Top             =   4440
      Width           =   1035
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H0080C0FF&
      Caption         =   "Year : "
      BeginProperty Font 
         Name            =   "Goudy Old Style"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   345
      Left            =   2160
      TabIndex        =   3
      Top             =   3720
      Width           =   735
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H0080C0FF&
      Caption         =   "Name : "
      BeginProperty Font 
         Name            =   "Goudy Old Style"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   345
      Left            =   2160
      TabIndex        =   2
      Top             =   3000
      Width           =   900
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H0080C0FF&
      Caption         =   "Student ID : "
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
      Left            =   2160
      TabIndex        =   1
      Top             =   2280
      Width           =   1635
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00400000&
      Caption         =   "Student Registration Form"
      BeginProperty Font 
         Name            =   "Perpetua Titling MT"
         Size            =   20.25
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000018&
      Height          =   480
      Left            =   3375
      TabIndex        =   0
      Top             =   480
      Width           =   6480
   End
   Begin VB.Image Image1 
      Height          =   6960
      Left            =   0
      Picture         =   "StdRegfrm.frx":698F7
      Stretch         =   -1  'True
      Top             =   -240
      Width           =   12600
   End
End
Attribute VB_Name = "StdRegfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim db As Database
Dim rs As Recordset


Private Sub Addcmd_Click()
If IDtxt.Text = "" Then
MsgBox ("Enter StudentID")
ElseIf Nametxt.Text = "" Then
MsgBox ("Enter Name")
ElseIf Yearcmb.Text = "Choose Year" Then
MsgBox ("Select year")
ElseIf Streamcmb.Text = "Choose Stream" Then
MsgBox ("Select Stream")
Else
rs.AddNew
rs.Fields(0).Value = IDtxt.Text
rs.Fields(1).Value = Nametxt.Text
rs.Fields(2).Value = Yearcmb.Text
rs.Fields(3).Value = Streamcmb.Text
rs.Update
MsgBox ("Registration is successful")
End If
End Sub

Private Sub clrcmd_Click()
IDtxt.Text = ""
Nametxt.Text = ""
Yearcmb.Text = ""
Streamcmb.Text = ""
IDtxt.SetFocus
End Sub


Private Sub Command1_Click()
Unload Me
Studentdetailsfrm.Show
End Sub

Private Sub Form_Load()
Set db = OpenDatabase("F:\own vb6\StdReg.mdb")
Set rs = db.OpenRecordset("select *from Regfrm")
End Sub

