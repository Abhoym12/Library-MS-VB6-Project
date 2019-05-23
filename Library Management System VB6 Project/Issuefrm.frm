VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Issuefrm 
   Caption         =   "Issue Form"
   ClientHeight    =   7485
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   12150
   LinkTopic       =   "Form1"
   ScaleHeight     =   7485
   ScaleWidth      =   12150
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command5 
      BackColor       =   &H0080FFFF&
      Height          =   495
      Left            =   120
      Picture         =   "Issuefrm.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   240
      Width           =   735
   End
   Begin VB.CommandButton clrcmd 
      BackColor       =   &H0080FFFF&
      Caption         =   "Clear"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4320
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   6360
      Width           =   1095
   End
   Begin VB.CommandButton issuecmd 
      BackColor       =   &H0080FFFF&
      Caption         =   "Issue"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5760
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   6360
      Width           =   1095
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   345
      Left            =   3600
      TabIndex        =   16
      Top             =   5640
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   609
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Goudy Old Style"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   114098177
      CurrentDate     =   43585
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   345
      Left            =   3600
      TabIndex        =   14
      Top             =   4920
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   609
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Goudy Old Style"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CalendarForeColor=   4194304
      Format          =   114098177
      CurrentDate     =   43585
   End
   Begin VB.TextBox StdNametxt 
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
      Height          =   345
      Left            =   3600
      TabIndex        =   12
      Top             =   4200
      Width           =   2415
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H0080FFFF&
      Caption         =   "Search"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   6360
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   3480
      Width           =   1215
   End
   Begin VB.TextBox IDtxt 
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
      Height          =   345
      Left            =   3600
      TabIndex        =   9
      Top             =   3480
      Width           =   2415
   End
   Begin VB.TextBox authortxt 
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
      Height          =   345
      Left            =   3600
      TabIndex        =   7
      Top             =   2760
      Width           =   2415
   End
   Begin VB.TextBox bnametxt 
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
      Height          =   345
      Left            =   3600
      TabIndex        =   5
      Top             =   2040
      Width           =   2415
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080FFFF&
      Caption         =   "Search"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6360
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1320
      Width           =   1215
   End
   Begin VB.TextBox accntxt 
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
      Height          =   345
      Left            =   3600
      TabIndex        =   2
      Top             =   1320
      Width           =   2415
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackColor       =   &H0080C0FF&
      Caption         =   "Due Date: "
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
      Left            =   1440
      TabIndex        =   15
      Top             =   5640
      Width           =   1245
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackColor       =   &H0080C0FF&
      Caption         =   "Issuing Date: "
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
      Left            =   1440
      TabIndex        =   13
      Top             =   4920
      Width           =   1530
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackColor       =   &H0080C0FF&
      Caption         =   "Student's Name: "
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
      Left            =   1440
      TabIndex        =   11
      Top             =   4200
      Width           =   1935
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackColor       =   &H0080C0FF&
      Caption         =   "Student ID: "
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
      Left            =   1440
      TabIndex        =   8
      Top             =   3480
      Width           =   1425
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H0080C0FF&
      Caption         =   "Author: "
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
      Left            =   1440
      TabIndex        =   6
      Top             =   2760
      Width           =   990
   End
   Begin VB.Label Label3 
      BackColor       =   &H0080C0FF&
      Caption         =   "Book Name: "
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
      Left            =   1440
      TabIndex        =   4
      Top             =   2040
      Width           =   1455
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H0080C0FF&
      Caption         =   "Book ID: "
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
      Left            =   1440
      TabIndex        =   1
      Top             =   1320
      Width           =   1155
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00400000&
      Caption         =   "Books Issue Form"
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
      Left            =   2985
      TabIndex        =   0
      Top             =   120
      Width           =   4110
   End
   Begin VB.Image Image1 
      Height          =   7560
      Left            =   0
      Picture         =   "Issuefrm.frx":6988A
      Stretch         =   -1  'True
      Top             =   0
      Width           =   12120
   End
End
Attribute VB_Name = "Issuefrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim db As Database
Dim rs As Recordset
Dim str As String
Dim str1 As String
Dim str2 As String


Private Sub clrcmd_Click()
accntxt.Text = ""
bnametxt.Text = ""
Authortxt.Text = ""
IDtxt.Text = ""
StdNametxt.Text = ""
DTPicker1.Value = Date
DTPicker2.Value = Date + 30
accntxt.SetFocus
End Sub

Private Sub Command1_Click()
Set db = OpenDatabase("F:\own vb6\StdReg.mdb")
str = "select * from Booksdb where BookID='" & accntxt.Text & "'"
Set rs = db.OpenRecordset(str)
If rs.EOF Then
MsgBox ("Record Not Found")
Else
If (accntxt.Text = rs.Fields(0).Value) And (rs.Fields(5).Value = "") Then
bnametxt.Text = rs.Fields(1).Value
Authortxt.Text = rs.Fields(2).Value
Else
MsgBox ("Return the book first")
End If
End If
End Sub

Private Sub Command2_Click()
Set db = OpenDatabase("F:\own vb6\StdReg.mdb")
str1 = "select * from Regfrm where StudentID='" & IDtxt.Text & "'"
Set rs = db.OpenRecordset(str1)
If rs.EOF Then
MsgBox ("Record Not Found")
Else
If rs.Fields(0).Value = IDtxt.Text Then
StdNametxt.Text = rs.Fields(1).Value
End If
End If
End Sub

Private Sub Command5_Click()
Unload Me
mainfrm.Show

End Sub

Private Sub Form_Load()
DTPicker1.Value = Date
DTPicker2.Value = Date + 30

End Sub

Private Sub issuecmd_Click()
Set db = OpenDatabase("F:\own vb6\StdReg.mdb")
str2 = "select * from Booksdb where BookID='" & accntxt.Text & "'"
Set rs = db.OpenRecordset(str2)
If accntxt.Text = "" Then
MsgBox ("Enter Book ID")
ElseIf IDtxt.Text = "" Then
MsgBox ("Enter Student ID")
ElseIf bnametxt.Text = "" Then
MsgBox ("Search for book name")
ElseIf IDtxt.Text = "" Then
MsgBox ("Search for Student's name")
ElseIf rs.Fields(0) = accntxt.Text Then
rs.Edit
rs.Fields(5).Value = IDtxt.Text
rs.Fields(6).Value = StdNametxt.Text
rs.Fields(7).Value = DTPicker1.Value
rs.Fields(8).Value = DTPicker2.Value
rs.Update
MsgBox ("Book Issued")
Else
MsgBox ("Book Not Available")
accntxt.Text = ""
bnametxt.Text = ""
Authortxt.Text = ""
IDtxt.Text = ""
StdNametxt.Text = ""
DTPicker1.Value = Date
DTPicker2.Value = Date + 30
accntxt.SetFocus
End If
End Sub
