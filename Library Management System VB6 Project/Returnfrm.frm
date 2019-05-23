VERSION 5.00
Begin VB.Form Returnfrm 
   Caption         =   "Book Return"
   ClientHeight    =   6720
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   10065
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6720
   ScaleWidth      =   10065
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Searchtxt 
      BackColor       =   &H0080FFFF&
      Caption         =   "Search"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   7920
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   1155
      Width           =   1215
   End
   Begin VB.CommandButton Backcmd 
      BackColor       =   &H0080FFFF&
      Height          =   375
      Left            =   240
      Picture         =   "Returnfrm.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   360
      Width           =   615
   End
   Begin VB.CommandButton Clearcmd 
      BackColor       =   &H0080FFFF&
      Caption         =   "Clear"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4680
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   4920
      Width           =   1215
   End
   Begin VB.CommandButton Returncmd 
      BackColor       =   &H0080FFFF&
      Caption         =   "Return"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6240
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   4920
      Width           =   1455
   End
   Begin VB.TextBox StdIDtxt 
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
      Left            =   4920
      TabIndex        =   10
      Top             =   3360
      Width           =   2775
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
      Left            =   4920
      TabIndex        =   9
      Top             =   4080
      Width           =   2775
   End
   Begin VB.TextBox Authortxt 
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
      Left            =   4920
      TabIndex        =   6
      Top             =   2640
      Width           =   2775
   End
   Begin VB.TextBox Titletxt 
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
      Left            =   4920
      TabIndex        =   4
      Top             =   1920
      Width           =   2775
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
      Left            =   4920
      TabIndex        =   2
      Top             =   1200
      Width           =   2775
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackColor       =   &H0080C0FF&
      Caption         =   "Student Name: "
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
      TabIndex        =   8
      Top             =   4080
      Width           =   1800
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
      Left            =   2160
      TabIndex        =   7
      Top             =   3360
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
      Left            =   2160
      TabIndex        =   5
      Top             =   2640
      Width           =   990
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H0080C0FF&
      Caption         =   "Title: "
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
      Top             =   1920
      Width           =   690
   End
   Begin VB.Label Label2 
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
      Left            =   2160
      TabIndex        =   1
      Top             =   1200
      Width           =   1125
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00400000&
      Caption         =   "Return Books Form"
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
      Left            =   2640
      TabIndex        =   0
      Top             =   240
      Width           =   4725
   End
   Begin VB.Image Image1 
      Height          =   6840
      Left            =   0
      Picture         =   "Returnfrm.frx":6988A
      Stretch         =   -1  'True
      Top             =   0
      Width           =   10080
   End
End
Attribute VB_Name = "Returnfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim db As Database
Dim rs As Recordset
Dim str As String
Dim str1 As String


Private Sub Backcmd_Click()
Unload Me
mainfrm.Show
End Sub

Private Sub Clearcmd_Click()
accntxt.Text = ""
Titletxt.Text = ""
Authortxt.Text = ""
StdIDtxt.Text = ""
StdNametxt.Text = ""
accntxt.SetFocus
End Sub

Private Sub Returncmd_Click()
Set db = OpenDatabase("F:\own vb6\StdReg.mdb")
str1 = "Select * from Booksdb where BookID='" & accntxt.Text & "'"
Set rs = db.OpenRecordset(str1)
If rs.Fields(5).Value = "" Then
MsgBox ("Issue the book first")
Else
If rs.Fields(0).Value = accntxt.Text Then
rs.Edit
rs.Fields(5).Value = ""
rs.Fields(6).Value = ""
rs.Fields(7).Value = ""
rs.Fields(8).Value = ""
rs.Update
MsgBox ("Book Returned")
'accntxt.Text = ""
'bnametxt.Text = ""
'authortxt.Text = ""
'IDtxt.Text = ""
'StdNametxt.Text = ""
'DTPicker1.Value = Date
'DTPicker2.Value = Date + 30
'accntxt.SetFocus
End If
End If
End Sub

Private Sub Searchtxt_Click()
Set db = OpenDatabase("F:\own vb6\StdReg.mdb")
str = "select * from Booksdb where BookID='" & accntxt.Text & "'"
Set rs = db.OpenRecordset(str)
If rs.EOF Then
MsgBox ("Record Not Found")
Else
If rs.Fields(0) = accntxt.Text Then
Titletxt.Text = rs.Fields(1).Value
Authortxt.Text = rs.Fields(2).Value
StdIDtxt.Text = rs.Fields(5).Value
StdNametxt.Text = rs.Fields(6).Value
End If
End If
End Sub
