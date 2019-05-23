VERSION 5.00
Begin VB.Form Datadeletefrm 
   Caption         =   "Delete Data Form"
   ClientHeight    =   5865
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   8715
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5865
   ScaleWidth      =   8715
   StartUpPosition =   2  'CenterScreen
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
      Left            =   7200
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   2640
      Width           =   1095
   End
   Begin VB.CommandButton backcmd 
      BackColor       =   &H0080FFFF&
      Height          =   375
      Left            =   240
      Picture         =   "Bookdeletefrm.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   240
      Width           =   615
   End
   Begin VB.CommandButton Deletecmd 
      BackColor       =   &H0080FFFF&
      Caption         =   "Delete"
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
      Left            =   5640
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   4800
      Width           =   1335
   End
   Begin VB.TextBox Text2 
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
      Height          =   435
      Left            =   4320
      TabIndex        =   6
      Top             =   4080
      Width           =   2655
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
      Height          =   435
      Left            =   4320
      TabIndex        =   4
      Top             =   2640
      Width           =   2655
   End
   Begin VB.ComboBox typecmb 
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
      Height          =   435
      ItemData        =   "Bookdeletefrm.frx":6988A
      Left            =   4320
      List            =   "Bookdeletefrm.frx":69894
      TabIndex        =   2
      Text            =   "Choose type of data"
      Top             =   1320
      Width           =   2655
   End
   Begin VB.Label Label4 
      BackColor       =   &H0080C0FF&
      Caption         =   "Book Title/Student Name: "
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
      Height          =   735
      Left            =   960
      TabIndex        =   5
      Top             =   3960
      Width           =   2475
   End
   Begin VB.Label Label3 
      BackColor       =   &H0080C0FF&
      Caption         =   "Enter Book ID or Student ID: "
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
      Height          =   735
      Left            =   960
      TabIndex        =   3
      Top             =   2520
      Width           =   2475
   End
   Begin VB.Label Label2 
      BackColor       =   &H0080C0FF&
      Caption         =   "Select type of data to be removed: "
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
      Height          =   735
      Left            =   960
      TabIndex        =   1
      Top             =   1200
      Width           =   2415
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00400000&
      Caption         =   "Data Removal Form"
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
      Left            =   1920
      TabIndex        =   0
      Top             =   360
      Width           =   4665
   End
   Begin VB.Image Image1 
      Height          =   5895
      Left            =   0
      Picture         =   "Bookdeletefrm.frx":698A9
      Stretch         =   -1  'True
      Top             =   0
      Width           =   8775
   End
End
Attribute VB_Name = "Datadeletefrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim db As Database
Dim rs As Recordset
Dim str2 As String
Dim str, str1, str3 As String



Private Sub Backcmd_Click()
Unload Me
mainfrm.Show

End Sub

Private Sub Command1_Click()
If typecmb.Text = "Books" Then
Set db = OpenDatabase("F:\own vb6\StdReg.mdb")
str2 = "select * from Booksdb where BookID='" & accntxt.Text & "'"
Set rs = db.OpenRecordset(str2)
If rs.EOF Then
MsgBox ("Record Not Found")
Else
Text2.Text = rs.Fields(1).Value
End If
Else
Set db = OpenDatabase("F:\own vb6\StdReg.mdb")
str3 = "select * from Regfrm where StudentID='" & accntxt.Text & "'"
Set rs = db.OpenRecordset(str3)
If rs.EOF Then
MsgBox ("Record Not Found")
Else
Text2.Text = rs.Fields(1).Value
End If
End If
End Sub

Private Sub Deletecmd_Click()
If typecmb.Text = "Books" Then
Set db = OpenDatabase("F:\own vb6\StdReg.mdb")
str = "select * from Booksdb where BookID='" & accntxt.Text & "'"
Set rs = db.OpenRecordset(str)
If rs.Fields(0) = accntxt.Text Then
rs.Delete
MsgBox ("Delete Successful")
End If
Else
Set db = OpenDatabase("F:\own vb6\StdReg.mdb")
str1 = "select * from Regfrm where StudentID='" & accntxt.Text & "'"
Set rs = db.OpenRecordset(str1)
If rs.Fields(0).Value = accntxt.Text Then
rs.Delete
MsgBox ("Delete Successful")
End If
End If
End Sub

