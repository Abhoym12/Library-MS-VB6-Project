VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Bookcnffrm 
   Caption         =   "Book Configuration Form"
   ClientHeight    =   6675
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   10650
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6675
   ScaleWidth      =   10650
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton backcmd 
      BackColor       =   &H0080FFFF&
      Height          =   495
      Left            =   120
      Picture         =   "Bookdetailsfrm.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   120
      Width           =   495
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   5400
      TabIndex        =   12
      Top             =   3840
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   661
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
      Format          =   113770497
      CurrentDate     =   43584
   End
   Begin VB.TextBox publishertxt 
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
      Height          =   375
      Left            =   5400
      TabIndex        =   11
      Top             =   3240
      Width           =   2535
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
      Height          =   375
      Left            =   5400
      TabIndex        =   10
      Top             =   2640
      Width           =   2535
   End
   Begin VB.CommandButton clrcmd 
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
      Height          =   375
      Left            =   5280
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   4680
      Width           =   1215
   End
   Begin VB.CommandButton savecmd 
      BackColor       =   &H0080FFFF&
      Caption         =   "Save"
      BeginProperty Font 
         Name            =   "Perpetua Titling MT"
         Size            =   15.75
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6720
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   4680
      Width           =   1095
   End
   Begin VB.TextBox titletxt 
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
      Left            =   5400
      TabIndex        =   4
      Top             =   2040
      Width           =   2535
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
      Left            =   5400
      TabIndex        =   2
      Top             =   1440
      Width           =   2535
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackColor       =   &H0080C0FF&
      Caption         =   "Date Acquired: "
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
      Left            =   3120
      TabIndex        =   7
      Top             =   3840
      Width           =   1830
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackColor       =   &H0080C0FF&
      Caption         =   "Publisher: "
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
      Left            =   3120
      TabIndex        =   6
      Top             =   3240
      Width           =   1215
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
      Left            =   3120
      TabIndex        =   5
      Top             =   2640
      Width           =   990
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H0080C0FF&
      Caption         =   "Book Title: "
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
      Left            =   3120
      TabIndex        =   3
      Top             =   2040
      Width           =   1365
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
      Left            =   3120
      TabIndex        =   1
      Top             =   1440
      Width           =   1125
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H0080C0FF&
      Height          =   4695
      Left            =   2760
      Shape           =   4  'Rounded Rectangle
      Top             =   840
      Width           =   5415
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00400000&
      Caption         =   "Books Configuration Form"
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
      Left            =   2160
      TabIndex        =   0
      Top             =   120
      Width           =   6480
   End
   Begin VB.Image Image1 
      Height          =   6660
      Left            =   0
      Picture         =   "Bookdetailsfrm.frx":6988A
      Stretch         =   -1  'True
      Top             =   0
      Width           =   10680
   End
End
Attribute VB_Name = "Bookcnffrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim db As Database
Dim rs As Recordset

Private Sub Command3_Click()
Unload Me
SplashFrm.Show
End Sub

Private Sub Backcmd_Click()
Unload Me
Bookdetailsfrm.Show

End Sub

Private Sub clrcmd_Click()
accntxt.Text = ""
Titletxt.Text = ""
Authortxt.Text = ""
publishertxt.Text = ""
DTPicker1.Value = Date
accntxt.SetFocus
End Sub

Private Sub Form_Load()
DTPicker1.Value = Date
Set db = OpenDatabase("F:\own vb6\StdReg.mdb")
Set rs = db.OpenRecordset("select *from Booksdb")
End Sub

Private Sub savecmd_Click()
rs.AddNew
rs.Fields(0).Value = accntxt.Text
rs.Fields(1).Value = Titletxt.Text
rs.Fields(2).Value = Authortxt.Text
rs.Fields(3).Value = publishertxt.Text
rs.Fields(4).Value = DTPicker1.Value
rs.Update
MsgBox ("Book Data Saved")
End Sub
