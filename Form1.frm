VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   8070
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   18750
   LinkTopic       =   "Form1"
   ScaleHeight     =   8070
   ScaleWidth      =   18750
   StartUpPosition =   3  'Windows Default
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   375
      Left            =   5040
      TabIndex        =   6
      Top             =   2640
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   661
      _Version        =   393216
      Format          =   141492225
      CurrentDate     =   44509
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   495
      Left            =   1080
      TabIndex        =   5
      Top             =   2520
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   873
      _Version        =   393216
      Format          =   141492225
      CurrentDate     =   44509
   End
   Begin VB.PictureBox DataGrid1 
      Height          =   2415
      Left            =   480
      ScaleHeight     =   2355
      ScaleWidth      =   17355
      TabIndex        =   4
      Top             =   4560
      Width           =   17415
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   11160
      TabIndex        =   3
      Text            =   "Combo2"
      Top             =   720
      Width           =   2535
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   8280
      TabIndex        =   2
      Text            =   "Combo1"
      Top             =   720
      Width           =   2175
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   4440
      TabIndex        =   1
      Text            =   "Text2"
      Top             =   720
      Width           =   2775
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   960
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   720
      Width           =   2175
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Dim cn As ADODB.Connection
Set cn = New ADODB.Connection
cn.ConnectionString = "DRIVER={PostgreSQL Unicode};SERVER=ServerName;port=5432;DATABASE=Database;UID=userID;PWD=password"
'// You'll have to substitute "ServerName", "Database", "userID", and "password"
cn.ConnectionTimeout = 10
cn.Open

End Sub
