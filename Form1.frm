VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7545
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11100
   LinkTopic       =   "Form1"
   ScaleHeight     =   7545
   ScaleWidth      =   11100
   StartUpPosition =   3  'Windows Default
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   330
      Left            =   4200
      Top             =   6960
      Visible         =   0   'False
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   1
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=MSDASQL.1;Persist Security Info=False;Data Source=PostgreSQL35W"
      OLEDBString     =   "Provider=MSDASQL.1;Persist Security Info=False;Data Source=PostgreSQL35W"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "Select estado from testing"
      Caption         =   "Adodc2"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.CommandButton btnBuscar 
      Caption         =   "Buscar"
      Height          =   495
      Left            =   8880
      TabIndex        =   8
      Top             =   1920
      Width           =   1455
   End
   Begin VB.CommandButton btnLimpiar 
      Caption         =   "Limpiar"
      Height          =   495
      Left            =   6960
      TabIndex        =   7
      Top             =   1920
      Width           =   1695
   End
   Begin MSComCtl2.DTPicker DTHasta 
      Height          =   375
      Left            =   3240
      TabIndex        =   6
      Top             =   1920
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   661
      _Version        =   393216
      Format          =   16711681
      CurrentDate     =   44511
   End
   Begin MSComCtl2.DTPicker DTDesde 
      Height          =   375
      Left            =   480
      TabIndex        =   5
      Top             =   1920
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   661
      _Version        =   393216
      Format          =   16711681
      CurrentDate     =   44511
   End
   Begin VB.ComboBox cmbEstado 
      Height          =   315
      ItemData        =   "Form1.frx":0000
      Left            =   6960
      List            =   "Form1.frx":0002
      TabIndex        =   4
      Text            =   "Seleccionar"
      Top             =   720
      Width           =   1815
   End
   Begin VB.ComboBox cmbPrevision 
      Height          =   315
      ItemData        =   "Form1.frx":0004
      Left            =   4320
      List            =   "Form1.frx":0006
      TabIndex        =   3
      Text            =   "Seleccionar"
      Top             =   720
      Width           =   2055
   End
   Begin VB.TextBox txtPaciente 
      Height          =   375
      Left            =   2040
      TabIndex        =   2
      Top             =   720
      Width           =   1815
   End
   Begin VB.TextBox txtRol 
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   720
      Width           =   1575
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Form1.frx":0008
      Height          =   2895
      Left            =   480
      TabIndex        =   0
      Top             =   2760
      Width           =   9735
      _ExtentX        =   17171
      _ExtentY        =   5106
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   360
      Top             =   6960
      Visible         =   0   'False
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   1
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=MSDASQL.1;Persist Security Info=False;Data Source=PostgreSQL35W"
      OLEDBString     =   "Provider=MSDASQL.1;Persist Security Info=False;Data Source=PostgreSQL35W"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "Select * from testing;"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Label Label6 
      Caption         =   "Hasta"
      Height          =   255
      Left            =   3240
      TabIndex        =   14
      Top             =   1560
      Width           =   2055
   End
   Begin VB.Label Label5 
      Caption         =   "Desde"
      Height          =   255
      Left            =   360
      TabIndex        =   13
      Top             =   1560
      Width           =   2055
   End
   Begin VB.Label Label4 
      Caption         =   "Estado"
      Height          =   375
      Left            =   7200
      TabIndex        =   12
      Top             =   360
      Width           =   1335
   End
   Begin VB.Label Label3 
      Caption         =   "Prevision"
      Height          =   255
      Left            =   4440
      TabIndex        =   11
      Top             =   360
      Width           =   1815
   End
   Begin VB.Label Label2 
      Caption         =   "Paciente"
      Height          =   375
      Left            =   2160
      TabIndex        =   10
      Top             =   360
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "Rol"
      Height          =   255
      Left            =   360
      TabIndex        =   9
      Top             =   360
      Width           =   1455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False





Private Sub btnBuscar_Click()
Dim Array_query(6) As String
Dim index As Integer
Dim c As Integer
index = 0
Dim query As String
query = ""
If cmbEstado.List(cmbEstado.ListIndex) <> "" Then
    Array_query(index) = "estado='" & cmbEstado.List(cmbEstado.ListIndex) & "'"
    index = index + 1
End If
If cmbPrevision.List(cmbPrevision.ListIndex) <> "" Then
        Array_query(index) = "prevision='" & cmbPrevision.List(cmbPrevision.ListIndex) & "'"
          index = index + 1
End If
If txtRol.Text <> "" Then
        Array_query(index) = "rol Like %'" & txtRol.Text & "'%"
          index = index + 1
End If
If txtPaciente.Text <> "" Then
        Array_query(index) = "paciente Like %'" & txtPaciente.Text & "'%"
          index = index + 1
End If
Array_query(index) = "alta between '" & Format(DTDesde.Value, "YYYY-MM-DD") & "' and '" & Format(DTHasta.Value, "YYYY-MM-DD") & "'"


For Each Item In Array_query
    If Len(query) > 0 And Item <> "" Then
        query = query & " and (" & Item & ")"
    ElseIf Item <> "" Then
        query = "where " & Item
    End If
Next

Adodc1.RecordSource = "Select * from testing " & query
Adodc1.Refresh
Adodc1.Caption = Adodc1.RecordSource
End Sub

Private Sub btnLimpiar_Click()
txtRol.Text = ""
txtPaciente.Text = ""
cmbEstado.ListIndex = -1
cmbPrevision.ListIndex = -1
DTDesde.Value = Now()
DTHasta.Value = Now()
Adodc1.RecordSource = "Select * from testing "
Adodc1.Refresh
Adodc1.Caption = Adodc1.RecordSource
End Sub

Private Sub Form_Load()
Adodc2.RecordSource = "Select estado from testing " & query
Adodc2.Refresh
Adodc2.Caption = Adodc2.RecordSource
For Each Item In Adodc2.Recordset.GetRows
    cmbEstado.AddItem (Item)
Next

Adodc2.RecordSource = "Select prevision from testing " & query
Adodc2.Refresh
Adodc2.Caption = Adodc2.RecordSource
For Each Item In Adodc2.Recordset.GetRows
    cmbPrevision.AddItem (Item)
Next
End Sub

