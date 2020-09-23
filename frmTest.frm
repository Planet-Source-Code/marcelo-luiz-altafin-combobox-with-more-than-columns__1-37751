VERSION 5.00
Object = "{BFE3972C-AB13-11D6-BE20-AB4C6EEAF16E}#1.0#0"; "altCB.ocx"
Begin VB.Form frmTest 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "altComboBox"
   ClientHeight    =   2820
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6390
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2820
   ScaleWidth      =   6390
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtAddress 
      Height          =   285
      Left            =   720
      TabIndex        =   4
      Top             =   2160
      Width           =   3135
   End
   Begin VB.TextBox txtCode 
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   2160
      Width           =   495
   End
   Begin altCB.altComboBox altComboBox1 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   503
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Address"
      Height          =   195
      Index           =   2
      Left            =   720
      TabIndex        =   5
      Top             =   1920
      Width           =   570
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Code"
      Height          =   195
      Index           =   1
      Left            =   120
      TabIndex        =   3
      Top             =   1920
      Width           =   375
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Name"
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   420
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim adoConnection As New ADODB.Connection

Private Sub Form_Load()
    adoConnection.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\altdb.mdb;Persist Security Info=False"
    With altComboBox1
        .adoConexao = adoConnection
        .SQL = "SELECT Code, Name, Address FROM tbClient"
        .ColumnReturn = 1
        .Columns(1).Caption = "Código"
        .Columns(1).Width = 629
        .Columns(1).Alignment = altCenter
        .Columns(2).Caption = "Nome"
        .Columns(2).Width = 1709
        .Columns(2).Alignment = altLeft
        .Columns(3).Caption = "Endereço"
        .Columns(3).Width = 1934
        .Columns(3).Alignment = altLeft
        .MyControls.Add txtCode, 0
        .MyControls.Add txtAddress, 2
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    adoConnection.Close
    Set adoConnection = Nothing
End Sub
