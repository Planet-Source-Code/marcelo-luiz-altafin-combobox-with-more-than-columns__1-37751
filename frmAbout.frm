VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "About"
   ClientHeight    =   2295
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2775
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2295
   ScaleWidth      =   2775
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmbOK 
      Caption         =   "&OK"
      Height          =   375
      Left            =   960
      TabIndex        =   2
      Top             =   1800
      Width           =   855
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      Caption         =   "Versão 0.2a"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   4
      Top             =   1320
      Width           =   2535
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      Caption         =   "Componente em desenvolvimento"
      ForeColor       =   &H000000C0&
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   3
      Top             =   960
      Width           =   2535
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      Caption         =   "Marcelo Luiz Altafin"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   2535
   End
   Begin VB.Label Label 
      Caption         =   "altComboBox"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2535
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmbOK_Click()
    Unload Me
End Sub
