VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.UserControl altComboBox 
   BackColor       =   &H0080C0FF&
   BackStyle       =   0  'Transparent
   ClientHeight    =   1020
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2220
   ScaleHeight     =   1020
   ScaleWidth      =   2220
   Begin MSDataGridLib.DataGrid dtgCB 
      Height          =   615
      Left            =   0
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   360
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   1085
      _Version        =   393216
      AllowUpdate     =   0   'False
      ColumnHeaders   =   -1  'True
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
            LCID            =   1046
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
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         MarqueeStyle    =   3
         RecordSelectors =   0   'False
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picBotao 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   1800
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   0
      Width           =   240
   End
   Begin VB.TextBox txtGeral 
      Height          =   285
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1695
   End
End
Attribute VB_Name = "altComboBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Dim adorsRegistros As New ADODB.Recordset
Dim vintColumns As New Columns
Dim vintMyControls As New MyControls
Dim vintADOConexao As Connection
Dim vintSQL As String
Dim vintColumnReturn As Integer
Dim vintText As String
Dim AlturaOriginal As Long
Dim LarguraOriginal As Long
Dim DispararEventoResize As Boolean

Const LarguraGrid = 2815
Const AlturaGrid = 560

Public Property Get Columns() As Columns
    Set Columns = vintColumns
End Property

Public Property Let Columns(ByVal vNewValue As Columns)
    Set vintColumns = vNewValue
End Property

Public Property Get MyControls() As MyControls
    Set MyControls = vintMyControls
End Property

Public Property Let MyControls(ByVal vNewValue As MyControls)
    Set vintMyControls = vNewValue
End Property

Public Property Get Text() As String
Attribute Text.VB_MemberFlags = "24"
    Text = vintText
End Property

Public Property Let Text(ByVal vNewValue As String)
    vintText = vNewValue
End Property

Public Property Get AdoConexao() As Connection
    Set AdoConexao = vintADOConexao
End Property

Public Property Let AdoConexao(vNewValue As Connection)
    Set vintADOConexao = vNewValue
End Property

Public Property Let SQL(vNewValue As String)
    vintSQL = vNewValue
    ObterColunas
End Property

Public Property Get SQL() As String
    SQL = vintSQL
End Property

Public Property Let ColumnReturn(vNewValue As Integer)
    vintColumnReturn = vNewValue
End Property

Public Property Get ColumnReturn() As Integer
    ColumnReturn = vintColumnReturn
End Property

Private Sub ObterColunas()
Dim tmp As String
Dim pos As Integer
Dim Coluna As String
Dim cont As Integer
    pos = InStr(1, vintSQL, "FROM", vbBinaryCompare)
    If pos > 0 Then
        tmp = Mid(vintSQL, 8, pos - 9)
        cont = 0
        While Not Trim(tmp) = Empty
            pos = InStr(1, tmp, ",", vbBinaryCompare)
            If pos <> 0 Then
                Coluna = Trim(Mid(tmp, 1, pos - 1))
                tmp = Trim(Mid(tmp, pos + 1))
            Else
                Coluna = Trim(tmp)
                tmp = Empty
            End If
            vintColumns.Add cont
            cont = cont + 1
            vintColumns.Item(cont).Caption = Coluna
        Wend
    Else
        MsgBox "Error in final position from SQL", vbCritical, "altcombobox"
    End If
End Sub

Private Sub ObterRegistroGrid()
    adorsRegistros.CursorLocation = adUseClient
    adorsRegistros.Open vintSQL, vintADOConexao, adOpenStatic, adLockReadOnly, adCmdText
    Set dtgCB.DataSource = adorsRegistros
End Sub

Private Sub dtgCB_LostFocus()
    OcultarGrid
End Sub

Private Sub dtgCB_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim LinhaAtual As Integer
    LinhaAtual = (Y \ Int(dtgCB.RowHeight))
    If LinhaAtual > 0 And LinhaAtual <= dtgCB.VisibleRows Then
        dtgCB.Row = LinhaAtual - 1
    End If
End Sub

Private Sub dtgCB_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim cont As Integer
    If Button = 1 Then
        txtGeral.Text = dtgCB.Columns(vintColumnReturn).Text
        'Obter quantidades de controles extras
        For cont = 1 To vintMyControls.Count
            vintMyControls(cont).NameControl = dtgCB.Columns(vintMyControls(cont).ColumnReturn)
        Next cont
        '---
        OcultarGrid
    End If
End Sub

Private Sub UserControl_Initialize()
    With txtGeral
        .Left = 0
        .Top = 0
    End With
    With picBotao
        .Top = 25
        .Left = (txtGeral.Width - .Width) - 25
        .Picture = LoadResPicture(101, vbResBitmap)
    End With
    With UserControl
        .Height = 275
        .Width = 1695
    End With
    vintColumnReturn = 0
End Sub

Private Sub UserControl_InitProperties()
    DispararEventoResize = True
End Sub

Private Sub UserControl_Resize()
    If DispararEventoResize Then
        With UserControl
            If .Height > 275 Or .Height < 275 Then .Height = 280
            If .Width < 375 Then .Width = 375
            txtGeral.Height = .Height
            txtGeral.Width = .Width
        End With
        With picBotao
            .Top = 25
            .Left = (txtGeral.Width - .Width) - 25
        End With
        With dtgCB
            .Left = 0
            .Top = txtGeral.Top + txtGeral.Height + 7
        End With
    Else
        DispararEventoResize = True
    End If
End Sub

Private Sub UserControl_Show()
    If Not Trim(vintSQL) = Empty Then
        ObterRegistroGrid
    End If
End Sub

Private Sub picBotao_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    picBotao.Picture = LoadResPicture(102, vbResBitmap)
    If Not dtgCB.Visible Then
        adorsRegistros.Requery
        ExibirGrid
        dtgCB.Refresh
        dtgCB.Col = 0
        dtgCB.Row = 1
        dtgCB.SetFocus
        DoEvents
    Else
        OcultarGrid
    End If
End Sub

Private Sub picBotao_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    picBotao.Picture = LoadResPicture(101, vbResBitmap)
End Sub

Private Sub ExibirGrid()
Dim cont As Integer
    AlturaOriginal = UserControl.Height
    LarguraOriginal = UserControl.Width
    DispararEventoResize = False
    UserControl.Height = AlturaOriginal + dtgCB.Height + AlturaGrid
    DispararEventoResize = False
    UserControl.Width = LarguraOriginal + LarguraGrid
    dtgCB.Width = UserControl.Width
    dtgCB.Height = dtgCB.Height + AlturaGrid
    For cont = 1 To vintColumns.Count
        dtgCB.Columns(cont - 1).Caption = vintColumns.Item(cont).Caption
        dtgCB.Columns(cont - 1).Width = vintColumns.Item(cont).Width
        dtgCB.Columns(cont - 1).Alignment = vintColumns.Item(cont).Alignment
    Next cont
    dtgCB.Visible = True
    DoEvents
End Sub

Private Sub OcultarGrid()
    dtgCB.ClearSelCols
    dtgCB.Visible = False
    dtgCB.Height = 615
    dtgCB.Width = 2175
    DispararEventoResize = False
    UserControl.Height = AlturaOriginal
    DispararEventoResize = False
    UserControl.Width = LarguraOriginal
    DoEvents
End Sub

Public Sub ShowAbouBox()
Attribute ShowAbouBox.VB_UserMemId = -552
    frmAbout.Show vbModal
End Sub
