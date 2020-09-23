VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Column"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Option Explicit

Public Enum Alignment
    altLeft = 0
    altRight = 1
    altCenter = 2
End Enum

Dim Index As Integer

Dim mvarCaption As String
Dim mvarWidth As Integer
Dim mvarAlignment As Integer

Public Property Let Caption(ByVal vData As String)
    mvarCaption = vData
End Property

Public Property Get Caption() As String
    Caption = mvarCaption
End Property

Public Property Let Width(ByVal vData As Integer)
    mvarWidth = vData
End Property

Public Property Get Width() As Integer
    Width = mvarWidth
End Property

Public Property Let Alignment(ByVal vData As Alignment)
    mvarAlignment = vData
End Property

Public Property Get Alignment() As Alignment
    Alignment = mvarAlignment
End Property

Private Sub Class_Initialize()
    mvarWidth = 1514
    mvarAlignment = 0
End Sub
