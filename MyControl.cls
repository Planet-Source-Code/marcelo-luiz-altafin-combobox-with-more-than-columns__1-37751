VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MyControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Public Key As String

'local variable(s) to hold property value(s)
Private mvarNameControl As Control  'local copy
Private mvarColumnReturn As Integer   'local copy

Friend Property Set NameControl(ByVal vData As Control)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.NameControl = Form1
    Set mvarNameControl = vData
End Property

Friend Property Get NameControl() As Control
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.NameControl
    Set NameControl = mvarNameControl
End Property

Friend Property Let ColumnReturn(ByVal vData As Integer)
    mvarColumnReturn = vData
End Property

Friend Property Get ColumnReturn() As Integer
    ColumnReturn = mvarColumnReturn
End Property
