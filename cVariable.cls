VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cVariable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private mKey As String
Private mVariable As String
Private mNombre As String

Public Property Get Key() As String
    Key = mKey
End Property
Public Property Let Key(ByVal newValue As String)
    mKey = newValue
End Property

Public Property Get Variable() As String
    Variable = mVariable
End Property
Public Property Let Variable(ByVal newValue As String)
    mVariable = newValue
End Property

Public Property Get Nombre() As String
    Nombre = mNombre
End Property
Public Property Let Nombre(ByVal newValue As String)
    mNombre = newValue
End Property



