VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cDato"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private mKey As String
Private mValor As Double
Private mFecha As String

Public Property Get Key() As String
    Key = mKey
End Property
Public Property Let Key(ByVal newValue As String)
    mKey = newValue
End Property

Public Property Get Valor() As Double
    Valor = mValor
End Property
Public Property Let Valor(ByVal newValue As Double)
    mValor = newValue
End Property
Public Property Get Fecha() As String
    Fecha = mFecha
End Property
Public Property Let Fecha(ByVal newValue As String)
    mFecha = newValue
End Property


