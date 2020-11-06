VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Pluviometro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Esta sus crea y almacena las caracteristicas de los pluviómetros
Private miNombrePluvi As String
Private miCodPluvi As String
'Private misubcuencaAsociada As Integer
Private miCoefThiessen As Double



'Coeficiente de thiessen
Public Property Get CoefThiessen() As Double
    CoefThiessen = miCoefThiessen
End Property
Public Property Let CoefThiessen(vNewValue As Double)
        miCoefThiessen = vNewValue
End Property

'Nombre del Pluviómetro
Public Property Get NombrePluvi() As String
    NombrePluvi = miNombrePluvi
End Property
Public Property Let NombrePluvi(vNewValue As String)
        miNombrePluvi = vNewValue
End Property

'Nombre del SubCuenca asociada
'Public Property Get subcuencaAsociada() As String
'    subcuencaAsociada = misubcuencaAsociada
'End Property
'Public Property Let subcuencaAsociada(vNewValue As String)
'        misubcuencaAsociada = vNewValue
'End Property

'Codigo del Pluviometro
Public Property Get codPluvi() As String
    codPluvi = miCodPluvi
End Property
Public Property Let codPluvi(vNewValue As String)
        miCodPluvi = vNewValue
End Property




