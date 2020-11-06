VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SUBCUENCA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Clase que almacena las variables de las subcuencas

Option Explicit  'Sentencia que obliga a declarar todas las variables

Private miNOMBRE_SUBCUENCA As String
Private miPluvi As Integer
Private miNumeroPluviometros As Integer

Public miColeccionPluviometros As New Collection
'Private miPluvi As New Pluviometro
Private miAREA_KM2 As Double

Private miNC_S As Double
Private miNC_N As Double
Private miNC_H As Double

Private miPUNTO_DRENAJE_ASOCIADO As String
Private miAFORO_AGUAS_ARRIBA As String

Private miTc As Double
Private miDesfase As Double 'Desfase temporal en horas entre las series de caudales aguas arriba y los medidos en el punto de drenaje

Public Sub AñadirPluvi(NombrePluvi As String, CoefThiessen As Double, codPluvi As String)
Dim cpluviometro As New Pluviometro
    
   cpluviometro.NombrePluvi = NombrePluvi
   cpluviometro.CoefThiessen = CoefThiessen
   cpluviometro.codPluvi = codPluvi
  '  Set cpluviometro.subcuencaAsociada = Me.NOMBRE_SUBCUENCA
    
    miColeccionPluviometros.Add cpluviometro, cpluviometro.NombrePluvi
 
End Sub

Public Sub BorrarPluvi(NombrePluvi As String)

    'miColeccionPluviometros.Item "NombrePluvi"
    miColeccionPluviometros.Remove NombrePluvi
 
End Sub

Public Sub compruebaThiessen() 'Ajusta los coeficiente para que sumen 1
Dim n%
Dim Suma As Double
Suma = 0
    For n = 1 To miColeccionPluviometros.Count
       Suma = miColeccionPluviometros.Item(n).CoefThiessen + Suma
    Next n
    
    If Format(Suma, "0.00") <> Format(1, "0.00") Then
        For n = 1 To miColeccionPluviometros.Count
          miColeccionPluviometros.Item(n).CoefThiessen = miColeccionPluviometros.Item(n).CoefThiessen / Suma
        Next n
    End If
End Sub

Public Property Get Pluviometros() As Collection  'Nombre de la subcuencas
  Set Pluviometros = miColeccionPluviometros
End Property

Public Property Get NOMBRE_SUBCUENCA() As String  'Nombre de la subcuencas
    NOMBRE_SUBCUENCA = miNOMBRE_SUBCUENCA
End Property

Public Property Let NOMBRE_SUBCUENCA(vNewValue As String)
    miNOMBRE_SUBCUENCA = vNewValue
End Property

Public Property Get AREA_KM2() As Double  'area en Km 2
    AREA_KM2 = miAREA_KM2
End Property

Public Property Let tc(vNewValue As Double)
    miTc = vNewValue
End Property

Public Property Get tc() As Double  'area en Km 2
    tc = miTc
End Property

Public Property Let AREA_KM2(vNewValue As Double)
    miAREA_KM2 = vNewValue
End Property


Public Property Get NC_S() As Double  'Numero de curva condición seca
    NC_S = miNC_S
End Property

Public Property Let NC_S(vNewValue As Double)
    miNC_S = vNewValue
End Property

Public Property Get NC_N() As Double  'Numero de curva condición normal
    NC_N = miNC_N
End Property

Public Property Let NC_N(vNewValue As Double)
    miNC_N = vNewValue
End Property

Public Property Get NC_H() As Double  'Numero de curva condición humeda
    NC_H = miNC_H
End Property

Public Property Let NC_H(vNewValue As Double)
    miNC_H = vNewValue
End Property


Public Property Get PUNTO_DRENAJE_ASOCIADO() As String  'coordenada x geográfica de la serie
    PUNTO_DRENAJE_ASOCIADO = miPUNTO_DRENAJE_ASOCIADO
End Property

Public Property Let PUNTO_DRENAJE_ASOCIADO(vNewValue As String)
    miPUNTO_DRENAJE_ASOCIADO = vNewValue
End Property


Public Property Get AFORO_AGUAS_ARRIBA() As String  'coordenada x geográfica de la serie
    AFORO_AGUAS_ARRIBA = miAFORO_AGUAS_ARRIBA
End Property

Public Property Let AFORO_AGUAS_ARRIBA(vNewValue As String)
    miAFORO_AGUAS_ARRIBA = vNewValue
End Property

Public Property Let Desfase(vNewValue As Double)
    miDesfase = vNewValue
End Property

Public Property Get Desfase() As Double  'area en Km 2
    Desfase = miDesfase
End Property


Private Sub Class_Initialize()
    'Me.NumeroPluviometros = 0
    Set miColeccionPluviometros = Nothing
   ' Set miPluvi = Nothing
End Sub
