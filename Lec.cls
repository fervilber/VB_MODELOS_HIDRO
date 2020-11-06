VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "modHidro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
Option Explicit  'Sentencia que obliga a declarar todas las variables
'Las variables que se declaran son variables internas de las propiedades de la clase (prefijo "mi")

'Para verificar si se ha aceptado el episodio
Dim miEscorrentia() As Double
Dim miIntervalo As Long
Dim miNumeroIntervalos As Long
Dim miHidroUniSCS() As Double
Dim miCaudalHidrounitario() As Double

Dim miIntervaloSCS As Integer
Dim miNumeroIntervalosSCS As Integer



Public Sub LluvianetaSCS(SerieTem, Intervalo As Integer, Optional NC As Double, _
Optional S As Double, Optional P_anterior As Double, Optional UmbralEscorrentia As Double)
'Esta sub calcula la lluvia neta a partir de la lluvia media segun el modelo del SCS
'Los datos de entrada son:
'   un vector "SerieTem" que contenga la serie temporal de lluvia en mm/h o acumulado en intervalo
'   Intervalo=intervalo temporal de la serie o = a -1 si es directamente acumulado en el intervalo
'   NC= Número de curva opcional ya que se puede dar el Umbral de escorrentía
'   S= capacidad de almacenamiento maxima, por defecto =5*Umbral
'   P_anterior= es la precipitación que ha caido antes del inicio de la serie temporal
'   UmbralEscorrentia= Umbral de escorrentia o Po en mm
'El resultado se almacena en una serie temporal o vector que se llama desde fuera de la clase así:
'  NombreEntidaddeClase.Intervalo = n, primero se identifica el instante que se quiere
'  NombreEntidaddeClase.Escorrentia
' es decir se guarda en la propiedad de la clase Escorrrentia, que es un vector

  Dim Dimension As Long
  Dim n As Long
  Dim vector()
  Dim P_acumulada()
  Dim Esc()
  Dim Dif_E As Double
     Dimension = UBound(SerieTem)
     Me.NumeroIntervalos = Dimension
     ReDim vector(1 To Dimension)
     ReDim P_acumulada(1 To Dimension)
     ReDim Esc(1 To Dimension)
     ReDim miEscorrentia(1 To Dimension)
     
        If NC <> 0 Then
          UmbralEscorrentia = (5000 / NC) - 50
         ElseIf S = 0 Then
          S = 5 * UmbralEscorrentia
         ElseIf P_anterior = 0 Then
          
         ElseIf UmbralEscorrentia = 0 Then
           NC = (5000 / (UmbralEscorrentia + 50))
        End If
     
    P_acumulada(1) = SerieTem(1) + P_anterior
    
      If P_acumulada(1) > UmbralEscorrentia Then
        Dif_E = P_acumulada(1)
        Esc(1) = (Dif_E) * ((Dif_E) / (Dif_E + S))
       Else
        Esc(1) = 0
      End If
    

    
   For n = 2 To Dimension
      P_acumulada(n) = SerieTem(n) + P_acumulada(n - 1)
      If P_acumulada(n) > UmbralEscorrentia Then
        Dif_E = (P_acumulada(n) - UmbralEscorrentia)
        Esc(n) = (Dif_E) * ((Dif_E) / (Dif_E + S)) - Esc(n - 1)
        Esc(n) = Esc(n - 1) + Esc(n)
       Else
        Esc(n) = 0
      End If
      
   Next n

        For n = 1 To Dimension
         miEscorrentia(n) = Esc(n)
        Next
End Sub


Public Property Get Escorrentia() As Double  'coordenada x geográfica de la serie
    Escorrentia = miEscorrentia(Me.Intervalo)
End Property

Public Property Let Escorrentia(vNewValue As Double)
    miEscorrentia(Me.Intervalo) = vNewValue
End Property


Public Property Get Intervalo() As Long 'Para el indice de la serie en curso
    Intervalo = miIntervalo
End Property
Public Property Let Intervalo(vNewValue As Long) 'Para el indice de la serie en curso
    If vNewValue > Me.NumeroIntervalos Then
        miIntervalo = Me.NumeroIntervalos
    ElseIf vNewValue < 1 Then
        miIntervalo = 1
    Else
        miIntervalo = vNewValue
    End If
End Property

Public Property Get NumeroIntervalos() As Long 'Para el número de intevalos
    NumeroIntervalos = miNumeroIntervalos
End Property
Public Property Let NumeroIntervalos(vNewValue As Long) 'Para el número de intevalos
    miNumeroIntervalos = vNewValue
End Property



Public Sub HidrogramaUnitSCS(D As Double, Tc As Double, S As Double) ', Optional Npuntos As Integer)
' Modelo de hidrograma unitario de SCS, los datos de entrada son:
'   D= Duración de la tormenta= Intervalo temporal de la serie EN HORAS
'   Tc= Tiempo de concentración en horas
'   S= Superficie de la cuenca en km2
'   Npuntos= numero de puntos en los que se interpola el hidrograma

Dim Tb As Double
Dim Tp As Double
Dim Qp As Double
Dim Numinterval As Integer
Dim n As Integer

Tb = Tc + D
Tp = (D / 2) + 0.35 * Tc
Qp = S / (1.8 * Tb)

 Numinterval = Int(Tb / D)

'Erase miHidroUniSCS

ReDim miHidroUniSCS(1 To Numinterval)
 For n = 1 To Numinterval
 If n * D < Tp Then
  miHidroUniSCS(n) = n * D * (Qp / Tp)
  Else
  miHidroUniSCS(n) = (Tb - n * D) * (Qp / (Tb - Tp))
 End If
 Next n
 
End Sub
Public Property Get HidroUniSCS() As Double  'coordenada x geográfica de la serie
    HidroUniSCS = miHidroUniSCS(Me.IntervaloSCS)
End Property

Public Property Let HidroUniSCS(vNewValue As Double)
    miHidroUniSCS(Me.IntervaloSCS) = vNewValue
End Property


Public Property Get IntervaloSCS() As Integer 'Para el indice de la serie en curso
    IntervaloSCS = miIntervaloSCS
End Property
Public Property Let IntervaloSCS(vNewValue As Integer) 'Para el indice de la serie en curso
    If vNewValue > Me.NumeroIntervalosSCS Then
        miIntervalo = Me.NumeroIntervalosSCS
    ElseIf vNewValue < 1 Then
        miIntervaloSCS = 1
    Else
        miIntervaloSCS = vNewValue
    End If
End Property

Public Property Get NumeroIntervalosSCS() As Integer 'Para el número de intevalos
    NumeroIntervalosSCS = miNumeroIntervalosSCS
End Property
Public Property Let NumeroIntervalosSCS(vNewValue As Integer) 'Para el número de intevalos
    miNumeroIntervalosSCS = vNewValue
End Property

Public Sub CaudalHidroUnit(SerieTem, VectorHidroUnitario, Optional Caudal_Inicial As Double)

 Dim Dimension As Long
 Dim DimensionHidro As Integer
 Dim n As Integer
 
    Dimension = UBound(SerieTem)
    DimensionHidro = UBound(VectorHidroUnitario)
ReDim miCaudalHidrounitario(1 To Dimension)

 miCaudalHidrounitario(0) = Caudal_Inicial
 
    For n = 1 To Dimension
    For i = 1 To DimensionHidro
     miCaudalHidrounitario(n) = SerieTem(n - i + 1) * VectorHidroUnitario(i)
    ' miCaudalHidrounitario(n) = miCaudalHidrounitario(n - 1) + A
    Next i
    Next n

End Sub

