Attribute VB_Name = "modeloHidrologicos"
'
Option Explicit  'Sentencia que obliga a declarar todas las variables
'Las variables que se declaran son variables internas de las propiedades de la clase (prefijo "mi")

Public IMAGsubcuencas()


'Para verificar si se ha aceptado el episodio
Public miEscorrentia() As Double
'Public miIntervalo As Long
Public miNumeroIntervalos As Long
Public miHidroUniSCS() As Double
Public miCaudalHidrounitario() As Double
Public miCaudalAguasArriba() As Double

Public NumCoorHidrUnitSCS As Integer
Public DimensionHidro As Integer 'dimension del hidrograma unitario

Public miIntervaloSCS As Integer
Public miNumeroIntervalosSCS As Integer

Public Intervalo_en_Horas As Double

Public LLuvia_media_acumulada As Double
Public LLuvia_neta_acumulada As Double
Public Max_Ll_neta As Double
Public Max_Ll_media As Double
Public Max_Caudal As Double

Public Volumen_Total As Double


Public Area_Cuenca_Km2 As Double
Public Factor_Intervalo As Double

Public ColecciondeSubcuencas As New Collection
Dim SerieTem()
Dim Superficie_Km2 As Double
Dim Caudal_Inicial

'Fin de definición de variables totales

Public Sub LluvianetaSCS(SerieTem, Optional NC As Double, Optional Factor_Intervalo As Double, _
Optional S As Double, Optional P_anterior As Double, Optional UmbralEscorrentia As Double) ', Intervalo As Integer,
'Esta sub calcula la lluvia neta a partir de la lluvia media segun el modelo del SCS
'Los datos de entrada son:
'   un vector "SerieTem" que contenga la serie temporal de lluvia en mm/h o acumulado en intervalo
'   Intervalo=intervalo temporal de la serie o = a -1 si es directamente acumulado en el intervalo
'   NC= Número de curva opcional ya que se puede dar el Umbral de escorrentía
'   S= capacidad de almacenamiento maxima, por defecto =5*Umbral
'   P_anterior= es la precipitación que ha caido antes del inicio de la serie temporal
'   UmbralEscorrentia= Umbral de escorrentia o Po en mm
'El resultado se almacena en una serie temporal o vector que se llama desde fuera de la clase así:
' el resultado se guarda en miescorrentia
  Dim Dimension As Long
  Dim n As Long
  Dim vector()
  Dim P_acumulada()
  Dim Esc()
  Dim Dif_E As Double
     Dimension = UBound(SerieTem)
     ReDim vector(1 To Dimension)
     ReDim P_acumulada(1 To Dimension)
     ReDim Esc(1 To Dimension)
     ReDim miEscorrentia(1 To Dimension)
     
 'El Factor_Intervalo  sirve para pasar los datos de lluvia en intensidad a acumulados _
  por ejemplo si los datos están en mm/h y el intervalo de la serie temporal es de 30 minutos _
  el factor valdría Factor_Intervalo=0.5
  Max_Ll_media = 0
 If Factor_Intervalo = 0 Then Factor_Intervalo = 1 'And Factor_Intervalo <> 1
    For n = 1 To Dimension
     'Multiplica cada valor de la serie temporal por el factor que pasa de intensidad _
     a acumulado en intervalo
     SerieTem(n) = Factor_Intervalo * SerieTem(n)
     
     'Comprueba que no hay valores menores que cero
     If SerieTem(n) < 0 Then SerieTem(n) = 0
     
     'Calcula el acumulado de la serie
     LLuvia_media_acumulada = LLuvia_media_acumulada + SerieTem(n)
     
     'Calcula el maximo de la serie (ojo que son mm y no intensidades)
     If SerieTem(n) > Max_Ll_media Then Max_Ll_media = SerieTem(n)
    Next n
'
        If NC <> 0 Then
          UmbralEscorrentia = (5000 / NC) - 50
         ElseIf NC = 0 Then
           NC = (5000 / (UmbralEscorrentia + 50))
        End If
     If S = 0 Then S = 5 * UmbralEscorrentia
     
      P_acumulada(1) = SerieTem(1) + P_anterior
      If P_acumulada(1) > UmbralEscorrentia Then
        Dif_E = (P_acumulada(1) - UmbralEscorrentia)
        Esc(1) = (Dif_E) * ((Dif_E) / (Dif_E + S))
       Else
        Esc(1) = 0
      End If
'
   For n = 2 To Dimension
      P_acumulada(n) = SerieTem(n) + P_acumulada(n - 1)
      
      If P_acumulada(n) > UmbralEscorrentia Then
      
        Dif_E = (P_acumulada(n) - UmbralEscorrentia)
        'lo siguiente es una adicción de prueba para ver si funciona mejor el modelo
        If Val(LEE_EPI.textparametro.Text) > 0 Then
         S = S + n * (Dif_E / Val(LEE_EPI.textparametro.Text)) / Dimension
        End If
        Esc(n) = (Dif_E) ^ 2 / ((Dif_E + S)) ' - Esc(n - 1)
       
       Else
        Esc(n) = 0
      End If
   Next n
   
        LLuvia_neta_acumulada = Esc(Dimension)
        
        miEscorrentia(1) = Esc(1)
        Max_Ll_neta = miEscorrentia(1)
        For n = 2 To Dimension
         miEscorrentia(n) = Esc(n) - Esc(n - 1)
         If miEscorrentia(n) < 0 Then miEscorrentia(n) = 0
         If miEscorrentia(n) > Max_Ll_neta Then Max_Ll_neta = miEscorrentia(n)
        Next
End Sub

Public Sub HidrogramaUnitSCS(D As Double, tc As Double, S As Double) ', Optional Npuntos As Integer)
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

Tb = tc + D
Tp = (D / 2) + 0.35 * tc
Qp = S / (1.8 * Tb)

 NumCoorHidrUnitSCS = Int(Tb / D)

Erase miHidroUniSCS

ReDim miHidroUniSCS(1 To NumCoorHidrUnitSCS)
 For n = 1 To NumCoorHidrUnitSCS
 If n * D < Tp Then
  miHidroUniSCS(n) = n * D * (Qp / Tp)
  Else
  miHidroUniSCS(n) = (Tb - n * D) * (Qp / (Tb - Tp))
 End If
 Next n
 
End Sub

Public Sub CaudalHidroUnit(SerieTemLLuvianeta, VectorHidroUnitario, Optional Caudal_Inicial As Double)

 Dim Dimension As Long
 
 Dim i As Integer
 Dim k%
 Dim hasta As Integer
 
    Dimension = UBound(SerieTemLLuvianeta)
    DimensionHidro = UBound(VectorHidroUnitario)
    ReDim miCaudalHidrounitario(0 To (Dimension + DimensionHidro))

    miCaudalHidrounitario(0) = Caudal_Inicial
 
    For k = 1 To (Dimension + DimensionHidro)
      miCaudalHidrounitario(k) = Caudal_Inicial
      For i = 1 To DimensionHidro - 1
       If (k - i) > 0 And (k - i) <= Dimension Then
        miCaudalHidrounitario(k) = SerieTemLLuvianeta(k - i) * VectorHidroUnitario(i + 1) + miCaudalHidrounitario(k)
       End If
      Next i
    Next k
    
'Parametros globales:
    Volumen_Total = 0
    Max_Caudal = 0
    For k = 1 To (Dimension + DimensionHidro)
     Volumen_Total = ((miCaudalHidrounitario(k) * Intervalo_en_Horas * 3600) / 1000000) + Volumen_Total
     If miCaudalHidrounitario(k) > Max_Caudal Then Max_Caudal = miCaudalHidrounitario(k)
    Next k


End Sub

Public Sub SubcuencaSCS(SerieTem, Intervalo_en_h As Double, Factor_Intervalo As Double, _
Tc_en_h As Double, Superficie_Km2 As Double, Caudal_Inicial As Double, _
Optional NC As Double, Optional P_anterior As Double, Optional UmbralEscorrentia As Double, _
Optional Desfase_en_h As Double = 0)
'Esta sub calcula la lluvia neta y el caudal a partir de la lluvia media
'Se calcula le lluvia neta
'If UmbralEscorrentia = 0 Then
Dim k As Integer

  Intervalo_en_Horas = Intervalo_en_h
If LEE_EPI.MSFlexGrid2.Visible = False Then
 LluvianetaSCS SerieTem, NC, Factor_Intervalo, , P_anterior, UmbralEscorrentia
Else
 Ll_Neta_Per_IniConst SerieTem, Val(LEE_EPI.MSFlexGrid2.TextMatrix(0, 1)), Val(LEE_EPI.MSFlexGrid2.TextMatrix(1, 1)), Val(LEE_EPI.MSFlexGrid2.TextMatrix(2, 1)), Factor_Intervalo
End If
'El resultado está en miEscorrentia

'Se calcula las coordenadas del hidrgrama unitario
 HidrogramaUnitSCS Intervalo_en_h, Tc_en_h, Superficie_Km2
'El resultado está en miHidroUniSCS y el numero de puntos que tiene es NumCoorHidrUnitSCS

'Se calcula el caudal segun SCS
 CaudalHidroUnit miEscorrentia, miHidroUniSCS, Caudal_Inicial
'Se almacena en miCaudalHidrounitario


If UBound(miCaudalAguasArriba) > 0 Then

'Se suma la a la serie el caudal aguas arriba mas el retraso
Retraso Desfase_en_h, Intervalo_en_h

For k = 1 To UBound(miCaudalHidrounitario) - 1
 miCaudalHidrounitario(k) = miCaudalAguasArriba(k) + miCaudalHidrounitario(k)
Next k
End If
End Sub

Public Sub Retraso(Desfase As Double, IntervaloTemporal)
'Calcula la nueva serie de caudales desfasados la unidades de tiempo especificadas
'Desfase es en unidades de tiempo (horas)
Dim NuevaSerie()
Dim NumerodeDatos As Long
Dim NumerodeDatos1 As Long
Dim Avance As Integer ' es el numero de puestos que hay que correr la serie
Dim n%

NumerodeDatos = UBound(miCaudalHidrounitario)
NumerodeDatos1 = UBound(miCaudalAguasArriba)
If NumerodeDatos1 > NumerodeDatos Then NumerodeDatos1 = NumerodeDatos

Avance = Int(Desfase / IntervaloTemporal)
ReDim Preserve miCaudalAguasArriba(1 To NumerodeDatos)
If Desfase = 0 Then Exit Sub
ReDim NuevaSerie(1 To NumerodeDatos)



'hace que no se reste el -1 de fallo
  For n = 1 To NumerodeDatos
    If miCaudalAguasArriba(n) = -1 Then miCaudalAguasArriba(n) = 0
  Next n
  
  For n = 1 To Avance
    NuevaSerie(n) = miCaudalAguasArriba(1)
  Next n
  
  For n = Avance + 1 To NumerodeDatos1 + Avance
    NuevaSerie(n) = miCaudalAguasArriba(n - Avance)
  Next n
  For n = NumerodeDatos1 + Avance To NumerodeDatos
    NuevaSerie(n) = miCaudalAguasArriba(NumerodeDatos1 + Avance)
  Next n

  For n = 1 To NumerodeDatos
   ' miCaudalAguasArriba(n) = NuevaSerie(n)
   miCaudalAguasArriba(n) = NuevaSerie(n)
  Next n

End Sub

Public Sub Inicia_nuevo_calculo()
'Inicializa las variable antes de hacer otro cálculo
 LLuvia_media_acumulada = 0
 LLuvia_neta_acumulada = 0
 Volumen_Total = 0
 
 Max_Ll_neta = 0
 Max_Ll_media = 0
 Max_Caudal = 0
 
ReDim miHidroUniSCS(0 To 0)

ReDim miEscorrentia(1 To 2)
ReDim miCaudalHidrounitario(1 To 2)

End Sub

Public Function FechaEpi(nombrefichero As String) As Date
'Esta función da la fecha de final de un fichero de episodio _
a partir de su nombre.
Dim Mes
Dim Dia
Dim Hora
Dim Año
Dim Min
'Fecha de creacion de la función 9-ene-2001

nombrefichero = Trim(nombrefichero)
Mes = Left(nombrefichero, 2)
Dia = Mid(nombrefichero, 3, 2)
Hora = Mid(nombrefichero, 5, 2)
Min = Mid(nombrefichero, 7, 2)
Año = Right(nombrefichero, 2)

FechaEpi = Dia & "/" & Mes & "/" & Año & " " & Hora & ":" & Min

End Function

Public Function RUT(RUTACOMPLETA As String) As String
'RUTACOMPLETA es la direccion completa de un fichero
' la funcion RUT quita de la parte final el nombre del fichero dejando solo la ruta
Dim i%
Dim A
For i = Len(RUTACOMPLETA) To 1 Step -1
    RUT = Mid(RUTACOMPLETA, 1, i)
    A = Right(RUT, 1)
    If A = "\" Then Exit For
Next
End Function


Public Function CombinacionLineal(SerieTem1, SerieTem2, A As Double, B As Double) As Collection
'Esta sub calcula la combinacion lineal de dos series temporales es decir:
' da = a*serietem1 + b*serietem2
 Dim Dimension As Long
 
 Dim i As Long
 Dim serieResultado()
 Dim hasta As Integer
 Dim DimensionMax As Long
 Dim DimensionMin As Long
 
 
    DimensionMax = UBound(SerieTem1)
    DimensionMin = LBound(SerieTem1)
    
    If DimensionMax <> UBound(SerieTem2) Then Exit Function
    
    ReDim serieResultado(DimensionMin To DimensionMax)
    
    For i = DimensionMin To DimensionMax
     serieResultado(i) = A * SerieTem1(i) + B * SerieTem2(i)
    Next i

End Function


Public Sub T_SubcuencaSCS(NOMBRE_SUBCUENCA As String, Intervalo_en_h As Double, Factor_Intervalo As Double, Tc_en_h As Double, Caudal_Inicial As Double, _
Optional NC As Double, Optional P_anterior As Double, Optional UmbralEscorrentia As Double)
'Esta sub calcula la lluvia neta y el caudal a partir de la lluvia media
'Se calcula le lluvia neta


Intervalo_en_Horas = Intervalo_en_h

If LEE_EPI.MSFlexGrid2.Visible = False Then
 LluvianetaSCS SerieTem, NC, Factor_Intervalo, , P_anterior, UmbralEscorrentia
Else
 Ll_Neta_Per_IniConst SerieTem, Val(LEE_EPI.MSFlexGrid2.TextMatrix(0, 0)), Val(LEE_EPI.MSFlexGrid2.TextMatrix(1, 0)), Val(LEE_EPI.MSFlexGrid2.TextMatrix(2, 0)), Factor_Intervalo
End If

'Se calcula las coordenadas del hidrgrama unitario
 HidrogramaUnitSCS Intervalo_en_h, Tc_en_h, Superficie_Km2
'El resultado está en miHidroUniSCS y el numero de puntos que tiene es NumCoorHidrUnitSCS

'Se calcula el caudal segun SCS
 CaudalHidroUnit miEscorrentia, miHidroUniSCS, Caudal_Inicial


End Sub


Public Sub Ll_Neta_Per_IniConst(SerieTem, UmbraldeEscorentia As Double, _
Capacidad_de_Infiltracion As Double, Precipitacion_Inicial As Double, _
Factor_Intervalo As Double)   ', Intervalo As Integer,

'Esta sub calcula la lluvia neta a partir de la lluvia media segun el modelo de pérdidas exponencial
'Los datos de entrada son:
'   un vector "SerieTem" que contenga la serie temporal de lluvia en mm/h o acumulado en intervalo
'   Intervalo=intervalo temporal de la serie o = a -1 si es directament
'   UmbraldeEscorentia= Umbral de escorrentía
'   Capacidad_de_Infiltracion = la lluvia que se infiltra en mm/h
'   Precipitacion_Inicial= es la precipitación que ha caido antes del inicio de la serie temporal
'El resultado se almacena en una serie temporal o vector que se llama desde fuera de la clase así:
' el resultado se guarda en miescorrentia
  Dim Dimension As Long
  Dim n As Long
  Dim vector()
  Dim P_acumulada()
  Dim Esc()
  Dim Dif_E As Double
     Dimension = UBound(SerieTem)
     ReDim vector(1 To Dimension)
     ReDim P_acumulada(1 To Dimension)
     ReDim Esc(1 To Dimension)
     ReDim miEscorrentia(1 To Dimension)

 'El Factor_Intervalo  sirve para pasar los datos de lluvia en intensidad a acumulados _
  por ejemplo si los datos están en mm/h y el intervalo de la serie temporal es de 30 minutos _
  el factor valdría Factor_Intervalo=0.5
  Max_Ll_media = 0
 If Factor_Intervalo = 0 Then Factor_Intervalo = 1 'And Factor_Intervalo <> 1
    For n = 1 To Dimension
     'Multiplica cada valor de la serie temporal por el factor que pasa de intensidad _
     a acumulado en intervalo
     SerieTem(n) = Factor_Intervalo * SerieTem(n)
     
     'Comprueba que no hay valores menores que cero
     If SerieTem(n) < 0 Then SerieTem(n) = 0
     
     'Calcula el acumulado de la serie
     LLuvia_media_acumulada = LLuvia_media_acumulada + SerieTem(n)
     
     'Calcula el maximo de la serie (ojo que son mm y no intensidades)
     If SerieTem(n) > Max_Ll_media Then Max_Ll_media = SerieTem(n)
    Next n
     
      P_acumulada(1) = SerieTem(1) + Precipitacion_Inicial
      If P_acumulada(1) > UmbraldeEscorentia Then
         If SerieTem(1) > Capacidad_de_Infiltracion Then
           Esc(1) = SerieTem(1) - Capacidad_de_Infiltracion
         Else
         Esc(1) = 0
         End If
      Else
         Esc(1) = 0
      End If
    
   For n = 2 To Dimension
      P_acumulada(n) = SerieTem(n) + P_acumulada(n - 1)
      If P_acumulada(n) > UmbraldeEscorentia Then
         If SerieTem(n) > Capacidad_de_Infiltracion Then
           Esc(n) = SerieTem(n) - Capacidad_de_Infiltracion
         Else
            Esc(n) = 0
         End If
       Else
        Esc(n) = 0
      End If
   Next n
   
        LLuvia_neta_acumulada = Esc(Dimension)
        miEscorrentia(1) = Esc(1)
        Max_Ll_neta = miEscorrentia(1)
        For n = 2 To Dimension
         miEscorrentia(n) = Esc(n) '- Esc(n - 1)
         If miEscorrentia(n) > Max_Ll_neta Then Max_Ll_neta = miEscorrentia(n)
        Next
End Sub

Public Sub Subcuenca(SerieTem, Intervalo_en_h As Double, Factor_Intervalo As Double, _
Tc_en_h As Double, Superficie_Km2 As Double, Caudal_Inicial As Double, _
Optional NC As Double, Optional P_anterior As Double, Optional UmbralEscorrentia As Double, _
Optional Desfase_en_h As Double = 0, Optional A)
'Esta sub calcula la lluvia neta y el caudal a partir de la lluvia media
'Se calcula le lluvia neta
'If UmbralEscorrentia = 0 Then
Dim k As Integer

  Intervalo_en_Horas = Intervalo_en_h
' Ll_Neta_Per_IniConst SerieTem, UmbralEscorrentia, Capacidad_de_Infiltracion, Precipitacion_Inicial
'El resultado está en miEscorrentia

'Se calcula las coordenadas del hidrgrama unitario
 HidrogramaUnitSCS Intervalo_en_h, Tc_en_h, Superficie_Km2
'El resultado está en miHidroUniSCS y el numero de puntos que tiene es NumCoorHidrUnitSCS

'Se calcula el caudal segun SCS
 CaudalHidroUnit miEscorrentia, miHidroUniSCS, Caudal_Inicial
'Se almacena en miCaudalHidrounitario


If UBound(miCaudalAguasArriba) > 0 Then

'Se suma la a la serie el caudal aguas arriba mas el retraso
Retraso Desfase_en_h, Intervalo_en_h

For k = 1 To UBound(miCaudalHidrounitario) - 1
 miCaudalHidrounitario(k) = miCaudalAguasArriba(k) + miCaudalHidrounitario(k)
Next k
End If
End Sub


