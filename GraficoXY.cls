VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GraficoXY"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Variables asociadas a propiedades publicas
Dim miContenedor As Object 'el objeto donde se pinta el grafico: puede ser un picturebox y  un contenedor
Dim miDivisionVertical As Integer 'las divisiones verticales del objeto donde se pinta
Dim miDivisionHorizontal As Integer 'lo mismo pero verticales
Dim miPosicionMarco As Integer 'la situacion donde situar el grafico
Dim miX1Etiqueta, miX2Etiqueta As Variant '
Dim miY1Etiqueta, miY2Etiqueta As Double
Dim miValorInvalido As Double
Dim miDato As Long 'indica la posicion del dato en los vectores de datos
Dim miNumeroDatos As Integer
Dim miXDato() As Variant 'vector donde se almacenan los valores de X
Dim miYDato() As Double  'vector donde se almacenar los valores de Y
Dim miTipoGrafico As String 'lineas o barras
Dim miFechaInicial As Date 'no hace falta
Dim miFechaFinal As Date  'no hace falta
Dim miRejilla As Boolean 'para pintar la rejilla
Dim miTituloGrafico As String 'el titulo del grafico
Dim miTituloEjeX As String 'el titulo del eje X
Dim miTituloEjeY As String 'el titulo del eje y
Dim miX1Marco, miX2Marco As Single 'extremos horizontales del marco
Dim miY1Marco, miY2Marco As Single  'extremos verticales del marco
Dim miAnchuraMarco As Single 'anchura del marco
Dim miAnchuraEjes As Single 'anchura de los ejes
Dim miAnchuraDatos As Single 'anchura de la linea del grafico
Dim miAnchuraRejilla As Single 'anchura rejilla
Dim miColorEjes As Long 'color de los ejes
Dim miColorDato() As Long 'color de cada uno de los datos
Dim miSiColorDato As Boolean 'datos con colores individualizados
Dim miColorDatos As Long 'color de los datos
Dim miColorRejilla As Long  'color de la rejilla
Dim miColorMarco As Long 'color del marco
Dim miSiEtiquetasY As Boolean
Dim miIntervaloTemporal As Integer  'para el tiempo que se toma para los intervalos temporales
            'por ejemplo, en series de episodio es de 60 minutos
'Variables internas de la clase
Dim A, B As Single 'transformacion de coordenadas
Dim c, D As Single 'transformacion de coordenadas
Dim miX1Eje, miX2Eje As Single 'extremos del eje X
Dim miY1Eje, miY2Eje As Single 'extremo del eje Y
Dim miMaximoX, miMinimoX As Single   'extremos de los valores de X
Dim miMaximoY, miMinimoY As Single 'extremos de los valores de Y
Dim miFormato As String 'formato para escribir las etiquetas
Dim miColor As Long 'color de las lineas
Dim miMarcosRellenos() As Boolean
Dim miEtiquetasPintables As Boolean 'para pintar las etiquetas de los ejes
Dim miTitulosPintables As Boolean 'para pintar los titulos
Dim miSegmentosPintables As Boolean 'para pintar segemetacion
'el valor del Error datos temporales 0.5 segundos
Const Error_T As Double = 1.15740740740741E-05 / 2#
Const Error_Y As Double = 0.0000001
Public Property Get SiColorDato() As Boolean
    SiColorDato = miSiColorDato
End Property
Public Property Let SiColorDato(vNewValue As Boolean)
    miSiColorDato = vNewValue
End Property

Public Property Get ColorDato() As Long
    ColorDato = miColorDato(Me.Dato)
End Property
Public Property Let ColorDato(vNewValue As Long)
    miColorDato(Me.Dato) = vNewValue
End Property
Private Sub Class_Initialize()
'iniciación de la etiquetas
Me.X1Etiqueta = -1E+38
Me.X2Etiqueta = 1E+38
Me.Y1Etiqueta = -1E+38
Me.Y2Etiqueta = 1E+38

'iniciacion de los marcos
Me.X1Marco = -1E+38
Me.Y1Marco = -1E+38
Me.X2Marco = 1E+38
Me.Y2Marco = 1E+38

Me.TipoGrafico = "LINEAS"
Me.ValorInvalido = -9999
Me.ColorMarco = 0
Me.Rejilla = False
Me.TituloGrafico = ""
Me.TituloEjeX = ""
Me.TituloEjeY = ""
miFormato = "#0"
miDato = 0

'anchuras
Me.AnchuraMarco = 1
Me.AnchuraEjes = 1
Me.AnchuraDatos = 1
Me.AnchuraRejilla = 1

'colores
Me.ColorMarco = QBColor(0)
Me.ColorEjes = QBColor(0)
Me.ColorDatos = QBColor(0)
Me.ColorRejilla = QBColor(0)

'etiquetas para pintar
miEtiquetasPintables = True

'datos extremos
miMaximoX = -1E+38
miMaximoY = -1E+38
miMinimoX = 1E+38
miMinimoY = 1E+38
miSiEtiquetasY = True


'para los datos temporales
miIntervaloTemporal = 60
End Sub



Public Property Get DivisionVertical() As Integer
    DivisionVertical = miDivisionVertical
End Property

Public Property Let DivisionVertical(vNewValue As Integer)
    miDivisionVertical = vNewValue
End Property



Public Property Get DivisionHorizontal() As Integer
    'las partes verticales en que se divide el contenedor
    DivisionHorizontal = miDivisionHorizontal
End Property

Public Property Let DivisionHorizontal(vNewValue As Integer)
    miDivisionHorizontal = vNewValue
    
    
End Property

Public Sub DibujaMarco()
    'Calculo las coordenadas extremas del marco y lo dibuja posteriormente
    Dim Fila, Columna As Integer
    Dim i As Integer
    Dim AnchuraContenedor As Single
    Dim AlturaContenedor As Single
    Dim Borde As Single
    Dim BordeVertical, BordeHorizontal As Single
    Dim aux As Single


'el borde es la distancia desde el marco hasta los ejes
    
    Borde = 1 / 200
    AnchuraContenedor = Me.Contenedor.ScaleWidth
    AlturaContenedor = Me.Contenedor.ScaleHeight

    BordeVertical = Borde * AlturaContenedor
    BordeHorizontal = Borde * AnchuraContenedor

'localizar la posicion vetical, en cual linea esta
    If Me.DivisionVertical <= 0 Then
        Me.DivisionVertical = 1
    End If

    If Me.DivisionHorizontal <= 0 Then
        Me.DivisionHorizontal = 1
    End If

    If Me.PosicionMarco <= 0 Then
        Me.PosicionMarco = 1
    End If
    
    If Me.PosicionMarco > Me.DivisionVertical * Me.DivisionHorizontal Then
        Me.PosicionMarco = 1
    End If

    For i = 1 To Me.DivisionVertical * Me.DivisionHorizontal
        If Me.PosicionMarco <= i * Me.DivisionVertical Then Exit For
    Next i
    Fila = i
    

    For i = 1 To Me.DivisionVertical
        If (Me.PosicionMarco - ((Fila - 1) * Me.DivisionVertical)) = i Then Exit For
    Next i
    Columna = i

    'calcular los extremos del marco
    If Me.X1Marco = -1E+38 And Me.Y1Marco = -1E+38 And Me.X2Marco = 1E+38 And Me.Y2Marco = 1E+38 Then
        'no hay eleccion externa
      'extremos inferior derecho
        miX1Marco = BordeHorizontal + (Columna - 1) * AnchuraContenedor / Me.DivisionVertical
        miY1Marco = BordeVertical + (Fila - 1) * (AlturaContenedor) / Me.DivisionHorizontal
       'extremos superior izquierdo
        miY2Marco = Fila * (AlturaContenedor) / Me.DivisionHorizontal - BordeVertical
        miX2Marco = Columna * AnchuraContenedor / Me.DivisionVertical - BordeHorizontal
        
        
    End If

'coherencia de marco
If miX1Marco >= miX2Marco Or miY1Marco >= miY2Marco Then
    MsgBox "miX1Marco >= miX2Marco O miY1Marco >= miY2Marco", vbOKOnly, "GráficoXY"
    Exit Sub
End If

'pintar el marco
'cuatro llamadas al procedimito PintarLinea
If Me.AnchuraMarco <> 0 Then
    Call PintarLinea(miX1Marco, miY1Marco, miX1Marco, miY2Marco, miColorMarco, Me.AnchuraMarco)
    Call PintarLinea(miX1Marco, miY2Marco, miX2Marco, miY2Marco, miColorMarco, Me.AnchuraMarco)
    Call PintarLinea(miX2Marco, miY2Marco, miX2Marco, miY1Marco, miColorMarco, Me.AnchuraMarco)
    Call PintarLinea(miX2Marco, miY1Marco, miX1Marco, miY1Marco, miColorMarco, Me.AnchuraMarco)
End If

End Sub

Public Property Let Contenedor(vNewValue As Object)
    Set miContenedor = vNewValue
    miContenedor.AutoRedraw = True
End Property
Public Property Get Contenedor() As Object
    Set Contenedor = miContenedor
End Property

Public Property Get XDato() As Variant
    XDato = miXDato(Me.Dato)
End Property
Public Property Let XDato(vNewValue As Variant)
    miXDato(Me.Dato) = vNewValue
    If vNewValue > miMaximoX Then
        miMaximoX = vNewValue
    End If
    If vNewValue < miMinimoX Then
        miMinimoX = vNewValue
    End If
End Property

Public Property Get YDato() As Double
    YDato = miYDato(Me.Dato)
End Property
Public Property Let YDato(vNewValue As Double)
    miYDato(Me.Dato) = vNewValue
    If vNewValue > miMaximoY And vNewValue <> miValorInvalido Then
        miMaximoY = vNewValue
    End If
    If vNewValue < miMinimoY And vNewValue <> miValorInvalido Then
        miMinimoY = vNewValue
    End If
End Property

Public Property Get TipoGrafico() As String
    'grafico del tipo lineas, barras o puntos
    TipoGrafico = miTipoGrafico
End Property
Public Property Let TipoGrafico(vNewValue As String)
    miTipoGrafico = vNewValue
End Property

Private Sub Ejes()
'este procedimiento calcula y dibuja los ejes
    Dim MargenSuperior, MargenInferior As Single
    Dim MargenDerecho, MargenIzquierdo As Single
    Dim XTamañoEtiquetas, YtamañoEtiquetas As Single
    Dim aux, aux2 As Single
    Dim i As Integer
    Dim Formato As String
    Dim TitulosVisiblesX As Single
    Dim TitulosVisiblesY As Single
    Dim EtiquetasVisiblesX As Single
    Dim EtiquetasVisiblesY As Single
    Dim SegmentosVisiblesX As Single
    Dim SegmentosVisiblesY As Single
    'valores iniciales
    miEtiquetasPintables = True
    miTitulosPintables = True
    miSegmentosPintables = True
    Formato = "#0.00"

    'coeficientes de visibilidad de titulos y etiquetas
    ' proporcionan longitudes en relación a (miX2Marco - miX1Marco) por
    ' producto de esta diferencia (ó miY2Marco - miY1Marco)
    TitulosVisiblesX = 1
    TitulosVisiblesY = 1
    EtiquetasVisiblesX = 1.5
    EtiquetasVisiblesY = 1.5
    SegmentosVisiblesX = 2#
    SegmentosVisiblesY = 2#

    'MARGENES PARA EL EJE Y
    'margen superior
    'separacion del marco
    aux = Me.Contenedor.TextHeight("0") 'separacion del marco
    
    'separacion por el titulo de grafico y medio espacio
    MargenSuperior = aux + 1.5 * Me.Contenedor.TextHeight(Me.TituloGrafico)
    
    'margen inferior
    'separcion del margen
    aux = 0.5 * Me.Contenedor.TextHeight("0")
     
    'separación por el titulo y medio espacio mas
    'si existe titulo
    If Me.TituloEjeX <> "" Then
        aux = aux + 1.5 * Me.Contenedor.TextHeight(Me.TituloEjeX)
    End If
     
    'separcion por las etiquetas y medio espacio mas por etiqueta
    Me.Dato = 1
    If IsDate(Me.XDato) Then
        'para fechas
        aux = aux + 2 * Me.Contenedor.TextHeight("0")
    Else
        'no fechas
        aux = aux + 1 * Me.Contenedor.TextHeight("0")
    End If
    
    'separacion por la segmentacion
    MargenInferior = aux + 0.5 * Me.Contenedor.TextHeight("0")
    
    'comparacion del primer criterio
    If (MargenSuperior + MargenInferior + 2 * Me.Contenedor.TextHeight("0")) > TitulosVisiblesY * (miY2Marco - miY1Marco) Then
        miTitulosPintables = False
    End If
    
    'segundo criterio, las etiquetas
    If (MargenSuperior + MargenInferior + 2 * Me.Contenedor.TextHeight("0")) > EtiquetasVisiblesY * (miY2Marco - miY1Marco) Then
        miEtiquetasPintables = False
        miTitulosPintables = False
    End If
    
    'tercer criterio
    If (MargenSuperior + MargenInferior + 2 * Me.Contenedor.TextHeight("0")) > SegmentosVisiblesY * (miY2Marco - miY1Marco) Then
        miEtiquetasPintables = False
        miTitulosPintables = False
        miSegmentosPintables = False
    End If
    
    'separacion del eje X
    'separcion con el marco izquierdo
    aux = 1 * Me.Contenedor.TextWidth("0") 'separacion del marco
    
    'separacion por el titulo de eje y un medio espacio mas
    If Me.TituloEjeY <> "" Then
        aux = aux + 2 * Me.Contenedor.TextWidth(Left(Me.TituloEjeY, 1))
    End If
    
    'separcion por las etiquetas
    If Me.Y1Etiqueta <> -1E+38 Then
        'elegida la etiqueta inferior del eje Y
            aux2 = Len(Format(Me.Y1Etiqueta, Formato))
    Else
        'no esta elejida
        If miMinimoY = 1E+38 Then
            miMinimoY = -1
        End If
        aux2 = Len(Format(miMinimoY, Formato))
    End If
    
    If Me.Y2Etiqueta <> 1E+38 Then
        'esta elegida la segunda etiqueta
        If Len(Format(Me.Y2Etiqueta, Formato)) > aux2 Then
            aux2 = Len(Format(Me.Y2Etiqueta, Formato))
        End If
    Else
        'no esta elejida
        If Len(Format(miMaximoY, Formato)) > aux2 Then
            If miMaximoY = -1E+38 Then
                miMaximoY = -1
            End If
            aux2 = Len(Format(miMaximoY, Formato))
        End If
    End If
    
    'conozco el tamaño mayor que pueden tener las etiquetas
    'y un espacio mas
    aux = aux + (aux2 + 0.5) * Me.Contenedor.TextWidth("0")
    
    'separacion por la segmetacion
    MargenIzquierdo = aux + 0.5 * Me.Contenedor.TextWidth("0")
    
    'separacion del marco derecho
    Me.Dato = Me.NumeroDatos
    If IsDate(Me.XDato) Then
        'para fechas
        MargenDerecho = 0.75 * Me.Contenedor.TextWidth(Format(Me.XDato, "dd/mm/yy")) + 0.5 * Me.Contenedor.TextWidth("0")
    Else
        'no fechas
        MargenDerecho = 0.75 * Me.Contenedor.TextWidth(Me.XDato) 'datos no fechas
    End If
    
    'primer criterio, los titulos
    If IsDate(Me.XDato) Then
        'son fechas
        If (MargenDerecho + MargenIzquierdo + 1 * Me.Contenedor.TextWidth(Format(Me.XDato, "dd/mm/yy"))) > TitulosVisiblesX * (miX2Marco - miX1Marco) Then
            miTitulosPintables = False
        End If
    Else
        If (MargenDerecho + MargenIzquierdo + 1.5 * Me.Contenedor.TextWidth(Format(miMaximoX, Formato))) > TitulosVisiblesX * (miX2Marco - miX1Marco) Then
            miTitulosPintables = False
        End If
    End If
    
    'segundo criterio, las etiquetas
    If IsDate(Me.XDato) Then
        'son fechas
        If (MargenDerecho + MargenIzquierdo + 1.5 * Me.Contenedor.TextWidth(Format(Me.XDato, "dd/mm/yy"))) > EtiquetasVisiblesX * (miX2Marco - miX1Marco) Then
            miEtiquetasPintables = False
            miTitulosPintables = False
        End If
    Else
        If (MargenDerecho + MargenIzquierdo + 1.5 * Me.Contenedor.TextWidth(Format(miMaximoX, Formato))) > EtiquetasVisiblesX * (miX2Marco - miX1Marco) Then
            miEtiquetasPintables = False
            miTitulosPintables = False
        End If
    End If
        
    'tercer criterio
    If (MargenDerecho + MargenIzquierdo + 1.5 * Me.Contenedor.TextHeight("0")) > SegmentosVisiblesX * (miX2Marco - miX1Marco) Then
        miEtiquetasPintables = False
        miTitulosPintables = False
        miSegmentosPintables = False
    End If
    
    'consecuencias del primer criterio
    If miTitulosPintables = False Then
        MargenSuperior = MargenSuperior - 1.5 * Me.Contenedor.TextHeight(Me.TituloGrafico)
        If Me.TituloEjeX <> "" Then
            MargenInferior = MargenInferior - 1.5 * Me.Contenedor.TextHeight(Me.TituloEjeX)
        End If
        MargenIzquierdo = MargenIzquierdo - 1.5 * Me.Contenedor.TextWidth(Left(Me.TituloEjeY, 1))
    End If
    
    'consecuencias del segundo criterio
    If miEtiquetasPintables = False Then
        MargenIzquierdo = MargenIzquierdo - (aux2) * Me.Contenedor.TextWidth("0")
        
        If IsDate(Me.XDato) Then
            MargenInferior = MargenInferior - 2 * Me.Contenedor.TextHeight("0")
            MargenDerecho = MargenDerecho - 0.5 * Me.Contenedor.TextWidth(Format(Me.XDato, "dd/mm/yy"))
        Else
            MargenInferior = MargenInferior - 1.5 * Me.Contenedor.TextHeight("0")
            MargenDerecho = MargenDerecho - 1.25 * Me.Contenedor.TextWidth("0")
        End If
    End If
    
    'consecuencias del tercer criterio
    If miSegmentosPintables = False Then
        MargenIzquierdo = 0
        MargenDerecho = 0
        MargenSuperior = 0
        MargenInferior = 0
    End If

    miX1Eje = miX1Marco + MargenIzquierdo
    miX2Eje = miX2Marco - MargenDerecho
    miY1Eje = miY2Marco - MargenInferior
    miY2Eje = miY1Marco + MargenSuperior

    
    'pintar los ejes
    If Me.AnchuraEjes <> 0 Then
        Call PintarLinea(miX1Eje, miY1Eje, miX2Eje, miY1Eje, Me.ColorEjes, Me.AnchuraEjes)   'eje x
        Call PintarLinea(miX1Eje, miY1Eje, miX1Eje, miY2Eje, Me.ColorEjes, Me.AnchuraEjes) 'eje y
    End If

End Sub

Private Sub Etiquetas()
        'este proceso calcula el numero de etiquetas puede haber en los ejes
       
        Call EtiquetasX
             
        'para saltar las etiquetas Y
        If Me.SiEtiquetasY = False Then Exit Sub
        
        Call EtiquetasY


    
End Sub




Public Property Get SiEtiquetasY() As Boolean
    SiEtiquetasY = miSiEtiquetasY
End Property

Public Property Let SiEtiquetasY(v As Boolean)
miSiEtiquetasY = v
End Property



Public Sub DibujaDatos()
    'distribuye el tipo de grafico a dibujar
    Select Case UCase(Me.TipoGrafico)
    Case "BARRAS"
        GraficoBarras
    Case "BARRAS H"
        GraficoBarras
    Case "LINEAS"
        GraficoLineas
    Case "PUNTOS"
        GraficoPuntos
    Case Else
        MsgBox "No reconoce el tipo de grafico", , "GraficoXY"
    End Select

End Sub



Private Sub GraficoBarras()

Dim X1, Y1 As Single
Dim X2, Y2 As Single
Dim Intervalo As Variant
Dim i As Integer
Dim condicion1 As Boolean
Dim condicion2 As Boolean
Dim colordat As Long

'transformacion
'bucle de datos
i = 1
Do While i <= Me.NumeroDatos
    Me.Dato = i
    If IsDate(Me.XDato) Then
        condicion1 = Me.XDato + Error_T >= Me.X1Etiqueta And Me.XDato - Error_T <= Me.X2Etiqueta
    Else
        condicion1 = (Me.XDato >= Me.X1Etiqueta And Me.XDato <= Me.X2Etiqueta)
    End If
    If condicion1 Then
        'para los datos en fallo
        Do While Me.YDato = Me.ValorInvalido
            DatoNulo (i)
            i = i + 1
            If i > Me.NumeroDatos Then Exit Sub
            Me.Dato = i
        Loop
    
        'los datos no tienen fallo
        
        If i = 1 Then
            X1 = X_a_Xc(Me.X1Etiqueta)
        Else
            Me.Dato = i - 1
            X1 = X_a_Xc(Me.XDato)
            Me.Dato = i
        End If
        X2 = X_a_Xc(Me.XDato)
    
        Me.Dato = i
        Y2 = Y_a_Yc(Y1Etiqueta) ' Y_a_Yc(0)  ¡las barras empiezan en el eje x, no en el valor 0
                            'las series van del mínimo al máximo, pero el mínimo no tiene por que ser 0
        
        'parte añadida
        If SiColorDato = True Then  'para cada dato un color (barra de colores múltiples)
            colordat = Me.ColorDato
        Else
            colordat = Me.ColorDatos  'todos los datos tienen el mismo color (barra de un color)
        End If
        
        Select Case Me.YDato
        Case Is > Me.Y2Etiqueta + Error_Y
            Y1 = Y_a_Yc(Me.Y2Etiqueta)
            Call PintarRectangulo(X1, Y1, X2, Y2, colordat, Me.AnchuraDatos)
            Y2 = Y_a_Yc(Me.Y2Etiqueta)
            Y1 = Y_a_Yc(Me.Y2Etiqueta) - 0.1 * (Y_a_Yc(Me.Y2Etiqueta) - Me.Y1Marco)
            Call PintarRectangulo(X1, Y1, X2, Y2, QBColor(12), Me.AnchuraDatos)
        Case Is < Me.Y1Etiqueta - Error_Y
            Y1 = Y_a_Yc(Me.Y1Etiqueta)
            Call PintarRectangulo(X1, Y1, X2, Y2, colordat, Me.AnchuraDatos)
            Y2 = Y_a_Yc(Me.Y1Etiqueta)
            Y1 = Y_a_Yc(Me.Y1Etiqueta) + 0.1 * (Me.Y2Marco - Y_a_Yc(Me.Y1Etiqueta))
            Call PintarRectangulo(X1, Y1, X2, Y2, QBColor(12), Me.AnchuraDatos)
        Case Else
            Y1 = Y_a_Yc(Me.YDato)
            Call PintarRectangulo(X1, Y1, X2, Y2, colordat, Me.AnchuraDatos)
        End Select
            
    End If
    i = i + 1
Loop

End Sub

Private Sub GraficoLineas()

Dim X1, Y1 As Single
Dim X2, Y2 As Single
Dim i As Integer

If Me.NumeroDatos = 1 Then
    GraficoPuntos
    Exit Sub
End If



'transformacion
For i = 1 To Me.NumeroDatos
    Me.Dato = i
    'para los datos en fallo
    Do While Me.YDato = Me.ValorInvalido
        DatoNulo (i)
        i = i + 1
        If i > Me.NumeroDatos Then Exit Sub
        Me.Dato = i
    Loop
    
    '
    If Me.XDato >= Me.X1Etiqueta And Me.XDato <= Me.X2Etiqueta Then
        If Me.Dato = 1 Then
            X1 = X_a_Xc(Me.XDato)
        Else
            If miYDato(Me.Dato - 1) = ValorInvalido Then
                X1 = X_a_Xc(miXDato(Me.Dato - 1))
            Else
                X1 = X_a_Xc(Me.XDato)
            End If
        End If
        Y1 = Y_a_Yc(Me.YDato)
    End If
    
    'Sustituir un dato en fallo
    If (i + 1) > Me.NumeroDatos Then Exit For
    Me.Dato = i + 1

    If Me.XDato >= Me.X1Etiqueta And Me.XDato <= (Me.X2Etiqueta + Error_T) Then
        X2 = X_a_Xc(Me.XDato)
        Y2 = Y_a_Yc(Me.YDato)
    End If
   
    'un vez calculada la transformacion hay que saber si hay que pintarla
    If (Me.YDato <> Me.ValorInvalido) Then
        If Y1 < miY2Eje And Y2 < miY2Eje Then
        'los dos puntos por debajo del eje
        Else
            If Y1 < miY2Eje Then
                'el punto incial esta fuera
                Y1 = miY2Eje
            End If
            If Y2 < miY2Eje Then
                'el punto final esta fuera
                Y2 = miY2Eje
            End If
            If Y1 > miY1Eje And Y2 > miY1Eje Then
            '
            Else
                If Y1 > miY1Eje Then
                    Y1 = miY1Eje
                End If
                If Y2 > miY1Eje Then
                    Y2 = miY1Eje
                End If
                'los datos estan dentro de los eje
                If Me.AnchuraDatos <> 0 And X1 <> Empty Then
                    Call PintarLinea(X1, Y1, X2, Y2, Me.ColorDatos, Me.AnchuraDatos)
                End If
            End If
        End If
    End If
Next i

End Sub

Private Sub DatoNulo(Posicion As Integer)
'este procedimiento pinta los datos nulos
'printa una barra de color rojo donde detecta un dato nulo
Dim X1, Y1 As Single
Dim X2, Y2 As Single
    
    If Posicion = 1 Then
        Me.Dato = Posicion
        X1 = X_a_Xc(Me.X1Etiqueta)
        Y1 = Y_a_Yc(Me.Y2Etiqueta)
    Else
        Me.Dato = Posicion - 1
        X1 = X_a_Xc(Me.XDato)
        Y1 = Y_a_Yc(Me.Y2Etiqueta)
    End If
    Me.Dato = Posicion
    X2 = X_a_Xc(Me.XDato)
    Y2 = Y_a_Yc(Me.Y1Etiqueta)
    
If TypeName(Contenedor) <> "Contenedor" Then

   
    miContenedor.Line (X1, Y1)-(X2, Y2), QBColor(12), BF
Else

   Call miContenedor.Line_(X1, Y1, X2, Y2, QBColor(12), "BF")
End If

End Sub

Public Property Get X1Etiqueta() As Variant
    X1Etiqueta = miX1Etiqueta
End Property
Public Property Let X1Etiqueta(vNewValue As Variant)
    miX1Etiqueta = vNewValue
End Property

Public Property Get X2Etiqueta() As Variant
    X2Etiqueta = miX2Etiqueta
End Property
Public Property Let X2Etiqueta(vNewValue As Variant)
    miX2Etiqueta = vNewValue
End Property

Public Property Get Y1Etiqueta() As Double
    Y1Etiqueta = miY1Etiqueta
End Property
Public Property Let Y1Etiqueta(vNewValue As Double)
    miY1Etiqueta = vNewValue
End Property

Public Property Get Y2Etiqueta() As Double
    Y2Etiqueta = miY2Etiqueta
End Property
Public Property Let Y2Etiqueta(vNewValue As Double)
    miY2Etiqueta = vNewValue
End Property

Public Function X_a_Xc(ByVal XCOORD) As Single
    X_a_Xc = A * (XCOORD - Me.X1Etiqueta) + B
End Function

Public Function Y_a_Yc(ByVal YCOORD) As Single
    Y_a_Yc = c * (YCOORD - Me.Y1Etiqueta) + D
End Function

Public Sub Dibuja()
    If Me.Contenedor Is Nothing Then
        MsgBox "No se especificó el contenedor", , "Gráfico XY"
        Exit Sub
    End If

    'este procedimiento regula el flujo de proceso de la clase GraficoXY
    'Primero distribuye el tipo grafico a pintar
    'Dibuja los datos en el interior del marco
    'Dibuja la rejilla
    
    'Calcula y dibuja los marcos
    DibujaMarco
    'Calculo y dibuja los ejes
    Ejes
    'Calcula y dibuja las etiquetas de los ejes
    Etiquetas
    'calculo de las constantes caractericticas del grafico
    CalculaConstantes
    DibujaDatos
    'Imprime los titulos del grafico y de los ejes
    Accesorios
    Final
End Sub

Public Property Get PosicionMarco() As Integer
    'situacion del grafico
    PosicionMarco = miPosicionMarco
End Property
Public Property Let PosicionMarco(vNewValue As Integer)
    miPosicionMarco = vNewValue
End Property

Public Property Get ValorInvalido() As Double
    ValorInvalido = miValorInvalido
End Property
Public Property Let ValorInvalido(vNewValue As Double)
    miValorInvalido = vNewValue
End Property

Public Property Get Dato() As Long
    'indice de la matriz de datos
    Dato = miDato
End Property
Public Property Let Dato(vNewValue As Long)
    miDato = vNewValue
End Property

Public Property Get FechaInicial() As Date
    FechaInicial = miFechaInicial
End Property
Public Property Let FechaInicial(vNewValue As Date)
    miFechaInicial = vNewValue
End Property

Public Property Let FechaFinal(vNewValue As Date)
    miFechaFinal = vNewValue
End Property
Public Property Get FechaFinal() As Date
    FechaFinal = miFechaFinal
End Property

Public Property Get Rejilla() As Boolean
    'opcion de dibujar la rejilla o no
    Rejilla = miRejilla
End Property
Public Property Let Rejilla(vNewValue As Boolean)
    miRejilla = vNewValue
End Property

Public Property Get TituloGrafico() As String
    TituloGrafico = miTituloGrafico
End Property
Public Property Let TituloGrafico(vNewValue As String)
    miTituloGrafico = vNewValue
End Property

Private Sub Accesorios()
    Dim TamañoTexto As Single
    Dim NumeroEspacios As Integer
    Dim i As Integer
    Dim Xflotante, Yflotante As Single
    Dim Titulo As String
    Dim LetraAnterior As Single

    'Pinta los titulos si existen
    'del grafico
    TamañoTexto = 0
    Titulo = ""
    If Trim(Me.TituloGrafico) <> "" Then
        Yflotante = miY1Marco + 0.5 * Me.Contenedor.TextHeight(Me.TituloGrafico)
    
    'existe grafico
        For i = 1 To Len(Me.TituloGrafico)
            If Mid(Me.TituloGrafico, i, 1) <> "^" Then
                TamañoTexto = TamañoTexto + Me.Contenedor.TextWidth(Mid(Me.TituloGrafico, i, 1))
            End If
        Next i
        Xflotante = miX1Marco + 0.5 * (miX2Marco - miX1Marco) - 0.5 * TamañoTexto
    
        For i = 1 To Len(Me.TituloGrafico)
            Select Case Mid(Me.TituloGrafico, i, 1)
                Case Is <> "^"
                    Titulo = Titulo + Mid(Me.TituloGrafico, i, 1)
                Case Is = "^"
                    LetraAnterior = Me.Contenedor.FontSize
                    Me.Contenedor.CurrentX = Xflotante
                    Me.Contenedor.CurrentY = Yflotante
                    If miTitulosPintables = True Then
                        If TypeName(Contenedor) <> "Contenedor" Then


                            Me.Contenedor.Print Titulo
                        Else
                            Me.Contenedor.Print_ Titulo
                        End If
                    End If
                    Xflotante = Xflotante + Me.Contenedor.TextWidth(Titulo)
                    Me.Contenedor.FontSize = 0.75 * LetraAnterior
                    i = i + 1
                    Me.Contenedor.CurrentX = Xflotante
                    Me.Contenedor.CurrentY = Yflotante
                    If miTitulosPintables = True Then
                        Me.Contenedor.Print_ Mid(Me.TituloGrafico, i, 1)
                        If TypeName(Contenedor) <> "Contenedor" Then


                            Me.Contenedor.Print Mid(Me.TituloGrafico, i, 1)
                        Else
                            Me.Contenedor.Print_ Mid(Me.TituloGrafico, i, 1)
                        End If
                    End If
                    Xflotante = Xflotante + Me.Contenedor.TextWidth(Mid(Me.TituloGrafico, i, 1))
                    Titulo = ""
                    Me.Contenedor.FontSize = LetraAnterior
            End Select
        Next i
        Me.Contenedor.CurrentX = Xflotante
        Me.Contenedor.CurrentY = Yflotante
        If miTitulosPintables = True Then
            If TypeName(Contenedor) <> "Contenedor" Then


                Me.Contenedor.Print Titulo
            Else
                Me.Contenedor.Print_ Titulo
            End If
        End If
    End If

    'del eje X
    Titulo = ""
    TamañoTexto = 0
    If Trim(Me.TituloEjeX) <> "" Then
        Yflotante = miY2Marco - 1.5 * Me.Contenedor.TextHeight(Me.TituloEjeX)
        'existe grafico
        For i = 1 To Len(Me.TituloEjeX)
            If Mid(Me.TituloEjeX, i, 1) <> "^" Then
                TamañoTexto = TamañoTexto + Me.Contenedor.TextWidth(Mid(Me.TituloEjeX, i, 1))
            End If
        Next i
        Xflotante = miX2Eje - TamañoTexto
    
        For i = 1 To Len(Me.TituloEjeX)
            Select Case Mid(Me.TituloEjeX, i, 1)
                Case Is <> "^"
                    Titulo = Titulo + Mid(Me.TituloEjeX, i, 1)
                Case Is = "^"
                    LetraAnterior = Me.Contenedor.FontSize
                    Me.Contenedor.CurrentX = Xflotante
                    Me.Contenedor.CurrentY = Yflotante
                    
                        If TypeName(Contenedor) <> "Contenedor" Then

                            Me.Contenedor.Print Titulo
                        Else
                            Me.Contenedor.Print_ Titulo
                        End If
                        
                    Xflotante = Xflotante + Me.Contenedor.TextWidth(Titulo)
                    Me.Contenedor.FontSize = 0.75 * LetraAnterior
                    i = i + 1
                    Me.Contenedor.CurrentX = Xflotante
                    Me.Contenedor.CurrentY = Yflotante
                    If miTitulosPintables = True Then
                        If TypeName(Contenedor) <> "Contenedor" Then

                            Me.Contenedor.Print Mid(Me.TituloEjeX, i, 1)
                        Else
                            Me.Contenedor.Print_ Mid(Me.TituloEjeX, i, 1)
                        End If
                    End If
                    Xflotante = Xflotante + Me.Contenedor.TextWidth(Mid(Me.TituloEjeX, i, 1))
                    Titulo = ""
                    Me.Contenedor.FontSize = LetraAnterior
            End Select
        Next i
        Me.Contenedor.CurrentX = Xflotante
        Me.Contenedor.CurrentY = Yflotante
        If miTitulosPintables = True Then
            If TypeName(Contenedor) <> "Contenedor" Then

                Me.Contenedor.Print Titulo
            Else
                Me.Contenedor.Print_ Titulo
            End If
        End If
    End If

    'del eje Y
    Titulo = ""
    Me.TituloEjeY = Trim(Me.TituloEjeY)
    If Me.TituloEjeY <> "" Then
        NumeroEspacios = Len(Trim(Me.TituloEjeY))
        If NumeroEspacios > 7 Then
            'pintar vertical
            For i = 1 To NumeroEspacios
                Titulo = Mid(Me.TituloEjeY, i, 1)
                TamañoTexto = Me.Contenedor.TextWidth(Titulo)
                Me.Contenedor.CurrentX = (miX1Eje - (3 * Me.Contenedor.TextWidth("0") + Me.Contenedor.TextWidth(Int(miMaximoY))))
                Me.Contenedor.CurrentX = Me.Contenedor.CurrentX - 0.5 * (Me.Contenedor.CurrentX - miX1Marco) - 0.5 * TamañoTexto
                Me.Contenedor.CurrentY = miY2Eje + 0.75 * (i - 1) * Me.Contenedor.TextHeight(Titulo)
                If miTitulosPintables = True Then
                    If TypeName(Contenedor) <> "Contenedor" Then

                        Me.Contenedor.Print Titulo
                    Else
                        Me.Contenedor.Print_ Titulo
                    End If
                End If
            Next i
        Else
            Yflotante = miY2Eje - 0.5 * Me.Contenedor.TextHeight(Me.TituloEjeY)
            'pintar horizontal
            LetraAnterior = Me.Contenedor.FontSize
            TamañoTexto = 0
            For i = 1 To Len(Me.TituloEjeY)
                If Mid(Me.TituloEjeY, i, 1) <> "^" Then
                    TamañoTexto = TamañoTexto + Me.Contenedor.TextWidth(Mid(Me.TituloEjeY, i, 1))
                End If
            Next i
            Xflotante = 0.5 * (miX1Eje - miX1Marco) - 0.75 * TamañoTexto
        
            For i = 1 To NumeroEspacios
                If Mid(Me.TituloEjeY, i, 1) = "^" Then
                    Me.Contenedor.FontSize = 0.75 * LetraAnterior
                    
                    i = i + 1
                
                    Me.Contenedor.CurrentX = Xflotante
                    Me.Contenedor.CurrentY = Yflotante
                    Me.Contenedor.Print_ Mid(Me.TituloEjeY, i, 1)
                    If TypeName(Contenedor) <> "Contenedor" Then

                        Me.Contenedor.Print Mid(Me.TituloEjeY, i, 1)
                    Else
                        Me.Contenedor.Print_ Mid(Me.TituloEjeY, i, 1)
                    End If
                    Xflotante = Xflotante + Me.Contenedor.TextWidth(Mid(Me.TituloEjeY, i, 1))
                    Me.Contenedor.FontSize = LetraAnterior
                Else
                    Me.Contenedor.CurrentX = Xflotante
                    Me.Contenedor.CurrentY = Yflotante
                    If TypeName(Contenedor) <> "Contenedor" Then

                        Me.Contenedor.Print Mid(Me.TituloEjeY, i, 1)
                    Else
                        Me.Contenedor.Print_ Mid(Me.TituloEjeY, i, 1)
                    End If
                    Xflotante = Xflotante + Me.Contenedor.TextWidth(Mid(Me.TituloEjeY, i, 1))
                End If
            Next i
            Me.Contenedor.CurrentX = Xflotante
            Me.Contenedor.CurrentY = Yflotante
            If miTitulosPintables = True Then
                    If TypeName(Contenedor) <> "Contenedor" Then

                        Me.Contenedor.Print Titulo
                    Else
                        Me.Contenedor.Print_ Titulo
                    End If
            End If
        End If
  
    End If

End Sub

Public Property Get TituloEjeX() As String
    TituloEjeX = miTituloEjeX
End Property
Public Property Let TituloEjeX(vNewValue As String)
    miTituloEjeX = vNewValue
End Property

Public Property Get TituloEjeY() As String
    TituloEjeY = miTituloEjeY
End Property
Public Property Let TituloEjeY(vNewValue As String)
    miTituloEjeY = vNewValue
End Property

Public Property Get NumeroDatos() As Integer
    'numero total de datos de la serie a representar
    NumeroDatos = miNumeroDatos
End Property
Public Property Let NumeroDatos(vNewValue As Integer)
    If vNewValue < 1 Then Exit Property
    
    miNumeroDatos = vNewValue
    ReDim miXDato(1 To miNumeroDatos)
    ReDim miYDato(1 To miNumeroDatos)
    ReDim miColorDato(1 To miNumeroDatos)
    
End Property

Public Property Get ColorMarco() As Long
    ColorMarco = miColorMarco
End Property
Public Property Let ColorMarco(vNewValue As Long)
    miColorMarco = vNewValue
End Property

Private Sub Final()
'para borrar de la memoria las etiquetas
'    If Me.DivisionHorizontal = 1 And Me.DivisionVertical = 1 Then
'        Me.X1Etiqueta = -1E+38
'        Me.X2Etiqueta = 1E+38
'        Me.Y1Etiqueta = -1E+38
'        Me.Y2Etiqueta = 1E+38
'    End If
    If Me.X1Marco <> -1E+38 Then
        Me.X1Marco = -1E+38
    End If
    If Me.Y1Marco <> -1E+38 Then
        Me.Y1Marco = -1E+38
    End If
    If Me.X2Marco <> 1E+38 Then
        Me.X2Marco = 1E+38
    End If
    If Me.Y2Marco <> 1E+38 Then
        Me.Y2Marco = 1E+38
    End If

End Sub

Public Property Get X1Marco() As Single
    X1Marco = miX1Marco
End Property
Public Property Let X1Marco(vNewValue As Single)
    miX1Marco = vNewValue
End Property

Public Property Get Y1Marco() As Single
    Y1Marco = miY1Marco
End Property
Public Property Let Y1Marco(vNewValue As Single)
    miY1Marco = vNewValue
End Property

Public Property Get X2Marco() As Single
    X2Marco = miX2Marco
End Property
Public Property Let X2Marco(vNewValue As Single)
    miX2Marco = vNewValue
End Property

Public Property Get Y2Marco() As Single
    Y2Marco = miY2Marco
End Property
Public Property Let Y2Marco(vNewValue As Single)
    miY2Marco = vNewValue
End Property

Public Function LocalizaGrafico(ByVal X, ByVal Y) As Boolean
    'localizar la situacion
    'localizar la coordenada X
    If X >= Me.X1Marco Then
        If X <= Me.X2Marco Then
        'localizar la coordenada Y
            If Y >= Me.Y1Marco Then
                If Y <= Me.Y2Marco Then
                    LocalizaGrafico = True
                End If
            End If
        End If
    End If

End Function
Private Sub CalculaConstantes()
    'eje x
    A = (miX2Eje - miX1Eje) / (Me.X2Etiqueta - Me.X1Etiqueta)
    B = miX1Eje
    'eje y
    If (Me.Y2Etiqueta - Me.Y1Etiqueta) <> 0 Then
        c = (miY2Eje - miY1Eje) / (Me.Y2Etiqueta - Me.Y1Etiqueta)
    Else
        c = 0  'no debe darse nunca este caso
    End If
    D = miY1Eje
End Sub

Public Function Xc_a_X(ByVal X)
    Xc_a_X = (X - B) / A + Me.X1Etiqueta
End Function

Public Function Yc_a_Y(ByVal Y) As Double
   Yc_a_Y = (Y - D) / c + Me.Y1Etiqueta
End Function


Public Sub PintarLinea(ByVal X1 As Single, ByVal Y1 As Single, ByVal X2 As Single, ByVal Y2 As Single, ColorLinea As Long, Anchura As Single)
    miContenedor.DrawWidth = Anchura 'anchura de la linea
    If TypeName(Contenedor) <> "Contenedor" Then

         miContenedor.Line (X1, Y1)-(X2, Y2), ColorLinea
    Else
          Call miContenedor.Line_(X1, Y1, X2, Y2, ColorLinea)
    End If
    
    
    
    
End Sub

Public Property Get AnchuraMarco() As Single
    AnchuraMarco = miAnchuraMarco
End Property
Public Property Let AnchuraMarco(vNewValue As Single)
    miAnchuraMarco = vNewValue
End Property

Public Property Get AnchuraEjes() As Single
    AnchuraEjes = miAnchuraEjes
End Property
Public Property Let AnchuraEjes(vNewValue As Single)
    miAnchuraEjes = vNewValue
End Property

Public Property Get AnchuraDatos() As Single
    AnchuraDatos = miAnchuraDatos
End Property
Public Property Let AnchuraDatos(vNewValue As Single)
    miAnchuraDatos = vNewValue
End Property

Public Property Get AnchuraRejilla() As Single
    AnchuraRejilla = miAnchuraRejilla
End Property
Public Property Let AnchuraRejilla(vNewValue As Single)
    miAnchuraRejilla = vNewValue
End Property

Public Sub PintarRectangulo(ByVal X1 As Single, ByVal Y1 As Single, ByVal X2 As Single, ByVal Y2 As Single, ColorLinea As Long, Anchura As Single)
    miContenedor.DrawWidth = Anchura 'anchura de la linea
 '  Call miContenedor.Line_(X1, Y1, X2, Y2, ColorLinea, "BF")
    
    If TypeName(Contenedor) <> "Contenedor" Then
        If Me.TipoGrafico = "BARRAS H" Then
            miContenedor.Line (X1, Y1)-(X2, Y2), ColorLinea, B
        Else
            miContenedor.Line (X1, Y1)-(X2, Y2), ColorLinea, BF
        End If
    Else
        If Me.TipoGrafico = "BARRAS H" Then
            Call miContenedor.Line_(X1, Y1, X2, Y2, ColorLinea, "B")
        Else
            Call miContenedor.Line_(X1, Y1, X2, Y2, ColorLinea, "BF")
        End If
    End If
    
End Sub

Public Property Get ColorEjes() As Long
    ColorEjes = miColorEjes
End Property
Public Property Let ColorEjes(vNewValue As Long)
    miColorEjes = vNewValue
End Property

Public Property Get ColorDatos() As Long
    ColorDatos = miColorDatos
End Property
Public Property Let ColorDatos(vNewValue As Long)
    miColorDatos = vNewValue
End Property

Public Property Get ColorRejilla() As Long
    ColorRejilla = miColorRejilla
End Property
Public Property Let ColorRejilla(vNewValue As Long)
    miColorRejilla = vNewValue
End Property

Private Sub PintarRejilla(ByVal X1 As Single, ByVal Y1 As Single, ByVal X2 As Single, ByVal Y2 As Single, ColorLinea As Long, Anchura As Single, tipo As String)
    Me.Contenedor.DrawStyle = 2 'lineas de puntos para la rejilla
    If Me.AnchuraRejilla = 0 Then
        Me.AnchuraRejilla = 1
    End If
    'para el tipo de rejilla horizotal o vertical
    If tipo = "V" Then
        'la rejilla es vertical
        If X1 <> miX1Eje Then 'no pintar sobre el eje Y
            Call PintarLinea(X1, Y1, X1, miY2Eje, Me.ColorRejilla, Me.AnchuraRejilla)
        End If
    End If
    If tipo = "H" Then
        'la rejilla es horizontal
        If Y1 <> miY1Eje Then 'no pintar sobre el eje X
            Call PintarLinea(X1, Y1, miX2Eje, Y1, Me.ColorRejilla, Me.AnchuraRejilla)
        End If
    End If
    Me.Contenedor.DrawStyle = 0 'retoma el valor por defecto
End Sub

Private Sub PintarSegmentosEtiquetas(ByVal Eje As String, ByVal Etiqueta As Variant, ByVal Posicion As Single, ByVal DistanciaEntreSegmentos As Double, ByVal NumeroEtiquetas As Variant, ByVal EtiquetaAnterior As String)
'procedimiento para calcular la posicion de los segmentos y de la etiquetas
'tambien envia controla el proceso de escribir la rejilla


Dim X1, Y1, X2, Y2 As Single
Dim Hora, Dia, DiaAnterior As Variant


If Eje = "X" Then
        'posicion de los segmentos
        X1 = miX1Eje + Posicion * DistanciaEntreSegmentos
        Y1 = miY1Eje
        X2 = X1
        
    If IsDate(Etiqueta) Then
        'fechas
        Hora = Format(Etiqueta, "hh:mm")
        Dia = Format(Etiqueta, "dd/mm/yy")
        Y2 = Y1 + 0.5 * miContenedor.TextHeight(Dia)
        
        'posicion de las etiquetas -HORA-
        Me.Contenedor.CurrentX = X1 - 0.5 * miContenedor.TextWidth(Hora)
        Me.Contenedor.CurrentY = Y2
        
        'pintar la etiqueta -hora-
        If miEtiquetasPintables = True Then
                    If TypeName(Contenedor) <> "Contenedor" Then

                        Me.Contenedor.Print Hora
                    Else
                        Me.Contenedor.Print_ Hora
                    End If
        End If
        
        'posicion de la etiqueta -dia-
        DiaAnterior = Format(EtiquetaAnterior, "dd/mm/yy")
        Me.Contenedor.CurrentX = X1 - 0.5 * miContenedor.TextWidth(Dia)
        
        'comprobacion del dia
        If (Dia <> DiaAnterior) And (Posicion + 1 < NumeroEtiquetas + 0.5) Then
            If miEtiquetasPintables = True Then
                    If TypeName(Contenedor) <> "Contenedor" Then

                        Me.Contenedor.Print Dia
                    Else
                        Me.Contenedor.Print_ Dia
                    End If
            End If
        End If
        If Posicion + 1 >= NumeroEtiquetas + 0.5 Then
            If miEtiquetasPintables = True Then
                    If TypeName(Contenedor) <> "Contenedor" Then

                        Me.Contenedor.Print Dia
                    Else
                        Me.Contenedor.Print_ Dia
                    End If
            End If
        End If
               
    Else
        'numeros
        Y2 = Y1 + 0.5 * miContenedor.TextHeight("0")
        
        'posicion de las etiquetas -numeros-
        Me.Contenedor.CurrentX = X1 - 0.5 * miContenedor.TextWidth(Etiqueta)
        Me.Contenedor.CurrentY = Y2

        'pintar las etiquetas
        If miEtiquetasPintables = True Then
                    If TypeName(Contenedor) <> "Contenedor" Then

                        Me.Contenedor.Print Etiqueta
                    Else
                        Me.Contenedor.Print_ Etiqueta
                    End If
        End If
                
    End If
        
        
End If
    
'para el caso del eje Y
If Eje = "Y" Then

    'posicion de los segmentos
    X1 = miX1Eje
    Y1 = miY1Eje - Posicion * DistanciaEntreSegmentos
    X2 = X1 - Me.Contenedor.TextWidth("0")
    Y2 = Y1
   
   
   'calcular la posicion de las etiquetas
   Me.Contenedor.CurrentX = X2 - (Me.Contenedor.TextWidth(Etiqueta) + 0.5 * Me.Contenedor.TextWidth("0"))
   Me.Contenedor.CurrentY = Y1 - 0.5 * Me.Contenedor.TextHeight(Etiqueta)
   
    'pintar las Etiquetas
    If miEtiquetasPintables = True Then
                    If TypeName(Contenedor) <> "Contenedor" Then

                        Me.Contenedor.Print Etiqueta
                    Else
                        Me.Contenedor.Print_ Etiqueta
                    End If
    End If
   

End If

'Pintar los segmentos
'If miSegmentosPintables = True Then
    Call PintarLinea(X1, Y1, X2, Y2, Me.ColorEjes, Me.AnchuraEjes)
'End If

'rejilla
If Me.Rejilla = True Then
    'llamada al procedimiento de pintar rejilla -HORIZONTAL y VERTICAL-
    Call PintarRejilla(X1, Y1, X1, miY2Eje, Me.ColorRejilla, Me.AnchuraRejilla, "V")
    Call PintarRejilla(X1, Y1, miX2Eje, Y1, Me.ColorRejilla, Me.AnchuraRejilla, "H")
End If



End Sub

Private Sub GraficoPuntos()
Dim X1, Y1 As Single
Dim X2, Y2 As Single
Dim i As Integer
Dim TamañoPunto As Single


Dim ScaleModeContenedor As Variant
Dim FillStyleContenedor As Variant


'Guardar las propiedades del contenedor
ScaleModeContenedor = Me.Contenedor.ScaleMode
FillStyleContenedor = Me.Contenedor.FillStyle

'se elige el modo usuario del contenedor
'
Me.Contenedor.ScaleMode = 0
TamañoPunto = 2

'se puede establecer si el punto es solido
Me.Contenedor.FillStyle = 0
For i = 1 To Me.NumeroDatos
    
    Me.Dato = i
    'para los datos en fallo
    'puede sobrar para los puntos
    Do While Me.YDato = Me.ValorInvalido
        DatoNulo (i)
        i = i + 1
        If i > Me.NumeroDatos Then Exit Sub
        Me.Dato = i
    Loop
   
    'calcular las coordenadas
    Me.Dato = i
    X1 = X_a_Xc(Me.XDato)
    Y1 = Y_a_Yc(Me.YDato)
           
    'pintar el punto
    If TypeName(Contenedor) <> "Contenedor" Then

        miContenedor.Circle (X1, Y1), TamañoPunto, Me.ColorDatos
    Else
    
'        X1 = CSng(X1)
'        Y1 = CSng(Y1)
'        TamañoPunto = CSng(TamañoPunto)
        
            
        Call miContenedor.Circle_(X1, Y1, TamañoPunto, Me.ColorDatos)
    End If
    
    
    
    
Next i


'restablecer las propiedades del contenedor
Me.Contenedor.ScaleMode = ScaleModeContenedor
Me.Contenedor.FillStyle = FillStyleContenedor
End Sub

Private Sub EtiquetasX()
Dim EtiquetasVisibles As Variant
Dim TamañoEtiqueta As Double
Dim NumeroEtiquetas As Variant
Dim EtiquetaInicial, EtiquetaFinal As Variant
Dim i As Long
Dim Etiqueta, EtiquetaAnterior As String
Dim DistanciaEntreSegmentos As Double
Dim X1, X2, Y1, Y2 As Single
Dim TiempoEntreDatos As Double 'long ' en minutos
Dim Intervalo As Double
Dim PrimerDato As Variant
Dim UltimoDato As Variant
Dim aux As Variant
Dim PasoEtiqueta As Variant
Dim Formato As String
Dim n As Single


'no se pintan las etiquetas
If miSegmentosPintables = False Then
    'eje x
    If Me.X1Etiqueta = -1E+38 Then  'la etiqueta X INICIAL no esta predeterminada
        Me.Dato = 1
        Me.X1Etiqueta = Me.XDato ' - Error_T
    End If
    If Me.X2Etiqueta = 1E+38 Then 'etiqueta X FINAL no esta predeterminada
        Me.Dato = Me.NumeroDatos
        Me.X2Etiqueta = Me.XDato ' + Error_T
    End If
    Exit Sub
        
End If

'caso especial: si el número de daros es de solo uno
'If Me.NumeroDatos = 1 Then
'    If Me.TipoGrafico = "LINEAS" Or Me.TipoGrafico = "PUNTOS" Then
    




'para las etiquetas del eje X
Me.Dato = 1
If IsDate(XDato) Then
    'las etiquetas son fechas
    EtiquetasVisibles = 1
    If Me.X1Etiqueta = -1E+38 Then
        'la etiqueta no esta predeterminada
        Me.Dato = 1
        Me.FechaInicial = Me.XDato
        Me.X1Etiqueta = Me.XDato 'es la primera etiqueta
    Else
        Me.FechaInicial = Me.X1Etiqueta 'la etiqueta esta elegida desde fuera
    End If
    If Me.X2Etiqueta = 1E+38 Then
        'la etiqueta no esta elegida
        Me.Dato = Me.NumeroDatos
        Me.FechaFinal = Me.XDato
        Me.X2Etiqueta = Me.XDato 'es la ultima etiqueta
    Else
        Me.FechaFinal = Me.X2Etiqueta 'la etiqueta esta elegida desde fuera
     End If
    
    '
    
    'precalculo
    'el tamaño de tiempo puede depender o no de las etiquetas
    Me.Dato = 1
    PrimerDato = Me.XDato
    Me.Dato = NumeroDatos
    UltimoDato = Me.XDato
    If Me.NumeroDatos = 1 Then
        TiempoEntreDatos = miIntervaloTemporal
    Else
        TiempoEntreDatos = DateDiff("n", PrimerDato, UltimoDato) / (Me.NumeroDatos - 1) ' CLng(DateDiff("n", PrimerDato, UltimoDato) / (Me.NumeroDatos - 1))
    End If
    NumeroEtiquetas = CInt(DateDiff("n", Me.FechaInicial, Me.FechaFinal) / TiempoEntreDatos)
    'TamañoEtiqueta = Me.Contenedor.TextWidth(Format(Me.FechaFinal, "dd/mm/yy"))
    TamañoEtiqueta = Me.Contenedor.TextWidth(Format("31/12/99", "dd/mm/yy"))
    
    'tanteo previo
    n = (miX2Eje - miX1Eje) / TamañoEtiqueta
    
    'etiquetas visibles
    If n / NumeroEtiquetas > 1 Then
        'entran todas las etiquetas
    Else
        'no entran todas las etiquetas
        EtiquetasVisibles = Int(NumeroEtiquetas / n) + 1
    End If
        
    'al menos en el eje debe haber dos etiquetas
    If NumeroEtiquetas = 0 Then
        NumeroEtiquetas = 1
    End If
        
    'parametros de impresion
    NumeroEtiquetas = NumeroEtiquetas / EtiquetasVisibles
    TiempoEntreDatos = TiempoEntreDatos * EtiquetasVisibles
    DistanciaEntreSegmentos = (miX2Eje - miX1Eje) / NumeroEtiquetas
    EtiquetaAnterior = ""
    
    'proceso del escritura del eje X
    For i = 0 To NumeroEtiquetas
        'las etiquetas pueden ser o no regulares en sus distancias
        If i + 1 > NumeroEtiquetas Then
            'es la ultima o es una distancia irregular
            Etiqueta = Format(DateAdd("n", NumeroEtiquetas * TiempoEntreDatos, Me.FechaInicial), "hh:mm dd/mm/yy")
            Call PintarSegmentosEtiquetas("X", Etiqueta, NumeroEtiquetas, DistanciaEntreSegmentos, NumeroEtiquetas, EtiquetaAnterior)
        Else
            'las etiquetas son regulares
            Etiqueta = Format(DateAdd("n", i * TiempoEntreDatos, Me.FechaInicial), "hh:mm dd/mm/yy")
            Call PintarSegmentosEtiquetas("X", Etiqueta, i, DistanciaEntreSegmentos, NumeroEtiquetas, EtiquetaAnterior)
        End If
             
        'necesaria para pintar la etiqueta del dia
        EtiquetaAnterior = Etiqueta
               
    Next i
     
Else
    'son numeros
    EtiquetasVisibles = 1
    'las etiquetas extremas pueden estar elegidas con anterioridad
    If Me.X1Etiqueta = -1E+38 Then  'condicion para la eleccion de etiquetas
        If miMinimoX - Int(miMinimoX) > 0 Then 'etiquetas no predeterminadas
            EtiquetaInicial = Int(miMinimoX)
        Else
            EtiquetaInicial = miMinimoX
        End If
        Me.X1Etiqueta = EtiquetaInicial
    Else
        If Me.X1Etiqueta - Int(Me.X1Etiqueta) > 0 Then 'etiquetas predeterminadas
            EtiquetaInicial = Int(Me.X1Etiqueta)  'no es un valor entero
        Else
            EtiquetaInicial = Me.X1Etiqueta  'si es un valor entero
        End If
    End If
    'para el caso de la etiqueta final del eje
    If Me.X2Etiqueta = 1E+38 Then
        If miMaximoX - Int(miMaximoX) > 0 Then  'etiqueta no predeterminada
            EtiquetaFinal = Int(miMaximoX) + 1  'el valor siguiente al entero maximo
        Else
            EtiquetaFinal = miMaximoX 'el maximo es un entero
        End If
        Me.X2Etiqueta = EtiquetaFinal
    Else
        If Me.X2Etiqueta - Int(Me.X2Etiqueta) > 0 Then  'la etiqueta esta predeterminada
            EtiquetaFinal = Int(Me.X2Etiqueta) + 1  'valor siguiente al entero maximo
        Else
            EtiquetaFinal = Me.X2Etiqueta
        End If
    End If
     
    'precalculo de las etiquetas necesarias
    NumeroEtiquetas = EtiquetaFinal - EtiquetaInicial 'todas las etiquetas entre los valores
    TamañoEtiqueta = 1.15 * Me.Contenedor.TextWidth(EtiquetaFinal) 'separacion entre etiquetas
    
    'tanteo previo
    n = (miX2Eje - miX1Eje) / TamañoEtiqueta
    
    'etiquetas visibles
    If n / NumeroEtiquetas > 1 Then
        'entran todas las etiquetas
    Else
        'no entran todas las etiquetas
        EtiquetasVisibles = Int(NumeroEtiquetas / n) + 1
    End If
    
    If NumeroEtiquetas = 0 Then
        NumeroEtiquetas = 1
    End If
    
    'redondear las etiquetas a valores multiplos de 5
    If EtiquetasVisibles > 5 Then
        EtiquetasVisibles = (Int(EtiquetasVisibles / 5) + 1) * 5
    End If
    
    'Numero de etiquetas
    NumeroEtiquetas = NumeroEtiquetas / EtiquetasVisibles
            
    'calcular el tamaño de cada intervalo entre etiquetas
    DistanciaEntreSegmentos = (miX2Eje - miX1Eje) / NumeroEtiquetas
    
    'bucle para la impresión de segmentos y etiquetas
    For i = 0 To NumeroEtiquetas
        If i > NumeroEtiquetas Then Exit For
            
        If i + 1 > NumeroEtiquetas Then
            'es la ultima o es una distancia irregular
            Etiqueta = Format(NumeroEtiquetas * EtiquetasVisibles + (EtiquetaInicial), miFormato)
            Call PintarSegmentosEtiquetas("X", Etiqueta, NumeroEtiquetas, DistanciaEntreSegmentos, NumeroEtiquetas, EtiquetaAnterior)
        Else
            'las etiquetas son regulares
            Etiqueta = Format(i * EtiquetasVisibles + (EtiquetaInicial), miFormato)
            Call PintarSegmentosEtiquetas("X", Etiqueta, i, DistanciaEntreSegmentos, NumeroEtiquetas, EtiquetaAnterior)
        End If
            
    Next i

End If


End Sub

Private Sub EtiquetasY()
Dim EtiquetasVisibles As Variant
Dim TamañoEtiqueta As Double
Dim NumeroEtiquetas As Variant
Dim EtiquetaInicial, EtiquetaFinal As Variant
Dim i As Long
Dim Etiqueta, EtiquetaAnterior As String
Dim DistanciaEntreSegmentos As Double
Dim X1, X2, Y1, Y2 As Single
Dim TiempoEntreDatos As Long ' en minutos
Dim Intervalo As Double
Dim PrimerDato As Variant
Dim UltimoDato As Variant
Dim aux As Variant
Dim PasoEtiqueta As Variant
Dim Formato As String
Dim n As Single


'no se pintan las etiquetas
    
If miSegmentosPintables = False Then
    'eje y
    If Me.Y1Etiqueta = -1E+38 Then 'la etiqueta Y INICIAL no esta predeterminada
        Me.Y1Etiqueta = miMinimoY
    End If
    If Me.Y2Etiqueta = 1E+38 Then 'etiqueta X FINAL no esta predeterminada
        Me.Y2Etiqueta = miMaximoY
    End If
    If (Me.Y1Etiqueta = Me.Y2Etiqueta) And Me.Y1Etiqueta <> 0 Then 'los valores de estiquetas Y son iguales
        Me.Y1Etiqueta = Me.Y1Etiqueta - 0.1
        Me.Y2Etiqueta = Me.Y2Etiqueta + 0.1
    End If
    Exit Sub
        
End If


    EtiquetasVisibles = 1
    If Me.Y1Etiqueta = -1E+38 Then
        EtiquetaInicial = miMinimoY
        Me.Y1Etiqueta = miMinimoY
    Else
        EtiquetaInicial = Me.Y1Etiqueta
    End If
    If Me.Y2Etiqueta = 1E+38 Then
        EtiquetaFinal = miMaximoY
        Me.Y2Etiqueta = miMaximoY
    Else
        EtiquetaFinal = Me.Y2Etiqueta
    End If
    
    'precalculo del eje Y
    NumeroEtiquetas = EtiquetaFinal - EtiquetaInicial
    TamañoEtiqueta = 1.25 * Me.Contenedor.TextHeight(EtiquetaFinal) 'altura
    'sin control de division
    Select Case NumeroEtiquetas
    Case Is = 0
        'el minimo y el maximo coinciden
        'tres etiquetas
        
        NumeroEtiquetas = 2
        EtiquetasVisibles = 1
        
        If EtiquetaFinal > 0 Then 'son distintas de cero y  positivas
            EtiquetaInicial = EtiquetaInicial - 0.1
            If EtiquetaInicial < 0 Then EtiquetaInicial = 0
            EtiquetaFinal = EtiquetaFinal + 0.1
        ElseIf EtiquetaFinal <= 0 Then '  negativas o nulas
            EtiquetaInicial = 0#
            EtiquetaFinal = 1#
        End If
        
        Me.Y1Etiqueta = EtiquetaInicial
        Me.Y2Etiqueta = EtiquetaFinal
        
         'tanteo previo
        n = (miX2Eje - miX1Eje) / TamañoEtiqueta

        'etiquetas visibles
        If n / NumeroEtiquetas > 1 Then
            'entran todas las etiquetas
        Else
            'no entran todas las etiquetas
            EtiquetasVisibles = Int(NumeroEtiquetas / n) + 1
        End If
        'control de division
        If NumeroEtiquetas = 0 Then
            NumeroEtiquetas = 1
        End If
        
        PasoEtiqueta = (EtiquetaFinal - EtiquetaInicial) / NumeroEtiquetas
        Formato = "#0.00"
    
    Case Is <= 1
        'valores muy proximos menores de uno
        'dos etiquetas
        
        EtiquetasVisibles = 1
        NumeroEtiquetas = Int((miY1Eje - miY2Eje) / TamañoEtiqueta)
        'control de division
        If NumeroEtiquetas = 0 Then
            NumeroEtiquetas = 1
        End If
        If Me.TipoGrafico = "BARRAS" And miMinimoY >= EtiquetaInicial Then
            EtiquetaFinal = EtiquetaFinal + 0.2
        Else
            EtiquetaInicial = EtiquetaInicial - 0.1
            EtiquetaFinal = EtiquetaFinal + 0.1
        End If
        If EtiquetaInicial < 0 Then EtiquetaInicial = 0
         'tanteo previo
        n = (miX2Eje - miX1Eje) / TamañoEtiqueta
    
        'etiquetas visibles
        If n / NumeroEtiquetas > 1 Then
            'entran todas las etiquetas
        Else
            'no entran todas las etiquetas
            EtiquetasVisibles = Int(NumeroEtiquetas / n) + 1
        End If
        'control de division
        If NumeroEtiquetas = 0 Then
            NumeroEtiquetas = 1
        End If
    
        PasoEtiqueta = (EtiquetaFinal - EtiquetaInicial) / NumeroEtiquetas
        Formato = "#0.00"
    Case Else
        'mayor que la unidad
        EtiquetasVisibles = 1
        NumeroEtiquetas = Int((miY1Eje - miY2Eje) / TamañoEtiqueta)
        
        'control de division
        If NumeroEtiquetas = 0 Then
            NumeroEtiquetas = 1
        End If
        
        'tanteo previo
        n = (miX2Eje - miX1Eje) / TamañoEtiqueta
    
        'etiquetas visibles
        If n / NumeroEtiquetas > 1 Then
        'entran todas las etiquetas
        Else
        'no entran todas las etiquetas
        EtiquetasVisibles = Int(NumeroEtiquetas / n) + 1
        End If
 
        If EtiquetasVisibles > 5 Then
            EtiquetasVisibles = (Int(EtiquetasVisibles / 5) + 1) * 5
        End If
  
        'siempre es regular la separacion
        PasoEtiqueta = (EtiquetaFinal - EtiquetaInicial) / NumeroEtiquetas
        Formato = "#0.0"
    End Select
 
    'calcular el tamaño de la segmentacion
    DistanciaEntreSegmentos = (miY1Eje - miY2Eje) / NumeroEtiquetas
    
    'bucle para la impresion del segmento y de etiquetas
    For i = 0 To NumeroEtiquetas
                
        Etiqueta = Format(i * PasoEtiqueta + EtiquetaInicial, Formato)
        'enviar al proceso pintarsegmentos y etiquetas
        Call PintarSegmentosEtiquetas("Y", Etiqueta, i, DistanciaEntreSegmentos, NumeroEtiquetas, EtiquetaAnterior)
    Next i

End Sub

Public Property Get IntervaloTemporal() As Integer   'en minutos
IntervaloTemporal = miIntervaloTemporal
End Property

Public Property Let IntervaloTemporal(v As Integer)
miIntervaloTemporal = v
End Property

Public Sub EstudiaEtiquetasY()

End Sub

Public Sub Zoom(Factor As Double, XCentro As Double, YCentro As Double)
  Dim ancho As Double
  Dim alto As Double
    
    ancho = (X1Etiqueta - X2Etiqueta) * Factor
    alto = (Y1Etiqueta - Y2Etiqueta) * Factor
    
    Y1Etiqueta = YCentro - alto / 2
    Y2Etiqueta = YCentro + alto / 2
    X1Etiqueta = XCentro - ancho / 2
    X2Etiqueta = XCentro + ancho / 2

End Sub

