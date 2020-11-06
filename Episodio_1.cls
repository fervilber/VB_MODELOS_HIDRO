VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Episodio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'clase que gestiona todo lo relativo a un episodio temporal: letra de archivos, escritura,
'y realiza cálculos sencillos con los datos de las series: valor acumulado, fechas extremas
'del episodio, etc.

Option Explicit  'Sentencia que obliga a declarar todas las variables
'Las variables que se ddeclaran son variables internas de las propiedades de la clase (prefijo "mi")

'Para verificar si se ha aceptado el episodio
Dim misComentarios As String

'Para los episodios

Dim miNombreFichero As String
Dim midirectorio As String
Dim miFecha As String
Dim miFechaLarga As String 'Guarda la fecha "traducida" a partir del nombre del fichero
Dim miNumeroIntervalos As Long
Dim miIntervalo As Long
Dim miLongitudIntervalos As Long
Dim miMaximaPrecipDeEpisodio As Double
Dim miMinimaPrecipDeEpisodio As Double
Dim miMaximoCaudalEmbalse As Double
Dim miMaximoCaudalAforo As Double
Dim miSeries As Long 'Almacena el número de series leidas en el fichero
Dim miSerie As Long
Dim miMsgLectura As String
Dim miOrdenacionSerie As String
Dim miSeleccionSerie As String
Dim miErrorReconocimientoFecha As String
Dim miErrorReconocimientoArchivo As String
Dim miErrorFicheroLeido As String
Dim miErrorFaltanDatos As String
Dim miErrorSobranDatos As String
Dim miValorInvalido As Double
Dim miCodigosAnsi As Boolean

' Para las series: matrices dinámicas que almacenan los valores leidos
Dim miFechaIntervalo() As Date
Dim miNombreSerie() As String
Dim miTipoSerie() As String
Dim miValorIntervaloSerie() As Double
Dim miValorMaximoSerie() As Double
Dim miValorMinimoSerie() As Double
Dim miValorMaximoEnIntervalo() As Double
Dim SiHeMaxIntervalo As Boolean
Dim miCantidadInvalidosSerie() As Double
Dim miXSerie() As Double
Dim miYSerie() As Double
Dim miZSerie() As Double
Dim miPSerie() As Double
' Para las selecciones y ordenaciones:
Dim MSeriesSeleccionadas() As Long 'matriz que almacena qué series han sido seleccionadas desde SelSeries
Dim MatrizAcumulados() As Double
Dim SiHeAcumulado As Boolean  'indica si se ha hecho el cálculo de serie acumulada
Dim miSeriesSel As Long
Dim miSerieSel As Long
Public Enum Orden_Y_Sel   'indica el tipo de ordenacion de las series del episodio
    OrdenDefecto
    OrdenAlfa
    OrdenNS
    OrdenEO
    OrdenTipos
End Enum
Public Enum Tipo_Seleccion   'indica el tipo de seleccion de las series del episodio
    verformulario
    SelPluviometros
    SelEmbalses
    SelAforos
    SelPluviometrosEmbalses
    SelPluviometrosAforos
    SelEmbalsesAforos
    SelTodo
    AnularSel
End Enum

Public Enum LecturasEpi   'indica si se ha leido correctamente el fichero
    EpiFichLeidoOK
    EpiFichLeidoerr
End Enum
Public Event FicheroLeido(ByVal res As LecturasEpi)  'evento que salta cuando se ha terminado de leer el fichero

'Dim Lecturas As New modHidro
Dim vector() As Double


Public Sub Acumula(Optional IntIni, Optional IntFin, Optional SI, Optional sf)
Dim i  As Integer
Dim j As Integer
Dim Suma As Double
Dim invalidos As Long

'calcula el acumulado de cada serie despues de haber leido el fichero
'
If IsMissing(IntIni) Then IntIni = 1
If IsMissing(IntFin) Then IntFin = Me.NumeroIntervalos
If IsMissing(SI) Then SI = 1
If IsMissing(sf) Then sf = Me.Series

For i = SI To sf
    Suma = 0
    invalidos = 0
    Me.Serie = i
    For j = IntIni To IntFin
        Me.Intervalo = j
        If Me.Valor = Me.ValorInvalido Then invalidos = invalidos + 1
        If Me.Valor <> Me.ValorInvalido Then
             Suma = Suma + Me.Valor
        End If
    Next j
    If invalidos = Me.NumeroIntervalos Then Suma = Me.ValorInvalido
    MatrizAcumulados(i) = Suma
Next i

SiHeAcumulado = True

End Sub


Public Function SiExisteSerie(tipo As String, NOMBRE As String)
'comprueba si la serie buscada esta en el episodio
    Dim S As Integer
    For S = 1 To Me.Series
        Me.Serie = S
        If UCase$(Trim$(Me.NombreSerie)) = UCase$(Trim$(NOMBRE)) Then
            If UCase$(Trim$(Me.TipoSerie)) = UCase$(Trim$(tipo)) Then
                SiExisteSerie = S
                Exit Function
            End If
        End If
    Next S
    SiExisteSerie = -1
    
    
End Function
Public Property Get SeleccionSerie() As String  'mensaje que indica que tipo de series se han seleccionado
    SeleccionSerie = miSeleccionSerie
End Property
Public Property Let SeleccionSerie(v As String)
    miSeleccionSerie = v
End Property

Public Property Get OrdenacionSerie() As String  'mensaje que indica como estan ordenadas las series seleccionadas
    OrdenacionSerie = miOrdenacionSerie
End Property
Public Property Let OrdenacionSerie(v As String)
    miOrdenacionSerie = v
End Property

Public Property Get Intervalo() As Long  'Identifica el índice del intervalo leido
    Intervalo = miIntervalo
End Property
Public Property Let Intervalo(vNewValue As Long) 'Identifica el índice del intervalo leido
    miIntervalo = vNewValue
End Property

Public Property Get Valor() As Double  'Para el valor del intervalo
   Valor = miValorIntervaloSerie(Me.Intervalo, Me.Serie)
End Property
Public Property Let Valor(vNewValue As Double) 'Para el valor del intervalo
    miValorIntervaloSerie(Me.Intervalo, Me.Serie) = vNewValue
End Property

Public Property Get ValorMinimo() As Double  'Para el valor mínimo
    ValorMinimo = miValorMinimoSerie(Me.Serie)
End Property
Public Property Let ValorMinimo(vNewValue As Double)  'Para el valor mínimo
    miValorMinimoSerie(Me.Serie) = vNewValue
End Property
Public Property Get ValorMaximo() As Double 'Para el valor máximo
    ValorMaximo = miValorMaximoSerie(Me.Serie)
End Property
Public Property Let ValorMaximo(vNewValue As Double) 'Para el valor máximo
    miValorMaximoSerie(Me.Serie) = vNewValue
End Property

Public Property Get MaximaPrecipDeEpisodio() As Double 'valor maximo del episodio
    MaximaPrecipDeEpisodio = miMaximaPrecipDeEpisodio
End Property
Public Property Let MaximaPrecipDeEpisodio(ByVal vNewValue As Double)
    miMaximaPrecipDeEpisodio = vNewValue
End Property
Public Property Get MinimaPrecipDeEpisodio() As Double 'valor minimo del episodio
    MinimaPrecipDeEpisodio = miMinimaPrecipDeEpisodio
End Property
Public Property Let MinimaPrecipDeEpisodio(ByVal vNewValue As Double)
    miMinimaPrecipDeEpisodio = vNewValue
End Property

Public Property Get CantidadInvalidosSerie() As Double 'Identifica el número de valores inválidos que hay en cada serie que se lee
    CantidadInvalidosSerie = miCantidadInvalidosSerie(Me.Serie)
End Property
Public Property Let CantidadInvalidosSerie(vNewValue As Double) 'Identifica el número de valores inválidos que hay en cada serie que se lee
    miCantidadInvalidosSerie(Me.Serie) = vNewValue
End Property

Public Property Get ValorInvalido() As Double  'Para el valor inválido que se especifique
    ValorInvalido = miValorInvalido
End Property
Public Property Let ValorInvalido(vNewValue As Double) 'Para el valor inválido que se especifique
    miValorInvalido = vNewValue
End Property

Public Property Get CodigosAnsi() As Boolean 'Identifica si los valores leidos son ANSI o no
    CodigosAnsi = miCodigosAnsi
End Property
Public Property Let CodigosAnsi(vNewValue As Boolean) 'Identifica si los valores leidos son ANSI o no
   miCodigosAnsi = vNewValue
End Property

Public Property Get XSerie() As Double  'coordenada x geográfica de la serie
    XSerie = miXSerie(Me.Serie)
End Property
Public Property Let XSerie(vNewValue As Double)
    miXSerie(Me.Serie) = vNewValue
End Property

Public Property Get YSerie() As Double 'coordenada y geográfica de la serie
    YSerie = miYSerie(Me.Serie)
End Property
Public Property Let YSerie(vNewValue As Double)
    miYSerie(Me.Serie) = vNewValue
End Property

Public Property Get ZSerie() As Double 'coordenada z geográfica de la serie
    ZSerie = miZSerie(Me.Serie)
End Property
Public Property Let ZSerie(vNewValue As Double)
    miZSerie(Me.Serie) = vNewValue
End Property

Public Property Get PSerie() As Double  'parametro
    PSerie = miPSerie(Me.Serie)
End Property

Public Property Let PSerie(vNewValue As Double)
    miPSerie(Me.Serie) = vNewValue
End Property

Public Sub IniciaSeries()
    
'Procedimiento que establece los valores iniciales de las propiedades y variables de la clase
' y los indices de dimensionamiento de las matrices dinámicas de almacenamiento de los valores de las series
    misComentarios = ""
    SiHeAcumulado = False
    SiHeMaxIntervalo = False
    miSeries = 0
    Me.Serie = 0
    Me.LongitudIntervalos = 0
    Me.NumeroIntervalos = 0
    Me.MaximaPrecipDeEpisodio = 0
    Me.MaximoCaudalEmbalse = 0
    Me.MaximoCaudalAforo = 0

    Me.Intervalo = 0
    Me.ValorInvalido = -1
    Me.ErrorReconocimientoArchivo = ""
    Me.ErrorReconocimientoFecha = ""
    Me.ErrorFaltanDatos = ""
    Me.ErrorSobranDatos = ""
    Me.ErrorFicheroLeido = ""
    Me.MsgLectura = ""
    
    'Establece los valores iniciales de los límites inferior y superior de las dimensiones de las matrices dinámicas
    ReDim miTipoSerie(0 To 0)
    ReDim miNombreSerie(0 To 0)
    ReDim miXSerie(0 To 0)
    ReDim miYSerie(0 To 0)
    ReDim miZSerie(0 To 0)
    ReDim miPSerie(0 To 0)
    ReDim miValorIntervaloSerie(0 To 0, 0 To 0)
    ReDim miValorMaximoSerie(0 To 0)
    ReDim miValorMinimoSerie(0 To 0)
    ReDim miCantidadInvalidosSerie(0 To 0)
    ReDim MSeriesSeleccionadas(0 To 0)
    ReDim MatrizAcumulados(0 To 0)
    ReDim miValorMaximoEnIntervalo(0 To 0)
    'Valores iniciales de los datos de las matrices
    miTipoSerie(0) = ""
    miNombreSerie(0) = ""
    miXSerie(0) = 0#
    miYSerie(0) = 0#
    miZSerie(0) = 0#
    miPSerie(0) = 0#
    miValorIntervaloSerie(0, 0) = 0#
    miValorMaximoSerie(0) = 0#
    miCantidadInvalidosSerie(0) = 0#
    miValorMinimoSerie(0) = 0#
    miValorMaximoEnIntervalo(0) = 0#
    MSeriesSeleccionadas(0) = 0
    
    'valores iniciales de los mensajes de formulario
    miOrdenacionSerie = "Por Defecto"
    miSeleccionSerie = "Fichero Completo"
    
   
End Sub

Public Function CalculaFechIni(LIn As Long, NIn As Long, fech As String)
'se calcula la fecha inicial del episodio
    Dim f As Date
    f = CDate(fech)
    CalculaFechIni = DateAdd("n", -1 * LIn * NIn, (f))
End Function
Public Function FechaInicial()
'funcion que devuelve la fecha inicial del episodio utilazando el procedimiento calculafechini
    FechaInicial = CalculaFechIni(Me.LongitudIntervalos, Me.NumeroIntervalos, Me.Fecha)
End Function

Private Function FechaLargaEpisodio(cadena As String) As String
    
'Transforma el nombre del fichero en un mensaje en el que indica la fecha final del intervalo
    
    Dim Año As String
    Dim Mes As String
    Dim Dia As String
    Dim Hora As String
    Dim Minuto As String
    Dim diasemana As String
    Dim cond As Boolean
    
    cond = True 'La variable "cond" se usa para identificar si un caracter del nombre del fichero es numerico o no, y asignar un mensaje de error en su caso
    
    'Analiza el año
    Año = Right$(cadena, 2)
    If Not IsNumeric(Año) Then cond = False
    
    'Analiza el mes
    Mes = Left$(cadena, 2)
    If Not IsNumeric(Mes) Then
        cond = False
    Else
        Select Case Val(Mes)
        Case 1
            Mes = "Enero"
        Case 2
            Mes = "Febrero"
        Case 3
            Mes = "Marzo"
        Case 4
            Mes = "Abril"
        Case 5
            Mes = "Mayo"
        Case 6
            Mes = "Junio"
        Case 7
            Mes = "Julio"
        Case 8
            Mes = "Agosto"
        Case 9
            Mes = "Septiembre"
        Case 10
            Mes = "Octubre"
        Case 11
            Mes = "Noviembre"
        Case 12
            Mes = "Diciembre"
        Case Else
            Mes = "Desconocido"
        End Select
    End If
    
    'Analiza el dia
    Dia = Mid$(cadena, 3, 2)
    If Not IsNumeric(Dia) Then cond = False
    If cond Then
        Select Case Weekday(Dia & "/" & Mes & "/" & Año)
        Case 1
            diasemana = "Domingo"
        Case 2
            diasemana = "Lunes"
        Case 3
            diasemana = "Martes"
        Case 4
            diasemana = "Miercoles"
        Case 5
            diasemana = "Jueves"
        Case 6
            diasemana = "Viernes"
        Case 7
            diasemana = "Sábado"
        End Select
    Else
        diasemana = ""
    End If
    
    'Analiza las horas y los minutos
    Hora = Mid$(cadena, 5, 2)
    Minuto = Mid$(cadena, 7, 2)
        
    'Si todo esta correcto "cond" es verdadero y se establece el mensaje de fecha
    'si no esta correcto, generamos un mensaje de error y se asigna a la propiedad correspondiente
    
    If cond = True Then
        FechaLargaEpisodio = diasemana & " " & Dia & " de " & Mes & " del " & Año & " a las " & Hora & ":" & Minuto
    Else
        Me.ErrorReconocimientoFecha = "Nombre no reconocido"
        FechaLargaEpisodio = Me.ErrorReconocimientoFecha
    End If

End Function

Private Function FechaEpisodio(cadena As String)
    
'Transforma el nombre del fichero en un formato de fecha estándar,
'tipo: DIA/MES/AÑO  HORA:MINUTO
    
    Dim aux As String
    Dim Año As String
    Dim Mes As String
    Dim Dia As String
    Dim Hora As String
    Dim Minuto As String
    Dim condicion As Integer
   
    'Analiza los caracteres del nombre del fichero, para ver si son numéricos
    'y los asigna a la variable correspondiente
    
    Año = Right$(cadena, 2)
    If Not IsNumeric(Año) Then Año = "?"
    Mes = Left$(cadena, 2)
    If Not IsNumeric(Mes) Then Mes = "?"
    Dia = Mid$(cadena, 3, 2)
    If Not IsNumeric(Dia) Then Dia = "?"
    Hora = Mid$(cadena, 5, 2)
    If Not IsNumeric(Hora) Then Hora = "?"
    Minuto = Mid$(cadena, 7, 2)
    If Not IsNumeric(Minuto) Then Minuto = "?"
    
    condicion = IsNumeric(Año) And IsNumeric(Mes) And IsNumeric(Dia)
    condicion = condicion And IsNumeric(Hora) And IsNumeric(Minuto)
    
    If Not condicion Then
        FechaEpisodio = "Desconocida"
        Exit Function
    End If
    
    'Traduce los caracteress leidos a un formato de fecha
    
    aux = DateSerial(Val(Año), Val(Mes), Val(Dia)) + TimeSerial(Val(Hora), Val(Minuto), 0)
    FechaEpisodio = Format(aux, "dd/mm/yy hh:nn")

End Function

Public Function FechaIntervalo()
'se calcula la fecha final del intervalo actual

    Dim k As Integer
    k = Me.Intervalo
    FechaIntervalo = DateAdd("n", Me.LongitudIntervalos * k, CalculaFechIni(Me.LongitudIntervalos, Me.NumeroIntervalos, Me.Fecha))
End Function

Public Function NombreIntervalo(PreExt As String)
     Dim Año As String
     Dim Mes As String
     Dim Dia As String
     Dim Hora As String
     Dim Minuto As String
     Dim Indice As String
    'se crea, a partir de la fecha del intervalo actual, el nombre del episodio que corresponde a ese intervalo
    ' con la nomenclatura  mmddhhss.Xyy
    
     
     Año = Year(Me.FechaIntervalo)
     Mes = Month(Me.FechaIntervalo)
     Dia = Day(Me.FechaIntervalo)
     Hora = Hour(Me.FechaIntervalo)
     Minuto = Minute(Me.FechaIntervalo)
     
     'analizamos los valores devueltos
     'año
     Año = Trim$(Año)
     Año = Mid$(Año, 3, 2)
     
     'Mes
     Mes = Trim$(Mes)
     If CInt(Mes) < 10 Then
        Mes = "0" & Mes
     End If
     
     'Dia
     Dia = Trim$(Dia)
     If CInt(Dia) < 10 Then
        Dia = "0" & Dia
     End If
     
     'hora
     Hora = Trim$(Hora)
     If CInt(Hora) < 10 Then
        Hora = "0" & Hora
     End If
     
     'minutos
     Minuto = Trim$(Minuto)
     If CInt(Minuto) < 10 Then
        Minuto = "0" & Minuto
     End If
    
    NombreIntervalo = Mes & Dia & Hora & Minuto & "." & Trim$(PreExt) & Año
        
End Function

Public Function Fecha()
'Asigna la variable interna "mifecha" a la de salida "fecha",
'validada en el Property Let de NombreFichero mediante la función FechaEpisodio
    Fecha = miFecha
End Function

Public Function FechaLarga()
'Asigna la variable interna "mifechalarga" a la de salida "fechalarga",
'validada en el Property Let de NombreFichero mediante la función FechaLargaEpisodio
    FechaLarga = miFechaLarga
End Function

Public Property Get nombrefichero() As String 'Indica el nombre del fichero de datos
    nombrefichero = miNombreFichero
End Property

Public Property Let nombrefichero(vNewValue As String) 'Indica el nombre del fichero de datos
'Hay que validar el nombre
    miNombreFichero = vNewValue
    miFecha = FechaEpisodio(vNewValue)
    miFechaLarga = FechaLargaEpisodio(vNewValue)
End Property

Public Property Let Directorio(vNewValue As String) 'Identifica el directorio de trabajo
    vNewValue = Trim$(vNewValue)
    If Right$(vNewValue, 1) <> "\" Then
        midirectorio = vNewValue & "\"
    Else
        midirectorio = vNewValue
    End If
End Property
Public Property Get Directorio() As String 'Identifica el directorio de trabajo
    Directorio = midirectorio
End Property

Public Property Get NumeroIntervalos() As Long 'Para el número de intevalos
    NumeroIntervalos = miNumeroIntervalos
End Property
Public Property Let NumeroIntervalos(vNewValue As Long) 'Para el número de intevalos
    miNumeroIntervalos = vNewValue
End Property

Public Property Get LongitudIntervalos() As Long 'Identifica la duración de cada intervalo
     LongitudIntervalos = miLongitudIntervalos
End Property

Public Property Let LongitudIntervalos(vNewValue As Long) 'Identifica la duración de cada intervalo
    miLongitudIntervalos = vNewValue
End Property

Public Property Get Serie() As Long 'Para el indice de la serie en curso
    Serie = miSerie
End Property
Public Property Let Serie(vNewValue As Long) 'Para el indice de la serie en curso
    If vNewValue > Me.Series Then
        miSerie = Me.Series
    ElseIf vNewValue < 1 Then
        miSerie = 1
    Else
        miSerie = vNewValue
    End If
End Property

Public Function Series() As Long 'Para el número de series leidas
    Series = miSeries
End Function

Public Property Get SerieSel() As Long 'Para el indice de la serie en curso
    SerieSel = miSerieSel
End Property
Public Property Let SerieSel(vNewValue As Long) 'Para el indice de la serie en curso
    If vNewValue > Me.SeriesSel Then
        miSerieSel = Me.SeriesSel
    ElseIf vNewValue < 1 Then
        miSerieSel = 1
    Else
        miSerieSel = vNewValue
        Me.Serie = MSeriesSeleccionadas(miSerieSel)
    End If
End Property


Public Property Get SerieAcumulada() As Double  'devuelve el valor acumulado de los intervalos de la serie
If SiHeAcumulado = False Then
    Acumula
    SiHeAcumulado = True
End If

    SerieAcumulada = MatrizAcumulados(Me.Serie)
End Property

Public Function SeriesSel() As Long
    SeriesSel = UBound(MSeriesSeleccionadas) - LBound(MSeriesSeleccionadas) + 1
End Function

Public Property Get TipoSerie() As String 'Para el tipo de serie
    TipoSerie = miTipoSerie(Me.Serie)
End Property
Public Property Let TipoSerie(vNewValue As String) 'Para el tipo de serie
     miTipoSerie(Me.Serie) = vNewValue
End Property
Public Function DescribeTipoSerie() As String
    Select Case UCase$(Me.TipoSerie)
    Case "P"
        DescribeTipoSerie = "  " & "Precipitaciones"
    Case "N"
        DescribeTipoSerie = "  " & "Niveles de embalse"
    Case "S"
        DescribeTipoSerie = "  " & "Caudales de salida de embalse"
    Case "C"
        DescribeTipoSerie = "  " & "Niveles en estación de aforos"
    Case "Q"
        DescribeTipoSerie = "  " & "Caudales en estación de aforos"
    Case "V"
        DescribeTipoSerie = "  " & "Volumen de embalse"
    Case Else
        DescribeTipoSerie = "  " & "Tipo de serie no especificado"
    End Select
End Function
Public Property Get NombreSerie() As String 'Para el nombre de la serie
    NombreSerie = miNombreSerie(Me.Serie)
End Property
Public Property Let NombreSerie(vNewValue As String) 'Para el nombre de la serie
     miNombreSerie(Me.Serie) = vNewValue
End Property

Public Function LeeNumeroYLargoInterv() As Boolean

'Función que lee las líneas iniciales del fichero,almacenando la longitud
'de intervalo y el número de intervalos del fichero (línea G)

On Error GoTo SiError_LeeNumeroYLargoInterv
    Dim nf As Integer
    Dim nfaux As Integer
    Dim caso As Integer
    Dim TipoFunc As Boolean  'Variable que indica si se ha leido correctamente la linea de información general (true) o no
    Dim Linea As String
    Dim cosa As String * 1
    Dim Posicion As Integer
    Dim CadenaCarA As String
    Dim CadenaCarB As String
    Dim msgA As String
    Dim Parametro As String
    Dim NumInter As Long
    Dim LargInter As Long
    
    IniciaSeries
    TipoFunc = True
    nf = FreeFile
    cosa = ""
    Open midirectorio & miNombreFichero For Input As nf
        
        'Con este bucle se saltan las primeras lineas (en blanco o
        'con asterisco (titulos del  fichero):
        Do While Not EOF(nf) And UCase(Left$(Linea, 1)) <> "G"
            Line Input #nf, Linea
            Linea = Trim(Linea)
        Loop
    
    
        If Not EOF(nf) Then
            nfaux = FreeFile  'Fichero auxiliar donde se escribe la linea G preparada para leer
            Open "LineaG.tmp" For Output As #nfaux
            Linea = LTrim$(Mid$(Linea, 2))
            Print #nfaux, Linea
            Close nfaux
            Open "LineaG.tmp" For Input As #nfaux
            
            'A continuación se lee la línea caracter a caracter y los
            'que sean numéricos, se almacenan en las propiedades correspondientes
            
            If Not EOF(nfaux) Then
                Do While Not EOF(nfaux) And IsNumeric(cosa) = False 'Lee hasta que encuentra un valor numérico
                    cosa = Input$(1, nfaux)
                Loop
                Do While Not EOF(nfaux) And IsNumeric(cosa) 'Lee los primeros valores numéricos
                    CadenaCarA = CadenaCarA & cosa
                    cosa = Input$(1, nfaux)
                Loop
                
                Me.NumeroIntervalos = CadenaCarA  'Se asignan a la propiedad
                
                Do While Not EOF(nfaux) And IsNumeric(cosa) = False  'Lee hasta que encuentra un valor numérico
                    cosa = Input$(1, nfaux)
                Loop
                
                Do While Not EOF(nfaux) And IsNumeric(cosa) 'Lee los segundos valores numericos
                    CadenaCarB = CadenaCarB & cosa
                    cosa = Input$(1, nfaux)
                Loop
                
                Me.LongitudIntervalos = CadenaCarB  'Se asignan a la propiedad
            
            Else
                GoTo SiError_LeeNumeroYLargoInterv  'Si se ha acabado el fichero "aux" antes de los esperado
            End If
            Close #nfaux
        Else
            GoTo SiError_LeeNumeroYLargoInterv  'Si se ha acabado el fichero "nf" antes de los esperado
        End If
                   
    Close #nf

    
    'Asigna el valor de tipofunc (indica si la lectura es correcta o no)
    'al valor de salida de la función, que será utilizado en Leefichero:
    
    LeeNumeroYLargoInterv = TipoFunc


Sal_LeeNumeroYLargoInterv:
    Exit Function
    
SiError_LeeNumeroYLargoInterv:
    Me.ErrorFicheroLeido = "No se encontraron las especificaciones generales"
    Me.ErrorReconocimientoArchivo = "Archivo no reconocido"
    Me.NumeroIntervalos = Me.ValorInvalido
    Me.LongitudIntervalos = Me.ValorInvalido
    Close #nf
    Close #nfaux
    TipoFunc = False  'Lectura con error
    LeeNumeroYLargoInterv = TipoFunc
    Resume Sal_LeeNumeroYLargoInterv

End Function

Public Sub LeeFichero()
'Función que lee los datos del archivo y los almacena en las matrices correspondientes
    On Error GoTo SiError_LeeFichero  'Control de errores en la lectura
    Dim nf As Integer
    Dim nftmp As Integer
    Dim caso As Integer
    Dim canalaux As Integer
    Dim Linea As String
    Dim i As Integer
    Dim j As Integer
    Dim Contadat As Integer
    Dim Posicion As Integer
    Dim msgB As String
    Dim aux_S As String
    Dim LecturaDeLineaG As Boolean
    
    
    IniciaSeries   'Se llama al procedimiento de inicialización de valores
    
    
    msgB = ""
    LecturaDeLineaG = True
    nf = FreeFile    'Se asigna un número de canal libre para abrir el fichero
    nftmp = FreeFile
    
    
'PRIMERA APERTURA DE FICHERO NF: PASA DE LAS LINEAS EN BLANCO Y CON ASTERISCO
'Y LEE SOLO LAS QUE EMPIEZAN CON G, DETERMINANDO EL NUMERO DE INTERVALOS Y LA LONGITUD ED LOS MISMOS
    
    
    Me.LeeNumeroYLargoInterv   'Llama a la función que lee la linea de especificaciones generales
    
    'Se asigna la variable booleana devuelta por la función LeeNumeroYLargoInterv
    'y que indica si ha habido o no un error de lectura de fichero
    
    LecturaDeLineaG = LeeNumeroYLargoInterv
    If LecturaDeLineaG = False Then Exit Sub
    
'SEGUNDA APERTURA DE EPISODIOtemp y de nf; ADEMAS, LEE EL TIPO
'PARTE EL FICHERO Y SE QUEDA CON EL NOMBRE,XYZP, Y LOS VALORES DE INTERVALOS
    
    
    nf = FreeFile
    Open midirectorio & miNombreFichero For Input As #nf  'Abre el fichero de episodios
    nftmp = FreeFile
    Open "Episodio.Tmp" For Output As #nftmp   'Crea un fichero temporal donde almacena los datos de las series
    
    Do While Not EOF(nf)
    
        Line Input #nf, Linea
        Linea = Trim$(Linea)  'Quita los espacios en blanco alrededor de la linea
        
        
        If Linea <> "" Then
            
            'Si llega a una linea de datos de las series:
            If Left$(Linea, 1) <> "*" And UCase$(Left$(Linea, 1)) <> "G" Then
                
                miSeries = miSeries + 1    'Asigna el valor inicial al indice de las series, estableciendo un contador que se incrementa cada vez que se lee una linea
                
                If miSeries = 1 Then
                    ReDim miTipoSerie(1 To miSeries)  'Para la primera linea leida
                Else
                    ReDim Preserve miTipoSerie(1 To miSeries)  'Para posteriores lineas,las guarda manteniendo lo anterior
                End If
                
                'Lee y analiza el tipo de la serie, formando un mensaje de error si el tipo no esta especificado
                miTipoSerie(miSeries) = UCase$(Left$(Linea, 1))
                
                Select Case miTipoSerie(miSeries)
                Case "P", "N", "S", "C", "Q", "V" 'Casos válidos de tipo
                   Me.MsgLectura = ""
                Case Else
                    aux_S = "Error al leer la especificacion de Tipo en las series: "
                    If (InStr(1, msgB, aux_S) = 0) Then
                        msgB = msgB + aux_S + Str$(miSeries)  'Si aún no se ha formado el mensaje de error, lo forma
                    Else
                        msgB = msgB & ";" & Str$(miSeries)  'Si esta  formado, añade el número de serie actual a las otras erróneas
                    End If
                End Select
                
                Print #nftmp, Mid$(Linea, 2)  'Escribe en el fichero temporal, la línea leida, eliminando el caracter inicial que indica el tipo de serie
            End If
        End If
    Loop  'Vuelve al inicio del bucle para leer una nueva linea
    
    Close #nftmp
    Close #nf
    
    'Asigna a la propiedad de error de lectura, el mensaje creado,
    'añadiéndolo a otros posibles mensajes anteriores
    '(que puedan provenir de LeeNumeroYLargoInterv):
    Me.ErrorFicheroLeido = Me.ErrorFicheroLeido & "  " & msgB

    
    'Redimensiona las matrices de datos de acuerdo con el numero de series leidas en el bucle anterior
    ReDim miNombreSerie(1 To miSeries)
    ReDim miXSerie(1 To miSeries)
    ReDim miYSerie(1 To miSeries)
    ReDim miZSerie(1 To miSeries)
    ReDim miPSerie(1 To miSeries)
    ReDim miValorIntervaloSerie(1 To Me.NumeroIntervalos, 1 To miSeries)
    ReDim miValorMaximoSerie(1 To miSeries)
    ReDim miValorMinimoSerie(1 To miSeries)
    ReDim miCantidadInvalidosSerie(1 To miSeries)
    ReDim miFechaIntervalo(1 To Me.NumeroIntervalos)
    ReDim MSeriesSeleccionadas(1 To miSeries)
    ReDim MatrizAcumulados(1 To miSeries)

    miMaximaPrecipDeEpisodio = -1E+36
    miMinimaPrecipDeEpisodio = 1E+36
    miSerie = 0
'VAMOS A LEER EL NOMBRE, COORDENADAS Y VALORES DE LA SERIE
    
    nftmp = FreeFile
    Open "Episodio.Tmp" For Input As #nftmp  'Abre el episodio temporal anterior
        
        canalaux = FreeFile   'Establece un canal libre para un nuevo fichero auxiliar a crear
        
    Do While Not EOF(nftmp)
        
        'En este bucle, se lee el fichero temporal linea por linea,
        'guardándolas en un nuevo fichero auxiliar (EpisodioAux) de donde se
        'leen los datos de cada serie y se almacenan en las matrices
        'correspondientes; esto se hace para evitar errores en la lectura
        'ocasionados por el modo de trabajar de las funciones input
        'y lineinput (de los procedimientos Lee y LeeLinea)
        
        'Line Input #
         Line Input #nftmp, Linea
        miSerie = miSerie + 1   'Contador de las series leidas
        
'        Open "EpisodioAux.tmp" For Output As #canalaux  'Crea el nuevo fichero auxiliar
'            Print #canalaux, linea  'Escribe la linea
'            linea = Trim$(linea)
'        Close #canalaux
  
'        Open "EpisodioAux.tmp" For Input As #canalaux
            
            Contadat = 1  'Variable contadora de cada dato que se lee de la linea
            j = 0   'Indice de los valores de cada serie
            
'            If Not Me.CodigosAnsi Then
'            'Si se cumple esto, los datos estan en código ASCII y hay que
'            'transformarlos (solo se cambia el nombre, pues los datos
'            'numéricos no dan problemas)
'             '   Lecturas.Lee canalaux, miNombreSerie(Me.Serie): Contadat = Contadat + 1:
'             '   miNombreSerie(Me.Serie) = Lecturas.AsciiAAnsi(miNombreSerie(Me.Serie))  'Llama a la función AsciiAAnsi del módulo CodASCII
'            Else
'                Lecturas.Lee canalaux, miNombreSerie(Me.Serie): Contadat = Contadat + 1:  'Lectura normal del nombre
'            End If
            
            
            'Se leen el resto de los datos de la linea:X,Y,Z,P,Valores
            LeeLineaFichero Linea, Me.NumeroIntervalos
           
'            Lecturas.Lee canalaux, miXSerie(Me.Serie): Contadat = Contadat + 1
'            Lecturas.Lee canalaux, miYSerie(Me.Serie): Contadat = Contadat + 1
'            Lecturas.Lee canalaux, miZSerie(Me.Serie): Contadat = Contadat + 1
'            Lecturas.Lee canalaux, miPSerie(Me.Serie): Contadat = Contadat + 1
            
            'Se inicializan los valores máximo y mínimo de la linea
            miValorMaximoSerie(Me.Serie) = -1E+36
            miValorMinimoSerie(Me.Serie) = 1E+36
            MSeriesSeleccionadas(Me.Serie) = Me.Serie
            
            'Lectura de valores, si es un valor inválido (-1),la propiedad
            'correspondiente se incrementa en una unidad:
            For j = 1 To Me.NumeroIntervalos
                miValorIntervaloSerie(j, Me.Serie) = vector(j + 4) 'ver sub leelineafichero
                
                'Lecturas.Lee canalaux, miValorIntervaloSerie(j, Me.Serie)
                If miValorIntervaloSerie(j, Me.Serie) = Me.ValorInvalido Then
                    miCantidadInvalidosSerie(Me.Serie) = miCantidadInvalidosSerie(Me.Serie) + 1
                Else
                    'Comprobación de extremos
                    If miValorIntervaloSerie(j, Me.Serie) > miValorMaximoSerie(Me.Serie) Then _
                        miValorMaximoSerie(Me.Serie) = miValorIntervaloSerie(j, Me.Serie)
                    If miValorIntervaloSerie(j, Me.Serie) < miValorMinimoSerie(Me.Serie) Then _
                        miValorMinimoSerie(Me.Serie) = miValorIntervaloSerie(j, Me.Serie)
                 End If
            Next j
            
            If Me.NumeroIntervalos = miCantidadInvalidosSerie(Me.Serie) Then
                miValorMinimoSerie(Me.Serie) = Me.ValorInvalido
                miValorMaximoSerie(Me.Serie) = Me.ValorInvalido
            Else
                If UCase(Me.TipoSerie) = "P" Then 'precipitaciones
                    If miValorMaximoSerie(Me.Serie) > miMaximaPrecipDeEpisodio Then _
                        miMaximaPrecipDeEpisodio = miValorMaximoSerie(Me.Serie)
                    If miValorMinimoSerie(Me.Serie) < miMinimaPrecipDeEpisodio Then _
                        miMinimaPrecipDeEpisodio = miValorMinimoSerie(Me.Serie)
                End If
                If UCase(Me.TipoSerie) = "S" Then 'caudal embalses
                    If miValorMaximoSerie(Me.Serie) > miMaximoCaudalEmbalse Then _
                        miMaximoCaudalEmbalse = miValorMaximoSerie(Me.Serie)
                End If
                If UCase(Me.TipoSerie) = "Q" Then 'caudal aforos
                    If miValorMaximoSerie(Me.Serie) > miMaximoCaudalAforo Then _
                        miMaximoCaudalAforo = miValorMaximoSerie(Me.Serie)
                End If
            End If
            
            'Si no se ha llegado al final de la linea despues de leer
            'el valor asignado de valores, se estima que hay un error
            '(el número de valores es superior al indicado
            'en la linea de especificaciones generales (linea G)):
           ' If EOF(canalaux) = False Then caso = 2
            'If EOF(canalaux) = False Then GoTo SiError_LeeFichero
        ' If EOF(nftmp) = False Then caso = 2
        ' If EOF(nftmp) = False Then GoTo SiError_LeeFichero
        'Close nftmp
    
    Loop  'Se vuelve al inicio del bucle para leer un nueva linea
    Close #nftmp
    
    Kill "Episodio.Tmp"    'Se eliminan de la memoria los ficheros temporales
    Kill "EpisodioAux.tmp"
    Me.MsgLectura = " Leidas " & Me.Series & " series en " & Me.nombrefichero & Chr(10) & Chr(13) & Me.MsgLectura
    Acumula
    SiHeAcumulado = True
    CalculaMaximoEnIntervalo
    SiHeMaxIntervalo = True
    RaiseEvent FicheroLeido(EpiFichLeidoOK)
Sal_LeeFichero:
    Exit Sub

SiError_LeeFichero:  'Rutina de control de errores
    If caso = 2 Then
        Me.ErrorSobranDatos = "Error al leer en la serie " & Me.Serie & " y en el dato" & Contadat + (j) & " .Compruebe que no sobren datos"     '  el numero de intervalos no supere al indicado al inicio del fichero de episodio"
    Else
        Me.ErrorFaltanDatos = "Error al leer en la serie " & Me.Serie & " y en el dato" & Contadat + (j) & " .Compruebe que no falten datos "     ' o comillas en el nombre de la serie temporal"
    End If
     Reset
    RaiseEvent FicheroLeido(EpiFichLeidoerr)
    Resume Sal_LeeFichero
End Sub

Public Property Get MsgLectura() As String  'Identifica el estado de la lectura del fichero
    MsgLectura = miMsgLectura
End Property
Public Property Let MsgLectura(vNewValue As String) 'Identifica el estado de la lectura del fichero
    miMsgLectura = vNewValue
End Property

Public Property Get ErrorFaltanDatos() As String 'Identifica el error que se produce cuando falta algún dato en la serie o comillas en el nombre de la misma
    ErrorFaltanDatos = miErrorFaltanDatos
End Property
Public Property Let ErrorFaltanDatos(vNewValue As String) 'Identifica el error que se produce cuando falta algún dato en la serie o comillas en el nombre de la misma
    miErrorFaltanDatos = vNewValue
End Property

Public Property Get ErrorSobranDatos() As String
'Identifica el error producido cuando sobran datos en la serie,
'de acuerdo con el número de intervalos leido en la linea de especificaciones generales (linea G)
    ErrorSobranDatos = miErrorSobranDatos
End Property
Public Property Let ErrorSobranDatos(vNewValue As String)
'Identifica el error producido cuando sobran datos en la serie,
'de acuerdo con el número de intervalos leido en la linea de especificaciones generales (linea G)
    miErrorSobranDatos = vNewValue
End Property

Public Property Get ErrorFicheroLeido() As String 'Identifica el error al leer la línea de especificaciones generales (linea G) o el tipo de la serie
    ErrorFicheroLeido = miErrorFicheroLeido
End Property
Public Property Let ErrorFicheroLeido(vNewValue As String) 'Identifica el error al leer la línea de especificaciones generales (linea G) o el tipo de la serie
    miErrorFicheroLeido = vNewValue
End Property

Public Property Get ErrorReconocimientoFecha() As String  'Identifica el error que se produce cuando el nombre del archivo no se puede traducir a una fecha de episodio
    ErrorReconocimientoFecha = miErrorReconocimientoFecha
End Property
Public Property Let ErrorReconocimientoFecha(vNewValue As String) 'Identifica el error que se produce cuando el nombre del archivo no se puede traducir a una fecha de episodio
    miErrorReconocimientoFecha = vNewValue
End Property

Public Property Get ErrorReconocimientoArchivo() As String
'Identifica el error producido cuando la línea de especificaciones generales
'(linea G) no corresponde con la de un fichero de episodios;
'el archivo no  se reconoce
    ErrorReconocimientoArchivo = miErrorReconocimientoArchivo
End Property
Public Property Let ErrorReconocimientoArchivo(vNewValue As String)
'Identifica el error producido cuando la línea de especificaciones generales
'(linea G) no corresponde con la de un fichero de episodios;
'el archivo no  se reconoce
    miErrorReconocimientoArchivo = vNewValue
End Property

Public Function OrganizaErrores() As String
'Función que prepara el formato de salida de los mensajes de error en la
'etiqueta InfoLab del formulario de información general
    If Me.ErrorFaltanDatos <> "" Then
       OrganizaErrores = Me.ErrorFicheroLeido & Chr(13) & Chr(10) & "   " & Me.ErrorFaltanDatos
    ElseIf Me.ErrorSobranDatos <> "" Then
        OrganizaErrores = Me.ErrorFicheroLeido & Chr(13) & Chr(10) & "   " & Me.ErrorSobranDatos
    Else
        OrganizaErrores = Me.ErrorFicheroLeido & Chr(13) & Chr(10) & "  " & Me.MsgLectura
    End If
End Function

Public Sub Selecciona()
'Procedimiento que organiza la lectura de episodios
   ' Dim f As New FormLeeFichEpi
   ' Load f
   ' Set f.EpisodioOrigen = Me
   ' f.Show 1
        
End Sub

Public Function TipoGraf() As String
    Select Case Me.TipoSerie
    
    Case "P"  'Precipitaciones
        TipoGraf = "BARRAS"
    Case "N"  'Nivel de embalse
        TipoGraf = "LINEAS"
    Case "S"  'Caudal de salida de embalse
        TipoGraf = "LINEAS"
    Case "C"  'Cota de estación de aforo
        TipoGraf = "LINEAS"
    Case "Q"  'Caudal de estación de aforo
        TipoGraf = "LINEAS"
    Case Else 'Desconocido
        TipoGraf = "BARRAS"
    End Select
    

End Function

Public Function TipoColor() As Long
    Select Case Me.TipoSerie
    
    Case "P" 'precipitaciones
        TipoColor = QBColor(1)  'azul oscuro
    Case "N"  'nivel de embalse
        TipoColor = RGB(0, 255, 255)  ' azul claro
    Case "S"  'Caudal de salida de embalse
        TipoColor = QBColor(6)  'amarillo
    Case "C"  'nivel en la estacion de aforos
        TipoColor = QBColor(1) 'azul oscuro
    Case "Q"  'Caudal en estacion de aforos
        TipoColor = RGB(255, 0, 0) 'rojo
    Case Else  'desconocido
        TipoColor = QBColor(0)  'negro
    End Select
End Function

'Public Sub Selecciona_Y_Ordena(umsg As Tipo_Seleccion, Optional orden As Orden_Y_Sel)
'    Dim i As Integer
'    Dim DimMatSel As Integer
'    Load SelSeriesForm
'    Call SelSeriesForm.CargaMatrices(miNombreSerie(), miTipoSerie(), miXSerie(), miYSerie())
'    If umsg = verformulario Then
'        SelSeriesForm.LabelINfo = "    " & Me.nombrefichero
'        SelSeriesForm.Caption = Me.nombrefichero & "  " & "Selección y ordenación de series temporales"
'        SelSeriesForm.Show 1
'    Else
'        Select Case orden
'        Case OrdenAlfa
'            SelSeriesForm.Ordena OrdenAlfa
'        Case OrdenNS
'            SelSeriesForm.Ordena OrdenNS
'        Case OrdenEO
'            SelSeriesForm.Ordena OrdenEO
'        Case OrdenTipos
'            SelSeriesForm.Ordena OrdenTipos
'        Case Else
'            SelSeriesForm.Ordena OrdenDefecto
'        End Select
'        SelSeriesForm.Selecciona umsg
'    End If
'
'    'El formulario se muestra en forma modal
'        If umsg = verformulario Then
'            If SelSeriesForm.Tag = "Cancelar" Then Exit Sub
'        End If
'
'        DimMatSel = SelSeriesForm.Elegidos
'        If DimMatSel > 0 Then
'            ReDim MSeriesSeleccionadas(1 To DimMatSel)
'            For i = 1 To DimMatSel
'                SelSeriesForm.ElementoSel = i
'                MSeriesSeleccionadas(i) = SelSeriesForm.ResultadoOrdenYSel
'            Next i
'        End If
'        miOrdenacionSerie = SelSeriesForm.MensajeOrd
'        miSeleccionSerie = SelSeriesForm.MensajeSel
'    Unload SelSeriesForm
'End Sub

Public Sub Redimensiona(NSeries As Long, NInterv As Long)
    ReDim miNombreSerie(1 To NSeries)
    ReDim miXSerie(1 To NSeries)
    ReDim miYSerie(1 To NSeries)
    ReDim miZSerie(1 To NSeries)
    ReDim miPSerie(1 To NSeries)
    ReDim miValorIntervaloSerie(1 To NInterv, 1 To NSeries)
    ReDim miValorMaximoSerie(1 To NSeries)
    ReDim miValorMinimoSerie(1 To NSeries)
    ReDim miCantidadInvalidosSerie(1 To NSeries)
    ReDim miFechaIntervalo(1 To NInterv)
    ReDim MSeriesSeleccionadas(1 To NSeries)
    ReDim MatrizAcumulados(1 To NSeries)
   
End Sub

Public Sub GuardaComo()
'procedimiento que abre un espacio en memoria para guardar un nuevo episodio
'al que se le da una ruta mediante el control CommonDialog



On Error GoTo Si_Error_GuardaComo

    Dim entrada As String
    Dim letras As String  'Variable para analizar si es un fichero de configuración o no
    Dim Titulo As String
    Dim Posicion As Integer
'    FormLeeFichEpi.CommonDialog1.CancelError = True 'Para control de elección de botón
    ' cancelar a través del control de errores 32755 (cdlCancel)
'    If Me.Directorio <> "" Then _
'        FormLeeFichEpi.CommonDialog1.InitDir = Me.Directorio
'    FormLeeFichEpi.CommonDialog1.DialogTitle = "Guardar fichero de episodios"
    'entrada = Left$(Right$(Me.NombreFichero, 3), 1) ' & "??"
    'FormLeeFichEpi.CommonDialog1.DefaultExt = entrada
'    FormLeeFichEpi.CommonDialog1.Filter = "Archivos de episodio (*.*)|*.*"
    'FormLeeFichEpi.CommonDialog1.Flags =
'    FormLeeFichEpi.CommonDialog1.FileName = Me.nombrefichero
'    FormLeeFichEpi.CommonDialog1.ShowSave
'    entrada = FormLeeFichEpi.CommonDialog1.FileName 'Asigna el fichero seleccionado en pantalla
'    entrada = Trim$(entrada)
'    Titulo = Trim$(FormLeeFichEpi.CommonDialog1.FileTitle)
'
            ' Se busca el directorio donde se encuentra el archivo SAIH.CNF
            ' Este será el directorio principal de aplicaciones (DirSAIH)
'            Posicion = InStr(entrada, Titulo) ' Calcula la posicion donde encuentra la variable Titulo
'            entrada = Trim$(Left$(entrada, Posicion - 1)) 'Prepara el nombre del fichero para que node lugar a error en la lectura del mismo
'            Me.Directorio = entrada   'Asigna el fichero a la propiedad
'            Me.nombrefichero = Titulo
'            Me.Escribeme
    
Sal_GuardaComo:
    Exit Sub
Si_Error_GuardaComo:
    If Err = cdlCancel Then
        Resume Sal_GuardaComo
    Else
        MsgBox "Error al guardar archivo de episodio", vbOKOnly + vbCritical
        Resume Sal_GuardaComo
    End If
    
End Sub

Public Sub Escribeme()  'éscribe en formato de episodio los datos de las series temporales
'se utiliza cuando se ha cortado o unido un episodio generando uno nuevo
    Dim canal As Integer
    Dim i As Long
    Dim j As Long
    Dim Linea As String
    Dim LineaG As String
    Dim COMDOB As String * 1
    Dim S As String * 8
    
    COMDOB = Chr$(34)
    
    LineaG = ""
    Linea = ""
    
    canal = FreeFile
    LineaG = "G" & "  " & Me.NumeroIntervalos & "  " & Me.LongitudIntervalos
    Open Me.Directorio & Me.nombrefichero For Output As #canal
        Print #canal, Me.Comentarios
        Print #canal, ""
        Print #canal, LineaG
        Print #canal, ""
        For i = 1 To Me.Series
            Me.Serie = i
            Print #canal, Me.TipoSerie & "   " & COMDOB & Me.NombreSerie & COMDOB;
            Print #canal, "   " & Me.XSerie & "   " & Me.YSerie & "   " & Me.ZSerie & "   " & Me.PSerie & "    ";
            For j = 1 To Me.NumeroIntervalos
                Me.Intervalo = j
                S = Format(Me.Valor, "#.00")
                Print #canal, S;
            Next j
            Print #canal, ""
        Next i
        
    Close #canal
    
End Sub

Public Sub AñadeSerie()

'se crea una serie nueva, redimensionando todas las matrices de almacenamiento de datos
'para el caso de la creación de nuevos episodios apartir, por ejemplo, de su union o corte
    If miSeries = 0 Then
       miSeries = 1
        ReDim miTipoSerie(1 To Me.Series)
        ReDim miNombreSerie(1 To Me.Series)
        ReDim miXSerie(1 To Me.Series)
        ReDim miYSerie(1 To Me.Series)
        ReDim miZSerie(1 To Me.Series)
        ReDim miPSerie(1 To Me.Series)
        ReDim miValorIntervaloSerie(1 To Me.NumeroIntervalos, 1 To Me.Series)
        ReDim miValorMaximoSerie(1 To Me.Series)
        ReDim miValorMinimoSerie(1 To Me.Series)
        ReDim miCantidadInvalidosSerie(1 To Me.Series)
        ReDim miFechaIntervalo(1 To Me.NumeroIntervalos)
        ReDim MSeriesSeleccionadas(1 To Me.Series)
        ReDim MatrizAcumulados(1 To Me.Series)
        MSeriesSeleccionadas(1) = 1
        
    Else
    
        miSeries = miSeries + 1
        ReDim Preserve miTipoSerie(1 To Me.Series)
        ReDim Preserve miNombreSerie(1 To Me.Series)
        ReDim Preserve miXSerie(1 To Me.Series)
        ReDim Preserve miYSerie(1 To Me.Series)
        ReDim Preserve miZSerie(1 To Me.Series)
        ReDim Preserve miPSerie(1 To Me.Series)
        ReDim Preserve miValorIntervaloSerie(1 To Me.NumeroIntervalos, 1 To Me.Series)
        ReDim Preserve miValorMaximoSerie(1 To Me.Series)
        ReDim Preserve miValorMinimoSerie(1 To Me.Series)
        ReDim Preserve miCantidadInvalidosSerie(1 To Me.Series)
        ReDim Preserve miFechaIntervalo(1 To Me.NumeroIntervalos)
        ReDim Preserve MSeriesSeleccionadas(1 To Me.Series)
        ReDim Preserve MatrizAcumulados(1 To Me.Series)
        MSeriesSeleccionadas(miSeries) = miSeries
        
    End If
    miValorMaximoSerie(Me.Series) = -1E+36
    miValorMinimoSerie(Me.Series) = 1E+36
End Sub

Public Sub CalculaExtremos() 'calcula los extremos (maximo y minimo) del episodio, cuando no se ha
'pasado por el proceso de lectura del fichero (cuando se ha creado desde fuera)
    Dim i As Integer, j As Long
    ' En Leefichero se obtienen en el mismo proceso de lectura
    ' con lo que se evita duplicar bucles
    CalculaCantidadInvalidos
    For i = 1 To Me.Series
        miValorMaximoSerie(i) = -1E+38
        miValorMinimoSerie(i) = 1E+38
        For j = 1 To Me.NumeroIntervalos
            If miValorIntervaloSerie(j, i) <> Me.ValorInvalido Then
                If miValorIntervaloSerie(j, i) > miValorMaximoSerie(i) Then miValorMaximoSerie(i) = miValorIntervaloSerie(j, i)
                If miValorIntervaloSerie(j, i) < miValorMinimoSerie(i) Then miValorMinimoSerie(i) = miValorIntervaloSerie(j, i)
            End If
        Next j
        If Me.NumeroIntervalos = miCantidadInvalidosSerie(i) Then
            miValorMinimoSerie(i) = Me.ValorInvalido
            miValorMaximoSerie(i) = Me.ValorInvalido
        Else
            If UCase(miTipoSerie(i)) = "P" Then 'precipitaciones
                If miValorMaximoSerie(i) > miMaximaPrecipDeEpisodio Then _
                    miMaximaPrecipDeEpisodio = miValorMaximoSerie(i)
                If miValorMinimoSerie(i) < miMinimaPrecipDeEpisodio Then _
                    miMinimaPrecipDeEpisodio = miValorMinimoSerie(i)
            End If
            If UCase(miTipoSerie(i)) = "S" Then 'caudal embalses
                If miValorMaximoSerie(i) > miMaximoCaudalEmbalse Then _
                    miMaximoCaudalEmbalse = miValorMaximoSerie(i)
            End If
            If UCase(miTipoSerie(i)) = "Q" Then 'caudal aforos
                If miValorMaximoSerie(i) > miMaximoCaudalAforo Then _
                    miMaximoCaudalAforo = miValorMaximoSerie(i)
            End If
            
        End If
    Next i

End Sub
Public Sub CalculaCantidadInvalidos()
    Dim i As Integer, j As Long
    ' Obtiene el número de valores inválidos en todas las series cuando no se ha pasado por el proceso de
    'lectura del fichero
    ' En Leefichero se obtienen en el mismo proceso de lectura
    ' con lo que se evita duplicar bucles
    For i = 1 To Me.Series
        miCantidadInvalidosSerie(i) = 0
        For j = 1 To Me.NumeroIntervalos
            If miValorIntervaloSerie(j, i) = Me.ValorInvalido Then _
            miCantidadInvalidosSerie(i) = miCantidadInvalidosSerie(i) + 1
        Next j
        If miCantidadInvalidosSerie(i) = Me.NumeroIntervalos Then
            miValorMaximoSerie(i) = Me.ValorInvalido
            miValorMinimoSerie(i) = Me.ValorInvalido
        End If
    Next i

End Sub


Public Property Get Comentarios() As String  'para guardar las notas iniciales que se quieran
    'escribir en el fichero, al escribir un nuevo episodio
    Comentarios = misComentarios
End Property

Public Property Let Comentarios(ByVal vNewValue As String)
    misComentarios = vNewValue
End Property

Private Sub Class_Initialize()
'Llama al procedimiento que inicializa los valores y propiedades de la clase
    IniciaSeries
    
End Sub

Private Sub Class_Terminate()
SiHeAcumulado = False
End Sub

Public Property Get MaximoEnIntervalo() As Double 'para el maximo valor del intervalo, a lo largo de todas las sseries del episodio
If SiHeMaxIntervalo = False Then
    CalculaMaximoEnIntervalo
    SiHeMaxIntervalo = True  'indica si se ha calculado el maximo valor del intervalo
End If

MaximoEnIntervalo = miValorMaximoEnIntervalo(Me.Intervalo)
End Property



Public Sub CalculaMaximoEnIntervalo() 'a lo largo de todas las sseries del episodio
Dim i As Integer
Dim j As Integer
Dim maxin As Double


ReDim miValorMaximoEnIntervalo(1 To Me.NumeroIntervalos)



For i = 1 To Me.NumeroIntervalos
    Me.Intervalo = i
    maxin = -1E+38
    For j = 1 To Me.SeriesSel
        Me.Serie = j
        If Me.TipoSerie = "P" Then
            If Me.Valor > maxin Then
                maxin = Me.Valor
            End If
        End If
    Next j
    miValorMaximoEnIntervalo(i) = maxin
Next i


SiHeMaxIntervalo = True

End Sub



Public Property Get MaximoCaudalEmbalse() As Double
MaximoCaudalEmbalse = miMaximoCaudalEmbalse
End Property

Public Property Let MaximoCaudalEmbalse(v As Double)
miMaximoCaudalEmbalse = v
End Property


Public Property Get MaximoCaudalAforo() As Double
MaximoCaudalAforo = miMaximoCaudalAforo
End Property

Public Property Let MaximoCaudalAforo(v As Double)
miMaximoCaudalAforo = v
End Property

Public Function QuitaComillas(plu As String)
 Dim longitud As Integer
 Dim i%
 Dim Lee As String
 longitud = Len(plu)
    For i = 1 To longitud
    Lee = Mid(plu, i, 1)
     If Lee = Chr(34) Then
        plu = Left(plu, i - 1) & Right(plu, longitud - i)
        longitud = Len(plu)
     End If
    Next i
    QuitaComillas = plu
End Function



Public Sub LeeLineaFichero(Linea As String, miNumeroIntervalos)
   Dim logLinea As Integer
   Dim pos1 As Integer
   Dim NombreSerie As String
   Dim Leido As String
   Dim n%
   Dim numero As Double
   ReDim vector(1 To miNumeroIntervalos + 4) ' esdecir todo incluso x y z p pi
   
     logLinea = Len(Linea)  'halla la longitud de la línea como string
     pos1 = 1 + InStr(5, Linea, Chr(34)) 'busca la posicion de las últimas comillas dobles
     NombreSerie = Mid(Linea, 3, pos1 - 3)
     NombreSerie = QuitaComillas(NombreSerie) 'pluvi = Trim(pluvi)
     miNombreSerie(Me.Serie) = NombreSerie
     Linea = Mid(Linea, pos1 + 1)
     logLinea = Len(Linea)
     n = 0
     
     Do While logLinea > 1
        n = n + 1
        Linea = Linea & " "
        pos1 = InStr(1, Linea, Chr(32))
        Leido = Mid(Linea, 1, pos1)
            vector(n) = Val(Leido)
        Linea = Mid(Linea, pos1)
        Linea = Trim(Linea)
        logLinea = Len(Linea)
     Loop
     
     miXSerie(Me.Serie) = vector(1)
     miYSerie(Me.Serie) = vector(2)
     miZSerie(Me.Serie) = vector(3)
     miPSerie(Me.Serie) = vector(4)

End Sub
