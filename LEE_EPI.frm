VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "Mswinsck.ocx"
Begin VB.Form LEE_EPI 
   AutoRedraw      =   -1  'True
   Caption         =   "Modelo de crecidas"
   ClientHeight    =   9705
   ClientLeft      =   2580
   ClientTop       =   1230
   ClientWidth     =   10815
   Icon            =   "LEE_EPI.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9705
   ScaleWidth      =   10815
   Begin MSComDlg.CommonDialog CD1 
      Left            =   9465
      Top             =   3555
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton btnImprimir 
      Caption         =   "Imprimir informe"
      Height          =   480
      Left            =   9450
      TabIndex        =   30
      Top             =   2385
      Width           =   1230
   End
   Begin VB.CheckBox Decimales 
      Caption         =   "Decimales con comas"
      Height          =   315
      Left            =   7890
      TabIndex        =   29
      Top             =   2010
      Width           =   2400
   End
   Begin VB.TextBox textparametro 
      BackColor       =   &H00FFC0FF&
      Height          =   285
      Left            =   4785
      TabIndex        =   28
      Text            =   "0"
      Top             =   1035
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   9240
      TabIndex        =   27
      Top             =   4560
      Visible         =   0   'False
      Width           =   975
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid2 
      Height          =   855
      Left            =   8400
      TabIndex        =   26
      Top             =   4200
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   1508
      _Version        =   393216
      Rows            =   3
      FixedRows       =   0
      Appearance      =   0
   End
   Begin VB.CommandButton btnsalir 
      Caption         =   "SALIR"
      Height          =   375
      Left            =   7920
      TabIndex        =   20
      Top             =   3480
      Width           =   1455
   End
   Begin MSWinsockLib.Winsock connTCP 
      Left            =   9600
      Top             =   2400
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton btnProceso 
      Caption         =   "Inicio"
      Height          =   255
      Left            =   7920
      TabIndex        =   19
      Top             =   3240
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      Caption         =   "Fechas"
      Height          =   1695
      Left            =   7800
      TabIndex        =   14
      Top             =   240
      Width           =   3015
      Begin VB.ComboBox Comboaño 
         Height          =   315
         Left            =   960
         TabIndex        =   23
         Text            =   "Serie"
         Top             =   1200
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.TextBox TextfechaFIN 
         Height          =   285
         Left            =   960
         TabIndex        =   16
         Text            =   "fin"
         Top             =   720
         Width           =   1695
      End
      Begin VB.TextBox TextfechaINI 
         Height          =   285
         Left            =   960
         TabIndex        =   15
         Text            =   "ini"
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label Label3 
         Caption         =   "Fecha FIN"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Fecha INI"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.CommandButton btnCalculoestandar 
      Caption         =   "Calculo estandar"
      Height          =   255
      Left            =   7920
      TabIndex        =   13
      ToolTipText     =   "Realiza el calculo con los 3 numeros de curva en estado seco, normal y  humedo"
      Top             =   3000
      Width           =   1455
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   5280
      TabIndex        =   12
      Top             =   1800
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   240
      TabIndex        =   11
      Top             =   240
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Frame Frame2 
      Caption         =   "Variables"
      Height          =   3495
      Left            =   4080
      TabIndex        =   5
      Top             =   480
      Width           =   3615
      Begin VB.CommandButton BTNPERDIDAS 
         Caption         =   "PER"
         Height          =   255
         Left            =   2760
         TabIndex        =   25
         Top             =   240
         Width           =   735
      End
      Begin VB.CommandButton btnSUB 
         Caption         =   "Guardar Fichero"
         Height          =   300
         Left            =   1665
         TabIndex        =   22
         Top             =   585
         Width           =   1815
      End
      Begin MSFlexGridLib.MSFlexGrid tablaSUB 
         Height          =   2415
         Left            =   120
         TabIndex        =   10
         Top             =   960
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   4260
         _Version        =   393216
      End
      Begin VB.TextBox Textfactor 
         Height          =   285
         Left            =   1680
         TabIndex        =   7
         Text            =   "1"
         ToolTipText     =   "factor que pasa los valores de intensidad a acumulados en el periodo de un intervalo"
         Top             =   600
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.TextBox Textintervalo 
         Height          =   285
         Left            =   1680
         TabIndex        =   6
         Text            =   "0.5"
         Top             =   225
         Width           =   975
      End
      Begin VB.Label Label10 
         Caption         =   "Factor intervalo"
         Height          =   375
         Left            =   240
         TabIndex        =   9
         Top             =   600
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Label Label6 
         Caption         =   "Intervalo en horas"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.ComboBox comboSUB 
      Height          =   315
      Left            =   5520
      TabIndex        =   4
      Text            =   "Selecciona Subcuenca"
      Top             =   105
      Width           =   2175
   End
   Begin VB.CheckBox Check1 
      Caption         =   "No borrar grafico"
      Height          =   255
      Left            =   8880
      TabIndex        =   3
      Top             =   3240
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   5415
      Left            =   0
      ScaleHeight     =   5355
      ScaleWidth      =   10515
      TabIndex        =   2
      Top             =   4080
      Width           =   10575
      Begin VB.Shape Shape1 
         BorderColor     =   &H000080FF&
         Height          =   735
         Left            =   5400
         Top             =   1320
         Visible         =   0   'False
         Width           =   1935
      End
   End
   Begin VB.CommandButton BTN2 
      BackColor       =   &H8000000B&
      Caption         =   "Calcula actual"
      Height          =   495
      Left            =   7920
      TabIndex        =   1
      ToolTipText     =   "Calcula con el numero de curva actual"
      Top             =   2400
      Width           =   1455
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   3855
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   6800
      _Version        =   393216
      Rows            =   4
      FixedRows       =   3
   End
   Begin VB.Label Label4 
      Caption         =   "Subcuenca:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4080
      TabIndex        =   24
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "FVB dic 2002"
      BeginProperty Font 
         Name            =   "Futura Lt BT"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   360
      TabIndex        =   21
      Top             =   9480
      Width           =   1050
   End
End
Attribute VB_Name = "LEE_EPI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
'Variables de la conexxion al SAIH
Private mConexion As cSaih
Private cColVari As Collection
Private bCancela As Boolean

'Resto de variables
Public Epi As Episodio
'Public NuevoModelo As New modHidro
Dim vector() ' vector que almacena la lluvia media antes de los calculos
'Dim VectorCaudalAArriba()
Dim x() ' almacena el tiempo o la x
Dim y() ' en principio es el caudal
Dim Y_Llmedia()
Dim Y_Llneta()
Dim migraf As New GraficoXY
Dim Ymax As Double
Dim NoBorrarGrafico As Boolean
Public JUCAR As Cuenca

'Private Etiqueta_dX 'Sirve para no cambiar la gráfica
Dim BorrarGrafico  As Boolean
Dim haceZoom As Boolean 'sirve para el zoom sobre el picture
Dim X1E, X2E, Y1E, Y2E

Dim SerieTemporal As Integer 'almacena la serie temporal global que lee del fichero config.ini

'Para guardar los gráficos
Private Declare Function DIWriteJpg Lib "DIjpg.dll" (ByVal DestPath As String, ByVal quality As Long, ByVal progressive As Long) As Long




Private Sub BTNPERDIDAS_Click()

If Me.MSFlexGrid2.Visible = True Then
 Me.MSFlexGrid2.Visible = False
Else
 Me.MSFlexGrid2.Visible = True
End If

End Sub

Private Sub btnsalir_Click()
 'FINALIZA EL PROGRAMA
 Unload Me
 End
End Sub

Private Sub btnSUB_Click()
'Guarda el fichero
Dim Indice
Dim Nfichero
  Indice = Me.comboSUB.ListIndex + 1
  
  Me.JUCAR.miColecSub(Indice).NOMBRE_SUBCUENCA = Me.tablaSUB.TextMatrix(1, 1)
  Me.JUCAR.miColecSub(Indice).AREA_KM2 = Me.tablaSUB.TextMatrix(2, 1)
  
  For i = 1 To Me.JUCAR.miColecSub(Indice).miColeccionPluviometros.Count
   Me.JUCAR.miColecSub(Indice).miColeccionPluviometros(i).CoefThiessen = Me.tablaSUB.TextMatrix(i + 1, 2)
   Me.JUCAR.miColecSub(Indice).miColeccionPluviometros(i).NombrePluvi = Me.tablaSUB.TextMatrix(i + 1, 3)
  Next i
  
  Me.JUCAR.miColecSub(Indice).NC_S = Me.tablaSUB.TextMatrix(3, 1)
  Me.JUCAR.miColecSub(Indice).NC_N = Me.tablaSUB.TextMatrix(4, 1)
  Me.JUCAR.miColecSub(Indice).NC_H = Me.tablaSUB.TextMatrix(5, 1)
  Me.JUCAR.miColecSub(Indice).PUNTO_DRENAJE_ASOCIADO = Me.tablaSUB.TextMatrix(6, 1)
  Me.JUCAR.miColecSub(Indice).tc = Me.tablaSUB.TextMatrix(7, 1)
  Me.JUCAR.miColecSub(Indice).AFORO_AGUAS_ARRIBA = Me.tablaSUB.TextMatrix(9, 1)
  Me.JUCAR.miColecSub(Indice).Desfase = Me.tablaSUB.TextMatrix(10, 1)

 '*******************************************************************
 Nfichero = FreeFile
 Open App.Path & "\Def_sub.txt" For Output As #Nfichero
 
 For n = 1 To JUCAR.miColecSub.Count
  Print #Nfichero, "************************************"
  Print #Nfichero, "NOMBRE_SUBCUENCA: " & JUCAR.miColecSub(n).NOMBRE_SUBCUENCA
  Print #Nfichero, "PLUVIOMETROS: "
  
  For i = 1 To Me.JUCAR.miColecSub(Indice).miColeccionPluviometros.Count
   Print #Nfichero, Me.JUCAR.miColecSub(n).miColeccionPluviometros(i).CoefThiessen & "  " & Me.JUCAR.miColecSub(n).miColeccionPluviometros(i).NombrePluvi
  Next i
  Print #Nfichero, "AREA_KM2:  " & JUCAR.miColecSub(n).AREA_KM2
  Print #Nfichero, "TC_HORAS:  " & JUCAR.miColecSub(n).tc
  Print #Nfichero, "NC_SECO_NORMAL_HUMEDO:  " & JUCAR.miColecSub(n).NC_S & ", " & JUCAR.miColecSub(n).NC_N & ", " & JUCAR.miColecSub(n).NC_H
  Print #Nfichero, "PUNTO_DRENAJE_ASOCIADO: " & JUCAR.miColecSub(n).PUNTO_DRENAJE_ASOCIADO
  Print #Nfichero, "AFORO_AGUAS_ARRIBA:     " & JUCAR.miColecSub(n).AFORO_AGUAS_ARRIBA
  Print #Nfichero, "DESFASE:                " & JUCAR.miColecSub(n).Desfase
 Next n
 Close Nfichero
End Sub

Private Sub Comboaño_Click()
'esta sub pone el año que corresponda
Dim FechaINI As Date
Dim FechaFIN As Date

Dim Resta As Integer

FechaINI = Me.TextfechaINI.Text
FechaFIN = Me.TextfechaFIN.Text

Resta = -Year(FechaINI) + (Me.Comboaño.Text)

Me.TextfechaINI.Text = DateAdd("yyyy", Resta, FechaINI)
Me.TextfechaFIN.Text = DateAdd("yyyy", Resta, FechaFIN)
End Sub

Private Sub connTCP_DataArrival(ByVal bytesTotal As Long)
    Dim strData         As Variant
 
    connTCP.GetData strData, vbString
    mConexion.Actualizar strData
End Sub
Public Sub EnviaDatosTCP(ByVal strComando As String)
   connTCP.SendData strComando
End Sub
Public Sub CierraConexion()
    connTCP.Close
End Sub
Private Sub prc_Generar(Variable)
'Esta es la sub donde se cogen los datos de OFELIA
    Dim k As Integer
    Dim i As Integer
    bCancela = False
   'Saca los datos de una sola variable
 
        prc_GeneraLLuvias Variable ' llama al procedimiento que lee
        
        Do While Not mConexion.OKSerieSaih And Not mConexion.ErrorSerie
            DoEvents
            If bCancela Then Exit Do
        Loop
        
        If (Not mConexion.ColDatos Is Nothing) And Not bCancela Then
           ' txtResultado.Text = txtResultado.Text & "Variable " & cColVari.Item(i).Variable & vbLf
           
        End If
        If bCancela Then
            If connTCP.State <> 0 Then
                connTCP.Close
            End If
        End If
       
      'Dibuja_grafica "LINEAS", cColVari.Item(Ind).Nombre
End Sub
Private Sub prc_GeneraLLuvias(ByVal strVariable As String)
Dim Serie As Integer

    If Not mConexion Is Nothing Then
        Set mConexion = Nothing
    End If
    
    Set mConexion = New cSaih
    
    mConexion.ErrorSerie = False
    mConexion.OKSerieSaih = False
    mConexion.Intervalo = Val(Me.Textintervalo.Text) * 60 '
    mConexion.FechaInicial = Format(TextfechaINI.Text, "yyyymmddhhnn") '& "0800"
    mConexion.FechaFinal = Format(TextfechaFIN.Text, "yyyymmddhhnn")   '& "0800"
    mConexion.Periodo = "0"
    mConexion.Variable = strVariable
    mConexion.Acum = "0" '"A"

    mConexion.Serie = SerieTemporal 'SeleccionSerie(Me.Comboaño.Text) '"0" '0=5min 1=diarias
    mConexion.FormTCP = Me
    prc_ConectarSaih

End Sub

Public Sub prc_ConectarSaih()

    If connTCP.State <> 0 Then
        connTCP.Close
    End If
    
    connTCP.RemotePort = mConexion.RemotePort
    connTCP.RemoteHost = mConexion.RemoteHost
    connTCP.Protocol = mConexion.Protocol
    connTCP.Connect mConexion.RemoteHost, mConexion.RemotePort
    DoEvents
    
End Sub




'*******************************************
 
Private Sub BTN2_Click()
If Me.MSFlexGrid1.Rows > 4 Then

NuevoCALCULO Val(Me.tablaSUB.TextMatrix(4, 1)), 2
 
Picture1.Cls
Dibuja_grafica 1, Me.MSFlexGrid1.Cols - 3, Me.MSFlexGrid1.Cols - 2, Me.MSFlexGrid1.Cols - 1
Dibuja_Una_Serie 3, "ROJO"

'Dibuja_Una_Serie Me.MSFlexGrid1.Cols - 1, "VERDE"
 Pinta_Linea_Actual
 
 'Guarda la imagen
  Me.GuardaIMG "grafico", Me.Picture1, 80
End If
End Sub

Private Sub btnCalculoestandar_Click()
'Realiza el calculo para los 3 numeros de curva principales
'SECO
    NuevoCALCULO Val(Me.tablaSUB.TextMatrix(3, 1)), 2
    Dibuja_Una_Serie Me.MSFlexGrid1.Cols - 1, "GRIS"
'HUMEDO
    NuevoCALCULO (Val(Me.tablaSUB.TextMatrix(5, 1))), 2
    Dibuja_Una_Serie Me.MSFlexGrid1.Cols - 1, "GRIS"
'NORMAL
    NuevoCALCULO (Val(Me.tablaSUB.TextMatrix(4, 1))), 2
    Dibuja_Una_Serie Me.MSFlexGrid1.Cols - 1, "AZUL"

End Sub




Private Sub Dibuja_Una_Serie(COL As Integer, Optional Color As String)
 Dim i%
 Dim NumeroPuntosDibu
 Dim LongInt As Integer
 Dim Fecha
 
  migraf.DivisionHorizontal = 2
  migraf.DivisionVertical = 1
  migraf.PosicionMarco = 2
  migraf.Contenedor = Me.Picture1
 NumeroPuntosDibu = migraf.NumeroDatos
 ' If NumeroPuntosDibu < Me.MSFlexGrid1.Rows Then
 '  NumeroPuntosDibu = Me.MSFlexGrid1.Rows - 3
 '  migraf.NumeroDatos = NumeroPuntosDibu
 ' End If
 'Reescribe las y
 
ReDim x(1 To NumeroPuntosDibu) 'Fechas
ReDim y(1 To NumeroPuntosDibu) 'caudales

 For i = 1 To NumeroPuntosDibu
    y(i) = Val(Me.MSFlexGrid1.TextMatrix(i + 2, COL)) 'valor
    Fecha = Me.MSFlexGrid1.TextMatrix(i + 2, col_Fecha)
    If Fecha = "" Then
     Fecha = (x(i - 1) + (Val(Textintervalo.Text) / 24))
     Me.MSFlexGrid1.TextMatrix(i + 2, 1) = Format(Fecha, "dd/mm/yyyy hh:nn")
    End If
    x(i) = CDate(Me.MSFlexGrid1.TextMatrix(i + 2, 1))
    
 Next i
 
  'Mete los puntos  en el gráfico
  If Ymax = 0 Then Ymax = Maximo(y)
  If Ymax = 0 Then Ymax = 1
    migraf.Y2Etiqueta = Ymax  'etiqueta del máximo valor
    migraf.Y1Etiqueta = 0
    
    migraf.TituloGrafico = ""
    migraf.TipoGrafico = "LINEAS"
    If UCase(Color) = "AZUL" Then
        migraf.ColorDatos = RGB(0, 0, 255)
    ElseIf UCase(Color) = "ROJO" Then
        migraf.ColorDatos = RGB(255, 0, 0)
    ElseIf UCase(Color) = "GRIS" Then
        migraf.ColorDatos = RGB(200, 200, 200)
    ElseIf UCase(Color) = "VERDE" Then
        migraf.ColorDatos = RGB(100, 100, 50)
    Else
        migraf.ColorDatos = RGB(250, 10, 5)
    End If
    '    migraf.ColorEjes = 0
    '    migraf.ColorRejilla = 0
    '    migraf.ColorMarco = 15
    '    migraf.Rejilla = True
        'en las etiquetas se meten los valores máximos y mínimos de las series.
        'migraf.X1Etiqueta = Xmin
        'migraf.X2Etiqueta = Xmax
        
       ' migraf.NumeroDatos = NumeroPuntosDibu
        'en el siguiente for... se meten los datos de caudal
        
        For i = 1 To NumeroPuntosDibu
            migraf.Dato = i      'nº de orden del dato
            migraf.XDato = x(i) 'valor de la x
            migraf.YDato = y(i) 'valor de la y
        Next i
        
        migraf.AnchuraDatos = 2 'ancho de la linea de datos
        migraf.Dibuja

End Sub


Private Sub Dibuja_grafica(col_Fecha As Integer, col_Llmedia As Integer, col_Llneta As Integer, col_Q As Integer)
'Se dibujan los datos en la gráfica. dado el número de columna en el que se encuentran los datos
Dim v As Integer
Dim Fecha
'este método dibuja un gráfico (en este caso la celda seleccionada)
'de la clase GraficoXY en un contenedor.
'Dim FechaInicial_1 As Date
'Dim LongInt
'    NumeroPuntosDibu = Me.MSFlexGrid1.Rows - 1
If Me.MSFlexGrid1.Cols < 3 Then Exit Sub
NumeroPuntosDibu = Me.MSFlexGrid1.Rows - 3 'mConexion.ColDatos.Count + UBound(miHidroUniSCS)

ReDim x(1 To NumeroPuntosDibu)
ReDim y(1 To NumeroPuntosDibu) 'caudales
ReDim Y_Llmedia(1 To NumeroPuntosDibu)
ReDim Y_Llneta(1 To NumeroPuntosDibu)

'FechaInicial_1 = Me.Epi.FechaInicial
'LongInt = Me.Epi.LongitudIntervalos


If NumeroPuntosDibu > Me.MSFlexGrid1.Rows Then Me.MSFlexGrid1.Rows = NumeroPuntosDibu

'Almacena los datos en las variables
For i = 1 To NumeroPuntosDibu
    y(i) = Val(Me.MSFlexGrid1.TextMatrix(i + 2, col_Q)) '4
    Y_Llmedia(i) = Val(Me.MSFlexGrid1.TextMatrix(i + 2, col_Llmedia)) '1
    Y_Llneta(i) = Val(Me.MSFlexGrid1.TextMatrix(i + 2, col_Llneta)) '2
    Fecha = Me.MSFlexGrid1.TextMatrix(i + 2, col_Fecha)
    If Fecha = "" Then
     Fecha = (x(i - 1) + (Val(Textintervalo.Text) / 24))
     Me.MSFlexGrid1.TextMatrix(i + 2, col_Fecha) = Format(Fecha, "dd/mm/yyyy hh:nn")
    End If
    x(i) = CDate(Me.MSFlexGrid1.TextMatrix(i + 2, col_Fecha))
Next i

        Me.Picture1.AutoRedraw = True
'        If NoBorrarGrafico Then
       '' If Ymax = 0 Then

         Ymax = Maximo(y)
         Ymax = Ymax + Ymax / 10
         If Ymax = 0 Then Ymax = 2
         migraf.Y1Etiqueta = 0
        'End If
         migraf.Y2Etiqueta = Ymax  'etiqueta del máximo valor
         
         If Me.MSFlexGrid1.TextMatrix(Me.MSFlexGrid1.Rows - 3, col_Fecha) = "" Then
         ElseIf (Me.MSFlexGrid1.TextMatrix(Me.MSFlexGrid1.Rows - 3, col_Fecha)) > (migraf.X2Etiqueta) Then
          migraf.X2Etiqueta = CDate(Format(Me.MSFlexGrid1.TextMatrix(Me.MSFlexGrid1.Rows - 3, col_Fecha), "dd/mm/yyyy hh:nn"))
         End If
         
'        Else
'         Ymax = Maximo(Y)
'         Ymax = Ymax + Ymax / 10
'         Me.Picture1.Cls
'         migraf.Y2Etiqueta = Ymax  'etiqueta del máximo valor
'         migraf.Y1Etiqueta = 0
'        End If
       migraf.Contenedor = Me.Picture1
       migraf.DivisionHorizontal = 2
        'indica las divisiones en el contenedor, si queremos hacer 4
        'gráficos en un mismo contenedor daremos 2 divisiones horizontales
        'y otras 2 verticales.
    
        migraf.DivisionVertical = 1
        'migraf.IntervaloTemporal = Me.Epi.Intervalo
'*******************************************************
'Pinta en el marco 2 los caudales
'*******************************************************

        migraf.PosicionMarco = 2 'indica en cual de los  submarcos nos encontramos
        'migraf.Rejilla = False
        migraf.TituloGrafico = "Caudales en " & Me.comboSUB.Text ' Me.Epi.NombreSerie
      '  migraf.TituloEjeX = "X"
      '  migraf.TituloEjeY = "Y"
        migraf.ValorInvalido = -1
        'migraf.Y1Etiqueta = Ymin

        migraf.TipoGrafico = "LINEAS"
        migraf.ColorDatos = RGB(0, 0, 255)
        
        migraf.ColorEjes = 0
        migraf.ColorRejilla = 0
        migraf.ColorMarco = 15
        migraf.Rejilla = True
      'En las etiquetas se meten los valores máximos y mínimos de las series.
      'migraf.X1Etiqueta = Xmin
      '  migraf.X2Etiqueta = Xmax
        
        migraf.NumeroDatos = NumeroPuntosDibu
      'En el siguiente for... se meten los datos de caudal
        
        For i = 1 To NumeroPuntosDibu
            migraf.Dato = i      'nº de orden del dato
            migraf.XDato = x(i) 'valor de la x
            migraf.YDato = y(i) 'valor de la y
        Next i
        
        migraf.AnchuraDatos = 2 'ancho de la linea de datos
        migraf.Dibuja
        Pinta_Linea_Actual
'****************************************
'Dibuja el gráfico de lluvia

        migraf.PosicionMarco = 1 'indica en cual de los  submarcos nos encontramos
        migraf.Rejilla = True
        migraf.TituloGrafico = "Precipitación en " & Me.comboSUB.Text ' & Me.Epi.NombreSerie
       ' migraf.TituloEjeX = "X"
       ' migraf.TituloEjeY = "Y"
        migraf.ValorInvalido = -1
        'migraf.Y1Etiqueta = Ymin
        migraf.Y2Etiqueta = Maximo(Y_Llmedia) + 5
        
        'migraf.ColorDato = False
        migraf.TipoGrafico = "BARRAS"
        migraf.ColorDatos = RGB(0, 0, 255)
        
        migraf.ColorEjes = 0
        migraf.ColorRejilla = 11
        migraf.ColorMarco = 0
        migraf.Rejilla = True
        'en las etiquetas se meten los valores máximos y mínimos de las series.
        ' migraf.X1Etiqueta = Xmin
        ' migraf.X2Etiqueta = Xmax
        
        migraf.NumeroDatos = NumeroPuntosDibu
        'en el siguiente for... se meten los datos.
        
     'Dibuja la lluvia media
        For i = 1 To NumeroPuntosDibu
            migraf.Dato = i      'nº de orden del dato
            migraf.XDato = x(i) 'valor de la x
            migraf.YDato = Y_Llmedia(i) 'valor de la y
        Next i
      migraf.Dibuja
      
      'Dibuja la lluvia neta
        For i = 1 To NumeroPuntosDibu
            migraf.Dato = i      'nº de orden del dato
            migraf.XDato = x(i) 'valor de la x
            migraf.YDato = Y_Llneta(i) 'valor de la y
        Next i
      migraf.ColorDatos = RGB(250, 200, 94) 'pinta la lluvia neta de naranja
      migraf.Dibuja
      Pinta_Linea_Actual
End Sub

Private Function Maximo(Valor)
Dim Dimension
    Dimension = UBound(Valor)
    Maximo = 0
    For i = 1 To Dimension
     If Valor(i) > Maximo Then Maximo = Valor(i)
    Next i

End Function

Private Function Minimo(Valor)
Dim Dimension
    Dimension = UBound(Valor)
    Minimo = Maximo(Valor)
    For i = 1 To Dimension
     If Valor(i) < Minimo Then Minimo = Valor(i)
    Next i

End Function

Private Sub btnProceso_Click()
'ESTA sub lleva el proceso principal de calculo
'Aquí se leen los datos
    Dim Variable As String
    Dim Valor As Double
    Dim Suma As Double
    Dim MaxValor As Double
    
If Me.comboSUB.ListIndex = -1 Then Exit Sub
    Me.MousePointer = 11
    Suma = 0
    
'Lee la variable de lluvia del SAIH
Variable = Me.JUCAR.miColecSub(Me.comboSUB.ListIndex + 1).Pluviometros.Item(1).NombrePluvi
prc_Generar Variable 'Llama al procedimiento de conexión al SAIH

If mConexion.ColDatos.Count = 0 Then Exit Sub
            Me.MSFlexGrid1.Clear
            Me.MSFlexGrid1.Rows = mConexion.ColDatos.Count + 3
            Me.MSFlexGrid1.Cols = 4
            Me.MSFlexGrid1.ColWidth(0) = 500
            Me.MSFlexGrid1.ColWidth(1) = 1500
            Me.MSFlexGrid1.TextMatrix(0, 0) = "SUM"
            Me.MSFlexGrid1.TextMatrix(1, 0) = "MAX"
            Me.MSFlexGrid1.TextMatrix(2, 0) = "n"
            Me.MSFlexGrid1.TextMatrix(2, 1) = "Fecha"
            Me.MSFlexGrid1.TextMatrix(2, 2) = "Ll_media"
            Me.MSFlexGrid1.TextMatrix(2, 3) = "Caudal"
            
            
            For k = 1 To mConexion.ColDatos.Count
                'Aqui leo los datos y los pongo donde quiera
                Me.MSFlexGrid1.TextMatrix(k + 2, 0) = k
                Valor = mConexion.ColDatos.Item(k).Valor
                Me.MSFlexGrid1.TextMatrix(k + 2, 1) = mConexion.ColDatos.Item(k).Fecha
                
                'si el sistema lee datos decimales con comas divide por 1000 el valor
                If Me.Decimales.Value = 1 Then
                  Valor = Valor / 1000
                End If
                
                Me.MSFlexGrid1.TextMatrix(k + 2, 2) = Format(Valor, "0.00")
                If Valor > 0 Then Suma = Suma + Valor
                If Valor > MaxValor Then MaxValor = Valor
            Next k
            Me.MSFlexGrid1.TextMatrix(0, 2) = Format(Suma, "0.00")
            Me.MSFlexGrid1.TextMatrix(1, 2) = Format(MaxValor, "0.00")
            
'Lee la variable de Caudal del SAIH
Variable = Me.JUCAR.miColecSub(Me.comboSUB.ListIndex + 1).PUNTO_DRENAJE_ASOCIADO
prc_Generar Variable 'Llama al procedimiento de conexión al SAIH
            Suma = 0
            Valor = 0
            For k = 1 To mConexion.ColDatos.Count
                'Aqui leo los datos y los pongo donde quiera
                 Valor = mConexion.ColDatos.Item(k).Valor
                
                
                'si el sistema lee datos decimales con comas divide por 1000 el valor
                If Me.Decimales.Value = 1 Then
                  Valor = Valor / 1000
                End If
                 
                 
                Me.MSFlexGrid1.TextMatrix(k + 2, 3) = Format(Valor, "0.00")
                If Valor > 0 Then Suma = Suma + Valor
                If Valor > MaxValor Then MaxValor = Valor
            Next k
            Suma = (Suma * Val(Textintervalo.Text) * 3600) / 1000000
            Me.MSFlexGrid1.TextMatrix(0, 3) = Format(Suma, "0.00")
            Me.MSFlexGrid1.TextMatrix(1, 3) = Format(MaxValor, "0.00")

'Nuevas cosas se lee
Variable = Me.JUCAR.miColecSub(Me.comboSUB.ListIndex + 1).AFORO_AGUAS_ARRIBA
If Variable <> "" Then
Me.MSFlexGrid1.Cols = 5
prc_Generar Variable 'Llama al procedimiento de conexión al SAIH
            Suma = 0
            Valor = 0
            For k = 1 To mConexion.ColDatos.Count
                'Aqui leo los datos y los pongo donde quiera
                 Valor = mConexion.ColDatos.Item(k).Valor
                
                'Si el sistema lee datos decimales con comas divide por 1000 el valor
                If Me.Decimales.Value = 1 Then
                  Valor = Valor / 1000
                End If
                
                Me.MSFlexGrid1.TextMatrix(k + 2, 4) = Format(Valor, "0.00")
                If Valor > 0 Then Suma = Suma + Valor
                If Valor > MaxValor Then MaxValor = Valor
            Next k
            Suma = (Suma * Val(Textintervalo.Text) * 3600) / 1000000
            Me.MSFlexGrid1.TextMatrix(0, 4) = Format(Suma, "0.00")
            Me.MSFlexGrid1.TextMatrix(1, 4) = Format(MaxValor, "0.00")
            Me.MSFlexGrid1.TextMatrix(2, 4) = "C A.Arriba"
End If

Set migraf = Nothing
Picture1.Cls

NuevoCALCULO Val(Me.tablaSUB.TextMatrix(4, 1)), 2
'Rellena las fechas que añade el modelo
Me.Rellenafechas

Dibuja_grafica 1, Me.MSFlexGrid1.Cols - 3, Me.MSFlexGrid1.Cols - 2, Me.MSFlexGrid1.Cols - 1
Dibuja_Una_Serie 3, "ROJO"

Me.MousePointer = 0

End Sub

Private Sub lee_fichero_subcuencas()
Dim NombreEpisodio As String
Dim Rutafichero As String
Dim n%

'    CD1.CancelError = True
'    On Error GoTo salir
'    ' Establecer los filtros
'    CD1.InitDir = App.Path
'    CD1.Filter = "Archivos de texto (*.txt)|*.txt|Archivos de episodio" & _
'    "(*.txt)|*.txt|"
'    ' Especificar el filtro predeterminado
'    CD1.FilterIndex = 1
'    ' Presentar el cuadro de diálogo Abrir
'    CD1.ShowOpen
    
 NombreEpisodio = App.Path & "\Def_sub.txt" 'CD1.FileName
 
 Set JUCAR = New Cuenca
 JUCAR.Nombre = "JUCAR"
 'Lee el fichero de cuencas
 JUCAR.LEE_FICH_SUB (NombreEpisodio)
 
 Me.comboSUB.Clear
 
 For n = 1 To JUCAR.NumSubCuencas
  Me.comboSUB.AddItem JUCAR.miColecSub(n).NOMBRE_SUBCUENCA
 Next n
 
salir:
End Sub

Private Sub Check1_Click()
 NoBorrarGrafico = Not NoBorrarGrafico
End Sub





Private Sub Form_Load()
 'carga la ventana
 Me.TextfechaINI.Text = Format(Now() - 2, "dd/mm/yyyy  hh:nn")
 Me.TextfechaFIN.Text = Format(Now() - ((1 / 24) / 5), "dd/mm/yyyy  hh:nn")
  
 NoBorrarGrafico = True
 
 Me.tablaSUB.AllowUserResizing = flexResizeBoth
 Me.MSFlexGrid1.AllowUserResizing = flexResizeBoth
 Me.MSFlexGrid2.AllowUserResizing = flexResizeBoth
 
 Me.MSFlexGrid2.Cols = 2
 Me.MSFlexGrid2.FixedCols = 1
 Me.MSFlexGrid2.FixedRows = 0
 Me.MSFlexGrid2.ColWidth(0) = 1000
 Me.MSFlexGrid2.Rows = 3
 Me.MSFlexGrid2.TextMatrix(0, 0) = "UmbralEsc"
 Me.MSFlexGrid2.TextMatrix(1, 0) = "Infiltración"
 Me.MSFlexGrid2.TextMatrix(2, 0) = "P_Inicial"
 
 'lee el numero de la serie temporal
    Dim nf
    Dim LIn As String
    nf = FreeFile
    
    Open App.Path & "\config.ini" For Input As nf
    Line Input #nf, LIn
    SerieTemporal = Val(LIn)
    Close #nf
    
 '*****
 
' Carga_Series
 lee_fichero_subcuencas
 
    gblnServidor = "goofy"
    gblPuerto = "22377"
    gblDireccion = "192.25.81.41"
    gblProtocolo = "0"
    
    Dim tLocaleInfo As cLocaleInfo
     
    Set tLocaleInfo = New cLocaleInfo
    
    With tLocaleInfo
        gblConfRegional = "." '.sDecimal
    End With
    
    Set tLocaleInfo = Nothing

End Sub

Private Sub comboSUB_Click()
'Selecciona la cuenca e inicia el proceso de cálculo
  Dim Indice As Integer
  Dim nomPuntoAsociado As String
  Dim n%
  Dim Totalseries%
  Dim indicesSeriesAforo%
  Dim exito As Boolean
  Dim Dimension
  Dim Suma As Double
  Dim MaxValor As Double
  
   ReDim miHidroUniSCS(0 To 0)
   ReDim miCaudalAguasArriba(0 To 0)
   
   Indice = Me.comboSUB.ListIndex + 1
   If Indice = 0 Then Exit Sub
   'If Year(Me.TextfechaFIN.Text) <> (Me.Comboaño.Text) Then Exit Sub
   'Comprueba que la fecha inicial sea antes de la final
   If CDate(Me.TextfechaINI.Text) > CDate(Me.TextfechaFIN.Text) Then
    MsgBox "Has introducido mal las fechas", vbInformation, "Fechas erroneas"
    Exit Sub
   End If
   'Da el formato a la Tabla de datos de cuencas
   FormatosTablaSubcuencas Indice

  nomPuntoAsociado = Me.tablaSUB.TextMatrix(6, 1)
  nomPuntoAsociado = Mid(Trim(nomPuntoAsociado), 2, Len(nomPuntoAsociado) - 2)
  If nomPuntoAsociado = "" Then Exit Sub
  'Me.MSFlexGrid1.ColWidth(1) = 1
  Me.MSFlexGrid1.Cols = 6
  Me.MSFlexGrid1.TextMatrix(0, 5) = "Q_drenaje"
  
  'En lo siguiente se hacen todos los cálculos
  btnProceso_Click
  
  Me.tablaSUB.Cols = Me.tablaSUB.Cols + 1
  Me.tablaSUB.TextMatrix(0, Me.tablaSUB.Cols - 1) = "HidroUNI"
  
  If UBound(miHidroUniSCS) > 0 Then
    For n = 1 To UBound(miHidroUniSCS)
      If n > Me.tablaSUB.Rows - 2 Then
       Me.tablaSUB.Rows = Me.tablaSUB.Rows + 1
      End If
      Me.tablaSUB.TextMatrix(n, (Me.tablaSUB.Cols - 1)) = Format(miHidroUniSCS(n), "0.00")
    Next n
  End If
  
Pinta_Linea_Actual
  
Me.GuardaIMG "grafico", Me.Picture1, 80
End Sub

Sub FormatosTablaSubcuencas(Indice As Integer)
'Da formato a la tabla que contendrá los parametros
  Me.tablaSUB.Cols = 4
  Me.tablaSUB.Rows = 11
  Me.tablaSUB.ColWidth(0) = 700
  Me.tablaSUB.ColWidth(1) = 1000
  Me.tablaSUB.ColWidth(2) = 1000
  Me.tablaSUB.ColWidth(3) = 1000

  
  Me.tablaSUB.TextMatrix(1, 0) = "Nombre"
  Me.tablaSUB.TextMatrix(2, 0) = "Area"
  Me.tablaSUB.TextMatrix(0, 2) = "Coef Thissen"
  Me.tablaSUB.TextMatrix(0, 3) = "Pluviometros"
  Me.tablaSUB.TextMatrix(3, 0) = "NC_S"
  Me.tablaSUB.TextMatrix(4, 0) = "NC_N"
  Me.tablaSUB.TextMatrix(5, 0) = "NC_H"
  Me.tablaSUB.TextMatrix(6, 0) = "DRENAJE"
  Me.tablaSUB.TextMatrix(7, 0) = "TC"
  Me.tablaSUB.TextMatrix(8, 0) = "Q base"
  Me.tablaSUB.TextMatrix(9, 0) = "Q A.A" 'CAUDAL AGUAS ARRIBA
  Me.tablaSUB.TextMatrix(10, 0) = "Desfase"
  
  
  Me.tablaSUB.TextMatrix(1, 1) = Me.JUCAR.miColecSub(Indice).NOMBRE_SUBCUENCA
  Me.tablaSUB.TextMatrix(2, 1) = Me.JUCAR.miColecSub(Indice).AREA_KM2
  
  'Me.tablaSUB.Cols = Me.JUCAR.miColecSub(Indice).miColeccionPluviometros.Count + 1
  
  For i = 1 To Me.JUCAR.miColecSub(Indice).miColeccionPluviometros.Count
   Me.tablaSUB.TextMatrix(i + 1, 2) = Me.JUCAR.miColecSub(Indice).miColeccionPluviometros(i).CoefThiessen
   Me.tablaSUB.TextMatrix(i + 1, 3) = Me.JUCAR.miColecSub(Indice).miColeccionPluviometros(i).NombrePluvi
  Next i
  
  Me.tablaSUB.TextMatrix(3, 1) = Me.JUCAR.miColecSub(Indice).NC_S
  Me.tablaSUB.TextMatrix(4, 1) = Me.JUCAR.miColecSub(Indice).NC_N
  Me.tablaSUB.Row = 1
  Me.tablaSUB.COL = 1
  Me.tablaSUB.CellBackColor = RGB(245, 222, 200) 'color de fondo de las celdas
  
  Me.tablaSUB.Row = 4
  Me.tablaSUB.COL = 1
  Me.tablaSUB.CellBackColor = RGB(249, 225, 113) 'color de fondo de las celdas
  
  Me.tablaSUB.TextMatrix(5, 1) = Me.JUCAR.miColecSub(Indice).NC_H
  Me.tablaSUB.TextMatrix(6, 1) = Me.JUCAR.miColecSub(Indice).PUNTO_DRENAJE_ASOCIADO
  Me.tablaSUB.TextMatrix(7, 1) = Me.JUCAR.miColecSub(Indice).tc
  
  Me.tablaSUB.TextMatrix(8, 1) = 0

  Me.tablaSUB.TextMatrix(9, 1) = Me.JUCAR.miColecSub(Indice).AFORO_AGUAS_ARRIBA
  Me.tablaSUB.TextMatrix(10, 1) = Me.JUCAR.miColecSub(Indice).Desfase
  
End Sub



'*************************************************************
'**************** EDIT MSFLEXGRID ****************************
'*************************************************************

Private Sub GridEdit(KeyAscii As Integer)
    ' Position the TextBox over the cell.
    Text1.Left = MSFlexGrid1.CellLeft + MSFlexGrid1.Left
    Text1.Top = MSFlexGrid1.CellTop + MSFlexGrid1.Top
    Text1.Width = MSFlexGrid1.CellWidth
    Text1.Height = MSFlexGrid1.CellHeight
    Text1.Visible = True
    Text1.SetFocus

    Select Case KeyAscii
        Case 0 To Asc(" ")
            Text1.Text = MSFlexGrid1.Text
            Text1.SelStart = Len(Text1.Text)
        Case Else
            Text1.Text = Chr$(KeyAscii)
            Text1.SelStart = 1
    End Select
End Sub



Private Sub Form_Resize()
'iguala el ancho de la imagen
 Me.Picture1.Width = Me.Width - 220
 If Me.Height >= 10110 Then
 Me.Picture1.Height = Me.Picture1.Height + Me.Height - 10110
 End If
End Sub

Private Sub MSFlexGrid1_Scroll()
 Me.MSFlexGrid1.ScrollTrack = True
End Sub



Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            ' Leave the text unchanged.
            Text1.Visible = False
            MSFlexGrid1.SetFocus

        Case vbKeyReturn
            ' Finish editing.
            MSFlexGrid1.SetFocus

        Case vbKeyDown
            ' Move down 1 row.
            MSFlexGrid1.SetFocus
            DoEvents
            If MSFlexGrid1.Row < MSFlexGrid1.Rows - 1 Then
                MSFlexGrid1.Row = MSFlexGrid1.Row + 1
            End If

        Case vbKeyUp
            ' Move up 1 row.
            MSFlexGrid1.SetFocus
            DoEvents
            If MSFlexGrid1.Row > MSFlexGrid1.FixedRows Then
                MSFlexGrid1.Row = MSFlexGrid1.Row - 1
            End If

    End Select
End Sub

'Do not beep on Return or Escape.
Private Sub Text1_KeyPress(KeyAscii As Integer)
    If (KeyAscii = vbKeyReturn) Or _
       (KeyAscii = vbKeyEscape) _
            Then KeyAscii = 0
End Sub



Private Sub MSFlexGrid1_DblClick()
    GridEdit Asc(" ")
End Sub

Private Sub MSFlexGrid1_KeyPress(KeyAscii As Integer)
Dim cadenaCopiada

GridEdit KeyAscii
 
 Clipboard.Clear

If KeyAscii = 3 Then ' si se pulsa control+C lo copia al portapapeles
 cadenaCopiada = Me.MSFlexGrid1.Clip
 Clipboard.SetText cadenaCopiada
End If

End Sub

Private Sub MSFlexGrid1_LeaveCell()
    If Text1.Visible Then
        MSFlexGrid1.Text = Text1.Text
        Text1.Visible = False
    End If
End Sub
Private Sub MSFlexGrid1_GotFocus()
    If Text1.Visible Then
        MSFlexGrid1.Text = Text1.Text
        Text1.Visible = False
    End If
End Sub

'***********************************************
Private Sub GridEdit2(KeyAscii As Integer)
    ' Position the TextBox over the cell.
    Text3.Left = tablaSUB.CellLeft + tablaSUB.Left + Me.Frame2.Left
    Text3.Top = tablaSUB.CellTop + tablaSUB.Top + Frame2.Top
    Text3.Width = tablaSUB.CellWidth
    Text3.Height = tablaSUB.CellHeight
    Text3.Visible = True
    Text3.SetFocus

    Select Case KeyAscii
        Case 0 To Asc(" ")
            Text3.Text = tablaSUB.Text
            Text3.SelStart = Len(Text3.Text)
        Case Else
            Text3.Text = Chr$(KeyAscii)
            Text3.SelStart = 1
    End Select
End Sub



Private Sub Text3_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            ' Leave the text unchanged.
            Text3.Visible = False
            tablaSUB.SetFocus

        Case vbKeyReturn
            ' Finish editing.
            tablaSUB.SetFocus

        Case vbKeyDown
            ' Move down 1 row.
            tablaSUB.SetFocus
            DoEvents
            If tablaSUB.Row < tablaSUB.Rows - 1 Then
                tablaSUB.Row = tablaSUB.Row + 1
            End If

        Case vbKeyUp
            ' Move up 1 row.
            tablaSUB.SetFocus
            DoEvents
            If tablaSUB.Row > tablaSUB.FixedRows Then
                tablaSUB.Row = tablaSUB.Row - 1
            End If

    End Select
End Sub

'Do not beep on Return or Escape.
Private Sub Text3_KeyPress(KeyAscii As Integer)
    If (KeyAscii = vbKeyReturn) Or _
       (KeyAscii = vbKeyEscape) _
            Then KeyAscii = 0
End Sub



Private Sub tablaSUB_DblClick()
    GridEdit2 Asc(" ")
End Sub

Private Sub tablaSUB_KeyPress(KeyAscii As Integer)
    GridEdit2 KeyAscii
End Sub

Private Sub tablaSUB_LeaveCell()
    If Text3.Visible Then
        tablaSUB.Text = Text3.Text
        Text3.Visible = False
    End If
End Sub
Private Sub tablaSUB_GotFocus()
    If Text3.Visible Then
        tablaSUB.Text = Text3.Text
        Text3.Visible = False
    End If
End Sub



Private Sub NuevoCALCULO(NC As Double, col_Llmedia As Integer)

Dim Dimension
Dim numpluvi
Dim nompluvi

Dim tc As Double
Dim NumeroPuntosDibu

Dim Des_fase As Double

 Dimension = mConexion.ColDatos.Count
If Dimension = 0 Then Exit Sub
If (Me.MSFlexGrid1.Rows - 3) > Dimension Then Dimension = (Me.MSFlexGrid1.Rows - 3)
 ReDim vector(1 To Dimension) 'Redimensiona el vector que almacenará la lluvia media
 
 n = 0
 
'ALMACENA LA LLUVIA MEDIA COMO BASE DE LOS CÁLCULOS
 For n = 1 To Dimension
  vector(n) = Val(Me.MSFlexGrid1.TextMatrix(n + 2, col_Llmedia))
 Next n

'ALMACENA el caudal aguas arriba
 If Me.tablaSUB.TextMatrix(9, 1) <> "" Then
  ReDim miCaudalAguasArriba(1 To Dimension) 'Redimensiona el vector que almacenará pero el caudal aguas arriba
  For n = 1 To Dimension
   miCaudalAguasArriba(n) = Val(Me.MSFlexGrid1.TextMatrix(n + 2, 4))
  Next n
  Else
  ReDim miCaudalAguasArriba(0 To 0)
  'Erase miCaudalAguasArriba
 End If
 '**********************
 Area_Cuenca_Km2 = Val(Me.tablaSUB.TextMatrix(2, 1))
 'NC = Val(Me.tablaSUB.TextMatrix(4, 1)) '
 tc = Val(Me.tablaSUB.TextMatrix(7, 1)) 'Me.JUCAR.miColecSub(indiceSub).tc
 Des_fase = Val(Me.tablaSUB.TextMatrix(10, 1))

 Call Inicia_nuevo_calculo 'inicializa vectores

If Factor_Intervalo = 0 Then
  Factor_Intervalo = 1
 Else
  Factor_Intervalo = Val(Me.Textfactor.Text)
End If


If Intervalo_en_Horas = 0 Then
  Intervalo_en_Horas = Val(Me.Textintervalo.Text)
End If

'se llama a la subrutina que calcula el caudal de una cuenca
 SubcuencaSCS vector, Intervalo_en_Horas, Factor_Intervalo, tc, Area_Cuenca_Km2, Val(Me.tablaSUB.TextMatrix(8, 1)), NC, , , Des_fase
 
 
 Me.MSFlexGrid1.Cols = Me.MSFlexGrid1.Cols + 3
 
 'Escribe en la tabla el caudal
 If (Dimension + DimensionHidro) + 2 >= Me.MSFlexGrid1.Rows Then Me.MSFlexGrid1.Rows = (Dimension + DimensionHidro) + 3
 For n = 1 To (Dimension + DimensionHidro)
  'PUNTO_DRENAJE_ASOCIADO
  Me.MSFlexGrid1.TextMatrix(n + 2, Me.MSFlexGrid1.Cols - 1) = Format(miCaudalHidrounitario(n), "0.00")
 Next n
 
 'Escribe en la tabla la lluvia neta, que esta almacenada en miEscorrentia
 For n = 1 To Dimension
  Me.MSFlexGrid1.TextMatrix(n + 2, Me.MSFlexGrid1.Cols - 2) = Format(miEscorrentia(n), "0.00")
 Next n
 
 'Escribe en la tabla la lluvia media que se almacena en vector
 For n = 1 To (Dimension)
  Me.MSFlexGrid1.TextMatrix(n + 2, Me.MSFlexGrid1.Cols - 3) = Format(vector(n), "0.00")
 Next n
 
  'Me.MSFlexGrid2.TextMatrix(2, 2) = Format(Max_Ll_neta, "0.00")
  
  Me.MSFlexGrid1.TextMatrix(2, Me.MSFlexGrid1.Cols - 1) = "Q"
  Me.MSFlexGrid1.TextMatrix(2, Me.MSFlexGrid1.Cols - 2) = "Lln"
  Me.MSFlexGrid1.TextMatrix(2, Me.MSFlexGrid1.Cols - 3) = "Llm"
  
 'Me.MSFlexGrid1.TextMatrix(0, Me.MSFlexGrid1.Cols - 1) = "NC=" & NC
  Me.MSFlexGrid1.TextMatrix(0, Me.MSFlexGrid1.Cols - 1) = Format(Volumen_Total, "0.00")
  Me.MSFlexGrid1.TextMatrix(1, Me.MSFlexGrid1.Cols - 1) = Format(Max_Caudal, "0.00")
'
'  Me.MSFlexGrid2.TextMatrix(0, Me.MSFlexGrid1.Cols - 2) = "LLn"
  Me.MSFlexGrid1.TextMatrix(0, Me.MSFlexGrid1.Cols - 2) = Format(LLuvia_neta_acumulada, "0.00")
  Me.MSFlexGrid1.TextMatrix(1, Me.MSFlexGrid1.Cols - 2) = Format(Max_Ll_neta, "0.00")
'
'  Me.MSFlexGrid2.TextMatrix(0, Me.MSFlexGrid1.Cols - 3) = "LLm"
  Me.MSFlexGrid1.TextMatrix(0, Me.MSFlexGrid1.Cols - 3) = Format(LLuvia_media_acumulada, "0.00")
  Me.MSFlexGrid1.TextMatrix(1, Me.MSFlexGrid1.Cols - 3) = Format(Max_Ll_media, "0.00")

End Sub

Public Sub Rellenafechas()
For k = 3 To Me.MSFlexGrid1.Rows - 1
    If Me.MSFlexGrid1.TextMatrix(k, 1) = "" Then
     Me.MSFlexGrid1.TextMatrix(k, 1) = Format(Me.MSFlexGrid1.TextMatrix(k - 1, 1) + CDate(Val(Textintervalo.Text) / 24), "dd/mm/yyyy hh:nn")
    End If
Next k
End Sub


'************ZOOM*************
Private Sub Picture1_MouseUp(Button As Integer, Shift As Integer, Xup As Single, Yup As Single)
Dim COL
Dim EX1, EX2, EY1, EY2
Dim AnchoSeleccion

If Me.comboSUB.ListIndex <= 0 Then Exit Sub
AnchoSeleccion = Abs(Me.Shape1.Left - Xup)

If AnchoSeleccion < 100 Then
 haceZoom = False
 Exit Sub
End If

If Button = 1 Then
Me.Shape1.Visible = False
'Me.Shape1.Width = Abs(Xup - Me.Shape1.Left)

Me.Picture1.Cls
'migraf.PosicionMarco = 1
EX1 = migraf.Xc_a_X(Me.Shape1.Left)
EX2 = migraf.Xc_a_X(Xup)
EY1 = migraf.Yc_a_Y(Yup)
EY2 = migraf.Yc_a_Y(Me.Shape1.Top)
 
 migraf.X1Etiqueta = EX1 'migraf.Xc_a_X(Me.Shape1.Left)
 migraf.X2Etiqueta = EX2 'migraf.Xc_a_X(Xup - 3000) 'Me.Shape1.Width + Me.Shape1.Left)
 If migraf.X2Etiqueta > X2E Then migraf.X2Etiqueta = X2E
 If EY1 < 0 Then EY1 = 0
 migraf.Y1Etiqueta = EY1 'migraf.Yc_a_Y(Yup)
 migraf.Y2Etiqueta = EY2 'migraf.Yc_a_Y(Me.Shape1.Top)
  
 migraf.DivisionHorizontal = 1
 migraf.DivisionVertical = 1
 migraf.PosicionMarco = 1
 
 
' Dibuja_Una_Serie Me.MSFlexGrid1.Cols - 1, "AZUL"
 migraf.Dibuja
 COL = Me.MSFlexGrid1.Cols - 1
 NumeroPuntosDibu = migraf.NumeroDatos
' For i = 1 To NumeroPuntosDibu
'    Y(i) = Val(Me.MSFlexGrid1.TextMatrix(i + 2, COL))
migraf.ColorDatos = RGB(0, 110, 255)

  For i = 1 To NumeroPuntosDibu
      migraf.Dato = i      'nº de orden del dato
      migraf.XDato = x(i) 'valor de la x
      migraf.YDato = Val(Me.MSFlexGrid1.TextMatrix(i + 2, COL))  'valor de la y
  Next i
 
 
 migraf.Dibuja
 Pinta_Linea_Actual
'Dibuja_Una_Serie Me.MSFlexGrid1.Cols - 1, "AZUL"
End If
haceZoom = False
End Sub


Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, Xx As Single, Yy As Single)
 ' Hace zoom x 2
If Me.comboSUB.ListIndex < 0 Then Exit Sub
If Button = 1 Then
 

 Me.Shape1.Top = Yy
 Me.Shape1.Left = Xx
 'If haceZoom = True Then
    X1E = migraf.X1Etiqueta
    X2E = migraf.X2Etiqueta
    Y1E = migraf.Y1Etiqueta
    Y2E = migraf.Y2Etiqueta
 'End If
 haceZoom = True
Else
Me.Picture1.Cls
'migraf.PosicionMarco = 1
migraf.X1Etiqueta = X1E
migraf.X2Etiqueta = X2E
migraf.Y1Etiqueta = Y1E
migraf.Y2Etiqueta = Y2E
Dibuja_grafica 1, Me.MSFlexGrid1.Cols - 3, Me.MSFlexGrid1.Cols - 2, Me.MSFlexGrid1.Cols - 1
Dibuja_Una_Serie 3, "ROJO"

End If

End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

If haceZoom Then
Me.Shape1.Width = Abs(x - Me.Shape1.Left)
Me.Shape1.Height = Abs(Me.Shape1.Top - y)
 Me.Shape1.Visible = True
End If
'Me.Label1.Caption = migraf.Xc_a_X(X)
'Me.Label2.Caption = migraf.Yc_a_Y(Y) ''
End Sub
'*********** FIN ZOOM*************


Private Sub Pinta_Linea_Actual()
'Pinta una linea vertical que indica el último momento de datos.
Dim x1, x2, y1, y2 As Double
 'X1 = CDate(Format(TextfechaFIN.Text, "yyyymmddhhnn"))
 x1 = CDate(Me.TextfechaFIN.Text) ', "dd/mm/yyy hh:nn"))
 'X1 = CDate(X1)
 x1 = migraf.X_a_Xc(x1)
 x2 = x1
 y1 = migraf.Y_a_Yc(migraf.Y1Etiqueta)
 y2 = migraf.Y_a_Yc(migraf.Y2Etiqueta)
 
 Me.Picture1.DrawWidth = 2
 Me.Picture1.Line (x1, y1)-(x2, y2), RGB(10, 10, 20)
 'Me.Picture1.Line (X1, 100)-(100, Y2), RGB(10, 10, 220)
End Sub



Public Sub Carga_Series()
    Me.Comboaño.Clear
    
    Me.Comboaño.AddItem "2003"
    Me.Comboaño.AddItem "2002"
    Me.Comboaño.AddItem "2001"
    Me.Comboaño.AddItem "2000"
    Me.Comboaño.AddItem "1999"
    Me.Comboaño.AddItem "1998"
    Me.Comboaño.AddItem "1997"
    Me.Comboaño.AddItem "1996"
    
    Me.Comboaño.ListIndex = 0
    
End Sub

Private Function SeleccionSerie(Texto)
'esta función ya no es necesaria pues hay una serie que engloba a todas

 'Select Case Texto
 '   Case "2002"
 '    SeleccionSerie = 0
 '   Case "2001"
 '    SeleccionSerie = 43
 '   Case "2000"
 '    SeleccionSerie = 42
 '   Case "1999"
 '    SeleccionSerie = 44
 '   Case "1998"
 '    SeleccionSerie = 45
 '   Case "1997"
 '    SeleccionSerie = 49
 '   Case "1996"
 '    SeleccionSerie = 50
 '   Case "1995"
 '    SeleccionSerie = 51
 '
 '   Case Else
 '    SeleccionSerie = 0
 'End Select

End Function

'*************************************************************
'**************** EDIT MSFLEXGRID_2 ****************************
'*************************************************************

Private Sub GridEdit3(KeyAscii As Integer)
    ' Position the TextBox over the cell.
    Text2.Left = MSFlexGrid2.CellLeft + MSFlexGrid2.Left
    Text2.Top = MSFlexGrid2.CellTop + MSFlexGrid2.Top
    Text2.Width = MSFlexGrid2.CellWidth
    Text2.Height = MSFlexGrid2.CellHeight
    Text2.Visible = True
    Text2.SetFocus

    Select Case KeyAscii
        Case 0 To Asc(" ")
            Text2.Text = MSFlexGrid2.Text
            Text2.SelStart = Len(Text3.Text)
        Case Else
            Text2.Text = Chr$(KeyAscii)
            Text2.SelStart = 1
    End Select
End Sub

Private Sub MSFlexGrid2_Scroll()
 Me.MSFlexGrid2.ScrollTrack = True
End Sub

Private Sub Text2_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            ' Leave the text unchanged.
            Text2.Visible = False
            MSFlexGrid2.SetFocus

        Case vbKeyReturn
            ' Finish editing.
            MSFlexGrid2.SetFocus

        Case vbKeyDown
            ' Move down 1 row.
            MSFlexGrid2.SetFocus
            DoEvents
            If MSFlexGrid2.Row < MSFlexGrid2.Rows - 1 Then
                MSFlexGrid2.Row = MSFlexGrid2.Row + 1
            End If

        Case vbKeyUp
            ' Move up 1 row.
            MSFlexGrid2.SetFocus
            DoEvents
            If MSFlexGrid2.Row > MSFlexGrid2.FixedRows Then
                MSFlexGrid2.Row = MSFlexGrid2.Row - 1
            End If

    End Select
End Sub

'Do not beep on Return or Escape.
Private Sub Text2_KeyPress(KeyAscii As Integer)
    If (KeyAscii = vbKeyReturn) Or _
       (KeyAscii = vbKeyEscape) _
            Then KeyAscii = 0
End Sub


Private Sub MSFlexGrid2_DblClick()
    GridEdit Asc(" ")
End Sub

Private Sub MSFlexGrid2_KeyPress(KeyAscii As Integer)
Dim cadenaCopiada

GridEdit3 KeyAscii
 
 Clipboard.Clear

If KeyAscii = 3 Then ' si se pulsa control+C lo copia al portapapeles
 cadenaCopiada = Me.MSFlexGrid2.Clip
 Clipboard.SetText cadenaCopiada
End If

End Sub

Private Sub MSFlexGrid2_LeaveCell()
    If Text2.Visible Then
        MSFlexGrid2.Text = Text2.Text
        Text2.Visible = False
    End If
End Sub
Private Sub MSFlexGrid2_GotFocus()
    If Text2.Visible Then
        MSFlexGrid2.Text = Text2.Text
        Text2.Visible = False
    End If
End Sub

'***********************************************
Public Sub GuardaIMG(Nombre As String, Origen As PictureBox, Optional calidad As Integer)
Dim cadena, retval
 cadena = App.Path & "\" & Nombre & ".jpg"
 
 If calidad = 0 Then calidad = 75
 
 SavePicture Origen.Image, "c:\tmp.bmp"
 'La guarda como JPEG
    retval = DIWriteJpg(cadena, calidad, 1)

    If retval = 1 Then  'Si lo hace con exito
    Else                'Si hay un error
        MsgBox "Error en la conversión a jpg"
    End If

    Kill "c:\tmp.bmp" ' borra el fichero temporal bmp

End Sub

Private Sub btnImprimir_Click()
 Dim Pu As New cls_Printer
 Const DistMargenIzq As Integer = 300
On Error GoTo salir
 Me.CD1.CancelError = True
 Me.CD1.Orientation = cdlPortrait
 Me.CD1.PrinterDefault = True
 CD1.Flags = cdlPDReturnIC
 Me.CD1.ShowPrinter

 Printer.Orientation = Me.CD1.Orientation
 Printer.Copies = Me.CD1.Copies
 'Set Printer.hDC = cdlPDReturnIC
 

 Set Pu.ClientPrinter = Printer
' Me.Picture1.Line (10 + (n * AnchoEscala), (Altura - 12))-(10 + ((n + 1) * AnchoEscala), (Altura - 12 - AnchoEscala / 2)), EscalaColor(n), BF
 Pu.PutText DistMargenIzq, 1100, "PROGRAMA MODELOS   -->  SIMULACIÓN DE CAUDALES DE AVENIDA EN CUENCAS SAIH", ptr_RightToPoint 'ptr_LeftToPoint
 Pu.PutImage DistMargenIzq, 1500, Me.Picture1.Image
 'Pu.PutParagraph DistMargenIzq, 2000 + Me.Picture1.Height, Me.Picture1.Width, 1000, _
   " Cálculo de previsiones de datos S.A.I.H."
 
 Pu.Rectangle DistMargenIzq, 1500, Me.Picture1.Width + DistMargenIzq, Me.Picture1.Height + 1500, 2
 'Pu.PintaRectanRelleno 5 + DistMargenIzq, 5 + 1500, 5 + DistMargenIzq + Me.Picture1.TextWidth(form_Origen.text_titulo.Text), 60, 1, rgb(255, 0, 0)
 Pu.PutText DistMargenIzq, 1500 + Me.Picture1.Height + 100, "Leyenda: Barra azul= lluvia medida ; barra naranja= lluvia neta ; línea roja = caudal medido ; línea azul = caudal simulado"
 
 Pu.PutText DistMargenIzq, 1500 + Me.Picture1.Height + 500, "Parámetros de cálculo:"
 Pu.PutText DistMargenIzq, 1500 + Me.Picture1.Height + 1000, "Nombre de la Cuenca = " & Me.tablaSUB.TextMatrix(1, 1)
 Pu.PutText DistMargenIzq, 1500 + Me.Picture1.Height + 1250, "Numero de Curva (seco-normal-humedo) = " & Me.tablaSUB.TextMatrix(3, 1) & " - " & Me.tablaSUB.TextMatrix(4, 1) & " - " & Me.tablaSUB.TextMatrix(5, 1)
 Pu.PutText DistMargenIzq, 1500 + Me.Picture1.Height + 1500, "Tiempo de Concentración = " & Me.tablaSUB.TextMatrix(7, 1) & " h"
 Pu.PutText DistMargenIzq, 1500 + Me.Picture1.Height + 1750, "Intervalo entre datos = " & Me.Textintervalo.Text & " h"
 Pu.PutText DistMargenIzq, 1500 + Me.Picture1.Height + 2000, "----------------------------------------------------------------"
 
 Pu.PutText DistMargenIzq, 1500 + Me.Picture1.Height + 2250, "Lluvia acumulada = " & Me.MSFlexGrid1.TextMatrix(0, 2) & " mm"
 Pu.PutText DistMargenIzq, 1500 + Me.Picture1.Height + 2500, "Lluvia máxima = " & Me.MSFlexGrid1.TextMatrix(1, 2) & " mm/" & Format("0.0", Val(Me.Textintervalo.Text) * 60) & "min"
 Pu.PutText DistMargenIzq, 1500 + Me.Picture1.Height + 2800, "Volumen real = " & Me.MSFlexGrid1.TextMatrix(0, 3) & " Hm3"
 Pu.PutText DistMargenIzq, 1500 + Me.Picture1.Height + 3050, "Caudal maximo real = " & Me.MSFlexGrid1.TextMatrix(1, 3) & " m3/s"
 Pu.PutText DistMargenIzq, 1500 + Me.Picture1.Height + 3350, "Volumen modelo = " & Me.MSFlexGrid1.TextMatrix(0, Me.MSFlexGrid1.Cols - 1) & " Hm3"
 Pu.PutText DistMargenIzq, 1500 + Me.Picture1.Height + 3600, "Caudal maximo modelo = " & Me.MSFlexGrid1.TextMatrix(1, Me.MSFlexGrid1.Cols - 1) & " m3/s"
 
 
 
 
 
 Pu.PutText DistMargenIzq + Me.Picture1.Width, 1500 + Me.Picture1.Height + 3600, "Impreso el " & Format(Now(), "dd/mm/yyyy hh:nn") & "       Área de Explotación - SAIHJúcar", ptr_LeftToPoint
 Pu.Imprimir




salir:
 Me.CD1.CancelError = False

End Sub
