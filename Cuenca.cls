VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Cuenca"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Clase que almacena las variables de las subcuencas en una Cuenca

Option Explicit

Public miColecSub As New Collection
Private miNombre As String 'nombre de la cuenca
Private miIndSub As Integer

Public Sub LEE_FICH_SUB(rutaNomfichero As String)
'Esta sub lee el fichero que define las subcuencas y asocia los valores de las series
'Se almacenan las subcuencas como un clase.
'El fichero que define las cuencas tiene el siguiente formato:
'*****************************************
'NOMBRE_SUBCUENCA: ARQUILLO
'PLUVIOMETROS:
'        1 1E02MVVR3   // coeficiente_thiessen variable
'AREA_KM2:   698
'TC_HORAS:   9.97
'Po_SECO_NORMAL_HUMEDO: 60, 46.89, 30
'o NC_SECO_NORMAL_HUMEDO: 12, 34, 45
'PUNTO_DRENAJE_ASOCIADO: 2QAM2VV3CDR
'AFORO_AGUAS_ARRIBA:  4E01EQS1MVVR
'DESFASE:  10

Dim Linea
Dim nf%
Dim Npluvi As Integer
Dim CoeficienteThiessen As Double
Dim VariablePluviometro As String
Dim ncorte, ncorte1
Dim newSub As New SUBCUENCA


nf = FreeFile
'Nsubcuencas = 0
Open rutaNomfichero For Input As nf

        Do While Not EOF(nf)
            Line Input #nf, Linea
            Linea = Trim(Linea)
Inicio:
          If Left(Linea, 17) = "NOMBRE_SUBCUENCA:" Then
            Linea = Trim(Linea)
            Set newSub = Nothing
            newSub.NOMBRE_SUBCUENCA = Trim(Mid$(Linea, 18, Len(Linea) - 17))
            
          ElseIf Linea = "PLUVIOMETROS:" Then
           Npluvi = 0
           Line Input #nf, Linea
           Linea = (Trim(Linea))
            Do While IsNumeric(Left(Linea, 1))
                Npluvi = Npluvi + 1
                ncorte = BuscaBlanco(Linea, " ")
                    CoeficienteThiessen = Val(Mid(Linea, 1, ncorte - 1))
                    VariablePluviometro = Trim(Mid(Linea, ncorte, Len(Linea)))
                    
                    newSub.AñadirPluvi VariablePluviometro, CoeficienteThiessen, Str(Npluvi)
                Line Input #nf, Linea
                Linea = Trim(Linea)
            Loop
            GoTo Inicio
            
          ElseIf Left(UCase(Linea), 8) = "AREA_KM2" Then
                newSub.AREA_KM2 = Val(Trim(Mid(Linea, 10, Len(Linea))))
               
          ElseIf Left(UCase(Linea), 8) = "TC_HORAS" Then
                newSub.tc = Val(Val(Mid(Linea, 10, Len(Linea))))
                    
          ElseIf Left(UCase(Linea), 21) = "NC_SECO_NORMAL_HUMEDO" Then
           Linea = Right(Linea, Len(Linea) - 22)
           
                ncorte = BuscaBlanco(Linea, ",")
                
                newSub.NC_S = Val(Mid(Linea, 1, ncorte))
                
                Linea = Right(Linea, Len(Linea) - ncorte)
                ncorte = BuscaBlanco(Linea, ",")
                newSub.NC_N = Val(Mid(Linea, 1, ncorte))
                
                Linea = Right(Linea, Len(Linea) - ncorte)
                newSub.NC_H = Val(Linea)
                
          ElseIf Left(UCase(Linea), 21) = "PO_SECO_NORMAL_HUMEDO" Then
            Linea = Right(Linea, Len(Linea) - 22)
            Linea = Trim(Linea)
            
            ncorte = BuscaBlanco(Linea, ",")
            newSub.NC_S = Format((5000 / (Val(Mid(Linea, 1, ncorte)) + 50)), "0.00")
            
            Linea = Right(Linea, Len(Linea) - ncorte)
            ncorte = BuscaBlanco(Linea, ",")
            newSub.NC_N = Format((5000 / (Val(Mid(Linea, 1, ncorte) + 50))), "0.00")
            
            Linea = Right(Linea, Len(Linea) - ncorte)
            newSub.NC_H = Format((5000 / (Val(Linea) + 50)), "0.00")
                
          ElseIf Left(UCase(Linea), 22) = "PUNTO_DRENAJE_ASOCIADO" Then
                
                newSub.PUNTO_DRENAJE_ASOCIADO = Trim(Mid(Linea, 24, Len(Linea)))
                
          ElseIf Left(UCase(Linea), 18) = "AFORO_AGUAS_ARRIBA" Then
                newSub.AFORO_AGUAS_ARRIBA = Trim(Mid(Linea, 20, Len(Linea)))
          
          ElseIf Left(UCase(Linea), 7) = "DESFASE" Then
                newSub.Desfase = Trim(Mid(Linea, 9, Len(Linea)))
            'Añade la subcuenca a la colesccion
            
            Me.AñadirSUB newSub
          End If
        
        
        Loop

Close nf
End Sub


Public Property Get Nombre() As String  'Nombre de la subcuencas
    Nombre = miNombre
End Property

Public Property Let Nombre(vNewValue As String)
    miNombre = vNewValue
End Property


Public Sub AñadirSUB(Sub_C As SUBCUENCA)
    
    miColecSub.Add Sub_C, Sub_C.NOMBRE_SUBCUENCA
 
End Sub

Public Property Get NumSubCuencas() As Integer  'indice de la subcuenca seleccionada
    NumSubCuencas = Me.miColecSub.Count
End Property

Public Property Get IndSub() As Integer 'Para el indice de la serie en curso
    If IndSub = 0 Then IndSub = 1
    IndSub = miIndSub
End Property

Public Property Let IndSub(vNewValue As Integer) 'Para el indice de la serie en curso
    If vNewValue > Me.NumSubCuencas Then
        miIndSub = Me.NumSubCuencas
    ElseIf vNewValue < 1 Then
        miIndSub = 1
    Else
        miIndSub = vNewValue
    End If
End Property


Private Sub Class_Initialize()
 Set miColecSub = Nothing
End Sub

Private Function BuscaBlanco(Linea, caracterBuscado As String, Optional PosicionINI As Integer) As Integer
'Halla en qué posición se encuentra el siguiente espacio en blanco de una frase (linea)

Dim k%
Dim caracter As String * 1
If PosicionINI = 0 Then PosicionINI = 1
For k = PosicionINI To Len(Linea)
caracter = Mid(Linea, k, 1)
If caracter = caracterBuscado Then
 BuscaBlanco = k
 Exit Function
End If
Next k
 
 BuscaBlanco = -1
End Function
