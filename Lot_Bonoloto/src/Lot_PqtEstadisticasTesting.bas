Attribute VB_Name = "Lot_PqtEstadisticasTesting"
' *============================================================================*
' *
' *     Fichero    : Lot_PqtEstadisticasTesting.bas
' *
' *     Autor      : Carlos Almela Baeza
' *     Creación   : mar, 17/01/2012 23:45
' *     Versión    : 1.0
' *     Propósito  : Pruebas unitarias de las clases del paquete Estadísticas
' *
' *
' *============================================================================*
Option Explicit
Option Base 0


'---------------------------------------------------------------------------------------
' Procedure : ParametrosMuestraTest
' Author    : Charly
' Date      : 19/03/2012
' Purpose   : Probar la clase ParametrosMuestra
'             Test: Bonoloto Ok
'                   Primitiva Ok
'                   Gordo Ok
'                   Euromillon Ok
'                   Gordo NOK
'                   Fecha ini + Fecha Fin
'                   Fecha Fin + registros
'                   Fecha Analisis Ok
'                   Fecha Analisis NOK
'---------------------------------------------------------------------------------------
'
Private Sub ParametrosMuestraTest()
    Dim m_objParMuestra As ParametrosMuestra
    '
    '   Bonoloto Ok
    '
    Set m_objParMuestra = New ParametrosMuestra
    With m_objParMuestra
        .Juego = Bonoloto
        .FechaAnalisis = #4/3/2017#          ' Lunes
        .FechaFinal = #4/1/2017#             ' Sabado
        .FechaInicial = #3/25/2017#          ' Sabado
    End With
    Print_ParametrosMuestra m_objParMuestra
    '
    '   Bonoloto Ok
    '
    Set m_objParMuestra = New ParametrosMuestra
    With m_objParMuestra
        .Juego = Bonoloto
        .FechaAnalisis = #8/21/2017#          ' Lunes
        .FechaFinal = #8/19/2017#             ' Sabado
        .FechaInicial = #7/22/2017#           ' Sabado
    End With
    Print_ParametrosMuestra m_objParMuestra
    '
    '   Primitiva Ok
    '
    Set m_objParMuestra = New ParametrosMuestra
    With m_objParMuestra
        .Juego = LoteriaPrimitiva
        .FechaAnalisis = #4/8/2017#          'Sabado
        .FechaFinal = #4/6/2017#             'Jueves
        .NumeroSorteos = 10
    End With
    Print_ParametrosMuestra m_objParMuestra
    '
    '   Euromillon Ok
    '
    Set m_objParMuestra = New ParametrosMuestra
    With m_objParMuestra
        .Juego = Euromillones
        .FechaAnalisis = #3/31/2017#         'Viernes
        .FechaFinal = #3/28/2017#            'Martes
        .NumeroSorteos = 10
    End With
    Print_ParametrosMuestra m_objParMuestra
    '
    '   Gordo Ok
    '
    Set m_objParMuestra = New ParametrosMuestra
    With m_objParMuestra
        .Juego = gordoPrimitiva
        .FechaAnalisis = #4/16/2017#         'Domingo
        .FechaFinal = #4/9/2017#             'Domingo
        .NumeroSorteos = 10
    End With
    Print_ParametrosMuestra m_objParMuestra
    
End Sub
'---------------------------------------------------------------------------------------
' Procedure : MuestraTest
' Author    : Charly
' Date      : 19/03/2012
' Purpose   : Probar la clase Muestra
'---------------------------------------------------------------------------------------
'
Private Sub MuestraTest()
    Dim m_objParMuestra     As ParametrosMuestra
    Dim m_objMuestra        As Muestra
    Dim m_objRg             As Range        'rango de datos
    Dim m_objBd             As New BdDatos  'base de datos
     
    m_objBd.Ir_A_Hoja ("Salida")
    Set m_objParMuestra = New ParametrosMuestra
    Set m_objMuestra = New Muestra
    
    With m_objParMuestra
        .FechaAnalisis = #2/6/2012#
        .FechaFinal = #2/4/2012#
        .FechaInicial = #1/25/2012#
    End With
    '
    '       Calcula la Muestra
    '
    '   obtiene el rango con los datos comprendido entre las dos fechas
    '
    Set m_objRg = m_objBd.Resultados_Fechas(m_objParMuestra.FechaInicial, _
                                            m_objParMuestra.FechaFinal)
    '
    '   se lo pasa al constructor de la clase y obtiene las estadisticas para cada bola
    '
    Set m_objMuestra.ParametrosMuestra = m_objParMuestra
    m_objMuestra.Constructor m_objRg, JUEGO_DEFECTO

    
    Pintar_Muestra m_objMuestra
End Sub



'---------------------------------------------------------------------------------------
' Procedure : Print_ParametrosMuestra
' Author    : Charly
' Date      : 15/04/2017
' Purpose   : Probar la clase ParametrosMuestra
'---------------------------------------------------------------------------------------
'
Private Sub Print_ParametrosMuestra(Obj As ParametrosMuestra)
    Debug.Print "==> Pruebas ParametrosMuestraTest"
    Debug.Print vbTab & "DiasAnalisis      =" & Obj.DiasAnalisis
    Debug.Print vbTab & "FechaAnalisis     =" & Obj.FechaAnalisis
    Debug.Print vbTab & "FechaFinal        =" & Obj.FechaFinal
    Debug.Print vbTab & "FechaInicial      =" & Obj.FechaInicial
    Debug.Print vbTab & "Juego             =" & Obj.Juego
    Debug.Print vbTab & "NumeroSorteos     =" & Obj.NumeroSorteos
    Debug.Print vbTab & "ResgistroAnalisis =" & Obj.ResgistroAnalisis
    Debug.Print vbTab & "ResgistroFinal    =" & Obj.ResgistroFinal
    Debug.Print vbTab & "ResgistroInicial  =" & Obj.ResgistroInicial
    On Error Resume Next
    Debug.Print vbTab & "ToString()         " & Obj.ToString()
    Debug.Print vbTab & "Validar()          " & Obj.Validar()
    Debug.Print vbTab & "GetMensaje()       " & Obj.GetMensaje()
    Err.Clear
End Sub

'---------------------------------------------------------------------------------------
' Procedure : BolaTest
' Author    : Charly
' Date      : 22/04/2017
' Purpose   : Probar la clase Bola
'---------------------------------------------------------------------------------------
'
Public Sub BolaTest()
    Dim mBola As BolaV2
    Dim mNumero As Numero
    Dim mTupla As TuplaAparicion
    
 On Error GoTo BolaTest_Error
    Set mNumero = New Numero
    mNumero.Valor = 48
    Set mBola = New BolaV2
    '
    '   Datos de la bola
    '
    Set mBola.Numero = mNumero
    mBola.Juego = Bonoloto
    mBola.TipoBola = 1  'Numero
    mBola.TotalNumeros = 98
    mBola.FechaAnalisis = #4/18/2017#
    mBola.RegistroAnalisis = 1731
'    mBola.FechaAnalisis = #4/17/2017#
'    mBola.RegistroAnalisis = 1730
    
    '
    '
    '
    Set mTupla = New TuplaAparicion
    mTupla.FechaAparicion = #4/7/2017#
    mTupla.NumeroRegistro = 1722
    mTupla.OrdenAparicion = 3
    mBola.Add mTupla
    '
    Set mTupla = New TuplaAparicion
    mTupla.FechaAparicion = #4/8/2017#
    mTupla.NumeroRegistro = 1723
    mTupla.OrdenAparicion = 7
    mBola.Add mTupla
    '
    Set mTupla = New TuplaAparicion
    mTupla.FechaAparicion = #4/12/2017#
    mTupla.NumeroRegistro = 1726
    mTupla.OrdenAparicion = 1
    mBola.Add mTupla
    '
    Set mTupla = New TuplaAparicion
    mTupla.FechaAparicion = #4/15/2017#
    mTupla.NumeroRegistro = 1729
    mTupla.OrdenAparicion = 7
    mBola.Add mTupla
    mBola.Actualizar
    
    Print_Bola mBola

    On Error GoTo 0
BolaTest__CleanExit:
    Exit Sub
            
BolaTest_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    '   Audita el error
    Call HandleException(ErrNumber, ErrDescription, "Lot_TestingEstadisticas.BolaTest", ErrSource)
    '   Informa del error
    Call MsgBox(ErrDescription, vbCritical Or vbSystemModal, ThisWorkbook.Name)
    Call Trace("CERRAR")
End Sub
'---------------------------------------------------------------------------------------
' Procedure : Print_ParametrosMuestra
' Author    : Charly
' Date      : 15/04/2017
' Purpose   : Probar la clase ParametrosMuestra
'---------------------------------------------------------------------------------------
'
Private Sub Print_Bola(Obj As BolaV2)
    Debug.Print "==> Pruebas BolaTest"
    Debug.Print vbTab & "Apariciones           =" & Obj.Apariciones
    Debug.Print vbTab & "Ausencias             =" & Obj.Ausencias
    Debug.Print vbTab & "ColorFrecuencia       =" & Obj.ColorFrecuencia
    Debug.Print vbTab & "ColorProbabilidad     =" & Obj.ColorProbabilidad
    Debug.Print vbTab & "ColorTiempoMedio      =" & Obj.ColorTiempoMedio
    Debug.Print vbTab & "DesviacionTiempoMedio =" & Obj.DesviacionTiempoMedio
    Debug.Print vbTab & "FechaAnalisis         =" & Obj.FechaAnalisis
    Debug.Print vbTab & "FechasAparicion       =" & Obj.FechasAparicion.Count
    Debug.Print vbTab & "FechaUltimaAparicion  =" & Obj.FechaUltimaAparicion
    Debug.Print vbTab & "Frecuencia            =" & Obj.Frecuencias.Count
    Debug.Print vbTab & "Juego                 =" & Obj.Juego
    Debug.Print vbTab & "MaximoTiempo          =" & Obj.MaximoTiempo
    Debug.Print vbTab & "Mediana               =" & Obj.Mediana
    Debug.Print vbTab & "MinimoTiempo          =" & Obj.MinimoTiempo
    Debug.Print vbTab & "Moda                  =" & Obj.Moda
    Debug.Print vbTab & "Numero                =" & Obj.Numero.Valor
    Debug.Print vbTab & "Probabilidad          =" & Obj.Probabilidad
    Debug.Print vbTab & "ProbabilidadFrecuencia =" & Obj.ProbabilidadFrecuencia
    Debug.Print vbTab & "probabilidadTiempo    =" & Obj.ProbabilidadTiempo
    Debug.Print vbTab & "ProximaFechaAparicion =" & Obj.ProximaFechaAparicion
    Debug.Print vbTab & "RegistroAnalisis      =" & Obj.RegistroAnalisis
    Debug.Print vbTab & "RegistroAparicion     =" & Obj.RegistroAparicion
    Debug.Print vbTab & "Tendencia             =" & Obj.Tendencia
    Debug.Print vbTab & "TiempoMedio           =" & Obj.TiempoMedio
    Debug.Print vbTab & "TipoAusencia          =" & Obj.TipoAusencia
    Debug.Print vbTab & "TotalNumeros          =" & Obj.TotalNumeros
    Debug.Print vbTab & "UltimoRegistro        =" & Obj.UltimoRegistro
End Sub

'------------------------------------------------------------------------------*
' Procedimiento  : BomboTest
' Fecha          : 06/05/2018
' Propósito      : Pruebas unitarias de la clase Bombo
'------------------------------------------------------------------------------*
'
Public Sub BomboTest()
    Dim Obj As BomboV2
    
 On Error GoTo BomboTest_Error
    '
    '   PU01: Bombo en vacio (Carga, Giros, Extraccion)
    '
    Set Obj = New BomboV2
    Print_Bombo Obj
    Obj.Cargar
    Obj.Girar
    Debug.Print "Extraemos bolas de bombo vacio:" & Obj.ExtraerBolas(1)
    Set Obj = Nothing
    '
    '   PU02: Bombo Bonoloto (Carga, Giros, Extraccion)
    '
    Set Obj = New BomboV2
    With Obj
        .Juego = Bonoloto
        .NumGiros = 100
        .TipoGiros = lotGiros
    End With
    Print_Bombo Obj
    Obj.Cargar
    Obj.Girar
    Debug.Print "Extraemos bolas de bombo :" & Obj.ExtraerBolas(7)
    Set Obj = Nothing
    '
    '   PU04: Bombo Euromillones (Carga, Giros, Extraccion)
    '
    Set Obj = New BomboV2
    With Obj
        .Juego = Euromillones
        .NumGiros = 100
        .TipoGiros = lotTiempo
        .TiempoGiro = #12:00:01 AM#   ' 1 segundo
    End With
    Print_Bombo Obj
    Obj.Cargar
    Obj.Girar
    Debug.Print "Extraemos bolas de bombo :" & Obj.ExtraerBolas(5)
    Set Obj = Nothing

    '
    '   PU05: Bombo Euromillones Estrellas (Carga, Giros, Extraccion)
    '
    Set Obj = New BomboV2
    With Obj
        .Juego = Euromillones
        .TipoBombo = 2
        .NumGiros = 50
        .TipoGiros = lotGiros
        .TiempoGiro = #12:00:01 AM#   ' 1 segundo
    End With
    Print_Bombo Obj
    Obj.Cargar
    Obj.Girar
    Debug.Print "Extraemos bolas de bombo :" & Obj.ExtraerBolas(2)
    Set Obj = Nothing

    '
    '   PU06: Bombo Gordo primitiva (Carga, Giros, Extraccion)
    '
    Set Obj = New BomboV2
    With Obj
        .Juego = gordoPrimitiva
        .TipoBombo = 1
        .NumGiros = 50
        .TipoGiros = lotGiros
    End With
    Obj.Cargar
    Print_Bombo Obj
    Obj.Girar
    Debug.Print "Extraemos bolas de bombo :" & Obj.ExtraerBolas(6)
    Set Obj = Nothing
    '
    '
    '   PU07: Bombo Gordo primitiva - Clave (Carga, Giros, Extraccion)
    '
    Set Obj = New BomboV2
    With Obj
        .Juego = gordoPrimitiva
        .TipoBombo = 2
        .NumGiros = 50
        .TipoGiros = lotGiros
    End With
    Obj.Cargar
    Print_Bombo Obj
    Obj.Girar
    Debug.Print "Extraemos bolas de bombo :" & Obj.ExtraerBolas(1)
    Set Obj = Nothing
    '
    '   PU08: Bombo Estrellas Cargado (Carga, Giros, Extraccion)
    '
    Dim mProb(11, 1) As Variant
    mProb(0, 0) = 1:    mProb(0, 1) = 0.025396825
    mProb(1, 0) = 2:    mProb(1, 1) = 0.026984127
    mProb(2, 0) = 3:    mProb(2, 1) = 0.00100475
    mProb(3, 0) = 4:    mProb(3, 1) = 0.028571429
    mProb(4, 0) = 5:    mProb(4, 1) = 0.026984127
    mProb(5, 0) = 6:    mProb(5, 1) = 0.05048
    mProb(6, 0) = 7:    mProb(6, 1) = 0.03015873
    mProb(7, 0) = 8:    mProb(7, 1) = 0.038095238
    mProb(8, 0) = 9:    mProb(8, 1) = 0.5084588
    mProb(9, 0) = 10:   mProb(9, 1) = 0.04069
    mProb(10, 0) = 11: mProb(10, 1) = 0.01689
    mProb(11, 0) = 12: mProb(11, 1) = 0.025396825
                
    Set Obj = New BomboV2
    With Obj
        .Juego = Euromillones
        .TipoBombo = 2
        .NumGiros = 50
        .TipoGiros = lotGiros
        .ProbabilidadesBolas = mProb
    End With
    Obj.Cargar
    Print_Bombo Obj
    Obj.Girar
    Debug.Print "Extraemos bolas de bombo :" & Obj.ExtraerBolas(2)
    Set Obj = Nothing
    '
    '   PU09: Bombo Bonoloto Reinicio
    '
    Set Obj = New BomboV2
    With Obj
        .Juego = Bonoloto
        .TipoBombo = 1
        .NumGiros = 50
        .TipoGiros = lotGiros
    End With
    Obj.Cargar
    Print_Bombo Obj
    Obj.Girar
    Debug.Print "Extraemos bolas de bombo :" & Obj.ExtraerBolas(10)
    Obj.Reiniciar
    Print_Bombo Obj
    Obj.Girar
    Debug.Print "Extraemos bolas de bombo :" & Obj.ExtraerBolas(10)
    
    Set Obj = Nothing
    
    '
    '   PU10: Bombo Bonoloto tipo giros
    '
    
 On Error GoTo 0
BomboTest__CleanExit:
    Exit Sub
            
BomboTest_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    '   Audita el error
    Call HandleException(ErrNumber, ErrDescription, "Lot_PqtEstadisticasTesting.BomboTest", ErrSource)
    '   Informa del error
    Call MsgBox(ErrDescription, vbCritical Or vbSystemModal, ThisWorkbook.Name)
    Call Trace("CERRAR")
End Sub


'Private Function GetIndex(datTotal As Integer) As Integer
'    Static b_rand As Boolean
'    If Not b_rand Then          'La primera vez que se ejecuta
'        b_rand = True           'la función RND se ceba la
'        Randomize               'la semilla
'    End If
'    GetIndex = Int(datTotal * Rnd)
'    If GetIndex = 0 Then
'        GetIndex = datTotal
'    ElseIf GetIndex >= datTotal Then
'        GetIndex = 1
'    End If
'End Function
'
'Public Sub GetIdexTest()
'    Dim i As Integer
'
'    For i = 1 To 100
'        Debug.Print GetIndex(5)
'    Next i
'End Sub

Private Sub Print_Bombo(Obj As BomboV2)
    Debug.Print "==> Pruebas BomboTest"
    Debug.Print vbTab & "Bolas.Count           =" & Obj.Bolas.Count
    Debug.Print vbTab & "Cargado               =" & Obj.Cargado
    Debug.Print vbTab & "Juego                 =" & Obj.Juego
    Debug.Print vbTab & "NumBolas              =" & Obj.NumBolas
    Debug.Print vbTab & "NumGiros              =" & Obj.NumGiros
    Debug.Print vbTab & "ProbabilidadesBolas   =" & UBound(Obj.ProbabilidadesBolas)
    Debug.Print vbTab & "Situacion             =" & Obj.Situacion
    Debug.Print vbTab & "TiempoGiro            =" & Obj.TiempoGiro
    Debug.Print vbTab & "TipoBombo             =" & Obj.TipoBombo
    Debug.Print vbTab & "TipoGiros             =" & Obj.TipoGiros
End Sub




' *===========(EOF): Lot_PqtEstadisticasTesting.bas
