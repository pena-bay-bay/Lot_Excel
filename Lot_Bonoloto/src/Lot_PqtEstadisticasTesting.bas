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
        .TipoMuestra = False
        .Juego = Bonoloto
        .FechaAnalisis = #4/3/2017#          ' Lunes
        .FechaFinal = #4/1/2017#             ' Sabado
        .DiasAnalisis = 30
    End With
    Print_ParametrosMuestra m_objParMuestra
    '
    '   Bonoloto Ok
    '
    Set m_objParMuestra = New ParametrosMuestra
    With m_objParMuestra
        .TipoMuestra = False
        .Juego = Bonoloto
        .FechaAnalisis = #8/21/2017#          ' Lunes
        .FechaFinal = #8/19/2017#             ' Sabado
        .DiasAnalisis = 15
    End With
    Print_ParametrosMuestra m_objParMuestra
    '
    '   Primitiva Ok
    '
    Set m_objParMuestra = New ParametrosMuestra
    With m_objParMuestra
        .TipoMuestra = True
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
        .TipoMuestra = True
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
        .TipoMuestra = True
        .Juego = GordoPrimitiva
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
     
    Ir_A_Hoja ("Salida")
    Set m_objParMuestra = New ParametrosMuestra
    Set m_objMuestra = New Muestra
    
    With m_objParMuestra
        .FechaAnalisis = #2/6/2012#
        .FechaFinal = #2/4/2012#
        .DiasAnalisis = 9
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
Private Sub Print_ParametrosMuestra(obj As ParametrosMuestra)
    Debug.Print "==> Pruebas ParametrosMuestraTest"
    Debug.Print vbTab & "DiasAnalisis      =" & obj.DiasAnalisis
    Debug.Print vbTab & "FechaAnalisis     =" & obj.FechaAnalisis
    Debug.Print vbTab & "FechaFinal        =" & obj.FechaFinal
    Debug.Print vbTab & "FechaInicial      =" & obj.FechaInicial
    Debug.Print vbTab & "Juego             =" & obj.Juego
    Debug.Print vbTab & "NumeroSorteos     =" & obj.NumeroSorteos
    Debug.Print vbTab & "ResgistroAnalisis =" & obj.ResgistroAnalisis
    Debug.Print vbTab & "ResgistroFinal    =" & obj.ResgistroFinal
    Debug.Print vbTab & "ResgistroInicial  =" & obj.ResgistroInicial
    On Error Resume Next
    Debug.Print vbTab & "ToString()         " & obj.ToString()
    Debug.Print vbTab & "Validar()          " & obj.Validar()
    Debug.Print vbTab & "GetMensaje()       " & obj.GetMensaje()
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
Private Sub Print_Bola(obj As BolaV2)
    Debug.Print "==> Pruebas BolaTest"
    Debug.Print vbTab & "Apariciones           =" & obj.Apariciones
    Debug.Print vbTab & "Ausencias             =" & obj.Ausencias
    Debug.Print vbTab & "ColorFrecuencia       =" & obj.ColorFrecuencia
    Debug.Print vbTab & "ColorProbabilidad     =" & obj.ColorProbabilidad
    Debug.Print vbTab & "ColorTiempoMedio      =" & obj.ColorTiempoMedio
    Debug.Print vbTab & "DesviacionTiempoMedio =" & obj.DesviacionTiempoMedio
    Debug.Print vbTab & "FechaAnalisis         =" & obj.FechaAnalisis
    Debug.Print vbTab & "FechasAparicion       =" & obj.FechasAparicion.Count
    Debug.Print vbTab & "FechaUltimaAparicion  =" & obj.FechaUltimaAparicion
    Debug.Print vbTab & "Frecuencia            =" & obj.Frecuencias.Count
    Debug.Print vbTab & "Juego                 =" & obj.Juego
    Debug.Print vbTab & "MaximoTiempo          =" & obj.MaximoTiempo
    Debug.Print vbTab & "Mediana               =" & obj.Mediana
    Debug.Print vbTab & "MinimoTiempo          =" & obj.MinimoTiempo
    Debug.Print vbTab & "Moda                  =" & obj.Moda
    Debug.Print vbTab & "Numero                =" & obj.Numero.Valor
    Debug.Print vbTab & "Probabilidad          =" & obj.Probabilidad
    Debug.Print vbTab & "ProbabilidadFrecuencia =" & obj.ProbabilidadFrecuencia
    Debug.Print vbTab & "probabilidadTiempo    =" & obj.ProbabilidadTiempo
    Debug.Print vbTab & "ProximaFechaAparicion =" & obj.ProximaFechaAparicion
    Debug.Print vbTab & "RegistroAnalisis      =" & obj.RegistroAnalisis
    Debug.Print vbTab & "RegistroAparicion     =" & obj.RegistroAparicion
    Debug.Print vbTab & "Tendencia             =" & obj.Tendencia
    Debug.Print vbTab & "TiempoMedio           =" & obj.TiempoMedio
    Debug.Print vbTab & "TipoAusencia          =" & obj.TipoAusencia
    Debug.Print vbTab & "TotalNumeros          =" & obj.TotalNumeros
    Debug.Print vbTab & "UltimoRegistro        =" & obj.UltimoRegistro
End Sub

'------------------------------------------------------------------------------*
' Procedimiento  : BomboTest
' Fecha          : 06/05/2018
' Propósito      : Pruebas unitarias de la clase Bombo
'------------------------------------------------------------------------------*
'
Public Sub BomboTest()
    Dim obj As BomboV2
    
 On Error GoTo BomboTest_Error
    '
    '   PU01: Bombo en vacio (Carga, Giros, Extraccion)
    '
    Set obj = New BomboV2
    Print_Bombo obj
    obj.Cargar
    obj.Girar
    Debug.Print "Extraemos bolas de bombo vacio:" & obj.ExtraerBolas(1)
    Set obj = Nothing
    '
    '   PU02: Bombo Bonoloto (Carga, Giros, Extraccion)
    '
    Set obj = New BomboV2
    With obj
        .Juego = Bonoloto
        .NumGiros = 100
        .TipoGiros = lotGiros
    End With
    Print_Bombo obj
    obj.Cargar
    obj.Girar
    Debug.Print "Extraemos bolas de bombo :" & obj.ExtraerBolas(7)
    Set obj = Nothing
    '
    '   PU04: Bombo Euromillones (Carga, Giros, Extraccion)
    '
    Set obj = New BomboV2
    With obj
        .Juego = Euromillones
        .NumGiros = 100
        .TipoGiros = lotTiempo
        .TiempoGiro = #12:00:01 AM#   ' 1 segundo
    End With
    Print_Bombo obj
    obj.Cargar
    obj.Girar
    Debug.Print "Extraemos bolas de bombo :" & obj.ExtraerBolas(5)
    Set obj = Nothing

    '
    '   PU05: Bombo Euromillones Estrellas (Carga, Giros, Extraccion)
    '
    Set obj = New BomboV2
    With obj
        .Juego = Euromillones
        .TipoBombo = 2
        .NumGiros = 50
        .TipoGiros = lotGiros
        .TiempoGiro = #12:00:01 AM#   ' 1 segundo
    End With
    Print_Bombo obj
    obj.Cargar
    obj.Girar
    Debug.Print "Extraemos bolas de bombo :" & obj.ExtraerBolas(2)
    Set obj = Nothing

    '
    '   PU06: Bombo Gordo primitiva (Carga, Giros, Extraccion)
    '
    Set obj = New BomboV2
    With obj
        .Juego = GordoPrimitiva
        .TipoBombo = 1
        .NumGiros = 50
        .TipoGiros = lotGiros
    End With
    obj.Cargar
    Print_Bombo obj
    obj.Girar
    Debug.Print "Extraemos bolas de bombo :" & obj.ExtraerBolas(6)
    Set obj = Nothing
    '
    '
    '   PU07: Bombo Gordo primitiva - Clave (Carga, Giros, Extraccion)
    '
    Set obj = New BomboV2
    With obj
        .Juego = GordoPrimitiva
        .TipoBombo = 2
        .NumGiros = 50
        .TipoGiros = lotGiros
    End With
    obj.Cargar
    Print_Bombo obj
    obj.Girar
    Debug.Print "Extraemos bolas de bombo :" & obj.ExtraerBolas(1)
    Set obj = Nothing
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
                
    Set obj = New BomboV2
    With obj
        .Juego = Euromillones
        .TipoBombo = 2
        .NumGiros = 50
        .TipoGiros = lotGiros
        .Cargar
        .ProbabilidadesBolas = mProb
    End With
    Print_Bombo obj
    obj.Girar
    Debug.Print "Extraemos bolas de bombo :" & obj.ExtraerBolas(2)
    Set obj = Nothing
    '
    '   PU09: Bombo Bonoloto Reinicio
    '
    Set obj = New BomboV2
    With obj
        .Juego = Bonoloto
        .TipoBombo = 1
        .NumGiros = 50
        .TipoGiros = lotGiros
    End With
    obj.Cargar
    Print_Bombo obj
    obj.Girar
    Debug.Print "Extraemos bolas de bombo :" & obj.ExtraerBolas(10)
    obj.Reiniciar
    Print_Bombo obj
    obj.Girar
    Debug.Print "Extraemos bolas de bombo :" & obj.ExtraerBolas(10)
    
    Set obj = Nothing
    
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

Private Sub Print_Bombo(obj As BomboV2)
    Debug.Print "==> Pruebas BomboTest"
    Debug.Print vbTab & "Bolas.Count           =" & obj.Bolas.Count
    Debug.Print vbTab & "Cargado               =" & obj.Cargado
    Debug.Print vbTab & "Juego                 =" & obj.Juego
    Debug.Print vbTab & "NumBolas              =" & obj.NumBolas
    Debug.Print vbTab & "NumGiros              =" & obj.NumGiros
    Debug.Print vbTab & "ProbabilidadesBolas   =" & UBound(obj.ProbabilidadesBolas)
    Debug.Print vbTab & "Situacion             =" & obj.Situacion
    Debug.Print vbTab & "TiempoGiro            =" & obj.TiempoGiro
    Debug.Print vbTab & "TipoBombo             =" & obj.TipoBombo
    Debug.Print vbTab & "TipoGiros             =" & obj.TipoGiros
End Sub




' *===========(EOF): Lot_PqtEstadisticasTesting.bas
