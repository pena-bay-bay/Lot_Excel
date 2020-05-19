Attribute VB_Name = "Lot_TestingEstadisticas"
' *============================================================================*
' *
' *     Fichero    : Lot_TestingEstadisticas
' *
' *     Autor      : Carlos Almela Baeza
' *     Creación   : mar, 17/01/2012 23:45
' *     Versión    : 1.0
' *     Propósito  : Pruebas unitarias de las clases del paquete Estadísticas
' *
' *
' *============================================================================*
Option Explicit
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


