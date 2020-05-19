VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RealizarSugerencia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' *============================================================================*
' *
' *     Fichero    : RealizarSugerencia.cls
' *
' *     Autor      : Carlos Almela Baeza
' *     Creación   : ju., 23/abr/2020 17:29:04
' *     Versión    : 1.0
' *     Propósito  : Motor que genera sugerencias para un método definido
' *
' *============================================================================*
Option Explicit
Option Base 0
'
'--- Variables Privadas -------------------------------------------------------*
Private mMetodo As Metodo                   ' Metodo de cálculo
Private mMuestra As Muestra                 ' Muestra estadística
Private mSuge As Sugerencia                 ' Sugerencia
Private mPronosticos As Integer             ' Pronosticos de la Sugerencia
Private mFilters As FiltrosCombinacion      ' Filtros de combinación
Private mNum As Integer                     ' Numero seleccionado para sugerecia
Private oNum As Numero                      ' Objeto número para sugerencia

'--- Constantes ---------------------------------------------------------------*
'--- Mensajes -----------------------------------------------------------------*
'--- Errores ------------------------------------------------------------------*
'--- Propiedades --------------------------------------------------------------*
'--- Métodos Privados ---------------------------------------------------------*
Private Sub Class_Initialize()
    Set mFilters = New FiltrosCombinacion
    Set mMetodo = New Metodo
    Set mMuestra = New Muestra
    Set mSuge = New Sugerencia
End Sub

Private Sub Class_Terminate()
    Set mFilters = Nothing
    Set mMetodo = Nothing
    Set mMuestra = Nothing
    Set mSuge = Nothing
End Sub



'------------------------------------------------------------------------------*
' Procedimiento  : getMtdAleatorio
' Fecha          : lu., 11/may/2020 18:25:19
' Propósito      : generar una sugerencia del método aleatorio
'------------------------------------------------------------------------------*
Private Sub getMtdAleatorio()
    Dim mMaxNum As Integer
    Dim mSetFiltros As Boolean
    Dim mMaxCiclos As Integer
    Dim mNCiclos As Integer
    
  On Error GoTo getMtdAleatorio_Error
    '
    '   Establecemos el número máximo de bolas del bombo
    '
    Select Case mMetodo.ModalidadJuego
        Case LP_LB_6_49: mMaxNum = 49
        Case GP_5_54: mMaxNum = 54
        Case EU_5_50: mMaxNum = 50
        Case EU_2_12: mMaxNum = 12
    End Select
    '
    '   Identificamos si se aplican filtros a la sugerencia
    '
    If mMetodo.Filtros.Count > 0 Then
        mSetFiltros = True
    Else
        mSetFiltros = False
    End If
    '
    '   Establecemos los limites del bucle
    '
    mMaxCiclos = 100
    mNCiclos = 0
    '
    '   Bucle principal de generación
    '
    While (mNCiclos < mMaxCiclos)
        '
        '   inicializamos la combinación de la sugerencia
        '
        mSuge.Combinacion.Clear
        '
        '   Inicializamos la semilla
        '
        Randomize
        '
        '   Generamos una combinación aleatoria válida
        '
        Do Until mSuge.Combinacion.Count = mPronosticos
            '
            '   Función aleatoria
            '
            mNum = Round(mMaxNum * Rnd()) + 1
            '
            '   Si el numero no está en la combinación
            '
            If Not mSuge.Combinacion.Contiene(mNum) Then
                '
                '   añadimos el numero a la combinación
                '
                Set oNum = New Numero
                oNum.Valor = mNum
                mSuge.Combinacion.Add oNum
            End If
        Loop
        '
        '   Si hay filtros comprobamos si es válido
        '
        If mSetFiltros Then
            If mFilters.Evaluar(mSuge.Combinacion) Then
                mNCiclos = mMaxCiclos
            Else
                mNCiclos = mNCiclos + 1
            End If
        Else
            mNCiclos = mMaxCiclos
        End If
    Wend
  
  On Error GoTo 0
getMtdAleatorio__CleanExit:
    Exit Sub
getMtdAleatorio_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    Call HandleException(ErrNumber, ErrDescription, "RealizarSugerencia.getMtdAleatorio", ErrSource)
    Err.Raise ErrNumber, "RealizarSugerencia.getMtdAleatorio", ErrDescription
End Sub


'------------------------------------------------------------------------------*
' Procedimiento  : getMtdBombo
' Fecha          :
' Propósito      : generar una sugerencia del método Bombo
'------------------------------------------------------------------------------*
Private Sub getMtdBombo()
    
  On Error GoTo getMtdBombo_Error
  
  On Error GoTo 0
getMtdBombo__CleanExit:
    Exit Sub
getMtdBombo_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    Call HandleException(ErrNumber, ErrDescription, "RealizarSugerencia.getMtdBombo", ErrSource)
    Err.Raise ErrNumber, "RealizarSugerencia.getMtdBombo", ErrDescription
End Sub



'------------------------------------------------------------------------------*
' Procedimiento  : getMtdBomboCargado
' Fecha          :
' Propósito      : generar una sugerencia del método Bombo
'------------------------------------------------------------------------------*
Private Sub getMtdBomboCargado()
    
  On Error GoTo getMtdBomboCargado_Error
  
  On Error GoTo 0
getMtdBomboCargado__CleanExit:
    Exit Sub
getMtdBomboCargado_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    Call HandleException(ErrNumber, ErrDescription, "RealizarSugerencia.getMtdBomboCargado", ErrSource)
    Err.Raise ErrNumber, "RealizarSugerencia.getMtdBomboCargado", ErrDescription
End Sub


Private Sub getMtdEstadistico()

End Sub


Private Sub getMtdEstadCombi()

End Sub

'--- Métodos Públicos ---------------------------------------------------------*
'------------------------------------------------------------------------------*
' Procedimiento  : GetSugerencia
' Fecha          :
' Propósito      :
' Parámetros     : Metodo de cálculo
'                  fecha de análisis de la muestra
'------------------------------------------------------------------------------*
Public Function GetSugerencia(datMetodo As Metodo, _
                              Optional datFecha As Date) As Sugerencia
    Dim mParam As ParametrosMuestra
  On Error GoTo GetSugerencia_Error
    '
    '   Si no nos pasan la fecha de sugerencia  asumimos la actual
    '
    If IsMissing(datFecha) Then
        datFecha = Date
    End If
    '
    '   Asignamos el metodo a la variable interna
    '
    Set mMetodo = datMetodo
    '
    '   Inicializamos la sugerencia
    '
    Set mSuge = New Sugerencia
    With mSuge
        Set .Metodo = mMetodo
        .Modalidad = mMetodo.ModalidadJuego
        .Entidad.Origen = 1
    End With
    '
    '   Parametros de la muestra: Bombo Cargado, Estadistica,
    '                             Estadistica combinación
    '
    If mMetodo.TipoProcedimiento = mtdBomboCargado Or _
       mMetodo.TipoProcedimiento = mtdEstadCombinacion Or _
       mMetodo.TipoProcedimiento = mtdEstadistico Then
            Set mParam = New ParametrosMuestra
            mParam.FechaAnalisis = datFecha
            If mMetodo.TipoMuestra Then
                mParam.NumeroSorteos = mMetodo.NumeroSorteos
            Else
                mParam.DiasAnalisis = mMetodo.DiasAnalisis
            End If
            CalcularMuestra mParam
    End If
    '
    '   Evaluar el número de pronosticos
    '
    mPronosticos = mMetodo.Pronosticos
    '
    '
    '   Seleccionamos el procedimiento
    '
    Select Case datMetodo.TipoProcedimiento
        Case mtdAleatorio: getMtdAleatorio
        Case mtdBombo: getMtdBombo
        Case mtdBomboCargado: getMtdBomboCargado
        Case mtdEstadistico: getMtdEstadistico
        Case mtdEstadCombinacion: getMtdEstadCombi
    End Select
    '
    '   Devolvemos la sugerencia
    '
    Set GetSugerencia = mSuge
  On Error GoTo 0
GetSugerencia__CleanExit:
    Exit Function
GetSugerencia_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    Call HandleException(ErrNumber, ErrDescription, "RealizarSugerencia.GetSugerencia", ErrSource)
    Err.Raise ErrNumber, "RealizarSugerencia.GetSugerencia", ErrDescription
End Function
'------------------------------------------------------------------------------*
' Procedimiento  : GetSugerencia
' Fecha          :
' Propósito      :
' Parámetros     : Metodo de cálculo
'                  fecha de análisis de la muestra
'------------------------------------------------------------------------------*
Public Function GetSugerencias(datMetodo As Metodo, _
                               datNumSugerencias As Integer, _
                      Optional datFecha As Date) As Sugerencias
  On Error GoTo GetSugerencias_Error
    '
    '
    Err.Raise ERR_TODO, "RealizarSugerencia.GetSugerencias", MSG_TODO
    '
  On Error GoTo 0
GetSugerencias__CleanExit:
    Exit Function
GetSugerencias_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    Call HandleException(ErrNumber, ErrDescription, "RealizarSugerencia.GetSugerencias", ErrSource)
    Err.Raise ErrNumber, "RealizarSugerencia.GetSugerencias", ErrDescription
End Function


Private Sub CalcularMuestra(datParam As ParametrosMuestra)
    Dim mRg             As Range                  ' rango de datos
    Dim mBd             As New BdDatos            ' base de datos
  On Error GoTo CalcularMuestra_Error
    '
    '   obtiene el rango con los datos comprendido entre las dos fechas
    '
    Set mRg = mBd.Resultados_Fechas(datParam.FechaInicial, _
                                    datParam.FechaFinal)
    '
    '   se lo pasa al constructor de la clase y obtiene las estadisticas para cada bola
    '
    Set mMuestra.ParametrosMuestra = datParam
    mMuestra.Constructor mRg, JUEGO_DEFECTO
    
On Error GoTo 0
CalcularMuestra__CleanExit:
    Exit Sub
CalcularMuestra_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    Call HandleException(ErrNumber, ErrDescription, "RealizarSugerencia.CalcularMuestra", ErrSource)
    Err.Raise ErrNumber, "RealizarSugerencia.CalcularMuestra", ErrDescription
End Sub
'' *===========(EOF): RealizarSugerencia
