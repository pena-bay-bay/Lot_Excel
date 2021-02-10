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
Private mMaxNum As Integer                  ' Numero máximo
Private mSetFiltros As Boolean              ' Indicador de aplicar filtro
Private mMaxCiclos As Integer               ' Maximo numero de ciclos para
                                            ' obtener filtros
Private mNCiclos As Integer                 ' indicador de filtros
Private mBombo As BomboV2                   ' clase Bombo para obtener bolas
Private mStr As String                      ' texto temporal
Private mParEstad As Variant                ' parametros de la estadistica
Private oOrden      As TipoOrdenacion       ' Tipo de ordenación de la Std
Private oAgrup      As TipoAgrupacion       ' Tipo de agrupación de la Std
Private bAscendente As Boolean              ' Orden ascendente / Descendente
Private vArray      As Variant              ' Matriz de probabilidades
Private vArray2     As Variant              ' Matriz de probabilidades aux
Private i           As Integer              ' indice
Private j           As Integer              ' indice

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
        Set mFilters = mMetodo.Filtros
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
            mNum = IIf(mNum > mMaxNum, 1, mNum)
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
            Debug.Print "Ciclo:", mNCiclos, mSuge.Combinacion.ToString(True)
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
' Fecha          : lu., 25/may/2020 19:26:39
' Propósito      : generar una sugerencia del método Bombo
'------------------------------------------------------------------------------*
Private Sub getMtdBombo()
    
  On Error GoTo getMtdBombo_Error
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
        Set mFilters = mMetodo.Filtros
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
        '   Inicializamos el bombo
        '
        Randomize
        Set mBombo = New BomboV2
        With mBombo
            .Juego = JUEGO_DEFECTO
            .TipoBombo = 1
            .NumGiros = 10
            .Cargar
            .Girar
        End With
        '
        '   Extraemos una combinación
        '
        mStr = mBombo.ExtraerBolas(mPronosticos)
        mSuge.Combinacion.Texto = mStr
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
' Fecha          : ma., 26/may/2020 17:36:05
' Propósito      : generar una sugerencia del método Bombo
'------------------------------------------------------------------------------*
Private Sub getMtdBomboCargado()
  
  On Error GoTo getMtdBomboCargado_Error
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
        Set mFilters = mMetodo.Filtros
    Else
        mSetFiltros = False
    End If
    '
    '   Obtenemos la matriz estadistica según criterio
    '
    mParEstad = mMuestra.GetMatrizBolas(mMetodo.CriteriosOrdenacion, mMetodo.SentidoOrdenacion)
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
        '   Inicializamos el bombo
        '
        Randomize
        Set mBombo = New BomboV2
        With mBombo
            .Juego = JUEGO_DEFECTO
            .TipoBombo = 1
            .NumGiros = 10
            .Cargar
            .ProbabilidadesBolas = mParEstad
            .Girar
        End With
        '
        '   Extraemos una combinación
        '
        mStr = mBombo.ExtraerBolas(mPronosticos)
        mSuge.Combinacion.Texto = mStr
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
getMtdBomboCargado__CleanExit:
    Exit Sub
getMtdBomboCargado_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    Call HandleException(ErrNumber, ErrDescription, "RealizarSugerencia.getMtdBomboCargado", ErrSource)
    Err.Raise ErrNumber, "RealizarSugerencia.getMtdBomboCargado", ErrDescription
End Sub


'------------------------------------------------------------------------------*
' Procedimiento  : getMtdEstadistico
' Fecha          : ma., 26/may/2020 17:35:41
' Propósito      : generar una sugerencia del método Bombo cargado estadisticas
'------------------------------------------------------------------------------*
Private Sub getMtdEstadistico()
  
  On Error GoTo getMtdEstadistico_Error
    i = 0
    j = 0
    '
    '   Obtenemos los parametros del metodo
    '
    With mMetodo
        oOrden = .CriteriosOrdenacion
        oAgrup = .CriteriosAgrupacion
        bAscendente = .SentidoOrdenacion
    End With
    '
    '   Obtenemos la matriz de Numeros ordenada por el valor
    '
    vArray = mMuestra.GetMatrizBolas(oOrden, bAscendente)
    '
    '
    '
    If oAgrup <> grpSinDefinir Then
        '
        ' Atender a la agrupación
        '   Crear matriz de 2 *2
        '
        vArray2 = AgrupaMatriz(vArray, oAgrup)
        '
        '
        '
        Do Until mSuge.Combinacion.Count = mPronosticos
            '
            '   Obtiene el Numero de la matriz
            '
            mNum = vArray2(i, j)
            '
            '   Si el  Numero no existe
            '
            If Not mSuge.Combinacion.Contiene(mNum) Then
                '
                '   añadimos el numero a la combinación
                '
                Set oNum = New Numero
                oNum.Valor = mNum
                mSuge.Combinacion.Add oNum
            End If
            '
            '   Cambia de grupo
            '
            If i >= UBound(vArray2, 1) Then
                i = 0
                j = j + 1
            Else
                i = i + 1
            End If
        Loop
        
    Else
        Do Until mSuge.Combinacion.Count = mPronosticos
            '
            '   Obtiene el Numero de la matriz
            '
            mNum = vArray(i, 0)
            '
            '   Si el  Numero no existe
            '
            If Not mSuge.Combinacion.Contiene(mNum) Then
                '
                '   añadimos el numero a la combinación
                '
                Set oNum = New Numero
                oNum.Valor = mNum
                mSuge.Combinacion.Add oNum
            End If
            i = i + 1
        Loop
    End If

  On Error GoTo 0
getMtdEstadistico__CleanExit:
    Exit Sub
getMtdEstadistico_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    Call HandleException(ErrNumber, ErrDescription, "RealizarSugerencia.getMtdEstadistico", ErrSource)
    Err.Raise ErrNumber, "RealizarSugerencia.getMtdEstadistico", ErrDescription
End Sub


'------------------------------------------------------------------------------*
' Procedimiento  : getMtdEstadCombi
' Fecha          :
' Propósito      : Sugerencia estadistica sujeta a filtro
'------------------------------------------------------------------------------*
Private Sub getMtdEstadCombi()
  On Error GoTo getMtdEstadCombi_Error
    
    i = 0
    j = 0
    '
    '   Obtenemos los parametros del metodo
    '
    With mMetodo
        oOrden = .CriteriosOrdenacion
        oAgrup = .CriteriosAgrupacion
        bAscendente = .SentidoOrdenacion
    End With
    '
    '   Obtenemos la matriz de Numeros ordenada por el valor
    '
    vArray = mMuestra.GetMatrizBolas(oOrden, bAscendente)
    '
    '
    '
    If oAgrup <> grpSinDefinir Then
        '
        ' Atender a la agrupación
        '   Crear matriz de 2 *2
        '
        vArray2 = AgrupaMatriz(vArray, oAgrup)
        '
        '
        '
        Do Until mSuge.Combinacion.Count = mPronosticos
            '
            '   Obtiene el Numero de la matriz
            '
            mNum = vArray2(i, j)
            '
            '   Si el  Numero no existe
            '
            If Not mSuge.Combinacion.Contiene(mNum) Then
                '
                '   añadimos el numero a la combinación
                '
                Set oNum = New Numero
                oNum.Valor = mNum
                mSuge.Combinacion.Add oNum
            End If
            '
            '   Cambia de grupo
            '
            If i >= UBound(vArray2, 1) Then
                i = 0
                j = j + 1
            Else
                i = i + 1
            End If
        Loop
        
    Else
        Do Until mSuge.Combinacion.Count = mPronosticos
            '
            '   Obtiene el Numero de la matriz
            '
            mNum = vArray(i, 0)
            '
            '   Si el  Numero no existe
            '
            If Not mSuge.Combinacion.Contiene(mNum) Then
                '
                '   añadimos el numero a la combinación
                '
                Set oNum = New Numero
                oNum.Valor = mNum
                mSuge.Combinacion.Add oNum
            End If
            i = i + 1
        Loop
    End If
    '
    '   #TODO: incluir filtros en la selección de la sugerencia
    '
    
  On Error GoTo 0
getMtdEstadCombi__CleanExit:
    Exit Sub
getMtdEstadCombi_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    Call HandleException(ErrNumber, ErrDescription, "RealizarSugerencia.getMtdEstadCombi", ErrSource)
    Err.Raise ErrNumber, "RealizarSugerencia.getMtdEstadCombi", ErrDescription
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
    Dim mDatos As BdDatos
    
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
            '
            '   Definimos los parámetros de la muestra
            '
            Set mParam = New ParametrosMuestra
            mParam.TipoMuestra = mMetodo.TipoMuestra
            mParam.FechaAnalisis = datFecha
            If mMetodo.TipoMuestra Then
                mParam.NumeroSorteos = mMetodo.NumeroSorteos
            Else
                mParam.DiasAnalisis = mMetodo.DiasAnalisis
            End If
            '
            '   Calculamos la muestra
            '
            On Error Resume Next
            CalcularMuestra mParam
            If Err.Number = 100 Then
                '
                '   Obtenemos una muestra estadistica de los
                '   ultimos 100 registros de la base de datos
                '
                Set mDatos = New BdDatos
                mParam.TipoMuestra = True
                mParam.FechaFinal = mDatos.UltimoResultado
                mParam.NumeroSorteos = 100
                CalcularMuestra mParam
            End If
            Set mSuge.Parametros = mParam
            On Error GoTo GetSugerencia_Error
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



'------------------------------------------------------------------------------*
' Procedimiento  : CalcularMuestra
' Fecha          :
' Propósito      :
' Parámetros     :
'------------------------------------------------------------------------------*
Private Sub CalcularMuestra(datParam As ParametrosMuestra)
    Dim mRg             As Range                  ' rango de datos
    Dim mBd             As New BdDatos            ' base de datos
  On Error GoTo CalcularMuestra_Error
    '
    '   obtiene el rango con los datos comprendido entre las dos fechas
    '
    Set mRg = mBd.GetSorteosInFechas(datParam.PeriodoDatos)
    '
    '   se lo pasa al constructor de la clase y obtiene las estadisticas para cada bola
    '
    Set mMuestra.ParametrosMuestra = datParam
    Select Case JUEGO_DEFECTO
        Case LoteriaPrimitiva, Bonoloto:
            mMuestra.Constructor mRg, ModalidadJuego.LP_LB_6_49
        
        Case GordoPrimitiva:
            mMuestra.Constructor mRg, ModalidadJuego.GP_5_54
        
        Case Euromillones:
            mMuestra.Constructor mRg, ModalidadJuego.EU_5_50
            
    End Select

On Error GoTo 0
CalcularMuestra__CleanExit:
    Exit Sub
CalcularMuestra_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    Call HandleException(ErrNumber, ErrDescription, "RealizarSugerencia.CalcularMuestra", ErrSource)
    Err.Raise ErrNumber, "RealizarSugerencia.CalcularMuestra", ErrDescription
End Sub

'---------------------------------------------------------------------------------------
' Procedure : AgrupaMatriz
' Author    : Charly
' Date      : 15/04/2012
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Function AgrupaMatriz(vMat As Variant, vAgrup As TipoAgrupacion) As Variant
    Dim a_Matriz()          As Variant
    Dim n                   As New Numero
    Dim i                   As Integer
    Dim a_cont()            As Variant
    Dim j                   As Integer
    
  On Error GoTo AgrupaMatriz_Error

    Select Case vAgrup
        
        Case grpDecenas:
            If JUEGO_DEFECTO = Bonoloto Or _
            JUEGO_DEFECTO = LoteriaPrimitiva Then
                ReDim a_Matriz(4, 10)
                ReDim a_cont(4, 1)
            Else
                ReDim a_Matriz(5, 10)
                ReDim a_cont(5, 1)
            End If
            
        Case grpSeptenas:
            ReDim a_Matriz(7, 7)
            ReDim a_cont(7, 1)
            
        Case grpParidad:
            ReDim a_Matriz(1, 25)
            ReDim a_cont(1, 1)
            
        Case grpPeso:
            ReDim a_Matriz(1, 25)
            ReDim a_cont(1, 1)
            
        Case grpTerminacion:
            ReDim a_Matriz(9, 4)
            ReDim a_cont(9, 1)

    End Select
    
    '
    '   Redimimensionar contador y Probabilidad
    '
    
    For i = 0 To UBound(vMat) - 1
        '
        '   Concretamos el Numero
        '
        n.Valor = vMat(i, 0)
        '
        '
        '
        Select Case vAgrup
            Case grpDecenas:
                j = n.Decena - 1
                a_Matriz(j, a_cont(j, 0)) = n.Valor
                a_cont(j, 0) = a_cont(j, 0) + 1
                a_cont(j, 1) = a_cont(j, 1) + vMat(i, 1)
                
            Case grpSeptenas:
                j = n.Septena - 1
                a_Matriz(j, a_cont(j, 0)) = n.Valor
                a_cont(j, 0) = a_cont(j, 0) + 1
                a_cont(j, 1) = a_cont(j, 1) + vMat(i, 1)
    
            Case grpParidad:
                j = IIf(n.Paridad = LT_PAR, 0, 1)
                a_Matriz(j, a_cont(j, 0)) = n.Valor
                a_cont(j, 0) = a_cont(j, 0) + 1
                a_cont(j, 1) = a_cont(j, 1) + vMat(i, 1)
    
            Case grpPeso:
                j = IIf(n.Peso = LT_BAJO, 0, 1)
                a_Matriz(j, a_cont(j, 0)) = n.Valor
                a_cont(j, 0) = a_cont(j, 0) + 1
                a_cont(j, 1) = a_cont(j, 1) + vMat(i, 1)
    
            Case grpTerminacion:
                j = n.Terminacion
                a_Matriz(j, a_cont(j, 0)) = n.Valor
                a_cont(j, 0) = a_cont(j, 0) + 1
                a_cont(j, 1) = a_cont(j, 1) + vMat(i, 1)
        
        End Select
    Next i
    
    AgrupaMatriz = a_Matriz

   On Error GoTo 0
   Exit Function

AgrupaMatriz_Error:
   Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
   ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
   Call HandleException(ErrNumber, ErrDescription, "RealizarSugerencia.AgrupaMatriz")
   Err.Raise ErrNumber, ErrSource, ErrDescription
End Function
'' *===========(EOF): RealizarSugerencia
