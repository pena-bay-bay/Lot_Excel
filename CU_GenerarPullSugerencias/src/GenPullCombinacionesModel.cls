VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GenPullCombinacionesModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' *============================================================================*
' *
' *     Fichero    : GenPullCombinacionesModel.cls
' *
' *     Autor      : Carlos Almela Baeza
' *     Creación   : lu., 21/dic/2020 18:56:50
' *     Versión    : 1.0
' *     Propósito  :
' *
' *============================================================================*
Option Explicit
Option Base 0
'
'--- Variables Privadas -------------------------------------------------------*
Private mMatrizNumeros As Variant
Private mNums As Collection
Private mFiltros As Variant
Private mFltrs As FiltrosCombinacion
Private mError As Long
Private mSorteos As Variant
Private mSrts As Sorteos
Private mCombinacionGanadora As Variant
Private mCmbGndra As Sorteo
Private mNumSugerencia As Integer
Private mTotNumCombinar As Integer
Private mTotCombinaciones As Long
Private mTotEvaluadas As Long
Private mTotRegComprobados As Long
Private mTotRegPremiados As Long
Private mTotCoste As Currency
Private mTotImportePremios As Currency
Private mFase As String
Private mFrm  As frmProgreso

'--- Constantes ---------------------------------------------------------------*
Private mErrorValidacion                    As Long       ' Control de errores de validación
Private Const ERR_NUMSUGOBLIGATIRIO         As Long = 1   ' Numeros de sugerencia obligatorios
Private Const ERR_MTRZOBLIGATORIA           As Long = 2   ' Matriz de números Obligatoria

'--- Mensajes -----------------------------------------------------------------*
Private Const MSG_ERRORESVALIDACION         As String = "Los parametros del proceso no cumple las siguientes validaciones:" & vbCrLf
Private Const MSG_NUMSUGOBLIGATIRIO         As String = "* El numero de sugerencias es obligatorio."
Private Const MSG_MTRZOBLIGATORIA           As String = "* El conjunto de numeros a combinar es obligatorio."



'--- Propiedades --------------------------------------------------------------*
Public Property Get MatrizNumeros() As Variant
    MatrizNumeros = mMatrizNumeros
End Property

Public Property Let MatrizNumeros(ByVal vNewValue As Variant)
    mMatrizNumeros = vNewValue
    If IsArray(vNewValue) Then
        SetNumerosToColleccion vNewValue
    End If
End Property

Public Property Get NumerosSugerencia() As Variant
    NumerosSugerencia = mNumSugerencia
End Property

Public Property Let NumerosSugerencia(ByVal vNewValue As Variant)
    If IsNumeric(vNewValue) Then
        mNumSugerencia = vNewValue
    End If
End Property

Public Property Get TotalNumerosCombinar() As Variant
    TotalNumerosCombinar = mTotNumCombinar
End Property

Public Property Let TotalNumerosCombinar(ByVal vNewValue As Variant)
    If IsNumeric(vNewValue) Then
        mTotNumCombinar = vNewValue
    End If
End Property

Public Property Get TotalCombinaciones() As Variant
    TotalCombinaciones = mTotCombinaciones
End Property

Public Property Let TotalCombinaciones(ByVal vNewValue As Variant)
    mTotCombinaciones = vNewValue
End Property

Public Property Get TotRegComprobados() As Long
    TotRegComprobados = mTotRegComprobados
End Property

Public Property Get TotRegPremiados() As Long
    TotRegPremiados = mTotRegPremiados
End Property

Public Property Get TotCoste() As Currency
    TotCoste = mTotCoste
End Property
Public Property Get TotImportePremios() As Currency
    TotImportePremios = mTotImportePremios
End Property
Public Property Get Filtros() As Variant
    Filtros = mFiltros
End Property

Public Property Let Filtros(ByVal vNewValue As Variant)
    mFiltros = vNewValue
    If IsArray(vNewValue) Then
        SetFiltrosToColleccion vNewValue
    End If
End Property

Public Property Get Sorteos() As Variant
    Sorteos = mSorteos
End Property

Public Property Let Sorteos(ByVal vNewValue As Variant)
    mSorteos = vNewValue
    If IsArray(vNewValue) Then
        SetSorteosToCollection vNewValue
    End If
End Property

Public Property Get CombinacionGanadora() As Variant
    CombinacionGanadora = mCombinacionGanadora
End Property

Public Property Let CombinacionGanadora(ByVal vNewValue As Variant)
    mCombinacionGanadora = vNewValue
    If IsArray(vNewValue) Then
        SetCmbnGndrToSorteo vNewValue
    End If
End Property

Public Property Get FaseProceso() As String
    FaseProceso = mFase
End Property

Public Property Let FaseProceso(vNewValue As String)
    mFase = vNewValue
End Property



'--- Métodos Privados ---------------------------------------------------------*
Private Sub Class_Initialize()
    mNumSugerencia = 0
    mTotNumCombinar = 0
    Set mFltrs = New FiltrosCombinacion
    Set mNums = New Collection
    Set mSrts = New Sorteos
    Set mCmbGndra = New Sorteo
End Sub


Private Sub Class_Terminate()
    Set mFltrs = Nothing
    Set mNums = Nothing
    Set mSrts = Nothing
    Set mCmbGndra = Nothing
End Sub



'------------------------------------------------------------------------------*
' Procedimiento  : SetNumerosToColleccion
' Fecha          : ma., 03/sep/2019 19:11:14
' Propósito      : Trasladar los numéros a una colección
' Parámetros     : Matriz de Numeros
'------------------------------------------------------------------------------*
Private Sub SetNumerosToColleccion(vNewValue As Variant)
    Dim i As Integer
    Dim mNum As Numero
    Dim mKey As String

On Error GoTo SetNumerosToColleccion_Error

    For i = 0 To UBound(vNewValue)
        '
        '   Creamos el Numero
        '
        Set mNum = New Numero
        '
        '   Asignamos el filtro
        '
        mNum.Valor = CInt(vNewValue(i))
        mKey = CStr(i)
        '
        '   Agregamos a la colección
        '
        If mNums.Count = 0 Then
            mNums.Add mNum, mKey
        Else
            mNums.Add mNum, mKey, , mNums.Count
        End If
    Next i
          
  On Error GoTo 0
SetNumerosToColleccion_CleanExit:
    Exit Sub
SetNumerosToColleccion_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    Call HandleException(ErrNumber, ErrDescription, "GenPullCombinacionesModel.SetNumerosToColleccion", ErrSource)
    Err.Raise ErrNumber, "GenPullCombinacionesModel.SetNumerosToColleccion", ErrDescription
End Sub




'------------------------------------------------------------------------------*
' Procedimiento  : SetFiltrosToColleccion
' Fecha          : ma., 03/sep/2019 19:10:24
' Propósito      : Trasladar los filtros seleccionados a una colección
' Parámetros     : Matriz de filtros
'------------------------------------------------------------------------------*
Private Sub SetFiltrosToColleccion(vNewValue As Variant)
    Dim i As Integer
    Dim j As Integer
    Dim mFilter As FiltroCombinacion
    Dim mTmpFilter As FiltroCombinacion
    Dim mTmp As String
    Dim mFltrNew As Boolean

On Error GoTo SetFiltrosToColleccion_Error
    For i = 0 To UBound(vNewValue)
        '
        '   Creamos el filtro
        '
        Set mFilter = New FiltroCombinacion
        '
        '   Asignamos el filtro
        '
        mTmp = vNewValue(i)
        '
        '  Si el filtro esta vacio se sale del bucle
        '
        If Len(mTmp) = 0 Then
            Exit For
        End If
        '
        '   Desempaquetamos el filtro
        '
        mFilter.Parse mTmp
        '
        '   Agregamos a la colección
        '
        mFltrs.Add mFilter
    Next i
          
  On Error GoTo 0
SetFiltrosToColleccion_CleanExit:
    Exit Sub
SetFiltrosToColleccion_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    Call HandleException(ErrNumber, ErrDescription, "GenPullCombinacionesModel.SetFiltrosToColleccion", ErrSource)
    Err.Raise ErrNumber, "GenPullCombinacionesModel.SetFiltrosToColleccion", ErrDescription
End Sub




'------------------------------------------------------------------------------*
' Procedimiento  : SetSorteosToCollection
' Fecha          : do., 31/ene/2021 12:42:53
' Propósito      : Trasladar los sorteos definidos para evaluar a una colección
' Parámetros     : Matriz de sorteos
'------------------------------------------------------------------------------*
Private Sub SetSorteosToCollection(vNewValue As Variant)
    Dim mSort   As Sorteo
    Dim i       As Integer
    Dim mStr    As String
    
  On Error GoTo SetSorteosToCollection_Error
    '
    '   Creamos la colección de sorteos
    '
    Set mSorteos = New Sorteos
    '
    '   Recorremos la matriz
    '
    For i = 1 To UBound(vNewValue)
        '
        '
        If IsNumeric(vNewValue(i, 1)) Then
            Set mSort = New Sorteo
            With mSort
                .Id = CInt(vNewValue(i, 1))
                .NumeroSorteo = vNewValue(i, 2)
                .Fecha = CDate(vNewValue(i, 3))
                .Dia = vNewValue(i, 4)
            '    .Semana = CInt(vNewValue(i, 5))
                mStr = vNewValue(i, 6) & "-" & _
                       vNewValue(i, 7) & "-" & _
                       vNewValue(i, 8) & "-" & _
                       vNewValue(i, 9) & "-" & _
                       vNewValue(i, 10) & "-" & _
                       vNewValue(i, 11)
                .Combinacion.Texto = mStr
                .Complementario = CInt(vNewValue(i, 12))
                .Reintegro = CInt(vNewValue(i, 13))
                .Ordenado = IIf(vNewValue(i, 14) = "Si", True, False)
                If Not IsEmpty(vNewValue(i, 15)) Then
                    .Premios.Parse (vNewValue(i, 15))
                End If
            End With
            mSorteos.Add mSort
        End If
    Next i
  
  On Error GoTo 0
SetSorteosToCollection_CleanExit:
    Exit Sub
SetSorteosToCollection_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    Call HandleException(ErrNumber, ErrDescription, "GenPullCombinacionesModel.SetSorteosToCollection", ErrSource)
    Err.Raise ErrNumber, "GenPullCombinacionesModel.SetSorteosToCollection", ErrDescription
End Sub




'------------------------------------------------------------------------------*
' Procedimiento  : SetCmbnGndrToSorteo
' Fecha          : do., 31/ene/2021 12:44:14
' Propósito      : Traslada el sorteo a comprobar a un objeto Sorteo
' Parámetros     : Matriz de valores de un sorteo
'------------------------------------------------------------------------------*
Private Sub SetCmbnGndrToSorteo(vNewValue As Variant)
    Dim mStr    As String
  
  On Error GoTo SetCmbnGndrToSorteo_Error
    '
    '   Creamos el objeto Sorteo
    '
    Set mCombinacionGanadora = New Sorteo
    '
    '   Recorremos la matriz
    '
    If IsNumeric(vNewValue(1, 1)) Then
        '
        '   Asignamos valores
        '
        With mCombinacionGanadora
            .Id = CInt(vNewValue(1, 1))
            .NumeroSorteo = vNewValue(1, 2)
            .Fecha = CDate(vNewValue(1, 3))
            .Dia = vNewValue(1, 4)
        '    .Semana = CInt(vNewValue(1, 5))
            mStr = vNewValue(1, 6) & "-" & _
                   vNewValue(1, 7) & "-" & _
                   vNewValue(1, 8) & "-" & _
                   vNewValue(1, 9) & "-" & _
                   vNewValue(1, 10) & "-" & _
                   vNewValue(1, 11)
            .Combinacion.Texto = mStr
            .Complementario = CInt(vNewValue(1, 12))
            .Reintegro = CInt(vNewValue(1, 13))
            .Ordenado = IIf(vNewValue(1, 14) = "Si", True, False)
            .Premios.Parse (vNewValue(1, 15))
        End With
    End If
  
  On Error GoTo 0
SetCmbnGndrToSorteo_CleanExit:
    Exit Sub
SetCmbnGndrToSorteo_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    Call HandleException(ErrNumber, ErrDescription, "GenPullCombinacionesModel.SetCmbnGndrToSorteo", ErrSource)
    Err.Raise ErrNumber, "GenPullCombinacionesModel.SetCmbnGndrToSorteo", ErrDescription
End Sub




'--- Métodos Públicos ---------------------------------------------------------*

'------------------------------------------------------------------------------*
' Procedimiento  : IsValid
' Fecha          : mi., 21/ago/2019 23:02:33
' Propósito      : Validar los parametros del proceso
' Retorno        : True, si es válido, False si no lo es
'------------------------------------------------------------------------------*
Public Function IsValid() As Boolean
  On Error GoTo IsValid_Error
    '
    '   #TODO: Validar los datos según la fase
    '          Fase Generar: matriz de numeros y numeros por combinacion
    '          Fase Filtrar: Combinaciones generadas > 0
    '          Fase Evaluar: matriz de sorteos con datos y combinaciones filtradas
    '          Fase Comprobar: Combinaciones evaluadas y Combinacion ganadora
    mErrorValidacion = 0
    '
    '
    If mNumSugerencia = 0 Then
        mErrorValidacion = mErrorValidacion + ERR_NUMSUGOBLIGATIRIO
    End If
    '
    '
    If UBound(mMatrizNumeros) = 1 Then
        mErrorValidacion = mErrorValidacion + ERR_MTRZOBLIGATORIA
    End If
    ' #TODO: Comprobar que los numeros son mayores que las combinaciones
    ' #TODO: Comprobar que todos los numeros están en el rango definido
    '
    '
    IsValid = IIf(mErrorValidacion = 0, True, False)
    
  On Error GoTo 0
IsValid__CleanExit:
    Exit Function
            
IsValid_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    Call HandleException(ErrNumber, ErrDescription, "GenPullCombinacionesModel.IsValid", ErrSource)
    Err.Raise ErrNumber, "GenPullCombinacionesModel.IsValid", ErrDescription
End Function




'------------------------------------------------------------------------------*
' Procedimiento  : GetMessage
' Fecha          : mi., 21/ago/2019 23:02:33
' Propósito      : Devuelve el mensaje de error de la validación
' Retorno        : Mensaje de error
'------------------------------------------------------------------------------*
Public Function GetMessage() As String
    Dim sResult As String
    '
    '   Si no hay error inicializamos el mensaje
    '
    If mErrorValidacion = 0 Then
        sResult = ""
    Else
        '
        '   Si hay algún error inicializamos la cabecera del error
        '
        sResult = MSG_ERRORESVALIDACION
    End If
    
    If (mErrorValidacion And ERR_NUMSUGOBLIGATIRIO) Then
        sResult = sResult & vbTab & MSG_NUMSUGOBLIGATIRIO & vbCrLf
    End If
        
    If (mErrorValidacion And ERR_MTRZOBLIGATORIA) Then
        sResult = sResult & vbTab & MSG_MTRZOBLIGATORIA & vbCrLf
    End If
    '
    '
    '
    '
    '   Devolvemos el mensaje
    '
    GetMessage = sResult

End Function
'------------------------------------------------------------------------------*
' Procedimiento  : IncrementarMatriz
' Fecha          : mi., 21/ago/2019 23:02:33
' Propósito      : Incrementa los punteros que representan las posiciones
'                  de los numeros
' Retorno        : True si ha habido cambios, false ha llegado al final
'------------------------------------------------------------------------------*
'      i =>      0   1   2   3   4   5
'      A(i,0)=>  1   2   3   4   5   6
'      B(i,1)=> 15  16  17  18  19  20
Private Function IncrementarMatriz(ByRef datMatriz As Variant) As Boolean
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    Dim iMax As Integer
    Dim mChange As Boolean
    
 On Error GoTo IncrementarMatriz_Error
    '
    '   Establecemos no hay cambios de indices
    mChange = False
    '
    '   Calculamos el numero max de indices
    iMax = mNumSugerencia - 1
    '
    '   bucle decreciente con todos los indices
    '
    For i = iMax To 0 Step -1
        '
        '   Si el indice iesimo ha llegado al max
        '
        If datMatriz(i, 0) = datMatriz(i, 1) Then
            '
            '   Comprobamos el indice anterior
            '
            j = i - 1
            '
            '   Si el anterior no ha llegado al máximo
            '
            If Not (datMatriz(j, 0) = datMatriz(j, 1)) Then
                '
                '   reiniciamos los indices desde la posicion j
                '   hasta el final
                For k = j To iMax
                    datMatriz(k, 0) = datMatriz(k, 0) + 1
                    If k > j Then
                        datMatriz(k, 0) = datMatriz(k - 1, 0) + 1
                    End If
                Next k
                mChange = True
                Exit For
            Else
                If j = 0 Then
                    mChange = False
                    Exit For
                End If
            End If
        Else
            datMatriz(i, 0) = datMatriz(i, 0) + 1
            mChange = True
            Exit For
        End If
    Next i
    IncrementarMatriz = Not mChange
    
  On Error GoTo 0
IncrementarMatriz__CleanExit:
    Exit Function
IncrementarMatriz_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    Call HandleException(ErrNumber, ErrDescription, "GenPullCombinacionesModel.IncrementarMatriz", ErrSource)
    Err.Raise ErrNumber, "GenPullCombinacionesModel.IncrementarMatriz", ErrDescription
End Function



'------------------------------------------------------------------------------*
' Procedimiento  : GetFiltrosOf
' Fecha          : lu., 21/dic/2020 18:55:25
' Propósito      : Obtener todas las categorias de un filtro para un numero
'                  determinado de sugerencias
' Parámetros     : Nombre del filtro, numeros de la combinacion
' Retorno        : Matriz con los posibles valores para el filtro
'------------------------------------------------------------------------------*
Public Function GetFiltrosOf(datFiltro As String, datNumSuges As Integer) As Variant
    Dim mMatriz As Variant
    Dim mFiltr As FiltroCombinacion
    
  On Error GoTo GetFiltrosOf_Error
    '
    '   Definimos un filtro  para obtener caracteristicas
    '
    Set mFiltr = New FiltroCombinacion
    mFiltr.NameFiltro = datFiltro
    '
    '   Obtenemos los valores posibles para el filtro
    '
    mMatriz = mFiltr.GetValoresFiltros(datNumSuges)
    '
    '   Devolvemos el filtos
    '
    GetFiltrosOf = mMatriz
  On Error GoTo 0
GetFiltrosOf__CleanExit:
    Exit Function
            
GetFiltrosOf_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    Call HandleException(ErrNumber, ErrDescription, "GenPullCombinacionesModel.GetFiltrosOf", ErrSource)
    Err.Raise ErrNumber, "GenPullCombinacionesModel.GetFiltrosOf", ErrDescription
End Function




'------------------------------------------------------------------------------*
' Procedimiento  : GenerarCombinaciones
' Fecha          : lu., 21/dic/2020 18:56:50
' Propósito      : Genera combinaciones de N numeros tomados de m en m
' Parámetros     : Celda de salida
' Retorno        : Total de combinaciones a calcular
'
'                #TODO: Si el máximo de combinaciones sobrepasa 1048574
'                       escribir en fichero
'------------------------------------------------------------------------------*
Public Sub GenerarCombinaciones(datCelda As Range, datMaximo As Long)
    Dim mMatrizIndex As Variant
    Dim i            As Integer
    Dim j            As Integer
    Dim mComb        As Combinacion
    Dim mNum         As Numero
    Dim mFinProceso  As Boolean
    Dim mIdComb      As Long
    Dim mSelComb     As Boolean

  On Error GoTo GenerarCombinaciones_Error
    '
    '   Configuramos el visor de progreso
    '
    Set mFrm = New frmProgreso
    With mFrm
        .Fase = "Generacion de combinaciones"
        .Maximo = datMaximo
    End With
    mFrm.Show
    '
    '   Creamos una matriz de tantos indices como numeros
    '   a combinar y con sus limites superiores
    '   ejemplo para 6 numeros y 20 numeros a combinar:
    '      i =>     0   1   2   3   4   5
    '      A(i,0)=> 1   2   3   4   5   6
    '      B(i,1)=>15  16  17  18  19  20
    '
    '
    i = NumerosSugerencia - 1
    ReDim mMatrizIndex(i, 1)
    
    j = TotalNumerosCombinar - NumerosSugerencia + 1
    
    For i = 0 To (NumerosSugerencia - 1)
        mMatrizIndex(i, 0) = i + 1
        mMatrizIndex(i, 1) = j
        j = j + 1
    Next i
    '
    '   Inicilializamos
    '
    mFinProceso = False
    mIdComb = 0
    While Not mFinProceso
        '
        '
        '
        Set mComb = New Combinacion
        For i = 1 To mNumSugerencia
            j = mMatrizIndex(i - 1, 0)
            Set mNum = New Numero
            Set mNum = mNums.Item(j)
            mComb.Add mNum
        Next i
        '
        '   Visualizar combinación
        '
        datCelda.Offset(mIdComb, 0).Value = mIdComb + 1
        datCelda.Offset(mIdComb, 1).Value = "'" & mComb.Texto
        datCelda.Offset(mIdComb, 2).Value = "'" & mComb.FormulaParidad
        datCelda.Offset(mIdComb, 3).Value = "'" & mComb.FormulaAltoBajo
        datCelda.Offset(mIdComb, 4).Value = "'" & mComb.FormulaConsecutivos
        datCelda.Offset(mIdComb, 5).Value = "'" & mComb.FormulaDecenas
        datCelda.Offset(mIdComb, 6).Value = "'" & mComb.FormulaSeptenas
        datCelda.Offset(mIdComb, 7).Value = "'" & mComb.Suma
        datCelda.Offset(mIdComb, 8).Value = "'" & mComb.FormulaTerminaciones
        mFrm.Valor = mIdComb
        '
        '   Incrementamos Id combinación
        '
        mIdComb = mIdComb + 1
        
        '
        '   Incrementamos matriz de Indices
        '
        mFinProceso = IncrementarMatriz(mMatrizIndex)
    Wend
    mFrm.Valor = mIdComb
    mFrm.DisProceso
    
    Set mFrm = Nothing
  
  On Error GoTo 0
GenerarCombinaciones__CleanExit:
    Exit Sub
GenerarCombinaciones_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    Call HandleException(ErrNumber, ErrDescription, "GenPullCombinacionesModel.GenerarCombinaciones", ErrSource)
    Err.Raise ErrNumber, "GenPullCombinacionesModel.GenerarCombinaciones", ErrDescription
End Sub




'------------------------------------------------------------------------------*
' Procedimiento  : FiltrarCombinaciones
' Fecha          : lu., 21/dic/2020 18:38:32
' Propósito      : Aplicar filtros a las combinaciones generadas
'------------------------------------------------------------------------------*
Public Sub FiltrarCombinaciones(datCeldaOrigen As Range, datCeldaDestino As Range, datMaximo As Long)
    Dim mRg         As Range
    Dim mRow        As Range
    Dim mRgDest     As Range
    Dim i           As Integer
    Dim xOffset     As Integer
    Dim mComb       As Combinacion
    
  On Error GoTo FiltrarCombinaciones_Error
    '
    '   Configuramos el visor de progreso
    '
    Set mFrm = New frmProgreso
    With mFrm
        .Fase = "Filtrado de combinaciones"
        .Maximo = datMaximo
    End With
    mFrm.Show
    '
    '   Definimos el rango de entrada
    '
    Set mRg = datCeldaOrigen.CurrentRegion
    i = mRg.Rows.Count
    Set mRg = mRg.Offset(2, 0).Resize(i - 2, mRg.Columns.Count)
    '
    '   Inicializamos contador
    '
    i = 0: xOffset = 0
    '
    '   Para cada combinación generada en la fila
    '
    For Each mRow In mRg.Rows
        '
        '   Creamos la combinación
        '
        Set mComb = New Combinacion
        mComb.Texto = mRow.Value2(1, 2)
        '
        '   Comprobamos la combinación con los filtros
        '   Si la combinación pasa los filtros
        '
        If mFltrs.Evaluar(mComb) Then
        '
        '       Escribimos la fila
        '       Incrementamos contador de salida
        
            datCeldaDestino.Offset(xOffset, 0).Value = mRow.Value2(1, 1)
            datCeldaDestino.Offset(xOffset, 1).Value = mRow.Value2(1, 2)
            datCeldaDestino.Offset(xOffset, 2).Value = "'" & mRow.Value2(1, 3)
            datCeldaDestino.Offset(xOffset, 3).Value = "'" & mRow.Value2(1, 4)
            datCeldaDestino.Offset(xOffset, 4).Value = "'" & mRow.Value2(1, 5)
            datCeldaDestino.Offset(xOffset, 5).Value = "'" & mRow.Value2(1, 6)
            datCeldaDestino.Offset(xOffset, 6).Value = "'" & mRow.Value2(1, 7)
            datCeldaDestino.Offset(xOffset, 7).Value = "'" & mRow.Value2(1, 8)
            datCeldaDestino.Offset(xOffset, 8).Value = "'" & mRow.Value2(1, 9)
            xOffset = xOffset + 1
        End If
        
        i = i + 1
        '
        '   Visualizamos el proceso
        '
        mFrm.Valor = i
    Next mRow
    '
    '   Visualizamos el proceso
    '
    mFrm.Valor = i
    mFrm.DisProceso
    
    
  On Error GoTo 0
FiltrarCombinaciones__CleanExit:
    Exit Sub
FiltrarCombinaciones_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    Call HandleException(ErrNumber, ErrDescription, "GenPullCombinacionesModel.FiltrarCombinaciones", ErrSource)
    Err.Raise ErrNumber, "GenPullCombinacionesModel.FiltrarCombinaciones", ErrDescription
End Sub





'------------------------------------------------------------------------------*
' Procedimiento  : EvaluarCombinaciones
' Fecha          :
' Propósito      :
'------------------------------------------------------------------------------*
Public Sub EvaluarCombinaciones(datCeldaOrigen As Range, _
                                datCeldaDestino As Range, _
                                datFiltradas As Long)
    Dim i           As Integer
    Dim mRg         As Range
    Dim mRow        As Range
    Dim xOffset     As Integer
    Dim mComb       As Sugerencia
    Dim mDSrt       As Integer      'Numero de dias con al menos un numero acertado
    Dim mPunt       As Integer
    Dim mCheck      As ComprobarSugerencia
    Dim mSorteo     As Sorteo
    
  On Error GoTo EvaluarCombinaciones_Error
    '
    '   Configuramos el visor de progreso
    '
    Set mFrm = New frmProgreso
    With mFrm
        .Fase = "Evaluar combinaciones con Sorteos"
        .Maximo = datFiltradas
    End With
    mFrm.Show
    '
    '   Creamos el comprobador de Sugerencias
    '
    Set mCheck = New ComprobarSugerencia
    '
    '   Definimos el rango de entrada
    '
    Set mRg = datCeldaOrigen.CurrentRegion
    i = mRg.Rows.Count
    Set mRg = mRg.Offset(2, 0).Resize(i - 2, mRg.Columns.Count)
    '
    '   Inicializamos contador
    '
    i = 0: xOffset = 0
    '
    '   Para cada combinación generada en la fila
    '
    For Each mRow In mRg.Rows
        '
        '   Creamos la Sugerencia
        '
        Set mComb = New Sugerencia
        mComb.Combinacion.Texto = mRow.Value2(1, 2)
        '
        '   Inicializamos totales por combinación
        '
        mDSrt = 0: mPunt = 0
        '
        '   Comprobamos la combinación con la colección de sorteos
        '
        For Each mSorteo In mSorteos.Items
            '
            '   Asignamos sorteo
            '
            Set mCheck.Sorteo = mSorteo
            '
            '   Comprobamos la sugerencia
            '
            mCheck.ComprobarSugerencia mComb
            If mCheck.NumerosAcertados > 0 Then
                mDSrt = mDSrt + 1
                mPunt = mPunt + mCheck.Puntuacion
            End If
        Next mSorteo
        '   Escribimos la fila
        '   Incrementamos contador de salida
    
        datCeldaDestino.Offset(xOffset, 0).Value = mRow.Value2(1, 1)    ' N
        datCeldaDestino.Offset(xOffset, 1).Value = mRow.Value2(1, 2)    ' Combinación
        datCeldaDestino.Offset(xOffset, 2).Value = mDSrt                ' Dias acertados
        datCeldaDestino.Offset(xOffset, 3).Value = mPunt                ' Puntuación
        xOffset = xOffset + 1
        
        i = i + 1
        '
        '   Visualizamos el proceso
        '
        mFrm.Valor = i
    Next mRow
    '
    '   Visualizamos el proceso
    '
    mFrm.Valor = i
    mFrm.DisProceso
      
  
  On Error GoTo 0
EvaluarCombinaciones__CleanExit:
    Exit Sub
EvaluarCombinaciones_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    Call HandleException(ErrNumber, ErrDescription, "GenPullCombinacionesModel.EvaluarCombinaciones", ErrSource)
    Err.Raise ErrNumber, "GenPullCombinacionesModel.EvaluarCombinaciones", ErrDescription
End Sub




'------------------------------------------------------------------------------*
' Procedimiento  : ComprobarCombinaciones
' Fecha          :
' Propósito      :
'------------------------------------------------------------------------------*
Public Sub ComprobarCombinaciones(datCeldaOrigen As Range, _
                                  datCeldaDestino As Range, _
                                  datEvaluadas As Long)
    Dim i           As Integer
    Dim mRg         As Range
    Dim mRow        As Range
    Dim xOffset     As Integer
    Dim mComb       As Sugerencia
    Dim mCheck      As ComprobarSugerencia
    
  On Error GoTo ComprobarCombinaciones_Error
    '
    '   Configuramos el visor de progreso
    '
    Set mFrm = New frmProgreso
    With mFrm
        .Fase = "Comprobar combinaciones con Combinación Ganadora"
        .Maximo = datEvaluadas
    End With
    mFrm.Show
    '
    '   Creamos el comprobador de Sugerencias
    '
    Set mCheck = New ComprobarSugerencia
    '
    '   Cargamos la combinación ganadora
    '
    Set mCheck.Sorteo = mCombinacionGanadora
    '
    '   Definimos el rango de entrada
    '
    Set mRg = datCeldaOrigen.CurrentRegion
    i = mRg.Rows.Count
    Set mRg = mRg.Offset(2, 0).Resize(i - 2, mRg.Columns.Count)
    '
    '   Inicializamos contadores
    '
    i = 0: xOffset = 0
    mTotRegPremiados = 0: mTotRegComprobados = 0
    mTotImportePremios = 0: mTotCoste = 0
    '
    '   Para cada combinación generada en la fila
    '
    For Each mRow In mRg.Rows
        '
        '   Creamos la Sugerencia
        '
        Set mComb = New Sugerencia
        mComb.Combinacion.Texto = mRow.Value2(1, 2)
        '
        '
        '   Comprobamos la sugerencia
        '
        mCheck.ComprobarSugerencia mComb
        '   Escribimos la fila
        '   Incrementamos contador de salida
        '
        datCeldaDestino.Offset(xOffset, 0).Value = mRow.Value2(1, 1)    ' N
        datCeldaDestino.Offset(xOffset, 1).Value = mRow.Value2(1, 2)    ' Combinación
        datCeldaDestino.Offset(xOffset, 2).Value = mRow.Value2(1, 4)    ' Puntuación Previa
        datCeldaDestino.Offset(xOffset, 3).Value = mCheck.Puntuacion    ' Puntuación
        If mCheck.CatPremioApuesta <> Ninguna Then
            datCeldaDestino.Offset(xOffset, 4).Value = mCheck.CategoriaPremioTxt
            datCeldaDestino.Offset(xOffset, 6).Value = mCheck.ImporteApuesta
            mTotImportePremios = mTotImportePremios + mCheck.ImporteApuesta
            mTotRegPremiados = mTotRegPremiados + 1
        Else
            datCeldaDestino.Offset(xOffset, 4).Value = mCheck.NumerosAcertados
        End If
        datCeldaDestino.Offset(xOffset, 5).Value = mComb.Coste
        mTotCoste = mTotCoste + mComb.Coste
        mTotRegComprobados = mTotRegComprobados + 1
        
        xOffset = xOffset + 1
        
        i = i + 1
        '
        '   Visualizamos el proceso
        '
        mFrm.Valor = i
    Next mRow
    '
    '   Visualizamos el proceso
    '
    mFrm.Valor = i
    mFrm.DisProceso
    
  On Error GoTo 0
ComprobarCombinaciones__CleanExit:
    Exit Sub
ComprobarCombinaciones_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    Call HandleException(ErrNumber, ErrDescription, "GenPullCombinacionesModel.ComprobarCombinaciones", ErrSource)
    Err.Raise ErrNumber, "GenPullCombinacionesModel.ComprobarCombinaciones", ErrDescription
End Sub
'' *===========(EOF): GenPullCombinacionesModel.cls
