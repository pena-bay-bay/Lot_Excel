VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GenCombinacionesModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' *============================================================================*
' *
' *     Fichero    : GeneradorModel.cls
' *
' *     Autor      : Carlos Almela Baeza
' *     Creación   : mi., 21/ago/2019 17:05:19
' *     Versión    : 1.0
' *     Propósito  :
' *
' *============================================================================*
Option Explicit
Option Base 0
'
'--- Variables Privadas -------------------------------------------------------*
Private mMatrizNumeros As Variant
Private mNumSugerencia As Integer
Private mTotNumCombinar As Integer
Private mTotCombinaciones As Long
Private mFiltros As Variant
Private mFltrs As Collection
Private mNums As Collection
Private mError As Long

'--- Constantes ---------------------------------------------------------------*
Private mErrorValidacion                    As Long       ' Control de errores de validación
Private Const ERR_NUMSUGOBLIGATIRIO         As Long = 1   ' Numeros de sugerencia obligatorios
Private Const ERR_MTRZOBLIGATORIA           As Long = 2   ' Matriz de números Obligatoria

'--- Mensajes -----------------------------------------------------------------*
'--- Errores ------------------------------------------------------------------*
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

Public Property Get Filtros() As Variant
    Filtros = mFiltros
End Property

Public Property Let Filtros(ByVal vNewValue As Variant)
    mFiltros = vNewValue
    If IsArray(vNewValue) Then
        SetFiltrosToColleccion vNewValue
    End If
End Property

'--- Métodos Privados ---------------------------------------------------------*
Private Sub Class_Initialize()
    mNumSugerencia = 0
    mTotNumCombinar = 0
    Set mFltrs = New Collection
    Set mNums = New Collection
End Sub

Private Sub Class_Terminate()
    Set mFltrs = Nothing
    Set mNums = Nothing
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
    Call HandleException(ErrNumber, ErrDescription, "GeneradorModel.SetNumerosToColleccion", ErrSource)
    Err.Raise ErrNumber, "GeneradorModel.SetNumerosToColleccion", ErrDescription
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
        If mFltrs.Count = 0 Then
            mFltrs.Add mFilter, CStr(i)
        Else
            mFltrNew = True
            '
            '   Si es un tipo de filtro existente se agrega el valor
            '
            For j = 1 To mFltrs.Count
                If mFltrs.Item(j).TipoFiltro = mFilter.TipoFiltro Then
                    mFltrs.Item(j).FilterValue = mFltrs.Item(j).FilterValue & ";" & mFilter.FilterValue
                    mFltrs.Item(j).MultiplesFiltros = True
                    mFltrNew = False
                End If
            Next j
            If mFltrNew Then
                mFltrs.Add mFilter, CStr(i), , mFltrs.Count
            End If
        End If
    Next i
          
  On Error GoTo 0
SetFiltrosToColleccion_CleanExit:
    Exit Sub
            
SetFiltrosToColleccion_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    Call HandleException(ErrNumber, ErrDescription, "GeneradorModel.SetFiltrosToColleccion", ErrSource)
    Err.Raise ErrNumber, "GeneradorModel.SetFiltrosToColleccion", ErrDescription
End Sub
'--- Métodos Públicos ---------------------------------------------------------*
'------------------------------------------------------------------------------*
' Procedimiento  : GenerarCombinaciones
' Fecha          : dd/MMM/yyyy
' Propósito      :
' Parámetros     :
' Retorno        :
'                   TODO: Si el máximo de combinaciones sobrepasa 1048574
'                         escribir en fichero
'------------------------------------------------------------------------------*
Public Sub GenerarCombinaciones(datCelda As Range)
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
        '   Filtrar combinación
        '
        mSelComb = True
        For i = 1 To mFltrs.Count
            mSelComb = mSelComb And mFltrs.Item(i).EvaluarCombinacion(mComb)
        Next i
        '
        '   Si se selecciona la combinación
        '
        If mSelComb Then
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
            '
            '   Incrementamos Id combinación
            '
            mIdComb = mIdComb + 1
        End If
        '
        '   Incrementamos matriz de Indices
        '
        mFinProceso = IncrementarMatriz(mMatrizIndex)
    Wend
  On Error GoTo 0
GenerarCombinaciones__CleanExit:
    Exit Sub
            
GenerarCombinaciones_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    Call HandleException(ErrNumber, ErrDescription, "GenCombinacionesModel.GenerarCombinaciones", ErrSource)
    Err.Raise ErrNumber, "GenCombinacionesModel.GenerarCombinaciones", ErrDescription
End Sub

'------------------------------------------------------------------------------*
' Procedimiento  : IsValid
' Fecha          : mi., 21/ago/2019 23:02:33
' Propósito      : Validar los parametros del proceso
' Retorno        : True, si es válido, False si no lo es
'------------------------------------------------------------------------------*
Public Function IsValid() As Boolean
  On Error GoTo IsValid_Error
    '
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
    ' TODO: Comprobar que los numeros son mayores que las combinaciones
    ' TODO: Comprobar que todos los numeros están en el rango definido
    '
    '
    IsValid = IIf(mErrorValidacion = 0, True, False)
    
  On Error GoTo 0
IsValid__CleanExit:
    Exit Function
            
IsValid_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    Call HandleException(ErrNumber, ErrDescription, "GeneradorModel.IsValid", ErrSource)
    Err.Raise ErrNumber, "GeneradorModel.IsValid", ErrDescription
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
'' *===========(EOF): GeneradorModel.cls
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
    Call HandleException(ErrNumber, ErrDescription, "GeneradorModel.IncrementarMatriz", ErrSource)
    Err.Raise ErrNumber, "GeneradorModel.IncrementarMatriz", ErrDescription
End Function

'------------------------------------------------------------------------------*
' Procedimiento  : GetFiltrosOf
' Fecha          :
' Propósito      : Obtener todas las categorias de un filtro para un numero
'                  determinado de sugerencias
' Parámetros     :
' Retorno        :
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
    Call HandleException(ErrNumber, ErrDescription, "GenCombinacionesModel.GetFiltrosOf", ErrSource)
    Err.Raise ErrNumber, "GenCombinacionesModel.GetFiltrosOf", ErrDescription
End Function
