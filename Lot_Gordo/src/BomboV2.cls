VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BomboV2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' *============================================================================*
' *
' *     Fichero    : BomboV2.cls
' *
' *     Autor      : Carlos Almela Baeza
' *     Creación   : mar, 08/may/2018 06:43:15
' *     Versión    : 1.0
' *     Propósito  : Clase que simula un Bombo
' *
' *============================================================================*
Option Explicit
Option Base 0
'
'--- Variables Privadas -------------------------------------------------------*
Private mJuego         As Juego       ' Tipo de juego para que se carga el bombo
Private mCol           As Collection  ' Colección de Bolas del Bombo
Private mColSel        As Collection  ' Colección de Bolas extraidas del bómbo
Private mSituacion     As Integer     ' Situación del Bombo: 1-Vacio, 2-lleno, 3-Cargado
Private mTipoBombo     As Integer     ' Tipo de Bombo: 1-Principal, 2-Secundario
Private mTipoGiros     As Integer     ' Tipo de Giro: 1-Por tiempo, 2-Por giros
Private mBola          As bola        ' Bola
Private mMaxBolas      As Integer     ' numero mínimo del juego
Private mMinBolas      As Integer     ' numero máximo del juego
Private i              As Integer     ' contador genérico
Private mKey           As String      ' clave del número en la colección
Private mNumGiros      As Integer     ' Numero de giros realizados en el bombo
Private mTiempoGiro    As Date        ' Duración de los giros
Private mProbNum       As Variant     ' Matriz de probabilidades
Private mBolasCargadas As Collection  ' Colección de Bolas cargadas
Private mTotSumaProb   As Double      ' Sumatorio total de probabilidades
Private mBomboCargado  As Boolean     ' Indicador de bombo cargado con proba

'--- Constantes ---------------------------------------------------------------*
'--- Mensajes -----------------------------------------------------------------*
Private Const MSG_VALTIPOGIRO = "Bombo: Error en Tipo de giro debe ser 1 ó 2."
Private Const MSG_BOMBOVACIO = "#Error: Bombo vacío."

'--- Errores ------------------------------------------------------------------*
Private Const ERR_VALTIPOGIRO = 3001
'
'--- Metodos Privados ---------------------------------------------------------*
Private Sub Class_Initialize()
    Set mCol = New Collection
    Set mColSel = New Collection
    mSituacion = lotVacio           ' Vacio
    mTipoBombo = 1                  ' Bombo principal
    mNumGiros = 0                   ' Inicio Giros
    mTiempoGiro = Time() - Time()   ' Inicio tiempo
    ReDim mProbNum(1, 1)            ' redimensionamos la matriz a 2x2
End Sub
Private Sub Class_Terminate()
    Set mCol = Nothing
    Set mColSel = Nothing
    Set mBolasCargadas = Nothing
End Sub


'------------------------------------------------------------------------------*
' Procedimiento  : RealizarGiro
' Fecha          : 13/may/2018
' Propósito      : Realiza un giro del bómbo
'------------------------------------------------------------------------------*
Private Sub RealizarGiro()
    Dim mIdxA  As Integer
    Dim mIdxB  As Integer
   
 On Error GoTo RealizarGiro_Error
    '
    '   seleccionamos un índice origen al azar
    '
    mIdxA = GetIndex()
    mIdxB = 0
    '
    '   seleccionamos un índice destino al azar
    '   distinto del origen
    '
    While ((mIdxA = mIdxB) _
    Or (mIdxB = 0))
        mIdxB = GetIndex()
    Wend
    '
    '   Intercambiamos las bolas entre indice origen y destino
    '
    Set mBola = mCol.Item(mIdxA)
    '
    '   Sustituimos los elementos
    '
    SustituirElem mCol, mIdxA, mCol.Item(mIdxB)
    SustituirElem mCol, mIdxB, mBola
    
 On Error GoTo 0
RealizarGiro__CleanExit:
    Exit Sub
            
RealizarGiro_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    Call HandleException(ErrNumber, ErrDescription, "BomboV2.RealizarGiro", ErrSource)
    Err.Raise ErrNumber, "BomboV2.RealizarGiro", ErrDescription
End Sub



'------------------------------------------------------------------------------*
' Procedimiento  : GetIndex
' Fecha          : 13/may/2018
' Propósito      : Devuelve un indice de la colección aleatoriamente
'------------------------------------------------------------------------------*
Private Function GetIndex() As Integer
    Dim mxBolas  As Integer
    Static b_rand As Boolean
    '
    '
    '
    If Not b_rand Then          'La primera vez que se ejecuta
        b_rand = True           'la función RND se ceba la
        Randomize               'la semilla
    End If
    '
    '   Total de bolas
    '
    mxBolas = mCol.Count
    '
    '   Obtenemos aleatoriamente un puntero
    '
    GetIndex = Int(mCol.Count * Rnd)
    '
    '   Si el puntero es 0 pasamas al último numero
    '
    If GetIndex = 0 Then
        GetIndex = mxBolas
    ElseIf GetIndex >= mxBolas Then
        GetIndex = 1
    End If
End Function

'------------------------------------------------------------------------------*
' Procedimiento  : GetRueda
' Fecha          :
' Propósito      : Devuelve un indice de la colección ejecutando la ruleta
'------------------------------------------------------------------------------*
Private Function GetRueda() As Integer
    Dim mTotSum         As Double
    Dim mMinProb        As Double
    Dim mMaxProb        As Double
    Dim mValue          As Double
    Dim mValueSemilla   As Double
    Dim i               As Integer
    Dim j               As Integer
    Static mRandOn      As Boolean
    Dim mDatos          As Variant
    Dim mElement        As Variant
    
 On Error GoTo GetRueda_Error
    '
    '   La primera vez que se ejecuta se ceba la función Randomice
    '
    If Not mRandOn Then
        mRandOn = True
        Randomize
    End If
    '
    '   Inicializamos parametros
    '
    mTotSum = 0
    mMinProb = 9999
    mMaxProb = -9999
    '
    '   Redimensionamos matriz
    '
    ReDim mDatos(mCol.Count - 1, 1)
    i = 0
    '
    '   Calcular sumatorio y valores extremos
    '
    For Each mElement In mBolasCargadas
        mKey = mElement(0)
        mValue = mElement(1)
        mDatos(i, 0) = mKey
        mDatos(i, 1) = mValue
        If mValue > mMaxProb Then
            mMaxProb = mValue
        End If
        If mValue < mMinProb Then
            mMinProb = mValue
        End If
        mTotSum = mTotSum + mValue
        i = i + 1
    Next mElement
    '
    '   Homogenenizamos valores a 100 si hay datos
    '
    For i = 0 To UBound(mDatos)
        mKey = mDatos(i, 0)
        mValue = mDatos(i, 1)
        '
        '   Si trabajamos con proabilidades
        '
        If mTotSum < 1 And mTotSum > 0 Then
            mValue = (mValue / mTotSum) * 100
        ElseIf mTotSum > 0 Then
            '
            '   Si trabajamos con Modas y Tm
            '
            mValue = mMaxProb - mValue
            mValue = (mValue / mTotSum) * 100
        Else
            mValue = 100 / mBolasCargadas.Count
        End If
        mDatos(i, 1) = mValue
    
'        Debug.Print "Value (" & mKey & ") => " & mValue

    Next i
    '
    '   Generamos una semilla entre 0 y 100
    '
    mValueSemilla = 100 * Rnd
    '
    '   Ejecutamos la rueda
    '
    mMinProb = 0
    mMaxProb = 0
    GetRueda = 1
    For i = 0 To mCol.Count - 1
        mKey = CStr(mCol.Item(i + 1).Numero.Valor)
        '
        '   Busca la probabilidad de la bola
        '
        For j = 0 To UBound(mDatos)
            If mKey = mDatos(j, 0) Then
                mMaxProb = mDatos(j, 1)
                mTotSum = mMinProb + mMaxProb
                Exit For
            End If
        Next j
        If mValueSemilla >= mMinProb And _
           mValueSemilla < mTotSum Then
                GetRueda = i + 1
                Exit For
        End If
        mMinProb = mMinProb + mMaxProb
    Next i
        
    
 
 On Error GoTo 0
GetRueda_CleanExit:
    Exit Function
GetRueda_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    Call HandleException(ErrNumber, ErrDescription, "BomboV2.GetRueda", ErrSource)
    Err.Raise ErrNumber, "BomboV2.GetRueda", ErrDescription
End Function


'--- Metodos Publicos ---------------------------------------------------------*


'------------------------------------------------------------------------------*
' Procedimiento  : Cargar
' Fecha          : 12/may/2018
' Propósito      : Cargar las Bolas en el Bombo segun el juego
' Parámetros     :
' Retorno        :
'------------------------------------------------------------------------------*
Public Sub Cargar()
    
 On Error GoTo Cargar_Error
    '
    '   Si el bombo está cargado salimos de la rutina
    '
    If mSituacion <> lotVacio Then
        Exit Sub
    End If
    '
    '   Deterninamos en numero de bolas del bombo
    '
    Select Case mJuego
        Case bonoloto, loteriaPrimitiva
            mMinBolas = 1
            mMaxBolas = 49
            
        Case Euromillones
            mMinBolas = 1
            If mTipoBombo = 1 Then
                mMaxBolas = 50
            Else
                mMaxBolas = 12
            End If
        
        Case gordoPrimitiva
            If mTipoBombo = 1 Then
                mMinBolas = 1
                mMaxBolas = 54
            Else
                mMinBolas = 0
                mMaxBolas = 9
            End If

    End Select
    '
    '   Cargamos las bolas del bombo si tenemos criterio
    '
    If mMinBolas < mMaxBolas Then
        '
        '   bucle de carga
        '
        For i = mMinBolas To mMaxBolas
            '
            '   Convertimos el numero en clave
            '
            mKey = CStr(i)
            '
            '   Creamos la bola
            '
            Set mBola = New bola
            '
            '   asignamos el número
            '
            mBola.Numero.Valor = i
            '
            '   agregamos bola
            '
            mCol.Add mBola, mKey
         
        Next i
        '
        '   Actualizamos la situación del bombo
        '
        mSituacion = lotLleno
        '
        '   Dimensionamos la matriz de
        '
        ReDim mProbNum(mMaxBolas - mMinBolas, 1)
        
    End If
    
 On Error GoTo 0
Cargar__CleanExit:
    Exit Sub
            
Cargar_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    Call HandleException(ErrNumber, ErrDescription, "BomboV2.Cargar", ErrSource)
    Err.Raise ErrNumber, "BomboV2.Cargar", ErrDescription
End Sub



'------------------------------------------------------------------------------*
' Procedimiento  : Reiniciar
' Fecha          : 12/may/2018
' Propósito      : R
' Parámetros     :
' Retorno        :
'------------------------------------------------------------------------------*
Public Sub Reiniciar()
 On Error GoTo Reiniciar_Error
    '
    '   Si el bombo está vacio salimos de la rutina
    '
    If mSituacion = lotVacio Then
        Exit Sub
    End If
    '
    '   Para cada bola en la colección de seleccionadas
    '
    For Each mBola In mColSel
        '
        '   Calculamos la clave de la bola
        '
        mKey = CStr(mBola.Numero.Valor)
        '
        '   Agregamos la bola a la colección
        '
        mCol.Add mBola, mKey
    Next mBola
    '
    '   Inicializamos Bolas seleccionadas
    '
    EliminarTodosElementos mColSel
    '
    '   Inicializamos la situacion
    '
    mSituacion = lotVacio
    
 On Error GoTo 0
Reiniciar__CleanExit:
    Exit Sub
            
Reiniciar_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    Call HandleException(ErrNumber, ErrDescription, "BomboV2.Reiniciar", ErrSource)
    Err.Raise ErrNumber, "BomboV2.Reiniciar", ErrDescription
End Sub


'------------------------------------------------------------------------------*
' Procedimiento  : Girar
' Fecha          : 12/may/2018
' Propósito      :
' Parámetros     :
' Retorno        : Gira el bombo según los criterios establecidos de tiempo o
'                  número de giros
'------------------------------------------------------------------------------*
Public Sub Girar()
    Dim mTFin                 As Date
    Dim mGirosRealizados      As Long
    
 On Error GoTo Girar_Error
    '
    '   Si el bombo está vacio sale de la rutina
    '
    If mSituacion = lotVacio Then
        Exit Sub
    End If
    '
    '   inicializamos el número de giros del bombo
    '
    mGirosRealizados = 0
    '
    '   En función del tipo de giro
    '
    If mTipoGiros = lotGiros Then
        '
        '   Numero de giros
        '
        For mGirosRealizados = 1 To mNumGiros
            RealizarGiro
            mGirosRealizados = mGirosRealizados + 1
        Next mGirosRealizados
        
    ElseIf mTipoGiros = lotTiempo Then
        '
        '   Tiempo de giro.
        '   Calculamos tiempo finalizacion
        '
        mTFin = Time() + mTiempoGiro
        While Time() < mTFin
            RealizarGiro
            mGirosRealizados = mGirosRealizados + 1
        Wend

    End If
 On Error GoTo 0
Girar__CleanExit:
    Exit Sub
            
Girar_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    Call HandleException(ErrNumber, ErrDescription, "BomboV2.Girar", ErrSource)
    Err.Raise ErrNumber, "BomboV2.Girar", ErrDescription
End Sub



'------------------------------------------------------------------------------*
' Procedimiento  : ExtraerBolas
' Fecha          : 12/may/2018
' Propósito      :
' Parámetros     :
' Retorno        : Cadena de texto con los números elegidos
'------------------------------------------------------------------------------*
Public Function ExtraerBolas(ByVal vNewValue As Integer) As String
    Dim mIdxA       As Integer
    
 On Error GoTo ExtraerBolas_Error
    '
    '   Si el bombo está vacio devolvemos un literal
    '
    If mSituacion = lotVacio Then
        ExtraerBolas = MSG_BOMBOVACIO
        Exit Function
    End If
    '
    '   Si el bombo contiene bolas
    '
    If mSituacion = lotLleno Then
        For i = 1 To vNewValue
            '
            '   Seleccionamos una bola
            '
            mIdxA = GetIndex()
            Set mBola = mCol.Item(mIdxA)
            mCol.Remove mIdxA
            '
            '   Agregamos la bola a la colección de seleccionadas
            '
            mColSel.Add mBola
        
        Next i
    End If
    '
    '   Si el bombo contiene bolas cargadas
    '
    If mSituacion = lotCargado Then
        For i = 1 To vNewValue
            '
            '   Seleccionamos una bola
            '
            mIdxA = GetRueda()
            Set mBola = mCol.Item(mIdxA)
            mCol.Remove mIdxA
            '
            '   Agregamos la bola a la colección de seleccionadas
            '
            mColSel.Add mBola
            '
            '   Reiniciamos probabilidades
            '
            SetProbToBolas
        Next i
    End If
    
    ExtraerBolas = ""
    '
    '   Montamos la cadena de bolas seleccionadas
    '
    For Each mBola In mColSel
        ExtraerBolas = ExtraerBolas & CStr(mBola.Numero.Valor) & "-"
    Next mBola
    '
    '   Eliminamos el último guion
    '
    ExtraerBolas = Left(ExtraerBolas, Len(ExtraerBolas) - 1)
 
 On Error GoTo 0
ExtraerBolas__CleanExit:
    Exit Function
            
ExtraerBolas_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    Call HandleException(ErrNumber, ErrDescription, "BomboV2.ExtraerBolas", ErrSource)
    Err.Raise ErrNumber, "BomboV2.ExtraerBolas", ErrDescription
End Function




'--- Propiedades --------------------------------------------------------------*
'------------------------------------------------------------------------------*
' Procedimiento  : Juego
' Fecha          : 10/may/2018
' Propósito      : Definir el juego para el que se carga el bombo
'------------------------------------------------------------------------------*
Public Property Get Juego() As Juego
    Juego = mJuego
End Property

Public Property Let Juego(ByVal vNewValue As Juego)
    mJuego = vNewValue
End Property


'------------------------------------------------------------------------------*
' Procedimiento  : NumBolas
' Fecha          : 10/may/2018
' Propósito      : Numero de bolas del Bombo
'------------------------------------------------------------------------------*
Public Property Get NumBolas() As Integer
    NumBolas = mCol.Count
End Property
'Public Property Let NumBolas(ByVal vNewValue As Variant)
'
'End Property


'------------------------------------------------------------------------------*
' Procedimiento  : Bolas
' Fecha          : 10/may/2018
' Propósito      : Colección de Bolas del Bombo
'------------------------------------------------------------------------------*
Public Property Get Bolas() As Collection
    Set Bolas = mCol
End Property
'Public Property Let Bolas(ByVal vNewValue As Variant)
'
'End Property



'------------------------------------------------------------------------------*
' Procedimiento  : TipoBombo
' Fecha          : 10/may/2018
' Propósito      : Tipo de Bombo según el juego: 1 Bombo principal,
'                  2 bombo secundario
'------------------------------------------------------------------------------*
Public Property Get TipoBombo() As Integer
    TipoBombo = mTipoBombo
End Property

Public Property Let TipoBombo(ByVal vNewValue As Integer)
    mTipoBombo = vNewValue
End Property



'------------------------------------------------------------------------------*
' Procedimiento  : TipoGiros
' Fecha          : 10/may/2018
' Propósito      : Tipo de función para girar el bombo: 1-Por Tiempo 2-Por Giros
'------------------------------------------------------------------------------*
Public Property Get TipoGiros() As Integer
    TipoGiros = mTipoGiros
End Property

Public Property Let TipoGiros(ByVal vNewValue As Integer)
    If vNewValue < 1 Or vNewValue > 2 Then
        Err.Raise ERR_VALTIPOGIRO, "BomboV2", MSG_VALTIPOGIRO
    End If
    mTipoGiros = vNewValue
End Property
'------------------------------------------------------------------------------*
' Procedimiento  : Situacion
' Fecha          : 12/may/2018
' Propósito      : Indicador de estado del bombo: 1-Vacio, 2-Cargado, 3-Rotando
'------------------------------------------------------------------------------*
Public Property Get Situacion() As Integer
    Situacion = mSituacion
End Property
'------------------------------------------------------------------------------*
' Procedimiento  : Cargado
' Fecha          : 12/may/2018
' Propósito      : Indicador de bombo cargado con bolas
'------------------------------------------------------------------------------*
Public Property Get Cargado() As Boolean
    If mCol.Count > 0 Then
        Cargado = True
    Else
        Cargado = False
    End If
End Property

'------------------------------------------------------------------------------*
' Procedimiento  : ProbabilidadesBolas
' Fecha          : 12/may/2018
' Propósito      : Matriz de probabilidades para cada bola (numero)
'------------------------------------------------------------------------------*
Public Property Get ProbabilidadesBolas() As Variant
    ProbabilidadesBolas = mProbNum
End Property

Public Property Let ProbabilidadesBolas(ByVal vNewValue As Variant)
'
'   cOMPROBAR QUE ES UNA MATRIZ DE NUMEROS
'
    If Not IsArray(vNewValue) Then
        Err.Raise 46545, "Bombo.ProbabilidadesBolas", "el argumeto debe ser una matriz de dos dimensiones"
    End If
    '
    '   Comprobar que la segunda dimensión son dos
    '
    If UBound(vNewValue, 2) <> 1 Then
        Err.Raise 12546, "Bombo.ProbabilidadesBolas", "el argumeto debe ser una matriz de dos dimensiones"
    End If
        
    mProbNum = vNewValue
    If mSituacion <> lotVacio Then
        SetProbToBolas
        mSituacion = lotCargado
    End If
End Property
'------------------------------------------------------------------------------*
' Procedimiento  : TiempoGiro
' Fecha          : 13/may/2018
' Propósito      :
'------------------------------------------------------------------------------*
Public Property Get TiempoGiro() As Date
    TiempoGiro = mTiempoGiro
End Property

Public Property Let TiempoGiro(ByVal vNewValue As Date)
    mTiempoGiro = vNewValue
End Property
'------------------------------------------------------------------------------*
' Procedimiento  : NumGiros
' Fecha          : 13/may/2018
' Propósito      :
'------------------------------------------------------------------------------*
Public Property Get NumGiros() As Integer
    NumGiros = mNumGiros
End Property

Public Property Let NumGiros(ByVal vNewValue As Integer)
    mNumGiros = vNewValue
End Property

'------------------------------------------------------------------------------*
' Procedimiento  : SetProbToBolas
' Fecha          : vi., 22/may/2020 16:38:36
' Propósito      : Cargar las probabilidad de cada bola
'------------------------------------------------------------------------------*
Private Sub SetProbToBolas()
    Dim i As Integer
    Dim mKey As String
    Dim mValue As Double
    Dim mBola As Variant
    
  On Error GoTo SetProbToBolas_Error
    If mCol.Count = 0 Then
        Exit Sub
    End If
    Set mBolasCargadas = New Collection
    '
    '
    '
    For Each mBola In mCol
        mKey = CStr(mBola.Numero.Valor)
        For i = 0 To UBound(mProbNum)
            If mProbNum(i, 0) = mKey Then
                mValue = mProbNum(i, 1)
                Exit For
            End If
        Next i
        If mBolasCargadas.Count = 0 Then
            mBolasCargadas.Add Array(mKey, mValue), mKey
        Else
            mBolasCargadas.Add Array(mKey, mValue), mKey, mBolasCargadas.Count
        End If
    Next mBola
    
'    '
'    '   cargamos las probabilidades de las bolas
'    '
'    For i = 0 To UBound(mProbNum, 1)
'        '
'        '   Obtenemos el número
'        '
'        mKey = CStr(mProbNum(i, 0))
'        '
'        '   Obtenemos su probabilidad
'        '
'        mValue = CDbl(mProbNum(i, 1))
'        If ExistenElem(mCol, mKey) Then
'            If mBolasCargadas.Count = 0 Then
'                mBolasCargadas.Add Array(mKey, mValue), mKey
'            Else
'                mBolasCargadas.Add Array(mKey, mValue), mKey, mBolasCargadas.Count
'            End If
'        End If
'    Next i
  On Error GoTo 0
SetProbToBolas__CleanExit:
    Exit Sub
            
SetProbToBolas_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    Call HandleException(ErrNumber, ErrDescription, "BomboV2.PARRAFO", ErrSource)
    Err.Raise ErrNumber, "File.PARRAFO", ErrDescription
End Sub
'' *===========(EOF): BomboV2.cls

