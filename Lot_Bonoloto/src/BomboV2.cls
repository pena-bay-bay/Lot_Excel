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
' *     Creaci�n   : mar, 08/may/2018 06:43:15
' *     Versi�n    : 1.0
' *     Prop�sito  : Clase que simula un Bombo
' *
' *============================================================================*
Option Explicit
Option Base 0
'
'--- Variables Privadas -------------------------------------------------------*
Dim mJuego         As Juego               ' Tipo de juego para que se carga el bombo
Dim mCol           As Collection          ' Colecci�n de Bolas del Bombo
Dim mColSel        As Collection          ' Colecci�n de Bolas extraidas del b�mbo
Dim mSituacion     As Integer             ' Situaci�n del Bombo: 1-Vacio, 2-lleno, 3-Cargado
Dim mTipoBombo     As Integer             ' Tipo de Bombo: 1-Principal, 2-Secundario
Dim mTipoGiros     As Integer             ' Tipo de Giro: 1-Por tiempo, 2-Por giros
Dim mBola          As bola                ' Bola
Dim mMaxBolas      As Integer             ' numero m�nimo del juego
Dim mMinBolas      As Integer             ' numero m�ximo del juego
Dim i              As Integer             ' contador gen�rico
Dim mKey           As String              ' clave del n�mero en la colecci�n
Dim mNumGiros      As Integer             ' Numero de giros realizados en el bombo
Dim mTiempoGiro    As Date                ' Duraci�n de los giros
Dim mProbNum       As Variant             ' Matriz de probabilidades
'--- Constantes ---------------------------------------------------------------*

'--- Mensajes -----------------------------------------------------------------*
Private Const MSG_VALTIPOGIRO = "Bombo: Error en Tipo de giro debe ser 1 � 2."
Private Const MSG_BOMBOVACIO = "#Error: Bombo vac�o."

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
End Sub
'------------------------------------------------------------------------------*
' Procedimiento  : RealizarGiro
' Fecha          : 13/may/2018
' Prop�sito      : Realiza un giro del b�mbo
'------------------------------------------------------------------------------*
Private Sub RealizarGiro()
    Dim mIdxA  As Integer
    Dim mIdxB  As Integer
   
 On Error GoTo RealizarGiro_Error
    '
    '   seleccionamos un �ndice origen al azar
    '
    mIdxA = GetIndex()
    mIdxB = 0
    '
    '   seleccionamos un �ndice destino al azar
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
    '   Audita el error
    Call HandleException(ErrNumber, ErrDescription, "BomboV2.RealizarGiro", ErrSource)
    '   Lanza el Error
    Err.Raise ErrNumber, "BomboV2.RealizarGiro", ErrDescription
End Sub
'------------------------------------------------------------------------------*
' Procedimiento  : GetIndex
' Fecha          : 13/may/2018
' Prop�sito      : Devuelve un indice de la colecci�n aleatoriamente
'------------------------------------------------------------------------------*
Private Function GetIndex() As Integer
    Dim mxBolas  As Integer
    Static b_rand As Boolean
    '
    '
    '
    If Not b_rand Then          'La primera vez que se ejecuta
        b_rand = True           'la funci�n RND se ceba la
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
    '   Si el puntero es 0 pasamas al �ltimo numero
    '
    If GetIndex = 0 Then
        GetIndex = mxBolas
    ElseIf GetIndex >= mxBolas Then
        GetIndex = 1
    End If
End Function



'--- Metodos Publicos ---------------------------------------------------------*
'------------------------------------------------------------------------------*
' Procedimiento  : Cargar
' Fecha          : 12/may/2018
' Prop�sito      : Cargar las Bolas en el Bombo segun el juego
' Par�metros     :
' Retorno        :
'------------------------------------------------------------------------------*
Public Sub Cargar()
    
 On Error GoTo Cargar_Error
    '
    '   Si el bombo est� cargado salimos de la rutina
    '
    If mSituacion = lotCargado Then
        Exit Sub
    End If
    '
    '   Deterninamos en numero de bolas del bombo
    '
    Select Case mJuego
        Case Bonoloto, LoteriaPrimitiva
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
            '   asignamos el n�mero
            '
            mBola.Numero.Valor = i
            '
            '   agregamos bola
            '
            mCol.Add mBola, mKey
         
        Next i
        '
        '   Actualizamos la situaci�n del bombo
        '
        mSituacion = lotCargado
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
    '   Audita el error
    Call HandleException(ErrNumber, ErrDescription, "BomboV2.Cargar", ErrSource)
    '   Lanza el Error
    Err.Raise ErrNumber, "BomboV2.Cargar", ErrDescription
    '   Informa del error
    Call MsgBox(ErrDescription, vbCritical Or vbSystemModal, ThisWorkbook.Name)
    Call Trace("CERRAR")
End Sub
'------------------------------------------------------------------------------*
' Procedimiento  : Reiniciar
' Fecha          : 12/may/2018
' Prop�sito      : R
' Par�metros     :
' Retorno        :
'------------------------------------------------------------------------------*
Public Sub Reiniciar()
 On Error GoTo Reiniciar_Error
    '
    '   Si el bombo est� vacio salimos de la rutina
    '
    If mSituacion = lotVacio Then
        Exit Sub
    End If
    '
    '   Para cada bola en la colecci�n de seleccionadas
    '
    For Each mBola In mColSel
        '
        '   Calculamos la clave de la bola
        '
        mKey = CStr(mBola.Numero.Valor)
        '
        '   Agregamos la bola a la colecci�n
        '
        mCol.Add mBola, mKey
    Next mBola
    '
    '   Inicializamos Bolas seleccionadas
    '
    EliminarTodosElementos mColSel
 
 On Error GoTo 0
Reiniciar__CleanExit:
    Exit Sub
            
Reiniciar_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    '   Audita el error
    Call HandleException(ErrNumber, ErrDescription, "BomboV2.Reiniciar", ErrSource)
    '   Lanza el Error
    Err.Raise ErrNumber, "BomboV2.Reiniciar", ErrDescription
    '   Informa del error
    Call MsgBox(ErrDescription, vbCritical Or vbSystemModal, ThisWorkbook.Name)
    Call Trace("CERRAR")
End Sub
'------------------------------------------------------------------------------*
' Procedimiento  : Girar
' Fecha          : 12/may/2018
' Prop�sito      :
' Par�metros     :
' Retorno        : Gira el bombo seg�n los criterios establecidos de tiempo o
'                  n�mero de giros
'------------------------------------------------------------------------------*
Public Sub Girar()
    Dim mTFin                 As Date
    Dim mGirosRealizados      As Long
    
 On Error GoTo Girar_Error
    '
    '   Si el bombo est� vacio sale de la rutina
    '
    If mSituacion = lotVacio Then
        Exit Sub
    End If
    '
    '   inicializamos el n�mero de giros del bombo
    '
    mGirosRealizados = 0
    '
    '   En funci�n del tipo de giro
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
    '   Audita el error
    Call HandleException(ErrNumber, ErrDescription, "BomboV2.Girar", ErrSource)
    '   Lanza el Error
    Err.Raise ErrNumber, "BomboV2.Girar", ErrDescription
End Sub
'------------------------------------------------------------------------------*
' Procedimiento  : ExtraerBolas
' Fecha          : 12/may/2018
' Prop�sito      :
' Par�metros     :
' Retorno        : Cadena de texto con los n�meros elegidos
'------------------------------------------------------------------------------*
Public Function ExtraerBolas(ByVal vNewValue As Integer) As String
    Dim mIdxA       As Integer
    
 On Error GoTo ExtraerBolas_Error
    '
    '   Si el bombo est� vacio devolvemos un literal
    '
    If mSituacion = lotVacio Then
        ExtraerBolas = MSG_BOMBOVACIO
        Exit Function
    End If
    For i = 1 To vNewValue
        '
        '   Seleccionamos una bola
        '
        mIdxA = GetIndex()
        Set mBola = mCol.Item(mIdxA)
        mCol.Remove mIdxA
        '
        '   Agregamos la bola a la colecci�n de seleccionadas
        '
        mColSel.Add mBola
    
    Next i
    ExtraerBolas = ""
    '
    '   Montamos la cadena de bolas seleccionadas
    '
    For Each mBola In mColSel
        ExtraerBolas = ExtraerBolas & CStr(mBola.Numero.Valor) & "-"
    Next mBola
    '
    '   Eliminamos el �ltimo guion
    '
    ExtraerBolas = Left(ExtraerBolas, Len(ExtraerBolas) - 1)
'    '
'    '
'    Err.Raise ERR_TODO, "BomboV2.ExtraerBolas", MSG_TODO
'    '
 On Error GoTo 0
ExtraerBolas__CleanExit:
    Exit Function
            
ExtraerBolas_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    '   Audita el error
    Call HandleException(ErrNumber, ErrDescription, "BomboV2.ExtraerBolas", ErrSource)
    '   Lanza el Error
    Err.Raise ErrNumber, "BomboV2.ExtraerBolas", ErrDescription
    '   Informa del error
    Call MsgBox(ErrDescription, vbCritical Or vbSystemModal, ThisWorkbook.Name)
    Call Trace("CERRAR")

End Function


'--- Propiedades --------------------------------------------------------------*
'------------------------------------------------------------------------------*
' Procedimiento  : Juego
' Fecha          : 10/may/2018
' Prop�sito      : Definir el juego para el que se carga el bombo
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
' Prop�sito      : Numero de bolas del Bombo
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
' Prop�sito      : Colecci�n de Bolas del Bombo
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
' Prop�sito      : Tipo de Bombo seg�n el juego: 1 Bombo principal,
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
' Prop�sito      : Tipo de funci�n para girar el bombo: 1-Por Tiempo 2-Por Giros
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
' Prop�sito      : Indicador de estado del bombo: 1-Vacio, 2-Cargado, 3-Rotando
'------------------------------------------------------------------------------*
Public Property Get Situacion() As Integer
    Situacion = mSituacion
End Property
'------------------------------------------------------------------------------*
' Procedimiento  : Cargado
' Fecha          : 12/may/2018
' Prop�sito      : Indicador de bombo cargado con bolas
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
' Prop�sito      : Matriz de probabilidades para cada bola (numero)
'------------------------------------------------------------------------------*
Public Property Get ProbabilidadesBolas() As Variant
'   Comprobar si es matriz y tiene los mismos elementos que el bombo
'   Asignar a cada elemento el valor de la propiedad
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
    '   Comprobar que la segunda dimensi�n son dos
    '
    If UBound(vNewValue, 2) <> 1 Then
        Err.Raise 12546, "Bombo.ProbabilidadesBolas", "el argumeto debe ser una matriz de dos dimensiones"
    End If
        
    mProbNum = vNewValue
End Property
'------------------------------------------------------------------------------*
' Procedimiento  : TiempoGiro
' Fecha          : 13/may/2018
' Prop�sito      :
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
' Prop�sito      :
'------------------------------------------------------------------------------*
Public Property Get NumGiros() As Integer
    NumGiros = mNumGiros
End Property

Public Property Let NumGiros(ByVal vNewValue As Integer)
    mNumGiros = vNewValue
End Property

'' *===========(EOF): BomboV2.cls
