VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Individuo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' *============================================================================*
' *
' *     Fichero    : Individuo.cls
' *
' *     Autor      : Carlos Almela Baeza
' *     Creación   : Dom, 06/may/2018 23:57:06
' *     Versión    : 1.0
' *     Propósito  : Clase que almacena la información de un Individuo de la
' *                  poblacion
' *
' *============================================================================*
Option Explicit
Option Base 0

'
'--- Variables Privadas -------------------------------------------------------*
Dim mId             As String
Dim mIdPoblacion    As String
Dim mFechaCreacion  As Date
Dim mFitness        As Double
Dim mGeneracion     As Integer
Dim mJuego          As Juego
Dim mCombinacion    As Combinacion
Dim mEstrellas      As Combinacion
Dim mViable         As Boolean
Dim mMutado         As Boolean
'--- Constantes ---------------------------------------------------------------*
Const LT_INDIVIDUO = "Individuo:$1, F.Creacion:$2, Juego:$3, Combinacion:$4, " & _
                     "Viable:$5, Mutado:$6, Poblacion:$7, Generacion:$8, Fitness:$9"
Const LT_SINGENOMA = "#Sin combinaciones"
'--- Mensajes -----------------------------------------------------------------*
'--- Errores ------------------------------------------------------------------*
'--- Metodos Privados ---------------------------------------------------------*
Private Sub Class_Initialize()
    Dim nSec As Integer
    nSec = CInt(Rnd() * 1000)
    Set mCombinacion = New Combinacion
    Set mEstrellas = New Combinacion
    mFechaCreacion = Now
    mId = "IND" & Format(mFechaCreacion, "yyyyMMddThhmmss") & Format(nSec, "00#")
    mFitness = 0
    mIdPoblacion = ""
    mViable = True
    mMutado = False
End Sub

Private Sub Class_Terminate()
    Set mCombinacion = Nothing
    Set mEstrellas = Nothing
End Sub

'--- Metodos Publicos ---------------------------------------------------------*
'------------------------------------------------------------------------------*
' Procedimiento  : ToString
' Fecha          : 25/may/2018
' Propósito      : Visualizar el individuo como una cadena de texto
' Retorno        : Expresión Individuo
'------------------------------------------------------------------------------*
'
Public Function ToString() As String
    Dim mTmp        As String
 On Error GoTo ToString_Error
    mTmp = Replace(LT_INDIVIDUO, "$1", mId)
    mTmp = Replace(mTmp, "$2", Format(mFechaCreacion, "dd/MM/yyyy hh:mm"))
    mTmp = Replace(mTmp, "$3", mJuego)
    mTmp = Replace(mTmp, "$4", GetGenoma)
    mTmp = Replace(mTmp, "$5", mViable)
    mTmp = Replace(mTmp, "$6", mMutado)
    mTmp = Replace(mTmp, "$7", mIdPoblacion)
    mTmp = Replace(mTmp, "$8", mGeneracion)
    mTmp = Replace(mTmp, "$9", Format(mFitness, "#0,##"))
                
    ToString = mTmp
 On Error GoTo 0
ToString__CleanExit:
    Exit Function
            
ToString_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    '   Audita el error
    Call HandleException(ErrNumber, ErrDescription, "Individuo.ToString", ErrSource)
    '   Lanza el Error
    Err.Raise ErrNumber, "Individuo.ToString", ErrDescription
End Function
'------------------------------------------------------------------------------*
' Procedimiento  : GetGenoma
' Fecha          : 25/may/2018
' Propósito      : Devuelve los números que componen el individuo
' Retorno        : Numeros de la combinación
'------------------------------------------------------------------------------*
'
Private Function GetGenoma() As String
    Dim mTmp        As String
    Dim mNum        As Numero
 On Error GoTo GetGenoma_Error
    '
    '   Serializamos los numeros
    '
    mTmp = mCombinacion.Texto
    '
    '   Si el juego es euromillon agregamos estrellas
    '
    If mJuego = Euromillones Then
        mTmp = mTmp & " E-" & mEstrellas.Texto
    End If
    '
    '   Si no hay datos ponemos el literal vacio
    '
    If Len(mTmp) = 0 Then
        mTmp = LT_SINGENOMA
    End If
    
    GetGenoma = mTmp
 On Error GoTo 0
GetGenoma__CleanExit:
    Exit Function
            
GetGenoma_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    '   Audita el error
    Call HandleException(ErrNumber, ErrDescription, "Individuo.GetGenoma", ErrSource)
    '   Lanza el Error
    Err.Raise ErrNumber, "Individuo.GetGenoma", ErrDescription
    
End Function

'--- Propiedades --------------------------------------------------------------*


'------------------------------------------------------------------------------*
' Procedimiento  : Id
' Fecha          : 25/may/2018
' Propósito      : Identificador único de individuo
'------------------------------------------------------------------------------*
'
Public Property Get Id() As String
    Id = mId
End Property

Public Property Let Id(ByVal vNewValue As String)
    mId = vNewValue
End Property

'------------------------------------------------------------------------------*
' Procedimiento  : IdPoblacion
' Fecha          : 25/may/2018
' Propósito      : Identificador de poblacion en la que pertenece
'------------------------------------------------------------------------------*
'
Public Property Get IdPoblacion() As String
    IdPoblacion = mIdPoblacion
End Property

Public Property Let IdPoblacion(ByVal vNewValue As String)
    mIdPoblacion = vNewValue
End Property

'------------------------------------------------------------------------------*
' Procedimiento  : Generacion
' Fecha          : 25/may/2018
' Propósito      : Numero de generación en la que se crea el individuo
'------------------------------------------------------------------------------*
'
Public Property Get Generacion() As Integer
    Generacion = mGeneracion
End Property

Public Property Let Generacion(ByVal vNewValue As Integer)
    mGeneracion = vNewValue
End Property

'------------------------------------------------------------------------------*
' Procedimiento  : Juego
' Fecha          : 25/may/2018
' Propósito      : Juego al que pertenece el individuo
'------------------------------------------------------------------------------*
'
Public Property Get Juego() As Juego
    Juego = mJuego
End Property

Public Property Let Juego(ByVal vNewValue As Juego)
    mJuego = vNewValue
End Property

'------------------------------------------------------------------------------*
' Procedimiento  : FechaCreacion
' Fecha          : 25/may/2018
' Propósito      : Instante en el que se crea el Individuo
'------------------------------------------------------------------------------*
'
Public Property Get FechaCreacion() As Date
    FechaCreacion = mFechaCreacion
End Property

Public Property Let FechaCreacion(ByVal vNewValue As Date)
    mFechaCreacion = vNewValue
End Property

'------------------------------------------------------------------------------*
' Procedimiento  : Genoma
' Fecha          : 25/may/2018
' Propósito      : Numeros que componen el genoma del individuo
'------------------------------------------------------------------------------*
'
Public Property Get Genoma() As Combinacion
    Set Genoma = mCombinacion
End Property

Public Property Set Genoma(ByVal vNewValue As Combinacion)
    Set mCombinacion = vNewValue
End Property


'------------------------------------------------------------------------------*
' Procedimiento  : Estrellas
' Fecha          : 25/may/2018
' Propósito      : Combinación de estrellas sugerencia para el juego euromillon
'------------------------------------------------------------------------------*
'
Public Property Get Estrellas() As Combinacion
    Set Estrellas = mEstrellas
End Property

Public Property Set Estrellas(ByVal vNewValue As Combinacion)
    Set mEstrellas = vNewValue
End Property

'------------------------------------------------------------------------------*
' Procedimiento  : Viable
' Fecha          : 25/may/2018
' Propósito      : Indicador de individuo valido, sin errores
'------------------------------------------------------------------------------*
'
Public Property Get Viable() As Boolean
    Viable = mViable
End Property

Public Property Let Viable(ByVal vNewValue As Boolean)
    mViable = vNewValue
End Property

'------------------------------------------------------------------------------*
' Procedimiento  : Mutado
' Fecha          : 25/may/2018
' Propósito      : Indicador de individuo sometido a mutación
'------------------------------------------------------------------------------*
'
Public Property Get Mutado() As Boolean
    Mutado = mMutado
End Property

Public Property Let Mutado(ByVal vNewValue As Boolean)
    mMutado = vNewValue
End Property

'------------------------------------------------------------------------------*
' Procedimiento  : Fitness
' Fecha          : 25/may/2018
' Propósito      : Puntuación de la sugerencia frente al problema
'------------------------------------------------------------------------------*
'
Public Property Get Fitness() As Double
    Fitness = mFitness
End Property

Public Property Let Fitness(ByVal vNewValue As Double)
    mFitness = vNewValue
End Property
' *===========(EOF): Individuo.cls

