VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Poblacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' *============================================================================*
' *
' *     Fichero    : Poblacion.cls
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
'--- Variables Privadas -------------------------------------------------------*
Dim mId             As String
Dim mFechaCreacion  As Date
Dim mGeneracion     As Integer
Dim mJuego          As Juego
Dim mCol            As Collection
Dim mMaxFitness     As Double
'--- Constantes ---------------------------------------------------------------*
Const LT_POBLACION = "Poblacion:$1, F.Creacion:$2, Juego:$3, Individuos:$4, Generacion:$5, MaxFitness:$6"
'--- Mensajes -----------------------------------------------------------------*
'--- Errores ------------------------------------------------------------------*
'--- Metodos Privados ---------------------------------------------------------*
Private Sub Class_Initialize()
    Set mCol = New Collection
    mFechaCreacion = Now
    mId = "POB" & Format(mFechaCreacion, "yyyyMMddThhmmss")
    mMaxFitness = 0
End Sub

Private Sub Class_Terminate()
    Set mCol = Nothing
End Sub
'--- Metodos Publicos ---------------------------------------------------------*
'------------------------------------------------------------------------------*
' Procedimiento  : ToString
' Fecha          : 26/may/2018
' Propósito      : Visualizar la población como una cadena de texto
' Retorno        : Expresión Población
'------------------------------------------------------------------------------*
'
Public Function ToString() As String
    Dim mTmp        As String
 On Error GoTo ToString_Error
    mTmp = Replace(LT_POBLACION, "$1", mId)
    mTmp = Replace(mTmp, "$2", Format(mFechaCreacion, "dd/MM/yyyy hh:mm"))
    mTmp = Replace(mTmp, "$3", mJuego)
    mTmp = Replace(mTmp, "$4", NumIndividuos)
    mTmp = Replace(mTmp, "$5", mGeneracion)
    mTmp = Replace(mTmp, "$6", Format(MaxFitness, "#0.##"))
                
    ToString = mTmp
 On Error GoTo 0
ToString__CleanExit:
    Exit Function
            
ToString_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    '   Audita el error
    Call HandleException(ErrNumber, ErrDescription, "Poblacion.ToString", ErrSource)
    '   Lanza el Error
    Err.Raise ErrNumber, "Poblacion.ToString", ErrDescription
End Function
'------------------------------------------------------------------------------*
' Procedimiento  : SortByFitness
' Fecha          : 26/may/2018
' Propósito      : Ordena la población por Fitness de Mayor a menor
'------------------------------------------------------------------------------*
'
Public Sub SortByFitness()
    Dim mTmpCol As Collection
    Dim mMatriz As Variant
    
 On Error GoTo SortByFitness_Error
    
    
    
 On Error GoTo 0
SortByFitness__CleanExit:
    Exit Sub
            
SortByFitness_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    '   Audita el error
    Call HandleException(ErrNumber, ErrDescription, "Poblacion.SortByFitness", ErrSource)
    '   Lanza el Error
    Err.Raise ErrNumber, "Poblacion.SortByFitness", ErrDescription
End Sub
'------------------------------------------------------------------------------*
' Procedimiento  : Add
' Fecha          : 26/may/2018
' Propósito      : Agrega un individuo a la población
'------------------------------------------------------------------------------*
'
Public Sub Add(vNewValue As Individuo)
    
 On Error GoTo Add_Error
    '
    '   Asignamos el Id de población
    '
    vNewValue.IdPoblacion = mId
    vNewValue.Generacion = mGeneracion
    '
    '
    '
    mCol.Add vNewValue
    
 On Error GoTo 0
Add__CleanExit:
    Exit Sub
            
Add_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    '   Audita el error
    Call HandleException(ErrNumber, ErrDescription, "Poblacion.Add", ErrSource)
    '   Lanza el Error
    Err.Raise ErrNumber, "Poblacion.Add", ErrDescription
End Sub
'------------------------------------------------------------------------------*
' Procedimiento  : GetIndividuo
' Fecha          : 26/may/2018
' Propósito      : Obtiene el individuo enésimo de la población
' Retorno        : Individuo
'------------------------------------------------------------------------------*
'
Public Function GetIndividuo(vNewValue As Integer) As Individuo
    
 On Error GoTo GetIndividuo_Error
    '
    '   Comprobamos que el indice está dentro de los elementos
    '
    If vNewValue > mCol.Count Then
        Err.Raise ERR_IDXINDIVIDUO, "Poblacion.GetIndividuo", MSG_IDXINDIVIDUO
    End If
    '
    '   Devolvemos el elementos Iesimo
    '
    Set GetIndividuo = mCol.Item(vNewValue)
    
 On Error GoTo 0
GetIndividuo__CleanExit:
    Exit Function
            
GetIndividuo_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    '   Audita el error
    Call HandleException(ErrNumber, ErrDescription, "Poblacion.GetIndividuo", ErrSource)
    '   Lanza el Error
    Err.Raise ErrNumber, "Poblacion.GetIndividuo", ErrDescription
End Function
'--- Propiedades --------------------------------------------------------------*
'------------------------------------------------------------------------------*
' Procedimiento  : Id
' Fecha          : 26/may/2018
' Propósito      : Identificador único de la poblacion
'------------------------------------------------------------------------------*
'
Public Property Get Id() As String
    Id = mId
End Property

Public Property Let Id(ByVal vNewValue As String)
    mId = vNewValue
End Property
'------------------------------------------------------------------------------*
' Procedimiento  : Generacion
' Fecha          : 26/may/2018
' Propósito      : Numero de generación en la que está la poblacion
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
' Fecha          : 26/may/2018
' Propósito      : Juego al que pertenece la Poblacion
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
' Fecha          : 26/may/2018
' Propósito      : Instante en el que se crea la poblacion
'------------------------------------------------------------------------------*
'
Public Property Get FechaCreacion() As Date
    FechaCreacion = mFechaCreacion
End Property

Public Property Let FechaCreacion(ByVal vNewValue As Date)
    mFechaCreacion = vNewValue
End Property
'------------------------------------------------------------------------------*
' Procedimiento  : NumIndividuos
' Fecha          : 26/may/2018
' Propósito      : Numero de Individuos de la población
'------------------------------------------------------------------------------*
'
Public Property Get NumIndividuos() As Integer
    NumIndividuos = mCol.Count
End Property
'------------------------------------------------------------------------------*
' Procedimiento  : MaxFitness
' Fecha          : 26/may/2018
' Propósito      : Numero de Individuos de la población
'------------------------------------------------------------------------------*
'
Public Property Get MaxFitness() As Double
    Dim mIdv As Individuo
    '
    '   Si no tenemos el valor máximo lo calculamos
    '
    If mMaxFitness = 0 Then
        '
        '   Para cada individuo en la población
        '
        For Each mIdv In mCol
            '
            '   Comprobamos cual es su fitness
            '
            If mIdv.Fitness > mMaxFitness Then
                mMaxFitness = mIdv.Fitness
            End If
        Next mIdv
    End If
    '
    '   Devolvemos el máximo
    '
    MaxFitness = mMaxFitness
End Property
' *===========(EOF): Poblacion.cls


