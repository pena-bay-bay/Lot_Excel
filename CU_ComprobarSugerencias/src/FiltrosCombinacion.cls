VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FiltrosCombinacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' *============================================================================*
' *
' *     Fichero    : FiltrosCombinacion.cls
' *
' *     Autor      : Carlos Almela Baeza
' *     Creación   : vi., 03/abr/2020 17:36:05
' *     Modificado : lu., 11/may/2020 18:50:57
' *     Versión    : 1.1
' *     Propósito  : Colección de filtros de combiación
' *
' *============================================================================*
Option Explicit
Option Base 0
'
'--- Variables Privadas -------------------------------------------------------*
Private mCol As Collection          ' Colección de Filtros
'--- Constantes ---------------------------------------------------------------*
'--- Mensajes -----------------------------------------------------------------*
'--- Errores ------------------------------------------------------------------*
'--- Propiedades --------------------------------------------------------------*
Public Property Get Count() As Integer
    Count = mCol.Count
End Property
'------------------------------------------------------------------------------*
' Procedimiento  : Items
' Fecha          : 16/04/2020 10:22
' Propósito      : Devolver la colección interna
' Retorno        : Colección de Objetos
'------------------------------------------------------------------------------*
Public Property Get Items() As Collection
    Set Items = mCol
End Property

Public Property Let Items(ByVal vNewValue As Collection)
    Set mCol = vNewValue
End Property

'--- Métodos Privados ---------------------------------------------------------*
Private Sub Class_Initialize()
    Set mCol = New Collection
End Sub

Private Sub Class_Terminate()
    Set mCol = Nothing
End Sub
'--- Métodos Públicos ---------------------------------------------------------*
'------------------------------------------------------------------------------*
' Procedimiento  : ToString
' Fecha          : lu., 13/abr/2020 17:24:35
' Propósito      : Expresar la colección de filtros como una cadena de texto
'------------------------------------------------------------------------------*
Public Function ToString() As String
    Dim mStr As String
    Dim i As Integer
  On Error GoTo ToString_Error
    '
    '   incializamos cadena de filtros
    '
    mStr = ""
    '
    '
    '
    For i = 1 To mCol.Count
        If i <> mCol.Count Then
            mStr = mStr & mCol.Item(i).ToString & "|"
        Else
            mStr = mStr & mCol.Item(i).ToString
        End If
    Next i
    '
    '   Devolvemos la cadena
    '
    ToString = mStr
    
  On Error GoTo 0
ToString__CleanExit:
    Exit Function
ToString_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    Call HandleException(ErrNumber, ErrDescription, "FiltrosCombinacion.ToString", ErrSource)
    Err.Raise ErrNumber, "FiltrosCombinacion.ToString", ErrDescription
End Function
'------------------------------------------------------------------------------*
' Procedimiento  : Parse
' Fecha          :
' Propósito      : Traducir una cadena de filtros en una colección
' Parámetros     : Cadena de filtros
'------------------------------------------------------------------------------*
Public Sub Parse(datFiltros As String)
    Dim mVar As Variant
    Dim mStr As String
    Dim mFlt As FiltroCombinacion
    Dim i As Integer
    
  On Error GoTo Parse_Error
    '
    '  si el campo esta vacio inicializamos
    '
    If Len(Trim(datFiltros)) = 0 Then
        Me.Clear
        Exit Sub
    End If
    '
    '   Desglosamos en filtros
    '
    mVar = Split(datFiltros, "|")
    '
    '   para cada filtro en la expresión
    '
    For i = 0 To UBound(mVar)
        '
        '   creamos un nuevo objeto filtro
        '
        Set mFlt = New FiltroCombinacion
        '
        '   validamos el filtro
        '
        mStr = mVar(i)
        mFlt.Parse mStr
        '
        '   agregamos a la colección
        '
        Me.Add mFlt
    Next i
    '
    '   destruimos el objeto
    '
    Set mFlt = Nothing
  On Error GoTo 0
Parse__CleanExit:
    Exit Sub
Parse_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    Call HandleException(ErrNumber, ErrDescription, "FiltrosCombinacion.Parse", ErrSource)
    Err.Raise ErrNumber, "FiltrosCombinacion.Parse", ErrDescription
End Sub
'------------------------------------------------------------------------------*
' Procedimiento  : Add
' Fecha          : lu., 13/abr/2020 16:26:46
' Propósito      : Agrega un filtro a la colección de filtros
' Parámetros     : FiltroCombinación
'------------------------------------------------------------------------------*
Public Sub Add(datFiltro As FiltroCombinacion)
    Dim mKey    As String
    Dim mFlt    As FiltroCombinacion
  On Error GoTo Add_Error
    '
    '   Definimos la clave de almacenamiento
    '
    mKey = CStr(datFiltro.TipoFiltro)
    '
    '   Comprobamos si hay filtros
    '
    If mCol.Count = 0 Then
        mKey = datFiltro.TipoFiltro
        mCol.Add datFiltro, mKey
    Else
    '
    '   Si no hay filtros buscamos si hay un filtro ya de este tipo
    '
        If ExistenElem(mCol, mKey) Then
            '
            '   #TODO: comprobar que no existe el valor seleccionado
            '
            Set mFlt = mCol.Item(mKey)
            mFlt.AddFilterValue datFiltro.FilterValue
        Else
            mCol.Add datFiltro, mKey, mCol.Count
        End If
    End If
    
  On Error GoTo 0
Add__CleanExit:
    Exit Sub
Add_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    Call HandleException(ErrNumber, ErrDescription, "FiltrosCombinacion.Add", ErrSource)
    Err.Raise ErrNumber, "FiltrosCombinacion.Add", ErrDescription
End Sub
'------------------------------------------------------------------------------*
' Procedimiento  : Clear
' Fecha          : dd/MMM/yyyy
' Propósito      :
' Parámetros     :
'------------------------------------------------------------------------------*
Public Sub Clear()
    Set mCol = New Collection
End Sub


'------------------------------------------------------------------------------*
' Procedimiento  : Evaluar
' Fecha          : lu., 11/may/2020 18:50:12
' Propósito      : Evaluar una combinación en todos sus filtros
' Parámetros     : True, la combinación cumple los filtros
'------------------------------------------------------------------------------*
Public Function Evaluar(NewValue As Combinacion) As Boolean
    Dim mSelComb As Boolean
    Dim i As Integer
  On Error GoTo Evaluar_Error
    mSelComb = True
    For i = 1 To mCol.Count
        mSelComb = mSelComb And mCol.Item(i).EvaluarCombinacion(NewValue)
    Next i
    Evaluar = mSelComb
  On Error GoTo 0
Evaluar_CleanExit:
    Exit Function
Evaluar_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    Call HandleException(ErrNumber, ErrDescription, "FiltrosCombinacion.Evaluar", ErrSource)
    Err.Raise ErrNumber, "FiltrosCombinacion.Evaluar", ErrDescription
End Function

'' *===========(EOF): FiltrosCombinacion.cls
