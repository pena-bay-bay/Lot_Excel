VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Metodos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' *============================================================================*
' *
' *     Fichero    : Metodos.cls
' *
' *     Autor      : Carlos Almela Baeza
' *     Creación   : do., 08/mar/2020 00:02:14
' *     Versión    : 1.0
' *     Propósito  : Colección de la clase Metodo
' *
' *============================================================================*
Option Explicit
Option Base 0
'--- Variables Privadas -------------------------------------------------------*
Private mCol As Collection                  ' Usado para almacenar la Clase
'--- Constantes ---------------------------------------------------------------*
'--- Mensajes -----------------------------------------------------------------*
'--- Errores ------------------------------------------------------------------*
'--- Propiedades --------------------------------------------------------------*
'------------------------------------------------------------------------------*
' Procedimiento  : Count
' Fecha          : lu., 23/mar/2020 17:57:32
' Propósito      : Devuelve el número de elementos de la colección
'------------------------------------------------------------------------------*
Public Property Get Count() As Variant
   On Error GoTo Count_Error
    If mCol Is Nothing Then
        Count = 0
    Else
        Count = mCol.Count
    End If
   On Error GoTo 0
   Exit Property
Count_Error:
   Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
   ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
   Call HandleException(ErrNumber, ErrDescription, "Metodo.Count")
   Err.Raise ErrNumber, ErrSource, ErrDescription
End Property



'------------------------------------------------------------------------------*
' Procedimiento  : Item
' Fecha          : lu., 23/mar/2020 17:57:32
' Propósito      : Devuelve el Metodo indicado en el Indice
'------------------------------------------------------------------------------*
Public Property Get Item(ByVal Index As Variant) As Metodo
    Dim mintCodeID As Integer
   On Error GoTo Item_Error
    Set Item = Nothing
    If mCol Is Nothing Then
        GoTo Item_CleanExit
    End If
    If mCol.Count = 0 Then
        GoTo Item_CleanExit
    End If
    'Make Sure you cast the index to a Integer otherwise sets won't work
    If Trim(Index & "") = "" Or Index <= 0 Then
        'raise an error?
        GoTo Item_CleanExit
    Else
        mintCodeID = CInt(Index)
        Set Item = mCol.Item(mintCodeID)
    End If

Item_CleanExit:
   On Error GoTo 0
   Exit Property

Item_Error:
   Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
   ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
   Call HandleException(ErrNumber, ErrDescription, "Metodo.Item")
   Err.Raise ErrNumber, ErrSource, ErrDescription
End Property



'------------------------------------------------------------------------------*
' Procedimiento  : Items
' Fecha          : 02/04/2019
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
' Procedimiento  : Add
' Fecha          : lu., 23/mar/2020 17:54:26
' Propósito      : Agrega un método a la colección
' Parámetros     : Metodo a agregar
'------------------------------------------------------------------------------*
Public Sub Add(ByVal datMetodo As Metodo)
    Dim mKey As String
   On Error GoTo Add_Error
    '
    '   Obtenemos la clave del parametro
    '
    mKey = CStr(datMetodo.Id)
    '
    '   Agregamos a la colección con la clave y al final de la misma
    '
    If mCol.Count <> 0 Then
        mCol.Add datMetodo, mKey, , mCol.Count
    Else
        mCol.Add datMetodo, mKey
    End If
    
Add_CleanExit:
   On Error GoTo 0
   Exit Sub

Add_Error:
   Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
   ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
   Call HandleException(ErrNumber, ErrDescription, "Metodo.Add")
   Err.Raise ErrNumber, ErrSource, ErrDescription
End Sub



'------------------------------------------------------------------------------*
' Procedimiento  : Clear
' Fecha          : lu., 23/mar/2020 17:54:26
' Propósito      : Inicializa la colección
'------------------------------------------------------------------------------*
Public Sub Clear()
  On Error GoTo Clear_Error
    'Clear Collection
    Set mCol = Nothing
    'Create Collection
    Set mCol = New Collection
  On Error GoTo 0
   Exit Sub

Clear_Error:
   Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
   ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
   Call HandleException(ErrNumber, ErrDescription, "Metodo.Clear")
   Err.Raise ErrNumber, ErrSource, ErrDescription
End Sub



'------------------------------------------------------------------------------*
' Procedimiento  : MarkForDelete
' Fecha          : lu., 23/mar/2020 18:06:43
' Propósito      : Marcar un elemento para borrar de la BBDD
' Parámetros     : Indice del elemento a borrar
' Retorno        :
'------------------------------------------------------------------------------*
Public Sub MarkForDelete(datIndex As Integer)
    Dim strError As String
    Dim mMetodo As Metodo
    
   On Error GoTo MarkForDelete_Error
    If datIndex > mCol.Count Then
        strError = Replace(MSG_INDEXERROR, "#", "#" & datIndex)
        Err.Raise ERR_INDEXERROR, "Metodos.MarkForDelete", strError
    Else
        Set mMetodo = mCol.Item(datIndex)
        mMetodo.EntidadNegocio.MarkForDelete = True
    End If
    
   On Error GoTo 0
   Exit Sub

MarkForDelete_Error:
   Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
   ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
   Call HandleException(ErrNumber, ErrDescription, "Metodos.MarkForDelete")
   Err.Raise ErrNumber, ErrSource, ErrDescription
End Sub

'------------------------------------------------------------------------------*
' Procedimiento  : Undelete
' Fecha          : lu., 23/mar/2020 18:08:39
' Propósito      : Desmarcar un elemento para borrar
' Parámetros     : Indice de la colección
'------------------------------------------------------------------------------*
Public Sub Undelete(datIndex As Integer)
    Dim strError As String
    Dim mMetodo As Metodo
   On Error GoTo Undelete_Error
    
    If datIndex > mCol.Count Then
        strError = Replace(MSG_INDEXERROR, "#", "#" & datIndex)
        Err.Raise ERR_INDEXERROR, "Metodos.Undelete", strError
    Else
        Set mMetodo = mCol.Item(datIndex)
        mMetodo.EntidadNegocio.MarkForDelete = False
    End If
   
   On Error GoTo 0
   Exit Sub

Undelete_Error:
   Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
   ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
   Call HandleException(ErrNumber, ErrDescription, "Metodos.Undelete")
   Err.Raise ErrNumber, ErrSource, ErrDescription
End Sub


'------------------------------------------------------------------------------*
' Procedimiento  : Delete
' Fecha          : lu., 23/mar/2020 17:54:26
' Propósito      : Borra un sorteo de la colección
'------------------------------------------------------------------------------*
'
Public Sub Delete(datIndex As Integer)
    Dim strError As String
    
  On Error GoTo Delete_Error
    If datIndex > mCol.Count Then
        strError = Replace(MSG_DELETEINDEXERROR, "#", "#" & datIndex)
        Err.Raise ERR_DELETEINDEXERROR, "Sorteos.Delete", strError
    Else
        mCol.Remove datIndex
    End If
    
  On Error GoTo 0
    Exit Sub

Delete_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    Call HandleException(ErrNumber, ErrDescription, "Metodos.Delete")
    Err.Raise ErrNumber, ErrSource, ErrDescription
End Sub


'' *===========(EOF): Metodos.cls

