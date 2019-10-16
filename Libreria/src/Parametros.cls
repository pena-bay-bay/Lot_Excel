VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Parametros"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' *============================================================================*
' *
' *     Fichero    : Parametros.cls
' *
' *     Autor      : Carlos Almela Baeza
' *     Creación   : Mie, 20/Dic/2018 00:03
' *     Versión    : 1.0
' *     Propósito  : Colección de parametros de aplicación
' *
' *============================================================================*
Option Explicit
Option Base 0
'
'--- Variables Privadas -------------------------------------------------------*
Private mCol As Collection
'--- Constantes ---------------------------------------------------------------*
'--- Mensajes -----------------------------------------------------------------*
'--- Errores ------------------------------------------------------------------*
'--- Metodos Privados ---------------------------------------------------------*
Private Sub Class_Initialize()
    Set mCol = New Collection
End Sub

Private Sub Class_Terminate()
    Set mCol = Nothing
End Sub
'--- Metodos Publicos ---------------------------------------------------------*
'------------------------------------------------------------------------------*
' Procedimiento  : Add
' Fecha          : 13/01/2018
' Propósito      : Agrega un parámetro a la colección
' Parámetros     : datParametro
'------------------------------------------------------------------------------*
Public Sub Add(datParametro As Parametro)
    Dim strKey  As String
  On Error GoTo Add_Error
    '
    '   Calculamos la clave
    '
    strKey = CStr(datParametro.Id)
    '
    '   Si ya tenemos elementos agregamos al final
    '
    If mCol.Count <> 0 Then
        mCol.Add datParametro, strKey, , mCol.Count
    Else
        mCol.Add datParametro, strKey
    End If

  On Error GoTo 0
   Exit Sub

Add_Error:
   Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
   ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
   '   Audita el error
   Call HandleException(ErrNumber, ErrDescription, "Parametros.Add")
   '   Lanza el Error
   Err.Raise ErrNumber, ErrSource, ErrDescription
End Sub

'------------------------------------------------------------------------------*
' Procedimiento  : Delete
' Fecha          : 13/01/2019
' Propósito      :
' Parámetros     :
' Retorno        :
'------------------------------------------------------------------------------*
Public Sub Delete(datIndex As Integer)
    Dim strError As String
    
  On Error GoTo Delete_Error
    If datIndex > mCol.Count Then
        strError = Replace(MSG_DELETEINDEXERROR, "#", "#" & datIndex)
        Err.Raise ERR_DELETEINDEXERROR, "Parametros.Delete", strError
    Else
        mCol.Remove datIndex
    End If
    
  On Error GoTo 0
    Exit Sub

Delete_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    '   Audita el error
    Call HandleException(ErrNumber, ErrDescription, "Parametros.Delete")
    '   Lanza el Error
    Err.Raise ErrNumber, ErrSource, ErrDescription
End Sub

'------------------------------------------------------------------------------*
' Procedimiento  : MarkForDelete
' Fecha          : 13/01/2019
' Propósito      : Marcar un elemento para borrar de la BBDD
' Parámetros     : Indice del elemento a borrar
' Retorno        :
'------------------------------------------------------------------------------*
Public Sub MarkForDelete(datIndex As Integer)
    Dim strError As String
    Dim mPar As Parametro
    
   On Error GoTo MarkForDelete_Error
    If datIndex > mCol.Count Then
        strError = Replace(MSG_INDEXERROR, "#", "#" & datIndex)
        Err.Raise ERR_INDEXERROR, "Parametros.MarkForDelete", strError
    Else
        Set mPar = mCol.Item(datIndex)
        mPar.EntidadNegocio.MarkForDelete = True
    End If
    
   On Error GoTo 0
   Exit Sub

MarkForDelete_Error:
   Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
   ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
   '   Audita el error
   Call HandleException(ErrNumber, ErrDescription, "Parametros.MarkForDelete")
   '   Lanza el Error
   Err.Raise ErrNumber, ErrSource, ErrDescription
End Sub

'------------------------------------------------------------------------------*
' Procedimiento  : Undelete
' Fecha          : 13/01/2018
' Propósito      : Desmarcar un elemento para borrar
' Parámetros     : Indice de la colección
'------------------------------------------------------------------------------*
Public Sub Undelete(datIndex As Integer)
    Dim strError As String
    Dim mPar As Parametro
   On Error GoTo Undelete_Error
    
    If datIndex > mCol.Count Then
        strError = Replace(MSG_INDEXERROR, "#", "#" & datIndex)
        Err.Raise ERR_INDEXERROR, "Parametros.Undelete", strError
    Else
        Set mPar = mCol.Item(datIndex)
        mPar.EntidadNegocio.MarkForDelete = False
    End If
   
   On Error GoTo 0
   Exit Sub

Undelete_Error:
   Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
   ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
   '   Audita el error
   Call HandleException(ErrNumber, ErrDescription, "Parametros.Undelete")
   '   Lanza el Error
   Err.Raise ErrNumber, ErrSource, ErrDescription
End Sub



'------------------------------------------------------------------------------*
' Procedimiento  : Clear
' Fecha          : 13/01/2019
' Propósito      : Elminar todos los elementos de la colección
'------------------------------------------------------------------------------*
Public Sub Clear()
   On Error GoTo Clear_Error
    Set mCol = Nothing
    Set mCol = New Collection
    
   On Error GoTo 0
   Exit Sub

Clear_Error:
   Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
   ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
   '   Audita el error
   Call HandleException(ErrNumber, ErrDescription, "Parametros.Clear")
   '   Lanza el Error
   Err.Raise ErrNumber, ErrSource, ErrDescription
End Sub
'--- Propiedades --------------------------------------------------------------*
'------------------------------------------------------------------------------*
' Procedimiento  : Count
' Fecha          : 06/08/2014
' Propósito      : Devuelve el número de parametros de la colección
' Retorno        : Numero de elementos
'------------------------------------------------------------------------------*
Public Property Get Count() As Integer

   On Error GoTo Count_Error

    Count = mCol.Count

   On Error GoTo 0
   Exit Property

Count_Error:
   Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
   ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
   '   Audita el error
   Call HandleException(ErrNumber, ErrDescription, "Parametros.Count")
   '   Lanza el Error
   Err.Raise ErrNumber, ErrSource, ErrDescription
End Property

'------------------------------------------------------------------------------*
' Procedimiento  : Items
' Fecha          : 08/01/2019
' Propósito      : Devolver la colección interna
' Retorno        : Colección de Objetos
'------------------------------------------------------------------------------*
Public Property Get Items() As Collection
    Set Items = mCol
End Property

Public Property Let Items(ByVal vNewValue As Collection)
    Set mCol = vNewValue
End Property

'' *===========(EOF): Parametros.cls

