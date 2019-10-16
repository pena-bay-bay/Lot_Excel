VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sorteos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' *============================================================================*
' *
' *     Fichero    : Sorteos.cls
' *
' *     Autor      : Carlos Almela Baeza
' *     Creación   : Dom, 17/jun/2018 23:52:06
' *     Versión    : 1.0
' *     Propósito  : Colección de Sorteos
' *
' *
' *============================================================================*
Option Explicit
Option Base 0

'--- Variables Privadas -------------------------------------------------------*
Private mCol As Collection                ' Usado para almacenar la Clase
Private mJuego As Juego                   ' Juego al que pertenece la coleccion

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
'---------------------------------------------------------------------------------------
' Procedure : Add
' Author    : CHARLY
' Date      : 27/03/2014
' Purpose   : Agrega un Sorteo a la colección
'---------------------------------------------------------------------------------------
'
Public Sub Add(vNewValue As Sorteo)
    Dim mKey As String
 On Error GoTo Add_Error
    '
    '   #TODO: Comprobar el juego del sorteo que se añade a la colección
    '           ERR_SORTEOJUEGOERROR
    '
    '   Obtenemos la clave del parametro
    '
    mKey = CStr(vNewValue.Id)
    '
    '   Agregamos a la colección con la clave y al final de la misma
    '
    If mCol.Count <> 0 Then
        mCol.Add vNewValue, mKey, , mCol.Count
    Else
        mCol.Add vNewValue, mKey
    End If
    
Add_CleanExit:
   On Error GoTo 0
   Exit Sub

Add_Error:
   Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
   ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
   '   Audita el error
   Call HandleException(ErrNumber, ErrDescription, "Sorteos.Add")
   '   Lanza el Error
   Err.Raise ErrNumber, ErrSource, ErrDescription
End Sub


'------------------------------------------------------------------------------*
' Procedimiento  : Delete
' Fecha          : 26/may/2018
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
    '   Audita el error
    Call HandleException(ErrNumber, ErrDescription, "Sorteos.Delete")
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
    Dim mSort As Sorteo
    
   On Error GoTo MarkForDelete_Error
    If datIndex > mCol.Count Then
        strError = Replace(MSG_INDEXERROR, "#", "#" & datIndex)
        Err.Raise ERR_INDEXERROR, "Sorteos.MarkForDelete", strError
    Else
        Set mSort = mCol.Item(datIndex)
        mSort.EntidadNegocio.MarkForDelete = True
    End If
    
   On Error GoTo 0
   Exit Sub

MarkForDelete_Error:
   Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
   ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
   '   Audita el error
   Call HandleException(ErrNumber, ErrDescription, "Sorteos.MarkForDelete")
   '   Lanza el Error
   Err.Raise ErrNumber, ErrSource, ErrDescription
End Sub

'------------------------------------------------------------------------------*
' Procedimiento  : Undelete
' Fecha          : 02/04/2019
' Propósito      : Desmarcar un elemento para borrar
' Parámetros     : Indice de la colección
'------------------------------------------------------------------------------*
Public Sub Undelete(datIndex As Integer)
    Dim strError As String
    Dim mSort As Sorteo
   On Error GoTo Undelete_Error
    
    If datIndex > mCol.Count Then
        strError = Replace(MSG_INDEXERROR, "#", "#" & datIndex)
        Err.Raise ERR_INDEXERROR, "Sorteos.Undelete", strError
    Else
        Set mSort = mCol.Item(datIndex)
        mSort.EntidadNegocio.MarkForDelete = False
    End If
   
   On Error GoTo 0
   Exit Sub

Undelete_Error:
   Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
   ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
   '   Audita el error
   Call HandleException(ErrNumber, ErrDescription, "Sorteos.Undelete")
   '   Lanza el Error
   Err.Raise ErrNumber, ErrSource, ErrDescription
End Sub

'------------------------------------------------------------------------------*
' Procedimiento  : Clear
' Fecha          : 26/may/2018
' Propósito      : Iniciliza la colección
'------------------------------------------------------------------------------*
'
Public Sub Clear()
    
 On Error GoTo Clear_Error
    '
    '   Inicializamos la colección
    '
    Set mCol = Nothing
    '
    '   Volvemos a crear la colección
    '
    Set mCol = New Collection
 On Error GoTo 0
Clear__CleanExit:
    Exit Sub
            
Clear_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    '   Audita el error
    Call HandleException(ErrNumber, ErrDescription, "Sorteos.Clear", ErrSource)
    '   Lanza el Error
    Err.Raise ErrNumber, "Sorteos.Clear", ErrDescription
End Sub

'--- Propiedades --------------------------------------------------------------*
'------------------------------------------------------------------------------*
' Procedimiento  : Count
' Fecha          : 06/08/2014
' Propósito      : Devuelve el número de Sorteos de la colección
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
'------------------------------------------------------------------------------*
' Procedimiento  : Juego
' Fecha          : 02/04/2019
' Propósito      : Juego al que pertenece la colección
' Retorno        : Colección de Objetos
'------------------------------------------------------------------------------*
Public Property Get Juego() As Juego
    Juego = mJuego
End Property

Public Property Let Juego(ByVal vNewValue As Juego)
    mJuego = vNewValue
End Property
' *===========(EOF): Sorteos.cls
