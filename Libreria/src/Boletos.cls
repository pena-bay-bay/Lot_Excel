VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Boletos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' *============================================================================*
' *
' *     Fichero    : Boletos.cls
' *
' *     Autor      : Carlos Almela Baeza
' *     Creación   : sá., 10/ago/2019 20:34:10
' *     Versión    : 1.0
' *     Propósito  :
' *
' *============================================================================*
Option Explicit
Option Base 0
'
'--- Variables Privadas -------------------------------------------------------*
Private mCol As Collection                ' Usado para almacenar la Clase
Private mJuego As Juego                   ' Juego al que pertenece la coleccion

'--- Constantes ---------------------------------------------------------------*
'--- Mensajes -----------------------------------------------------------------*
'--- Errores ------------------------------------------------------------------*
'--- Propiedades --------------------------------------------------------------*
'------------------------------------------------------------------------------*
' Procedimiento  : Count
' Fecha          : sá., 10/ago/2019 20:40:44
' Propósito      : Devuelve el número de Boletos de la colección
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
   Call HandleException(ErrNumber, ErrDescription, "Boletos.Count")
   Err.Raise ErrNumber, ErrSource, ErrDescription
End Property
'------------------------------------------------------------------------------*
' Procedimiento  : Items
' Fecha          : sá., 10/ago/2019 20:40:57
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
' Fecha          : sá., 10/ago/2019 20:41:15
' Propósito      : Juego al que pertenece la colección
' Retorno        : Colección de Objetos
'------------------------------------------------------------------------------*
Public Property Get Juego() As Juego
    Juego = mJuego
End Property

Public Property Let Juego(ByVal vNewValue As Juego)
    mJuego = vNewValue
End Property
'--- Metodos Privados ---------------------------------------------------------*
Private Sub Class_Initialize()
    Set mCol = New Collection
End Sub

Private Sub Class_Terminate()
    Set mCol = Nothing
End Sub

'--- Métodos Públicos ---------------------------------------------------------*
'------------------------------------------------------------------------------*
' Procedimiento  : Add
' Fecha          : sá., 10/ago/2019 20:35:47
' Propósito      : Agrega una apuesta a la colección
' Parámetros     : Boleto
'------------------------------------------------------------------------------*
Public Sub Add(vNewValue As Boleto)
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
   Call HandleException(ErrNumber, ErrDescription, "Boletos.Add")
   Err.Raise ErrNumber, ErrSource, ErrDescription
End Sub
'------------------------------------------------------------------------------*
' Procedimiento  : Delete
' Fecha          : sá., 10/ago/2019 20:36:54
' Propósito      : Borra una Boleto de la colección
'------------------------------------------------------------------------------*
'
Public Sub Delete(datIndex As Integer)
    Dim strError As String
    
  On Error GoTo Delete_Error
    If datIndex > mCol.Count Then
        strError = Replace(MSG_DELETEINDEXERROR, "#", "#" & datIndex)
        Err.Raise ERR_DELETEINDEXERROR, "Boletos.Delete", strError
    Else
        mCol.Remove datIndex
    End If
    
  On Error GoTo 0
    Exit Sub

Delete_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    Call HandleException(ErrNumber, ErrDescription, "Boletos.Delete")
    Err.Raise ErrNumber, ErrSource, ErrDescription
End Sub
'------------------------------------------------------------------------------*
' Procedimiento  : MarkForDelete
' Fecha          : sá., 10/ago/2019 20:37:54
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
        Err.Raise ERR_INDEXERROR, "Boletos.MarkForDelete", strError
    Else
        Set mSort = mCol.Item(datIndex)
        mSort.EntidadNegocio.MarkForDelete = True
    End If
    
   On Error GoTo 0
   Exit Sub

MarkForDelete_Error:
   Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
   ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
   Call HandleException(ErrNumber, ErrDescription, "Boletos.MarkForDelete")
   Err.Raise ErrNumber, ErrSource, ErrDescription
End Sub

'------------------------------------------------------------------------------*
' Procedimiento  : Undelete
' Fecha          : sá., 10/ago/2019 20:38:30
' Propósito      : Desmarcar un elemento para borrar
' Parámetros     : Indice de la colección
'------------------------------------------------------------------------------*
Public Sub Undelete(datIndex As Integer)
    Dim strError As String
    Dim mSort As Sorteo
   On Error GoTo Undelete_Error
    
    If datIndex > mCol.Count Then
        strError = Replace(MSG_INDEXERROR, "#", "#" & datIndex)
        Err.Raise ERR_INDEXERROR, "Boletos.Undelete", strError
    Else
        Set mSort = mCol.Item(datIndex)
        mSort.EntidadNegocio.MarkForDelete = False
    End If
   
   On Error GoTo 0
   Exit Sub

Undelete_Error:
   Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
   ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
   Call HandleException(ErrNumber, ErrDescription, "Boletos.Undelete")
   Err.Raise ErrNumber, ErrSource, ErrDescription
End Sub

'------------------------------------------------------------------------------*
' Procedimiento  : Clear
' Fecha          : sá., 10/ago/2019 20:38:54
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
    Call HandleException(ErrNumber, ErrDescription, "Boletos.Clear", ErrSource)
    Err.Raise ErrNumber, "Boletos.Clear", ErrDescription
End Sub

'' *===========(EOF): Apuestas.cls

