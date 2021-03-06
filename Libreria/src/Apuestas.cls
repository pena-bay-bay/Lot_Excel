VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Apuestas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' *============================================================================*
' *
' *     Fichero    : Apuestas.cls
' *
' *     Autor      : Carlos Almela Baeza
' *     Creaci�n   : s�., 10/ago/2019 20:34:10
' *     Versi�n    : 1.0
' *     Prop�sito  :
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
' Fecha          : s�., 10/ago/2019 20:40:44
' Prop�sito      : Devuelve el n�mero de Apuestas de la colecci�n
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
   Call HandleException(ErrNumber, ErrDescription, "Apuestas.Count")
   Err.Raise ErrNumber, ErrSource, ErrDescription
End Property
'------------------------------------------------------------------------------*
' Procedimiento  : Items
' Fecha          : s�., 10/ago/2019 20:40:57
' Prop�sito      : Devolver la colecci�n interna
' Retorno        : Colecci�n de Objetos
'------------------------------------------------------------------------------*
Public Property Get Items() As Collection
    Set Items = mCol
End Property

Public Property Let Items(ByVal vNewValue As Collection)
    Set mCol = vNewValue
End Property
'------------------------------------------------------------------------------*
' Procedimiento  : Juego
' Fecha          : s�., 10/ago/2019 20:41:15
' Prop�sito      : Juego al que pertenece la colecci�n
' Retorno        : Colecci�n de Objetos
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

'--- M�todos P�blicos ---------------------------------------------------------*
'------------------------------------------------------------------------------*
' Procedimiento  : Add
' Fecha          : s�., 10/ago/2019 20:35:47
' Prop�sito      : Agrega una apuesta a la colecci�n
' Par�metros     : Apuesta
'------------------------------------------------------------------------------*
Public Sub Add(vNewValue As Apuesta)
    Dim mKey As String
 On Error GoTo Add_Error
    '
    '   #TODO: Comprobar el juego del sorteo que se a�ade a la colecci�n
    '           ERR_SORTEOJUEGOERROR
    '
    '   Obtenemos la clave del parametro
    '
    mKey = CStr(vNewValue.Id)
    '
    '   Agregamos a la colecci�n con la clave y al final de la misma
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
   Call HandleException(ErrNumber, ErrDescription, "Apuestas.Add")
   Err.Raise ErrNumber, ErrSource, ErrDescription
End Sub
'------------------------------------------------------------------------------*
' Procedimiento  : Delete
' Fecha          : s�., 10/ago/2019 20:36:54
' Prop�sito      : Borra una apuesta de la colecci�n
'------------------------------------------------------------------------------*
'
Public Sub Delete(datIndex As Integer)
    Dim strError As String
    
  On Error GoTo Delete_Error
    If datIndex > mCol.Count Then
        strError = Replace(MSG_DELETEINDEXERROR, "#", "#" & datIndex)
        Err.Raise ERR_DELETEINDEXERROR, "Apuestas.Delete", strError
    Else
        mCol.Remove datIndex
    End If
    
  On Error GoTo 0
    Exit Sub

Delete_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    Call HandleException(ErrNumber, ErrDescription, "Apuestas.Delete")
    Err.Raise ErrNumber, ErrSource, ErrDescription
End Sub
'------------------------------------------------------------------------------*
' Procedimiento  : MarkForDelete
' Fecha          : s�., 10/ago/2019 20:37:54
' Prop�sito      : Marcar un elemento para borrar de la BBDD
' Par�metros     : Indice del elemento a borrar
' Retorno        :
'------------------------------------------------------------------------------*
Public Sub MarkForDelete(datIndex As Integer)
    Dim strError As String
    Dim mApt     As Apuesta
    
   On Error GoTo MarkForDelete_Error
    If datIndex > mCol.Count Then
        strError = Replace(MSG_INDEXERROR, "#", "#" & datIndex)
        Err.Raise ERR_INDEXERROR, "Apuestas.MarkForDelete", strError
    Else
        Set mApt = mCol.Item(datIndex)
        mApt.EntidadNegocio.MarkForDelete = True
    End If
    
   On Error GoTo 0
   Exit Sub

MarkForDelete_Error:
   Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
   ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
   Call HandleException(ErrNumber, ErrDescription, "Apuestas.MarkForDelete")
   Err.Raise ErrNumber, ErrSource, ErrDescription
End Sub

'------------------------------------------------------------------------------*
' Procedimiento  : Undelete
' Fecha          : s�., 10/ago/2019 20:38:30
' Prop�sito      : Desmarcar un elemento para borrar
' Par�metros     : Indice de la colecci�n
'------------------------------------------------------------------------------*
Public Sub Undelete(datIndex As Integer)
    Dim strError As String
    Dim mApt     As Apuesta
   On Error GoTo Undelete_Error
    
    If datIndex > mCol.Count Then
        strError = Replace(MSG_INDEXERROR, "#", "#" & datIndex)
        Err.Raise ERR_INDEXERROR, "Apuestas.Undelete", strError
    Else
        Set mApt = mCol.Item(datIndex)
        mApt.EntidadNegocio.MarkForDelete = False
    End If
   
   On Error GoTo 0
   Exit Sub

Undelete_Error:
   Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
   ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
   Call HandleException(ErrNumber, ErrDescription, "Apuestas.Undelete")
   Err.Raise ErrNumber, ErrSource, ErrDescription
End Sub

'------------------------------------------------------------------------------*
' Procedimiento  : Clear
' Fecha          : s�., 10/ago/2019 20:38:54
' Prop�sito      : Iniciliza la colecci�n
'------------------------------------------------------------------------------*
'
Public Sub Clear()
    
 On Error GoTo Clear_Error
    '
    '   Inicializamos la colecci�n
    '
    Set mCol = Nothing
    '
    '   Volvemos a crear la colecci�n
    '
    Set mCol = New Collection
 On Error GoTo 0
Clear__CleanExit:
    Exit Sub
            
Clear_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    Call HandleException(ErrNumber, ErrDescription, "Apuestas.Clear", ErrSource)
    Err.Raise ErrNumber, "Apuestas.Clear", ErrDescription
End Sub

'' *===========(EOF): Apuestas.cls
