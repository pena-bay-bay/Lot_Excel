VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sugerencias"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' *============================================================================*
' *
' *     Fichero    : Sugerencias.cls
' *
' *     Autor      : Carlos Almela Baeza
' *     Creaci�n   : ju., 23/abr/2020 19:01:11
' *     Versi�n    : 1.0
' *     Prop�sito  :
' *
' *============================================================================*
Option Explicit
Option Base 0

'--- Variables Privadas -------------------------------------------------------*
Private mCol As Collection                ' Usado para almacenar la Clase

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
' Date      :
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Sub Add(vNewValue As Sugerencia)
    Dim mKey As String
 On Error GoTo Add_Error
    '
    '   Obtenemos la clave del parametro
    '
    mKey = CStr(vNewValue.Entidad.Id)
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
   Call HandleException(ErrNumber, ErrDescription, "Sugerencias.Add")
   Err.Raise ErrNumber, ErrSource, ErrDescription
End Sub


'------------------------------------------------------------------------------*
' Procedimiento  : Delete
' Fecha          :
' Prop�sito      : Borra una sugerencia de la colecci�n
'------------------------------------------------------------------------------*
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
    Call HandleException(ErrNumber, ErrDescription, "Sugerencias.Delete")
    Err.Raise ErrNumber, ErrSource, ErrDescription
End Sub

'------------------------------------------------------------------------------*
' Procedimiento  : MarkForDelete
' Fecha          :
' Prop�sito      : Marcar un elemento para borrar de la BBDD
' Par�metros     : Indice del elemento a borrar
' Retorno        :
'------------------------------------------------------------------------------*
Public Sub MarkForDelete(datIndex As Integer)
    Dim strError As String
    Dim mSuge As Sugerencia
    
   On Error GoTo MarkForDelete_Error
    If datIndex > mCol.Count Then
        strError = Replace(MSG_INDEXERROR, "#", "#" & datIndex)
        Err.Raise ERR_INDEXERROR, "Sorteos.MarkForDelete", strError
    Else
        Set mSuge = mCol.Item(datIndex)
        mSuge.Entidad.MarkForDelete = True
    End If
    
   On Error GoTo 0
   Exit Sub

MarkForDelete_Error:
   Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
   ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
   Call HandleException(ErrNumber, ErrDescription, "Sugerencias.MarkForDelete")
   Err.Raise ErrNumber, ErrSource, ErrDescription
End Sub



'------------------------------------------------------------------------------*
' Procedimiento  : Undelete
' Fecha          :
' Prop�sito      : Desmarcar un elemento para borrar
' Par�metros     : Indice de la colecci�n
'------------------------------------------------------------------------------*
Public Sub Undelete(datIndex As Integer)
    Dim strError As String
    Dim mSuge As Sugerencia
   On Error GoTo Undelete_Error
    
    If datIndex > mCol.Count Then
        strError = Replace(MSG_INDEXERROR, "#", "#" & datIndex)
        Err.Raise ERR_INDEXERROR, "Sorteos.Undelete", strError
    Else
        Set mSuge = mCol.Item(datIndex)
        mSuge.Entidad.MarkForDelete = False
    End If
   
   On Error GoTo 0
   Exit Sub

Undelete_Error:
   Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
   ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
   Call HandleException(ErrNumber, ErrDescription, "Sugerencias.Undelete")
   Err.Raise ErrNumber, ErrSource, ErrDescription
End Sub



'------------------------------------------------------------------------------*
' Procedimiento  : Clear
' Fecha          : 26/may/2018
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
    Call HandleException(ErrNumber, ErrDescription, "Sugerencias.Clear", ErrSource)
    Err.Raise ErrNumber, "Sorteos.Clear", ErrDescription
End Sub

'--- Propiedades --------------------------------------------------------------*
'------------------------------------------------------------------------------*
' Procedimiento  : Count
' Fecha          : 06/08/2014
' Prop�sito      : Devuelve el n�mero de Sorteos de la colecci�n
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
   Call HandleException(ErrNumber, ErrDescription, "Sugerencias.Count")
   Err.Raise ErrNumber, ErrSource, ErrDescription
End Property


'------------------------------------------------------------------------------*
' Procedimiento  : Items
' Fecha          : 02/04/2019
' Prop�sito      : Devolver la colecci�n interna
' Retorno        : Colecci�n de Objetos
'------------------------------------------------------------------------------*
Public Property Get Items() As Collection
    Set Items = mCol
End Property

Public Property Let Items(ByVal vNewValue As Collection)
    Set mCol = vNewValue
End Property

'' *===========(EOF): Sugerencias.cls
