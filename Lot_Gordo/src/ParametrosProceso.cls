VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ParametrosProceso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' *============================================================================*
' *
' *     Fichero    : ParametrosProceso.cls
' *
' *     Autor      : Carlos Almela Baeza
' *     Creación   : Dom, 06/may/2018 23:57:06
' *     Versión    : 1.0
' *     Propósito  : Colección de variables ParamProceso
' *
' *
' *============================================================================*
Option Explicit
Option Base 0

'--- Variables Privadas -------------------------------------------------------*
Dim mCol        As Collection

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
' Fecha          : 26/may/2018
' Propósito      : Agrega un parametro a la colección
'------------------------------------------------------------------------------*
'
Public Sub Add(vNewValue As ParamProceso)
    Dim mKey As String
    
 On Error GoTo Add_Error
    '
    '   Obtenemos la clave del parametro
    '
    mKey = vNewValue.Nombre
    '
    '   Agregamos a la colección con la clave y al final de la misma
    '
    If mCol.Count <> 0 Then
        mCol.Add vNewValue, mKey, , mCol.Count
    Else
        mCol.Add vNewValue, mKey
    End If
    
 On Error GoTo 0
Add__CleanExit:
    Exit Sub
            
Add_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    '   Audita el error
    Call HandleException(ErrNumber, ErrDescription, "ParametrosProceso.Add", ErrSource)
    '   Lanza el Error
    Err.Raise ErrNumber, "ParametrosProceso.Add", ErrDescription
End Sub
'------------------------------------------------------------------------------*
' Procedimiento  : Delete
' Fecha          : 26/may/2018
' Propósito      : Agrega un parametro a la colección
'------------------------------------------------------------------------------*
'
Public Sub Delete(vNewValue As ParamProceso)
    Dim mIndex  As Integer
    Dim mTmp    As ParamProceso
    Dim mKey    As String
    
 On Error GoTo Delete_Error
    '
    '   Si no hay variable nada que borrar
    '
    If mCol.Count = 0 Then
        Exit Sub
    End If
    '
    '   Inicializamos el indice
    '
    mIndex = 0
    '
    '   Obtenemos la clave
    '
    mKey = vNewValue.Nombre
    '
    '   Buscamos el indice de la variable que queremos borrar
    '
    For Each mTmp In mCol
        '
        '   Si es el elemento salimos del bucle
        '
        If (mTmp.Nombre = mKey) Then
            Exit For
        End If
        '
        '   Incrementamos el indice
        '
        mIndex = mIndex + 1
    Next mTmp
    '
    '   Borramos la variable
    '
    mCol.Remove mIndex
    '
    '
    '
 On Error GoTo 0
Delete__CleanExit:
    Exit Sub
            
Delete_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    '   Audita el error
    Call HandleException(ErrNumber, ErrDescription, "ParametrosProceso.Delete", ErrSource)
    '   Lanza el Error
    Err.Raise ErrNumber, "ParametrosProceso.Delete", ErrDescription
End Sub
'------------------------------------------------------------------------------*
' Procedimiento  : Clear
' Fecha          : 26/may/2018
' Propósito      : Agrega un parametro a la colección
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
    Call HandleException(ErrNumber, ErrDescription, "ParametrosProceso.Clear", ErrSource)
    '   Lanza el Error
    Err.Raise ErrNumber, "ParametrosProceso.Clear", ErrDescription
End Sub
'------------------------------------------------------------------------------*
' Procedimiento  : GetVariable
' Fecha          : 26/may/2018
' Propósito      : Agrega un parametro a la colección
'------------------------------------------------------------------------------*
'
Public Function GetVariable(vNewValue As String) As ParamProceso
    Dim mTmp As ParamProceso
    
 On Error GoTo GetVariable_Error
    '
    '
    '
    Set mTmp = New ParamProceso
    '
    '
    '
    Set mTmp = mCol.Item(vNewValue)
    '
    '
    '
    Set GetVariable = mTmp
    
 On Error GoTo 0
GetVariable__CleanExit:
    Exit Function
            
GetVariable_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    '   Audita el error
    Call HandleException(ErrNumber, ErrDescription, "ParametrosProceso.GetVariable", ErrSource)
    '   Lanza el Error
    Err.Raise ErrNumber, "ParametrosProceso.GetVariable", ErrDescription
End Function


'--- Propiedades --------------------------------------------------------------*
Public Property Get Items() As Collection
    Set Items = mCol
End Property

Public Property Get Count() As Integer
    Count = mCol.Count
End Property

' *===========(EOF): ParametrosProceso.cls



