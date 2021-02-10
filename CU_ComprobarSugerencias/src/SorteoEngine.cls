VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SorteoEngine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' *============================================================================*
' *
' *     Fichero    : SorteoEngine.cls
' *
' *     Autor      : Carlos Almela Baeza
' *     Creación   : mar, 09/sep/2014 23:28:15
' *     Revisión   : lu., 06/jul/2020 17:52:14
' *     Versión    : 1.1
' *     Propósito  : Motor de generación de Sorteos
' *
' *============================================================================*
Option Explicit
Option Base 0

'--- Variables Privadas -------------------------------------------------------*
Private Db              As BdDatos           ' Objeto Base de Datos
Private rgSorteos       As Range             ' Rango de varios sorteos
Private oSorteo         As Sorteo            ' Objeto Sorteo
Private oFila           As Range
Private sKey            As String

'--- Constantes ---------------------------------------------------------------*
'--- Mensajes -----------------------------------------------------------------*
'--- Errores ------------------------------------------------------------------*
'--- Propiedades --------------------------------------------------------------*
'--- Métodos Privados ---------------------------------------------------------*
Private Sub Class_Initialize()
    Set Db = New BdDatos
End Sub



'--- Métodos Públicos ---------------------------------------------------------*
'---------------------------------------------------------------------------------------
' Procedure : GetSorteoByFecha
' Author    : CHARLY
' Date      : mar, 09/sep/2014 23:33:57
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Function GetSorteoByFecha(vNewValue As Date) As Sorteo
   
   On Error GoTo GetSorteoByFecha_Error
    '
    '   Creamos el objeto Sorteo
    '
    Set oSorteo = New Sorteo
    '
    '
    '
    Set oFila = Db.GetSorteoByFecha(vNewValue)
    '
    '
    '
    If Not (oFila Is Nothing) Then
        '
        '   Creamos el sorteo
        '
        oSorteo.Constructor oFila
        '
        '   Si corresponde a la fecha solicitada
        '
        If oSorteo.Fecha = vNewValue Then
            '
            '   Devolvemos el objeto
            '
            Set GetSorteoByFecha = oSorteo
        Else
            Set GetSorteoByFecha = Nothing
        End If
    Else
        Set GetSorteoByFecha = Nothing
    End If
                
  On Error GoTo 0
    Exit Function
    
GetSorteoByFecha_Error:
   Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
   ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
   Call HandleException(ErrNumber, ErrDescription, ErrSource, "SorteoEngine.GetSorteoByFecha")
   Err.Raise ErrNumber, "SorteoEngine.GetSorteoByFecha", ErrDescription
End Function




' *===========(EOF): <<nombre fichero>>

