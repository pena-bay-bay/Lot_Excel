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
' *     Versión    : 1.0
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
'--- Métodos Privados ---------------------------------------------------------*
Private Sub Class_Initialize()
    Set mCol = New Collection
End Sub

Private Sub Class_Terminate()
    Set mCol = Nothing
End Sub
'--- Métodos Públicos ---------------------------------------------------------*
Public Function ToString() As String

End Function
Public Sub Parse(datFiltros As String)

End Sub
'------------------------------------------------------------------------------*
' Procedimiento  : <<nombre propiedad>>
' Fecha          : dd/MMM/yyyy
' Propósito      :
' Parámetros     :
'------------------------------------------------------------------------------*
'' *===========(EOF): FiltrosCombinacion.cls
