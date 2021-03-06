VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "HojaSugerenciasUI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
' *============================================================================*
' *
' *     Fichero    : HojaSugerenciasUI.cls
' *
' *     Autor      : Carlos Almela Baeza
' *     Creaci�n   : ma., 05/mar/2019 23:50:06
' *     Versi�n    : 1.0
' *     Prop�sito  :
' *
' *============================================================================*
Option Explicit
Option Base 0
'
'--- Variables Privadas -------------------------------------------------------*
'--- Constantes ---------------------------------------------------------------*
Private Const RgContexto As String = "M1:N3"
'--- Mensajes -----------------------------------------------------------------*
'--- Errores ------------------------------------------------------------------*
'--- Metodos Privados ---------------------------------------------------------*
'------------------------------------------------------------------------------*
' Procedimiento  : Worksheet_BeforeDoubleClick
' Fecha          : 03/mar/2019
' Prop�sito      : Intercepta el doble click del raton para interactuar con
'                  areas sensibles
' Par�metros     : Target Rango donde se ha pulsado el doble click
'                  Cancel indicador de cancelaci�n del docle click
'------------------------------------------------------------------------------*
Private Sub Worksheet_BeforeDoubleClick(ByVal Target As RAnge, Cancel As Boolean)
    Debug.Print Target.Address
End Sub
'--- Metodos Publicos ---------------------------------------------------------*
'--- Propiedades --------------------------------------------------------------*
'' *===========(EOF): HojaSugerenciasUI.cls
