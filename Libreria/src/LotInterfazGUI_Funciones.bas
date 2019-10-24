Attribute VB_Name = "LotInterfazGUI_Funciones"
' *============================================================================*
' *
' *     Fichero    : LotInterfazGUI_Funciones.bas
' *
' *     Autor      : Carlos Almela Baeza
' *     Creación   : mi., 06/mar/2019 00:11:02
' *     Versión    : 1.0
' *     Propósito  : Funciones comunes de la interfaz de usuario de la
' *                  aplicación
' *============================================================================*
Option Explicit
Option Base 0

'--- Constantes ---------------------------------------------------------------*
Public Enum Modo
    mdSinDefinir = 0
    mdAlta = 1
    mdEdicion = 2
    mdConsulta = 3
End Enum

Public Const LT_ALTA As String = "Alta"
Public Const LT_EDICION As String = "Edicion"
Public Const LT_CONSULTA As String = "Consulta"

'--- Metodos Publicos ---------------------------------------------------------*
Sub Salir()
    '
    '   TODO: Preguntar si queremos salir
    '
    ActiveWorkbook.Close True
End Sub

Sub Go2Inicio()
    '
    '   TODO: verificar si la información esta actualizada, sino actualizar
    '
    ActiveWorkbook.Sheets("Portada").Activate
End Sub

Sub Go2Sorteos()
    ActiveWorkbook.Sheets("Sorteos").Activate
End Sub

Sub Go2Apuestas()
    ActiveWorkbook.Sheets("Apuestas").Activate
End Sub

Sub Go2Sugerencias()
    ActiveWorkbook.Sheets("Sugerencias").Activate
End Sub

Sub Go2Contabilidad()
    ActiveWorkbook.Sheets("Contabilidad").Activate
End Sub

Sub Go2Estadistica()
    ActiveWorkbook.Sheets("Estadistica").Activate
End Sub

'--- Metodos Publicos ---------------------------------------------------------*
'--- Propiedades --------------------------------------------------------------*
'------------------------------------------------------------------------------*
' Procedimiento  : <<nombre propiedad>>
' Fecha          : dd/MMM/yyyy
' Propósito      :
' Parámetros     :
' Retorno        :
'------------------------------------------------------------------------------*
'' *===========(EOF): LotInterfazGUI_Funciones.bas
