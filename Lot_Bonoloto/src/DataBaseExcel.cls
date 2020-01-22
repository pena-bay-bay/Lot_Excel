VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DataBaseExcel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' *============================================================================*
' *
' *     Fichero    : DataBaseExcel.cls
' *
' *     Autor      : Carlos Almela Baeza
' *     Creación   : Mie, 02/may/2018 16:54:26
' *     Versión    : 1.0
' *     Propósito  : Gestionar los libros de datos como una base de datos
' *                  Trabajando con rangos
' *                  Se encarga de abrir los libros de los distintos juegos
' *                  y suministrar datos definiendo los rangos de los datos
' *============================================================================*
Option Explicit
Option Base 0
'--- Constantes ---------------------------------------------------------------*
Private Const NOMBRE_LIBROSDATOS = "Lot_Bonoloto_v2018(beta).xlsm;Lot_Primitiva_v2018(beta).xlsm;" & _
                                   "Lot_Euromillones.xlsm;Lot_GordoPrimitiva.xlsm"
Private Const NOMBRE_HOJASDATOS = "Resultados;Resultados;Resultados;Resultados;"
'--- Variables Privadas -------------------------------------------------------*
Private mWorkbookDatos  As Workbook
Private mRutaLibro      As String
Private mListaLibros    As Variant
Private mRango          As Range
Private mHoja           As Worksheet
'--- Mensajes -----------------------------------------------------------------*
'--- Errores ------------------------------------------------------------------*
'--- Metodos Privados ---------------------------------------------------------*
Private Sub Class_Initialize()
    '   Componer la matriz de libros y la matriz de resultados
    '
    '
End Sub

Private Sub Class_Terminate()
    '
    '   Borrar objetos
    '
End Sub
'--- Metodos Publicos ---------------------------------------------------------*
Public Sub Abrir(vNewValue As Juego)
    '
    '   Comprobar que el libro esta abierto, sino es asi abrirlo
    '       Componer el nombre del libro donde están los resultados
    '       Abrir el fichero establecer el libro de datos al abierto
    '
End Sub

'--- Propiedades --------------------------------------------------------------*

Public Function GetResultados(vNewValue As Juego, vNewPeriodo As Periodo) As Range

End Function

Public Function GetResultadoById(vNewValue As Juego, vNewId As Integer) As Range

End Function





' *===========(EOF): <<nombre fichero>>


