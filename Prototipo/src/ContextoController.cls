VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ContextoController"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' *============================================================================*
' *
' *     Fichero    : ContextoController.cls
' *
' *     Autor      : Carlos Almela Baeza
' *     Creaci�n   : mi., 06/mar/2019 00:01:00
' *     Versi�n    : 1.0
' *     Prop�sito  : Controlador del CU Definir Contexto
' *                  Se encarga de mantener actualizada la informaci�n del
' *                  contexto de la aplicaci�n y la informaci�n relevante
' *============================================================================*
Option Explicit
Option Base 0
'
'--- Variables Privadas -------------------------------------------------------*
Private rgAlmacen          As RAnge

'--- Constantes ---------------------------------------------------------------*
Private Const RangoAlmacen As String = "Contexto!A1"
'--- Mensajes -----------------------------------------------------------------*
'--- Errores ------------------------------------------------------------------*
'--- Propiedades --------------------------------------------------------------*
Public JuegoText            As String           'Texto del juego seleccionado
Public JuegoSeleccionado    As Juego            'Juego Seleccionado
Public PeriodoText          As String           'Texto del periodo seleccionado
Public PeriodoSeleccionado  As Periodo          'Periodo seleccionado
Public EstadisticasText     As String           'Texto de las Estadisticas
                                                'Par�metros de estadisticas
Public UsuarioText          As String           'Nombre del usuario

'--- Metodos Privados ---------------------------------------------------------*
Private Sub Class_Initialize()
    '
    '   Creamos el rango de almacenaje de datos
    '
    Set rgAlmacen = ThisWorkbook.Sheets("Contexto").RAnge(RangoAlmacen)
    '
    '   Creamos el periodo de selecci�n de datos
    '
'    Set PeriodoSeleccionado = New Periodo
    '
    '   TODO: Comprobar que existe la hoja Contexto
    '   TODO: Si la hoja contexto no tiene datos inicializar con datos por defecto
    '   TODO: Definir parametros de estadistica (pendiente de clase).
End Sub

Private Sub Class_Terminate()
    Set rgAlmacen = Nothing
    Set PeriodoSeleccionado = Nothing
End Sub
'--- Metodos Publicos ---------------------------------------------------------*
Public Sub Consultar()
    '
    '   Cargamos el dato de la celda
    '
    JuegoText = rgAlmacen.Offset(2, 1).Value
    JuegoSeleccionado = rgAlmacen.Offset(2, 1).Value
'    PeriodoText =
    
End Sub
Public Sub Guardar()

End Sub
'------------------------------------------------------------------------------*
' Procedimiento  : <<nombre propiedad>>
' Fecha          : dd/MMM/yyyy
' Prop�sito      :
' Par�metros     :
' Retorno        :
'------------------------------------------------------------------------------*
'' *===========(EOF): ContextoController.cls
