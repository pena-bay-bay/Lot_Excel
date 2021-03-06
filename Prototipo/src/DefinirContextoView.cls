VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DefinirContextoView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' *============================================================================*
' *
' *     Fichero    : DefinirContextoView.cls
' *
' *     Autor      : Carlos Almela Baeza
' *     Creaci�n   : do., 07/abr/2019 19:09:14
' *     Versi�n    : 1.0
' *     Prop�sito  : Vista de definici�n de variables de contexto:
' *                  Juego Seleccionado
' *                  Periodo de tiempo aplicable
' *                  Parametros estadisticos
' *============================================================================*
Option Explicit
Option Base 0
'
'--- Variables Privadas -------------------------------------------------------*
Private mController  As ContextoController
'--- Propiedades --------------------------------------------------------------*
Public JuegoText            As String           'Texto del juego seleccionado
Public JuegoSeleccionado    As Juego            'Juego Seleccionado
Public PeriodoText          As String           'Texto del periodo seleccionado
Public PeriodoSeleccionado  As Periodo          'Periodo seleccionado
Public EstadisticasText     As String           'Texto de las Estadisticas
                                                'Par�metros de estadisticas
Public UsuarioText          As String           'Nombre del usuario
            

'--- Constantes ---------------------------------------------------------------*
'--- Mensajes -----------------------------------------------------------------*
'--- Errores ------------------------------------------------------------------*
'--- M�todos Privados ---------------------------------------------------------*
Private Sub Class_Initialize()
    Set mController = New ContextoController
End Sub

Private Sub Class_Terminate()
    Set mController = Nothing
End Sub
'--- M�todos P�blicos ---------------------------------------------------------*
'------------------------------------------------------------------------------*
' Procedimiento  : Refrescar
' Fecha          : 07/abr/2019
' Prop�sito      : Actualizar los datos de la vista
'------------------------------------------------------------------------------*
Public Sub Refrescar()

End Sub
'------------------------------------------------------------------------------*
' Procedimiento  : Refrescar
' Fecha          : 07/abr/2019
' Prop�sito      : Actualizar los datos de la vista
'------------------------------------------------------------------------------*
Public Sub EditarContexto()

End Sub
'' *===========(EOF): DefinirContextoView.cls
