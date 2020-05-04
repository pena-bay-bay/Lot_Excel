VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sugerencia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' *============================================================================*
' *
' *     Fichero    : Sugerencia.cls
' *
' *     Autor      : Carlos Almela Baeza
' *     Creación   : vie, 13/abr/2012 hh:mm:ss
' *     Versión    : 1.0
' *     Propósito  :
' *
' *============================================================================*
Option Explicit
Option Base 0
'
'--- Variables Privadas -------------------------------------------------------*
Private m_colBolas As Bolas
Private m_eModalidad As ModalidadJuego
Private m_objCombinacion As Combinacion
'--- Constantes ---------------------------------------------------------------*
'--- Mensajes -----------------------------------------------------------------*
'--- Errores ------------------------------------------------------------------*
'--- Propiedades --------------------------------------------------------------*
Public Entidad As EntidadNegocio
Public Metodo As Metodo
Public Parametros As ParametrosMuestra

Public Property Get Bolas() As Bolas
    Set Bolas = m_colBolas
End Property

Public Property Set Bolas(colBolas As Bolas)
    Set m_colBolas = colBolas
End Property

Public Property Get Modalidad() As ModalidadJuego
    Modalidad = m_eModalidad
End Property

Public Property Let Modalidad(ByVal eModalidad As ModalidadJuego)
    m_eModalidad = eModalidad
End Property

Public Property Get Combinacion() As Combinacion
    Set Combinacion = m_objCombinacion
End Property

Public Property Set Combinacion(objCombinacion As Combinacion)
    Set m_objCombinacion = objCombinacion
End Property


'--- Métodos Privados ---------------------------------------------------------*
Private Sub Class_Initialize()
    Set m_colBolas = New Bolas
    Set m_objCombinacion = New Combinacion
    m_eModalidad = LP_LB_6_49
End Sub

Private Sub Class_Terminate()
    Set m_colBolas = Nothing
    Set m_objCombinacion = Nothing
End Sub
'--- Métodos Públicos ---------------------------------------------------------*
'------------------------------------------------------------------------------*
' Procedimiento  : <<nombre propiedad>>
' Fecha          : dd/MMM/yyyy
' Propósito      :
' Parámetros     :
'------------------------------------------------------------------------------*
Public Function Validar() As Boolean
    Dim bResult As Boolean
    Validar = bResult
End Function

Public Function MensajeError() As String
    Dim sResult As String
    MensajeError = sResult
End Function

Public Function ToString() As String
    Dim sResult As String
    ToString = sResult
End Function

'' *===========(EOF): Sugerencia.cls






