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
' *     Propósito  : Contiene una sugerencia para un sorteo y sus parametros
' *                  de generación
' *============================================================================*
Option Explicit
Option Base 0
'
'--- Variables Privadas -------------------------------------------------------*
Private mModalidad As ModalidadJuego
Private mCombinacion As Combinacion
'--- Constantes ---------------------------------------------------------------*
'--- Mensajes -----------------------------------------------------------------*
'--- Errores ------------------------------------------------------------------*
'--- Propiedades --------------------------------------------------------------*
Public Entidad As EntidadNegocio
Public Metodo As Metodo
Public Parametros As ParametrosMuestra

Public Property Get Modalidad() As ModalidadJuego
    Modalidad = mModalidad
End Property

Public Property Let Modalidad(ByVal eModalidad As ModalidadJuego)
    mModalidad = eModalidad
End Property

Public Property Get Combinacion() As Combinacion
    Set Combinacion = mCombinacion
End Property

Public Property Set Combinacion(objCombinacion As Combinacion)
    Set mCombinacion = objCombinacion
End Property


'--- Métodos Privados ---------------------------------------------------------*
Private Sub Class_Initialize()
    Set mCombinacion = New Combinacion
    Set Entidad = New EntidadNegocio
    Set Metodo = New Metodo
    Set Parametros = New ParametrosMuestra
    mModalidad = LP_LB_6_49
End Sub

Private Sub Class_Terminate()
    Set mCombinacion = Nothing
    Set Entidad = Nothing
    Set Metodo = Nothing
    Set Parametros = Nothing
End Sub
'--- Métodos Públicos ---------------------------------------------------------*
'------------------------------------------------------------------------------*
' Procedimiento  : IsValid
' Fecha          :
' Propósito      : validar los atributos de la clase
'------------------------------------------------------------------------------*
Public Function IsValid() As Boolean
    Dim bResult As Boolean
    If Not (Metodo.IsValid) Then
        bResult = False
    Else
        bResult = True
    End If
    If Not (Parametros.Validar) Then
        bResult = False
    Else
        bResult = True
    End If
    IsValid = bResult
End Function

Public Function MensajeError() As String
    Dim sResult As String
    MensajeError = sResult
End Function

Public Function ToString() As String
    Dim sResult As String
    Dim mModJuego As String
    Select Case mModalidad
        Case LP_LB_6_49: mModJuego = "6/49"
        Case GP_5_54: mModJuego = "5/54"
        Case EU_5_50: mModJuego = "5/50"
        Case EU_2_12: mModJuego = "2/12"
    End Select
        
    ' Sugerencia (Modalidad-fecha): nn-nn-nn-nn-nn-nn-nn-nn, Metodo
    sResult = "Sugerencia: ($1-$2) $3 , $4"
    sResult = Replace(sResult, "$1", mModJuego)
    sResult = Replace(sResult, "$2", Format(Parametros.FechaAnalisis, "dd/mm/yyyy"))
    sResult = Replace(sResult, "$3", mCombinacion.TextoOrdenado)
    sResult = Replace(sResult, "$4", Metodo.ToString)
    ToString = sResult
End Function

'' *===========(EOF): Sugerencia.cls






