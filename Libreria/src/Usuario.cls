VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Usuario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' *============================================================================*
' *
' *     Fichero    : Usuario.cls
' *
' *     Autor      : Carlos Almela Baeza
' *     Creación   : mi., 28/ago/2019 11:46:04
' *     Versión    : 1.0
' *     Propósito  : Clase que representa al usuario del sistema
' *
' *============================================================================*
Option Explicit
Option Base 0
'
'--- Variables Privadas -------------------------------------------------------*

Private mIdUsuario As String
Private mEmail As String
Private mRol  As Integer

'--- Constantes ---------------------------------------------------------------*
'--- Mensajes -----------------------------------------------------------------*
'--- Errores ------------------------------------------------------------------*
'--- Propiedades --------------------------------------------------------------*
Public Persona As Persona

Public Property Get IdUsuario() As String
     IdUsuario = mIdUsuario
End Property

Public Property Let IdUsuario(ByVal vNewValue As String)
    mIdUsuario = vNewValue
End Property

Public Property Get Nombre() As String
    Nombre = Persona.Nombre
End Property

Public Property Let Nombre(ByVal vNewValue As String)
    Persona.Nombre = vNewValue
End Property

Public Property Get Apellidos() As String
    Apellidos = Persona.Apellidos
End Property

Public Property Get Email() As String
    Email = mEmail
End Property

Public Property Let Email(ByVal vNewValue As String)
    mEmail = vNewValue
End Property

Public Property Get Rol() As Integer
    Rol = mRol
End Property

Public Property Let Rol(ByVal vNewValue As Integer)
    mRol = vNewValue
End Property

'--- Métodos Privados ---------------------------------------------------------*
Private Sub Class_Initialize()
    Set Persona = New Persona
End Sub

Private Sub Class_Terminate()
    Set Persona = Nothing
End Sub
'--- Métodos Públicos ---------------------------------------------------------*
'------------------------------------------------------------------------------*
' Procedimiento  : ToString
' Fecha          : dd/MMM/yyyy
' Propósito      :
' Parámetros     :
' Retorno        :
'------------------------------------------------------------------------------*
Public Function ToString() As String
    Err.Raise ERR_TODO, "Usuario.ToString", MSG_TODO
End Function
'------------------------------------------------------------------------------*
' Procedimiento  : IsValid
' Fecha          : dd/MMM/yyyy
' Propósito      :
' Parámetros     :
' Retorno        :
'------------------------------------------------------------------------------*
Public Function IsValid() As Boolean
    Err.Raise ERR_TODO, "Usuario.IsValid", MSG_TODO
End Function
'------------------------------------------------------------------------------*
' Procedimiento  : GetMensaje
' Fecha          : dd/MMM/yyyy
' Propósito      :
' Parámetros     :
' Retorno        :
'------------------------------------------------------------------------------*
Public Function GetMensaje() As String
    Err.Raise ERR_TODO, "Usuario.GetMensaje", MSG_TODO
End Function


'' *===========(EOF): Usuario.cls
