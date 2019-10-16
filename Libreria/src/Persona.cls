VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Persona"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' *============================================================================*
' *
' *     Fichero    : Persona.cls
' *
' *     Autor      : Carlos Almela Baeza
' *     Creación   : mi., 28/ago/2019 12:19:00
' *     Versión    : 1.0
' *     Propósito  : Clase que representa los atributos de una persona
' *
' *============================================================================*
Option Explicit
Option Base 0
'
'--- Variables Privadas -------------------------------------------------------*
Private mNombre As String
Private mApellido1 As String
Private mApellido2 As String
Private mFechaNacimiento As Date
Private mEstado As Integer
Private mSexo As Integer
'--- Constantes ---------------------------------------------------------------*
'--- Mensajes -----------------------------------------------------------------*
'--- Errores ------------------------------------------------------------------*
'--- Propiedades --------------------------------------------------------------*
Public Property Get Nombre() As String
    Nombre = mNombre
End Property

Public Property Let Nombre(ByVal vNewValue As String)
    mNombre = vNewValue
End Property

Public Property Get Apellido1() As String
    Apellido1 = mApellido1
End Property

Public Property Let Apellido1(ByVal vNewValue As String)
    mApellido1 = vNewValue
End Property

Public Property Get Apellido2() As String
    Apellido2 = mApellido2
End Property

Public Property Let Apellido2(ByVal vNewValue As String)
    mApellido2 = vNewValue
End Property

Public Property Get FechaNacimiento() As Date
    FechaNacimiento = mFechaNacimiento
End Property

Public Property Let FechaNacimiento(ByVal vNewValue As Date)
    mFechaNacimiento = vNewValue
End Property

Public Property Get Edad() As Integer
    Edad = 0
End Property

Public Property Get EstadoCivil() As Integer
    EstadoCivil = mEstado
End Property

Public Property Let EstadoCivil(ByVal vNewValue As Integer)
    mEstado = vNewValue
End Property

Public Property Get NombreInverso() As String
    NombreInverso = Apellidos & ", " & mNombre
End Property

Public Property Get Apellidos() As String
    Apellidos = mApellido1 & " " & mApellido2
End Property

Public Property Get Sexo() As Integer
    Sexo = mSexo
End Property

Public Property Let Sexo(ByVal vNewValue As Integer)
    mSexo = vNewValue
End Property

'--- Métodos Privados ---------------------------------------------------------*
Private Sub Class_Initialize()
    
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
    Err.Raise ERR_TODO, "Persona.ToString", MSG_TODO
End Function
'------------------------------------------------------------------------------*
' Procedimiento  : IsValid
' Fecha          : dd/MMM/yyyy
' Propósito      :
' Parámetros     :
' Retorno        :
'------------------------------------------------------------------------------*
Public Function IsValid() As Boolean
    Err.Raise ERR_TODO, "Persona.IsValid", MSG_TODO
End Function
'------------------------------------------------------------------------------*
' Procedimiento  : GetMensaje
' Fecha          : dd/MMM/yyyy
' Propósito      :
' Parámetros     :
' Retorno        :
'------------------------------------------------------------------------------*
Public Function GetMensaje() As String
    Err.Raise ERR_TODO, "Persona.GetMensaje", MSG_TODO
End Function

'' *===========(EOF): Persona.cls
