VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EntidadNegocio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' *============================================================================*
' *
' *     Fichero    : EntidadNegocio.cls
' *
' *     Autor      : Carlos Almela Baeza
' *     Creación   : vi., 21/sep/2018 18:07:57
' *     Versión    : 1.0
' *     Propósito  : Define las caracteristicas de instanciación de una
'                    clase que es de Negocio
' *
' *============================================================================*
Option Explicit
Option Base 0
'
'--- Variables Publicas -------------------------------------------------------*
Public Id                   As Long
Public FechaAlta            As Date
Public FechaModificacion    As Date
Public FechaBaja            As Date
Public ClassStorage         As Boolean
Public IsDirty              As Boolean
Public IsNew                As Boolean
Public Situacion            As Integer
Public MarkForDelete        As Boolean
Public Origen               As Integer

'--- Constantes ---------------------------------------------------------------*
'--- Mensajes -----------------------------------------------------------------*
'--- Errores ------------------------------------------------------------------*
'--- Metodos Privados ---------------------------------------------------------*
Private Sub Class_Initialize()
    Id = 1
    FechaAlta = Now
    FechaModificacion = Now
    ClassStorage = False
    IsDirty = False
    IsNew = True
    MarkForDelete = False
    Situacion = 0
    Origen = 0
End Sub
'--- Metodos Publicos ---------------------------------------------------------*
'--- Propiedades --------------------------------------------------------------*
'' *===========(EOF): EntidadNegocio.cls
