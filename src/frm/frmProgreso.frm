VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmProgreso 
   Caption         =   "Información del Proceso"
   ClientHeight    =   1620
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4560
   OleObjectBlob   =   "frmProgreso.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "frmProgreso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' *============================================================================*
' *
' *     Fichero    : frmProgreso.cls
' *
' *     Autor      : Carlos Almela Baeza
' *     Creación   : 11/08/2009
' *     Modificado : lu., 28/dic/2020 19:22:39
' *     Versión    : 1.1
' *     Propósito  : Mostrar el progreso del proceso y el tiempo empleado
' *
' *============================================================================*
Option Explicit
'--- Variables Privadas -------------------------------------------------------*
Private m_iMaximo       As Long
Private m_iValor        As Long
Private m_dbPorcentaje  As Double
Private m_longMax       As Double
Private m_sLabel        As String
Private m_slblFase      As String
Private m_sFase         As String
Private m_slblTiempo    As String
Private m_dInicio       As Date
Private m_dFin          As Date
Private m_dDuracion     As Date
'--- Constantes ---------------------------------------------------------------*
Private Const FRM_HEIGHT_MIN As Integer = 110
Private Const FRM_HEIGHT_MAX As Integer = 140

'--- Propiedades --------------------------------------------------------------*
'------------------------------------------------------------------------------*
' Propiedad      : Maximo
' Fecha          : 11/08/2009
' Propósito      : Total de Items del proceso
'------------------------------------------------------------------------------*
'
Public Property Get Maximo() As Long
    Maximo = m_iMaximo
End Property

Public Property Let Maximo(ByVal iMaximo As Long)
    m_iMaximo = iMaximo
End Property


'------------------------------------------------------------------------------*
' Propiedad      : Valor
' Fecha          : 11/08/2009
' Propósito      : Item actual del proceso
'------------------------------------------------------------------------------*
'
Public Property Get Valor() As Long
    Valor = m_iValor
End Property

Public Property Let Valor(ByVal iValor As Long)
    m_iValor = iValor
    CalPorcentaje
    SetGraficos
End Property

'------------------------------------------------------------------------------*
' Propiedad      : Fase
' Fecha          : 11/08/2009
' Propósito      : Literal de la fase del proceso
'------------------------------------------------------------------------------*
'
Public Property Get Fase() As String
    Fase = m_sFase
End Property

Public Property Let Fase(ByVal fFase As String)
    m_sFase = fFase
    lblFase.Caption = Replace(m_slblFase, "$f", fFase)
End Property


'--- Métodos Privados ---------------------------------------------------------*
'------------------------------------------------------------------------------*
' Procedimiento  : CalPorcentaje
' Fecha          : 11/08/2009
' Propósito      : Calcular el porcentaje de progreso del proceso
'------------------------------------------------------------------------------*
'
Private Sub CalPorcentaje()
    m_dbPorcentaje = m_iValor / m_iMaximo
End Sub

'------------------------------------------------------------------------------*
' Procedimiento  : SetGraficos
' Fecha          : 11/08/2009
' Propósito      : Redibujar los componentes del formulario con sus valores
'------------------------------------------------------------------------------*
'
Private Sub SetGraficos()
    Static i As Integer
    If imgBarraGris.Visible = False Then
        imgBarraGris.Width = 0
        imgBarraGris.Visible = True
    End If
    imgBarraGris.Width = m_longMax * m_dbPorcentaje
    lblPorcentaje.Caption = Replace(m_sLabel, "$p", Format(m_dbPorcentaje, "Percent"))
    If i > 50 Then
        i = 0
        Me.Repaint
    Else
        i = i + 1
    End If
End Sub

'------------------------------------------------------------------------------*
' Procedimiento  : cmdClose_Click
' Fecha          : 11/08/2009
' Propósito      : Evento asociado a Cerrar el formulario
'------------------------------------------------------------------------------*
'
Private Sub cmdClose_Click()
    Me.Hide
End Sub

'------------------------------------------------------------------------------*
' Procedimiento  : UserForm_Initialize
' Fecha          : 11/08/2009
' Propósito      : Evento asociado a la inicialización del formulario
'------------------------------------------------------------------------------*
'
Private Sub UserForm_Initialize()
   m_longMax = imgBarraFondo.Width
   m_sLabel = lblPorcentaje.Caption
   m_slblFase = lblFase.Caption
   m_slblTiempo = lblTiempos.Caption
   m_dInicio = Now
   Me.Height = FRM_HEIGHT_MIN
   lblTiempos.Visible = False
   cmdClose.Visible = False
End Sub


'--- Métodos Públicos ---------------------------------------------------------*
'------------------------------------------------------------------------------*
' Propiedad      : DisProceso
' Fecha          : lu., 28/dic/2020 19:31:04
' Propósito      : Visualiza el tiempo empleado por el proceso
'------------------------------------------------------------------------------*
'
Public Sub DisProceso()
    Dim mFmt As String
    '
    '   Calculamos tiempo
    '
    m_dFin = Now
    m_dDuracion = m_dFin - m_dInicio
    '
    '   formateamos salida y editamos etiqueta
    '
    mFmt = Format(m_dDuracion, "ttttt")
    lblTiempos.Caption = Replace(m_slblTiempo, "$t", mFmt)
    '
    '   Formateamos formulario
    Me.Height = FRM_HEIGHT_MAX
    '   visualizamos etiquetas de tiempos
    '   mostramos botón
    lblTiempos.Visible = True
    cmdClose.Visible = True
End Sub
' *===========(EOF): frmProgreso.cls
