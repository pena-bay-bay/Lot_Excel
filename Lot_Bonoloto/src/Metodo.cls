VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Metodo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Metodo de Sugerencias de Apuestas"
' *============================================================================*
' *
' *     Fichero    : Metodo.cls
' *
' *     Autor      : Carlos Almela Baeza
' *     Creación   : ma., 28/ene/2020 20:40:03
' *     Versión    : 1.0
' *     Propósito  : Clase que contiene los parámetros de una sugerencia
' *
' *============================================================================*
Option Explicit
Option Base 0
'
'--- Variables Privadas -------------------------------------------------------*
Private mSentidoOrdenacion    As Boolean
Private mOrden                As Integer
Private mCriteriosOrdenacion  As TipoOrdenacion
Private mCriteriosAgrupacion  As TipoAgrupacion
Private mModalidadJuego       As ModalidadJuego
Private mPronosticos          As Integer
Private mDiasAnalisis         As Integer
Private mNumeroSorteos        As Integer
Private mErrorValidacion      As Long        ' Control de errores de validación

'--- Constantes ---------------------------------------------------------------*
Private Const ERR_NUMERO01              As Long = 1
Private Const ERR_NUMERO02              As Long = 2
Private Const ERR_NUMERO03              As Long = 4
Private Const ERR_NUMERO04              As Long = 8

'--- Mensajes -----------------------------------------------------------------*
Private Const MSG_ERRORESVALIDACION         As String = "Los datos del formulario no cumplen las siguientes validaciones:" & vbCrLf
Private Const MSG_CAMPOOBLIGATORIO          As String = "* El campo xxxx es un campo obligatorio."
Private Const MSG_CAMPONONUMERICO           As String = "* Debe introducir un dato numérico."
Private Const MSG_FECHANOVALIDA             As String = "* La fecha no es una fecha válida."
Private Const MSG_DATONOSELECCIONADO        As String = "* No se ha seleccionado ningún elemento."

'--- Errores ------------------------------------------------------------------*
'--- Propiedades --------------------------------------------------------------*
Public EntidadNegocio       As EntidadNegocio
Public TipoProcedimiento    As ProcedimientoMetodo
Public TipoMuestra          As Boolean              'False = dias, True = Registros

'--- Métodos Privados ---------------------------------------------------------*
Private Sub Class_Initialize()
    Set EntidadNegocio = New EntidadNegocio
    mSentidoOrdenacion = True         '   Ascendente
    EntidadNegocio.Id = 1
    mOrden = 1
    mCriteriosOrdenacion = ordSinDefinir
    mCriteriosAgrupacion = grpSinDefinir
    mModalidadJuego = LP_LB_6_49
    mPronosticos = 0
    mDiasAnalisis = 0
    mNumeroSorteos = 0
End Sub
'---------------------------------------------------------------------------------------
' Procedure : Id
' Author    : Charly
' Date      : 12/03/2012
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get Id() As Integer
    Id = Me.EntidadNegocio.Id
End Property
Public Property Let Id(ByVal iId As Integer)
     Me.EntidadNegocio.Id = iId
End Property

'---------------------------------------------------------------------------------------
' Procedure : Orden
' Author    : Charly
' Date      : 12/03/2012
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get Orden() As Integer
    Orden = mOrden
End Property

Public Property Let Orden(ByVal iOrden As Integer)
    mOrden = iOrden
End Property

'---------------------------------------------------------------------------------------
' Procedure : CriteriosOrdenacion
' Author    : Charly
' Date      : 12/03/2012
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get CriteriosOrdenacion() As TipoOrdenacion
    CriteriosOrdenacion = mCriteriosOrdenacion
End Property

Public Property Let CriteriosOrdenacion(ByVal eCriteriosOrdenacion As TipoOrdenacion)
    mCriteriosOrdenacion = eCriteriosOrdenacion
End Property

'---------------------------------------------------------------------------------------
' Procedure : SentidoOrdenacion
' Author    : Charly
' Date      : 12/03/2012
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get SentidoOrdenacion() As Boolean
    SentidoOrdenacion = mSentidoOrdenacion
End Property

Public Property Let SentidoOrdenacion(ByVal bSentidoOrdenacion As Boolean)
    mSentidoOrdenacion = bSentidoOrdenacion
End Property

'---------------------------------------------------------------------------------------
' Procedure : CriteriosAgrupacion
' Author    : Charly
' Date      : 12/03/2012
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get CriteriosAgrupacion() As TipoAgrupacion
    CriteriosAgrupacion = mCriteriosAgrupacion
End Property

Public Property Let CriteriosAgrupacion(ByVal eCriteriosAgrupacion As TipoAgrupacion)
    mCriteriosAgrupacion = eCriteriosAgrupacion
End Property

'---------------------------------------------------------------------------------------
' Procedure : Pronosticos
' Author    : Charly
' Date      : 12/03/2012
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get Pronosticos() As Integer
    Pronosticos = mPronosticos
End Property

Public Property Let Pronosticos(ByVal iPronosticos As Integer)
    mPronosticos = iPronosticos
End Property

'---------------------------------------------------------------------------------------
' Procedure : ModalidadJuego
' Author    : Charly
' Date      : 12/03/2012
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get ModalidadJuego() As ModalidadJuego
    ModalidadJuego = mModalidadJuego
End Property

Public Property Let ModalidadJuego(ByVal eModalidadJuego As ModalidadJuego)
    mModalidadJuego = eModalidadJuego
End Property

'---------------------------------------------------------------------------------------
' Procedure : DiasAnalisis
' Author    : Charly
' Date      : 12/03/2012
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get DiasAnalisis() As Integer
    DiasAnalisis = mDiasAnalisis
End Property

Public Property Let DiasAnalisis(ByVal iDiasAnalisis As Integer)
    mDiasAnalisis = iDiasAnalisis
End Property

'---------------------------------------------------------------------------------------
' Procedure : NumeroSorteos
' Author    : Charly
' Date      : 12/03/2012
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get NumeroSorteos() As Integer
    NumeroSorteos = mNumeroSorteos
End Property

Public Property Let NumeroSorteos(ByVal iNumeroSorteos As Integer)
    mNumeroSorteos = iNumeroSorteos
End Property
'--- Métodos Públicos ---------------------------------------------------------*
'---------------------------------------------------------------------------------------
' Procedure : TipoProcedimientoTostring
' Author    : Charly
' Date      : 10/04/2012
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Function TipoProcedimientoTostring() As String
    Dim m_vMatriz As Variant
    m_vMatriz = Split(NOMBRES_PROCEDIMIENTOMETODO, ";")
    TipoProcedimientoTostring = m_vMatriz(TipoProcedimiento)
End Function
'---------------------------------------------------------------------------------------
' Procedure : ToString
' Author    : Charly
' Date      : 12/03/2012
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Function ToString() As String
    Dim strTmp As String
  On Error GoTo ToString_Error
  
    strTmp = "#" + CStr(Me.EntidadNegocio.Id)
    strTmp = strTmp + " " + OrdenacionToString()
    strTmp = strTmp + "(" + IIf(mSentidoOrdenacion, "A", "D") + ")"
    strTmp = strTmp + " " + AgrupacionToString()
    strTmp = strTmp + " Dias(" + CStr(mDiasAnalisis) + ")"
    ToString = strTmp
 
  On Error GoTo 0
ToString__CleanExit:
    Exit Function
ToString_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    '   Audita el error
    Call HandleException(ErrNumber, ErrDescription, "Metodo.ToString", ErrSource)
    '   Lanza el Error
    Err.Raise ErrNumber, "Metodo.ToString", ErrDescription
End Function

'---------------------------------------------------------------------------------------
' Procedure : OrdenacionToString
' Author    : Charly
' Date      : 19/03/2012
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Function OrdenacionToString() As String
    Dim m_vMatriz As Variant
    m_vMatriz = Split(NOMBRES_ORDENACION, ";")
    OrdenacionToString = m_vMatriz(mCriteriosOrdenacion)
End Function

'---------------------------------------------------------------------------------------
' Procedure : OrdenacionToString
' Author    : Charly
' Date      : 19/03/2012
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Function AgrupacionToString() As String
    Dim m_vMatriz As Variant
    m_vMatriz = Split(NOMBRES_AGRUPACION, ";")
    AgrupacionToString = m_vMatriz(mCriteriosAgrupacion)
End Function

'---------------------------------------------------------------------------------------
' Procedure : EsValido
' Author    : Charly
' Date      : 19/02/2020
' Purpose   : Comprobar que los parametros son correctos en función del tipo de método
'---------------------------------------------------------------------------------------
'
Public Function EsValido() As Boolean
  On Error GoTo EsValido_Error
    '
    '   Inicializamos variables de control
    '
     mErrorValidacion = 0
    '
    ' TODO: Validar consistencia de metodo
    '
    '       Validación de campo obligatorio
    '
'    If (Len(txtTexto.Text) = 0) Then
'        mErrorValidacion = mErrorValidacion + ERR_CAMPOOBLIGATORIO
'    End If
    '
    '       Validación de campo no numérico
    '
'    If (Len(txtNumero.Text) > 0) And (Not IsNumeric(txtNumero.Text)) Then
'        mErrorValidacion = mErrorValidacion + ERR_CAMPONONUMERICO
'    End If
    '
    '       Validación de Campo no fecha
    '
'    If (Len(txtFecha.Text) = 0) and (not IsDate(txtFecha.Text) Then
'        mErrorValidacion = mErrorValidacion + ERR_FECHANOVALIDA
'    End If
    '
    '       Validación de combo no seleccionado
    '
'    If (cboControl.ListIndex = -1) Then
'        mErrorValidacion = mErrorValidacion + ERR_DATONOSELECCIONADO
'    End If
    '
    '
    '   Evaluamos el código de error
    '
    If mErrorValidacion = 0 Then
        EsValido = True
    Else
        EsValido = False
    End If
  
  On Error GoTo 0
EsValido__CleanExit:
    Exit Function
EsValido_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    '   Audita el error
    Call HandleException(ErrNumber, ErrDescription, "Metodo.EsValido", ErrSource)
    '   Lanza el Error
    Err.Raise ErrNumber, "Metodo.EsValido", ErrDescription
End Function
'' *===========(EOF): Metodo.cls
