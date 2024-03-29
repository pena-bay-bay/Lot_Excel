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
' *     Creaci�n   : ma., 28/ene/2020 20:40:03
' *     Versi�n    : 1.0
' *     Prop�sito  : Clase que contiene los par�metros de una sugerencia
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
Private mErrorValidacion      As Long        ' Control de errores de validaci�n
Private mMensajeValidacion    As String
Private mColFilters           As FiltrosCombinacion  '

'--- Constantes ---------------------------------------------------------------*
'--- Errores ------------------------------------------------------------------*
Private Const ERR_TIPOPROCOBLIGATORIO   As Long = 1
Private Const ERR_PRONOOBLOGATORIO      As Long = 2
Private Const ERR_PRONOOUTRANGE         As Long = 4
Private Const ERR_ORDENACIONOBLIGATORIO As Long = 8
Private Const ERR_MUESTRAOBLIGATORIA    As Long = 16
Private Const ERR_NUMDIASIGUALCERO      As Long = 32
Private Const ERR_NUMREGISTROSIGUALCERO As Long = 64
Private Const ERR_FILTROSOBLIGATORIOS   As Long = 128

'--- Mensajes -----------------------------------------------------------------*
Private Const MSG_ERRORESVALIDACION         As String = "La informaci�n del Metodo de Sugerencia no cumplen las siguientes validaciones:" & vbCrLf
Private Const MSG_TIPOPROCOBLIGATORIO       As String = "* El tipo de procedimiento es obligatorio."
Private Const MSG_PRONOOBLOGATORIO          As String = "* Con filtros de combinaci�n, los pronosticos son obligatorios."
Private Const MSG_PRONOOUTRANGE             As String = "* Pronostico fuera de rango para el concurso seleccionado."
Private Const MSG_ORDENACIONOBLIGATORIO     As String = "* Criterio de ordenaci�n obligatorio."
Private Const MSG_MUESTRAOBLIGATORIA        As String = "* Criterio de muestra obligatorio."
Private Const MSG_NUMDIASIGUALCERO          As String = "* N�mero de d�as debe ser mayor que 0."
Private Const MSG_NUMREGISTROSIGUALCERO     As String = "* N�mero de registros debe ser mayor que 0."
Private Const MSG_FILTROSOBLIGATORIOS       As String = "* Filtros de combinaci�n obligatorios."


'--- Propiedades --------------------------------------------------------------*
Public EntidadNegocio       As EntidadNegocio
Public TipoProcedimiento    As ProcedimientoMetodo
Public TipoMuestra          As Boolean              'False = dias, True = Registros

Public Property Set Filtros(datFiltro As FiltrosCombinacion)
    Set mColFilters = datFiltro
End Property

Public Property Get Filtros() As FiltrosCombinacion
    Set Filtros = mColFilters
End Property

'--- M�todos Privados ---------------------------------------------------------*
Private Sub Class_Initialize()
    Set EntidadNegocio = New EntidadNegocio
    Set mColFilters = New FiltrosCombinacion
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
' Date      : 28/ene/2020 20:40:03
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
' Procedure : CriteriosOrdenacion
' Author    : Charly
' Date      : 28/ene/2020 20:40:03
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
' Date      : 28/ene/2020 20:40:03
' Purpose   : Sentido de la ordenaci�n: True - Ascendente, False - Descendente
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
' Date      : 28/ene/2020 20:40:03
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
' Date      : 28/ene/2020 20:40:03
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
' Date      : 28/ene/2020 20:40:03
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
' Date      : 28/ene/2020 20:40:03
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
' Date      : 28/ene/2020 20:40:03
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get NumeroSorteos() As Integer
    NumeroSorteos = mNumeroSorteos
End Property

Public Property Let NumeroSorteos(ByVal iNumeroSorteos As Integer)
    mNumeroSorteos = iNumeroSorteos
End Property

'--- M�todos P�blicos ---------------------------------------------------------*


'---------------------------------------------------------------------------------------
' Procedure : TipoProcedimientoTostring
' Author    : Charly
' Date      : 28/ene/2020 20:40:03
' Purpose   :
'---------------------------------------------------------------------------------------
Public Function TipoProcedimientoTostring() As String
    Dim m_vMatriz As Variant
    m_vMatriz = Split(NOMBRES_PROCEDIMIENTOMETODO, ";")
    TipoProcedimientoTostring = m_vMatriz(TipoProcedimiento)
End Function



'---------------------------------------------------------------------------------------
' Procedure : ToString
' Author    : Charly
' Date      : 28/ene/2020 20:40:03
' Purpose   : Expresi�n del objeto m�todo en formato string
'---------------------------------------------------------------------------------------
Public Function ToString() As String
    Dim strTmp As String
  On Error GoTo ToString_Error
  
    strTmp = "#" + CStr(Me.EntidadNegocio.Id)
    strTmp = strTmp + " " + TipoProcedimientoTostring
    If TipoProcedimiento = mtdEstadistico Or _
       TipoProcedimiento = mtdBomboCargado Or _
       TipoProcedimiento = mtdEstadCombinacion Then
        strTmp = strTmp + " " + OrdenacionToString()
        strTmp = strTmp + "(" + IIf(mSentidoOrdenacion, "A", "D") + ")"
        If mCriteriosAgrupacion <> grpSinDefinir Then
            strTmp = strTmp + " " + AgrupacionToString()
        End If
        If TipoMuestra Then
            strTmp = strTmp + " Registros(" + CStr(mNumeroSorteos) + ")"
        Else
            strTmp = strTmp + " Dias(" + CStr(mDiasAnalisis) + ")"
        End If
    End If
    
    If mColFilters.Count > 0 Then
        strTmp = strTmp + " Filtros(" + mColFilters.ToString + ")"
    End If
    
    ToString = strTmp
 
  On Error GoTo 0
ToString__CleanExit:
    Exit Function
ToString_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    Call HandleException(ErrNumber, ErrDescription, "Metodo.ToString", ErrSource)
    Err.Raise ErrNumber, "Metodo.ToString", ErrDescription
End Function

'---------------------------------------------------------------------------------------
' Procedure : OrdenacionToString
' Author    : Charly
' Date      : 28/ene/2020 20:40:03
' Purpose   :
'---------------------------------------------------------------------------------------
Public Function OrdenacionToString() As String
    Dim m_vMatriz As Variant
    m_vMatriz = Split(NOMBRES_ORDENACION, ";")
    OrdenacionToString = m_vMatriz(mCriteriosOrdenacion)
End Function

'---------------------------------------------------------------------------------------
' Procedure : OrdenacionToString
' Author    : Charly
' Date      : 28/ene/2020 20:40:03
' Purpose   :
'---------------------------------------------------------------------------------------
Public Function AgrupacionToString() As String
    Dim m_vMatriz As Variant
    m_vMatriz = Split(NOMBRES_AGRUPACION, ";")
    AgrupacionToString = m_vMatriz(mCriteriosAgrupacion)
End Function



'---------------------------------------------------------------------------------------
' Procedure : EsValido
' Author    : Charly
' Date      : ju., 16/abr/2020 18:36:09
' Purpose   : Comprobar que los parametros son correctos en funci�n del tipo de m�todo
'---------------------------------------------------------------------------------------
Public Function IsValid() As Boolean
  On Error GoTo IsValid_Error
    '
    '   Inicializamos variables de control
    '
    mErrorValidacion = 0
    '
    '   Tipo de m�todo obligatorio
    '
    If Me.TipoProcedimiento = 0 Then
        mErrorValidacion = mErrorValidacion + ERR_TIPOPROCOBLIGATORIO
    End If
    '
    '   Si los m�todos son: Aleatorio o bombo
    '
    If (Me.TipoProcedimiento = mtdAleatorio Or _
       Me.TipoProcedimiento = mtdBombo) Then
        '
        '   evaluamos pron�sticos
        '
        If mPronosticos = 0 Then
            mErrorValidacion = mErrorValidacion + ERR_PRONOOBLOGATORIO
        Else
            If mModalidadJuego = LP_LB_6_49 _
            And (mPronosticos < 5 Or mPronosticos > 11) Then
                mErrorValidacion = mErrorValidacion + ERR_PRONOOUTRANGE
            End If
            '
            If mModalidadJuego = GP_5_54 _
            And (mPronosticos < 5 Or mPronosticos > 11) Then
                mErrorValidacion = mErrorValidacion + ERR_PRONOOUTRANGE
            End If
            '
            If mModalidadJuego = EU_5_50 _
            And (mPronosticos < 5 Or mPronosticos > 10) Then
                mErrorValidacion = mErrorValidacion + ERR_PRONOOUTRANGE
            End If
            '
            If mModalidadJuego = EU_2_12 _
            And (mPronosticos < 2 Or mPronosticos > 5) Then
                mErrorValidacion = mErrorValidacion + ERR_PRONOOUTRANGE
            End If
        End If
    End If
    '
    '   Si los m�todos son: Bombo Cargado, Estadistico o estadistico combinaci�n
    '
    If (Me.TipoProcedimiento = mtdBomboCargado Or _
       Me.TipoProcedimiento = mtdEstadistico Or _
       Me.TipoProcedimiento = mtdEstadCombinacion) Then
        If mCriteriosOrdenacion = ordSinDefinir Then
            mErrorValidacion = mErrorValidacion + ERR_ORDENACIONOBLIGATORIO
        End If
        
        If Me.TipoMuestra = False And mDiasAnalisis = 0 Then
            mErrorValidacion = mErrorValidacion + ERR_NUMDIASIGUALCERO
        ElseIf Me.TipoMuestra = True And mNumeroSorteos = 0 Then
            mErrorValidacion = mErrorValidacion + ERR_NUMREGISTROSIGUALCERO
        End If

    End If
    If Me.TipoProcedimiento = mtdEstadCombinacion Then
        If mColFilters.Count = 0 Then
            mErrorValidacion = mErrorValidacion + ERR_FILTROSOBLIGATORIOS
        End If
    End If
    '
    '   Evaluamos el c�digo de error
    '
    If mErrorValidacion = 0 Then
        IsValid = True
    Else
        IsValid = False
    End If
  
  On Error GoTo 0
IsValid__CleanExit:
    Exit Function
IsValid_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    Call HandleException(ErrNumber, ErrDescription, "Metodo.IsValid", ErrSource)
    Err.Raise ErrNumber, "Metodo.IsValid", ErrDescription
End Function




'------------------------------------------------------------------------------*
' Procedimiento  : GetMessage
' Fecha          : ju., 16/abr/2020 18:36:09
' Prop�sito      : Devuelve el mensaje de error de validaci�n del m�todo
'------------------------------------------------------------------------------*
Public Function GetMessage() As String
    Dim sResult             As String
  On Error GoTo GetMessage_Error
    '
    '   Si no hay error inicializamos el mensaje
    '
    If mErrorValidacion = 0 Then
        sResult = ""
    Else
        '
        '   Si hay alg�n error inicializamos la cabecera del error
        '
        sResult = MSG_ERRORESVALIDACION
    End If
    '
    If (mErrorValidacion And ERR_TIPOPROCOBLIGATORIO) Then
        sResult = sResult & vbTab & MSG_TIPOPROCOBLIGATORIO & vbCrLf
    End If
    '
    If (mErrorValidacion And ERR_PRONOOBLOGATORIO) Then
        sResult = sResult & vbTab & MSG_PRONOOBLOGATORIO & vbCrLf
    End If
    '
    If (mErrorValidacion And ERR_PRONOOUTRANGE) Then
        sResult = sResult & vbTab & MSG_PRONOOUTRANGE & vbCrLf
    End If
    '
    If (mErrorValidacion And ERR_ORDENACIONOBLIGATORIO) Then
        sResult = sResult & vbTab & MSG_ORDENACIONOBLIGATORIO & vbCrLf
    End If
    '
    If (mErrorValidacion And ERR_MUESTRAOBLIGATORIA) Then
        sResult = sResult & vbTab & MSG_MUESTRAOBLIGATORIA & vbCrLf
    End If
    '
    If (mErrorValidacion And ERR_NUMDIASIGUALCERO) Then
        sResult = sResult & vbTab & MSG_NUMDIASIGUALCERO & vbCrLf
    End If
    '
    If (mErrorValidacion And ERR_NUMREGISTROSIGUALCERO) Then
        sResult = sResult & vbTab & MSG_NUMREGISTROSIGUALCERO & vbCrLf
    End If
    '
    If (mErrorValidacion And ERR_FILTROSOBLIGATORIOS) Then
        sResult = sResult & vbTab & MSG_FILTROSOBLIGATORIOS & vbCrLf
    End If
    '
    '   Devolvemos el mensaje
    '
    GetMessage = sResult
  On Error GoTo 0
GetMessage__CleanExit:
    Exit Function
GetMessage_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    Call HandleException(ErrNumber, ErrDescription, "Metodo.GetMessage", ErrSource)
    Err.Raise ErrNumber, "Metodo.GetMessage", ErrDescription
End Function

'' *===========(EOF): Metodo.cls
