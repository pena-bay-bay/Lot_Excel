VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ParametrosMetodo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'---------------------------------------------------------------------------------------
' Module    : ParametrosMetodo
' Author    : Charly
' Date      : 12/03/2012
' Purpose   : Definir los parámetros de un método
'---------------------------------------------------------------------------------------


Option Explicit
Private m_bSentidoOrdenacion    As Boolean
Private m_iId                   As Integer
Private m_iOrden                As Integer
Private m_eCriteriosOrdenacion  As TipoOrdenacion
Private m_eCriteriosAgrupacion  As TipoAgrupacion
Private m_eModalidadJuego       As ModalidadJuego
Private m_iPronosticos          As Integer
Private m_iDiasAnalisis         As Integer
Private m_iNumeroSorteos        As Integer
'---------------------------------------------------------------------------------------
' Procedure : Id
' Author    : Charly
' Date      : 12/03/2012
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get Id() As Integer
    Id = m_iId
End Property

'---------------------------------------------------------------------------------------
' Procedure : Id
' Author    : Charly
' Date      : 12/03/2012
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Let Id(ByVal iId As Integer)
    m_iId = iId
End Property

'---------------------------------------------------------------------------------------
' Procedure : Orden
' Author    : Charly
' Date      : 12/03/2012
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get Orden() As Integer

    Orden = m_iOrden

End Property

'---------------------------------------------------------------------------------------
' Procedure : Orden
' Author    : Charly
' Date      : 12/03/2012
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Let Orden(ByVal iOrden As Integer)

    m_iOrden = iOrden

End Property

'---------------------------------------------------------------------------------------
' Procedure : CriteriosOrdenacion
' Author    : Charly
' Date      : 12/03/2012
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get CriteriosOrdenacion() As TipoOrdenacion

    CriteriosOrdenacion = m_eCriteriosOrdenacion

End Property

'---------------------------------------------------------------------------------------
' Procedure : CriteriosOrdenacion
' Author    : Charly
' Date      : 12/03/2012
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Let CriteriosOrdenacion(ByVal eCriteriosOrdenacion As TipoOrdenacion)

    m_eCriteriosOrdenacion = eCriteriosOrdenacion

End Property

'---------------------------------------------------------------------------------------
' Procedure : SentidoOrdenacion
' Author    : Charly
' Date      : 12/03/2012
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get SentidoOrdenacion() As Boolean

    SentidoOrdenacion = m_bSentidoOrdenacion

End Property

'---------------------------------------------------------------------------------------
' Procedure : SentidoOrdenacion
' Author    : Charly
' Date      : 12/03/2012
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Let SentidoOrdenacion(ByVal bSentidoOrdenacion As Boolean)

    m_bSentidoOrdenacion = bSentidoOrdenacion

End Property

'---------------------------------------------------------------------------------------
' Procedure : CriteriosAgrupacion
' Author    : Charly
' Date      : 12/03/2012
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get CriteriosAgrupacion() As TipoAgrupacion

    CriteriosAgrupacion = m_eCriteriosAgrupacion

End Property

'---------------------------------------------------------------------------------------
' Procedure : CriteriosAgrupacion
' Author    : Charly
' Date      : 12/03/2012
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Let CriteriosAgrupacion(ByVal eCriteriosAgrupacion As TipoAgrupacion)

    m_eCriteriosAgrupacion = eCriteriosAgrupacion

End Property

'---------------------------------------------------------------------------------------
' Procedure : Pronosticos
' Author    : Charly
' Date      : 12/03/2012
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get Pronosticos() As Integer

    Pronosticos = m_iPronosticos

End Property

'---------------------------------------------------------------------------------------
' Procedure : Pronosticos
' Author    : Charly
' Date      : 12/03/2012
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Let Pronosticos(ByVal iPronosticos As Integer)

    m_iPronosticos = iPronosticos

End Property

'---------------------------------------------------------------------------------------
' Procedure : ModalidadJuego
' Author    : Charly
' Date      : 12/03/2012
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get ModalidadJuego() As ModalidadJuego

    ModalidadJuego = m_eModalidadJuego

End Property

'---------------------------------------------------------------------------------------
' Procedure : ModalidadJuego
' Author    : Charly
' Date      : 12/03/2012
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Let ModalidadJuego(ByVal eModalidadJuego As ModalidadJuego)

    m_eModalidadJuego = eModalidadJuego

End Property

'---------------------------------------------------------------------------------------
' Procedure : DiasAnalisis
' Author    : Charly
' Date      : 12/03/2012
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get DiasAnalisis() As Integer

    DiasAnalisis = m_iDiasAnalisis

End Property

'---------------------------------------------------------------------------------------
' Procedure : DiasAnalisis
' Author    : Charly
' Date      : 12/03/2012
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Let DiasAnalisis(ByVal iDiasAnalisis As Integer)

    m_iDiasAnalisis = iDiasAnalisis

End Property

'---------------------------------------------------------------------------------------
' Procedure : NumeroSorteos
' Author    : Charly
' Date      : 12/03/2012
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get NumeroSorteos() As Integer

    NumeroSorteos = m_iNumeroSorteos

End Property

'---------------------------------------------------------------------------------------
' Procedure : NumeroSorteos
' Author    : Charly
' Date      : 12/03/2012
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Let NumeroSorteos(ByVal iNumeroSorteos As Integer)

    m_iNumeroSorteos = iNumeroSorteos

End Property

'---------------------------------------------------------------------------------------
' Procedure : Class_Initialize
' Author    : Charly
' Date      : 12/03/2012
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub Class_Initialize()
    m_bSentidoOrdenacion = True         '   Ascendente
    m_iId = 1
    m_iOrden = 1
    m_eCriteriosOrdenacion = ordSinDefinir
    m_eCriteriosAgrupacion = grpSinDefinir
    m_eModalidadJuego = LP_LB_6_49
    m_iPronosticos = 0
    m_iDiasAnalisis = 0
    m_iNumeroSorteos = 0
End Sub

'---------------------------------------------------------------------------------------
' Procedure : ToString
' Author    : Charly
' Date      : 12/03/2012
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Function ToString() As String
    Dim strTmp As String
    strTmp = "#" + CStr(m_iId)
    strTmp = strTmp + " " + OrdenacionToString()
    strTmp = strTmp + "(" + IIf(m_bSentidoOrdenacion, "A", "D") + ")"
    strTmp = strTmp + " " + AgrupacionToString()
    strTmp = strTmp + " Dias(" + CStr(m_iDiasAnalisis) + ")"
    ToString = strTmp
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
    OrdenacionToString = m_vMatriz(m_eCriteriosOrdenacion)
End Function

'---------------------------------------------------------------------------------------
' Procedure : AgrupacionToString
' Author    : Charly
' Date      : 19/03/2012
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Function AgrupacionToString() As String
Attribute AgrupacionToString.VB_Description = "Devuelve la cadena de texto correspondiente al código de agrupación"
    Dim m_vMatriz As Variant
    m_vMatriz = Split(NOMBRES_AGRUPACION, ";")
    AgrupacionToString = m_vMatriz(m_eCriteriosAgrupacion)
End Function

