VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ParametrosMetodoOld"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Private mIdentificador          As Integer
Private mOrdenacion             As TiposMetodos
Private mDiasMuestra            As Integer
Private mRetardo                As Integer
Private m_iTipoMetodo           As Integer
Private m_bSentidoOrdenacion    As Boolean
Private m_iRegistrosMuestra     As Integer
Private m_iAgrupacion           As Integer
'
'
'
Public Property Get Id() As Variant
    Id = mIdentificador
End Property

Public Property Let Id(ByVal vNewValue As Variant)
    mIdentificador = vNewValue
End Property
'
'
'
'
Public Property Get Ordenacion() As TiposMetodos
    Ordenacion = mOrdenacion
End Property

Public Property Let Ordenacion(ByVal vNewValue As TiposMetodos)
    mOrdenacion = vNewValue
End Property
'
'
'
Public Property Get DiasMuestra() As Variant
    DiasMuestra = mDiasMuestra
End Property

Public Property Let DiasMuestra(ByVal vNewValue As Variant)
    mDiasMuestra = vNewValue
End Property
'
'
'
Public Property Get DiasRetardo() As Variant
    DiasRetardo = mRetardo
End Property

Public Property Let DiasRetardo(ByVal vNewValue As Variant)
    mRetardo = vNewValue
End Property
'
'
'
Public Function ToString() As String
    Dim m As MetodoOld, strNombres As Variant
    Dim strTmp As String
    Set m = New MetodoOld
    strTmp = "#" + CStr(mIdentificador)
    strNombres = m.LiteralesMetodos
    strTmp = strTmp + " " + strNombres(mOrdenacion)
    strTmp = strTmp + " Dias(" + CStr(mDiasMuestra) + ")/"
    strTmp = strTmp + "Rtrd(" + CStr(mRetardo) + ")"
    ToString = strTmp
End Function

Public Property Get TipoMetodo() As Integer

    TipoMetodo = m_iTipoMetodo

End Property

Public Property Let TipoMetodo(ByVal iTipoMetodo As Integer)

    m_iTipoMetodo = iTipoMetodo

End Property

Public Property Get SentidoOrdenacion() As Boolean

    SentidoOrdenacion = m_bSentidoOrdenacion

End Property

Public Property Let SentidoOrdenacion(ByVal bSentidoOrdenacion As Boolean)

    m_bSentidoOrdenacion = bSentidoOrdenacion

End Property

Public Property Get RegistrosMuestra() As Integer

    RegistrosMuestra = m_iRegistrosMuestra

End Property

Public Property Let RegistrosMuestra(ByVal iRegistrosMuestra As Integer)

    m_iRegistrosMuestra = iRegistrosMuestra

End Property

Public Property Get Agrupacion() As Integer

    Agrupacion = m_iAgrupacion

End Property

Public Property Let Agrupacion(ByVal iAgrupacion As Integer)

    m_iAgrupacion = iAgrupacion

End Property
