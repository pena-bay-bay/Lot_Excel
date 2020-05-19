VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Boleto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'---------------------------------------------------------------------------------------
' Module    : Boleto
' DateTime  : 12/ago/2007 19:09
' Author    : Carlos Almela Baeza
' Purpose   : Objeto Documento que contiene la informaci�n del boleto jugado
'---------------------------------------------------------------------------------------
Private m_iId As Integer
Private m_iSemana As Integer
Private m_sJuego As String
Private m_dtFecha_Validez As Date
Private m_dtFecha_Caducidad As Date
Private m_bSemanal As Boolean
Private m_iReintegro As Integer
Private m_iNumero_Apuestas As Integer
Private m_cCoste As Currency
Private m_cPremios As Currency
Private m_cReintegros As Currency

'---------------------------------------------------------------------------------------
' Procedure : Id
' DateTime  : 12/ago/2007 19:10
' Author    : Carlos Almela Baeza
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get Id() As Integer
    Id = m_iId
End Property

Public Property Let Id(ByVal iId As Integer)
    m_iId = iId
End Property

'---------------------------------------------------------------------------------------
' Procedure : Semana
' DateTime  : 12/ago/2007 19:10
' Author    : Carlos Almela Baeza
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get Semana() As Integer
    Semana = m_iSemana
End Property

Public Property Let Semana(ByVal iSemana As Integer)
    m_iSemana = iSemana
End Property

'---------------------------------------------------------------------------------------
' Procedure : Juego
' DateTime  : 12/ago/2007 19:10
' Author    : Carlos Almela Baeza
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get Juego() As String
    Juego = m_sJuego
End Property

Public Property Let Juego(ByVal sJuego As String)
    m_sJuego = sJuego
End Property

'---------------------------------------------------------------------------------------
' Procedure : Dias_Validez
' DateTime  : 12/ago/2007 19:10
' Author    : Carlos Almela Baeza
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get Dias_Validez() As Integer
    Select Case True
    Case Semanal And (m_sJuego = "BL"):
        Dias_Validez = 4
    Case Semanal And (m_sJuego = "LP"):
        Dias_Validez = 2
    Case Else
        Dias_Validez = 1
    End Select
End Property
'---------------------------------------------------------------------------------------
' Procedure : Fecha_Validez
' DateTime  : 12/ago/2007 19:10
' Author    : Carlos Almela Baeza
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get Fecha_Validez() As Date
    Fecha_Validez = m_dtFecha_Validez
End Property

Public Property Let Fecha_Validez(ByVal dtFecha_Validez As Date)
    m_dtFecha_Validez = dtFecha_Validez
End Property

'---------------------------------------------------------------------------------------
' Procedure : Fecha_Caducidad
' DateTime  : 12/ago/2007 19:10
' Author    : Carlos Almela Baeza
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get Fecha_Caducidad() As Date
    Fecha_Caducidad = m_dtFecha_Caducidad
End Property

Public Property Let Fecha_Caducidad(ByVal dtFecha_Caducidad As Date)
    m_dtFecha_Caducidad = dtFecha_Caducidad
End Property

'---------------------------------------------------------------------------------------
' Procedure : Semanal
' DateTime  : 12/ago/2007 19:10
' Author    : Carlos Almela Baeza
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get Semanal() As Boolean
    Semanal = m_bSemanal
End Property

Public Property Let Semanal(ByVal bSemanal As Boolean)
    m_bSemanal = bSemanal
End Property

'---------------------------------------------------------------------------------------
' Procedure : Reintegro
' DateTime  : 12/ago/2007 19:10
' Author    : Carlos Almela Baeza
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get Reintegro() As Integer
    Reintegro = m_iReintegro
End Property

Public Property Let Reintegro(ByVal iReintegro As Integer)
    m_iReintegro = iReintegro
End Property

'---------------------------------------------------------------------------------------
' Procedure : Numero_Apuestas
' DateTime  : 12/ago/2007 19:10
' Author    : Carlos Almela Baeza
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get Numero_Apuestas() As Integer
    Numero_Apuestas = m_iNumero_Apuestas
End Property

Public Property Let Numero_Apuestas(ByVal iNumero_Apuestas As Integer)
    m_iNumero_Apuestas = iNumero_Apuestas
End Property

'---------------------------------------------------------------------------------------
' Procedure : Coste
' DateTime  : 12/ago/2007 19:10
' Author    : Carlos Almela Baeza
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get Coste() As Currency
    Coste = m_cCoste
End Property

Public Property Let Coste(ByVal cCoste As Currency)
    m_cCoste = cCoste
End Property

'---------------------------------------------------------------------------------------
' Procedure : Premios
' DateTime  : 12/ago/2007 19:10
' Author    : Carlos Almela Baeza
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get Premios() As Currency
    Premios = m_cPremios
End Property

Public Property Let Premios(ByVal cPremios As Currency)
    m_cPremios = cPremios
End Property

'---------------------------------------------------------------------------------------
' Procedure : Premio_Reintegros
' DateTime  : 12/ago/2007 19:10
' Author    : Carlos Almela Baeza
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get Premio_Reintegros() As Currency
    Reintegros = m_cReintegros
End Property

Public Property Let Premio_Reintegros(ByVal cReintegros As Currency)
    m_cReintegros = cReintegros
End Property

'---------------------------------------------------------------------------------------
' Procedure : Constructor
' DateTime  : 12/ago/2007 19:11
' Author    : Carlos Almela Baeza
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Sub Constructor(Fila As Range)
    With Fila
        m_iId = .Cells(1, 1).Value
        m_iSemana = .Cells(1, 2).Value
        m_sJuego = .Cells(1, 3).Value
        m_dtFecha_Validez = CDate(.Cells(1, 4).Value)
        m_dtFecha_Caducidad = CDate(.Cells(1, 5).Value)
        m_bSemanal = IIf(.Cells(1, 6).Value = "Si", True, False)
        m_iReintegro = IIf(IsEmpty(.Cells(1, 7).Value), -1, .Cells(1, 7).Value)
        m_iNumero_Apuestas = .Cells(1, 8).Value
        m_cCoste = .Cells(1, 9).Value
        m_cPremios = .Cells(1, 10).Value
        m_cReintegros = .Cells(1, 11).Value
    End With
End Sub

'---------------------------------------------------------------------------------------
' Procedure : Get_FechaValidez
' DateTime  : 12/ago/2007 19:11
' Author    : Carlos Almela Baeza
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Function Get_FechaValidez(Dia As Integer) As Date
    Get_FechaValidez = m_dtFecha_Validez
    'solo si el boleto es semanal se devuelven varias fechas de validez
    If Semanal Then
        ' Si el juebo es bonoloto el cuarto dia es el viernes
        If m_sJuego = "BL" Then
            If Dia < 3 Then
                Get_FechaValidez = m_dtFecha_Validez + Dia - 1
            Else
                Get_FechaValidez = m_dtFecha_Caducidad
            End If
        End If
        If m_sJuego = "LP" Then
            If Dia = 1 Then
                Get_FechaValidez = m_dtFecha_Validez
            Else
                Get_FechaValidez = m_dtFecha_Caducidad
            End If
        End If
    End If
End Function