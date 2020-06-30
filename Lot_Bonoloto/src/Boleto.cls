VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Boleto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' *============================================================================*
' *
' *     Fichero    : Boleto.cls
' *
' *     Autor      : Carlos Almela Baeza
' *     Creación   : 12/ago/2007 19:09
' *     Modificado : ma., 30/jun/2020 18:21:54
' *     Versión    : 1.1
' *     Propósito  : Objeto que contiene toda la información de un Boleto
' *
' *============================================================================*
Option Explicit
Option Base 0

'--- Variables Privadas -------------------------------------------------------*
Private m_iId As Integer
Private m_iSemana As Integer
Private m_sDia  As String
Private m_sJuego As String
Private m_dtFecha_Validez As Date
Private m_dtFecha_Caducidad As Date
Private m_bSemanal As Boolean
Private m_bMultiple As Boolean
Private m_iReintegro As Integer
Private m_lJoker As Long
Private m_sElMillon As String
Private m_iNumeroApuestas As Integer
Private m_cCoste As Currency
Private m_cPremios As Currency
Private m_sCatPremios As String
Private m_sNotas As String

'--- Métodos Privados ---------------------------------------------------------*
Private Sub Class_Initialize()
    Set Apuestas = New Collection
End Sub

'--- Propiedades --------------------------------------------------------------*
Public Apuestas As Collection
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
' Procedure : Constructor
' DateTime  : 12/ago/2007 19:11
' Author    : Carlos Almela Baeza
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Sub Constructor(Fila As Range)
    
    
    Select Case JUEGO_DEFECTO
    Case Bonoloto, LoteriaPrimitiva:
        With Fila
            m_iId = .Cells(1, 1).Value
            m_sJuego = .Cells(1, 2).Value
            m_dtFecha_Validez = CDate(.Cells(1, 3).Value)
            m_iSemana = .Cells(1, 4).Value
            m_sDia = .Cells(1, 5).Value
            m_dtFecha_Caducidad = CDate(.Cells(1, 6).Value)
            m_bMultiple = IIf(.Cells(1, 7).Value = "Si", True, False)
            m_bSemanal = IIf(.Cells(1, 8).Value = "Si", True, False)
            m_iReintegro = .Cells(1, 9).Value
            m_lJoker = .Cells(1, 10).Value
'            m_sElMillon = .Cells(1, 11).Value
            m_iNumero_Apuestas = .Cells(1, 12).Value
            m_cCoste = .Cells(1, 13).Value
            m_sCatPremios = .Cells(1, 14).Value
            m_cPremios = .Cells(1, 15).Value
            m_sNotas = .Cells(1, 16).Value
        End With
        
    Case GordoPrimitiva:
        With Fila
            m_iId = .Cells(1, 1).Value
            m_sJuego = .Cells(1, 2).Value
            m_dtFecha_Validez = CDate(.Cells(1, 3).Value)
            m_iSemana = .Cells(1, 4).Value
            m_sDia = .Cells(1, 5).Value
            m_dtFecha_Caducidad = CDate(.Cells(1, 6).Value)
            m_bMultiple = IIf(.Cells(1, 7).Value = "Si", True, False)
            m_bSemanal = IIf(.Cells(1, 8).Value = "Si", True, False)
            m_iReintegro = .Cells(1, 9).Value
'            m_lJoker = .Cells(1, 10).Value
'            m_sElMillon = .Cells(1, 11).Value
            m_iNumero_Apuestas = .Cells(1, 12).Value
            m_cCoste = .Cells(1, 13).Value
            m_sCatPremios = .Cells(1, 14).Value
            m_cPremios = .Cells(1, 15).Value
            m_sNotas = .Cells(1, 16).Value
        End With
    
    Case Euromillones:
        With Fila
            m_iId = .Cells(1, 1).Value
            m_sJuego = .Cells(1, 2).Value
            m_dtFecha_Validez = CDate(.Cells(1, 3).Value)
            m_iSemana = .Cells(1, 4).Value
            m_sDia = .Cells(1, 5).Value
            m_dtFecha_Caducidad = CDate(.Cells(1, 6).Value)
            m_bMultiple = IIf(.Cells(1, 7).Value = "Si", True, False)
            m_bSemanal = IIf(.Cells(1, 8).Value = "Si", True, False)
'            m_iReintegro = .Cells(1, 9).Value
'            m_lJoker = .Cells(1, 10).Value
            m_sElMillon = .Cells(1, 11).Value
            m_iNumero_Apuestas = .Cells(1, 12).Value
            m_cCoste = .Cells(1, 13).Value
            m_sCatPremios = .Cells(1, 14).Value
            m_cPremios = .Cells(1, 15).Value
            m_sNotas = .Cells(1, 16).Value
        End With
    
    
    End Select
End Sub


Public Property Get CategoriaPremio() As String
    CategoriaPremio = m_sCatPremios
End Property

Public Property Let CategoriaPremio(ByVal vNewValue As String)
    m_sCatPremios = vNewValue
End Property

Public Property Get Joker() As Long
    Joker = m_lJoker
End Property

Public Property Let Joker(ByVal vNewValue As Long)
    m_lJoker = vNewValue
End Property

' *===========(EOF): Boleto.cls
