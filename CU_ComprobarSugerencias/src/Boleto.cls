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

' #TODO: Incluir un metodo para evaluar el contenido del boleto y las apuestas
'        IsValid
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

'--- Propiedades --------------------------------------------------------------*
Public Apuestas As Collection
Public Juego As Juego
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

Public Property Get Multiple() As Boolean
    Multiple = m_bMultiple
End Property

Public Property Let Multiple(ByVal vNewValue As Boolean)
    m_bMultiple = vNewValue
End Property

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

Public Property Get Notas() As String
    Notas = m_sNotas
End Property

Public Property Let Notas(ByVal vNewValue As String)
    m_sNotas = vNewValue
End Property

Public Property Get Dia() As String
    Dia = m_sDia
End Property

Public Property Let Dia(ByVal vNewValue As String)
    m_sDia = vNewValue
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
' Procedure : JuegoTexto
' DateTime  : 12/ago/2007 19:10
' Author    : Carlos Almela Baeza
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get JuegoTexto() As String
    JuegoTexto = m_sJuego
End Property

Public Property Let JuegoTexto(ByVal sJuego As String)
    m_sJuego = sJuego
End Property

Public Property Get DiasValidez() As Integer
    Select Case True
    Case Semanal And (m_sJuego = "BL"):
        DiasValidez = 5
    Case Semanal And (m_sJuego = "LP"):
        DiasValidez = 2
    Case Semanal And (m_sJuego = "EU"):
        DiasValidez = 2
    Case Else
        DiasValidez = 1
    End Select
End Property

Public Property Get FechaValidez() As Date
    FechaValidez = m_dtFecha_Validez
End Property

Public Property Let FechaValidez(ByVal vNewValue As Date)
    m_dtFecha_Validez = vNewValue
End Property

'---------------------------------------------------------------------------------------
' Procedure : Fecha_Caducidad
' DateTime  : 12/ago/2007 19:10
' Author    : Carlos Almela Baeza
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get FechaCaducidad() As Date
    FechaCaducidad = m_dtFecha_Caducidad
End Property

Public Property Let FechaCaducidad(ByVal vNewValue As Date)
    m_dtFecha_Caducidad = vNewValue
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
Public Property Get NumeroApuestas() As Integer
    NumeroApuestas = m_iNumeroApuestas
End Property

Public Property Let NumeroApuestas(ByVal vNewValue As Integer)
    m_iNumeroApuestas = vNewValue
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
Public Property Get ImportePremios() As Currency
    ImportePremios = m_cPremios
End Property

Public Property Let ImportePremios(ByVal cPremios As Currency)
    m_cPremios = cPremios
End Property

Public Property Get ElMillon() As String
    ElMillon = m_sElMillon
End Property

Public Property Let ElMillon(ByVal vNewValue As String)
    m_sElMillon = vNewValue
End Property

'--- Métodos Privados ---------------------------------------------------------*
Private Sub Class_Initialize()
    Set Apuestas = New Collection
End Sub

'---------------------------------------------------------------------------------------
' Procedure : Constructor
' DateTime  : 12/ago/2007 19:11
' Author    : Carlos Almela Baeza
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Sub Constructor(Fila As Range)
  On Error GoTo Constructor_Error
    '
    '   Carga información según en Juego
    '
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
            m_iNumeroApuestas = .Cells(1, 12).Value
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
            m_iNumeroApuestas = .Cells(1, 12).Value
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
            m_iNumeroApuestas = .Cells(1, 12).Value
            m_cCoste = .Cells(1, 13).Value
            m_sCatPremios = .Cells(1, 14).Value
            m_cPremios = .Cells(1, 15).Value
            m_sNotas = .Cells(1, 16).Value
        End With
    
    End Select
    Select Case m_sJuego
        Case "BL":          Juego = Bonoloto
        Case "LP":          Juego = LoteriaPrimitiva
        Case "GP":          Juego = GordoPrimitiva
        Case "EU":          Juego = Euromillones
    End Select
    
  On Error GoTo 0
Constructor__CleanExit:
    Exit Sub
Constructor_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    Call HandleException(ErrNumber, ErrDescription, "Boleto.Constructor", ErrSource)
    Err.Raise ErrNumber, "Boleto.Constructor", ErrDescription
End Sub



'------------------------------------------------------------------------------*
' Procedimiento  : SetApuestas
' Fecha          : vi., 10/jul/2020 17:45:41
' Propósito      : carga las apuestas del boleto
'------------------------------------------------------------------------------*
'
Public Sub SetApuestas()
    Dim oBd     As BdDatos
    Dim oApt    As Apuesta
    Dim oRng    As Range
    Dim oFila   As Range
    
  On Error GoTo SetApuestas_Error
    '
    '   Definimos Apuesta y Base de datos
    '
    Set oBd = New BdDatos
    '
    '   Inicializamos la colección de apuestas
    '
    Set Apuestas = New Collection
    '
    '   Localizamos el rango de apuestas
    '
    Set oRng = oBd.GetApuestaByBoleto(m_iId)
    '
    '   Si encontramos apuestas
    '
    If Not (oRng Is Nothing) Then
        '
        '   Recoremos el rango de apuestas para montar la colección
        '
        For Each oFila In oRng.Rows
            '
            '   Creamos un objeto apuesta
            '
            Set oApt = New Apuesta
            '
            '   Cargamos la apuesta
            '
            oApt.Constructor oFila
            '
            '   Agregamos apuestas a la colección
            '
            Apuestas.Add oApt
        Next oFila
    Else
        Call HandleException(999, "Error, Boleto (" & CStr(m_iId) & ") sin apuestas. Fecha:" & Format(Me.FechaValidez, "dd/MM/yyyy"), _
        "Boleto.SetApuestas", "Boleto.SetApuestas")
    End If
    '
    '   Si hay una apuesta y es múltiple el boleto es multiple y las apuest
    '
    If Apuestas.Count = 1 Then
        If oApt.EsMultiple Then
            m_bMultiple = True
            m_iNumeroApuestas = oApt.NumeroApuestas
        End If
    Else
        m_iNumeroApuestas = Apuestas.Count
    End If

  
  On Error GoTo 0
    Exit Sub
SetApuestas_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    Call HandleException(ErrNumber, ErrDescription, "Boleto.SetApuestas", ErrSource)
    Err.Raise ErrNumber, "Boleto.SetApuestas", ErrDescription
End Sub
' *===========(EOF): Boleto.cls

