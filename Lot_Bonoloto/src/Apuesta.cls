VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Apuesta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'---------------------------------------------------------------------------------------
' Module    : Apuesta
' Author    : Charly
' Date      : 10/11/2013
' Purpose   : Nueva version de Apuesta
'---------------------------------------------------------------------------------------

Option Explicit

Public Combinacion      As Combinacion
Public IdBoleto         As Integer
Public SeHaJugado       As Boolean
Public EntidadNegocio   As EntidadNegocio
Public EsMultiple       As Boolean
Public NumeroApuestas   As Integer
Public FechaAlta        As Date
Public FechaFinVigencia As Date
Public Metodo           As String
Public Juego            As Juego
'
Private m_dCoste As Double

Private Sub Class_Initialize()
    Set Combinacion = New Combinacion
    Set EntidadNegocio = New EntidadNegocio
    IdBoleto = 0
    SeHaJugado = False
    EsMultiple = False
    NumeroApuestas = 0
    FechaAlta = Date
End Sub

'---------------------------------------------------------------------------------------
' Procedure : Texto
' Author    : Charly
' Date      : 17/11/2013
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get Texto() As String

  On Error GoTo Texto_Error

    Texto = Me.Combinacion.Texto

   On Error GoTo 0
   Exit Property

Texto_Error:
     Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
   ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
   '   Audita el error
   Call HandleException(ErrNumber, ErrDescription, "Apuesta.Texto")
   '   Lanza el Error
   Err.Raise ErrNumber, ErrSource, ErrDescription
End Property

'---------------------------------------------------------------------------------------
' Procedure : Pronosticos
' Author    : Charly
' Date      : 17/11/2013
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get Pronosticos() As Integer

  On Error GoTo Pronosticos_Error
    Me.EsMultiple = True
    Pronosticos = Me.Combinacion.Count

    Select Case Pronosticos
        Case 5: NumeroApuestas = 44
            If Me.Juego = Euromillones Then
                Me.EsMultiple = False
            End If
        Case 6:
            NumeroApuestas = 1
            If Me.Juego = PrimitivaBonoloto Then
                Me.EsMultiple = False
            End If
        Case 7: NumeroApuestas = 7
        Case 8: NumeroApuestas = 28
        Case 9: NumeroApuestas = 84
        Case 10: NumeroApuestas = 210
        Case 11: NumeroApuestas = 462
        Case Else: NumeroApuestas = 0
    End Select
    
   On Error GoTo 0
   Exit Property

Pronosticos_Error:
     Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
   ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
   '   Audita el error
   Call HandleException(ErrNumber, ErrDescription, "Apuesta.Pronosticos")
   '   Lanza el Error
   Err.Raise ErrNumber, ErrSource, ErrDescription

End Property

Public Property Get Coste(datJuego As Juego) As Double
    Select Case datJuego
    Case LoteriaPrimitiva
        m_dCoste = NumeroApuestas * 1
    Case Bonoloto
        m_dCoste = NumeroApuestas * 0.5
    Case GordoPrimitiva
        m_dCoste = NumeroApuestas * 1.5
    Case Euromillones
       m_dCoste = NumeroApuestas * 2
    End Select

    Coste = m_dCoste
    

End Property
