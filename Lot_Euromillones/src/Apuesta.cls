VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Apuesta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' *============================================================================*
' *
' *     Fichero    : Apuesta.cls
' *
' *     Autor      : Carlos Almela Baeza
' *     Creación   : 10/11/2013
' *     Revisión   :
' *     Versión    : 1.0
' *     Propósito  : Representa una sugerencia o una apuesta de un Boleto
' *
' *============================================================================*
Option Explicit
Option Base 0
'
'--- Variables Privadas -------------------------------------------------------*

'--- Propiedades --------------------------------------------------------------*
Public EntidadNegocio   As EntidadNegocio
Public Combinacion      As Combinacion
Public Estrellas        As Combinacion
Public Fecha            As Date
Public IdBoleto         As Integer
Public SeHaJugado       As Boolean
Public FechaFinVigencia As Date
Public Metodo           As String
Public Juego            As Juego
Public Semana           As Integer
Public Dia              As String
Public Aciertos         As Integer
Public CategoriaPremio  As String
Public ImportePremio    As Currency

'---------------------------------------------------------------------------------------
' Procedure : NumeroApuestas
' Author    : Charly
' Date      : 04/09/2020
' Purpose   : Deduce según los pronosticos y estrellas si es una apuesta múltiple
'---------------------------------------------------------------------------------------
'
Public Property Get NumeroApuestas()
Select Case Juego
    Case LoteriaPrimitiva, Bonoloto:
        Select Case TotNumeros
            Case 5: NumeroApuestas = 44
            Case 6: NumeroApuestas = 1
            Case 7: NumeroApuestas = 7
            Case 8: NumeroApuestas = 28
            Case 9: NumeroApuestas = 84
            Case 10: NumeroApuestas = 210
            Case 11: NumeroApuestas = 462
            Case Else: NumeroApuestas = 0
        End Select
    
    Case GordoPrimitiva:
        Select Case TotNumeros
            Case 5: NumeroApuestas = 1
            Case 6: NumeroApuestas = 6
            Case 7: NumeroApuestas = 21
            Case 8: NumeroApuestas = 56
            Case 9: NumeroApuestas = 126
            Case 10: NumeroApuestas = 252
            Case 11: NumeroApuestas = 462
            Case Else: NumeroApuestas = 0
        End Select
    
    Case Euromillones:
        Select Case TotNumeros
        Case 5:
            Select Case TotEstrellas
                Case Is = 2: NumeroApuestas = 1
                Case Is = 3: NumeroApuestas = 3
                Case Is = 4: NumeroApuestas = 6
                Case Is = 5: NumeroApuestas = 10
            End Select
        Case 6:
            Select Case TotEstrellas
                Case Is = 2: NumeroApuestas = 6
                Case Is = 3: NumeroApuestas = 18
                Case Is = 4: NumeroApuestas = 36
                Case Is = 5: NumeroApuestas = 60
            End Select
            
        Case 7:
            Select Case TotEstrellas
                Case Is = 2: NumeroApuestas = 21
                Case Is = 3: NumeroApuestas = 63
                Case Is = 4: NumeroApuestas = 126
                Case Is = 5: NumeroApuestas = 210
            End Select
            
        Case 8:
            Select Case TotEstrellas
                Case Is = 2: NumeroApuestas = 56
                Case Is = 3: NumeroApuestas = 168
                Case Is = 4: NumeroApuestas = 336
                Case Is = 5: NumeroApuestas = 560
            End Select
            
        Case 9:
            Select Case TotEstrellas
                Case Is = 2: NumeroApuestas = 126
                Case Is = 3: NumeroApuestas = 378
                Case Is = 4: NumeroApuestas = 756
                Case Is = 5: NumeroApuestas = 1260
            End Select
            
        Case 10:
            Select Case TotEstrellas
                Case Is = 2: NumeroApuestas = 252
                Case Is = 3: NumeroApuestas = 756
                Case Is = 4: NumeroApuestas = 1512
                Case Is = 5: NumeroApuestas = 2520
            End Select
        
        Case Else:
            NumeroApuestas = 0
        End Select
End Select
End Property


'---------------------------------------------------------------------------------------
' Procedure : EsMultiple
' Author    : Charly
' Date      : 26/08/2020
' Purpose   : Deduce según los pronosticos y estrellas si es una apuesta múltiple
'---------------------------------------------------------------------------------------
'
Public Property Get EsMultiple() As Boolean
    Select Case Juego
        Case LoteriaPrimitiva, Bonoloto:
            If Pronosticos = 5 Or Pronosticos > 6 Then
                EsMultiple = True
            Else
                EsMultiple = False
            End If
            
        Case GordoPrimitiva:
            If Pronosticos > 5 Then
                EsMultiple = True
            Else
                EsMultiple = False
            End If
            
        Case Euromillones:
            If Pronosticos = 5 And TotEstrellas = 2 Then
                EsMultiple = False
            Else
                EsMultiple = True
            End If
    End Select
End Property


'---------------------------------------------------------------------------------------
' Procedure : Texto
' Author    : Charly
' Date      : 17/11/2013
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get Texto() As String
    Select Case Juego
        Case Bonoloto, LoteriaPrimitiva, GordoPrimitiva:
            Texto = Me.Combinacion.ToString(True)
        
        Case Euromillones
            Texto = Me.Combinacion.ToString(True) & "E" & Me.Estrellas.ToString(True)
    End Select
End Property

'---------------------------------------------------------------------------------------
' Procedure : Total Numeros
' Author    : Charly
' Date      : 23/08/2020
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get TotNumeros() As Integer
    TotNumeros = Me.Combinacion.Count
End Property
'---------------------------------------------------------------------------------------
' Procedure : Total Estrellas
' Author    : Charly
' Date      : 23/08/2020
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get TotEstrellas() As Integer
    TotEstrellas = Me.Estrellas.Count
End Property

'---------------------------------------------------------------------------------------
' Procedure : Pronosticos
' Author    : Charly
' Date      : 17/11/2013
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get Pronosticos() As Integer
    Pronosticos = TotNumeros
End Property



'---------------------------------------------------------------------------------------
' Procedure : Coste
' Author    : Charly
' Date      : 17/11/2013
' Purpose   : Calcular el coste total de la apuesta
'---------------------------------------------------------------------------------------
'
Public Property Get Coste() As Double
    Select Case Juego
        Case LoteriaPrimitiva:
            Coste = NumeroApuestas * 1
            
        Case Bonoloto:
            Coste = NumeroApuestas * 0.5
            
        Case GordoPrimitiva:
            Coste = NumeroApuestas * 1.5
            
        Case Euromillones:
           Coste = NumeroApuestas * 2.5
    End Select
End Property
'--- Constantes ---------------------------------------------------------------*
'--- Mensajes -----------------------------------------------------------------*
'--- Errores ------------------------------------------------------------------*
'--- Métodos Privados ---------------------------------------------------------*
Private Sub Class_Initialize()
    Set Combinacion = New Combinacion
    Set Estrellas = New Combinacion
    Set EntidadNegocio = New EntidadNegocio
    Juego = JUEGO_DEFECTO
    IdBoleto = 0
    SeHaJugado = False
    CategoriaPremio = ""
    ImportePremio = 0
End Sub

Private Sub Class_Terminate()
    Set Combinacion = Nothing
    Set Estrellas = Nothing
    Set EntidadNegocio = Nothing
End Sub


Private Function CalApuestas() As Integer
    Dim mPronos As Integer
    mPronos = Pronosticos
    
End Function
'---------------------------------------------------------------------------------------
' Procedure : Constructor
' Author    : Charly
' Date      :
' Purpose   : Constructor del objeto a partir de un rango de Excel
'---------------------------------------------------------------------------------------
'
Public Sub Constructor(vNewValue As Range)
    Dim mTexto As String
    Dim i As Integer
    Dim l As Double
  On Error GoTo Constructor_Error
    
    Select Case JUEGO_DEFECTO
        Case Bonoloto:
            With vNewValue
                EntidadNegocio.Id = .Cells(1, 1).Value
                IdBoleto = .Cells(1, 2).Value
                Juego = IIf(.Cells(1, 3).Value = "BL", Bonoloto, JUEGO_DEFECTO)
                Fecha = CDate(.Cells(1, 4).Value)
                Semana = CInt(.Cells(1, 5).Value)
                Dia = .Cells(1, 6).Value
                mTexto = .Cells(1, 7).Value & "-" & _
                         .Cells(1, 8).Value & "-" & _
                         .Cells(1, 9).Value & "-" & _
                         .Cells(1, 10).Value & "-" & _
                         .Cells(1, 11).Value & "-" & _
                         .Cells(1, 12).Value & "-" & _
                         .Cells(1, 13).Value & "-" & _
                         .Cells(1, 14).Value & "-" & _
                         .Cells(1, 15).Value & "-" & _
                         .Cells(1, 16).Value & "-" & _
                         .Cells(1, 17).Value
                Combinacion.Texto = mTexto
                CategoriaPremio = .Cells(1, 23).Value
                If IsNumeric(.Cells(1, 25).Value) Then
                    ImportePremio = CCur(.Cells(1, 24).Value)
                End If
                Metodo = .Cells(1, 25).Value
            End With
            
        Case LoteriaPrimitiva:
            With vNewValue
                EntidadNegocio.Id = .Cells(1, 1).Value
                IdBoleto = .Cells(1, 2).Value
                Juego = IIf(.Cells(1, 3).Value = "LP", LoteriaPrimitiva, JUEGO_DEFECTO)
                Fecha = CDate(.Cells(1, 4).Value)
                Semana = CInt(.Cells(1, 5).Value)
                Dia = .Cells(1, 6).Value
                mTexto = .Cells(1, 7).Value & "-" & _
                         .Cells(1, 8).Value & "-" & _
                         .Cells(1, 9).Value & "-" & _
                         .Cells(1, 10).Value & "-" & _
                         .Cells(1, 11).Value & "-" & _
                         .Cells(1, 12).Value & "-" & _
                         .Cells(1, 13).Value & "-" & _
                         .Cells(1, 14).Value & "-" & _
                         .Cells(1, 15).Value & "-" & _
                         .Cells(1, 16).Value & "-" & _
                         .Cells(1, 17).Value
                Combinacion.Texto = mTexto
                CategoriaPremio = .Cells(1, 23).Value
                If IsNumeric(.Cells(1, 25).Value) Then
                    ImportePremio = CCur(.Cells(1, 24).Value)
                End If
                Metodo = .Cells(1, 25).Value
            End With
            
        Case GordoPrimitiva:
             With vNewValue
                EntidadNegocio.Id = .Cells(1, 1).Value
                IdBoleto = .Cells(1, 2).Value
                Juego = IIf(.Cells(1, 3).Value = "GP", GordoPrimitiva, JUEGO_DEFECTO)
                Fecha = CDate(.Cells(1, 4).Value)
                Semana = CInt(.Cells(1, 5).Value)
                Dia = .Cells(1, 6).Value
                mTexto = .Cells(1, 7).Value & "-" & _
                         .Cells(1, 8).Value & "-" & _
                         .Cells(1, 9).Value & "-" & _
                         .Cells(1, 10).Value & "-" & _
                         .Cells(1, 11).Value & "-" & _
                         .Cells(1, 12).Value & "-" & _
                         .Cells(1, 13).Value & "-" & _
                         .Cells(1, 14).Value & "-" & _
                         .Cells(1, 15).Value & "-" & _
                         .Cells(1, 16).Value & "-" & _
                         .Cells(1, 17).Value
                Combinacion.Texto = mTexto
                CategoriaPremio = .Cells(1, 23).Value
                If IsNumeric(.Cells(1, 25).Value) Then
                    ImportePremio = CCur(.Cells(1, 24).Value)
                End If
                Metodo = .Cells(1, 25).Value
            End With
       
        Case Euromillones:
             With vNewValue
                EntidadNegocio.Id = .Cells(1, 1).Value
                IdBoleto = .Cells(1, 2).Value
                Juego = IIf(.Cells(1, 3).Value = "EU", Euromillones, JUEGO_DEFECTO)
                Fecha = CDate(.Cells(1, 4).Value)
                Semana = CInt(.Cells(1, 5).Value)
                Dia = .Cells(1, 6).Value
                mTexto = .Cells(1, 7).Value & "-" & _
                         .Cells(1, 8).Value & "-" & _
                         .Cells(1, 9).Value & "-" & _
                         .Cells(1, 10).Value & "-" & _
                         .Cells(1, 11).Value & "-" & _
                         .Cells(1, 12).Value & "-" & _
                         .Cells(1, 13).Value & "-" & _
                         .Cells(1, 14).Value & "-" & _
                         .Cells(1, 15).Value & "-" & _
                         .Cells(1, 16).Value & "-" & _
                         .Cells(1, 17).Value
                Combinacion.Texto = mTexto
                mTexto = .Cells(1, 18).Value & "-" & _
                         .Cells(1, 19).Value & "-" & _
                         .Cells(1, 20).Value & "-" & _
                         .Cells(1, 21).Value & "-" & _
                         .Cells(1, 22).Value
                Estrellas.Texto = mTexto
                CategoriaPremio = .Cells(1, 23).Value
                If IsNumeric(.Cells(1, 24).Value) Then
                    ImportePremio = CCur(.Cells(1, 24).Value)
                End If
                Metodo = .Cells(1, 25).Value
            End With
    
    End Select
    '
    '   Calculamos los pronosticos
    '
    i = Pronosticos
    
   On Error GoTo 0
   Exit Sub
Constructor_Error:
   Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
   ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
   Call HandleException(ErrNumber, ErrDescription, "Apuesta.Constructor")
   Err.Raise ErrNumber, ErrSource, ErrDescription
End Sub

' *===========(EOF): Apuesta.cls
