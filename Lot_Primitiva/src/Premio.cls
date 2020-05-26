VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Premio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' *============================================================================*
' *
' *     Fichero    : Premio
' *
' *     Autor      : Carlos Almela Baeza
' *     Creación   : dom, 12/04/2009 22:47
' *     Versión    : 1.0
' *     Propósito  :
' *
' *
' *============================================================================*
Option Explicit
Option Base 0
Public key As String
Public BolasAcertadas As Integer
Public ComplementarioAcertado As Boolean
Public ClaveAcertada As Boolean
Public NumeroEstrellasAcertadas As Integer
Public FechaSorteo As Date
Public ModalidadJuego As Juego
Public Pronosticos As Integer
Private m_eNumerosAcertados As TiposAciertos
'
'
'
Private TablaMultiples(5) As Variant 'Coordenadas Aciertos / apuestas Premio ver Boleto
Private ImportesPremiosBL As Variant
Private ImportesPremiosLP As Variant
Private ImportesPremiosEU As Variant
Private ImportesPremiosGP As Variant
Private m_eCategoriaPremio As CategoriaPremio
Private m_sCategoriaPremio As String

' *============================================================================*
' *     Procedure  : GetPremioEsperado
' *     Fichero    : Premio
' *     Autor      : Carlos Almela Baeza
' *     Creacion   : dom, 12/04/2009
' *     Asunto     :
' *============================================================================*
'
Public Function GetPremioEsperado() As Currency
    Dim i As Integer, j As Integer
    Dim cResult As Currency
    Dim vImportes As Variant, vVar As Variant
  On Error GoTo GetPremioEsperado_Error
    '
    '   Seleccionar la matriz de importes segun la modalidad del juego
    '
    If ModalidadJuego = Bonoloto Then
        vImportes = ImportesPremiosBL
    Else
        If ModalidadJuego = LoteriaPrimitiva Then
            vImportes = ImportesPremiosLP
        End If
    End If
    '
    '
    '
    If Pronosticos <> 6 Then
        
        '
        '   Se selecciona la matriz para juego multiples pronosticos
        '
        CargaTablaMultiples
        '
        '
        '
        For i = 0 To UBound(TablaMultiples)
            vVar = TablaMultiples(i)
            If (vVar(0) = Me.NumerosAcertados) Then
                
                For j = 1 To 5
                    '  El acumulador de importes = Numero de apuestas por el precio de la apuesta
                    cResult = cResult + (vVar(j) * vImportes(j - 1))
                Next j
            End If
        Next i
    Else
        Select Case Me.CategoriaPremio
            Case Especial, Primera: cResult = vImportes(0)
            Case Segunda: cResult = vImportes(1)
            Case Tercera: cResult = vImportes(2)
            Case Cuarta: cResult = vImportes(3)
            Case Quinta: cResult = vImportes(4)
        End Select
    End If
        
    GetPremioEsperado = cResult

   On Error GoTo 0
   Exit Function

GetPremioEsperado_Error:
     Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
   ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
   '   Audita el error
   Call HandleException(ErrNumber, ErrDescription, "Premio.GetPremioEsperado")
   '   Lanza el Error
   Err.Raise ErrNumber, ErrSource, ErrDescription
End Function

'---------------------------------------------------------------------------------------
' Procedure : CategoriaPremio
' Author    : Charly
' Date      : 24/11/2013
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get CategoriaPremio() As CategoriaPremio

  On Error GoTo CategoriaPremio_Error
    m_eCategoriaPremio = Ninguna
    '
    '   Seleccionar modalidad del juego
    '
    Select Case ModalidadJuego
        Case Bonoloto
            Select Case BolasAcertadas
                Case 7:
                    m_eCategoriaPremio = Primera
                    NumerosAcertados = SeisAciertos
                
                Case 6:
                    If ComplementarioAcertado Then
                        m_eCategoriaPremio = Segunda
                        NumerosAcertados = CincoAciertosMasC
                    Else
                        m_eCategoriaPremio = Primera
                        NumerosAcertados = SeisAciertos
                    End If

                   
                Case 5:
                    If ComplementarioAcertado Then
                        m_eCategoriaPremio = Cuarta
                        NumerosAcertados = CuatroAciertos
                    Else
                        m_eCategoriaPremio = Tercera
                        NumerosAcertados = CincoAciertos
                    End If
                
                Case 4:
                    If ComplementarioAcertado Then
                        m_eCategoriaPremio = Quinta
                        NumerosAcertados = TresAciertos
                    Else
                        m_eCategoriaPremio = Cuarta
                        NumerosAcertados = CuatroAciertos
                    End If
                
                Case 3:
                    If ComplementarioAcertado Then
                        m_eCategoriaPremio = Ninguna
                        NumerosAcertados = TresAciertos
                    Else
                        m_eCategoriaPremio = Quinta
                        NumerosAcertados = TresAciertos
                    End If
            End Select
            
        Case LoteriaPrimitiva
            Select Case BolasAcertadas
                Case 7:
                    If ComplementarioAcertado Then
                        m_eCategoriaPremio = Especial
                        NumerosAcertados = SeisAciertosMasC
                    Else
                        m_eCategoriaPremio = Primera
                        NumerosAcertados = SeisAciertos
                    End If
                    
                Case 6:
                    If ComplementarioAcertado Then
                        m_eCategoriaPremio = Segunda
                        NumerosAcertados = CincoAciertosMasC
                    Else
                        m_eCategoriaPremio = Primera
                        NumerosAcertados = SeisAciertos
                    End If

                   
                Case 5:
                    If ComplementarioAcertado Then
                        m_eCategoriaPremio = Cuarta
                        NumerosAcertados = CuatroAciertos
                    Else
                        m_eCategoriaPremio = Tercera
                        NumerosAcertados = CincoAciertos
                    End If
                
                Case 4:
                    If ComplementarioAcertado Then
                        m_eCategoriaPremio = Quinta
                        NumerosAcertados = TresAciertos
                    Else
                        m_eCategoriaPremio = Cuarta
                        NumerosAcertados = CuatroAciertos
                    End If
                
                Case 3:
                    If ComplementarioAcertado Then
                        m_eCategoriaPremio = Ninguna
                        NumerosAcertados = TresAciertos
                    Else
                        m_eCategoriaPremio = Quinta
                        NumerosAcertados = TresAciertos
                    End If
            End Select

        Case gordoPrimitiva
            Select Case BolasAcertadas
                Case 5:
                    If ClaveAcertada Then
                        m_eCategoriaPremio = Primera
                    Else
                        m_eCategoriaPremio = Segunda
                    End If
                Case 4:
                    If ClaveAcertada Then
                        m_eCategoriaPremio = Tercera
                    Else
                        m_eCategoriaPremio = Cuarta
                    End If
                Case 3:
                    If ClaveAcertada Then
                        m_eCategoriaPremio = Quinta
                    Else
                        m_eCategoriaPremio = Sexta
                    End If
                Case 2:
                    If ClaveAcertada Then
                        m_eCategoriaPremio = Septima
                    Else
                        m_eCategoriaPremio = Octava
                    End If
            End Select
            
        Case Euromillones
            Select Case BolasAcertadas
            Case 5:
                Select Case NumeroEstrellasAcertadas
                    Case 2: m_eCategoriaPremio = Primera
                    Case 1: m_eCategoriaPremio = Segunda
                    Case Else: m_eCategoriaPremio = Tercera
                End Select
            Case 4:
                Select Case NumeroEstrellasAcertadas
                    Case 2: m_eCategoriaPremio = Cuarta
                    Case 1: m_eCategoriaPremio = Quinta
                    Case Else: m_eCategoriaPremio = Sexta
                End Select
            Case 3:
                Select Case NumeroEstrellasAcertadas
                    Case 2: m_eCategoriaPremio = Septima
                    Case 1: m_eCategoriaPremio = Novena
                    Case Else: m_eCategoriaPremio = Decima
                End Select
            Case 2:
                Select Case NumeroEstrellasAcertadas
                    Case 2: m_eCategoriaPremio = Octava
                    Case 1: m_eCategoriaPremio = Duodecima
                    Case Else: m_eCategoriaPremio = Trigesimotercera
                End Select
            Case 1:
                Select Case NumeroEstrellasAcertadas
                    Case 2: m_eCategoriaPremio = Undecima
                End Select
            End Select
        Case Else
            m_eCategoriaPremio = Ninguna
    End Select
   
    CategoriaPremio = m_eCategoriaPremio

   On Error GoTo 0
   Exit Property

CategoriaPremio_Error:
     Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
   ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
   '   Audita el error
   Call HandleException(ErrNumber, ErrDescription, "Premio.CategoriaPremio")
   '   Lanza el Error
   Err.Raise ErrNumber, ErrSource, ErrDescription

End Property

' *============================================================================*
' *     Procedure  : Class_Initialize
' *     Fichero    : Premio
' *     Autor      : Carlos Almela Baeza
' *     Creacion   : dom, 12/04/2009
' *     Asunto     :
' *============================================================================*
'
Private Sub Class_Initialize()
    '
    '   Importes médios de la Bonoloto para las categorias 1ª, 2ª, 3ª, 4ª y 5ª
    '
    ImportesPremiosBL = Array(450000, 100000, 1000, 30, 4)
    '
    '   Importes médios de la Primitiva para las categorias 1ª, 2ª, 3ª, 4ª y 5ª
    '
    ImportesPremiosLP = Array(1000000, 75000, 3000, 85, 8)
    BolasAcertadas = 0
    ModalidadJuego = Bonoloto
    ComplementarioAcertado = False
    ClaveAcertada = False
    NumeroEstrellasAcertadas = 0
    Pronosticos = 0
    key = "000000"
End Sub

' *============================================================================*
' *     Procedure  : CargaTablaMultiples
' *     Fichero    : Premio
' *     Autor      : Carlos Almela Baeza
' *     Creacion   : dom, 12/04/2009
' *     Asunto     :
' *============================================================================*
'
Private Sub CargaTablaMultiples()
    Dim vVar As Variant
    
    Erase TablaMultiples
    
    Select Case Pronosticos
        Case 5:
            vVar = Array(CincoAciertos, 1, 1, 42, 0, 0): TablaMultiples(0) = vVar
            vVar = Array(CuatroAciertosMasC, 0, 2, 0, 42, 0): TablaMultiples(1) = vVar
            vVar = Array(CuatroAciertos, 0, 0, 2, 42, 0): TablaMultiples(2) = vVar
            vVar = Array(TresAciertos, 0, 0, 0, 3, 41): TablaMultiples(3) = vVar
            vVar = Array(DosAciertos, 0, 0, 0, 0, 4): TablaMultiples(4) = vVar
            vVar = Array(Ninguna, 0, 0, 0, 0, 0): TablaMultiples(5) = vVar
        
        Case 6:
            vVar = Array(SeisAciertosMasC, 1, 0, 0, 0, 0): TablaMultiples(0) = vVar
            vVar = Array(SeisAciertos, 1, 0, 0, 0, 0): TablaMultiples(1) = vVar
            vVar = Array(CincoAciertosMasC, 0, 1, 0, 0, 0): TablaMultiples(2) = vVar
            vVar = Array(CincoAciertos, 0, 0, 1, 0, 0): TablaMultiples(3) = vVar
            vVar = Array(CuatroAciertos, 0, 0, 0, 1, 0): TablaMultiples(4) = vVar
            vVar = Array(TresAciertos, 0, 0, 0, 0, 1): TablaMultiples(5) = vVar
        
        Case 7:
            vVar = Array(SeisAciertosMasC, 1, 6, 0, 0, 0): TablaMultiples(0) = vVar
            vVar = Array(SeisAciertos, 1, 0, 6, 0, 0): TablaMultiples(1) = vVar
            vVar = Array(CincoAciertosMasC, 0, 1, 1, 5, 0): TablaMultiples(2) = vVar
            vVar = Array(CincoAciertos, 0, 0, 2, 5, 0): TablaMultiples(3) = vVar
            vVar = Array(CuatroAciertos, 0, 0, 0, 3, 4): TablaMultiples(4) = vVar
            vVar = Array(TresAciertos, 0, 0, 0, 0, 4): TablaMultiples(5) = vVar
        
        Case 8:
            vVar = Array(SeisAciertosMasC, 1, 6, 6, 15, 0): TablaMultiples(0) = vVar
            vVar = Array(SeisAciertos, 1, 0, 12, 15, 0): TablaMultiples(1) = vVar
            vVar = Array(CincoAciertosMasC, 0, 1, 2, 15, 10): TablaMultiples(2) = vVar
            vVar = Array(CincoAciertos, 0, 0, 3, 15, 10): TablaMultiples(3) = vVar
            vVar = Array(CuatroAciertos, 0, 0, 0, 6, 16): TablaMultiples(4) = vVar
            vVar = Array(TresAciertos, 0, 0, 0, 0, 10): TablaMultiples(5) = vVar
            
        Case 9:
            vVar = Array(SeisAciertosMasC, 1, 6, 12, 45, 20): TablaMultiples(0) = vVar
            vVar = Array(SeisAciertos, 1, 0, 18, 45, 20): TablaMultiples(1) = vVar
            vVar = Array(CincoAciertosMasC, 0, 1, 3, 30, 40): TablaMultiples(2) = vVar
            vVar = Array(CincoAciertos, 0, 0, 4, 30, 40): TablaMultiples(3) = vVar
            vVar = Array(CuatroAciertos, 0, 0, 0, 10, 40): TablaMultiples(4) = vVar
            vVar = Array(TresAciertos, 0, 0, 0, 0, 20): TablaMultiples(5) = vVar
           
        Case 10:
            vVar = Array(SeisAciertosMasC, 1, 6, 18, 90, 80): TablaMultiples(0) = vVar
            vVar = Array(SeisAciertos, 1, 0, 24, 90, 80): TablaMultiples(1) = vVar
            vVar = Array(CincoAciertosMasC, 0, 1, 4, 50, 100): TablaMultiples(2) = vVar
            vVar = Array(CincoAciertos, 0, 0, 5, 50, 100): TablaMultiples(3) = vVar
            vVar = Array(CuatroAciertos, 0, 0, 0, 15, 80): TablaMultiples(4) = vVar
            vVar = Array(TresAciertos, 0, 0, 0, 0, 35): TablaMultiples(5) = vVar
                                                                                      
                                                                                      
        Case 11:
            vVar = Array(SeisAciertosMasC, 1, 6, 24, 150, 200): TablaMultiples(0) = vVar
            vVar = Array(SeisAciertos, 1, 0, 30, 150, 200): TablaMultiples(1) = vVar
            vVar = Array(CincoAciertosMasC, 0, 1, 5, 75, 200): TablaMultiples(2) = vVar
            vVar = Array(CincoAciertos, 0, 0, 6, 75, 200): TablaMultiples(3) = vVar
            vVar = Array(CuatroAciertos, 0, 0, 0, 21, 140): TablaMultiples(4) = vVar
            vVar = Array(TresAciertos, 0, 0, 0, 0, 56): TablaMultiples(5) = vVar
            
        Case Else
            vVar = Array(SeisAciertosMasC, 0, 0, 0, 0, 0): TablaMultiples(0) = vVar
            vVar = Array(SeisAciertos, 0, 0, 0, 0, 0): TablaMultiples(1) = vVar
            vVar = Array(CincoAciertosMasC, 0, 0, 0, 0, 0): TablaMultiples(2) = vVar
            vVar = Array(CincoAciertos, 0, 0, 0, 0, 0): TablaMultiples(3) = vVar
            vVar = Array(CuatroAciertos, 0, 0, 0, 0, 0): TablaMultiples(4) = vVar
            vVar = Array(TresAciertos, 0, 0, 0, 0, 0): TablaMultiples(5) = vVar
    End Select
End Sub


'---------------------------------------------------------------------------------------
' Procedure : LiteralCategoriaPremio
' Author    : CHARLY
' Date      : 09/03/2014
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get LiteralCategoriaPremio() As String
    Select Case m_eCategoriaPremio
        Case Especial: m_sCategoriaPremio = " Ex"
        Case Primera: m_sCategoriaPremio = " 1ª"
        Case Segunda: m_sCategoriaPremio = " 2ª"
        Case Tercera: m_sCategoriaPremio = " 3ª"
        Case Cuarta: m_sCategoriaPremio = " 4ª"
        Case Quinta: m_sCategoriaPremio = " 5ª"
        Case Sexta: m_sCategoriaPremio = " 6ª"
        Case Septima: m_sCategoriaPremio = " 7ª"
        Case Octava: m_sCategoriaPremio = " 8ª"
        Case Novena: m_sCategoriaPremio = " 9ª"
        Case Decima: m_sCategoriaPremio = "10ª"
        Case Undecima: m_sCategoriaPremio = "11ª"
        Case Duodecima: m_sCategoriaPremio = "12ª"
        Case Trigesimotercera: m_sCategoriaPremio = "13ª"
    End Select
    
    LiteralCategoriaPremio = m_sCategoriaPremio

End Property

'---------------------------------------------------------------------------------------
' Procedure : NumerosAcertados
' Author    : CHARLY
' Date      : 09/03/2014
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get NumerosAcertados() As TiposAciertos

   On Error GoTo NumerosAcertados_Error

    Select Case ModalidadJuego
        Case LoteriaPrimitiva, Bonoloto
            Select Case BolasAcertadas
                Case 7:
                    m_eNumerosAcertados = SeisAciertosMasC
                    m_eNumerosAcertados = SeisAciertos
                    
                Case 6:
                    If ComplementarioAcertado Then
                        m_eNumerosAcertados = CincoAciertosMasC
                    Else
                        m_eNumerosAcertados = SeisAciertos
                    End If

                   
                Case 5:
                    If ComplementarioAcertado Then
                        m_eNumerosAcertados = CuatroAciertosMasC
                        m_eNumerosAcertados = CuatroAciertos
                    Else
                        m_eNumerosAcertados = CincoAciertos
                    End If
                
                Case 4:
                    If ComplementarioAcertado Then
                        m_eNumerosAcertados = TresAciertos
                    Else
                        m_eNumerosAcertados = CuatroAciertos
                    End If
                
                Case 3:
                    If ComplementarioAcertado Then
                        m_eNumerosAcertados = DosAciertos
                    Else
                        m_eNumerosAcertados = TresAciertos
                    End If
                Case 2: m_eNumerosAcertados = DosAciertos
                Case 1: m_eNumerosAcertados = UnAcierto
                Case Else: m_eNumerosAcertados = SinAciertos
            End Select
      End Select
    NumerosAcertados = m_eNumerosAcertados

   On Error GoTo 0
   Exit Property

NumerosAcertados_Error:
   Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
   ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
   '   Audita el error
   Call HandleException(ErrNumber, ErrDescription, "Premio.NumerosAcertados")
   '   Lanza el Error
   Err.Raise ErrNumber, ErrSource, ErrDescription
End Property

'---------------------------------------------------------------------------------------
' Procedure : NumerosAcertados
' Author    : CHARLY
' Date      : 09/03/2014
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Let NumerosAcertados(ByVal eNumerosAcertados As TiposAciertos)

   On Error GoTo NumerosAcertados_Error

    m_eNumerosAcertados = eNumerosAcertados

   On Error GoTo 0
   Exit Property

NumerosAcertados_Error:
   Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
   ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
   '   Audita el error
   Call HandleException(ErrNumber, ErrDescription, "Premio.NumerosAcertados")
   '   Lanza el Error
   Err.Raise ErrNumber, ErrSource, ErrDescription
End Property
