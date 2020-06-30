VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Premio2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'---------------------------------------------------------------------------------------
' Module    : Premio2
' Author    : CHARLY
' Date      : 06/04/2014
' Purpose   :
'---------------------------------------------------------------------------------------

Option Explicit
Option Base 0
'
' Propiedades Públicas
'
Public CategoriaPremio As CategoriaPremio
Public NumeroAcertantesEspaña As Integer
Public NumeroAcertantesEuropa As Integer
Public Importe As Currency
Public Juego As Juego
'
'  Variables privadas
'
Private m_sCategoriaTexto As String
Private m_error As Integer

'---------------------------------------------------------------------------------------
' Procedure : CategoriaTexto
' Author    : CHARLY
' Date      : 06/04/2014
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get CategoriaTexto() As String

   On Error GoTo CategoriaTexto_Error
    Select Case Me.Juego
        Case bonoloto
            Select Case Me.CategoriaPremio
                Case Primera: m_sCategoriaTexto = "1ª (6 Aciertos)"
                Case Segunda: m_sCategoriaTexto = "2ª (5 Aciertos + C)"
                Case Tercera: m_sCategoriaTexto = "3ª (5 Aciertos)"
                Case Cuarta: m_sCategoriaTexto = "4ª (4 Aciertos)"
                Case Quinta: m_sCategoriaTexto = "5ª (3 Aciertos)"
                Case Reintegro: m_sCategoriaTexto = "Reintegro"
                Case Else: m_sCategoriaTexto = "Ninguna"
            End Select
        
        Case loteriaPrimitiva
            Select Case Me.CategoriaPremio
                Case Especial: m_sCategoriaTexto = "Especial (6 Aciertos + R)"
                Case Primera: m_sCategoriaTexto = "1ª (6 Aciertos)"
                Case Segunda: m_sCategoriaTexto = "2ª (5 Aciertos + C)"
                Case Tercera: m_sCategoriaTexto = "3ª (5 Aciertos)"
                Case Cuarta: m_sCategoriaTexto = "4ª (4 Aciertos)"
                Case Quinta: m_sCategoriaTexto = "5ª (3 Aciertos)"
                Case Reintegro: m_sCategoriaTexto = "Reintegro"
                Case Else: m_sCategoriaTexto = "Ninguna"
            End Select
            
        Case Euromillones
            Select Case Me.CategoriaPremio
                Case Primera: m_sCategoriaTexto = "1ª 5 + 2"
                Case Segunda: m_sCategoriaTexto = "2ª 5 + 1"
                Case Tercera: m_sCategoriaTexto = "3ª 5 + 0"
                Case Cuarta: m_sCategoriaTexto = "4ª 4 + 2"
                Case Quinta: m_sCategoriaTexto = "5ª 4 + 1"
                Case Sexta: m_sCategoriaTexto = "6ª 4 + 0"
                Case Septima: m_sCategoriaTexto = "7ª 3 + 2"
                Case Octava: m_sCategoriaTexto = "8ª 2 + 2"
                Case Novena: m_sCategoriaTexto = "9ª 3 + 1"
                Case Decima: m_sCategoriaTexto = "10ª 3 + 0"
                Case Undecima: m_sCategoriaTexto = "11ª 1 + 2"
                Case Duodecima: m_sCategoriaTexto = "12ª 2 + 1"
                Case Trigesimotercera: m_sCategoriaTexto = "13ª 2 + 0"
                Case Else: m_sCategoriaTexto = "Ninguna"
            End Select
         
        Case gordoPrimitiva
            Select Case Me.CategoriaPremio
                Case Primera: m_sCategoriaTexto = "1ª (5 + 1)"
                Case Segunda: m_sCategoriaTexto = "2ª (5 + 0)"
                Case Tercera: m_sCategoriaTexto = "3ª (4 + 1)"
                Case Cuarta: m_sCategoriaTexto = "4ª (4 + 0)"
                Case Quinta: m_sCategoriaTexto = "5ª (3 + 1)"
                Case Sexta: m_sCategoriaTexto = "6ª (3 + 0)"
                Case Septima: m_sCategoriaTexto = "7ª (2 + 1)"
                Case Octava: m_sCategoriaTexto = "8ª (2 + 0)"
                Case Reintegro: m_sCategoriaTexto = "Reintegro"
                Case Else: m_sCategoriaTexto = "Ninguna"
            End Select
        
        Case Else
            m_sCategoriaTexto = "Ninguna"
    End Select
    
    CategoriaTexto = m_sCategoriaTexto

   On Error GoTo 0
   Exit Property

CategoriaTexto_Error:
   Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
   ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
   '   Audita el error
   Call HandleException(ErrNumber, ErrDescription, "Premio2.CategoriaTexto")
   '   Lanza el Error
   Err.Raise ErrNumber, ErrSource, ErrDescription
End Property

'---------------------------------------------------------------------------------------
' Procedure : Class_Initialize
' Author    : CHARLY
' Date      : 06/04/2014
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub Class_Initialize()

   On Error GoTo Class_Initialize_Error

    Me.CategoriaPremio = Ninguna
    Me.Importe = 0
    Me.NumeroAcertantesEuropa = 0
    Me.NumeroAcertantesEspaña = 0
   On Error GoTo 0
   Exit Sub

Class_Initialize_Error:
   Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
   ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
   '   Audita el error
   Call HandleException(ErrNumber, ErrDescription, "Premio2.Class_Initialize")
   '   Lanza el Error
   Err.Raise ErrNumber, ErrSource, ErrDescription
End Sub

'---------------------------------------------------------------------------------------
' Procedure : EsValido
' Author    : CHARLY
' Date      : 06/04/2014
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Function EsValido() As Boolean

    Dim bResult As Boolean

   On Error GoTo EsValido_Error
    m_error = 0
    
    If Me.CategoriaPremio = Ninguna Then
        m_error = m_error + 1
    End If
    
    If Me.Juego = 0 Then
        m_error = m_error + 2
    End If
    
    If m_error > 0 Then
        bResult = False
    Else
        bResult = True
    End If
    
    EsValido = bResult

   On Error GoTo 0
   Exit Function

EsValido_Error:
   Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
   ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
   '   Audita el error
   Call HandleException(ErrNumber, ErrDescription, "Premio2.EsValido")
   '   Lanza el Error
   Err.Raise ErrNumber, ErrSource, ErrDescription
End Function

'---------------------------------------------------------------------------------------
' Procedure : ToString
' Author    : CHARLY
' Date      : 06/04/2014
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Function ToString() As String
    Dim sResult As String
    Dim mVar As Variant
    Dim sTexto As String
   On Error GoTo ToString_Error
    If Me.EsValido Then
        sResult = "Juego: $0, Categoria: $1 = $2, Importe: $3 Euros, Acertantes: $4 y $5"
        '
        ' Juego
        '
        mVar = Split(NOMBRE_JUEGOS, ";")
        sTexto = mVar(Me.Juego - 1)
        sResult = Replace(sResult, "$0", sTexto)
        '
        ' Categoria
        '
        sResult = Replace(sResult, "$1", Me.CategoriaPremio)
        sResult = Replace(sResult, "$2", Me.CategoriaTexto)
        '
        ' Importe
        '
        sResult = Replace(sResult, "$3", Me.Importe)
        sResult = Replace(sResult, "$4", Me.NumeroAcertantesEspaña)
        sResult = Replace(sResult, "$5", Me.NumeroAcertantesEuropa)
    Else
        sResult = "#Premio no valido"
    End If
    
    ToString = sResult

   On Error GoTo 0
   Exit Function

ToString_Error:
   Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
   ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
   '   Audita el error
   Call HandleException(ErrNumber, ErrDescription, "Premio2.ToString")
   '   Lanza el Error
   Err.Raise ErrNumber, ErrSource, ErrDescription
End Function


'---------------------------------------------------------------------------------------
' Procedure : Parse
' Author    : CHARLY
' Date      : 18/05/2014
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Sub Parse(datTexto As String)

    Dim vMatriz As Variant
    Dim i As Integer
    Dim vDat As Variant
    Dim mTexto As String
   On Error GoTo Parse_Error

    If Len(datTexto) = 0 Then
        Exit Sub
    End If
    
    vMatriz = Split(datTexto, ",")

    For i = 0 To UBound(vMatriz)
        Select Case i
        ' Juego: $0
        Case Is = 0
            vDat = Split(vMatriz(i), ":")
            mTexto = vDat(1)
            Me.Juego = getJuego(mTexto)
            
        ' Categoria: $1 = $2
        Case Is = 1
            vDat = Split(vMatriz(i), "=")
            mTexto = vDat(0)
            Me.CategoriaPremio = Val(Right(mTexto, 1))
        
        'Importe: $3 Euros
        Case Is = 2
            vDat = Split(Trim(vMatriz(i)), " ")
            mTexto = vDat(1)
            Me.Importe = Val(mTexto)
            
        'Acertantes: $4 y $5
        Case Is = 3
            vDat = Split(Trim(vMatriz(i)), " ")
            mTexto = vDat(1)
            Me.NumeroAcertantesEspaña = Val(mTexto)
            mTexto = vDat(3)
            Me.NumeroAcertantesEuropa = Val(mTexto)
            
        End Select
    Next i
   On Error GoTo 0
   Exit Sub

Parse_Error:
   Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
   ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
   '   Audita el error
   Call HandleException(ErrNumber, ErrDescription, "Premio2.Parse")
   '   Lanza el Error
   Err.Raise ErrNumber, ErrSource, ErrDescription

End Sub

'---------------------------------------------------------------------------------------
' Procedure : getJuego
' Author    : CHARLY
' Date      : 18/05/2014
' Purpose   :
'---------------------------------------------------------------------------------------
'
Friend Function getJuego(datNombre As String) As Juego

    Dim eResult As Juego
    Dim mVar As Variant
    Dim i As Integer
   On Error GoTo getJuego_Error
    mVar = Split(NOMBRE_JUEGOS, ";")
    For i = 0 To UBound(mVar)
        If mVar(i) = Trim(datNombre) Then
            eResult = i
            Exit For
        End If
    Next i
    getJuego = eResult

   On Error GoTo 0
   Exit Function

getJuego_Error:
   Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
   ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
   '   Audita el error
   Call HandleException(ErrNumber, ErrDescription, "Premio2.getJuego")
   '   Lanza el Error
   Err.Raise ErrNumber, ErrSource, ErrDescription

End Function
