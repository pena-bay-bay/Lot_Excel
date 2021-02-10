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
' *     Fichero    : Premio.cls
' *
' *     Autor      : Carlos Almela Baeza
' *     Creación   : do., 06/abr/2014 00:00:00
' *     Modificado : ma., 14/jul/2020 18:49:45
' *     Versión    : 1.1
' *     Propósito  : Representa la clase datos de un premio de un sorteo
' *
' *============================================================================*
Option Explicit
Option Base 0
'---------------------------------------------------------------------------------------
' Module    : Premio
' Author    : CHARLY
' Date      : 06/04/2014
' Purpose   :
'---------------------------------------------------------------------------------------

'
' Propiedades Públicas
'
Public Id As Integer                        ' Identificador unico en la colección
Public CategoriaPremio As CategoriaPremio   ' Codigo de caegoria de premio
Public NumeroAcertantesEspaña As Long       ' Numero de acertantes
Public NumeroAcertantesEuropa As Long       ' Numero de acertantes en europa
Public Importe As Currency                  ' Importe del premio
Public Juego As Juego                       ' Concurso
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
    '
    '   Segun el juego interpretamos la categoría
    '
    Select Case Me.Juego
        Case Bonoloto
            Select Case Me.CategoriaPremio
                Case Primera: m_sCategoriaTexto = "1ª (6 Aciertos)"
                Case Segunda: m_sCategoriaTexto = "2ª (5 Aciertos + C)"
                Case Tercera: m_sCategoriaTexto = "3ª (5 Aciertos)"
                Case Cuarta: m_sCategoriaTexto = "4ª (4 Aciertos)"
                Case Quinta: m_sCategoriaTexto = "5ª (3 Aciertos)"
                Case Reintegro: m_sCategoriaTexto = "Reintegro"
                Case Else: m_sCategoriaTexto = "Ninguna"
            End Select
        
        Case LoteriaPrimitiva
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
         
        Case GordoPrimitiva
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
   Call HandleException(ErrNumber, ErrDescription, "Premio.CategoriaTexto")
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
    Me.CategoriaPremio = Ninguna
    Me.Importe = 0
    Me.NumeroAcertantesEuropa = 0
    Me.NumeroAcertantesEspaña = 0
    Me.Juego = JUEGO_DEFECTO
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
   Call HandleException(ErrNumber, ErrDescription, "Premio.EsValido")
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
        If Juego = Euromillones Then
            sResult = "Juego: $0, Categoria: $1 = $2, Importe: $3 Euros, Acertantes: $4 Esp y $5 Eur"
        Else
            sResult = "Juego: $0, Categoria: $1 = $2, Importe: $3 Euros, Acertantes: $4"
        End If
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
        sResult = Replace(sResult, "$3", Str(Me.Importe))
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
   Call HandleException(ErrNumber, ErrDescription, "Premio.ToString")
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
    Dim j As Integer
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
            j = Len(mTexto) - InStr(1, mTexto, ":", 1)
            Me.CategoriaPremio = Val(Right(mTexto, j))
        
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
            If Juego = Euromillones Then
                mTexto = vDat(4)
                Me.NumeroAcertantesEuropa = Val(mTexto)
            End If
        End Select
    Next i
    
  On Error GoTo 0
   Exit Sub
Parse_Error:
   Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
   ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
   Call HandleException(ErrNumber, ErrDescription, "Premio.Parse")
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
            eResult = i + 1
            Exit For
        End If
    Next i
    getJuego = eResult

  On Error GoTo 0
   Exit Function
getJuego_Error:
   Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
   ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
   Call HandleException(ErrNumber, ErrDescription, "Premio.getJuego")
   Err.Raise ErrNumber, ErrSource, ErrDescription
End Function



'------------------------------------------------------------------------------*
' Función        : Pack
' Fecha          : ma., 14/jul/2020 19:11:40
' Propósito      : Estructura la información para guardar en una cadena
' Retorno        : devuelve una cadena según el juego
'                  "Categoria, acertantes, importe, acertantes Europa"
'------------------------------------------------------------------------------*
'
Public Function Pack() As String
    Dim mStr As String
    
  On Error GoTo Pack_Error
    '
    '   Segun el juego seleccionamos un patrón
    '
    Select Case Juego
        Case Bonoloto:
            mStr = "$1,$2,$3"
        Case LoteriaPrimitiva:
            mStr = "$1,$2,$3"
        Case GordoPrimitiva:
            mStr = "$1,$2,$3"
        Case Euromillones:
            mStr = "$1,$2,$3,$4"
    End Select
    mStr = Replace(mStr, "$1", CStr(CategoriaPremio))
    mStr = Replace(mStr, "$2", CStr(NumeroAcertantesEspaña))
    mStr = Replace(mStr, "$4", CStr(NumeroAcertantesEuropa))
    mStr = Replace(mStr, "$3", Str(Importe))

    Pack = mStr
  On Error GoTo 0
   Exit Function
Pack_Error:
   Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
   ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
   Call HandleException(ErrNumber, ErrDescription, "Premio.Pack")
   Err.Raise ErrNumber, ErrSource, ErrDescription
End Function



'------------------------------------------------------------------------------*
' Función        : UnPack
' Fecha          : ma., 14/jul/2020 19:14:15
' Propósito      : Realiza un parser de la información
'------------------------------------------------------------------------------*
'
Public Sub UnPack(vNewValue As String)
    Dim vMatriz As Variant
    
  On Error GoTo UnPack_Error
    If Len(vNewValue) = 0 Then
        Exit Sub
    End If
    '
    '   Desempaquetamos la estructura
    '
    vMatriz = Split(vNewValue, ",")
    '
    '
    Select Case Juego
        Case Bonoloto, LoteriaPrimitiva, GordoPrimitiva:
            Me.CategoriaPremio = Val(vMatriz(0))
            Me.NumeroAcertantesEspaña = Val(vMatriz(1))
            Me.Importe = Val(vMatriz(2))
            
        Case Euromillones:
            Me.CategoriaPremio = Val(vMatriz(0))
            Me.NumeroAcertantesEspaña = Val(vMatriz(1))
            Me.Importe = Val(vMatriz(2))
            Me.NumeroAcertantesEuropa = Val(vMatriz(3))
    End Select
  
  On Error GoTo 0
   Exit Sub
UnPack_Error:
   Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
   ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
   Call HandleException(ErrNumber, ErrDescription, "Premio.UnPack")
   Err.Raise ErrNumber, ErrSource, ErrDescription
End Sub

