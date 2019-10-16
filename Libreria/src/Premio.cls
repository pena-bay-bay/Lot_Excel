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
' *     Creaci�n   : Sab, 26/ene/2019 11:57
' *     Versi�n    : 1.0
' *     Prop�sito  : Define un premio del sorteo
' *
' *============================================================================*
Option Explicit
Option Base 0
'
'--- Variables Privadas -------------------------------------------------------*
Private m_sCategoriaTexto   As String
Private m_error             As Integer
Private mDatoDefecto        As Boolean
Private mEnt                As EntidadNegocio
'--- Propiedades --------------------------------------------------------------*
'------------------------------------------------------------------------------*
' Procedimiento  : Categoria
' Fecha          : dd/MMM/yyyy
' Prop�sito      : Define la categor�a del premio: primera, segunda,etc.
'------------------------------------------------------------------------------*
Public Categoria As CategoriaPremio
'------------------------------------------------------------------------------*
' Procedimiento  : NumeroAcertantesEspa�a
' Fecha          : 27/01/2019
' Prop�sito      : Numero de acertantes en Espa�a
'------------------------------------------------------------------------------*
Public NumeroAcertantesEspa�a As Integer
'------------------------------------------------------------------------------*
' Procedimiento  : NumeroAcertantesEuropa
' Fecha          : 27/01/2019
' Prop�sito      : Numero de acertantes en Europa (solo Euromillones)
'------------------------------------------------------------------------------*
Public NumeroAcertantesEuropa As Integer
'------------------------------------------------------------------------------*
' Procedimiento  : Importe
' Fecha          : 27/01/2019
' Prop�sito      : Importe del premio para una apuesta acertada
'------------------------------------------------------------------------------*
Public Importe As Currency
'------------------------------------------------------------------------------*
' Procedimiento  : Juego
' Fecha          : 27/01/2019
' Prop�sito      : Juego al que pertenece el premio
'------------------------------------------------------------------------------*
Public Juego As Juego
'------------------------------------------------------------------------------*
' Procedimiento  : ImporteDefault
' Fecha          : 27/01/2019
' Prop�sito      : Si se modifica por el usuario es un importe real del premio,
'                  por defecto se asignan un importe medio de premios deducido
'                  por la estad�stica
'------------------------------------------------------------------------------*
Public Property Get ImporteDefault() As Boolean
    ImporteDefault = mDatoDefecto
End Property

Public Property Let ImporteDefault(ByVal vNewValue As Boolean)
    mDatoDefecto = vNewValue
End Property
'------------------------------------------------------------------------------*
' Procedimiento  : MarkForDelete
' Fecha          : 27/01/2019
' Prop�sito      :
'------------------------------------------------------------------------------*
Public Property Get MarkForDelete() As Boolean
    MarkForDelete = mEnt.MarkForDelete
End Property
Public Property Let MarkForDelete(ByVal vNewValue As Boolean)
    mEnt.MarkForDelete = vNewValue
End Property

'---------------------------------------------------------------------------------------
' Procedure : CategoriaTexto
' Author    : CHARLY
' Date      : 06/04/2014
' Purpose   : Aplicando el n�mero de bolas acertadas, determinar la categor�a del premio
'---------------------------------------------------------------------------------------
'
Public Property Get CategoriaTexto() As String
Attribute CategoriaTexto.VB_Description = "Literal de la categor�a del Juego"

   On Error GoTo CategoriaTexto_Error
    Select Case Me.Juego
        Case bonoloto
            Select Case Me.Categoria
                Case Primera:   m_sCategoriaTexto = "1� (6 Aciertos)"
                Case Segunda:   m_sCategoriaTexto = "2� (5 Aciertos + C)"
                Case Tercera:   m_sCategoriaTexto = "3� (5 Aciertos)"
                Case Cuarta:    m_sCategoriaTexto = "4� (4 Aciertos)"
                Case Quinta:    m_sCategoriaTexto = "5� (3 Aciertos)"
                Case Reintegro: m_sCategoriaTexto = "Reintegro"
                Case Else: m_sCategoriaTexto = "Ninguna"
            End Select
        
        Case LoteriaPrimitiva
            Select Case Me.Categoria
                Case Especial:  m_sCategoriaTexto = "Especial (6 Aciertos + R)"
                Case Primera:   m_sCategoriaTexto = "1� (6 Aciertos)"
                Case Segunda:   m_sCategoriaTexto = "2� (5 Aciertos + C)"
                Case Tercera:   m_sCategoriaTexto = "3� (5 Aciertos)"
                Case Cuarta:    m_sCategoriaTexto = "4� (4 Aciertos)"
                Case Quinta:    m_sCategoriaTexto = "5� (3 Aciertos)"
                Case Reintegro: m_sCategoriaTexto = "Reintegro"
                Case Else: m_sCategoriaTexto = "Ninguna"
            End Select
            
        Case Euromillones
            Select Case Me.Categoria
                Case Primera:   m_sCategoriaTexto = "1� (5 + 2)"
                Case Segunda:   m_sCategoriaTexto = "2� (5 + 1)"
                Case Tercera:   m_sCategoriaTexto = "3� (5 + 0)"
                Case Cuarta:    m_sCategoriaTexto = "4� (4 + 2)"
                Case Quinta:    m_sCategoriaTexto = "5� (4 + 1)"
                Case sexta:     m_sCategoriaTexto = "6� (4 + 0)"
                Case Septima:   m_sCategoriaTexto = "7� (3 + 2)"
                Case Octava:    m_sCategoriaTexto = "8� (2 + 2)"
                Case Novena:    m_sCategoriaTexto = "9� (3 + 1)"
                Case Decima:    m_sCategoriaTexto = "10� (3 + 0)"
                Case Undecima:  m_sCategoriaTexto = "11� (1 + 2)"
                Case duodecima: m_sCategoriaTexto = "12� (2 + 1)"
                Case Trigesimotercera: m_sCategoriaTexto = "13� (2 + 0)"
                Case Else: m_sCategoriaTexto = "Ninguna"
            End Select
         
        Case gordoPrimitiva
            Select Case Me.Categoria
                Case Primera: m_sCategoriaTexto = "1� (5 + 1)"
                Case Segunda: m_sCategoriaTexto = "2� (5 + 0)"
                Case Tercera: m_sCategoriaTexto = "3� (4 + 1)"
                Case Cuarta: m_sCategoriaTexto = "4� (4 + 0)"
                Case Quinta: m_sCategoriaTexto = "5� (3 + 1)"
                Case sexta: m_sCategoriaTexto = "6� (3 + 0)"
                Case Septima: m_sCategoriaTexto = "7� (2 + 1)"
                Case Octava: m_sCategoriaTexto = "8� (2 + 0)"
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
   Call HandleException(ErrNumber, ErrDescription, "Premio.CategoriaTexto")
   '   Lanza el Error
   Err.Raise ErrNumber, ErrSource, ErrDescription
End Property

'--- Constantes ---------------------------------------------------------------*
'--- Mensajes -----------------------------------------------------------------*
'--- Errores ------------------------------------------------------------------*
'--- Metodos Privados ---------------------------------------------------------*
Private Sub Class_Initialize()
    Categoria = Ninguna
    Importe = 0
    NumeroAcertantesEuropa = 0
    NumeroAcertantesEspa�a = 0
    mDatoDefecto = True
    Set mEnt = New EntidadNegocio
End Sub

'--- Metodos Publicos ---------------------------------------------------------*
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
    
    If Me.Categoria = Ninguna Then
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
   Call HandleException(ErrNumber, ErrDescription, "Premio.EsValido")
   '   Lanza el Error
   Err.Raise ErrNumber, ErrSource, ErrDescription
End Function

'---------------------------------------------------------------------------------------
' Procedure : ToString
' Author    : CHARLY
' Date      : 06/04/2014
' Purpose   : Formatea en texto las propiedades del Premio
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
        sResult = Replace(sResult, "$1", Me.Categoria)
        sResult = Replace(sResult, "$2", Me.CategoriaTexto)
        '
        ' Importe
        '
        sResult = Replace(sResult, "$3", Me.Importe)
        sResult = Replace(sResult, "$4", Me.NumeroAcertantesEspa�a)
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
   Call HandleException(ErrNumber, ErrDescription, "Premio.ToString")
   '   Lanza el Error
   Err.Raise ErrNumber, ErrSource, ErrDescription
End Function


'---------------------------------------------------------------------------------------
' Procedure : Parse
' Author    : CHARLY
' Date      : 18/05/2014
' Purpose   : Analiza el texto y rellena las propiedades del objeto
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
            Me.Categoria = Val(Right(Trim(mTexto), 1))
        
        'Importe: $3 Euros
        Case Is = 2
            vDat = Split(Trim(vMatriz(i)), " ")
            mTexto = vDat(1)
            Me.Importe = Val(mTexto)
            Me.ImporteDefault = False
            
        'Acertantes: $4 y $5
        Case Is = 3
            vDat = Split(Trim(vMatriz(i)), " ")
            mTexto = vDat(1)
            Me.NumeroAcertantesEspa�a = Val(mTexto)
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
   Call HandleException(ErrNumber, ErrDescription, "Premio.Parse")
   '   Lanza el Error
   Err.Raise ErrNumber, ErrSource, ErrDescription

End Sub
'------------------------------------------------------------------------------*
' Procedimiento  : GetImportePremio
' Fecha          : 27/01/2019
' Prop�sito      : Devuelve el importe de un numero definido de apuestas
' Par�metros     : datApuestas numero de apuestas a valorar
' Retorno        : Importe total
'------------------------------------------------------------------------------*
'
Public Function GetImportePremio(datApuestas As Integer) As Currency
    GetImportePremio = datApuestas * Importe
End Function

'---------------------------------------------------------------------------------------
' Procedure : getJuego
' Author    : CHARLY
' Date      : 18/05/2014
' Purpose   : Devuelve el nombre del juego al que pertenece el premio
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
   '   Audita el error
   Call HandleException(ErrNumber, ErrDescription, "Premio.getJuego")
   '   Lanza el Error
   Err.Raise ErrNumber, ErrSource, ErrDescription

End Function

'' *===========(EOF): Premio.cls








