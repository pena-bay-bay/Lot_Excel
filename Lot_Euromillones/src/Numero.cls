VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Numero"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'/**
' * Clase   Numero
' * Created 11/01/2004 23:30
' * version 1.0.01
' **/

'//Propiedades privadas
Option Explicit
Private m_valor As Integer
Private m_orden As Integer
Private m_error As Integer
Private m_bEsPar As Boolean

Private Sub Class_Initialize()
    m_error = 0
    m_valor = 0
End Sub

'/**
' * Property    Valor
' *
' * @int        Valor del número, comprendido entre 1 y 49
' **/
Public Property Get Valor() As Variant
    Valor = m_valor
End Property

Public Property Let Valor(ByVal vNewValue As Variant)
    If vNewValue > 0 And vNewValue < 55 Then
        m_valor = vNewValue
    Else
        m_error = 1
    End If
End Property

'/**
' * method      peso
' *
' * @string     Devuelve la cadena alto si el Numero es igual o superior a 25 y bajo
' *             si el número es inferior
' **/
Public Property Get Peso() As String
    If (Not ver_valor) Then
        Peso = LT_ERROR
    Else
        If m_valor < 25 Then
            Peso = LT_BAJO
        Else
            Peso = LT_ALTO
        End If
    End If
End Property

'/**
' * method      Paridad
' *
' * @string     Devuelve la paridad del Numero: par, impar o #err
' **/
Public Property Get Paridad() As String
    Dim tmp As Integer
    If (Not ver_valor) Then
        Paridad = LT_ERROR
    Else
        tmp = m_valor Mod 2
        If tmp = 0 Then
            Paridad = LT_PAR
        Else
            Paridad = LT_IMPAR
        End If
    End If
End Property

'/**
' * method      Terminacion
' *
' * @int        Devuelve la terminación del Numero comprendida entre 0..9 ó #err
' **/
Public Property Get Terminacion() As Integer
    Dim tmp As Integer
    If (Not ver_valor) Then
        Terminacion = -1
    Else
        tmp = m_valor Mod 10
        Terminacion = tmp
    End If
End Property

'/**
' * method      decena
' *
' * @int        Devuelve la decena a la que pertenece siendo 1-> 0..9, 2->10..19, y así
' *             sucesivamente hasta 5 ó #err
' **/
Public Property Get Decena() As Integer
    Dim tmp As Integer
    If (Not ver_valor) Then
        Decena = -1
    Else
        tmp = Fix(m_valor / 10) + 1
        Decena = tmp
    End If
End Property

'/**
' * method      septena
' *
' * @int        Devuelve la septena a la que pertenece siendo 1-> 0..7, 2->8..14, y así
' *             sucesivamente hasta 5 ó #err
' **/
Public Property Get Septena() As Integer
    Dim tmp As Integer
    If (Not ver_valor) Then
        Septena = -1
    Else
        If (m_valor Mod 7) = 0 Then
            tmp = Fix(m_valor / 7)
        Else
            tmp = Fix(m_valor / 7) + 1
        End If
        Septena = tmp
    End If
End Property

'---------------------------------------------------------------------------------------
' Procedimiento : Orden
' Creación      : 05-dic-2006 19:59
' Autor         : Carlos Almela Baeza
' Objeto        : Lugar que ocupa el número en la extracción de un sorteo
'---------------------------------------------------------------------------------------
'
Public Property Get Orden() As Variant
    Orden = m_orden
End Property

Public Property Let Orden(ByVal vNewValue As Variant)
    m_orden = vNewValue
End Property

'/**
' * Property    error
' *
' * @int        código de error producido en la clase
' **/
Public Property Get Error() As Variant
    Error = m_error
End Property

''/**
'' * Property    mensaje_error
'' *
'' * @string     Cadena de texto con la explicación del error
'' **/
'Public Property Get Mensaje_error() As Variant
'    Select Case m_error
'     Case 0:    Mensaje_error = MSG_NOERROR
'     Case 1:    Mensaje_error = MSG_MALRANGO
'     Case 2:    Mensaje_error = MSG_FALTANumero
'    End Select
'End Property

'---------------------------------------------------------------------------------------
' Procedimiento : ver_valor
' Creación      : 05-dic-2006 19:59
' Autor         : Carlos Almela Baeza
' Objeto        : Comprueba el Numero introducido
'---------------------------------------------------------------------------------------
'
Private Function ver_valor()
    If (m_valor = 0) Then
        m_error = 2
        ver_valor = False
    Else
        ver_valor = True
    End If
End Function


'---------------------------------------------------------------------------------------
' Procedure : EsPar
' Author    : CHARLY
' Date      : 15/03/2014
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get EsPar() As Boolean
    Dim tmp As Integer
    tmp = m_valor Mod 2
    If tmp = 0 Then
        m_bEsPar = True
    Else
        m_bEsPar = False
    End If

    EsPar = m_bEsPar

End Property

'---------------------------------------------------------------------------------------
' Procedure : GetMensaje
' Author    : CHARLY
' Date      : 15/03/2014
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Function GetMensaje() As String

    Dim sResult As String


   On Error GoTo GetMensaje_Error
    Select Case m_error
     Case 0:    sResult = MSG_NOERROR
     Case 1:    sResult = MSG_MALRANGO
     Case 2:    sResult = MSG_FALTANumero
    End Select

    GetMensaje = sResult

   On Error GoTo 0
   Exit Function

GetMensaje_Error:
   Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
   ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
   '   Audita el error
   Call HandleException(ErrNumber, ErrDescription, "Numero.GetMensaje")
   '   Lanza el Error
   Err.Raise ErrNumber, ErrSource, ErrDescription
End Function

'---------------------------------------------------------------------------------------
' Procedure : EsValido
' Author    : CHARLY
' Date      : 15/03/2014
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Function EsValido(datJuego As Juego) As Boolean

    Dim bResult As Boolean

   On Error GoTo EsValido_Error
    
    bResult = True
    
    Select Case datJuego

        Case LoteriaPrimitiva:
            If (m_valor < 1) Or (m_valor > 49) Then
                bResult = False
                m_error = 1
            End If
        
        Case Bonoloto
            If (m_valor < 1) Or (m_valor > 49) Then
                bResult = False
                m_error = 1
            End If

        Case gordoPrimitiva
            If (m_valor < 1) Or (m_valor > 54) Then
                bResult = False
                m_error = 1
            End If

        Case Euromillones
            If (m_valor < 1) Or (m_valor > 50) Then
                bResult = False
                m_error = 1
            End If

        Case PrimitivaBonoloto
            If (m_valor < 1) Or (m_valor > 49) Then
                bResult = False
                m_error = 1
            End If

    End Select
    
    EsValido = bResult

   On Error GoTo 0
   Exit Function

EsValido_Error:
   Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
   ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
   '   Audita el error
   Call HandleException(ErrNumber, ErrDescription, "Numero.EsValido")
   '   Lanza el Error
   Err.Raise ErrNumber, ErrSource, ErrDescription
End Function
