VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Numero"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' *============================================================================*
' *
' *     Fichero    : Numero.cls
' *
' *     Autor      : Carlos Almela Baeza
' *     Creación   : 11/01/2004 23:30 hh:mm:ss
' *     Versión    : 1.0.01
' *     Propósito  : Contiene las propiedades de un número con arreglo a un
' *
' *============================================================================*
Option Explicit
Option Base 0
'
'--- Variables Privadas -------------------------------------------------------*
Private mValor As Integer               ' Valor asignado al numero
Private mOrden As Integer               ' Orden que ocupa en la combinación
Private mError As Integer               ' Error asociado al numero en función
                                        ' del juego
Private mJuego As Juego                 ' Juego al que pertenece el número
'--- Constantes ---------------------------------------------------------------*
'--- Errores ------------------------------------------------------------------*
Private Const ERR_NOVALIDJUEGO06_49 = 1
Private Const ERR_NOVALIDJUEGO05_54 = 2
Private Const ERR_NOVALIDJUEGO05_50 = 3
Private Const ERR_NUMERONOASIGNADO = 4
Private Const ERR_NOVALIDJUEGO02_12 = 5
'--- Mensajes -----------------------------------------------------------------*
Private Const MSG_NUMVALID = "Número válido"
Private Const MSG_NOVALIDJUEGO06_49 = "El rango del Numero debe estar comprendido entre 1 y 49, ambos inclusive."
Private Const MSG_NOVALIDJUEGO05_54 = "El rango del Numero debe estar comprendido entre 1 y 54, ambos inclusive"
Private Const MSG_NOVALIDJUEGO05_50 = "El rango del Numero debe estar comprendido entre 1 y 50, ambos inclusive"
Private Const MSG_NOVALIDJUEGO02_12 = "El rango del Numero debe estar comprendido entre 1 y 12, ambos inclusive"
Private Const MSG_NUMERONOASIGNADO = "Número no asignado."




'--- Metodos Privados ---------------------------------------------------------*
Private Sub Class_Initialize()
    mError = 0
    mValor = -1
    mOrden = -1
    mJuego = JUEGO_DEFECTO
End Sub

'--- Propiedades --------------------------------------------------------------*
'------------------------------------------------------------------------------*
' Procedimiento  : Valor
' Fecha          : 11/01/2004 23:30
' Propósito      : Valor numérico asignado al objeto Numero
'------------------------------------------------------------------------------*
Public Property Get Valor() As Integer
    Valor = mValor
End Property

Public Property Let Valor(ByVal vNewValue As Integer)
    mValor = vNewValue
End Property

'------------------------------------------------------------------------------*
' Procedimiento  : Peso
' Fecha          : 11/01/2004 23:30
' Propósito      : Devuelve la cadena Alto si el Numero está por encima de
'                  la mitad del rango del juego
' Retorno        :
'------------------------------------------------------------------------------*
Public Property Get Peso() As String
    Dim mMedium As Integer
    
    Select Case mJuego
        Case Bonoloto, LoteriaPrimitiva: mMedium = 25
        Case Euromillones: mMedium = 26
        Case GordoPrimitiva: mMedium = 28
    End Select
    
    If mValor = -1 Then
        Peso = LT_ERROR
    Else
        
        If mValor < mMedium Then
            Peso = LT_BAJO
        Else
            Peso = LT_ALTO
        End If
    End If
End Property

'------------------------------------------------------------------------------*
' Procedimiento  : Paridad
' Fecha          : 11/01/2004 23:30
' Propósito      : Devuelve la paridad del Numero: par, impar o #err
'------------------------------------------------------------------------------*
Public Property Get Paridad() As String
    Dim tmp As Integer
    If mValor = -1 Then
        Paridad = LT_ERROR
    Else
        tmp = mValor Mod 2
        If tmp = 0 Then
            Paridad = LT_PAR
        Else
            Paridad = LT_IMPAR
        End If
    End If
End Property

'------------------------------------------------------------------------------*
' Procedimiento  : Terminacion
' Fecha          : 15/08/2018
' Propósito      : Devuelve la terminación del Numero comprendida entre 0..9
' Retorno        : 0, 1, 2, 3..9
'------------------------------------------------------------------------------*
Public Property Get Terminacion() As Integer
    Dim tmp As Integer
    If mValor = -1 Then
        Terminacion = -1
    Else
        tmp = mValor Mod 10
        Terminacion = tmp
    End If
End Property
'------------------------------------------------------------------------------*
' Procedimiento  : Decena
' Fecha          : dd/MMM/yyyy
' Propósito      : Devuelve la decena a la que pertenece
' Retorno        : 1-> 0..9, 2->10..19, y así sucesivamente hasta 6
'------------------------------------------------------------------------------*
Public Property Get Decena() As Integer
    Dim tmp As Integer
    If mValor = -1 Then
        Decena = -1
    Else
        tmp = Fix(mValor / 10) + 1
        Decena = tmp
    End If
End Property

'------------------------------------------------------------------------------*
' Procedimiento  : Septena
' Fecha          : 15/08/2018
' Propósito      : Devuelve la septena a la que pertenece
' Retorno        : 1-> 0..7, 2->8..14 y así sucesivamente hasta 8
'------------------------------------------------------------------------------*
Public Property Get Septena() As Integer
    Dim tmp As Integer
    If mValor = -1 Then
        Septena = -1
    Else
        If (mValor Mod 7) = 0 Then
            tmp = Fix(mValor / 7)
        Else
            tmp = Fix(mValor / 7) + 1
        End If
        Septena = tmp
    End If
End Property

'------------------------------------------------------------------------------*
' Procedimiento  : Orden
' Fecha          : 05-dic-2006 19:59
' Propósito      : Lugar que ocupa el número en la extracción de un sorteo
'------------------------------------------------------------------------------*
Public Property Get Orden() As Integer
    Orden = mOrden
End Property

Public Property Let Orden(ByVal vNewValue As Integer)
    mOrden = vNewValue
End Property

'------------------------------------------------------------------------------*
' Procedimiento  : EsPar
' Fecha          : 15/03/2014
' Propósito      : Devuelve True si el numero es par
' Retorno        : True si es PAR, False si es Impar
'------------------------------------------------------------------------------*
Public Property Get EsPar() As Boolean
    Dim tmp As Integer
    tmp = mValor Mod 2
    If tmp = 0 Then
        EsPar = True
    Else
        EsPar = False
    End If
End Property

'--- Metodos Publicos ---------------------------------------------------------*
'------------------------------------------------------------------------------*
' Procedimiento  : ToString
' Fecha          : 15/03/2014
' Propósito      : Formatea el número en dos dígitos
' Retorno        : Numero formateado
'------------------------------------------------------------------------------*
Public Function ToString()
    If mValor = -1 Then
        ToString = LT_ERROR
    Else
        ToString = Format(mValor, "00")
    End If
End Function
'------------------------------------------------------------------------------*
' Procedimiento  : GetMensaje
' Fecha          : 15/03/2014
' Propósito      : Devuelve el mensaje de error asociado al número según
'                  el tipo de juego
' Retorno        : Mensaje de error
'------------------------------------------------------------------------------*
Public Function GetMensaje() As String
    Dim sResult As String
  On Error GoTo GetMensaje_Error
    
    Select Case mError
     Case 0:    sResult = MSG_NUMVALID
     Case 1:    sResult = MSG_NOVALIDJUEGO06_49
     Case 2:    sResult = MSG_NOVALIDJUEGO05_54
     Case 3:    sResult = MSG_NOVALIDJUEGO05_50
     Case 4:    sResult = MSG_NUMERONOASIGNADO
     Case 5:    sResult = MSG_NOVALIDJUEGO02_12
    End Select

    GetMensaje = sResult

   On Error GoTo 0
   Exit Function

GetMensaje_Error:
   Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
   ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
   Call HandleException(ErrNumber, ErrDescription, "Numero.GetMensaje")
   Err.Raise ErrNumber, ErrSource, ErrDescription
End Function

'------------------------------------------------------------------------------*
' Procedimiento  : EsValido
' Fecha          : 15/03/2014
' Propósito      : Validar el contenido del numero con respecto al juego
' Parámetros     : datJuego
' Retorno        : True si el numero es válido, false si hay errores
'------------------------------------------------------------------------------*
Public Function EsValido(Optional datTipoJuego As ModalidadJuego) As Boolean
  On Error GoTo EsValido_Error
  
    If IsMissing(datTipoJuego) Then
        datTipoJuego = ModalidadJuego.LP_LB_6_49
    End If
    
    mError = 0
 
    If mValor = -1 Then
        mError = ERR_NUMERONOASIGNADO
    Else
        Select Case datTipoJuego
    
            Case LP_LB_6_49:
                If (mValor < 1) Or (mValor > 49) Then
                    mError = ERR_NOVALIDJUEGO06_49
                End If
           
            Case GP_5_54:
                If (mValor < 1) Or (mValor > 54) Then
                    mError = ERR_NOVALIDJUEGO05_54
                End If
    
            Case EU_5_50
                If (mValor < 1) Or (mValor > 50) Then
                    mError = ERR_NOVALIDJUEGO05_50
                End If
                
            Case EU_2_12
                If (mValor < 1) Or (mValor > 12) Then
                    mError = ERR_NOVALIDJUEGO02_12
                End If
                
        End Select
        
    End If
    
    EsValido = IIf(mError = 0, True, False)
    
   On Error GoTo 0
   Exit Function

EsValido_Error:
   Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
   ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
   Call HandleException(ErrNumber, ErrDescription, "Numero.EsValido")
   Err.Raise ErrNumber, ErrSource, ErrDescription
End Function
'' *===========(EOF): Numero.cls

