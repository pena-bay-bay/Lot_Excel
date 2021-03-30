VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sugerencia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' *============================================================================*
' *
' *     Fichero    : Sugerencia.cls
' *
' *     Autor      : Carlos Almela Baeza
' *     Creación   : vie, 13/abr/2012 hh:mm:ss
' *     Versión    : 1.0
' *     Propósito  : Contiene una sugerencia para un sorteo y sus parametros
' *                  de generación
' *============================================================================*
Option Explicit
Option Base 0
'
'--- Variables Privadas -------------------------------------------------------*
Private mJuego                  As Juego
Private mCombinacion            As Combinacion
Private mEstrellas              As Combinacion
Private mReintegro              As Integer
Private mError                  As Integer              ' Código de Error

'--- Propiedades --------------------------------------------------------------*
Public Entidad As EntidadNegocio
Public Metodo As String
Public Parametros As String



Public Property Get Juego() As Juego
    Juego = mJuego
End Property
Public Property Let Juego(ByVal vNewValue As Juego)
    mJuego = vNewValue
End Property



Public Property Get Combinacion() As Combinacion
    Set Combinacion = mCombinacion
End Property
Public Property Set Combinacion(objCombinacion As Combinacion)
    Set mCombinacion = objCombinacion
End Property


Public Property Get Estrellas() As Combinacion
    Set Estrellas = mEstrellas
End Property

Public Property Set Estrellas(objCombinacion As Combinacion)
    Set mEstrellas = objCombinacion
End Property

Public Property Get Reintegro() As Integer
    Reintegro = mReintegro
End Property

Public Property Let Reintegro(vNewValue As Integer)
    mReintegro = vNewValue
End Property

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
            If mCombinacion.Count = 5 Or mCombinacion.Count > 6 Then
                EsMultiple = True
            Else
                EsMultiple = False
            End If
            
        Case GordoPrimitiva:
            If mCombinacion.Count > 5 Then
                EsMultiple = True
            Else
                EsMultiple = False
            End If
            
        Case Euromillones:
            If mCombinacion.Count = 5 And mEstrellas.Count = 2 Then
                EsMultiple = False
            Else
                EsMultiple = True
            End If
    End Select
End Property

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
        Select Case mCombinacion.Count
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
        Select Case mCombinacion.Count
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
        Select Case mCombinacion.Count
        Case 5:
            Select Case mEstrellas.Count
                Case Is = 2: NumeroApuestas = 1
                Case Is = 3: NumeroApuestas = 3
                Case Is = 4: NumeroApuestas = 6
                Case Is = 5: NumeroApuestas = 10
            End Select
        Case 6:
            Select Case mEstrellas.Count
                Case Is = 2: NumeroApuestas = 6
                Case Is = 3: NumeroApuestas = 18
                Case Is = 4: NumeroApuestas = 36
                Case Is = 5: NumeroApuestas = 60
            End Select
            
        Case 7:
            Select Case mEstrellas.Count
                Case Is = 2: NumeroApuestas = 21
                Case Is = 3: NumeroApuestas = 63
                Case Is = 4: NumeroApuestas = 126
                Case Is = 5: NumeroApuestas = 210
            End Select
            
        Case 8:
            Select Case mEstrellas.Count
                Case Is = 2: NumeroApuestas = 56
                Case Is = 3: NumeroApuestas = 168
                Case Is = 4: NumeroApuestas = 336
                Case Is = 5: NumeroApuestas = 560
            End Select
            
        Case 9:
            Select Case mEstrellas.Count
                Case Is = 2: NumeroApuestas = 126
                Case Is = 3: NumeroApuestas = 378
                Case Is = 4: NumeroApuestas = 756
                Case Is = 5: NumeroApuestas = 1260
            End Select
            
        Case 10:
            Select Case mEstrellas.Count
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


'--- Métodos Privados ---------------------------------------------------------*
Private Sub Class_Initialize()
    Set mCombinacion = New Combinacion
    Set Entidad = New EntidadNegocio
    mJuego = JUEGO_DEFECTO
End Sub

Private Sub Class_Terminate()
    Set mCombinacion = Nothing
    Set Entidad = Nothing
End Sub


'--- Métodos Públicos ---------------------------------------------------------*
'------------------------------------------------------------------------------*
' Procedimiento  : IsValid
' Fecha          : mi., 20/may/2020 17:04:14
' Propósito      : validar los atributos de la clase
'------------------------------------------------------------------------------*
Public Function IsValid() As Boolean
 On Error GoTo IsValid_Error
    mError = 0
    '
    '   Validar el contenido de la sugerencia
    '
    If Me.Combinacion.Count = 0 Then
        mError = 1
    End If
'    '
'    '   Validar el método de sugerencia
'    '
'    If Not (Metodo.IsValid) Then
'        mError = mError + 2
'    End If
'    '
'    '   Validar los parámetros estadisticos
'    '
'    If Metodo.TipoProcedimiento = mtdBomboCargado Or _
'    Metodo.TipoProcedimiento = mtdEstadCombinacion Or _
'    Metodo.TipoProcedimiento = mtdEstadistico Then
'        '
'        '   Validar parámetros
'        '
'        If Not Me.Parametros.Validar Then
'            mError = mError + 4
'        End If
'    End If
    
    If mError = 0 Then
        IsValid = True
    Else
        IsValid = False
    End If
  
  On Error GoTo 0
IsValid__CleanExit:
    Exit Function
IsValid_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    Call HandleException(ErrNumber, ErrDescription, "Sugerencia.cls.IsValid", ErrSource)
    Err.Raise ErrNumber, "Sugerencia.cls.IsValid", ErrDescription
End Function





'------------------------------------------------------------------------------*
' Procedimiento  : GetMensaje
' Fecha          : mi., 20/may/2020 17:11:21
' Propósito      : Devuelve el error de inconsistencia
'------------------------------------------------------------------------------*
Public Function GetMensaje() As String
    Dim m_Mensaje           As String
  On Error GoTo GetMensaje_Error
    
    If mError = 0 Then
         m_Mensaje = MSG_NOERROR
    Else
         m_Mensaje = MSG_HAYERRORES & vbCrLf
    End If
    If mError And 1 Then
         m_Mensaje = m_Mensaje & MSG_COMBISUGEVACIA & vbCrLf
    End If
    If mError And 2 Then
         m_Mensaje = m_Mensaje & MSG_METODOSUGERROR & vbCrLf
    End If
    If mError And 4 Then
         m_Mensaje = m_Mensaje & MSG_PARAMSUGERROR & vbCrLf
    End If
    '
    '
    GetMensaje = m_Mensaje
   
On Error GoTo 0
GetMensaje_CleanExit:
    Exit Function

GetMensaje_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    Call HandleException(ErrNumber, ErrDescription, "ParametrosMuestra.GetMensaje", ErrSource)
    Err.Raise ErrNumber, "ParametrosMuestra.GetMensaje", ErrDescription
End Function



'------------------------------------------------------------------------------*
' Procedimiento  : ToString
' Fecha          : mi., 20/may/2020 17:04:14
' Propósito      :
'------------------------------------------------------------------------------*
Public Function ToString() As String
    Dim sResult As String
    Dim mModJuego As String
        
    If mError = 0 Then
        Select Case mJuego
            Case Bonoloto, LoteriaPrimitiva: mModJuego = "6/49"
            Case GordoPrimitiva: mModJuego = "5/54"
            Case Euromillones: mModJuego = "5/50"
        End Select
            
        ' Sugerencia (Modalidad-fecha): nn-nn-nn-nn-nn-nn-nn-nn, Metodo
        sResult = "Sugerencia: ($1-$2) $3 , $4"
        sResult = Replace(sResult, "$1", mModJuego)
'        sResult = Replace(sResult, "$2", Format(Parametros.FechaAnalisis, "dd/mm/yyyy"))
        sResult = Replace(sResult, "$3", mCombinacion.ToString(True))
'        sResult = Replace(sResult, "$4", Metodo.ToString)
        ToString = sResult
    Else
        ToString = MSG_SUGERENCIAERROR
    End If
End Function

'------------------------------------------------------------------------------*
' Procedimiento  : Sort
' Fecha          :
' Propósito      :
'------------------------------------------------------------------------------*
Public Sub Sort(Optional datAscendente As Boolean)
    Dim mComb    As Combinacion
    Dim mText    As String
    
    Set mComb = New Combinacion
    
    If IsMissing(datAscendente) Then
        datAscendente = True
    End If
    mText = mCombinacion.ToString(True)
    mComb.Texto = mText
    Set mCombinacion = mComb
    
    Set mComb = Nothing
End Sub
'' *===========(EOF): Sugerencia.cls


