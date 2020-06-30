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
Private mModalidad              As ModalidadJuego
Private mCombinacion            As Combinacion
Private mError                  As Integer      ' Código de Error

'--- Propiedades --------------------------------------------------------------*
Public Entidad As EntidadNegocio
Public Metodo As Metodo
Public Parametros As ParametrosMuestra



Public Property Get Modalidad() As ModalidadJuego
    Modalidad = mModalidad
End Property
Public Property Let Modalidad(ByVal eModalidad As ModalidadJuego)
    mModalidad = eModalidad
End Property



Public Property Get Combinacion() As Combinacion
    Set Combinacion = mCombinacion

End Property
Public Property Set Combinacion(objCombinacion As Combinacion)
    Set mCombinacion = objCombinacion
End Property




'--- Métodos Privados ---------------------------------------------------------*
Private Sub Class_Initialize()
    Set mCombinacion = New Combinacion
    Set Entidad = New EntidadNegocio
    Set Metodo = New Metodo
    Set Parametros = New ParametrosMuestra
    mModalidad = LP_LB_6_49
End Sub

Private Sub Class_Terminate()
    Set mCombinacion = Nothing
    Set Entidad = Nothing
    Set Metodo = Nothing
    Set Parametros = Nothing
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
    '
    '   Validar el método de sugerencia
    '
    If Not (Metodo.IsValid) Then
        mError = mError + 2
    End If
    '
    '   Validar los parámetros estadisticos
    '
    If Metodo.TipoProcedimiento = mtdBomboCargado Or _
    Metodo.TipoProcedimiento = mtdEstadCombinacion Or _
    Metodo.TipoProcedimiento = mtdEstadistico Then
        '
        '   Validar parámetros
        '
        If Not Me.Parametros.Validar Then
            mError = mError + 4
        End If
    End If
    
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
        Select Case mModalidad
            Case LP_LB_6_49: mModJuego = "6/49"
            Case GP_5_54: mModJuego = "5/54"
            Case EU_5_50: mModJuego = "5/50"
            Case EU_2_12: mModJuego = "2/12"
        End Select
            
        ' Sugerencia (Modalidad-fecha): nn-nn-nn-nn-nn-nn-nn-nn, Metodo
        sResult = "Sugerencia: ($1-$2) $3 , $4"
        sResult = Replace(sResult, "$1", mModJuego)
        sResult = Replace(sResult, "$2", Format(Parametros.FechaAnalisis, "dd/mm/yyyy"))
        sResult = Replace(sResult, "$3", mCombinacion.TextoOrdenado)
        sResult = Replace(sResult, "$4", Metodo.ToString)
        ToString = sResult
    Else
        ToString = MSG_SUGERENCIAERROR
    End If
End Function

'' *===========(EOF): Sugerencia.cls
