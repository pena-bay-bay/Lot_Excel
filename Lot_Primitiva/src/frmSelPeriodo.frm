VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSelPeriodo 
   Caption         =   "Selección de un Periodo de Fechas"
   ClientHeight    =   2130
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   6525
   OleObjectBlob   =   "frmSelPeriodo.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "frmSelPeriodo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



'---------------------------------------------------------------------------------------
' Module    : frmSelPeriodo
' Author    : Charly
' Date      : 04/04/2012
' Purpose   : Formulario de captura de un periodo predefinido
'---------------------------------------------------------------------------------------

Option Explicit
'
'   Objeto interno de periodo
'
Private m_objRangoFechas As Periodo
'
'   Variable de control de validación
'
Private m_varFechasDefinidas As Variant
'
'
'
Private m_lErrorValidacion As Long
'
'   Validaciones de la clase
'
Private Const ERR_VALFEINICIOMAL            As Long = 1
Private Const ERR_VALFEFINMAL               As Long = 2
Private Const ERR_VALFEFINMENORFEINICIAL    As Long = 64
'
'   Mensajes de validación
'
Private Const MSG_ERRORESVALIDACION         As String = "Los datos del formulario no cumplen las siguientes validaciones:" & vbCrLf
Private Const MSG_VALFEINICIOMAL            As String = "* La fecha de inicio no es una fecha válida."
Private Const MSG_VALFEFINMAL               As String = "* La fecha de fin no es una fecha válida."
Private Const MSG_VALFEFINMENORFEINICIAL    As String = "* La fecha de fin es inferior a la de inicio."

'---------------------------------------------------------------------------------------
' Procedure : RangoFechas
' Author    : Charly
' Date      : 04/04/2012
' Purpose   : Propiedad del formulario un rango de fechas
'---------------------------------------------------------------------------------------
'
Public Property Get RangoFechas() As Periodo

  On Error GoTo RangoFechas_Error

    Set RangoFechas = m_objRangoFechas

   On Error GoTo 0
   Exit Property

RangoFechas_Error:
     Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
   ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
   '   Audita el error
   Call HandleException(ErrNumber, ErrDescription, "frmSelPeriodo.RangoFechas")
   '   Lanza el Error
   Err.Raise ErrNumber, ErrSource, ErrDescription
End Property

'---------------------------------------------------------------------------------------
' Procedure : RangoFechas
' Author    : Charly
' Date      : 04/04/2012
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Set RangoFechas(objRangoFechas As Periodo)

  On Error GoTo RangoFechas_Error

    Set m_objRangoFechas = objRangoFechas
    '
    '   Configuramos el combo con el tipo de fechas
    '
    On Error Resume Next
    cboPerMuestra.ListIndex = m_objRangoFechas.Tipo_Fecha
    If Err.Number <> 0 Then
        cboPerMuestra.ListIndex = -1
    End If
   On Error GoTo 0
   Exit Property

RangoFechas_Error:
     Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
   ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
   '   Audita el error
   Call HandleException(ErrNumber, ErrDescription, "frmSelPeriodo.RangoFechas")
   '   Lanza el Error
   Err.Raise ErrNumber, ErrSource, ErrDescription
End Property

'---------------------------------------------------------------------------------------
' Procedure : UserForm_Initialize
' Author    : Charly
' Date      : 05/04/2012
' Purpose   : Inicializa los controles del formulario
'---------------------------------------------------------------------------------------
'
Private Sub UserForm_Initialize()
    '
    '   Inicializa el objeto principal
    '
    Set m_objRangoFechas = New Periodo
    '
    '
    '
    m_varFechasDefinidas = Array(ctHoy, ctAyer, ctUltimaSemana, _
                           ctUltimaQuincena, ctUltimoMes, ctLoQueVadeSemana, _
                           ctLoQueVadeMes, ctLoQueVadeTrimestre, _
                           ctUltimoTrimestre, ctPersonalizadas)
    '
    '   Carga el combo de las fechas
    '
    m_objRangoFechas.CargaCombo cboPerMuestra, m_varFechasDefinidas
    '
    '   Establece un periodo última semana (al cambiar el combo
    '   se actualiza el resto de controles)
    '
    cboPerMuestra.ListIndex = 2
End Sub

'---------------------------------------------------------------------------------------
' Procedure : VisualizaControles
' Author    : Charly
' Date      : 05/04/2012
' Purpose   : Asigna valores a los controles y visualiza
'---------------------------------------------------------------------------------------
'
Private Sub VisualizaControles()

  On Error GoTo VisualizaControles_Error
    '
    '   Formatea la fecha inicial y la coloca en el correspondiente caja de texto
    '
    txtFechaMuestraIni.Text = Format(m_objRangoFechas.FechaInicial, "dd/mm/yyyy")
    '
    '   Formatea la fecha final y la coloca en el correspondiente caja de texto
    '
    txtFechaMuestraFin.Text = Format(m_objRangoFechas.FechaFinal, "dd/mm/yyyy")
    '
    '   Si son fecha predefinidas bloqueamos los controles
    '
    If m_objRangoFechas.Tipo_Fecha <> ctPersonalizadas Then
        '
        '   Desactiva los controles
        '
        txtFechaMuestraIni.Enabled = False
        txtFechaMuestraFin.Enabled = False
    Else
        '
        '   Si la fecha es personalizada activa los controles
        '
        txtFechaMuestraIni.Enabled = True
        txtFechaMuestraFin.Enabled = True
    End If
    
    
   On Error GoTo 0
   Exit Sub

VisualizaControles_Error:
     Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
   ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
   '   Audita el error
   Call HandleException(ErrNumber, ErrDescription, "frmSelPeriodo.VisualizaControles")
   '   Lanza el Error
   Err.Raise ErrNumber, ErrSource, ErrDescription
End Sub

'---------------------------------------------------------------------------------------
' Procedure : cboPerMuestra_Change
' Author    : Charly
' Date      : 05/04/2012
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub cboPerMuestra_Change()
    
  On Error GoTo cboPerMuestra_Change_Error
    '
    '   Actualizamos el tipo de periodo seleccionado
    '
    m_objRangoFechas.Texto = cboPerMuestra.Text
    '
    '   Actualizamos los controles
    '
    VisualizaControles
   On Error GoTo 0
   Exit Sub

cboPerMuestra_Change_Error:
     Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
   ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
   '   Audita el error
   Call HandleException(ErrNumber, ErrDescription, "frmSelPeriodo.cboPerMuestra_Change")
   '   Lanza el Error
   Err.Raise ErrNumber, ErrSource, ErrDescription
End Sub

'---------------------------------------------------------------------------------------
' Procedure : cmdAceptar_Click
' Author    : Charly
' Date      : 04/04/2012
' Purpose   : Se ha pulsado el boton aceptar
'---------------------------------------------------------------------------------------
'
Private Sub cmdAceptar_Click()
    Dim strMensaje As String
    If IsValid Then
        Me.Tag = EJECUTAR
        Me.Hide
    Else
        strMensaje = MensajeValidacion
        Call MsgBox(strMensaje, vbExclamation Or vbSystemModal, Application.Name)
    End If
End Sub

'---------------------------------------------------------------------------------------
' Procedure : cmdCancelar_Click
' Author    : Charly
' Date      : 04/04/2012
' Purpose   : Se ha pulsado el boton cancelar
'---------------------------------------------------------------------------------------
'
Private Sub cmdCancelar_Click()
    Me.Tag = BOTON_CERRAR
    Me.Hide
End Sub

'---------------------------------------------------------------------------------------
' Procedure : IsValid
' Author    : Charly
' Date      : 04/04/2012
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Function IsValid() As Boolean
    Dim bResult As Boolean
   
  On Error GoTo IsValid_Error

    m_lErrorValidacion = 0
    '
    '   Fecha Inicial
    '
    If IsDate(txtFechaMuestraIni.Text) Then
        m_objRangoFechas.FechaInicial = CDate(txtFechaMuestraIni.Text)
    Else
        m_lErrorValidacion = m_lErrorValidacion + ERR_VALFEINICIOMAL
    End If
    '
    '   Fecha final
    '
    If IsDate(txtFechaMuestraFin.Text) Then
        m_objRangoFechas.FechaFinal = CDate(txtFechaMuestraFin.Text)
    Else
        m_lErrorValidacion = m_lErrorValidacion + ERR_VALFEFINMAL
    End If
    '
    '   Relación entre fechas
    '
    If (m_objRangoFechas.FechaInicial > m_objRangoFechas.FechaFinal) Then
        m_lErrorValidacion = m_lErrorValidacion + ERR_VALFEFINMENORFEINICIAL
    End If
    
    
    bResult = IIf(m_lErrorValidacion = 0, True, False)
    
    IsValid = bResult


   On Error GoTo 0
   Exit Function

IsValid_Error:
     Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
   ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
   '   Audita el error
   Call HandleException(ErrNumber, ErrDescription, "frmSelPeriodo.IsValid")
   '   Lanza el Error
   Err.Raise ErrNumber, ErrSource, ErrDescription
End Function

'---------------------------------------------------------------------------------------
' Procedure : MensajeValidacion
' Author    : Charly
' Date      : 04/04/2012
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Function MensajeValidacion() As String
    Dim sResult             As String
    
  On Error GoTo MensajeValidacion_Error
    '
    '   Si no hay error inicializamos el mensaje
    '
    If m_lErrorValidacion = 0 Then
        sResult = ""
    Else
        '
        '   Si hay algun error inicializamos la cabecera del error
        '
        sResult = MSG_ERRORESVALIDACION
    End If
    '
    '   Si se ha producido el error de fecha de inicio mal
    '   agregamos el mensaje a la cabecera
    '
    If (m_lErrorValidacion And ERR_VALFEINICIOMAL) Then
        sResult = sResult & vbTab & MSG_VALFEINICIOMAL & vbCrLf
    End If
    
    If (m_lErrorValidacion And ERR_VALFEFINMAL) Then
        sResult = sResult & vbTab & MSG_VALFEFINMAL & vbCrLf
    End If
    
    If (m_lErrorValidacion And ERR_VALFEFINMENORFEINICIAL) Then
        sResult = sResult & vbTab & MSG_VALFEFINMENORFEINICIAL & vbCrLf
    End If
    '
    '   Devolvemos el mensaje
    '
    MensajeValidacion = sResult

   On Error GoTo 0
   Exit Function

MensajeValidacion_Error:
     Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
   ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
   '   Audita el error
   Call HandleException(ErrNumber, ErrDescription, "frmSelPeriodo.MensajeValidacion")
   '   Lanza el Error
   Err.Raise ErrNumber, ErrSource, ErrDescription

End Function


Public Property Get Periodo() As Periodo
    Set Periodo = m_objRangoFechas
End Property

Public Property Set Periodo(ByVal vNewValue As Periodo)
    Set m_objRangoFechas = vNewValue
    '
    '   Actualizamos los controles
    '
    VisualizaControles
End Property
