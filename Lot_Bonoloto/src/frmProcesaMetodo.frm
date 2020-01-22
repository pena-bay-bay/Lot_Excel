VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmProcesaMetodo 
   Caption         =   "Configuración del Metodo de Sugerencia"
   ClientHeight    =   5175
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   5850
   OleObjectBlob   =   "frmProcesaMetodo.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "frmProcesaMetodo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



'---------------------------------------------------------------------------------------
' Module    : frmProcesaMetodo
' Author    : Charly
' Date      : 26/03/2012
' Purpose   :
'---------------------------------------------------------------------------------------
Option Explicit
'
'   Control de errores de validación
'
Private m_lErrorValidacion As Long
'
'
'   Validaciones de la clase
'
Private Const ERR_VALFEINICIOMAL            As Long = 1
Private Const ERR_VALFEFINMAL               As Long = 2
Private Const ERR_VALCOMBOORDEN             As Long = 4
Private Const ERR_VALDIASMUESTRACERO        As Long = 8
Private Const ERR_VALREGISTROSCERO          As Long = 16
Private Const ERR_VALCOMBOAGRUPACION        As Long = 32
Private Const ERR_VALPRONOSTICOSCERO        As Long = 64

'
'   Mensajes de validación
'
Private Const MSG_ERRORESVALIDACION         As String = "Los datos del formulario no cumplen las siguientes validaciones:" & vbCrLf
Private Const MSG_VALFEINICIOMAL            As String = "* La fecha de inicio no es una fecha válida."
Private Const MSG_VALFEFINMAL               As String = "* La fecha de fin no es una fecha válida."
Private Const MSG_VALCOMBOORDEN             As String = "* Debe seleccionar un método de ordenación."
Private Const MSG_VALDIASMUESTRACERO        As String = "* Los días de la muestra no pueden ser cero."
Private Const MSG_VALREGISTROSCERO          As String = "* El número de registros no puede ser cero."
Private Const MSG_VALCOMBOAGRUPACION        As String = "* Debe seleccionar un método de agrupacion."
Private Const MSG_VALPRONOSTICOSCERO        As String = "* El número de pronosticos no puede ser cero."
'
'   Variables internas del formulario
'
Private m_objMetodo                         As metodo    ' Objeto método a cumplimentar
Private m_objPeriodoSorteos                 As Periodo      ' Parámetro del proceso rango de sorteos

'---------------------------------------------------------------------------------------
' Procedure : Metodo
' Author    : Charly
' Date      : 31/03/2012
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get metodo() As metodo

    Set metodo = m_objMetodo

End Property

'---------------------------------------------------------------------------------------
' Procedure : Metodo
' Author    : Charly
' Date      : 31/03/2012
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Set metodo(objMetodo As metodo)

    Set m_objMetodo = objMetodo
    '
    '   Visualizar el metodo
    '
    VisualizarMetodo
End Property

'---------------------------------------------------------------------------------------
' Procedure : PeriodoSorteos
' Author    : Charly
' Date      : 31/03/2012
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get PeriodoSorteos() As Periodo

  On Error GoTo PeriodoSorteos_Error

    Set PeriodoSorteos = m_objPeriodoSorteos

   On Error GoTo 0
   Exit Property

PeriodoSorteos_Error:
     Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
   ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
   '   Audita el error
   Call HandleException(ErrNumber, ErrDescription, "frmProcesaMetodo.PeriodoSorteos")
   '   Lanza el Error
   Err.Raise ErrNumber, ErrSource, ErrDescription
End Property

'---------------------------------------------------------------------------------------
' Procedure : PeriodoSorteos
' Author    : Charly
' Date      : 31/03/2012
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Set PeriodoSorteos(objPeriodoSorteos As Periodo)

  On Error GoTo PeriodoSorteos_Error

    Set m_objPeriodoSorteos = objPeriodoSorteos
    '
    '   Visualizar los valores
    '
    VisualizarPeriodoSorteo
   On Error GoTo 0
   Exit Property

PeriodoSorteos_Error:
     Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
   ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
   '   Audita el error
   Call HandleException(ErrNumber, ErrDescription, "frmProcesaMetodo.PeriodoSorteos")
   '   Lanza el Error
   Err.Raise ErrNumber, ErrSource, ErrDescription
End Property

'---------------------------------------------------------------------------------------
' Procedure : UserForm_Initialize
' Author    : Charly
' Date      : 31/03/2012
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub UserForm_Initialize()
    Dim m_objMetEng                 As MetodoEngine         ' Motor de Metodos
  On Error GoTo UserForm_Initialize_Error
    '
    '   Carga los combos del formulario
    '
    Set m_objMetEng = New MetodoEngine
    m_objMetEng.CargaOrdenaciones Me.cboOrdenacion
    m_objMetEng.CargaAgrupaciones Me.cboAgrupacion
    '
    '   Crea el objeto Metodo
    '
    Set m_objMetodo = New metodo
    '
    '   Crea el periodo de sorteos y se configura a la última semana
    '
    Set m_objPeriodoSorteos = New Periodo
    m_objPeriodoSorteos.Tipo_Fecha = ctUltimaSemana
    '
    '   Consulta si existe un método por defecto
    '
'    If m_objMetEng.HasDefaultMetodo Then
'        '
'        '   Si existe lo carga
'        '
'        Set m_objMetodo = m_objMetEng.GetDefaultMetodo()
'    End If
    
    '
    '   Inicializa controles del formulario
    '
    Inicializa_Controles_Formulario
    VisualizarPeriodoSorteo


   On Error GoTo 0
   Exit Sub

UserForm_Initialize_Error:
     Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
   ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
   '   Audita el error
   Call HandleException(ErrNumber, ErrDescription, "frmProcesaMetodo.UserForm_Initialize")
   '   Lanza el Error
   Err.Raise ErrNumber, ErrSource, ErrDescription
End Sub

'---------------------------------------------------------------------------------------
' Procedure : Inicializa_Controles_Formulario
' Author    : Charly
' Date      : 03/04/2012
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub Inicializa_Controles_Formulario()

  On Error GoTo Inicializa_Controles_Formulario_Error
    '
    '   Inicializa los Combos
    '
    cboOrdenacion.ListIndex = 1 ' m_objMetodo.ParametrosMetodo.or
    cboAgrupacion.ListIndex = 1 ' m_objMetodo.
    '
    '
    '
    chkSentido.Value = True
    chkSentido.Caption = IIf(chkSentido.Value, "Ascendente", "Descendente")
    
    chkCriterioMuestra.Value = True
    chkCriterioMuestra.Caption = IIf(chkCriterioMuestra.Value, "Por Dias", "Por Registros")
    lblDiasMuestras.Caption = IIf(chkCriterioMuestra.Value, "Dias Análisis", "N. Registros")
    '
    '   Inicializa cajas de texto
    '
    txtPronosticos.Text = Format(6, "00")
    '
    '
    '
    txtDiasMuestra.Text = Format(45, "00")
    
   On Error GoTo 0
   Exit Sub

Inicializa_Controles_Formulario_Error:
     Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
   ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
   '   Audita el error
   Call HandleException(ErrNumber, ErrDescription, "frmProcesaMetodo.Inicializa_Controles_Formulario")
   '   Lanza el Error
   Err.Raise ErrNumber, ErrSource, ErrDescription
End Sub

'---------------------------------------------------------------------------------------
' Procedure : VisualizarPeriodoSorteo
' Author    : Charly
' Date      : 05/04/2012
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub VisualizarPeriodoSorteo()
    txtFechaIni.Text = Format(m_objPeriodoSorteos.FechaInicial, "dd/mm/yyyy")
    txtFechaFin.Text = Format(m_objPeriodoSorteos.FechaFinal, "dd/mm/yyyy")
End Sub

'---------------------------------------------------------------------------------------
' Procedure : VisualizarMetodo
' Author    : Charly
' Date      : 05/04/2012
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub VisualizarMetodo()
    
    cboOrdenacion.ListIndex = m_objMetodo.Parametros.CriteriosOrdenacion
    chkSentido.Value = m_objMetodo.Parametros.SentidoOrdenacion
    cboAgrupacion.ListIndex = m_objMetodo.Parametros.CriteriosAgrupacion

    txtPronosticos.Text = Format(m_objMetodo.Parametros.Pronosticos, "00")
    chkCriterioMuestra.Value = m_objMetodo.TipoMuestra
    If (m_objMetodo.TipoMuestra) Then
        txtDiasMuestra.Text = Format(m_objMetodo.Parametros.DiasAnalisis, "00")
    Else
        txtDiasMuestra.Text = Format(m_objMetodo.Parametros.NumeroSorteos, "00")
    End If
    
End Sub

'---------------------------------------------------------------------------------------
' Procedure : chkCriterioMuestra_Click
' Author    : Charly
' Date      : 06/04/2012
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub chkCriterioMuestra_Click()
    chkCriterioMuestra.Caption = IIf(chkCriterioMuestra.Value, "Por Dias", "Por Registros")
    lblDiasMuestras.Caption = IIf(chkCriterioMuestra.Value, "Dias Análisis", "N. Registros")
End Sub

'---------------------------------------------------------------------------------------
' Procedure : chkSentido_Click
' Author    : Charly
' Date      : 06/04/2012
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub chkSentido_Click()
    chkSentido.Caption = IIf(chkSentido.Value, "Ascendente", "Descendente")
End Sub

'---------------------------------------------------------------------------------------
' Procedure : cmdGetPeriodo_Click
' Author    : Charly
' Date      : 05/04/2012
' Purpose   : Invoca al formulario para seleccionar fechas
'---------------------------------------------------------------------------------------
'
Private Sub cmdGetPeriodo_Click()
    Dim m_objForm           As frmSelPeriodo
    '
    '   Creamos el objeto
    '
    Set m_objForm = New frmSelPeriodo
    '
    '   Establecemos el periodo de sorteos
    '
    Set m_objForm.RangoFechas = m_objPeriodoSorteos
    '
    '   Se muestra de forma modal
    '
    m_objForm.Show vbModal
    '
    '
    '
    If m_objForm.Tag = EJECUTAR Then
        '
        '   Obtenemos el periodo seleccionado
        '
        Set m_objPeriodoSorteos = m_objForm.RangoFechas
        '
        '   Lo visualizamos
        '
        txtFechaIni.Text = Format(m_objPeriodoSorteos.FechaInicial, "dd/mm/yyyy")
        txtFechaFin.Text = Format(m_objPeriodoSorteos.FechaFinal, "dd/mm/yyyy")
    End If
    '
    '   Elimina el objeto de memoria
    '
    Set m_objForm = Nothing
End Sub

'---------------------------------------------------------------------------------------
' Procedure : cmdCancelar_Click
' Author    : Charly
' Date      : 03/04/2012
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub cmdCancelar_Click()
    Me.Tag = BOTON_CERRAR
    Me.Hide
End Sub

'---------------------------------------------------------------------------------------
' Procedure : cmdEjecutar_Click
' Author    : Charly
' Date      : 03/04/2012
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub cmdEjecutar_Click()
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
' Procedure : IsValid
' Author    : Charly
' Date      : 03/04/2012
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
    If IsDate(txtFechaIni.Text) Then
        m_objPeriodoSorteos.FechaInicial = CDate(txtFechaIni.Text)
    Else
        m_lErrorValidacion = m_lErrorValidacion + ERR_VALFEINICIOMAL
    End If
    '
    '   Fecha final
    '
    If IsDate(txtFechaFin.Text) Then
        m_objPeriodoSorteos.FechaFinal = CDate(txtFechaFin.Text)
    Else
        m_lErrorValidacion = m_lErrorValidacion + ERR_VALFEFINMAL
    End If
    '
    '   Numero de Pronosticos
    '
    If IsNumeric(txtPronosticos.Text) Then
        m_objMetodo.Parametros.Pronosticos = CInt(txtPronosticos.Text)
        If (m_objMetodo.Parametros.Pronosticos = 0) Then
            m_lErrorValidacion = m_lErrorValidacion + ERR_VALPRONOSTICOSCERO
        End If
    Else
        m_lErrorValidacion = m_lErrorValidacion + ERR_VALPRONOSTICOSCERO
    End If
    '
    '   Seleccion de Combo de ordenacion
    '
    If cboOrdenacion.ListIndex = -1 Then
         m_lErrorValidacion = m_lErrorValidacion + ERR_VALCOMBOORDEN
    Else
         m_objMetodo.Parametros.CriteriosOrdenacion = cboOrdenacion.ListIndex
    End If
    '
    '   Seleccion de Combo de agrupacion
    '
    If cboAgrupacion.ListIndex = -1 Then
         m_lErrorValidacion = m_lErrorValidacion + ERR_VALCOMBOAGRUPACION
    Else
         m_objMetodo.Parametros.CriteriosAgrupacion = cboAgrupacion.ListIndex
    End If
    
    
    
    '
    '   Se inicilizan parametros del objeto
    '
    m_objMetodo.Parametros.SentidoOrdenacion = chkSentido.Value
    m_objMetodo.TipoMuestra = chkCriterioMuestra.Value
    '
    '   Dias de analisis / registros
    '
    If IsNumeric(txtDiasMuestra.Text) Then
        If m_objMetodo.TipoMuestra Then
            m_objMetodo.Parametros.DiasAnalisis = CInt(txtDiasMuestra.Text)
            If (m_objMetodo.Parametros.DiasAnalisis = 0) Then
                m_lErrorValidacion = m_lErrorValidacion + ERR_VALDIASMUESTRACERO
            End If
        Else
            m_objMetodo.Parametros.NumeroSorteos = CInt(txtDiasMuestra.Text)
            If (m_objMetodo.Parametros.NumeroSorteos = 0) Then
                m_lErrorValidacion = m_lErrorValidacion + ERR_VALREGISTROSCERO
            End If
        End If
    Else
        If m_objMetodo.TipoMuestra Then
            m_lErrorValidacion = m_lErrorValidacion + ERR_VALDIASMUESTRACERO
        Else
            m_lErrorValidacion = m_lErrorValidacion + ERR_VALREGISTROSCERO
        End If
    End If
    '
    '
    '
    bResult = IIf(m_lErrorValidacion = 0, True, False)
    
    IsValid = bResult

   On Error GoTo 0
   Exit Function

IsValid_Error:
     Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
   ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
   '   Audita el error
   Call HandleException(ErrNumber, ErrDescription, "frmProcesaMetodo.IsValid")
   '   Lanza el Error
   Err.Raise ErrNumber, ErrSource, ErrDescription
End Function

'---------------------------------------------------------------------------------------
' Procedure : MensajeValidacion
' Author    : Charly
' Date      : 03/04/2012
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Function MensajeValidacion() As String
    Dim sResult             As String
    
    
  On Error GoTo MensajeValidacion_Error

    If m_lErrorValidacion = 0 Then
        sResult = ""
    Else
        sResult = MSG_ERRORESVALIDACION
    End If
    
    If (m_lErrorValidacion And ERR_VALFEINICIOMAL) Then
        sResult = sResult & vbTab & MSG_VALFEINICIOMAL & vbCrLf
    End If
    
    If (m_lErrorValidacion And ERR_VALFEFINMAL) Then
        sResult = sResult & vbTab & MSG_VALFEFINMAL & vbCrLf
    End If
    
    If (m_lErrorValidacion And ERR_VALPRONOSTICOSCERO) Then
        sResult = sResult & vbTab & MSG_VALPRONOSTICOSCERO & vbCrLf
    End If
    
    If (m_lErrorValidacion And ERR_VALREGISTROSCERO) Then
        sResult = sResult & vbTab & MSG_VALREGISTROSCERO & vbCrLf
    End If
    
    If (m_lErrorValidacion And ERR_VALCOMBOAGRUPACION) Then
        sResult = sResult & vbTab & MSG_VALCOMBOAGRUPACION & vbCrLf
    End If
    
    If (m_lErrorValidacion And ERR_VALDIASMUESTRACERO) Then
        sResult = sResult & vbTab & MSG_VALDIASMUESTRACERO & vbCrLf
    End If
    
    If (m_lErrorValidacion And ERR_VALCOMBOORDEN) Then
        sResult = sResult & vbTab & MSG_VALCOMBOORDEN & vbCrLf
    End If
    
    MensajeValidacion = sResult

   On Error GoTo 0
   Exit Function

MensajeValidacion_Error:
     Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
   ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
   '   Audita el error
   Call HandleException(ErrNumber, ErrDescription, "frmProcesaMetodo.MensajeValidacion")
   '   Lanza el Error
   Err.Raise ErrNumber, ErrSource, ErrDescription
End Function

