VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmMuestra 
   Caption         =   "Parámetros de la Muestra"
   ClientHeight    =   3885
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5040
   OleObjectBlob   =   "frmMuestra.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "frmMuestra"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' *============================================================================*
' *
' *     Fichero    : frmMuestra.frm
' *
' *     Autor      : Carlos Almela Baeza
' *     Creación   : 10/04/2007 23:25
' *     Versión    : 1.0
' *     Propósito  : Seleccionar los parámetros de la muestra estadistica
' *
' *============================================================================*
Option Explicit
Option Base 0
    
Private m_objParMuestra             As ParametrosMuestra  ' Parametros de la Muestra
Private DB                          As New BdDatos        ' Base de datos
Dim m_ini_DataBase                  As Date               ' Fecha inicial de la base de datos
Dim m_fin_DataBase                  As Date               ' Fecha final de la base de datos
'
'   Mensajería
'
Private m_sMensaje As String            ' Mensaje al operador
Private Const MSG_HAYERROR = "Existe errores en los parámetros de entrada: " & vbCrLf
Private Const MSG_FALTADATE = " * No ha definido la fecha de análisis." & vbCrLf
Private Const MSG_NODATE = " * La fecha de análisis no es valida." & vbCrLf
Private Const MSG_FALTANUMREG = " * No ha definido el número de registros." & vbCrLf
Private Const MSG_NONUMREG = " * El número de registros no es un Numero." & vbCrLf
Private Const MSG_NOCOMBO = " * No ha seleccionado elementos del rango." & vbCrLf
Private Const MSG_NUMREGNOCERO = " * El número de registros no puede ser 0." & vbCrLf
Private Const MSG_FALTANUMRANG = " * Falta número de rango." & vbCrLf
Private Const MSG_NONUMRANG = " * Falta Rango definido." & vbCrLf

'---------------------------------------------------------------------------------------
' Procedure : UserForm_Initialize
' DateTime  : 02/05/2007 23:28
' Author    : Carlos Almela Baeza
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub UserForm_Initialize()
    Dim m_dtFecha As Date
    Dim m_oInfo   As New InfoSorteo
   On Error GoTo UserForm_Initialize_Error
   
    Set m_objParMuestra = New ParametrosMuestra
    m_dtFecha = Date                            ' Toma la fecha de hoy
    m_ini_DataBase = DB.PrimerResultado         ' Obtiene la menor de las fechas de la
                                                ' base de datos
    m_fin_DataBase = DB.UltimoResultado         ' Obtiene la mayor de las fechas de la
                                                ' base de datos
                                                ' Formatea y asigna la fecha inicial
    '
    ' Si la fecha sugerida no es fecha del sorteo
    ' pasar a la siguiente fecha
    If Not m_oInfo.EsFechaSorteo(m_dtFecha) Then
        m_dtFecha = m_oInfo.GetProximoSorteo(m_dtFecha)
    End If
'    '
'    ' Si la fecha del proximo sorteo es la del último resultado
'    ' pasar a la siguiente fecha
'    '
'    If m_dtFecha = m_fin_DataBase Then
'        m_dtFecha = m_oInfo.GetProximoSorteo(m_dtFecha)
'    End If
    
    '
    ' inicializamos el contenido de la caja de texto con la fecha fomrateada
    '
    txtFechaAnalisis.Text = Format(m_dtFecha, "dd/mm/yyyy")
    '
    ' Indicamos por defecto que el intervalo seran de 40 registros o 5 semanas
    ' TODO: guardar la ultima seleccion en una tabla de parametros y recuperarla
    '
    txtNumReg.Text = Format(40, "00")
    txtNumRang.Text = Format(5, "00")
    '
    ' Indicamos que el análisis se ralizará por dias
    '
    optTipo.Value = True
    optTipo2.Value = False
    '
    '  Cargamos el combo con las  opciones temporales
    '
    cboRango.Clear
    cboRango.AddItem "dias", 0
    cboRango.AddItem "semanas", 1
    cboRango.AddItem "meses", 2
    cboRango.AddItem "trimestres", 3
    cboRango.AddItem "semestres", 4
    cboRango.AddItem "años", 5
    '
    ' Seleccionamos 1- Semanas
    '
    cboRango.ListIndex = 1

   On Error GoTo 0
     Exit Sub

UserForm_Initialize_Error:

    Dim sNumber As Integer
    Dim sDescription As String
    Dim sSource As String
    With Err
        sNumber = .Number
        sDescription = .Description
        sSource = .Source
    End With
    Call HandleException(sNumber, sDescription, sSource)
    '    Sube el error
    Err.Raise sNumber, sSource, sDescription
End Sub

'---------------------------------------------------------------------------------------
' Procedure : IsValid
' Author    : CAB3780Y
' Date      : 29/03/2011
' Purpose   : Función de validación de los controles del Formulario
'---------------------------------------------------------------------------------------
'
Private Function IsValid() As Boolean
    Dim m_bError As Boolean
   
   On Error GoTo IsValid_Error
    m_bError = True                  ' Ok de inicio
    m_sMensaje = ""
    '
    '   Evaluar cada uno de los controles del formulario
    '
    '
    '   Fecha de Análisis
    '
    If (Len(txtFechaAnalisis.Text) = 0) Then
        m_sMensaje = m_sMensaje & MSG_FALTADATE
        m_bError = False
    End If
    '
    '
    '
    If (Len(txtFechaAnalisis.Text) > 0) And (Not IsDate(txtFechaAnalisis.Text)) Then
        m_sMensaje = m_sMensaje & MSG_NODATE
        m_bError = False
    End If
    '
    '   Si la opción seleccionada es por Dias
    '   se evalua el rango de dias
    '
    If optTipo.Value Then
        If (Len(txtNumRang.Text) = 0) Then
            m_sMensaje = m_sMensaje & MSG_FALTANUMRANG
            m_bError = False
        End If
        If (Len(txtNumRang.Text) > 0) _
        And (Not IsNumeric(txtNumRang.Text)) Then
            m_sMensaje = m_sMensaje & MSG_NONUMRANG
            m_bError = False
        End If
        If (cboRango.ListIndex = -1) Then
            m_sMensaje = m_sMensaje & MSG_NOCOMBO
            m_bError = False
        End If
    End If
    '
    '   Si la opción seleccionada es por Registros
    '   se evalua Numero de registros
    '
    If optTipo2.Value Then
        If (Len(txtNumReg.Text) = 0) Then
            m_sMensaje = m_sMensaje & MSG_FALTANUMREG
            m_bError = False
        ElseIf (Not IsNumeric(txtNumReg.Text)) Then
            m_sMensaje = m_sMensaje & MSG_NONUMREG
            m_bError = False
        ElseIf (CInt(txtNumReg.Text) = 0) Then
            m_sMensaje = m_sMensaje & MSG_NUMREGNOCERO
            m_bError = False
        End If
    End If
    '
    '   sE VALID
    '
    If Not m_objParMuestra.Validar Then
        m_sMensaje = m_sMensaje & m_objParMuestra.GetMensaje
            m_bError = False
    End If
    '
    '
    '
    If (Not m_bError) Then
        m_sMensaje = MSG_HAYERROR & m_sMensaje
    End If
    
    IsValid = m_bError
   
   On Error GoTo 0
     Exit Function

IsValid_Error:

    Dim sNumber As Integer
    Dim sDescription As String
    Dim sSource As String
    With Err
        sNumber = .Number
        sDescription = .Description
        sSource = .Source
    End With
    Call HandleException(sNumber, sDescription, sSource)
    '    Sube el error
    Err.Raise sNumber, sSource, sDescription
End Function

'---------------------------------------------------------------------------------------
' Procedure : optTipo_Click
' Author    : CAB3780Y
' Date      : 29/03/2011
' Purpose   : Opcion cálculo por número de días
'---------------------------------------------------------------------------------------
'
Private Sub optTipo_Click()
    If optTipo.Value Then
        txtNumReg.Enabled = False
        txtNumRang.Enabled = True
        cboRango.Enabled = True
    End If
End Sub

'---------------------------------------------------------------------------------------
' Procedure : optTipo2_Click
' Author    : CAB3780Y
' Date      : 29/03/2011
' Purpose   : Opción Cálculo por Numero de registros
'---------------------------------------------------------------------------------------
'
Private Sub optTipo2_Click()
    If optTipo2.Value Then
        txtNumReg.Enabled = True
        txtNumRang.Enabled = False
        cboRango.Enabled = False
    End If
End Sub

'---------------------------------------------------------------------------------------
' Procedure : cmdCancelar_Click
' DateTime  : 02/05/2007 23:26
' Author    : Carlos Almela Baeza
' Purpose   : Cancela la presentación del formulario
'---------------------------------------------------------------------------------------
'
Private Sub cmdCancelar_Click()
    Me.Tag = BOTON_CERRAR               ' Asigna en la etiqueta del formulario
                                        ' la clave de cerrar
    Me.Hide                             ' Oculta el formulario
End Sub

'---------------------------------------------------------------------------------------
' Procedure : cmdEjecutar_Click
' DateTime  : 02/05/2007 23:27
' Author    : Carlos Almela Baeza
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub cmdEjecutar_Click()
    EstablecerPeriodo
    If IsValid() Then                       ' Si los datos del formulario son Válidos
        Me.Tag = EJECUTAR                   ' Se indica que el comando es EJECUTAR
                                            ' en la etiqueta del formulario
        Me.Hide                             ' Se oculta el fromulario
    Else
                                            ' Se obtienen los errores del formulario
                                            ' en la variable m_sMensaje
        MsgBox m_sMensaje, vbExclamation + vbOKOnly, Me.Caption
    End If
    
End Sub

'---------------------------------------------------------------------------------------
' Procedure : EstablecerPeriodo
' Author    : CAB3780Y
' Date      : 29/03/2011
' Purpose   : Define el periodo inicial para la muestra
'---------------------------------------------------------------------------------------
'
Private Sub EstablecerPeriodo()
    Dim m_dias As Integer                   'Número de días a calcular
    Dim m_reg As Integer                    'Número de días a calcular
    Dim m_factorMultiplicador As Integer    'Factor multiplicador de dias
    Dim m_info As InfoSorteo                'Info Sorteo
    
   On Error GoTo EstablecerPeriodo_Error
    Set m_info = New InfoSorteo
    m_info.Constructor Bonoloto
    '
    '   Asigna fecha de analisis
    '
    m_objParMuestra.FechaAnalisis = CDate(txtFechaAnalisis.Text)
    '
    '  Se obtiene el día anterior a la fecha de analisis
    '
    m_objParMuestra.FechaFinal = m_info.GetAnteriorSorteo(m_objParMuestra.FechaAnalisis)
    '
    '  Si la fecha supera las fechas de la base de datos
    '  Se asigna la última
    '
    If m_objParMuestra.FechaFinal > m_fin_DataBase Then
        m_objParMuestra.FechaFinal = m_fin_DataBase
    End If
    '
    '   Tratamiento por dias
    '
    If (optTipo.Value) Then
        m_factorMultiplicador = 1
        Select Case cboRango.ListIndex
            Case 0:  m_factorMultiplicador = 1          'Dias
            Case 1:  m_factorMultiplicador = 7          'Semanas
            Case 2:  m_factorMultiplicador = 30         'Meses
            Case 3:  m_factorMultiplicador = 90         'Trimestres
            Case 4:  m_factorMultiplicador = 180        'Semestres
            Case 5:  m_factorMultiplicador = 365        'años
            Case Else: m_factorMultiplicador = 1
        End Select
        m_dias = CInt(txtNumRang.Text)
        m_dias = m_dias * m_factorMultiplicador
        m_objParMuestra.FechaInicial = m_info.GetAnteriorSorteo(m_objParMuestra.FechaFinal - m_dias)
    '
    '   Tratamiento por registros
    '
    ElseIf (optTipo2.Value) Then
        m_reg = CInt(txtNumReg.Text)
        m_objParMuestra.FechaInicial = DB.GetFecha(m_objParMuestra.FechaFinal, m_reg)
    End If

    

   On Error GoTo 0
     Exit Sub

EstablecerPeriodo_Error:

    Dim sNumber As Integer
    Dim sDescription As String
    Dim sSource As String
    With Err
        sNumber = .Number
        sDescription = .Description
        sSource = .Source
    End With
    Call HandleException(sNumber, sDescription, sSource)
    '    Sube el error
    Err.Raise sNumber, sSource, sDescription
End Sub

' *============================================================================*
' *     Procedure  : ParMuestra
' *     Fichero    : frmMuestra
' *     Autor      : Carlos Almela Baeza
' *     Creacion   : sáb, 21/01/2012 21:51
' *     Asunto     :
' *============================================================================*
'
Public Property Get ParMuestra() As ParametrosMuestra
    Set ParMuestra = m_objParMuestra
End Property

Public Property Set ParMuestra(objParMuestra As ParametrosMuestra)
    Set m_objParMuestra = objParMuestra
End Property

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    Me.Tag = BOTON_CERRAR               ' Asigna en la etiqueta del formulario
                                        ' la clave de cerrar
    Me.Hide                             ' Oculta el formulario
End Sub
