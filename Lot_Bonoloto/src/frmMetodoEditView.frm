VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmMetodoEditView 
   Caption         =   "Método de Sugerencia"
   ClientHeight    =   7530
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5715
   OleObjectBlob   =   "frmMetodoEditView.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "frmMetodoEditView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' *============================================================================*
' *
' *     Fichero    : frmMetodoEditView.frm
' *
' *     Autor      : Carlos Almela Baeza
' *     Creación   : vi., 27/mar/2020 19:40:08
' *     Versión    : 1.0
' *     Propósito  : VIsta de edición del Metodo
' *
' *============================================================================*
Option Explicit
Option Base 0
'
'--- Variables Privadas -------------------------------------------------------*
'
'   Control de errores de validación
'
Private mCtrl               As MetodoController
Private mModel              As MetodoModel
Private mFilters            As FiltrosCombinacion
Private mFilter             As FiltroCombinacion
Private lngFraActive        As Long
Private lngFraDesactive     As Long
'--- Errores ------------------------------------------------------------------*
'--- Mensajes -----------------------------------------------------------------*
Private Const MSG_AVSBORRARMETODO     As String = "¿Está seguro que quie" & _
                                      "re eliminar este método de sugerencia?"
Private Const MSG_AVSBORRARFILTROS    As String = "¿Está seguro que quie" & _
                                      "re eliminar todos los filtros definidos?"
Private Const MSG_AVSFILTROERROR      As String = "Falta el tipo de filt" & _
                                      "ro o el valor del mismo."
'--- Constantes ---------------------------------------------------------------*
Private Const LT_ASCENDENTE As String = "Ascendente"
Private Const LT_DESCENDENTE As String = "Descendente"
Private Const LT_PORDIAS As String = "Por días"
Private Const LT_POREGISTROS As String = "Por registros"
Private Const LT_LBLDIAS As String = "Dias Análisis"
Private Const LT_LBLREGISTROS As String = "Número Sorteos"

'--- Propiedades --------------------------------------------------------------*
Public ModoAlta As Boolean

'--- Métodos Privados ---------------------------------------------------------*
'------------------------------------------------------------------------------*
' Procedimiento  : UserForm_Initialize
' Fecha          : vi., 27/mar/2020 19:40:08
' Propósito      : configura el formulario según el método
'------------------------------------------------------------------------------*
Private Sub UserForm_Initialize()
    Dim mMatriz       As Variant
    Dim mNombre       As String
    Dim i             As Integer
  
  On Error GoTo UserForm_Initialize_Error
    '
    '   Inicializamos combo ordenación
    '
    mMatriz = Split(NOMBRES_ORDENACION, ";")
    cboOrdenacion.Clear
    For i = 0 To UBound(mMatriz)
        mNombre = mMatriz(i)
        cboOrdenacion.AddItem mNombre, i
    Next i
    '
    '   Inicializamos Combo de Agrupación
    '
    mMatriz = Split(NOMBRES_AGRUPACION, ";")
    cboAgrupacion.Clear
    For i = 0 To UBound(mMatriz)
        mNombre = mMatriz(i)
        cboAgrupacion.AddItem mNombre, i
    Next i
    '
    '   Inicializamos Combo Tipo de procedimientos
    '
    mMatriz = Split(NOMBRES_PROCEDIMIENTOMETODO, ";")
    cboTipoProcedimiento.Clear
    For i = 0 To UBound(mMatriz)
        mNombre = mMatriz(i)
        cboTipoProcedimiento.AddItem mNombre, i
    Next i
    '
    '   Configuramos el tipo de muestra a días
    '
    chkCriterioMuestra = True
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
    cboRango.ListIndex = -1
    '
    '   Inicializamos Combo Tipo de filtros
    '
    mMatriz = Split(NOMBRES_TIPOS_FILTRO, ";")
    cboTipoFiltro.Clear
    For i = 0 To UBound(mMatriz)
        mNombre = mMatriz(i)
        cboTipoFiltro.AddItem mNombre, i
    Next i
    '
    '   Creamos la variable filtro
    '
    Set mFilter = New FiltroCombinacion
    Set mFilters = New FiltrosCombinacion
    '
    '   Definimos colores de activación y desactivación de los frames
    '
    lngFraActive = fraFiltros.BackColor
    lngFraDesactive = RGB(255, 255, 255)
    '
    '   Creamos el controlador
    '
    Set mCtrl = New MetodoController
    
  On Error GoTo 0
UserForm_Initialize__CleanExit:
    Exit Sub
UserForm_Initialize_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    Call HandleException(ErrNumber, ErrDescription, "frmMetodoEditView.UserForm_Initialize", ErrSource)
    Err.Raise ErrNumber, "frmSugerencia.UserForm_Initialize", ErrDescription
End Sub


Private Sub UserForm_Terminate()
    Set mCtrl = Nothing
    Set mFilter = Nothing
    Set mFilters = Nothing
End Sub


'------------------------------------------------------------------------------*
' Procedimiento  : cboTipoFiltro_Change
' Fecha          : vi., 10/abr/2020 11:24:17
' Propósito      : configura el combo de valores para el filtro
'------------------------------------------------------------------------------*
Private Sub cboTipoFiltro_Change()
    Dim mVar As Variant
    Dim mTxt As String
    Dim i As Integer
    
    If cboTipoFiltro.ListIndex > -1 And IsNumeric(txtPronosticos.Text) Then
        mFilter.TipoFiltro = cboTipoFiltro.ListIndex + 1
        '
        '   Obtenemos los valores para un filtro determinado
        '
        mVar = mFilter.GetValoresFiltros(CInt(txtPronosticos.Text))
        cboValorFiltro.Clear
        For i = 0 To UBound(mVar)
            cboValorFiltro.AddItem mVar(i)
        Next i
    Else
        cboValorFiltro.Clear
    End If
End Sub




'------------------------------------------------------------------------------*
' Procedimiento  : cboTipoProcedimiento_Change
' Fecha          : vi., 10/abr/2020 11:24:17
' Propósito      : configura el formulario según el método
'------------------------------------------------------------------------------*
Private Sub cboTipoProcedimiento_Change()

    Init_Formulario
    
    Select Case cboTipoProcedimiento.ListIndex
        Case mtdSinDefinir:
            fraMuestra.Enabled = False
            fraMuestra.BackColor = lngFraDesactive
            fraParametros.Enabled = False
            fraParametros.BackColor = lngFraDesactive
            fraFiltros.Enabled = False
            fraFiltros.BackColor = lngFraDesactive
            
        Case mtdAleatorio, mtdBombo:
            fraMuestra.Enabled = False
            fraMuestra.BackColor = lngFraDesactive
            fraParametros.Enabled = False
            fraParametros.BackColor = lngFraDesactive
            fraFiltros.Enabled = True
            fraFiltros.BackColor = lngFraActive
            
        Case mtdBomboCargado, mtdEstadCombinacion:
            fraMuestra.Enabled = True
            fraMuestra.BackColor = lngFraActive
            fraParametros.Enabled = True
            fraParametros.BackColor = lngFraActive
            fraFiltros.Enabled = True
            fraFiltros.BackColor = lngFraActive
            
        Case mtdEstadistico
            fraMuestra.Enabled = True
            fraMuestra.BackColor = lngFraActive
            fraParametros.Enabled = True
            fraParametros.BackColor = lngFraActive
            fraFiltros.Enabled = False
            fraFiltros.BackColor = lngFraDesactive
            
    End Select
End Sub




'------------------------------------------------------------------------------*
' Procedimiento  : chkCriterioMuestra_Click
' Fecha          : vi., 10/abr/2020 11:24:17
' Propósito      : configura el formulario según el método
'------------------------------------------------------------------------------*
Private Sub chkCriterioMuestra_Click()
    If chkCriterioMuestra.Value Then
        chkCriterioMuestra.Caption = LT_POREGISTROS
        lblTipoMuestra = LT_LBLREGISTROS
        cboRango.Visible = False
    Else
        chkCriterioMuestra.Caption = LT_PORDIAS
        lblTipoMuestra = LT_LBLDIAS
        cboRango.Visible = True
    End If
End Sub




'------------------------------------------------------------------------------*
' Procedimiento  : chkCriterioMuestra_Click
' Fecha          : vi., 10/abr/2020 11:24:17
' Propósito      : configura el formulario según el método
'------------------------------------------------------------------------------*
Private Sub chkSentido_Click()
    If chkSentido.Value Then
        chkSentido.Caption = LT_ASCENDENTE
    Else
        chkSentido.Caption = LT_DESCENDENTE
    End If
End Sub



'------------------------------------------------------------------------------*
' Procedimiento  : cmdAgregarFiltro_Click
' Fecha          : vi., 10/abr/2020 10:45:11
' Propósito      : Agrega un filtro a la colección
'------------------------------------------------------------------------------*
Private Sub cmdAgregarFiltro_Click()
      
  On Error GoTo cmdAgregarFiltro_Click_Error
    '
    '   Analizamos el tipo de filtro y el valor
    '
    If cboTipoFiltro.ListIndex > -1 And Len(cboValorFiltro.Value) > 0 Then
        '
        '   Definimos el filtro seleccionado
        '
        Set mFilter = New FiltroCombinacion
        mFilter.TipoFiltro = cboTipoFiltro.ListIndex + 1
        mFilter.FilterValue = cboValorFiltro.Value
        '
        '   agregamos el filtro a la colección
        '
        mFilters.Add mFilter
        '
        '   Refrescamos lista de valores
        '
        lstFiltros_Refresh mFilters
        '
        '   desactivamos en campo de numero de pronósticos
        '
        If mFilters.Count > 2 Then
            txtPronosticos.Enabled = False
        Else
            txtPronosticos.Enabled = True
        End If
    Else
        MsgBox MSG_AVSFILTROERROR, vbOKOnly + vbExclamation, Me.Caption
    End If
    
  On Error GoTo 0
cmdAgregarFiltro_Click__CleanExit:
    Exit Sub
cmdAgregarFiltro_Click_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    Call HandleException(ErrNumber, ErrDescription, "frmMetodoEditView.cmdAgregarFiltro_Click", ErrSource)
    Call MsgBox(ErrDescription, vbCritical Or vbSystemModal, ThisWorkbook.Name)
End Sub


'------------------------------------------------------------------------------*
' Procedimiento  : lstFiltros_Refresh
' Fecha          : vi., 27/mar/2020 19:40:08
' Propósito      : Actualiza los filtros de la lista
'------------------------------------------------------------------------------*
Private Sub lstFiltros_Refresh(datNewFiltros As FiltrosCombinacion)
    Dim i As Integer
    Dim mStr As String
    lstFiltros.Clear
    For i = 1 To datNewFiltros.Count
        mStr = datNewFiltros.Items(i).ToString()
        lstFiltros.AddItem mStr
    Next i
End Sub


'------------------------------------------------------------------------------*
' Procedimiento  : cmdCancel_Click
' Fecha          : vi., 27/mar/2020 19:40:08
' Propósito      : Cerrar el formulario
'------------------------------------------------------------------------------*
Private Sub cmdCancel_Click()
    Me.Tag = BOTON_CERRAR
    Me.Hide
End Sub

'------------------------------------------------------------------------------*
' Procedimiento  : cmdDelete_Click
' Fecha          : ma., 21/abr/2020 18:18:37
' Propósito      : Borra el método de la lista
'------------------------------------------------------------------------------*
Private Sub cmdDelete_Click()
  On Error GoTo cmdDelete_Click_Error
    '
    '   Preguntamos si quiere borrar el método
    '
    If MsgBox(MSG_AVSBORRARMETODO, vbYesNo + vbQuestion, Me.Caption) Then
        Me.Tag = BORRAR
        Me.Hide
    End If
  On Error GoTo 0
cmdDelete_Click__CleanExit:
    Exit Sub
cmdDelete_Click_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    Call HandleException(ErrNumber, ErrDescription, "frmMetodoEditView.cmdDelete_Click", ErrSource)
    Call MsgBox(ErrDescription, vbCritical Or vbSystemModal, ThisWorkbook.Name)
End Sub



'------------------------------------------------------------------------------*
' Procedimiento  : cmdDeleteFiltro_Click
' Fecha          : vi., 10/abr/2020 11:47:38
' Propósito      : Borrar los filtros establecidos
'------------------------------------------------------------------------------*
Private Sub cmdDeleteFiltro_Click()
  On Error GoTo cmdDeleteFiltro_Click_Error
    '
    '   Si no tenemos filtros definidos el boton no actua
    '
    If mFilters.Count > 0 Then
        '
        '   Preguntamos si se quieren borrar
        '
        If MsgBox(MSG_AVSBORRARFILTROS, vbQuestion + vbYesNo, Me.Caption) Then
            '
            '   Si se confirma se borran los filtros
            '
            mFilters.Clear
            lstFiltros.Clear
        End If
    End If
    
  On Error GoTo 0
cmdDeleteFiltro_Click__CleanExit:
    Exit Sub
cmdDeleteFiltro_Click_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    Call HandleException(ErrNumber, ErrDescription, "frmMetodoEditView.cmdDeleteFiltro_Click", ErrSource)
    Call MsgBox(ErrDescription, vbCritical Or vbSystemModal, ThisWorkbook.Name)
End Sub

'------------------------------------------------------------------------------*
' Procedimiento  : cmdSave_Click
' Fecha          :
' Propósito      : Guardar el metodo en la base de datos
'------------------------------------------------------------------------------*
Private Sub cmdSave_Click()
    Dim mMtd        As Metodo
  On Error GoTo cmdSave_Click_Error
    '
    '   Actualizamos el método del modelo
    '
    Set mMtd = mModel.Metodo
    With mMtd
        .TipoProcedimiento = cboTipoProcedimiento.ListIndex
        .CriteriosAgrupacion = cboAgrupacion.ListIndex
        .CriteriosOrdenacion = cboOrdenacion.ListIndex
        .SentidoOrdenacion = chkSentido.Value
        .TipoMuestra = chkCriterioMuestra.Value
        If .TipoMuestra Then
            If IsNumeric(txtDiasMuestra.Text) Then
                .NumeroSorteos = CInt(txtDiasMuestra.Text)
            Else
                .NumeroSorteos = 0
            End If
        Else
            Select Case cboRango.ListIndex
                Case 1: .DiasAnalisis = CInt(txtDiasMuestra.Text) * 7
                Case 2: .DiasAnalisis = CInt(txtDiasMuestra.Text) * 30
                Case 3: .DiasAnalisis = CInt(txtDiasMuestra.Text) * 90
                Case 4: .DiasAnalisis = CInt(txtDiasMuestra.Text) * 180
                Case 5: .DiasAnalisis = CInt(txtDiasMuestra.Text) * 365
                Case Else:
                    If IsNumeric(txtDiasMuestra.Text) Then
                        .DiasAnalisis = CInt(txtDiasMuestra.Text)
                    Else
                        .DiasAnalisis = 0
                    End If
            End Select
        End If
        .Pronosticos = CInt(txtPronosticos.Text)
        Set .Filtros = mFilters
    End With
    '
    '   Preguntar si es valido el método
    '
    If mMtd.IsValid Then
        Set mModel.Metodo = mMtd
        Me.Tag = EJECUTAR
        Me.Hide
    Else
        MsgBox mMtd.GetMessage, vbOKOnly + vbExclamation, Me.Caption
    End If
    
    
  On Error GoTo 0
cmdSave_Click__CleanExit:
    Exit Sub
cmdSave_Click_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    Call HandleException(ErrNumber, ErrDescription, "frmMetodoEditView.cmdSave_Click", ErrSource)
    Call MsgBox(ErrDescription, vbCritical Or vbSystemModal, ThisWorkbook.Name)
End Sub




'------------------------------------------------------------------------------*


'--- Métodos Públicos ---------------------------------------------------------*
'------------------------------------------------------------------------------*
' Procedimiento  : Refresh
' Fecha          : vi., 03/abr/2020 16:42:10
' Propósito      : Actualizar los controles del formulario con los datos del
'                  modelo
'------------------------------------------------------------------------------*
Public Sub Refresh(datModel As MetodoModel)
    Dim mMtd        As Metodo
    Dim i           As Integer
  On Error GoTo Refresh_Error
    '
    '   Actualizamos el modelo del formulario
    '
    Set mModel = datModel
    Set mMtd = mModel.Metodo
    '
    '   Establecemos el tipo de procedimiento
    '
    cboTipoProcedimiento.ListIndex = mMtd.TipoProcedimiento
    ' si es alta dejar la palabra New
    txtId.Text = Format(mMtd.Id, "#0")
    txtId.Enabled = False
    '
    '   Establecemos parametros de estadistica
    '
    cboOrdenacion.ListIndex = mMtd.CriteriosOrdenacion
    chkSentido.Value = mMtd.SentidoOrdenacion
    cboAgrupacion.ListIndex = mMtd.CriteriosAgrupacion
    '
    '   Establecemos parametros de muestra
    '
    chkCriterioMuestra.Value = mMtd.TipoMuestra
    If mMtd.TipoMuestra Then
        ' Muestra en número de registros
        txtDiasMuestra.Text = Format(mMtd.NumeroSorteos, "#0")
    Else
        If mMtd.DiasAnalisis > 0 Then
            Select Case True
            Case mMtd.DiasAnalisis Mod 365 = 0:
                cboRango.ListIndex = 5
                txtDiasMuestra.Text = Format(mMtd.DiasAnalisis / 365, "#0")
            
            Case mMtd.DiasAnalisis Mod 180 = 0:
                cboRango.ListIndex = 4
                txtDiasMuestra.Text = Format(mMtd.DiasAnalisis / 180, "#0")
            
            Case mMtd.DiasAnalisis Mod 90 = 0:
                cboRango.ListIndex = 3
                txtDiasMuestra.Text = Format(mMtd.DiasAnalisis / 90, "#0")
            
            Case mMtd.DiasAnalisis Mod 30 = 0:
                cboRango.ListIndex = 2
                txtDiasMuestra.Text = Format(mMtd.DiasAnalisis / 30, "#0")
            
            Case mMtd.DiasAnalisis Mod 7 = 0:
                cboRango.ListIndex = 1
                txtDiasMuestra.Text = Format(mMtd.DiasAnalisis / 7, "#0")
            
            Case Else
                cboRango.ListIndex = 0
                txtDiasMuestra.Text = Format(mMtd.DiasAnalisis, "#0")
            End Select
        Else
            cboRango.ListIndex = 0
            txtDiasMuestra.Text = Format(mMtd.DiasAnalisis, "#0")
        End If
    End If
    '
    '   Establecemos parametros del filtro
    '
    txtPronosticos.Text = Format(mMtd.Pronosticos, "#0")
    '
    '   Cargamos filtros en listBox
    '
    lstFiltros.Clear
    Set mFilters = mMtd.Filtros
    '
    '   Cargamos el listbox de filtros
    '
    For i = 1 To mFilters.Count
        Set mFilter = mFilters.Items(i)
        lstFiltros.AddItem mFilter.ToString
    Next i
    '
    '   Configuramos los botones según el modo de pantalla
    '
    If ModoAlta Then
        cmdDelete.Enabled = False
    Else
        cmdDelete.Enabled = True
    End If
    cmdSave.Enabled = True
    
  On Error GoTo 0
Refresh__CleanExit:
    Exit Sub
Refresh_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    Call HandleException(ErrNumber, ErrDescription, "frmMetodoEditView.Refresh", ErrSource)
    Call MsgBox(ErrDescription, vbCritical Or vbSystemModal, ThisWorkbook.Name)
End Sub

Private Sub Init_Formulario()
    '
    '   Establecemos parametros de estadistica
    '
    cboOrdenacion.ListIndex = ordSinDefinir
    chkSentido.Value = True
    cboAgrupacion.ListIndex = grpSinDefinir
    '
    '   Establecemos parametros de muestra
    '
    chkCriterioMuestra.Value = True
    txtDiasMuestra.Text = Empty
    cboRango.ListIndex = -1
    '
    '   Configuramos datos del filtro
    '
    txtPronosticos.Text = Empty
    cboTipoFiltro.ListIndex = -1
    cboValorFiltro.ListIndex = -1
    lstFiltros.Clear
End Sub
' *===========(EOF): frmMetodoEditView


