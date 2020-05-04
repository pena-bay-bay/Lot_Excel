VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmMetodoSelectView 
   Caption         =   "Generador de Sugerencias"
   ClientHeight    =   4785
   ClientLeft      =   45
   ClientTop       =   495
   ClientWidth     =   7815
   OleObjectBlob   =   "frmMetodoSelectView.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "frmMetodoSelectView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' *============================================================================*
' *
' *     Fichero    : frmMetodoSelectView.frm
' *
' *     Autor      : Carlos Almela Baeza
' *     Creación   : ma., 10/mar/2020 14:04:34
' *     Versión    : 1.0
' *     Propósito  : Seleccionar los métodos de sugerencias
' *
' *============================================================================*
Option Explicit
Option Base 0

'
'--- Variables Privadas -------------------------------------------------------*
Private mCtrl                       As MetodoController ' Controlador
Private mCol                        As Collection       ' Ids seleccionados
Private mAllMetodos                 As Boolean          ' Selección de todos
Private mJuego                      As Juego            ' sorteo
Private mFechaAnalisis              As Date             ' Fecha de pronostico
Private mPronosticos                As Integer          ' Numero de pronosticos
Private mCurrentId                  As Integer          ' Id seleccionado
Private mCurrentSelected            As Boolean          ' Id marcado
Private mPaginaActual               As Integer          ' Página actual
Private mTotalPaginas               As Integer          ' Total de paginas
Private mTotalRecords               As Integer          ' Total registros
Private mInfo                       As InfoSorteo       ' Información de los sorteos
Private i                           As Integer
Private m_sMensaje                  As String           ' Mensaje de error
'--- Constantes ---------------------------------------------------------------*
Private Const LT_NOMETODOS As String = "( No existen métodos de sugerencia )"
Private Const LT_LABELPAGINAS As String = "Página #1/#2"

'--- Mensajes -----------------------------------------------------------------*
Private Const MSG_HAYERROR As String = "El formulario contiene los siguientes errores: " & vbCrLf
Private Const MSG_NODATE As String = "- se requiere una fecha." & vbCrLf
Private Const MSG_DATENOVALID As String = "- no es una fecha válida." & vbCrLf
Private Const MSG_NODATESORTEO As String = "- La fecha no es una fecha del juego." & vbCrLf
Private Const MSG_NONUMERO As String = "- se requiere un número de pronosticos." & vbCrLf
Private Const MSG_NUMERONOVALID As String = "- se requiere un número válido para el pronostico." & vbCrLf
Private Const MSG_NUMOUTRANGE As String = "- pronostico fuera de rango: [5..11]." & vbCrLf
Private Const MSG_NOMETODOS As String = "- no se han seleccionado métodos." & vbCrLf


'--- Errores ------------------------------------------------------------------*
'--- Propiedades --------------------------------------------------------------*
Public Property Get SelectedIds() As Collection
    Set SelectedIds = mCol
End Property

'------------------------------------------------------------------------------*
' Procedimiento  : SelectedIds
' Fecha          : ma., 31/mar/2020 19:13:03
' Propósito      : Métodos seleccionados
'------------------------------------------------------------------------------*
Public Property Set SelectedIds(ByVal vNewValue As Collection)
    Set mCol = vNewValue
End Property
'------------------------------------------------------------------------------*
' Procedimiento  : AllMetodosSelected
' Fecha          : mi., 01/abr/2020 08:08:50
' Propósito      : Indicador de todos los métodos
'------------------------------------------------------------------------------*
Public Property Get AllMetodosSelected() As Boolean
    AllMetodosSelected = mAllMetodos
End Property
'------------------------------------------------------------------------------*
' Procedimiento  : lstMetodos_Change
' Fecha          : ma., 31/mar/2020 19:13:03
' Propósito      : Registro seleccionado
'------------------------------------------------------------------------------*
'
Private Sub lstMetodos_Change()
    Dim Selected As Boolean
    If lstMetodos.ListIndex > -1 And mTotalRecords > 0 Then
        mCurrentId = lstMetodos.List(lstMetodos.ListIndex, 0)
        mCurrentSelected = lstMetodos.Selected(lstMetodos.ListIndex)
        SelectRegistro
    End If
End Sub

'--- Métodos Privados ---------------------------------------------------------*
'------------------------------------------------------------------------------*
' Procedimiento  : UserForm_Initialize
' Fecha          : ma., 10/mar/2020 14:11:34
' Propósito      : Inicializa los objetos del formulario
'------------------------------------------------------------------------------*
Private Sub UserForm_Initialize()
  On Error GoTo UserForm_Initialize_Error
    '
    '   Definimos colección de Ids seleccionados
    '
    Set mCol = New Collection
    '
    '   Creamos el controlador del caso de uso
    '
    Set mCtrl = New MetodoController
    '
    '   Invocamos al controlador a inicializar la vista
    '
    mCtrl.InitSelectView Me
    
    
  On Error GoTo 0
UserForm_Initialize__CleanExit:
    Exit Sub
UserForm_Initialize_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    Call HandleException(ErrNumber, ErrDescription, "frmMetodoSelectView.UserForm_Initialize", ErrSource)
    Err.Raise ErrNumber, "frmMetodoSelectView.UserForm_Initialize", ErrDescription
End Sub


Private Sub UserForm_Terminate()
    Set mCtrl = Nothing
    Set mCol = Nothing
End Sub

'------------------------------------------------------------------------------*
' Procedimiento  : chkSelAllMetodos_Click
' Fecha          : ma., 31/mar/2020 19:14:10
' Propósito      : Seleccionar todos los metodos
'------------------------------------------------------------------------------*
Private Sub chkSelAllMetodos_Click()
  On Error GoTo chkSelAllMetodos_Error
    '
    '   Si está seleccionado desmarcar los metodos
    '
    mAllMetodos = chkSelAllMetodos.Value
    '
    '   Actualiza formulario
    '
    mCtrl.RefreshSelect Me
    
  On Error GoTo 0
chkSelAllMetodos__CleanExit:
    Exit Sub
chkSelAllMetodos_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    Call HandleException(ErrNumber, ErrDescription, "frmMetodoSelectView.chkSelAllMetodos", ErrSource)
    Err.Raise ErrNumber, "frmMetodoSelectView.chkSelAllMetodos", ErrDescription
End Sub

'------------------------------------------------------------------------------*
' Procedimiento  : cmdAgregar
' Fecha          : ma., 31/mar/2020 19:16:20
' Propósito      : Agregar un método nuevo
'------------------------------------------------------------------------------*
Private Sub cmdAgregar_Click()
    Dim i As Integer
    Dim mModel As MetodoModel
    
  On Error GoTo cmdAgregar_Click_Error
    '
    '   Obtenemos pronosticos
    '
    If IsNumeric(txtPronosticos.Text) Then
        i = CInt(txtPronosticos.Text)
    Else
        i = 0
    End If
    '
    '   Invocamos al controlador  con Id =  0
    '
    Set mModel = mCtrl.InitEditView(0, i)
    '
    '   Refrescamos el formulario
    '
    Me.Refresh mModel
  On Error GoTo 0
cmdAgregar_Click__CleanExit:
    Exit Sub
cmdAgregar_Click_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    Call HandleException(ErrNumber, ErrDescription, "frmMetodoSelectView.cmdAgregar_Click", ErrSource)
    Call MsgBox(ErrDescription, vbCritical Or vbSystemModal, ThisWorkbook.Name)
End Sub

'------------------------------------------------------------------------------*
' Procedimiento  : cmdCancelar_Click
' Fecha          : ma., 10/mar/2020 14:12:30
' Propósito      : Cerrar el formulario
'------------------------------------------------------------------------------*
Private Sub cmdCancelar_Click()
    Me.Tag = BOTON_CERRAR
    Me.Hide
End Sub

'------------------------------------------------------------------------------*
' Procedimiento  : cmdEditar
' Fecha          : ma., 31/mar/2020 19:17:20
' Propósito      : Editar el metodo seleccionado
'------------------------------------------------------------------------------*
Private Sub cmdEditar_Click()
    Dim mModel As MetodoModel
    
  On Error GoTo cmdEditar_Click_Error
    '
    '   Calcular Current Id
    '
    If lstMetodos.ListIndex > -1 Then
        '
        '   Obtenemos el Id seleccionado
        '
        mCurrentId = lstMetodos.List(lstMetodos.ListIndex, 0)
        '
        '   Invocamos al controlador con Id
        '
        Set mModel = mCtrl.InitEditView(mCurrentId, mPronosticos)
        '
        '   Refrescamos el formulario
        '
        Me.Refresh mModel
    Else
        MsgBox " No se ha seleccionado ningún método", Me.Caption
    End If
        
  On Error GoTo 0
cmdEditar_Click__CleanExit:
    Exit Sub
cmdEditar_Click_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    Call HandleException(ErrNumber, ErrDescription, "frmMetodoSelectView.cmdEditar_Click", ErrSource)
    Call MsgBox(ErrDescription, vbCritical Or vbSystemModal, ThisWorkbook.Name)
End Sub

'------------------------------------------------------------------------------*
' Procedimiento  : cmdEjecutar_Click
' Fecha          : ma., 31/mar/2020 19:20:55
' Propósito      : Ejecutar el proceso
'------------------------------------------------------------------------------*
Private Sub cmdEjecutar_Click()
  On Error GoTo cmdEjecutar_Click_Error
    '
    '   Si la información del formulario es válida
    '
    If IsValid Then
        Me.Tag = EJECUTAR
        Me.Hide
    Else
        MsgBox m_sMensaje, vbOKOnly + vbExclamation, "Formulario de Selección de Metodos"
    End If
  
  On Error GoTo 0
cmdEjecutar_Click__CleanExit:
    Exit Sub
cmdEjecutar_Click_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    Call HandleException(ErrNumber, ErrDescription, "frmMetodoSelectView.cmdEjecutar_Click", ErrSource)
    Call MsgBox(ErrDescription, vbCritical Or vbSystemModal, ThisWorkbook.Name)
End Sub

'------------------------------------------------------------------------------*
' Procedimiento  : cmdGoFirst_Click
' Fecha          : ma., 31/mar/2020 19:54:32
' Propósito      : Ir a la primera página
'------------------------------------------------------------------------------*
Private Sub cmdGoFirst_Click()
    Dim mModel As MetodoModel
  On Error GoTo cmdGoFirst_Click_Error
    '
    '   Ir a la página primera
    '
    Set mModel = mCtrl.GoPageNumber(1)
    '
    '
    '
    Refresh mModel
  On Error GoTo 0
cmdGoFirst_Click__CleanExit:
    Exit Sub
cmdGoFirst_Click_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    Call HandleException(ErrNumber, ErrDescription, "frmMetodoSelectView.cmdGoFirst_Click", ErrSource)
    Call MsgBox(ErrDescription, vbCritical Or vbSystemModal, ThisWorkbook.Name)
End Sub

'------------------------------------------------------------------------------*
' Procedimiento  : cmdGoLast_Click
' Fecha          : ma., 31/mar/2020 19:53:57
' Propósito      : Ir a la última página
'------------------------------------------------------------------------------*
Private Sub cmdGoLast_Click()
    Dim mModel As MetodoModel
  On Error GoTo cmdGoLast_Click_Error
    '
    '   Ir a la ultima página
    '
    Set mModel = mCtrl.GoPageNumber(mTotalPaginas)
    '
    '
    '
    Refresh mModel
  On Error GoTo 0
cmdGoLast_Click__CleanExit:
    Exit Sub
cmdGoLast_Click_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    Call HandleException(ErrNumber, ErrDescription, "frmMetodoSelectView.cmdGoLast_Click", ErrSource)
    Call MsgBox(ErrDescription, vbCritical Or vbSystemModal, ThisWorkbook.Name)
End Sub

'------------------------------------------------------------------------------*
' Procedimiento  : cmdNextPage_Click
' Fecha          : ma., 31/mar/2020 20:30:45
' Propósito      : Ir a la siguiente página
'------------------------------------------------------------------------------*
Private Sub cmdNextPage_Click()
    Dim mModel As MetodoModel
  On Error GoTo cmdNextPage_Click_Error
    '
    '   Ir a la pagina siguiente
    '
    Set mModel = mCtrl.GoPageNumber(mPaginaActual + 1)
    '
    '
    '
    Refresh mModel
  On Error GoTo 0
cmdNextPage_Click__CleanExit:
    Exit Sub
cmdNextPage_Click_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    Call HandleException(ErrNumber, ErrDescription, "frmMetodoSelectView.cmdNextPage_Click", ErrSource)
    Call MsgBox(ErrDescription, vbCritical Or vbSystemModal, ThisWorkbook.Name)
End Sub

'------------------------------------------------------------------------------*
' Procedimiento  : cmdPrevPage_Click
' Fecha          : ma., 31/mar/2020 19:53:19
' Propósito      :
'------------------------------------------------------------------------------*
Private Sub cmdPrevPage_Click()
    Dim mModel As MetodoModel
  On Error GoTo cmdPrevPage_Click_Error
    '
    '   Ir a la pagina anterior
    '
    Set mModel = mCtrl.GoPageNumber(mPaginaActual - 1)
    '
    '
    '
    Refresh mModel
  On Error GoTo 0
cmdPrevPage_Click__CleanExit:
    Exit Sub
cmdPrevPage_Click_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    Call HandleException(ErrNumber, ErrDescription, "frmMetodoSelectView.cmdPrevPage_Click", ErrSource)
    Call MsgBox(ErrDescription, vbCritical Or vbSystemModal, ThisWorkbook.Name)
End Sub

'------------------------------------------------------------------------------*
' Procedimiento  : txtFechaAnalisis_Change
' Fecha          : ma., 31/mar/2020 19:53:19
' Propósito      :
'------------------------------------------------------------------------------*
Private Sub txtFechaAnalisis_Change()
    If Not IsDate(txtFechaAnalisis.Text) Then
        txtFechaAnalisis.BackColor = RGB(255, 252, 162)    ' amarillo claro
    Else
        txtFechaAnalisis.BackColor = RGB(255, 255, 255)    'Blanco
    End If
End Sub

'------------------------------------------------------------------------------*
' Procedimiento  : txtPronosticos
' Fecha          : ma., 24/mar/2020 00:30:36
' Propósito      : Comprueba los pronosticos
'------------------------------------------------------------------------------*
Private Sub txtPronosticos_Change()
    If Not IsNumeric(txtPronosticos.Text) Then
        txtPronosticos.BackColor = RGB(255, 252, 162)     ' Amarillo claro
    Else
        mPronosticos = CInt(txtPronosticos.Text)
        If (mPronosticos >= 5 And mPronosticos <= 11) Then
            txtPronosticos.BackColor = RGB(255, 255, 255)     ' Blanco
        Else
            txtPronosticos.BackColor = RGB(255, 252, 162)     ' Amarillo claro
        End If
    End If
End Sub

'------------------------------------------------------------------------------*
' Procedimiento  : lstMetodos_DblClick
' Fecha          : ma., 31/mar/2020 20:31:18
' Propósito      : Seleccionar un método
' Parámetros     :
'------------------------------------------------------------------------------*
Private Sub lstMetodos_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    '
    '   Invocamos al botón editar si hay registros
    '
    If mTotalRecords > 0 Then
        cmdEditar_Click
    End If
End Sub



'--- Métodos Públicos ---------------------------------------------------------*
'------------------------------------------------------------------------------*
' Procedimiento  : InitForm
' Fecha          : ma., 31/mar/2020 19:24:22
' Propósito      : Inicializa los controles del informe
'------------------------------------------------------------------------------*
Public Sub InitForm()
    Dim mVar        As Variant
    Dim mHoy        As Date
    Dim DB          As BdDatos
  On Error GoTo InitForm_Error
    '
    '   Inicializamos pronosticos a 6
    '   TODO: en función del juego seleccionar 5 para euromillon y el gordo
    '
    mPronosticos = 6
    '
    '   Cargamos el Combo de juegos
    '
    mVar = Split(NOMBRE_JUEGOS, ";")
    For i = 0 To UBound(mVar)
        cboJuegos.AddItem mVar(i)
    Next i
    '
    '   Establecemos el juego por defecto
    '
    cboJuegos.ListIndex = JUEGO_DEFECTO - 1
    '
    '   Establecemos la próxima fecha de sugerencia
    '
    mHoy = Date
    Set mInfo = New InfoSorteo
    mInfo.Constructor JUEGO_DEFECTO
    
    Set DB = New BdDatos
    If mHoy = DB.UltimoResultado Then
            mFechaAnalisis = mInfo.GetProximoSorteo(mHoy)
    Else
        If mInfo.EsFechaSorteo(mHoy) Then
            mFechaAnalisis = mHoy
        Else
            mFechaAnalisis = mInfo.GetProximoSorteo(mHoy)
        End If
    End If
    '
    '   Configurar el listBox
    '
    With lstMetodos
        .ColumnHeads = False
        .ColumnWidths = "20;40"
        .MultiSelect = fmMultiSelectMulti
        .ListStyle = fmListStyleOption
        .BoundColumn = 1  ' Id
    End With
    
  On Error GoTo 0
InitForm__CleanExit:
    Exit Sub
InitForm_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    Call HandleException(ErrNumber, ErrDescription, "frmMetodoSelectView.InitForm", ErrSource)
    Call MsgBox(ErrDescription, vbCritical Or vbSystemModal, ThisWorkbook.Name)
End Sub




'------------------------------------------------------------------------------*
' Procedimiento  : Refresh
' Fecha          : ma., 31/mar/2020 19:24:22
' Propósito      : Inicializa los controles del informe
' Parámetros     : Modelo de Metodo
'------------------------------------------------------------------------------*
Public Sub Refresh(datModel As MetodoModel)
    Dim mMtd            As Metodo
    Dim mTmp            As Variant
    Dim i               As Integer
  On Error GoTo Refresh_Error
    '
    '   Actualizamos datos de registros
    '
    mTotalPaginas = datModel.TotalPages
    mPaginaActual = datModel.CurrentPage
    mTotalRecords = datModel.TotalRecords
    
    
    '
    '   Fijamos el control para que no se modifique
    '
    cboJuegos.Enabled = False
    '
    '   Actualizamos fecha
    '
    txtFechaAnalisis.Text = Format(mFechaAnalisis, "dd/mm/yyyy")
    '
    '   Actualizamos pronosticos
    '
    txtPronosticos.Text = Format(mPronosticos, "#0")
    '
    '   vaciamos la rejilla
    '
    lstMetodos.Clear
    '
    '   Cargar lista
    '
    i = 0
    If datModel.TotalRecords > 0 Then
        '
        '   Cargamos rejilla
        '
        For Each mMtd In datModel.Metodos.Items
            '
            '   Agregamos el Código al Index
            '
            lstMetodos.AddItem mMtd.Id
            '
            '   Agregamos la columna 2
            '
            lstMetodos.List(i, 1) = mMtd.ToString
            '
            '   Comprobar si esta seleccionado y marcarlo
            '
            lstMetodos.Selected(i) = ItemSelected(CStr(mMtd.Id)) Or mAllMetodos
            '
            i = i + 1
            Set mMtd = Nothing
        Next mMtd
        '
        '   Actualizar página
        '
        mTmp = Replace(LT_LABELPAGINAS, "#1", mPaginaActual)
        mTmp = Replace(mTmp, "#2", mTotalPaginas)
    Else
        mTmp = ""
        lstMetodos.AddItem 0
        lstMetodos.List(i, 1) = LT_NOMETODOS
    End If
    lblPagina = mTmp
    '
    '   Configurar editar
    '
    If mTotalRecords = 0 Then
        '
        '   Edición desactivada
        '
        cmdEditar.Enabled = False
    Else
        '
        '   Edición desactivada
        '
        cmdEditar.Enabled = True
    End If
    '
    '   Botones de navegación activar o desactivar
    '   Evaluar totalpaginas y pagina actual
    '
    If mTotalPaginas = mPaginaActual Then
        cmdNextPage.Enabled = False
        cmdGoLast.Enabled = False
    Else
        cmdNextPage.Enabled = True
        cmdGoLast.Enabled = True
    End If
    '
    If mPaginaActual = 1 Then
        cmdGoFirst.Enabled = False
        cmdPrevPage.Enabled = False
    Else
        cmdGoFirst.Enabled = True
        cmdPrevPage.Enabled = True
    End If
  On Error GoTo 0
Refresh__CleanExit:
    Exit Sub
Refresh_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    Call HandleException(ErrNumber, ErrDescription, "frmMetodoSelectView.Refresh", ErrSource)
    Call MsgBox(ErrDescription, vbCritical Or vbSystemModal, ThisWorkbook.Name)
End Sub

'------------------------------------------------------------------------------*
' Procedimiento  : IsValid
' Fecha          : ma., 31/mar/2020 19:34:13
' Propósito      : Valida los datos del formulario para la ejecución
'------------------------------------------------------------------------------*
'
Public Function IsValid() As Boolean
    Dim m_bError As Boolean    'Variable de control de al menos un error
    
  On Error GoTo IsValid_Error
    m_bError = True            'No existen errores por defecto
    m_sMensaje = ""
    Set mInfo = New InfoSorteo
    mInfo.Constructor JUEGO_DEFECTO
    '
    '       Validar fecha de sugerencia: requerida, valida y del juego
    '
    If (Len(txtFechaAnalisis.Text) = 0) Then
        m_sMensaje = m_sMensaje & MSG_NODATE
        m_bError = False
    ElseIf (Not IsDate(txtFechaAnalisis.Text)) Then
        m_sMensaje = m_sMensaje & MSG_DATENOVALID
        m_bError = False
    ElseIf Not (mInfo.EsFechaSorteo(txtFechaAnalisis.Text)) Then
        m_sMensaje = m_sMensaje & MSG_NODATESORTEO
        m_bError = False
    End If
    '
    '       Validar pronosticos
    '
    If (Len(txtPronosticos.Text) = 0) Then
        m_sMensaje = m_sMensaje & MSG_NONUMERO
        m_bError = False
    ElseIf (Not IsNumeric(txtPronosticos.Text)) Then
        m_sMensaje = m_sMensaje & MSG_NUMERONOVALID
        m_bError = False
    ElseIf (CInt(txtPronosticos.Text) < 5) _
    Or (CInt(txtPronosticos.Text) > 11) Then
        m_sMensaje = m_sMensaje & MSG_NUMOUTRANGE
        m_bError = False
    End If
    '
    '       Validar que se ha seleccionado algún método
    '
    If (mCol.Count = 0) And (Not mAllMetodos) Then
        m_sMensaje = m_sMensaje & MSG_NOMETODOS
        m_bError = False
    End If
    '
    '   Configuramos el mensaje
    '
    If (Not m_bError) Then
        m_sMensaje = MSG_HAYERROR & m_sMensaje
    End If
    '
    '   Devolvemos el resultado
    '
    IsValid = m_bError
    
    
  On Error GoTo 0
IsValid__CleanExit:
    Exit Function
IsValid_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    Call HandleException(ErrNumber, ErrDescription, "frmMetodoSelectView.IsValid", ErrSource)
    Call MsgBox(ErrDescription, vbCritical Or vbSystemModal, ThisWorkbook.Name)
End Function



'------------------------------------------------------------------------------*
' Procedimiento  : SelectRegistro
' Fecha          :
' Propósito      : Selecciona el Id del registro para la colección
'------------------------------------------------------------------------------*
Private Sub SelectRegistro()
    Dim mKey As String
    Dim mIsIncollection As Boolean
    Dim i As Integer
    Dim mIndex As Integer
    '
    '   obtiene la clave
    '
    mKey = CStr(mCurrentId)
    '
    '   Buscar si está en la colección,
    '
    mIndex = 0
    mIsIncollection = False
    For i = 1 To mCol.Count
        If mCol.Item(i) = mKey Then
            mIsIncollection = True
            mIndex = i
            Exit For
        End If
    Next i
    '   mCurrentId
    '   mCurrentSelected
    '   Si esta en la colección y ha sido deseleccionado
    '       eliminarlo de la colección
    '
    If mIsIncollection And Not mCurrentSelected Then
        mCol.Remove mIndex
    End If
    '   Si no esta en la colección y ha sido seleccionado
    '       agregarlo a la colección
    If Not mIsIncollection And mCurrentSelected Then
        If mCol.Count = 0 Then
            mCol.Add mKey
        Else
            mCol.Add mKey, , mCol.Count
        End If
    End If
End Sub

'------------------------------------------------------------------------------*
' Procedimiento  : ItemSelected
' Fecha          : ma., 21/abr/2020 19:47:31
' Propósito      : Indica si un elemento está seleccionado
'------------------------------------------------------------------------------*
Private Function ItemSelected(datKey As String) As Boolean
    Dim i As Integer
    Dim mIndex As Integer
    
    ItemSelected = False
    
    For i = 1 To mCol.Count
        If mCol.Item(i) = datKey Then
            ItemSelected = True
            Exit For
        End If
    Next i
End Function
'' *===========(EOF): frmMetodoSelectView.frm
