VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmMetodos 
   Caption         =   "Proceso de Simulacion de Varios Métodos"
   ClientHeight    =   4725
   ClientLeft      =   30
   ClientTop       =   330
   ClientWidth     =   7185
   OleObjectBlob   =   "frmMetodos.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "frmMetodos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False





' *============================================================================*
' *
' *     Fichero    : frmMetodos
' *
' *     Tipo       : Formulario
' *     Autor      : CAB3780Y
' *     Creacion   : vie, 13/06/2008  09:42
' *     Version    : 1.0
' *     Asunto     : Formulario de simulación de varios métodos de Sugerencia
' *
' *============================================================================*
Private DB          As New BdDatos           ' Base de datos

Enum Modo
    Consulta = 0                             ' Modo de actuación consulta, solo visualiza
    alta = 1                                 ' Modo de Agregar registros a la lista
    Modificacion = 2                         ' Modo para modificar un elemento de la lista
End Enum

Dim m_ini_DataBase  As Date                  ' Fecha inicial de la base de datos
Dim m_fin_DataBase  As Date                  ' Fecha final de la base de datos
Dim mPeriodo        As Periodo               ' Objeto que facilita el manejo de dos fechas
Dim mParametros     As ParametrosSimulacion  ' Parámetros de simulacion
Dim mParaMetodo     As ParametrosMetodoOld   ' Parámetros del método
Dim mMetodo         As MetodoOld             ' Metodo de trabajo
Dim mModo           As Modo                  ' Modo del formulario
Dim mMensaje        As String                ' Mensajes de validación

' *============================================================================*
' *     Procedimiento       : Parametros ( Property )
' *     Version             : 1.0 vie, 13/06/2008 09:44
' *     Autor               : CAB3780Y
' *     Retorno             : ParametrosSimulacion
' *     Parámetros          : <nombre_par1>     (I)]
' *     Objetivo            :
' *
' *============================================================================*
Public Property Get Parametros() As ParametrosSimulacion
    Set Parametros = mParametros
End Property

Private Sub txtFechaMuestraFin_Change()
    If mPeriodo Is Nothing Then Exit Sub
    If IsDate(txtFechaMuestraFin.Text) Then
        mPeriodo.FechaFinal = CDate(txtFechaMuestraFin.Text)
    End If
    
End Sub

Private Sub txtFechaMuestraIni_Change()
    If mPeriodo Is Nothing Then Exit Sub
    If IsDate(txtFechaMuestraIni.Text) Then
        mPeriodo.FechaInicial = CDate(txtFechaMuestraIni.Text)
    End If
   
End Sub

' *============================================================================*
' *     Procedimiento       : UserForm_Initialize ( Sub )
' *     Version             : 1.0 vie, 13/06/2008 09:43
' *     Autor               : CAB3780Y
' *     Retorno             :
' *     Parámetros          : <nombre_par1>     (I)]
' *     Objetivo            :
' *
' *============================================================================*
Private Sub UserForm_Initialize()
   
   On Error GoTo UserForm_Initialize_Error
    
    m_ini_DataBase = DB.PrimerResultado         ' Obtiene la menor de las fechas de la
                                                ' base de datos
    m_fin_DataBase = DB.UltimoResultado         ' Obtiene la mayor de las fechas de la
                                                ' base de datos
                                                ' Formatea y asigna la fecha inicial
    txtFechaMuestraIni.Text = Format(m_ini_DataBase, "dd/mm/yyyy")
                                                ' Formatea y asigna la fecha final
    txtFechaMuestraFin.Text = Format(m_fin_DataBase, "dd/mm/yyyy")
                                                
    Set mPeriodo = New Periodo                  ' Inicializa el objeto periodo de análisis
    Set mMetodo = New metodo                    ' Crea el método de trabajo
    Set mParametros = New ParametrosSimulacion  ' Crea el metodo de simulación
    
    mPeriodo.Init m_ini_DataBase, m_fin_DataBase
    mPeriodo.CargaTabla cboPerMuestra           ' Carga el combo con los periodos
                                                ' predefinidos
    LoadParametros
                                                
    cboPerMuestra.ListIndex = mPeriodo.Tipo_Fecha
    
'    cboPerMuestra.ListIndex = ctLoQueVadeSemana    ' Selecciona la opción Lo que va de mes
    mMetodo.CargaTabla Me.cboTipoOrdenacion
    Set mParametros.RangoAnalisis = mPeriodo    ' Asigna el periodo inicial
'    mParametros.Pronosticos = 7                 ' Asigna los pronosticos por defecto
    
    txtPronosticos.Text = CStr(mParametros.Pronosticos) '

    Cargar_Metodos                              ' Carga los metodos desde un rango
    
    Actualiza_Lista                             ' Actualiza la listbox
    
    MultiPage1.Value = 0                        ' Se posiciona en la página 0
   
   On Error GoTo 0
   Exit Sub
   
UserForm_Initialize_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & _
    ") in procedure UserForm_Initialize of Formulario frmMetodos"
End Sub

' *============================================================================*
' *     Procedimiento       : cboPerMuestra_Change ( Sub )
' *     Version             : 1.0 vie, 13/06/2008 09:45
' *     Autor               : CAB3780Y
' *     Retorno             :
' *     Parámetros          : <nombre_par1>     (I)]
' *     Objetivo            :
' *
' *============================================================================*
Private Sub cboPerMuestra_Change()
   
   On Error GoTo cboPerMuestra_Change_Error
    
     mPeriodo.Tipo_Fecha = cboPerMuestra.ListIndex   ' Actualizamos el tipo de periodo
                                                     ' seleccionado
'    If mPeriodo.Tipo_Fecha <> 0 Then                ' Si el periodo no es personalizado
'                                                    ' Si la fecha inicial no es menor
'                                                    ' que la base de datos
'        If (mPeriodo.FechaInicial < m_ini_DataBase) Then
'                                                    ' Selecciona la fecha menor de la
'                                                    ' base de datos
'            txtFechaMuestraIni.Text = Format(m_ini_DataBase, "dd/mm/yyyy")
'        Else
'                                                    ' Formatea la fecha inicial y la
'                                                    ' la coloca en la correspondiente caja
'                                                    ' de texto
'            txtFechaMuestraIni.Text = Format(mPeriodo.FechaInicial, "dd/mm/yyyy")
'        End If
'                                                    ' Si la fecha final es superior a
'                                                    ' la fecha de la base de datos
'        If (mPeriodo.FechaFinal > m_fin_DataBase) Then
'                                                    ' Selecciona la fecha mayor de la base
'                                                    ' de datos
'            mPeriodo.FechaFinal = m_fin_DataBase
'            txtFechaMuestraFin.Text = Format(m_fin_DataBase, "dd/mm/yyyy")
'        Else
'                                                    ' Formatea la fecha final y la
'                                                    ' coloca en el correspondiente
'                                                    ' caja de texto
'            txtFechaMuestraFin.Text = Format(mPeriodo.FechaFinal, "dd/mm/yyyy")
'        End If
'    mPeriodo.Tipo_Fecha = cboPerMuestra.ListIndex   ' Actualizamos el tipo de periodo
'                                                    ' seleccionado
     If mPeriodo.Tipo_Fecha <> 0 Then                ' Si el periodo no es personalizado
                                                     ' Formatea la fecha inicial y la
                                                     ' la coloca en la correspondiente caja
                                                     ' de texto
        txtFechaMuestraIni.Text = Format(mPeriodo.FechaInicial, "dd/mm/yyyy")
                                                     ' Formatea la fecha final y la
                                                     ' coloca en el correspondiente
                                                     ' caja de texto
        txtFechaMuestraFin.Text = Format(mPeriodo.FechaFinal, "dd/mm/yyyy")

        txtFechaMuestraIni.Enabled = False
        txtFechaMuestraFin.Enabled = False
    Else
        txtFechaMuestraIni.Enabled = True
        txtFechaMuestraFin.Enabled = True
    End If
    
    txtDiasMuestra.Text = Format(mPeriodo.Dias, "#,##0")
    txtDiasMuestra.Enabled = False
   
   On Error GoTo 0
   Exit Sub

cboPerMuestra_Change_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & _
    ") in procedure cboPeriodo_Change of Formulario frmProbDosNumeros"
End Sub

' *============================================================================*
' *     Procedimiento       : MultiPage1_Change ( Sub )
' *     Version             : 1.0 vie, 13/06/2008 09:44
' *     Autor               : CAB3780Y
' *     Retorno             :
' *     Parámetros          : <nombre_par1>     (I)]
' *     Objetivo            :
' *
' *============================================================================*
Private Sub MultiPage1_Change()
    If (MultiPage1.Value = 1) Then
        Configura_Modo
    End If
End Sub

'Private Sub btnSugerir_Click()
'    If (Not IsNumeric(txtPronosticos.Text)) Then
'        MsgBox "No ha introducido pronosticos Numérico", vbExclamation + vbOKCancel, Me.Caption
'        txtPronosticos.SetFocus
'        Exit Sub
'    Else
'        mParametros.Pronosticos = CInt(txtPronosticos.Text)
'    End If
'    SaveParametros
'    Guardar_Metodos
'    Me.Tag = SIMULAR_METODOS
'    Me.Hide
'End Sub

Private Sub cmdEliminar_Click()
    If mModo = Consulta Then
        If mParametros.NumMetodos > 0 Then
            mParametros.Remove mParaMetodo
            Actualiza_Lista
            Configura_Modo
        End If
    Else
        mModo = Consulta
        Actualiza_Lista
        Configura_Modo
    End If
End Sub

' *============================================================================*
' *     Procedimiento       : lstConsulta_Click ( Sub )
' *     Version             : 1.0 vie, 13/06/2008 09:44
' *     Autor               : CAB3780Y
' *     Retorno             :
' *     Parámetros          : <nombre_par1>     (I)]
' *     Objetivo            :
' *
' *============================================================================*
Private Sub lstConsulta_Click()
    Dim IdParametro As Integer
    
    IdParametro = lstConsulta.ListIndex + 1         ' Localiza el elemento
                                                    ' Si no tenemos definido
                                                    ' los parámetros del método
                                                    ' se crea
    DisplayMetodo IdParametro
    
End Sub

' *============================================================================*
' *     Procedimiento       : lstConsulta_DblClick ( Sub )
' *     Version             : 1.0 vie, 13/06/2008 09:44
' *     Autor               : CAB3780Y
' *     Retorno             :
' *     Parámetros          : <nombre_par1>     (I)]
' *     Objetivo            :
' *
' *============================================================================*
Private Sub lstConsulta_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Dim IdParametro As Integer
    
    IdParametro = lstConsulta.ListIndex + 1
    
    DisplayMetodo IdParametro
    
    mModo = Modificacion
    Configura_Modo
End Sub

' *============================================================================*
' *     Procedimiento       : cmdAgregar_Click ( Sub )
' *     Version             : 1.0 vie, 13/06/2008 09:44
' *     Autor               : CAB3780Y
' *     Retorno             :
' *     Parámetros          : <nombre_par1>     (I)]
' *     Objetivo            :
' *
' *============================================================================*
Private Sub cmdAgregar_Click()
    If mModo = Consulta Then
        mModo = alta
        Configura_Modo
        Set mParaMetodo = Parametros.GetNewMetodo
    Else
        If (IsValid) Then
            If mModo = alta Then
                mParaMetodo.Ordenacion = cboTipoOrdenacion.ListIndex
                mParaMetodo.DiasMuestra = CInt(txtDiasAnalisis.Text)
                mParaMetodo.DiasRetardo = CInt(txtDiasRetardo.Text)
                mParametros.Add mParaMetodo
            Else
                mParaMetodo.Ordenacion = cboTipoOrdenacion.ListIndex
                mParaMetodo.DiasMuestra = CInt(txtDiasAnalisis.Text)
                mParaMetodo.DiasRetardo = CInt(txtDiasRetardo.Text)
            End If
            mModo = Consulta
            Actualiza_Lista
            Configura_Modo
        Else
            MsgBox mMensaje, vbError + vbOKCancel, Me.Caption
        End If
    End If
End Sub

' *============================================================================*
' *     Procedimiento       : cmdCancelar_Click ( Sub )
' *     Version             : 1.0 vie, 13/06/2008 09:44
' *     Autor               : CAB3780Y
' *     Retorno             :
' *     Parámetros          : <nombre_par1>     (I)]
' *     Objetivo            :
' *
' *============================================================================*
Private Sub cmdCancelar_Click()
    '   Preguntar si guardar metodos
    '
    '
    Me.Tag = BOTON_CERRAR
    Me.Hide
End Sub

' *============================================================================*
' *     Procedimiento       : cmdEjecutar_Click ( Sub )
' *     Version             : 1.0 vie, 13/06/2008 09:44
' *     Autor               : CAB3780Y
' *     Retorno             :
' *     Parámetros          : <nombre_par1>     (I)]
' *     Objetivo            :
' *
' *============================================================================*
Private Sub cmdEjecutar_Click()
    If (Not IsNumeric(txtPronosticos.Text)) Then
        MsgBox "No ha introducido pronosticos Numérico", vbExclamation + vbOKCancel, Me.Caption
        txtPronosticos.SetFocus
        Exit Sub
    Else
        mParametros.Pronosticos = CInt(txtPronosticos.Text)
    End If
    
    SaveParametros
    Guardar_Metodos
    Me.Tag = EJECUTAR
    Me.Hide
End Sub

' *============================================================================*
' *     Procedimiento       : Actualiza_Lista ( Sub )
' *     Version             : 1.0 vie, 13/06/2008 09:45
' *     Autor               : CAB3780Y
' *     Retorno             :
' *     Parámetros          : <nombre_par1>     (I)]
' *     Objetivo            :
' *
' *============================================================================*
Private Sub Actualiza_Lista()
    Dim mMtd As ParametrosMetodo
    lstConsulta.Clear
    For Each mMtd In mParametros.Metodos
        lstConsulta.AddItem mMtd.ToString
    Next mMtd
End Sub

' *============================================================================*
' *     Procedimiento       : Configura_Modo ( Sub )
' *     Version             : 1.0 vie, 13/06/2008 09:45
' *     Autor               : CAB3780Y
' *     Retorno             :
' *     Parámetros          : <nombre_par1>     (I)]
' *     Objetivo            :
' *
' *============================================================================*
Private Sub Configura_Modo()
    If IsNull(mModo) Then
        mModo = Consulta
    End If
    If (lstConsulta.ListCount = 0) Then
       mModo = alta
    End If
   
    Select Case mModo
        Case Consulta
            Me.frmDetalle.Enabled = False
            Me.cmdAgregar.Enabled = True
            Me.cmdAgregar.Caption = "Agregar"
            Me.cmdEliminar.Caption = "Eliminar"
            InitCampos
            lstConsulta.SetFocus
            
        Case alta
            InitCampos
            Me.txtId.Text = "##"
            Me.frmDetalle.Enabled = True
            Me.cmdAgregar.Caption = "Guardar"
            Me.cmdEliminar.Caption = "Cancelar"
            Me.cboTipoOrdenacion.SetFocus
       
        Case Modificacion
            Me.frmDetalle.Enabled = True
            Me.cmdAgregar.Caption = "Guardar"
            Me.cmdEliminar.Caption = "Cancelar"
            Me.cboTipoOrdenacion.SetFocus
        
    End Select
End Sub

' *============================================================================*
' *     Procedimiento       : InitCampos ( Sub )
' *     Version             : 1.0 vie, 13/06/2008 09:45
' *     Autor               : CAB3780Y
' *     Retorno             :
' *     Parámetros          : <nombre_par1>     (I)]
' *     Objetivo            :
' *
' *============================================================================*
Private Sub InitCampos()
    txtId.Text = ""
    txtId.Enabled = False
    cboTipoOrdenacion.ListIndex = -1
    txtDiasAnalisis.Text = ""
    txtDiasRetardo.Text = ""
End Sub

' *============================================================================*
' *     Procedimiento       : Cargar_Metodos ( Sub )
' *     Version             : 1.0 vie, 13/06/2008 09:46
' *     Autor               : CAB3780Y
' *     Retorno             :
' *     Parámetros          : <nombre_par1>     (I)]
' *     Objetivo            :
' *
' *============================================================================*
Private Sub Cargar_Metodos()
        Dim mRango As Range, mMtd As ParametrosMetodoOld, mFila As Range
        Dim mCols As Integer, mFils As Integer
    
    On Error Resume Next
        If IsEmpty(THISLIBRO) Or (THISLIBRO = "") Then
            THISLIBRO = ActiveWorkbook.Name
        End If
        Set mParametros.RangoAnalisis = mPeriodo
        mParametros.Clear
        Set mRango = Workbooks(THISLIBRO).Worksheets("Variables").Range("F1").CurrentRegion
        mCols = mRango.Columns.Count
        mFils = mRango.Rows.Count - 2
        Set mRango = mRango.Offset(2, 0).Resize(mFils, mCols)
        For Each mFila In mRango.Rows
            Set mMtd = mParametros.GetNewMetodo()
            mMtd.Ordenacion = mFila.Cells(1, 2).Value
            mMtd.DiasMuestra = mFila.Cells(1, 4).Value
            mMtd.DiasRetardo = mFila.Cells(1, 5).Value
            mParametros.Add mMtd
        Next mFila
End Sub

' *============================================================================*
' *     Procedimiento       : DisplayMetodo ( Sub )
' *     Version             : 1.0 vie, 13/06/2008 10:07
' *     Autor               : CAB3780Y
' *     Retorno             :
' *     Parámetros          : <nombre_par1>     (I)]
' *     Objetivo            :
' *
' *============================================================================*
Private Sub DisplayMetodo(Index As Integer)
    
    If (IsNull(mParaMetodo)) Then
       Set mParamMetodos = New ParametrosMetodo
    End If
                                                    ' Lee el parámetro de la
                                                    ' colección
    Set mParaMetodo = mParametros.Metodos(Index)
                                                    ' Mueve datos a pantalla
    txtId.Text = CStr(mParaMetodo.Id)
    cboTipoOrdenacion.ListIndex = CInt(mParaMetodo.Ordenacion)
    txtDiasAnalisis.Text = CStr(mParaMetodo.DiasMuestra)
    txtDiasRetardo.Text = CStr(mParaMetodo.DiasRetardo)
End Sub

' *============================================================================*
' *     Procedimiento       : Guardar_Metodos ( Sub )
' *     Version             : 1.0 vie, 13/06/2008 09:46
' *     Autor               : CAB3780Y
' *     Retorno             :
' *     Parámetros          : <nombre_par1>     (I)]
' *     Objetivo            :
' *
' *============================================================================*
Private Sub Guardar_Metodos()
        Dim mRango As Range, mMtd As ParametrosMetodoOld
        Dim mCols As Integer, mFils As Integer
        Dim i As Integer
    
    On Error Resume Next
        If IsEmpty(THISLIBRO) Or (THISLIBRO = "") Then
            THISLIBRO = ActiveWorkbook.Name
        End If

        Set mRango = Workbooks(THISLIBRO).Worksheets("Variables").Range("F3").CurrentRegion
        mCols = mRango.Columns.Count
        mFils = mRango.Rows.Count
        If (mFils > 2) Then
            mFils = mFils - 2
            Set mRango = mRango.Offset(2, 0).Resize(mFils, mCols)
            mRango.Clear
            Set mRango = mRango.Resize(1, 1)
        Else
            Set mRango = mRango.Offset(2, 0).Resize(1, 1)
        End If

        i = 0
        For Each mMtd In mParametros.Metodos
             mRango.Offset(i, 0).Value = mMtd.Id
             mRango.Offset(i, 1).Value = mMtd.Ordenacion
             mRango.Offset(i, 2).Value = mMetodo.LiteralesMetodos(mMtd.Ordenacion)
             mRango.Offset(i, 3).Value = mMtd.DiasMuestra
             mRango.Offset(i, 4).Value = mMtd.DiasRetardo
             i = i + 1
        Next mMtd
        
End Sub

' *============================================================================*
' *     Procedure  : LoadParametros
' *     Fichero    : frmMetodos
' *     Autor      : Carlos Almela Baeza
' *     Creacion   : mar, 12/10/2010
' *     Asunto     :
' *============================================================================*
'
Private Sub LoadParametros()

  On Error GoTo LoadParametros_Error
    If IsEmpty(THISLIBRO) Or (THISLIBRO = "") Then
        THISLIBRO = ActiveWorkbook.Name
    End If

    With Workbooks(THISLIBRO).Worksheets("Variables")
       mPeriodo.Tipo_Fecha = .Range("B3").Value
       If mPeriodo.Tipo_Fecha = ctPersonalizadas Then
            mPeriodo.FechaInicial = .Range("B4").Value
            mPeriodo.FechaFinal = .Range("B5").Value
        End If
        mParametros.Pronosticos = .Range("B7").Value
    End With
    

LoadParametros_CleanExit:
   On Error GoTo 0
    Exit Sub

LoadParametros_Error:
    

    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    '   Informa del error
    Call MsgBox(ErrDescription, vbError Or vbSystemModal, NOMBRE_APLICACION)
End Sub

' *============================================================================*
' *     Procedure  : SaveParametros
' *     Fichero    : frmMetodos
' *     Autor      : Carlos Almela Baeza
' *     Creacion   : mar, 12/10/2010
' *     Asunto     :
' *============================================================================*
'
Private Sub SaveParametros()
  On Error GoTo SaveParametros_Error
    If IsEmpty(THISLIBRO) Or (THISLIBRO = "") Then
        THISLIBRO = ActiveWorkbook.Name
    End If

    With Workbooks(THISLIBRO).Worksheets("Variables")
        .Range("A3").Value = "Periodo Comprobacion"
        .Range("A4").Value = "Fecha Inicio"
        .Range("A5").Value = "Fecha Fin"
        .Range("A6").Value = "Dias Analisis"
        .Range("A7").Value = "Pronosticos"
        
        .Range("B3").Value = mPeriodo.Tipo_Fecha
        .Range("C3").Value = mPeriodo.Texto
        .Range("B4").Value = mPeriodo.FechaInicial
        .Range("B5").Value = mPeriodo.FechaFinal
        .Range("B6").Value = mPeriodo.Dias
        .Range("B7").Value = mParametros.Pronosticos
    End With

SaveParametros_CleanExit:
   On Error GoTo 0
    Exit Sub

SaveParametros_Error:
    

    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    '   Informa del error
    Call MsgBox(ErrDescription, vbError Or vbSystemModal, NOMBRE_APLICACION)
End Sub

' *============================================================================*
' *     Procedimiento       : IsValid ( Function )
' *     Version             : 1.0 vie, 13/06/2008 09:45
' *     Autor               : CAB3780Y
' *     Retorno             : Boolean
' *     Parámetros          : <nombre_par1>     (I)]
' *     Objetivo            :
' *
' *============================================================================*
Private Function IsValid() As Boolean
    mMensaje = ""
    Dim mError As Boolean
    mError = False
    
    If (Not IsNumeric(txtDiasMuestra.Text)) Then
        mError = True
        mMensaje = mMensaje + "No ha introducido Dias de Muestra Numérico" + vbCrLf
    End If
    
    If (Not IsNumeric(txtDiasRetardo.Text)) Then
        mError = True
        mMensaje = mMensaje + "No ha introducido Dias de Retardo Numérico" + vbCrLf
    End If
    
    If (cboTipoOrdenacion.ListIndex = -1) Then
        mError = True
        mMensaje = mMensaje + "No ha seleccionado níngún método." + vbCrLf
    End If
    
    IsValid = Not mError
End Function

