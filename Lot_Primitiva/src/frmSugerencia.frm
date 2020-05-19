VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSugerencia 
   Caption         =   "Formulario de Sugerencia"
   ClientHeight    =   2430
   ClientLeft      =   45
   ClientTop       =   495
   ClientWidth     =   4230
   OleObjectBlob   =   "frmSugerencia.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "frmSugerencia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




Option Explicit
Private DB          As New BdDatos           ' Base de datos

Dim m_ini_DataBase  As Date                  ' Fecha inicial de la base de datos
Dim m_fin_DataBase  As Date                  ' Fecha final de la base de datos
Dim mPeriodo        As Periodo               ' Objeto que facilita el manejo de dos fechas
Dim mParametros     As ParametrosSimulacion  ' Parámetros de simulacion
'Dim mParaMetodo     As ParametrosMetodo      ' Parámetros del método
Dim mMetodo         As metodo                ' Metodo de trabajo
'Dim mModo           As Modo                  ' Modo del formulario
'Dim mMensaje        As String                ' Mensajes de validación

Public Property Get Parametros() As ParametrosSimulacion
    Set Parametros = mParametros
End Property

Private Sub cboPerFecha_Change()
     
     mPeriodo.Tipo_Fecha = cboPerFecha.ListIndex   ' Actualizamos el tipo de periodo
                                                     ' seleccionado
     If mPeriodo.Tipo_Fecha <> 0 Then                ' Si el periodo no es personalizado
                                                     ' Formatea la fecha final y la
                                                     ' coloca en el correspondiente
                                                     ' caja de texto
        txtFecha.Text = Format(mPeriodo.FechaFinal, "dd/mm/yyyy")

        txtFecha.Enabled = False
    Else
        txtFecha.Enabled = True
    End If
   
End Sub

Private Sub txtFecha_Change()
    If mPeriodo Is Nothing Then Exit Sub
    If IsDate(txtFecha.Text) Then
        mPeriodo.FechaFinal = CDate(txtFecha.Text)
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
 
    m_ini_DataBase = DB.PrimerResultado         ' Obtiene la menor de las fechas de la
                                                ' base de datos
    m_fin_DataBase = DB.UltimoResultado         ' Obtiene la mayor de las fechas de la
                                                ' base de datos
                                                ' Formatea y asigna la fecha final
    txtFecha.Text = Format(m_fin_DataBase, "dd/mm/yyyy")
                                                
    Set mPeriodo = New Periodo                  ' Inicializa el objeto periodo de análisis
    Set mMetodo = New metodo                    ' Crea el método de trabajo
    Set mParametros = New ParametrosSimulacion  ' Crea el metodo de simulación
    
    mPeriodo.Init m_ini_DataBase, m_fin_DataBase
    mPeriodo.CargaTabla cboPerFecha             ' Carga el combo con los periodos
                                                ' predefinidos
    LoadParametros
                                                
    cboPerFecha.ListIndex = mPeriodo.Tipo_Fecha
    
    Set mParametros.RangoAnalisis = mPeriodo    ' Asigna el periodo inicial
    txtPronosticos.Text = CStr(mParametros.Pronosticos) '

    Cargar_Metodos                              ' Carga los metodos desde un rango
    
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
'    Me.Tag = SIMULAR_METODOS
'    Me.Hide
'End Sub


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
    Me.Tag = EJECUTAR
    Me.Hide
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
'        Dim i As Integer
    
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
' *     Procedimiento       : Guardar_Metodos ( Sub )
' *     Version             : 1.0 vie, 13/06/2008 09:46
' *     Autor               : CAB3780Y
' *     Retorno             :
' *     Parámetros          : <nombre_par1>     (I)]
' *     Objetivo            :
' *
' *============================================================================*
'Private Sub Guardar_Metodos()
'        Dim mRango As Range, mMtd As ParametrosMetodo
'        Dim mCols As Integer, mFils As Integer
'        Dim i As Integer
'
'    On Error Resume Next
'        If IsEmpty(THISLIBRO) Or (THISLIBRO = "") Then
'            THISLIBRO = ActiveWorkbook.Name
'        End If
'
'        Set mRango = Workbooks(THISLIBRO).Worksheets("Variables").Range("F3").CurrentRegion
'        mCols = mRango.Columns.Count
'        mFils = mRango.Rows.Count
'        If (mFils > 2) Then
'            mFils = mFils - 2
'            Set mRango = mRango.Offset(2, 0).Resize(mFils, mCols)
'            mRango.Clear
'            Set mRango = mRango.Resize(1, 1)
'        Else
'            Set mRango = mRango.Offset(2, 0).Resize(1, 1)
'        End If
'
'        i = 0
'        For Each mMtd In mParametros.Metodos
'             mRango.Offset(i, 0).Value = mMtd.Id
'             mRango.Offset(i, 1).Value = mMtd.Ordenacion
'             mRango.Offset(i, 2).Value = mMetodo.LiteralesMetodos(mMtd.Ordenacion)
'             mRango.Offset(i, 3).Value = mMtd.DiasMuestra
'             mRango.Offset(i, 4).Value = mMtd.DiasRetardo
'             i = i + 1
'        Next mMtd
'
'End Sub

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


