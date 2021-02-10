VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GenPullCombinacionesView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' *============================================================================*
' *
' *     Fichero    : GenPullCombinacionesView.cls
' *
' *     Autor      : Carlos Almela Baeza
' *     Creación   : ma., 15/dic/2020 23:42:33
' *     Versión    : 1.0
' *     Propósito  : Interfaz del caso de uso que obtiene y visualiza información
' *                  de la hoja de portada
' *============================================================================*
Option Explicit
Option Base 0
'
'--- Variables Privadas -------------------------------------------------------*
Private mwsUINumeros     As Worksheet
Private mwsUITotCombi    As Worksheet
Private mwsUIFiltros     As Worksheet
Private mwsUIFiltCombi   As Worksheet
Private mwsUIEvalCombi   As Worksheet
Private mwsUISorteos     As Worksheet
Private mwsUISorteoRes   As Worksheet
Private mwsUICompCombi   As Worksheet
Private mwsParametros    As Worksheet
'
Private mRg              As Range       ' Rango de trabajo para obtener info
Private mMtz             As Variant     ' Matriz de datos donde se devuele info
'
'--- Constantes ---------------------------------------------------------------*
Private Const CELDA_SUGERENCIAS As String = "B4"
Private Const CELDA_TOTNUMEROS As String = "B5"
Private Const CELDA_TOTCOMBINACIONES As String = "B6"
Private Const CELDA_COMBIGNERADAS As String = "D6"
Private Const CELDA_COMBIEVALUADAS As String = "A1"
Private Const CELDA_TIPOFILTRO As String = "B5"
Private Const CELDA_VALORFILTRO As String = "B6"
Private Const CELDA_MATRIZFILTROS As String = "B11"
Private Const CELDA_CATEGORIAFILTERS As String = "D1"
Private Const HOJA_PARAMETROS As String = "Parametros!"

'--- Mensajes -----------------------------------------------------------------*
'--- Errores ------------------------------------------------------------------*

'--- Métodos Privados ---------------------------------------------------------*
Private Sub Class_Initialize()
    Set mwsUINumeros = ThisWorkbook.Sheets("Fase001")
    Set mwsUITotCombi = ThisWorkbook.Sheets("Fase002")
    Set mwsUIFiltros = ThisWorkbook.Sheets("Fase003")
    Set mwsUIFiltCombi = ThisWorkbook.Sheets("Fase004")
    Set mwsUISorteos = ThisWorkbook.Sheets("Fase005")
    Set mwsUIEvalCombi = ThisWorkbook.Sheets("Fase006")
    Set mwsUISorteoRes = ThisWorkbook.Sheets("Fase007")
    Set mwsUICompCombi = ThisWorkbook.Sheets("Fase008")
    Set mwsParametros = ThisWorkbook.Sheets("Parametros")
End Sub

Private Sub Class_Terminate()
    Set mwsUINumeros = Nothing
    Set mwsUITotCombi = Nothing
    Set mwsUIFiltros = Nothing
    Set mwsUIFiltCombi = Nothing
    Set mwsUIEvalCombi = Nothing
    Set mwsUISorteos = Nothing
    Set mwsUISorteoRes = Nothing
    Set mwsUICompCombi = Nothing
End Sub

'--- Propiedades --------------------------------------------------------------*

'------------------------------------------------------------------------------*
' Propiedad      : NumSugerencias
' Fecha          : vi., 29/ene/2021 19:45:03
' Propósito      : Numeros que forman una sugerencia o combinación
'------------------------------------------------------------------------------*
Public Property Get NumSugerencias() As Integer
    If Len(mwsUINumeros.Range(CELDA_SUGERENCIAS).Value) > 0 Then
        NumSugerencias = CInt(mwsUINumeros.Range(CELDA_SUGERENCIAS).Value)
    Else
        NumSugerencias = 0
    End If
End Property

Public Property Let NumSugerencias(ByVal vNewValue As Integer)
    mwsUINumeros.Range(CELDA_SUGERENCIAS).Value = vNewValue
End Property

'------------------------------------------------------------------------------*
' Propiedad      : RangoSalida
' Fecha          : vi., 29/ene/2021 19:46:04
' Propósito      : Celda donde iniciar la salida del proceso de generación
'------------------------------------------------------------------------------*
Public Property Get RangoSalida() As Range
    Set RangoSalida = mwsUITotCombi.Range("A3")
End Property


'------------------------------------------------------------------------------*
' Propiedad      : RangoSalidaFiltros
' Fecha          : vi., 29/ene/2021 19:47:14
' Propósito      : Celda donde iniciar la salida del proceso de filtrado de
'                  combinaciones
'------------------------------------------------------------------------------*
Public Property Get RangoSalidaFiltros() As Range
    Set RangoSalidaFiltros = mwsUIFiltCombi.Range("A3")
End Property

'------------------------------------------------------------------------------*
' Propiedad      : RangoSalidaEvaluadas
' Fecha          : vi., 29/ene/2021 19:47:45
' Propósito      : Celda donde iniciar la salida del proceso de evaluación
'------------------------------------------------------------------------------*
Public Property Get RangoSalidaEvaluadas() As Range
    Set RangoSalidaEvaluadas = mwsUIEvalCombi.Range("A3")
End Property

'------------------------------------------------------------------------------*
' Propiedad      : RangoSalidaComprobacion
' Fecha          : mi., 03/feb/2021 17:07:58
' Propósito      : Celda donde iniciar la salida del proceso de comprobación
'------------------------------------------------------------------------------*
Public Property Get RangoSalidaComprobacion() As Range
    Set RangoSalidaComprobacion = mwsUICompCombi.Range("A3")
End Property

'------------------------------------------------------------------------------*
' Propiedad      : TotalNumeros
' Fecha          : vi., 29/ene/2021 19:48:21
' Propósito      : Total de numeros que se utilizarán para las sugerencias
'------------------------------------------------------------------------------*
Public Property Get TotalNumeros() As Integer
    If Len(mwsUINumeros.Range(CELDA_TOTNUMEROS).Value) > 0 Then
        TotalNumeros = CInt(mwsUINumeros.Range(CELDA_TOTNUMEROS).Value)
    Else
        TotalNumeros = 0
    End If
End Property

'------------------------------------------------------------------------------*
' Propiedad      : TotalCombinaciones
' Fecha          : vi., 29/ene/2021 19:48:55
' Propósito      : Numero total de combinaciones generadas ( fórmula )
'------------------------------------------------------------------------------*
Public Property Get TotalCombinaciones() As Long
    If Len(mwsUINumeros.Range(CELDA_TOTCOMBINACIONES).Value) > 0 Then
        TotalCombinaciones = CLng(mwsUINumeros.Range(CELDA_TOTCOMBINACIONES).Value)
    Else
        TotalCombinaciones = 0
    End If
End Property

'------------------------------------------------------------------------------*
' Propiedad      : CombinacionesGeneradas
' Fecha          : vi., 29/ene/2021 19:50:10
' Propósito      : Numero total de combinaciones generadas en el proceso
'------------------------------------------------------------------------------*
Public Property Get CombinacionesGeneradas() As Long
    If Len(mwsUINumeros.Range(CELDA_COMBIGNERADAS).Value) > 0 Then
        CombinacionesGeneradas = CInt(mwsUINumeros.Range(CELDA_COMBIGNERADAS).Value)
    Else
        CombinacionesGeneradas = 0
    End If
End Property

'------------------------------------------------------------------------------*
' Propiedad      : CombinacionesFiltradas
' Fecha          : vi., 29/ene/2021 19:53:23
' Propósito      : Numero total de combinaciones filtradas
'------------------------------------------------------------------------------*
Public Property Get CombinacionesFiltradas() As Long
    Set mRg = mwsUIFiltCombi.Range("A2").CurrentRegion
    If mRg Is Nothing Then
        CombinacionesFiltradas = 0
    Else
        CombinacionesFiltradas = mRg.Rows.Count - 2
    End If
End Property

'------------------------------------------------------------------------------*
' Propiedad      : CombinacionesEvaluadas
' Fecha          : vi., 29/ene/2021 19:50:40
' Propósito      : Numero total de combinaciones evaluadas
'------------------------------------------------------------------------------*
Public Property Get CombinacionesEvaluadas() As Long
    Set mRg = mwsUIEvalCombi.Range(CELDA_COMBIEVALUADAS).CurrentRegion
    If mRg Is Nothing Then
        CombinacionesEvaluadas = 0
    Else
        CombinacionesEvaluadas = mRg.Rows.Count - 2
    End If
End Property


'------------------------------------------------------------------------------*
' Propiedad      : TipoFiltro
' Fecha          : vi., 29/ene/2021 19:53:35
' Propósito      : Tipo de filtro seleccionado para la colección
'------------------------------------------------------------------------------*
Public Property Get TipoFiltro() As String
    TipoFiltro = mwsUIFiltros.Range(CELDA_TIPOFILTRO).Value
End Property

'------------------------------------------------------------------------------*
' Propiedad      : ValorFiltro
' Fecha          : vi., 29/ene/2021 19:54:04
' Propósito      : Valor del filtro seleccionado
'------------------------------------------------------------------------------*
Public Property Get ValorFiltro() As String
    ValorFiltro = mwsUIFiltros.Range(CELDA_VALORFILTRO).Value
End Property

'------------------------------------------------------------------------------*
' Propiedad      : TotalFiltros
' Fecha          : ju., 04/feb/2021 16:33:30
' Propósito      : Total de filtros aplicados
'------------------------------------------------------------------------------*
Public Property Get TotalFiltros() As Integer
    Set mRg = mwsUIFiltros.Range(CELDA_MATRIZFILTROS).CurrentRegion
    If mRg Is Nothing Then
        TotalFiltros = 0
    Else
        TotalFiltros = mRg.Rows.Count - 2
    End If
End Property


'------------------------------------------------------------------------------*
' Propiedad      : RegistrosComprobados
' Fecha          : ju., 04/feb/2021 16:33:30
' Propósito      :
'------------------------------------------------------------------------------*
Public Property Get RegistrosComprobados() As Long
    If IsEmpty(mwsUISorteoRes.Range("E6")) Then
       RegistrosComprobados = 0
    Else
       RegistrosComprobados = CLng(mwsUISorteoRes.Range("E6").Value)
    End If
End Property

Public Property Let RegistrosComprobados(vNewValue As Long)
    mwsUISorteoRes.Range("E6").Value = vNewValue
End Property

'------------------------------------------------------------------------------*
' Propiedad      : RegistrosPremiados
' Fecha          : ju., 04/feb/2021 16:33:30
' Propósito      :
'------------------------------------------------------------------------------*
Public Property Get RegistrosPremiados() As Long
    If IsEmpty(mwsUISorteoRes.Range("E7")) Then
       RegistrosPremiados = 0
    Else
       RegistrosPremiados = CLng(mwsUISorteoRes.Range("E7").Value)
    End If
End Property
Public Property Let RegistrosPremiados(vNewValue As Long)
    mwsUISorteoRes.Range("E7").Value = vNewValue
End Property

'------------------------------------------------------------------------------*
' Propiedad      : TotalCoste
' Fecha          : ju., 04/feb/2021 16:33:30
' Propósito      :
'------------------------------------------------------------------------------*
Public Property Get TotalCoste() As Currency
    If IsEmpty(mwsUISorteoRes.Range("E8")) Then
       TotalCoste = 0
    Else
       TotalCoste = CCur(mwsUISorteoRes.Range("E8").Value)
    End If
End Property
Public Property Let TotalCoste(vNewValue As Currency)
    mwsUISorteoRes.Range("E8").Value = vNewValue
End Property

'------------------------------------------------------------------------------*
' Propiedad      : TotalImporte
' Fecha          : ju., 04/feb/2021 16:33:30
' Propósito      :
'------------------------------------------------------------------------------*
Public Property Get TotalImporte() As Currency
    If IsEmpty(mwsUISorteoRes.Range("E8")) Then
       TotalImporte = 0
    Else
       TotalImporte = CCur(mwsUISorteoRes.Range("E8").Value)
    End If
End Property

Public Property Let TotalImporte(vNewValue As Currency)
    mwsUISorteoRes.Range("E9").Value = vNewValue
End Property

'--- Métodos Públicos ---------------------------------------------------------*
'------------------------------------------------------------------------------*
' Procedimiento  : ClearSalidaFiltros
' Fecha          : lu., 21/dic/2020 18:44:47
' Propósito      : Inicializar las celdas de salida para los filtros
'------------------------------------------------------------------------------*
Public Sub ClearSalidaFiltros()
    Dim rgBorra As Range
    Dim nRows  As Integer
    
  On Error GoTo ClearSalidaFiltros_Error
    Application.ScreenUpdating = False
    Set rgBorra = mwsUIFiltCombi.Range("A2").CurrentRegion
    nRows = rgBorra.Rows.Count
    '
    '   Mantenemos la cabecera
    '
    If nRows > 2 Then
        Set rgBorra = rgBorra.Offset(2, 0).Resize(nRows - 2, rgBorra.Columns.Count)
        rgBorra.ClearContents
    End If
    Application.ScreenUpdating = True
    Set rgBorra = Nothing
    
    
  On Error GoTo 0
ClearSalidaFiltros__CleanExit:
    Exit Sub
ClearSalidaFiltros_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    Call HandleException(ErrNumber, ErrDescription, "GenCombinacionesView.ClearSalidaFiltros", ErrSource)
    Err.Raise ErrNumber, "GenCombinacionesView.ClearSalidaFiltros", ErrDescription
End Sub



'------------------------------------------------------------------------------*
' Procedimiento  : ClearSalidaEvaluacion
' Fecha          : vi., 29/ene/2021 18:00:58
' Propósito      : Inicializar celdas salida para las combinaciones Evaluadas
'------------------------------------------------------------------------------*
Public Sub ClearSalidaEvaluacion()
    Dim rgBorra As Range
    Dim nRows  As Integer
    
  On Error GoTo ClearSalidaEvaluacion_Error
    Application.ScreenUpdating = False
    Set rgBorra = mwsUIEvalCombi.Range("A2").CurrentRegion
    nRows = rgBorra.Rows.Count
    '
    '   Mantenemos la cabecera
    '
    If nRows > 2 Then
        Set rgBorra = rgBorra.Offset(2, 0).Resize(nRows - 2, rgBorra.Columns.Count)
        rgBorra.ClearContents
    End If
    Application.ScreenUpdating = True
    Set rgBorra = Nothing
    
    
  On Error GoTo 0
ClearSalidaEvaluacion__CleanExit:
    Exit Sub
ClearSalidaEvaluacion_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    Call HandleException(ErrNumber, ErrDescription, "GenCombinacionesView.ClearSalidaEvaluacion", ErrSource)
    Err.Raise ErrNumber, "GenCombinacionesView.ClearSalidaEvaluacion", ErrDescription
End Sub



'------------------------------------------------------------------------------*
' Procedimiento  : GetParametrosProceso
' Fecha          : mi., 21/ago/2019 19:16:17
' Propósito      : Devuelve la información de la hoja de portada para el modelo
' Retorno        : Modelo
'------------------------------------------------------------------------------*
Public Function GetParametrosProceso() As GenPullCombinacionesModel
    Dim mMatriz As Variant
   
 On Error GoTo GetParametrosProceso_Error
    '
    '   Creamos el modelo
    '
    Set GetParametrosProceso = New GenPullCombinacionesModel
    '
    '   Asignamos Datos de la vista al modelo
    '
    With GetParametrosProceso
        .MatrizNumeros = GetNumeros
        .NumerosSugerencia = Me.NumSugerencias
        .TotalNumerosCombinar = Me.TotalNumeros
        .TotalCombinaciones = Me.TotalCombinaciones
        .Filtros = GetFiltros
        .Sorteos = GetSorteos
        .CombinacionGanadora = GetCombinacionGanadora
    End With
    '
    '
    '
  On Error GoTo 0
GetParametrosProceso__CleanExit:
    Exit Function
            
GetParametrosProceso_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    Call HandleException(ErrNumber, ErrDescription, "GenCombinacionesView.GetParametrosProceso", ErrSource)
    Err.Raise ErrNumber, "GenCombinacionesView.GetParametrosProceso", ErrDescription
End Function



'------------------------------------------------------------------------------*
' Procedimiento  : ClearSalida
' Fecha          : mi., 21/ago/2019 19:10:10
' Propósito      : Inicializar las celdas de salida de la combinación
'------------------------------------------------------------------------------*
Public Sub ClearSalida()
    Dim rgBorra As Range
    Dim nRows  As Integer
    
  On Error GoTo ClearSalida_Error
    Application.ScreenUpdating = False
    Set rgBorra = mwsUITotCombi.Range("A2").CurrentRegion
    nRows = rgBorra.Rows.Count
    '
    '   Mantenemos la cabecera
    '
    If nRows > 2 Then
        Set rgBorra = rgBorra.Offset(2, 0).Resize(nRows - 2, rgBorra.Columns.Count)
        rgBorra.ClearContents
    End If
    Application.ScreenUpdating = True
    Set rgBorra = Nothing
  On Error GoTo 0
ClearSalida__CleanExit:
    Exit Sub
            
ClearSalida_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    Call HandleException(ErrNumber, ErrDescription, "GenCombinacionesView.ClearSalida", ErrSource)
    Err.Raise ErrNumber, "GenCombinacionesView.ClearSalida", ErrDescription
End Sub




'------------------------------------------------------------------------------*
' Procedimiento  : ClearSalidaComprobacion
' Fecha          : mi., 03/feb/2021 16:54:44
' Propósito      : Inicializar las celdas de salida para el proceso de
'                  comprobación
'------------------------------------------------------------------------------*
Public Sub ClearSalidaComprobacion()
    Dim nRows  As Integer
    
  On Error GoTo ClearSalidaComprobacion_Error
    Application.ScreenUpdating = False
    Set mRg = mwsUICompCombi.Range("A2").CurrentRegion
    nRows = mRg.Rows.Count
    '
    '   Mantenemos la cabecera
    '
    If nRows > 2 Then
        Set mRg = mRg.Offset(2, 0).Resize(nRows - 2, mRg.Columns.Count)
        mRg.ClearContents
    End If
    Application.ScreenUpdating = True
    Set mRg = Nothing
  
  On Error GoTo 0
ClearSalidaComprobacion__CleanExit:
    Exit Sub
ClearSalidaComprobacion_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    Call HandleException(ErrNumber, ErrDescription, "GenCombinacionesView.ClearSalidaComprobacion", ErrSource)
    Err.Raise ErrNumber, "GenCombinacionesView.ClearSalidaComprobacion", ErrDescription
End Sub




'------------------------------------------------------------------------------*
' Procedimiento  : GetFiltros
' Fecha          : mi., 28/ago/2019 18:48:23
' Propósito      : Obtiene la matriz de filtro
'------------------------------------------------------------------------------*
Private Function GetFiltros() As Variant
    Dim mMtrz() As String
    Dim mFiltr  As String
    Dim i       As Integer
    
 On Error GoTo GetFiltros_Error
    '
    '   Inicializamos la Matriz con un elemento vacio
    '
    i = 0
    ReDim Preserve mMtrz(i)
    mMtrz(i) = Empty
    '
    '   Nos posicionamos en la celda de inicio
    '
    Set mRg = mwsUIFiltros.Range("B11")
    '
    '   Repetir hasta que la celda esté vacia o espacios
    '
    Do
        '
        '   Obtenemos el contenido de la celda
        '
        mFiltr = mRg.Offset(i, 0).Value
        '
        '   Si tiene contenido
        '
        If Len(mFiltr) > 0 Then
            ReDim Preserve mMtrz(i)
            mMtrz(i) = mFiltr
        End If
        i = i + 1
    Loop Until (IsEmpty(mFiltr) Or Len(mFiltr) = 0)
    '
    '
    '
    GetFiltros = mMtrz

  On Error GoTo 0
GetFiltros__CleanExit:
    Exit Function
            
GetFiltros_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    Call HandleException(ErrNumber, ErrDescription, "GenPullCombinacionesView.GetFiltros", ErrSource)
    Err.Raise ErrNumber, "GenPullCombinacionesView.GetFiltros", ErrDescription
End Function


'------------------------------------------------------------------------------*
' Procedimiento  : GetNumeros
' Fecha          : mi., 28/ago/2019 19:30:23
' Propósito      : Obtiene la matriz de Numeros
'------------------------------------------------------------------------------*
Private Function GetNumeros() As Variant
    Dim mMtrz() As String
    Dim mNum    As String
    Dim i       As Integer
    
 On Error GoTo GetNumeros_Error
    '
    '   Inicializamos la Matriz con un elemento vacio
    '
    i = 0
    ReDim Preserve mMtrz(i)
    mMtrz(i) = Empty
    '
    '   Nos posicionamos en la celda de inicio
    '
    Set mRg = mwsUINumeros.Range("B8")
    '
    '   Repetir hasta que la celda esté vacia o espacios
    '
    Do
        '
        '   Obtenemos el contenido de la celda
        '
        mNum = mRg.Offset(i, 0).Value
        '
        '   Si tiene contenido
        '
        If Len(mNum) > 0 Then
            ReDim Preserve mMtrz(i)
            mMtrz(i) = mNum
        End If
        i = i + 1
    Loop Until (IsEmpty(mNum) Or Len(mNum) = 0)
    '
    '
    '
    GetNumeros = mMtrz

  On Error GoTo 0
GetNumeros__CleanExit:
    Exit Function
            
GetNumeros_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    Call HandleException(ErrNumber, ErrDescription, "GenPullCombinacionesView.GetNumeros", ErrSource)
    Err.Raise ErrNumber, "GenPullCombinacionesView.GetNumeros", ErrDescription
End Function


'------------------------------------------------------------------------------*
' Procedimiento  : AgregarFiltro
' Fecha          : do., 01/sep/2019 20:42:52
' Propósito      : Agregar un nuevo filtro a la matriz
'------------------------------------------------------------------------------*
Public Sub AgregarFiltro()
    Dim mFiltro As String
    Dim i As Integer
    Dim mStr As Variant
    Dim mFltr As FiltroCombinacion
    
  On Error GoTo AgregarFiltro_Error
    '
    '   Comprobar que valor filtro no es vacio
    '
    If Len(Me.ValorFiltro) = 0 Then
        MsgBox "El valor del filtro está vacio.", vbOKOnly + vbExclamation, "Agregar Filtro Combinacion"
        GoTo AgregarFiltro__CleanExit
    End If
    '
    '   Pasamos los valores al modelo y nos devuelve un filtro
    '
    Set mFltr = New FiltroCombinacion
    mFltr.NameFiltro = Me.TipoFiltro
    mFltr.FilterValue = Me.ValorFiltro
    mFiltro = mFltr.ToString
    '
    '   Inicializamos
    '
    i = 0
    '
    '   Nos posicionamos en la matriz de filtros
    '
    Set mRg = mwsUIFiltros.Range(CELDA_MATRIZFILTROS)
    '
    '   Buscamos la primera celda vacia
    '
    Do
        '
        '   Obtenemos el contenido de la celda
        '
        mStr = mRg.Offset(i, 0).Value
        '
        '   Si tiene contenido
        '
        If Len(mStr) > 0 Then
            If mStr = mFiltro Then
                MsgBox "El filtro ya está definido.", vbOKOnly + vbExclamation, "Agregar Filtro Combinacion"
                GoTo AgregarFiltro__CleanExit
            End If
            i = i + 1
        End If
        
    Loop Until (IsEmpty(mStr) Or Len(mStr) = 0)
    '
    '   Si hay más de 10 filtro avisamos
    '
    If i > 11 Then
        MsgBox "El número máximo de filtros admitidos es 10.", vbOKOnly + vbExclamation, "Agregar Filtro Combinacion"
        GoTo AgregarFiltro__CleanExit
    Else
        mRg.Offset(i, 0).Value = mFiltro
    End If
  On Error GoTo 0
AgregarFiltro__CleanExit:
    Exit Sub
            
AgregarFiltro_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    Call HandleException(ErrNumber, ErrDescription, "GenPullCombinacionesView.AgregarFiltro", ErrSource)
    Err.Raise ErrNumber, "GenPullCombinacionesView.AgregarFiltro", ErrDescription
End Sub



'------------------------------------------------------------------------------*
' Procedimiento  : BorrarFiltros
' Fecha          : do., 01/sep/2019 20:39:59
' Propósito      : Eliminar el contenido de las celdas donde se agregan filtros
'------------------------------------------------------------------------------*
Public Sub BorrarFiltros()
    Dim nRows  As Integer
    
    Dim i       As Integer
    Dim mStr    As Variant
    Dim mRg     As Range
  On Error GoTo BorrarFiltros_Error
    '
    '   Obtenemos el área de los filtros
    '
    Set mRg = mwsUIFiltros.Range(CELDA_MATRIZFILTROS).CurrentRegion
    '
    '   Modificamos el area a una columna
    '
    nRows = mRg.Rows.Count
    Set mRg = mRg.CurrentRegion.Offset(1, 1).Resize(nRows, 1)
    
    mRg.ClearContents
    
 On Error GoTo 0
BorrarFiltros__CleanExit:
    Exit Sub
            
BorrarFiltros_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    Call HandleException(ErrNumber, ErrDescription, "GenPullCombinacionesView.BorrarFiltros", ErrSource)
    Err.Raise ErrNumber, "GenPullCombinacionesView.BorrarFiltros", ErrDescription
End Sub



'------------------------------------------------------------------------------*
' Procedimiento  : SetFiltros
' Fecha          : lu., 02/sep/2019 18:59:43
' Propósito      : Establece los posibles valores para un filtro
'------------------------------------------------------------------------------*
Public Sub SetFiltros(datArray As Variant)
    Dim i As Integer
    Dim mAddress As String
    Dim nRows  As Integer
    
  On Error GoTo SetFiltros_Error
    '
    '   Borramos los filtros existentes
    '
    Set mRg = mwsParametros.Range(CELDA_CATEGORIAFILTERS).CurrentRegion
    nRows = mRg.Rows.Count
    Set mRg = mRg.CurrentRegion.Offset(1, 0).Resize(nRows, 1)
    mRg.ClearContents
    '
    '       Trasladamos la matriz a las celdas
    '
    Set mRg = mwsParametros.Range(CELDA_CATEGORIAFILTERS).Offset(1, 0)
    For i = 0 To UBound(datArray)
        mRg.Offset(i, 0).Value = datArray(i)
    Next i
    '
    '   Configuramos el rango de validación
    '
    mAddress = "=" & HOJA_PARAMETROS & "$D2:$D" & CInt(i + 1)
    '
    '   Actualizamos la validación de la celda
    '
    With mwsUIFiltros.Range(CELDA_VALORFILTRO).Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:=mAddress
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With
    '
    '   Inicializamos la celda del valor del filtro
    '
    mwsUIFiltros.Range(CELDA_VALORFILTRO).ClearContents
 
 On Error GoTo 0
SetFiltros__CleanExit:
    Exit Sub
            
SetFiltros_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    Call HandleException(ErrNumber, ErrDescription, "GenPullCombinacionesView.SetFiltros", ErrSource)
    Err.Raise ErrNumber, "GenPullCombinacionesView.SetFiltros", ErrDescription
End Sub




'------------------------------------------------------------------------------*
' Procedimiento  : GetSorteos
' Fecha          : mi., 03/feb/2021 18:40:27
' Propósito      : Obtiene los sorteos en una matriz bidimensional
'------------------------------------------------------------------------------*
Public Function GetSorteos() As Variant
    Dim i As Integer
    
  On Error GoTo GetSorteos_Error
    '
    '   Obtenemos el rango de los sorteos
    '
    Set mRg = mwsUISorteos.Range("A1").CurrentRegion
    i = mRg.Rows.Count
    '
    '   Eliminamos la cabecera del rango
    '
    Set mRg = mRg.Offset(2, 0).Resize(i - 2, mRg.Columns.Count)
    '
    '   Devolvemos la matriz de valores
    '
    GetSorteos = mRg.Value2
  
  On Error GoTo 0
GetSorteos__CleanExit:
    Exit Function
GetSorteos_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    Call HandleException(ErrNumber, ErrDescription, "GenPullCombinacionesView.GetSorteos", ErrSource)
    Err.Raise ErrNumber, "GenPullCombinacionesView.GetSorteos", ErrDescription
End Function



'------------------------------------------------------------------------------*
' Procedimiento  : GetCombinacionGanadora
' Fecha          : mi., 03/feb/2021 18:51:53
' Propósito      : Obtiene la combinación ganadora para comparar
'------------------------------------------------------------------------------*
Public Function GetCombinacionGanadora() As Variant
    
  On Error GoTo GetCombinacionGanadora_Error
    '
    '   Obtenemos el rango de los sorteos
    '
    Set mRg = mwsUISorteoRes.Range("$A$3:$O$3")
    '
    '   Devolvemos la matriz de valores
    '
    GetCombinacionGanadora = mRg.Value2
  
  On Error GoTo 0
GetCombinacionGanadora__CleanExit:
    Exit Function
GetCombinacionGanadora_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    Call HandleException(ErrNumber, ErrDescription, "GenPullCombinacionesView.GetCombinacionGanadora", ErrSource)
    Err.Raise ErrNumber, "GenPullCombinacionesView.GetCombinacionGanadora", ErrDescription
End Function
'' *===========(EOF): GenPullCombinacionesView.cls
