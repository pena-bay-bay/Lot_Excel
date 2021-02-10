VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GenPullCombinacionesController"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' *============================================================================*
' *
' *     Fichero    : GenPullCombinacionesController.cls
' *
' *     Autor      : Carlos Almela Baeza
' *     Creación   : lu., 14/dic/2020 23:28:36
' *     Versión    : 1.1
' *     Propósito  : Controlador del caso de uso Generar pull de combinaciones
' *
' *============================================================================*
Option Explicit
Option Base 0
'
'--- Variables Privadas -------------------------------------------------------*
Private mModel          As GenPullCombinacionesModel   ' Modelo de negocio
Private mView           As GenPullCombinacionesView    ' Vista del modelo

'--- Métodos Privados ---------------------------------------------------------*
Private Sub Class_Initialize()
    Set mModel = New GenPullCombinacionesModel
    Set mView = New GenPullCombinacionesView
End Sub

Private Sub Class_Terminate()
    Set mModel = Nothing
    Set mView = Nothing
End Sub

'--- Métodos Públicos ---------------------------------------------------------*
'------------------------------------------------------------------------------*
' Procedimiento  : GenerarCombinaciones
' Fecha          : mi., 21/ago/2019 22:53:11
' Propósito      : Generar las combinaciones
'------------------------------------------------------------------------------*
Public Sub GenerarCombinaciones()
  On Error GoTo GenerarCombinaciones_Error
    '
    '   Obtenemos los datos de la hoja
    '
    Set mModel = mView.GetParametrosProceso
    '
    '   Si los parametros no son correctos
    '
    If Not mModel.IsValid Then
        '
        '   Mensaje del error
        '
        MsgBox mModel.GetMessage, vbOK + vbInformation, ThisWorkbook.Name
        Exit Sub
    End If
    '
    '   Si es correcto inicializamos el área de salida
    '
    mView.ClearSalida
    '
    '   Invocamos al modelo para que genere combinaciones pasamos el total
    '
    mModel.GenerarCombinaciones mView.RangoSalida, mView.TotalCombinaciones
    
  On Error GoTo 0
GenerarCombinaciones__CleanExit:
    Exit Sub
GenerarCombinaciones_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    Call HandleException(ErrNumber, ErrDescription, "GenPullCombinacionesController.GenerarCombinaciones", ErrSource)
    Err.Raise ErrNumber, "GenPullCombinacionesController.GenerarCombinaciones", ErrDescription
End Sub




'------------------------------------------------------------------------------*
' Procedimiento  : GenerarFiltros
' Fecha          : do., 01/sep/2019 00:32:05
' Propósito      : Generar las categorias de filtros para un tipo de filtro
'------------------------------------------------------------------------------*
Public Sub GenerarFiltros()
    Dim mFiltro As String
    Dim mArray  As Variant
    Dim mNum    As Integer
    
  On Error GoTo GenerarFiltros_Error
    '
    '   Obtiene el tipo de filtro
    '
    mFiltro = mView.TipoFiltro
    mNum = mView.NumSugerencias
    '
    '   Llama al modelo que genere una matriz
    '
    mArray = mModel.GetFiltrosOf(mFiltro, mNum)
    '
    '   Pasa la matriz a la vista para que lo actualice
    '
    mView.SetFiltros mArray
  
  On Error GoTo 0
GenerarFiltros__CleanExit:
    Exit Sub
GenerarFiltros_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    Call HandleException(ErrNumber, ErrDescription, "GenPullCombinacionesController.GenerarFiltros", ErrSource)
    Err.Raise ErrNumber, "GenPullCombinacionesController.GenerarFiltros", ErrDescription
End Sub





'------------------------------------------------------------------------------*
' Procedimiento  : AgregarFiltro
' Fecha          : do., 01/sep/2019 00:32:05
' Propósito      : Agrega un filtro definido a la lista de filtros a aplicar
'------------------------------------------------------------------------------*
Public Sub AgregarFiltro()
  On Error GoTo AgregarFiltro_Error
    '
    '  Llama a la vista para que agregue a la lista
    '
    mView.AgregarFiltro
  
  On Error GoTo 0
AgregarFiltro__CleanExit:
    Exit Sub
            
AgregarFiltro_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    Call HandleException(ErrNumber, ErrDescription, "GenPullCombinacionesController.AgregarFiltro", ErrSource)
    Err.Raise ErrNumber, "GenPullCombinacionesController.AgregarFiltro", ErrDescription
End Sub





'------------------------------------------------------------------------------*
' Procedimiento  : BorrarFiltros
' Fecha          : do., 01/sep/2019 10:46:29
' Propósito      : Borra el listado de filtros pendientes
'------------------------------------------------------------------------------*
Public Sub BorrarFiltros()
  On Error GoTo BorrarFiltros_Error
    '
    '   Llama a la vista para que borre la matriz de filtros
    '
    mView.BorrarFiltros
  
 On Error GoTo 0
BorrarFiltros__CleanExit:
    Exit Sub
BorrarFiltros_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    Call HandleException(ErrNumber, ErrDescription, "GenPullCombinacionesController.BorrarFiltros", ErrSource)
    Err.Raise ErrNumber, "GenPullCombinacionesController.BorrarFiltros", ErrDescription
End Sub





'------------------------------------------------------------------------------*
' Procedimiento  : FiltrarCombinaciones
' Fecha          : lu., 21/dic/2020 18:38:32
' Propósito      : Aplicar filtros
'------------------------------------------------------------------------------*
Public Sub FiltrarCombinaciones()
  On Error GoTo FiltrarCombinaciones_Error
    '
    '   Obtenemos los datos de la hoja
    '
    Set mModel = mView.GetParametrosProceso
    '
    '   Establecemos la fase del proceso
    '
    mModel.FaseProceso = FASE_FILTRAR
    '
    '   Si los parametros no son correctos
    '
    If Not mModel.IsValid Then
        '
        '   Mensaje del error
        '
        MsgBox mModel.GetMessage, vbOK + vbInformation, ThisWorkbook.Name
        Exit Sub
    End If
    '
    '   Si es correcto inicializamos el área de salida
    '
    mView.ClearSalidaFiltros
    '
    '   Invocamos al modelo para que genere combinaciones
    '
    mModel.FiltrarCombinaciones mView.RangoSalida, _
                                mView.RangoSalidaFiltros, _
                                mView.TotalCombinaciones
                                
 On Error GoTo 0
FiltrarCombinaciones__CleanExit:
    Exit Sub
FiltrarCombinaciones_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    Call HandleException(ErrNumber, ErrDescription, "GenPullCombinacionesController.FiltrarCombinaciones", ErrSource)
    Err.Raise ErrNumber, "GenPullCombinacionesController.FiltrarCombinaciones", ErrDescription
End Sub




'------------------------------------------------------------------------------*
' Procedimiento  : EvaluarCombinaciones
' Fecha          : mi., 03/feb/2021 16:50:05
' Propósito      : Evalua las combinaciones filtradas con los sorteos ya
'                  celebrados
'------------------------------------------------------------------------------*
Public Sub EvaluarCombinaciones()
  On Error GoTo EvaluarCombinaciones_Error
    '
    '   Obtenemos los datos de la hoja
    '
    Set mModel = mView.GetParametrosProceso
    '
    '   Establecemos la fase de evaluación
    '
    mModel.FaseProceso = FASE_EVALUAR
    '
    '   Si los parametros no son correctos
    '
    If Not mModel.IsValid Then
        '
        '   Mensaje del error
        '
        MsgBox mModel.GetMessage, vbOK + vbInformation, ThisWorkbook.Name
        Exit Sub
    End If
    '
    '   Si es correcto inicializamos el área de salida
    '
    mView.ClearSalidaEvaluacion
    '
    '   Invocamos al modelo para que genere combinaciones
    '
    mModel.EvaluarCombinaciones mView.RangoSalidaFiltros, _
                                mView.RangoSalidaEvaluadas, _
                                mView.CombinacionesFiltradas
                                
 On Error GoTo 0
EvaluarCombinaciones__CleanExit:
    Exit Sub
EvaluarCombinaciones_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    Call HandleException(ErrNumber, ErrDescription, "GenPullCombinacionesController.EvaluarCombinaciones", ErrSource)
    Err.Raise ErrNumber, "GenPullCombinacionesController.EvaluarCombinaciones", ErrDescription
End Sub





'------------------------------------------------------------------------------*
' Procedimiento  : ComprobarCombinaciones
' Fecha          : mi., 03/feb/2021 16:47:13
' Propósito      : Comprobar las combinaciones evaluadas con la combinación
'                  Ganadora
'------------------------------------------------------------------------------*
Public Sub ComprobarCombinaciones()
  On Error GoTo ComprobarCombinaciones_Error
    '
    '   Obtenemos los datos de la hoja
    '
    Set mModel = mView.GetParametrosProceso
    '
    '   Establecemos la fase de Comprobación
    '
    mModel.FaseProceso = FASE_COMPROBAR
    '
    '   Si los parametros no son correctos
    '
    If Not mModel.IsValid Then
        '
        '   Mensaje del error
        '
        MsgBox mModel.GetMessage, vbOK + vbInformation, ThisWorkbook.Name
        Exit Sub
    End If
    '
    '   Si es correcto inicializamos el área de salida
    '
    mView.ClearSalidaComprobacion
    '
    '   Inicialiamos estadistica
    '
    mView.RegistrosComprobados = mModel.TotRegComprobados
    mView.RegistrosPremiados = mModel.TotRegPremiados
    mView.TotalCoste = mModel.TotCoste
    mView.TotalImporte = mModel.TotImportePremios
    '
    '   Invocamos al modelo para que genere combinaciones
    '
    mModel.ComprobarCombinaciones mView.RangoSalidaEvaluadas, _
                                  mView.RangoSalidaComprobacion, _
                                  mView.CombinacionesEvaluadas
                                  
    '
    '   Visualizamos estadistica de la comprobación
    '
    mView.RegistrosComprobados = mModel.TotRegComprobados
    mView.RegistrosPremiados = mModel.TotRegPremiados
    mView.TotalCoste = mModel.TotCoste
    mView.TotalImporte = mModel.TotImportePremios
    
    
 On Error GoTo 0
ComprobarCombinaciones__CleanExit:
    Exit Sub
ComprobarCombinaciones_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    Call HandleException(ErrNumber, ErrDescription, "GenPullCombinacionesController.ComprobarCombinaciones", ErrSource)
    Err.Raise ErrNumber, "GenPullCombinacionesController.ComprobarCombinaciones", ErrDescription
End Sub

'' *===========(EOF): GenPullCombinacionesController.cls
