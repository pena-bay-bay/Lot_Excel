VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GenCombinacionesController"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' *============================================================================*
' *
' *     Fichero    : GenCombinacionesController.cls
' *
' *     Autor      : Carlos Almela Baeza
' *     Creación   : mi., 21/ago/2019 18:34:50
' *     Versión    : 1.0
' *     Propósito  : Controlador del caso de uso Generar combinaciones
' *
' *============================================================================*
Option Explicit
Option Base 0
'
'--- Variables Privadas -------------------------------------------------------*
Private mModel          As GenCombinacionesModel   ' Modelo de negocio
Private mView           As GenCombinacionesView    ' CU Validar Parametros

'--- Métodos Privados ---------------------------------------------------------*
Private Sub Class_Initialize()
    Set mModel = New GenCombinacionesModel
    Set mView = New GenCombinacionesView
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
    '   Invocamos al modelo para que genere combinaciones
    '
    mModel.GenerarCombinaciones mView.RangoSalida
  On Error GoTo 0
GenerarCombinaciones__CleanExit:
  Exit Sub
            
GenerarCombinaciones_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    Call HandleException(ErrNumber, ErrDescription, "GenCombinacionesController.GenerarCombinaciones", ErrSource)
    Err.Raise ErrNumber, "GenCombinacionesController.GenerarCombinaciones", ErrDescription
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
    Call HandleException(ErrNumber, ErrDescription, "GenCombinacionesController.GenerarFiltros", ErrSource)
    Err.Raise ErrNumber, "GenCombinacionesController.GenerarFiltros", ErrDescription
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
    Call HandleException(ErrNumber, ErrDescription, "GenCombinacionesController.AgregarFiltro", ErrSource)
    Err.Raise ErrNumber, "GenCombinacionesController.AgregarFiltro", ErrDescription
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
    Call HandleException(ErrNumber, ErrDescription, "GenCombinacionesController.BorrarFiltros", ErrSource)
    Err.Raise ErrNumber, "GenCombinacionesController.BorrarFiltros", ErrDescription
End Sub
'' *===========(EOF): GenCombinacionesController.cls
