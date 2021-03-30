Attribute VB_Name = "CU_GenerarEstadisticas"
' *============================================================================*
' *
' *     Fichero    : CU_GenerarEstadisticas.bas
' *
' *     Autor      : Carlos Almela Baeza
' *     Creación   : mi., 21/ago/2019 18:29:26
' *     Versión    : 1.0
' *     Propósito  : Generador de combinaciones con un numero limitado de
' *                  elementos
' *============================================================================*
Option Explicit
Option Base 0

'--- Variables Privadas -------------------------------------------------------*
Private GenCtrl     As GenCombinacionesController

'--- Constantes ---------------------------------------------------------------*
Public Const NOMBRES_TIPOS_FILTRO As String = "Paridad;Peso;Consecutivos;Decenas;Septenas;Suma;Terminaciones"

Public Enum TipoFiltro
    tfParidad = 1
    tfAltoBajo = 2
    tfConsecutivos = 3
    tfDecenas = 4
    tfSeptenas = 5
    tfSuma = 6
    tfTerminaciones = 7
End Enum

'--- Métodos Públicos ---------------------------------------------------------*
'------------------------------------------------------------------------------*
' Procedimiento  : GenerarCombinaciones
' Fecha          : do., 01/sep/2019 00:24:46
' Propósito      : Invocar al proceso de generación de combinaciones
'------------------------------------------------------------------------------*
Public Sub GenerarCombinaciones()
  On Error GoTo GenerarCombinaciones_Error
    '
    '   Creamos el controlador del Caso de Uso
    '
    If GenCtrl Is Nothing Then
        Set GenCtrl = New GenCombinacionesController
    End If
    '
    '   Invocamos al controlador para que genere las combinaciones
    '
    GenCtrl.GenerarCombinaciones
  
  
  On Error GoTo 0
GenerarCombinaciones__CleanExit:
    Exit Sub
            
GenerarCombinaciones_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    Call HandleException(ErrNumber, ErrDescription, "CU_GenerarEstadisticas.GenerarCombinaciones", ErrSource)
    Call MsgBox(ErrDescription, vbCritical Or vbSystemModal, ThisWorkbook.Name)
    Call Trace("CERRAR")
End Sub

'------------------------------------------------------------------------------*
' Procedimiento  : AgregarFiltro
' Fecha          : do., 01/sep/2019 00:24:46
' Propósito      : Agregar un filtro a la lista de filtros a aplicar
'------------------------------------------------------------------------------*
Public Sub AgregarFiltro()
  On Error GoTo AgregarFiltro_Error
    '
    '   Creamos el controlador del Caso de Uso
    '
    If GenCtrl Is Nothing Then
        Set GenCtrl = New GenCombinacionesController
    End If
    '
    '   Invocamos al controlador para que genere las combinaciones
    '
    GenCtrl.AgregarFiltro
  
  On Error GoTo 0
AgregarFiltro__CleanExit:
    Exit Sub
            
AgregarFiltro_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    Call HandleException(ErrNumber, ErrDescription, "CU_GenerarEstadisticas.AgregarFiltro", ErrSource)
    Call MsgBox(ErrDescription, vbCritical Or vbSystemModal, ThisWorkbook.Name)
    Call Trace("CERRAR")
End Sub
'------------------------------------------------------------------------------*
' Procedimiento  : GenerarFiltros
' Fecha          : do., 01/sep/2019 00:24:46
' Propósito      : Generar todos los filtros posibles para un tipo de filtro
'------------------------------------------------------------------------------*
Public Sub GenerarFiltros()
  On Error GoTo GenerarFiltros_Error
    '
    '   Creamos el controlador del Caso de Uso
    '
    If GenCtrl Is Nothing Then
        Set GenCtrl = New GenCombinacionesController
    End If
    '
    '   Invocamos al controlador para que genere las combinaciones
    '
    GenCtrl.GenerarFiltros
  
  On Error GoTo 0
GenerarFiltros__CleanExit:
    Exit Sub
            
GenerarFiltros_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    Call HandleException(ErrNumber, ErrDescription, "CU_GenerarEstadisticas.GenerarFiltros", ErrSource)
    Call MsgBox(ErrDescription, vbCritical Or vbSystemModal, ThisWorkbook.Name)
    Call Trace("CERRAR")
End Sub
'------------------------------------------------------------------------------*
' Procedimiento  : BorrarFiltros
' Fecha          : do., 01/sep/2019 10:46:29
' Propósito      : Borra el listado de filtros para la generación
'------------------------------------------------------------------------------*
Public Sub BorrarFiltros()
  On Error GoTo BorrarFiltros_Error
    '
    '   Creamos el controlador del Caso de Uso
    '
    If GenCtrl Is Nothing Then
        Set GenCtrl = New GenCombinacionesController
    End If
    '
    '   Invocamos al controlador para que borre los filtros
    '
    GenCtrl.BorrarFiltros
  
  On Error GoTo 0
BorrarFiltros__CleanExit:
    Exit Sub
            
BorrarFiltros_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    Call HandleException(ErrNumber, ErrDescription, "CU_GenerarEstadisticas.BorrarFiltros", ErrSource)
    Call MsgBox(ErrDescription, vbCritical Or vbSystemModal, ThisWorkbook.Name)
    Call Trace("CERRAR")
End Sub
' *===========(EOF): CU_GenerarEstadisticas.bas
