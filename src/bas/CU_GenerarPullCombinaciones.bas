Attribute VB_Name = "CU_GenerarPullCombinaciones"
' *============================================================================*
' *
' *     Fichero    : CU_GenerarPullCombinaciones.bas
' *
' *     Autor      : Carlos Almela Baeza
' *     Creación   : sab, 12/dic/2020 12:44:32
' *     Versión    : 1.0
' *     Propósito  : Caso de uso que genera combinaciones de un conjunto
' *                  de numeros
' *============================================================================*
Option Explicit
Option Base 0

'
'--- Variables Privadas -------------------------------------------------------*
Private GenCtrl     As GenPullCombinacionesController

'--- Constantes ---------------------------------------------------------------*
Public Const NOMBRES_TIPOS_FILTRO As String = "Paridad;Peso;Consecutivos;Decena" & _
    "s;Septenas;Suma;Terminaciones"

Public Enum TipoFiltro
    tfParidad = 1
    tfAltoBajo = 2
    tfConsecutivos = 3
    tfDecenas = 4
    tfSeptenas = 5
    tfSuma = 6
    tfTerminaciones = 7
End Enum

Public Const FASE_GENERAR As String = "Generar"
Public Const FASE_FILTRAR As String = "Filtrar"
Public Const FASE_EVALUAR As String = "Evaluar"
Public Const FASE_COMPROBAR As String = "Comprobar"

'--- Mensajes -----------------------------------------------------------------*
'--- Errores ------------------------------------------------------------------*
'--- Métodos Públicos ---------------------------------------------------------*

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
        Set GenCtrl = New GenPullCombinacionesController
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
    Call HandleException(ErrNumber, ErrDescription, "CU_GenerarPullCombinaciones.AgregarFiltro", ErrSource)
    Call MsgBox(ErrDescription, vbCritical Or vbSystemModal, ThisWorkbook.Name)
    Call Trace("CERRAR")
End Sub



'------------------------------------------------------------------------------*
' Procedimiento  : GenerarFiltros
' Fecha          : do., 01/sep/2019 00:24:46
' Propósito      : Generar todas las combinaciones de filtro posibles
'                  para un tipo de filtro
'------------------------------------------------------------------------------*
Public Sub GenerarFiltros()
  On Error GoTo GenerarFiltros_Error
    '
    '   Creamos el controlador del Caso de Uso
    '
    If GenCtrl Is Nothing Then
        Set GenCtrl = New GenPullCombinacionesController
    End If
    '
    '   Invocamos al controlador para que genere las combinaciones de un filtro
    '
    GenCtrl.GenerarFiltros
  
  On Error GoTo 0
GenerarFiltros__CleanExit:
    Exit Sub
GenerarFiltros_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    Call HandleException(ErrNumber, ErrDescription, "CU_GenerarPullCombinaciones.GenerarFiltros", ErrSource)
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
        Set GenCtrl = New GenPullCombinacionesController
    End If
    '
    '   Invocamos al controlador para que ejecute Borrar Filtros
    '
    GenCtrl.BorrarFiltros
    
  On Error GoTo 0
BorrarFiltros__CleanExit:
    Exit Sub
BorrarFiltros_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    Call HandleException(ErrNumber, ErrDescription, "CU_GenerarPullCombinaciones.BorrarFiltros", ErrSource)
    Call MsgBox(ErrDescription, vbCritical Or vbSystemModal, ThisWorkbook.Name)
    Call Trace("CERRAR")
End Sub


'------------------------------------------------------------------------------*
' Procedimiento  : GenerarCombinaciones
' Fecha          : do., 20/dic/2020 18:10:56
' Propósito      : Genera todas las combinaciones posibles con un subconjunto
'                  de números.
'------------------------------------------------------------------------------*
Public Sub GenerarCombinaciones()
 On Error GoTo GenerarCombinaciones_Error
    '
    '   Creamos el controlador del Caso de Uso
    '
    If GenCtrl Is Nothing Then
        Set GenCtrl = New GenPullCombinacionesController
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
    Call HandleException(ErrNumber, ErrDescription, "CU_GenerarPullCombinaciones.GenerarCombinaciones", ErrSource)
    Call MsgBox(ErrDescription, vbCritical Or vbSystemModal, ThisWorkbook.Name)
    Call Trace("CERRAR")
End Sub




'------------------------------------------------------------------------------*
' Procedimiento  : FiltrarCombinaciones
' Fecha          : lu., 21/dic/2020 18:36:15
' Propósito      : Aplica los filtros de a las combinaciones generadas
'------------------------------------------------------------------------------*
Public Sub FiltrarCombinaciones()
  On Error GoTo FiltrarCombinaciones_Error
    '
    '   Creamos el controlador del Caso de Uso
    '
    If GenCtrl Is Nothing Then
        Set GenCtrl = New GenPullCombinacionesController
    End If
    '
    '   Invocamos al controlador para que genere las combinaciones
    '
    GenCtrl.FiltrarCombinaciones
  
  On Error GoTo 0
FiltrarCombinaciones__CleanExit:
    Exit Sub
FiltrarCombinaciones_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    Call HandleException(ErrNumber, ErrDescription, "CU_GenerarPullCombinaciones.FiltrarCombinaciones", ErrSource)
    Call MsgBox(ErrDescription, vbCritical Or vbSystemModal, ThisWorkbook.Name)
    Call Trace("CERRAR")
End Sub





'------------------------------------------------------------------------------*
' Procedimiento  : EvaluarCombinaciones
' Fecha          :
' Propósito      :
'------------------------------------------------------------------------------*
Public Sub EvaluarCombinaciones()
  On Error GoTo EvaluarCombinaciones_Error
    '
    '   Creamos el controlador del Caso de Uso
    '
    If GenCtrl Is Nothing Then
        Set GenCtrl = New GenPullCombinacionesController
    End If
    '
    '   Invocamos la función del controlador
    '
    GenCtrl.EvaluarCombinaciones
  
  
  On Error GoTo 0
EvaluarCombinaciones__CleanExit:
    Exit Sub
EvaluarCombinaciones_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    Call HandleException(ErrNumber, ErrDescription, "CU_GenerarPullCombinaciones.EvaluarCombinaciones", ErrSource)
    Call MsgBox(ErrDescription, vbCritical Or vbSystemModal, ThisWorkbook.Name)
    Call Trace("CERRAR")
End Sub



'------------------------------------------------------------------------------*
' Procedimiento  : ComprobarCombinaciones
' Fecha          :
' Propósito      : Comprueba las combinaciones seleccionadas con el sorteo
'                  resultantes
'------------------------------------------------------------------------------*
Public Sub ComprobarCombinaciones()
  On Error GoTo ComprobarCombinaciones_Error
    '
    '   Creamos el controlador del Caso de Uso
    '
    If GenCtrl Is Nothing Then
        Set GenCtrl = New GenPullCombinacionesController
    End If
    '
    '   Invocamos la función del controlador
    '
    GenCtrl.ComprobarCombinaciones
    
  On Error GoTo 0
ComprobarCombinaciones__CleanExit:
    Exit Sub

ComprobarCombinaciones_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    Call HandleException(ErrNumber, ErrDescription, "CU_GenerarPullCombinaciones.ComprobarCombinaciones", ErrSource)
    Call MsgBox(ErrDescription, vbCritical Or vbSystemModal, ThisWorkbook.Name)
    Call Trace("CERRAR")
End Sub
' *===========(EOF): CU_GenerarPullCombinaciones.bas
