VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SorteoUIConsultar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
' *============================================================================*
' *
' *     Fichero    : SorteoUIConsultar.cls
' *
' *     Autor      : Carlos Almela Baeza
' *     Creación   : mi., 08/may/2019 23:51:31
' *     Versión    : 1.0
' *     Propósito  : Hoja UI Seleccionar un sorteo
' *
' *============================================================================*
Option Explicit
Option Base 0

'--- Variables Privadas -------------------------------------------------------*
Private SrtCtrl     As SorteoController
Private mId         As Integer
Private mPeriodo    As Periodo

'--- Métodos Privados ---------------------------------------------------------*
Private Sub Worksheet_Activate()
    Set SrtCtrl = New SorteoController
End Sub

'------------------------------------------------------------------------------*
' Procedimiento  : Worksheet_BeforeDoubleClick
' Fecha          : vi., 26/jul/2019 20:32:16
' Propósito      : Interceptar el evento Doble Click de selección de registro
'                  dentro de la rejilla
'------------------------------------------------------------------------------*'
Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
  On Error GoTo Worksheet_BeforeDoubleClick_Error
    '
    '   Si la fila está contenida en la rejilla de datos
    '
    If (Target.Column >= 2) And (Target.Column <= 8) _
    And (Target.Row >= 12) And (Target.Row <= 18) Then
        '
        '   Obtenemos el identificador del registro
        '
        mId = CInt(Target.Offset(0, 2 - Target.Column).value)
        '
        '   Si es numérico lo editamos
        '
        If mId > 0 Then
            '
            '   Creamos el controlador si no existe
            '
            If SrtCtrl Is Nothing Then
                Set SrtCtrl = New SorteoController
            End If
            '
            '   Invocamos al controlador para crear un nuevo registro
            '
            SrtCtrl.EditarPorId mId, Target.Row
        End If
    End If
    
  On Error GoTo 0
Worksheet_BeforeDoubleClick__CleanExit:
    Exit Sub
            
Worksheet_BeforeDoubleClick_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    '   Audita el error
    Call HandleException(ErrNumber, ErrDescription, "SorteoUIConsultar.Worksheet_BeforeDoubleClick", ErrSource)
    '   Informa del error
    Call MsgBox(ErrDescription, vbCritical Or vbSystemModal, ThisWorkbook.Name)
    Call Trace("CERRAR")
    
End Sub

Private Sub Worksheet_Deactivate()
    Set SrtCtrl = Nothing
End Sub

'------------------------------------------------------------------------------*
' Procedimiento  : Worksheet_Change
' Fecha          :
' Propósito      : Interceptar el evento de cambio del contenido de ciertas
'                  celdas para modificar contenido de otras
'------------------------------------------------------------------------------*'
Private Sub Worksheet_Change(ByVal Target As Range)
    Dim mId As Integer
  On Error GoTo Worksheet_Change_Error
    '
    '   Comprobamos la selección del periodo
    '
    If Target.Address = "$C$7" Then
        '
        '   Creamos el periodo si no existe
        '
        If mPeriodo Is Nothing Then
            Set mPeriodo = New Periodo
        End If
        '
        '   Asignamos el texto del periodo a la clase
        '
        mPeriodo.Texto = Target.value
        '
        '   Si el contenido está vacío, vaciamos las celdas relacionadas
        '
        If IsEmpty(Target.value) Then
            Target.Offset(1, 0).ClearContents
            Target.Offset(2, 0).ClearContents
        Else
            Target.Offset(1, 0).value = mPeriodo.FechaInicial
            Target.Offset(2, 0).value = mPeriodo.FechaFinal
        End If
    End If
  
  On Error GoTo 0
Worksheet_Change__CleanExit:
    Exit Sub
            
Worksheet_Change_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    '   Audita el error
    Call HandleException(ErrNumber, ErrDescription, "SorteoUIConsultar.Worksheet_Change", ErrSource)
    '   Informa del error
    Call MsgBox(ErrDescription, vbCritical Or vbSystemModal, ThisWorkbook.Name)
    Call Trace("CERRAR")
 End Sub

'------------------------------------------------------------------------------*
' Procedimiento  : BtnNuevo
' Fecha          : ju., 16/may/2019 23:34:53
' Propósito      : Edita un nuevo registro
'------------------------------------------------------------------------------*'
Public Sub BtnNuevo()
  On Error GoTo BtnNuevo_Error
    
    If SrtCtrl Is Nothing Then
        Set SrtCtrl = New SorteoController
    End If
    '
    '   Invocamos al controlador para crear un nuevo registro
    '
    SrtCtrl.Nuevo
  
  On Error GoTo 0
BtnNuevo__CleanExit:
    Exit Sub
            
BtnNuevo_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    '   Audita el error
    Call HandleException(ErrNumber, ErrDescription, "SorteoUIConsultar.BtnNuevo", ErrSource)
    '   Informa del error
    Call MsgBox(ErrDescription, vbCritical Or vbSystemModal, ThisWorkbook.Name)
    Call Trace("CERRAR")
End Sub
'------------------------------------------------------------------------------*
' Procedimiento  : BtnBuscar
' Fecha          : ju., 16/may/2019 23:38:51
' Propósito      : Ejecuta una busqueda por los criterios definidos
'------------------------------------------------------------------------------*'
Public Sub BtnBuscar()
  On Error GoTo BtnBuscar_Error
    
    If SrtCtrl Is Nothing Then
        Set SrtCtrl = New SorteoController
    End If
    '
    '
    '
    '
    '   Invocamos al controlador para realizar una busqueda
    '
    SrtCtrl.Buscar
  
  On Error GoTo 0
BtnBuscar__CleanExit:
    Exit Sub
            
BtnBuscar_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    '   Audita el error
    Call HandleException(ErrNumber, ErrDescription, "SorteoUIConsultar.BtnBuscar", ErrSource)
    '   Informa del error
    Call MsgBox(ErrDescription, vbCritical Or vbSystemModal, ThisWorkbook.Name)
    Call Trace("CERRAR")
End Sub
'------------------------------------------------------------------------------*
' Procedimiento  : BtnInicio
' Fecha          : ju., 16/may/2019 23:45:53
' Propósito      : Nos visualiza la primera página del filtro seleccionado
'------------------------------------------------------------------------------*'
Public Sub BtnInicio()
  On Error GoTo BtnInicio_Error
    
    If SrtCtrl Is Nothing Then
        Set SrtCtrl = New SorteoController
    End If
    '
    '   Invocamos al controlador para Visualizar la primera página
    '
    SrtCtrl.BuscarFirstPage
  
  On Error GoTo 0
BtnInicio__CleanExit:
    Exit Sub
            
BtnInicio_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    '   Audita el error
    Call HandleException(ErrNumber, ErrDescription, "SorteoUIConsultar.BtnInicio", ErrSource)
    '   Informa del error
    Call MsgBox(ErrDescription, vbCritical Or vbSystemModal, ThisWorkbook.Name)
    Call Trace("CERRAR")
End Sub
'------------------------------------------------------------------------------*
' Procedimiento  : BtnFin
' Fecha          : ju., 16/may/2019 23:46:58
' Propósito      : Visualiza la ultima página del filtro seleccionado
'------------------------------------------------------------------------------*'
Public Sub BtnFin()
  On Error GoTo BtnFin_Error
    
    If SrtCtrl Is Nothing Then
        Set SrtCtrl = New SorteoController
    End If
    '
    '   Invocamos al controlador para Visualizar la última página
    '
    SrtCtrl.BuscarLastPage
  
  On Error GoTo 0
BtnFin__CleanExit:
    Exit Sub
            
BtnFin_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    '   Audita el error
    Call HandleException(ErrNumber, ErrDescription, "SorteoUIConsultar.BtnFin", ErrSource)
    '   Informa del error
    Call MsgBox(ErrDescription, vbCritical Or vbSystemModal, ThisWorkbook.Name)
    Call Trace("CERRAR")
End Sub
'------------------------------------------------------------------------------*
' Procedimiento  : BtnPagAnterior
' Fecha          : ju., 16/may/2019 23:51:20
' Propósito      : Pagina a la pagina anterior a la actual según el filtro
'------------------------------------------------------------------------------*'
Public Sub BtnPagAnterior()
  On Error GoTo BtnPagAnterior_Error
    
    If SrtCtrl Is Nothing Then
        Set SrtCtrl = New SorteoController
    End If
    '
    '   Invocamos al controlador para Visualizar la pagina anterior
    '
    SrtCtrl.BuscarPrevPage
  
  On Error GoTo 0
BtnPagAnterior__CleanExit:
    Exit Sub
            
BtnPagAnterior_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    '   Audita el error
    Call HandleException(ErrNumber, ErrDescription, "SorteoUIConsultar.BtnPagAnterior", ErrSource)
    '   Informa del error
    Call MsgBox(ErrDescription, vbCritical Or vbSystemModal, ThisWorkbook.Name)
    Call Trace("CERRAR")
End Sub

'------------------------------------------------------------------------------*
' Procedimiento  : BtnPagSiguiente
' Fecha          : ju., 16/may/2019 23:53:03
' Propósito      : Pagina a la siguiente página del filtro establecido
'------------------------------------------------------------------------------*'
Public Sub BtnPagSiguiente()
  On Error GoTo BtnPagSiguiente_Error
    
    If SrtCtrl Is Nothing Then
        Set SrtCtrl = New SorteoController
    End If
    '
    '   Invocamos al controlador para Visualizar página siguiente
    '
    SrtCtrl.BuscarNextPage
  
  On Error GoTo 0
BtnPagSiguiente__CleanExit:
    Exit Sub
            
BtnPagSiguiente_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    '   Audita el error
    Call HandleException(ErrNumber, ErrDescription, "SorteoUIConsultar.BtnPagSiguiente", ErrSource)
    '   Informa del error
    Call MsgBox(ErrDescription, vbCritical Or vbSystemModal, ThisWorkbook.Name)
    Call Trace("CERRAR")
End Sub
'------------------------------------------------------------------------------*
' Procedimiento  : BtnReset
' Fecha          : lu., 05/ago/2019 23:45:32
' Propósito      : Inicializa el filtro de búsqueda
'------------------------------------------------------------------------------*'
Public Sub BtnReset()
  On Error GoTo BtnReset_Error
    
    If SrtCtrl Is Nothing Then
        Set SrtCtrl = New SorteoController
    End If
    '
    '   Invocamos al controlador para inicializar el filtro
    '
    SrtCtrl.ResetFiltro
  
  On Error GoTo 0
BtnReset__CleanExit:
    Exit Sub
            
BtnReset_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    '   Audita el error
    Call HandleException(ErrNumber, ErrDescription, "SorteoUIConsultar.BtnReset", ErrSource)
    '   Informa del error
    Call MsgBox(ErrDescription, vbCritical Or vbSystemModal, ThisWorkbook.Name)
    Call Trace("CERRAR")
    
End Sub
' *===========(EOF): SorteoUIConsultar.cls
