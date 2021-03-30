VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SorteoUIEditar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
' *============================================================================*
' *
' *     Fichero    : SorteoUIEditar.cls
' *
' *     Autor      : Carlos Almela Baeza
' *     Creación   : sá., 11/may/2019 18:26:00
' *     Versión    : 1.0
' *     Propósito  : Hoja UI Editar un sorteo
' *
' *============================================================================*
Option Explicit
Option Base 0

'--- Variables Privadas -------------------------------------------------------*
Private SrtCtrl     As SorteoController
Private mId         As Integer

'--- Métodos Privados ---------------------------------------------------------*
Private Sub Worksheet_Activate()
    Set SrtCtrl = New SorteoController
End Sub
'------------------------------------------------------------------------------*
' Procedimiento  : Worksheet_Change
' Fecha          : vi., 26/jul/2019 21:14:25
' Propósito      : Capturar el evento cambio de juego
'------------------------------------------------------------------------------*'
Private Sub Worksheet_Change(ByVal Target As Range)
    '
    '   Detectamos el cambio en la celda de Juego
    '
    If Target.Address = "$C$5" Then
        '
        '   Creamos el controlador si no está creado
        '
        If SrtCtrl Is Nothing Then
            Set SrtCtrl = New SorteoController
        End If
        '
        '   Invocamos a establecer juego si se asigna un juego
        '
        If Len(Target.value) > 0 Then
            SrtCtrl.SetJuego Target.value
        End If
    End If
End Sub

Private Sub Worksheet_Deactivate()
    Set SrtCtrl = Nothing
End Sub

'------------------------------------------------------------------------------*
' Procedimiento  : BtnGuardar
' Fecha          : sá., 11/may/2019 18:21:28
' Propósito      : Guarda el actual registro en el repositorio
'------------------------------------------------------------------------------*'
Public Sub BtnGuardar()
  On Error GoTo BtnGuardar_Error
    
    If SrtCtrl Is Nothing Then
        Set SrtCtrl = New SorteoController
    End If
    '
    '   Invocamos al controlador para que guarde el actual registro
    '
    SrtCtrl.Guardar
    
  On Error GoTo 0
BtnGuardar__CleanExit:
    Exit Sub
            
BtnGuardar_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    '   Audita el error
    Call HandleException(ErrNumber, ErrDescription, "SorteoController.BtnGuardar", ErrSource)
    '   Informa del error
    Call MsgBox(ErrDescription, vbCritical Or vbSystemModal, ThisWorkbook.Name)
    Call Trace("CERRAR")
End Sub

'------------------------------------------------------------------------------*
' Procedimiento  : BtnNuevo
' Fecha          : sá., 11/may/2019 18:21:28
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
    Call HandleException(ErrNumber, ErrDescription, "SorteoUIEditar.BtnNuevo", ErrSource)
    '   Informa del error
    Call MsgBox(ErrDescription, vbCritical Or vbSystemModal, ThisWorkbook.Name)
    Call Trace("CERRAR")
End Sub

'------------------------------------------------------------------------------*
' Procedimiento  : BtnEliminar
' Fecha          : sá., 11/may/2019 18:21:28
' Propósito      : Elimina el registro actual
'------------------------------------------------------------------------------*'
Public Sub BtnEliminar()
  On Error GoTo BtnEliminar_Error
  
    If SrtCtrl Is Nothing Then
        Set SrtCtrl = New SorteoController
    End If
    '
    '   Invocamos al controlador para que elimine el actual registro
    '
    mId = CInt(Me.Range("C4").value)
    SrtCtrl.Eliminar mId
  
  On Error GoTo 0
BtnEliminar__CleanExit:
    Exit Sub
            
BtnEliminar_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    '   Audita el error
    Call HandleException(ErrNumber, ErrDescription, "SorteoUIEditar.BtnEliminar", ErrSource)
    '   Informa del error
    Call MsgBox(ErrDescription, vbCritical Or vbSystemModal, ThisWorkbook.Name)
    Call Trace("CERRAR")
End Sub


'------------------------------------------------------------------------------*
' Procedimiento  : BtnIrPrimero
' Fecha          : sá., 11/may/2019 18:21:28
' Propósito      : Visualizar el primer registro del repositorio
'------------------------------------------------------------------------------*'
Public Sub BtnIrPrimero()
  On Error GoTo BtnIrPrimero_Error
    
    If SrtCtrl Is Nothing Then
        Set SrtCtrl = New SorteoController
    End If
    '
    '   Invocamos al controlador para que nos visualice el primer registro
    '
    SrtCtrl.IrAlPrimero
  
  On Error GoTo 0
BtnIrPrimero__CleanExit:
    Exit Sub
            
BtnIrPrimero_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    '   Audita el error
    Call HandleException(ErrNumber, ErrDescription, "SorteoUIEditar.BtnIrPrimero", ErrSource)
    '   Informa del error
    Call MsgBox(ErrDescription, vbCritical Or vbSystemModal, ThisWorkbook.Name)
    Call Trace("CERRAR")
End Sub
'------------------------------------------------------------------------------*
' Procedimiento  : BtnIrUltimo
' Fecha          : sá., 11/may/2019 18:21:28
' Propósito      : Visualizar el último registro del repositorio
'------------------------------------------------------------------------------*'
Public Sub BtnIrUltimo()
  On Error GoTo BtnIrUltimo_Error
    
    If SrtCtrl Is Nothing Then
        Set SrtCtrl = New SorteoController
    End If
    '
    '   Invocamos al controlador para que nos visualice el último registro
    '
    SrtCtrl.IrAlUltimo
  
  On Error GoTo 0
BtnIrUltimo__CleanExit:
    Exit Sub
            
BtnIrUltimo_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    '   Audita el error
    Call HandleException(ErrNumber, ErrDescription, "SorteoUIEditar.BtnIrUltimo", ErrSource)
    '   Informa del error
    Call MsgBox(ErrDescription, vbCritical Or vbSystemModal, ThisWorkbook.Name)
    Call Trace("CERRAR")
End Sub

'------------------------------------------------------------------------------*
' Procedimiento  : BtnIrSiguiente
' Fecha          : sá., 11/may/2019 18:21:28
' Propósito      : Visualizar el registro siguiente
'------------------------------------------------------------------------------*'
Public Sub BtnIrSiguiente()
  On Error GoTo BtnIrSiguiente_Error
    
    If SrtCtrl Is Nothing Then
        Set SrtCtrl = New SorteoController
    End If
    '
    '   Invocamos al controlador para que nos visualice el registro siguiente
    '
    mId = CInt(Me.Range("C4").value)
    SrtCtrl.Siguiente mId
  
  On Error GoTo 0
BtnIrSiguiente__CleanExit:
    Exit Sub
            
BtnIrSiguiente_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    '   Audita el error
    Call HandleException(ErrNumber, ErrDescription, "SorteoUIEditar.BtnIrSiguiente", ErrSource)
    '   Informa del error
    Call MsgBox(ErrDescription, vbCritical Or vbSystemModal, ThisWorkbook.Name)
    Call Trace("CERRAR")
End Sub
'------------------------------------------------------------------------------*
' Procedimiento  : BtnIrAnterior
' Fecha          : sá., 11/may/2019 18:20:28
' Propósito      : Visualizar el registro anterior
'------------------------------------------------------------------------------*'
Public Sub BtnIrAnterior()
  On Error GoTo BtnIrAnterior_Error
    
    If SrtCtrl Is Nothing Then
        Set SrtCtrl = New SorteoController
    End If
    '
    '   Invocamos al controlador para que nos visualice el registro anterior
    '
    mId = CInt(Me.Range("C4").value)
    SrtCtrl.Anterior mId
      
  On Error GoTo 0
BtnIrAnterior__CleanExit:
    Exit Sub
            
BtnIrAnterior_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    '   Audita el error
    Call HandleException(ErrNumber, ErrDescription, "SorteoUIEditar.BtnIrAnterior", ErrSource)
    '   Informa del error
    Call MsgBox(ErrDescription, vbCritical Or vbSystemModal, ThisWorkbook.Name)
    Call Trace("CERRAR")
End Sub
'' *===========(EOF): SorteoUIEditar.cls
