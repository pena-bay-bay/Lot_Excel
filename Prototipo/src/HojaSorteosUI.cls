VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "HojaSorteosUI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
' *============================================================================*
' *
' *     Fichero    : HojaSorteoUI.cls
' *
' *     Autor      : Carlos Almela Baeza
' *     Creación   : do., 03/mar/2019 20:23:30
' *     Versión    : 1.0
' *     Propósito  : Interfaz de usuario del CU Definir Sorteo
' *
' *============================================================================*
Option Explicit
Option Base 0
'
'--- Variables Privadas -------------------------------------------------------*
Private SrtCtrl     As SorteoController
Private mId         As Integer
Private mPeriodo    As Periodo

'--- Constantes ---------------------------------------------------------------*
Private Const RgContexto As String = "$AC$1:$AC$4"
Private Const RgSorteoEditar As String = "$E$4:$R$20"
Private Const RgSorteoConsultar As String = "$T$4:$Z$20"
'--- Mensajes -----------------------------------------------------------------*
'--- Errores ------------------------------------------------------------------*
'--- Metodos Privados ---------------------------------------------------------*
Private Sub Worksheet_Activate()
    Set SrtCtrl = New SorteoController
End Sub
Private Sub Worksheet_Deactivate()
    Set SrtCtrl = Nothing
End Sub
'------------------------------------------------------------------------------*
' Procedimiento  : Worksheet_BeforeDoubleClick
' Fecha          : 03/mar/2019
' Propósito      : Intercepta el doble click del raton para interactuar con
'                  areas sensibles
' Parámetros     : Target Rango donde se ha pulsado el doble click
'                  Cancel indicador de cancelación del docle click
'------------------------------------------------------------------------------*
'
Private Sub Worksheet_BeforeDoubleClick(ByVal Target As RAnge, Cancel As Boolean)
  
  On Error GoTo Worksheet_BeforeDoubleClick_Error
    '
    '   Si la fila está contenida en la rejilla de datos
    '
    If (Target.Column >= 2) And (Target.Column <= 8) _
    And (Target.Row >= 12) And (Target.Row <= 18) Then
        '
        '   Obtenemos el identificador del registro
        '
        mId = CInt(Target.Offset(0, 2 - Target.Column).Value)
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
    Call HandleException(ErrNumber, ErrDescription, "HojaSorteosUI.Worksheet_BeforeDoubleClick", ErrSource)
    '   Informa del error
    Call MsgBox(ErrDescription, vbCritical Or vbSystemModal, ThisWorkbook.Name)
    Call Trace("CERRAR")
End Sub
'------------------------------------------------------------------------------*
' Procedimiento  : Worksheet_Change
' Fecha          :
' Propósito      : Interceptar el evento de cambio del contenido de ciertas
'                  celdas para modificar contenido de otras
'------------------------------------------------------------------------------*'
'
Private Sub Worksheet_Change(ByVal Target As RAnge)
    
  On Error GoTo Worksheet_Change_Error
    '
    '   Comprobamos la selección del periodo
    '
    If Target.Address = "$W$7" Then
        '
        '   Creamos el periodo si no existe
        '
        If mPeriodo Is Nothing Then
            Set mPeriodo = GetPeriodo
        End If
        '
        '   Asignamos el texto del periodo a la clase
        '
        mPeriodo.Texto = Target.Value
        '
        '   Si el contenido está vacío, vaciamos las celdas relacionadas
        '
        If IsEmpty(Target.Value) Then
            Target.Offset(1, 0).ClearContents
            Target.Offset(2, 0).ClearContents
        Else
            Target.Offset(1, 0).Value = mPeriodo.FechaInicial
            Target.Offset(2, 0).Value = mPeriodo.FechaFinal
        End If
    End If
    '
    '   Detectamos el cambio en la celda de Juego si estamos en
    '   Modo Alta
    '
    If Target.Address = "$I$7" Then
        If Me.ModoSorteo = mdAlta Then
            '
            '   Creamos el controlador si no está creado
            '
            If SrtCtrl Is Nothing Then
                Set SrtCtrl = New SorteoController
            End If
            '
            '   Invocamos a establecer juego si se asigna un juego
            '
            If Len(Target.Value) > 0 Then
                SrtCtrl.SetJuego Target.Value
            End If
        Else
            '
            '   En modo modificación no podemos cambiar el tipo de juego
            '
            Exit Sub
        End If
    End If
  
  On Error GoTo 0
Worksheet_Change__CleanExit:
    Exit Sub
            
Worksheet_Change_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    '   Audita el error
    Call HandleException(ErrNumber, ErrDescription, "HojaSorteosUI.Worksheet_Change", ErrSource)
    '   Informa del error
    Call MsgBox(ErrDescription, vbCritical Or vbSystemModal, ThisWorkbook.Name)
    Call Trace("CERRAR")
 End Sub


'------------------------------------------------------------------------------*
' Procedimiento  : SetModo
' Fecha          : ma., 22/oct/2019 13:29:05
' Propósito      : Establecer la configuración de la hoja según el modo
'                  seleccionado
' Parámetros     : Modo de edición del Caso de Uso
'------------------------------------------------------------------------------*
'
Private Sub SetModo(datModo As Modo)
    Dim mWrksheet As Worksheet
    Dim mShape As Shape
    
  On Error GoTo SetModo_Error
    '
    '   Establecemos la hoja de Sorteos
    '
    Set mWrksheet = Worksheets("Sorteos")
    '
    '   Establecemos el modo
    '
    Me.ModoSorteo = datModo
    '
    '   cambiamos la visualización de la hoja
    '
    Select Case datModo
        Case mdAlta
            Columns("T:AA").Select
            Selection.EntireColumn.Hidden = True
            
            Columns("E:R").Select
            Selection.EntireColumn.Hidden = False
            
        Case mdEdicion
            Columns("T:AA").Select
            Selection.EntireColumn.Hidden = True
            
            Columns("E:R").Select
            Selection.EntireColumn.Hidden = False
        
        Case mdConsulta
            Columns("T:AA").Select
            Selection.EntireColumn.Hidden = False
            
            Columns("E:R").Select
            Selection.EntireColumn.Hidden = True
    End Select
    '
    '   Cambiamos el literal de los Botones
    '
    Select Case datModo
        Case mdAlta, mdEdicion:
            Set mShape = mWrksheet.Shapes("btnSorteo1")
            mShape.TextFrame2.TextRange.Characters.Text = "Agregar"
            Set mShape = mWrksheet.Shapes("btnSorteo2")
            mShape.TextFrame2.TextRange.Characters.Text = "Guardar"
            Set mShape = mWrksheet.Shapes("btnSorteo3")
            mShape.TextFrame2.TextRange.Characters.Text = "Eliminar"
            Set mShape = mWrksheet.Shapes("btnSorteo4")
            mShape.TextFrame2.TextRange.Characters.Text = "Consultar"
        
        Case mdConsulta
            Set mShape = mWrksheet.Shapes("btnSorteo1")
            mShape.TextFrame2.TextRange.Characters.Text = "Buscar"
            Set mShape = mWrksheet.Shapes("btnSorteo2")
            mShape.TextFrame2.TextRange.Characters.Text = "Editar"
            Set mShape = mWrksheet.Shapes("btnSorteo3")
            mShape.TextFrame2.TextRange.Characters.Text = "Agregar"
            Set mShape = mWrksheet.Shapes("btnSorteo4")
            mShape.TextFrame2.TextRange.Characters.Text = "Resetear"

    End Select
    
  On Error GoTo 0
SetModo__CleanExit:
    Exit Sub
SetModo_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    '   Audita el error
    Call HandleException(ErrNumber, ErrDescription, "HojaSorteosUI.SetModo", ErrSource)
    '   Informa del error
    Call MsgBox(ErrDescription, vbCritical Or vbSystemModal, ThisWorkbook.Name)
    Call Trace("CERRAR")
End Sub

'--- Metodos Publicos ---------------------------------------------------------*
'------------------------------------------------------------------------------*
' Procedimiento  : btnSorteo1
' Fecha          : ma., 22/oct/2019 13:34:48
' Propósito      : Funcionalidad del Boton 1 ( Agregar | Buscar )
'------------------------------------------------------------------------------*
'
Public Sub btnSorteo1()
  On Error GoTo btnSorteo1_Error
    '   Anulamos la actualizacion
    '
    Application.ScreenUpdating = False
    '
    '   Creamos el controlador
    '
    If SrtCtrl Is Nothing Then
        Set SrtCtrl = New SorteoController
    End If
    '
    '   Analizamos el modo actual
    '
    If (Me.ModoSorteo = mdAlta) Or (Me.ModoSorteo = mdEdicion) Then
        '
        '   Función agregar registro
        '
        SetModo mdAlta
        '
        '   Trasladamos la función al Controlador
        '
        SrtCtrl.Nuevo
        RAnge("I7").Select
        
    ElseIf (Me.ModoSorteo = mdConsulta) Then
        '
        '   Función buscar registros
        '
        SetModo mdConsulta
        '
        '   Trasladamos la función al Controlador
        '
        SrtCtrl.Buscar
        '
        'TODO: Posicionar cursor en rejilla
    End If
    '   Activamos la actualizacion
    '
    Application.ScreenUpdating = True
    
  On Error GoTo 0
btnSorteo1__CleanExit:
    Exit Sub

btnSorteo1_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    Call HandleException(ErrNumber, ErrDescription, "HojaSorteosUI.btnSorteo1", ErrSource)
    Call MsgBox(ErrDescription, vbCritical Or vbSystemModal, ThisWorkbook.Name)
    Call Trace("CERRAR")
End Sub
'------------------------------------------------------------------------------*
' Procedimiento  : btnSorteo2
' Fecha          : ma., 22/oct/2019 18:19:18
' Propósito      : Funcionalidad del Boton 2 ( Guardar | Buscar )
'------------------------------------------------------------------------------*
'
Public Sub btnSorteo2()
  On Error GoTo btnSorteo2_Error
    '   Anulamos la actualizacion
    '
    Application.ScreenUpdating = False
    '
    '   Creamos el controlador
    '
    If SrtCtrl Is Nothing Then
        Set SrtCtrl = New SorteoController
    End If
    '
    '   Analizamos el modo actual
    '
    If (Me.ModoSorteo = mdAlta) Or (Me.ModoSorteo = mdEdicion) Then
        '
        '   Trasladamos la función al Controlador
        '
        SrtCtrl.Guardar
        '
        '   Establecemos el modo Edición
        '
        SetModo mdEdicion
        RAnge("I7").Select
        
    ElseIf (Me.ModoSorteo = mdConsulta) Then
        '
        '   Función buscar registros
        '
        SetModo mdConsulta
        '
        '   Trasladamos la función al Controlador
        '
        SrtCtrl.Buscar
        '
        'TODO: Posicionar cursor en rejilla
    End If
    '   Activamos la actualizacion
    '
    Application.ScreenUpdating = True

  On Error GoTo 0
btnSorteo2__CleanExit:
    Exit Sub
btnSorteo2_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    '   Audita el error
    Call HandleException(ErrNumber, ErrDescription, "HojaSorteosUI.btnSorteo2", ErrSource)
    '   Informa del error
    Call MsgBox(ErrDescription, vbCritical Or vbSystemModal, ThisWorkbook.Name)
    Call Trace("CERRAR")
End Sub
'------------------------------------------------------------------------------*
' Procedimiento  : btnSorteo3
' Fecha          : ma., 22/oct/2019 18:19:18
' Propósito      : Funcionalidad del Boton 1 ( Eliminar | Agregar)
'------------------------------------------------------------------------------*
'
Public Sub btnSorteo3()
  On Error GoTo btnSorteo3_Error
    '   Anulamos la actualizacion
    '
    Application.ScreenUpdating = False
    '
    '   Creamos el controlador
    '
    If SrtCtrl Is Nothing Then
        Set SrtCtrl = New SorteoController
    End If
    '
    '   Analizamos el modo actual
    '
    If (Me.ModoSorteo = mdAlta) Or (Me.ModoSorteo = mdEdicion) Then
        '
        '   Obtenemos el Id
        '
        mId = CInt(Me.RAnge("I6").Value)
        '
        '   Función Eliminar Registro
        '
        SrtCtrl.Eliminar mId
        '
        '   Establecemos el modo consulta
        '
        SetModo mdConsulta
        RAnge("W5").Select
        
    ElseIf (Me.ModoSorteo = mdConsulta) Then
        '
        '   Función Nuevo registro
        '
        SetModo mdAlta
        '
        '   Trasladamos la función al Controlador
        '
        SrtCtrl.Nuevo
        RAnge("I7").Select
    End If
    '   Activamos la actualizacion
    '
    Application.ScreenUpdating = True

  On Error GoTo 0
btnSorteo3__CleanExit:
    Exit Sub

btnSorteo3_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    '   Audita el error
    Call HandleException(ErrNumber, ErrDescription, "HojaSorteosUI.btnSorteo3", ErrSource)
    '   Informa del error
    Call MsgBox(ErrDescription, vbCritical Or vbSystemModal, ThisWorkbook.Name)
    Call Trace("CERRAR")
End Sub
'------------------------------------------------------------------------------*
' Procedimiento  : btnSorteo4
' Fecha          : ma., 22/oct/2019 18:19:18
' Propósito      : Funcionalidad del Boton 4 ( Consultar | Resetear )
'------------------------------------------------------------------------------*
'
Public Sub btnSorteo4()
  On Error GoTo btnSorteo4_Error
    '   Anulamos la actualizacion
    '
    Application.ScreenUpdating = False
    '
    '   Creamos el controlador
    '
    If SrtCtrl Is Nothing Then
        Set SrtCtrl = New SorteoController
    End If
    '
    '   Analizamos el modo actual
    '
    If (Me.ModoSorteo = mdAlta) Or (Me.ModoSorteo = mdEdicion) Then
        '
        '   Saltamos a la pantalla de consulta
        '
        SetModo mdConsulta
        '
        '   Invocamos el reseteo del filtro
        '
        SrtCtrl.ResetFiltro
        RAnge("W5").Select
        
    ElseIf (Me.ModoSorteo = mdConsulta) Then
        '
        '   Función buscar registros
        '
        SetModo mdConsulta
        '
        '   Trasladamos la función al Controlador
        '
        SrtCtrl.ResetFiltro
        '
        '   Posicionamos el cursor en filtro Juego
        '
        RAnge("W5").Select
    End If
    '   Activamos la actualizacion
    '
    Application.ScreenUpdating = True

  On Error GoTo 0
btnSorteo4__CleanExit:
    Exit Sub
btnSorteo4_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    '   Audita el error
    Call HandleException(ErrNumber, ErrDescription, "HojaSorteosUI.btnSorteo4", ErrSource)
    '   Informa del error
    Call MsgBox(ErrDescription, vbCritical Or vbSystemModal, ThisWorkbook.Name)
    Call Trace("CERRAR")
End Sub


'------------------------------------------------------------------------------*
' Procedimiento  : btnFirstSorteo
' Fecha          : ma., 22/oct/2019 18:19:18
' Propósito      : Funcionalidad Ir al primer registro o primera página
'                  segun el modo
'------------------------------------------------------------------------------*
'
Public Sub btnFirstSorteo()
  On Error GoTo btnFirstSorteo_Error
    '   Anulamos la actualizacion
    '
    Application.ScreenUpdating = False
    '
    '   Creamos el controlador
    '
    If SrtCtrl Is Nothing Then
        Set SrtCtrl = New SorteoController
    End If
    '
    '   Si estamos en modo consulta realizamos una paginación
    '
    If Me.ModoSorteo = mdConsulta Then
        '
        '   Invocamos al controlador para Visualizar la primera página
        '
        SrtCtrl.BuscarFirstPage
    Else
        '
        '   Si estamos en modo alta o modifiación vamos al primero
        '
        SrtCtrl.IrAlPrimero
    End If
    '   Activamos la actualizacion
    '
    Application.ScreenUpdating = True
    
  On Error GoTo 0
btnFirstSorteo__CleanExit:
    Exit Sub
            
btnFirstSorteo_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    '   Audita el error
    Call HandleException(ErrNumber, ErrDescription, "HojaSorteosUI.btnFirstSorteo", ErrSource)
    '   Informa del error
    Call MsgBox(ErrDescription, vbCritical Or vbSystemModal, ThisWorkbook.Name)
    Call Trace("CERRAR")
End Sub
'------------------------------------------------------------------------------*
' Procedimiento  : btnLastSorteo
' Fecha          : ma., 22/oct/2019 18:20:08
' Propósito      : Funcionalidad Ir al ultimo registro o ultima página
'                  segun el modo
'------------------------------------------------------------------------------*
'
Public Sub btnLastSorteo()
  On Error GoTo btnLastSorteo_Error
    '   Anulamos la actualizacion
    '
    Application.ScreenUpdating = False
    '
    '   Creamos el controlador
    '
    If SrtCtrl Is Nothing Then
        Set SrtCtrl = New SorteoController
    End If
    '
    '   Si estamos en modo consulta realizamos una paginación
    '
    If Me.ModoSorteo = mdConsulta Then
        '
        '   Invocamos al controlador para Visualizar la ultima página
        '
        SrtCtrl.BuscarLastPage
    Else
        '
        '   Si estamos en modo alta o modifiación vamos al ultimo
        '
        SrtCtrl.IrAlUltimo
    End If
    '   Activamos la actualizacion
    '
    Application.ScreenUpdating = True
  
  On Error GoTo 0
btnLastSorteo__CleanExit:
    Exit Sub
            
btnLastSorteo_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    '   Audita el error
    Call HandleException(ErrNumber, ErrDescription, "HojaSorteosUI.btnLastSorteo", ErrSource)
    '   Informa del error
    Call MsgBox(ErrDescription, vbCritical Or vbSystemModal, ThisWorkbook.Name)
    Call Trace("CERRAR")
End Sub
'------------------------------------------------------------------------------*
' Procedimiento  : btnNextSorteo
' Fecha          : ma., 22/oct/2019 18:23:00
' Propósito      : Funcionalidad Ir al siguiente registro o siguiente página
'                  segun el modo
'------------------------------------------------------------------------------*
'
Public Sub btnNextSorteo()
  On Error GoTo btnNextSorteo_Error
    '   Anulamos la actualizacion
    '
    Application.ScreenUpdating = False
    '
    '   Creamos el controlador si no lo esta
    '
    If SrtCtrl Is Nothing Then
        Set SrtCtrl = New SorteoController
    End If
    '
    '   Si estamos en modo consulta realizamos una paginación
    '
    If Me.ModoSorteo = mdConsulta Then
        '
        '   Invocamos al controlador para Visualizar la ultima página
        '
        SrtCtrl.BuscarNextPage
    Else
        '
        '   Obtenemos el Id del registro actual
        '
        mId = CInt(Me.RAnge("I6").Value)
        '
        '   Si estamos en modo alta o modifiación vamos al ultimo
        '
        SrtCtrl.Siguiente mId
    End If
    '   Activamos la actualizacion
    '
    Application.ScreenUpdating = True
  
  On Error GoTo 0
btnNextSorteo__CleanExit:
    Exit Sub
            
btnNextSorteo_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    '   Audita el error
    Call HandleException(ErrNumber, ErrDescription, "HojaSorteosUI.btnNextSorteo", ErrSource)
    '   Informa del error
    Call MsgBox(ErrDescription, vbCritical Or vbSystemModal, ThisWorkbook.Name)
    Call Trace("CERRAR")
End Sub

'------------------------------------------------------------------------------*
' Procedimiento  : btnPrevSorteo
' Fecha          : ma., 22/oct/2019 18:23:00
' Propósito      : Funcionalidad Ir al anterior registro o anterior página
'                  segun el modo
'------------------------------------------------------------------------------*
'
Public Sub btnPrevSorteo()
  On Error GoTo btnPrevSorteo_Error
    '   Anulamos la actualizacion
    '
    Application.ScreenUpdating = False
    '
    '   Creamos el controlador si no lo esta
    '
    If SrtCtrl Is Nothing Then
        Set SrtCtrl = New SorteoController
    End If
    '
    '   Si estamos en modo consulta realizamos una paginación
    '
    If Me.ModoSorteo = mdConsulta Then
        '
        '   Invocamos al controlador para Visualizar la ultima página
        '
        SrtCtrl.BuscarPrevPage
    Else
        '
        '   Obtenemos el Id del registro actual
        '
        mId = CInt(Me.RAnge("I6").Value)
        '
        '   Si estamos en modo alta o modifiación vamos al ultimo
        '
        SrtCtrl.Anterior mId
    End If
    '   Activamos la actualizacion
    '
    Application.ScreenUpdating = True
  
  
  On Error GoTo 0
btnPrevSorteo__CleanExit:
    Exit Sub
            
btnPrevSorteo_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    '   Audita el error
    Call HandleException(ErrNumber, ErrDescription, "HojaSorteosUI.btnPrevSorteo", ErrSource)
    '   Informa del error
    Call MsgBox(ErrDescription, vbCritical Or vbSystemModal, ThisWorkbook.Name)
    Call Trace("CERRAR")
End Sub
'--- Propiedades --------------------------------------------------------------*
'------------------------------------------------------------------------------*
' Propiedad      : ModoSorteo
' Fecha          : vi., 18/oct/2019 18:48:58
' Propósito      : Modo del caso de uso Definir Sorteo:
'                 { Alta, Modificación, Consulta }
'------------------------------------------------------------------------------*
'
Public Property Get ModoSorteo() As Modo
    Dim mSrt As String
    Dim mRg As RAnge
    
    Set mRg = RAnge("A1")
    mSrt = mRg.Value
    Select Case mSrt
        Case LT_ALTA: ModoSorteo = mdAlta
        Case LT_EDICION: ModoSorteo = mdEdicion
        Case LT_CONSULTA: ModoSorteo = mdConsulta
        Case Else
            ModoSorteo = mdSinDefinir
    End Select
End Property

Public Property Let ModoSorteo(ByVal vNewValue As Modo)
    Dim mRg As RAnge
    Set mRg = RAnge("A1")
    Select Case vNewValue
        Case mdAlta: mRg.Value = LT_ALTA
        Case mdEdicion: mRg.Value = LT_EDICION
        Case mdConsulta: mRg.Value = LT_CONSULTA
        Case Else: mRg.Value = ""
    End Select
End Property

'' *===========(EOF): HojaSorteoUI.cls
