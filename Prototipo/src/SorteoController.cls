VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SorteoController"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' *============================================================================*
' *
' *     Fichero    : SorteoController.cls
' *
' *     Autor      : Carlos Almela Baeza
' *     Creación   : mi., 17/abr/2019 22:59:32
' *     Versión    : 1.0
' *     Propósito  : Controlador de sorteo
' *
' *============================================================================*
Option Explicit
Option Base 0
'
'--- Variables Privadas -------------------------------------------------------*
Private mModel          As SorteoModel                  ' Modelo de negocio
Private mView           As SorteoView                   ' Vista del CU
Private mPagActual      As Integer                      ' Página Actual de datos

'--- Métodos Privados ---------------------------------------------------------*
'
'   Creamos el modelo y la vista
'
Private Sub Class_Initialize()
    Set mModel = GetSorteoModel
    Set mView = New SorteoView
End Sub
'
'
'
Private Sub Class_Terminate()
    Set mModel = Nothing
    Set mView = Nothing
End Sub
'-- Propiedades
'--- Métodos Públicos ---------------------------------------------------------*
'------------------------------------------------------------------------------*
' Procedimiento  : Guardar
' Fecha          :
' Propósito      : Invoca al modelo para que guarde el Sorteo de la interfaz
'------------------------------------------------------------------------------*
'
Public Sub Guardar()
  On Error GoTo Guardar_Error
    '
    '   Invoca al modelo que guarde el Sorteo que devuelve la interfaz
    '
    If mModel.GuardarSorteoRecord(mView.GetDisplaySorteo) Then
        '
        '   Mensaje confirmación
        MsgBox "El registro ha sido guardado correctamente.", vbOK + vbInformation, ThisWorkbook.Name
        Exit Sub
    End If
  
  On Error GoTo 0
Guardar__CleanExit:
    Exit Sub
Guardar_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    Call HandleException(ErrNumber, ErrDescription, "SorteoController.Guardar", ErrSource)
    Err.Raise ErrNumber, "SorteoController.Guardar", ErrDescription
End Sub
'------------------------------------------------------------------------------*
' Procedimiento  : Nuevo
' Fecha          :
' Propósito      : Invoca al modelo para que cree un nuevo Sorteo
'------------------------------------------------------------------------------*
Public Sub Nuevo()
  On Error GoTo Nuevo_Error
    '
    '   Inicializamos la rejilla de consulta
    '
    mView.ClearGrid
    mView.ClearFiltros
    '
    '   Inicializamos el area de trabajo
    '
    mView.ClearSorteoDisplay True
    '
    '   Invoca al modelo que cree un nuevo sorteo
    '
    mModel.NuevoSorteoRecord
    '
    '   Pasamos el modelo con el nuevo objeto a la vista
    '
    mView.DisplaySorteoRecord mModel
    '
    '   Seleccionamos la hoja de edición
    '
    'ThisWorkbook.Sheets("Editar").Activate
    'RAnge("C5").Select
  On Error GoTo 0
Nuevo__CleanExit:
    Exit Sub
Nuevo_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    Call HandleException(ErrNumber, ErrDescription, "SorteoController.Nuevo", ErrSource)
    Err.Raise ErrNumber, "SorteoController.Nuevo", ErrDescription
End Sub
'------------------------------------------------------------------------------*
' Procedimiento  : Eliminar
' Fecha          :
' Propósito      : Elimina el sorteo actual
'------------------------------------------------------------------------------*
Public Sub Eliminar(ByVal IdSorteo As Integer)
  On Error GoTo Eliminar_Error
    '
    '   Solicitar confirmación
    '
    If MsgBox("¿Está seguro que quiere elminar este registro?", vbQuestion + vbYesNo, _
                ThisWorkbook.Name) = vbYes Then

        '
        '   Invoca al modelo que elimine el Sorteo
        '
        mModel.EliminarSorteoRecord IdSorteo
        '
        '
        '
        mView.ClearSorteoDisplay False
    End If
  On Error GoTo 0
Eliminar__CleanExit:
    Exit Sub
Eliminar_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    Call HandleException(ErrNumber, ErrDescription, "SorteoController.Eliminar", ErrSource)
    Err.Raise ErrNumber, "SorteoController.Eliminar", ErrDescription
End Sub
'------------------------------------------------------------------------------*
' Procedimiento  : Siguiente
' Fecha          :
' Propósito      : Visualiza el sorteo siguiente al actual
'------------------------------------------------------------------------------*
Public Sub Siguiente(ByVal IdSorteo As Integer)
  On Error GoTo Siguiente_Error
    '
    '   Invoca al modelo que carge el siguiente registro
    '
    If mModel.GetNextSorteoRecord(IdSorteo) Then
        '
        '   Pasamos el modelo con el nuevo objeto a la vista
        '
        mView.DisplaySorteoRecord mModel
    Else
        '
        '   Mensaje y salir
        '
        MsgBox "Es el último sorteo.", vbOK + vbInformation, ThisWorkbook.Name
        Exit Sub
    End If
  On Error GoTo 0
Siguiente__CleanExit:
    Exit Sub
Siguiente_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    Call HandleException(ErrNumber, ErrDescription, "SorteoController.Siguiente", ErrSource)
    Err.Raise ErrNumber, "SorteoController.Siguiente", ErrDescription
End Sub
'------------------------------------------------------------------------------*
' Procedimiento  : Anterior
' Fecha          :
' Propósito      : Visualiza el sorteo anterior al Id actual
'------------------------------------------------------------------------------*
Public Sub Anterior(ByVal IdSorteo As Integer)
  On Error GoTo Anterior_Error
    '
    '   Invoca al modelo que carge el anterior registro
    '
    If mModel.GetPrevSorteoRecord(IdSorteo) Then
        '
        '   Pasamos el modelo con el nuevo objeto a la vista
        '
        mView.DisplaySorteoRecord mModel
    Else
        '
        '   Mensaje y salir
        '
        MsgBox "Es el primer sorteo.", vbOK + vbInformation, ThisWorkbook.Name
        Exit Sub
    End If
  On Error GoTo 0
Anterior__CleanExit:
    Exit Sub
Anterior_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    Call HandleException(ErrNumber, ErrDescription, "SorteoController.Anterior", ErrSource)
    Err.Raise ErrNumber, "SorteoController.Anterior", ErrDescription
End Sub
'------------------------------------------------------------------------------*
' Procedimiento  : IrAlPrimero
' Fecha          :
' Propósito      : Visualiza el primer Sorteo
'------------------------------------------------------------------------------*
Public Sub IrAlPrimero()
  On Error GoTo IrAlPrimero_Error
    '
    '   Invoca al modelo que carge el primer registro
    '
    mModel.GetFirstSorteo
    '
    '   Pasamos el modelo con el nuevo objeto a la vista
    '
    mView.DisplaySorteoRecord mModel
  
  On Error GoTo 0
IrAlPrimero__CleanExit:
    Exit Sub
IrAlPrimero_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    Call HandleException(ErrNumber, ErrDescription, "SorteoController.IrAlPrimero", ErrSource)
    Err.Raise ErrNumber, "SorteoController.IrAlPrimero", ErrDescription
End Sub
'------------------------------------------------------------------------------*
' Procedimiento  : IrAlUltimo
' Fecha          :
' Propósito      : Visualizar el último Sorteo
'------------------------------------------------------------------------------*
Public Sub IrAlUltimo()
  On Error GoTo IrAlUltimo_Error
    '
    '   Invoca al modelo que carge el ultimo registro
    '
    mModel.GetLastSorteo
    '
    '   Pasamos el modelo con el nuevo objeto a la vista
    '
    mView.DisplaySorteoRecord mModel
  
  On Error GoTo 0
IrAlUltimo__CleanExit:
    Exit Sub
IrAlUltimo_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    Call HandleException(ErrNumber, ErrDescription, "SorteoController.IrAlUltimo", ErrSource)
    Err.Raise ErrNumber, "SorteoController.IrAlUltimo", ErrDescription
End Sub
'------------------------------------------------------------------------------*
' Procedimiento  : Buscar
' Fecha          :
' Propósito      : Buscar sorteos con las condiciones del filtro
'------------------------------------------------------------------------------*
Public Sub Buscar()
  On Error GoTo Buscar_Error
    '
    '   Invoca al modelo que busque Sorteos con el filtro definido en la interfaz
    '
    Set mModel = mView.GetFiltroBusqueda
    '
    '   Indicamos la primera página de la busqueda
    '
    mModel.PaginaActual = 1
    '   Se invoca al metodo buscar del modelo
    '   y si encuentra elementos se visualizan
    '   Si no se da un mensaje
    '
    If mModel.SearchSorteos() Then
        '
        '   Inicializa la rejilla y visualiza los sorteos encontrados
        '
        mView.ClearGrid
        mView.AddSorteosToGrid mModel
    Else
        MsgBox "No se han encontrado Sorteos con los filtros indicados", vbOK + vbInformation, ThisWorkbook.Name
    End If
  
  On Error GoTo 0
Buscar__CleanExit:
    Exit Sub
Buscar_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    Call HandleException(ErrNumber, ErrDescription, "SorteoController.Buscar", ErrSource)
    Err.Raise ErrNumber, "SorteoController.Buscar", ErrDescription
End Sub
'------------------------------------------------------------------------------*
' Procedimiento  : BuscarFirstPage
' Fecha          :
' Propósito      : Ir a la primera página de la selección realizada
'------------------------------------------------------------------------------*
Public Sub BuscarFirstPage()
  On Error GoTo BuscarFirstPage_Error
    '
    '   Invoca al modelo que busque Sorteos con el filtro definido en la interfaz
    '
    Set mModel = mView.GetFiltroBusqueda
    '
    '   Obtenemos la página actual
    '
    mPagActual = mView.PaginaActual
    '
    '   Comprobamos el Inicio
    '
    If mPagActual = 1 Then
        '
        '   Mensaje y salir
        '
        MsgBox "Es la primera página para este filtro", vbOK + vbInformation, ThisWorkbook.Name
        Exit Sub
    End If
    '
    '   Vamos a la primera página
    '
    mModel.PaginaActual = 1
    '
    '   Se invoca al metodo buscar del modelo
    '   y si encuentra elementos se visualizan
    '   Si no se da un mensaje
    '
    If mModel.SearchSorteos() Then
        '
        '   Inicializa la rejilla y visualiza los sorteos encontrados
        '
        mView.ClearGrid
        mView.AddSorteosToGrid mModel
    Else
        MsgBox "No se han encontrado Sorteos con los filtros indicados", vbOK + vbInformation, ThisWorkbook.Name
    End If
  
  On Error GoTo 0
BuscarFirstPage__CleanExit:
    Exit Sub
BuscarFirstPage_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    Call HandleException(ErrNumber, ErrDescription, "SorteoController.BuscarFirstPage", ErrSource)
    Err.Raise ErrNumber, "SorteoController.BuscarFirstPage", ErrDescription
End Sub
'------------------------------------------------------------------------------*
' Procedimiento  : BuscarNextPage
' Fecha          :
' Propósito      : Presenta la siguiente página a la visualizada
'------------------------------------------------------------------------------*
Public Sub BuscarNextPage()
  On Error GoTo BuscarNextPage_Error
    '
    '   Invoca al modelo que busque Sorteos con el filtro definido en la interfaz
    '
    Set mModel = mView.GetFiltroBusqueda
    '
    '   Obtenemos la página actual
    '
    mPagActual = mView.PaginaActual
    '
    '   Comprobamos el final
    '
    If mPagActual >= mView.TotalPaginas Then
        '
        '   Mensaje y salir
        '
        MsgBox "Es la última página para este filtro", vbOK + vbInformation, ThisWorkbook.Name
        Exit Sub
    End If
    '
    '   Incrementamos la página
    '
    mModel.PaginaActual = mPagActual + 1
    '
    '   Se invoca al metodo buscar del modelo
    '   y si encuentra elementos se visualizan
    '   Si no se da un mensaje
    '
    If mModel.SearchSorteos() Then
        '
        '   Inicializa la rejilla y visualiza los sorteos encontrados
        '
        mView.ClearGrid
        mView.AddSorteosToGrid mModel
    Else
        MsgBox "No se han encontrado Sorteos con los filtros indicados", vbOK + vbInformation, ThisWorkbook.Name
    End If
  
  On Error GoTo 0
BuscarNextPage__CleanExit:
    Exit Sub
BuscarNextPage_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    Call HandleException(ErrNumber, ErrDescription, "SorteoController.BuscarNextPage", ErrSource)
    Err.Raise ErrNumber, "SorteoController.BuscarNextPage", ErrDescription
End Sub
'------------------------------------------------------------------------------*
' Procedimiento  : BuscarPrevPage
' Fecha          :
' Propósito      : Ir a la página anterior de la selección
'------------------------------------------------------------------------------*
Public Sub BuscarPrevPage()
  On Error GoTo BuscarPrevPage_Error
    '
    '   Invoca al modelo que busque Sorteos con el filtro definido en la interfaz
    '
    Set mModel = mView.GetFiltroBusqueda
    '
    '   Obtenemos la página actual
    '
    mPagActual = mView.PaginaActual
    '
    '   Comprobamos el Inicio
    '
    If mPagActual = 1 Then
        '
        '   Mensaje y salir
        '
        MsgBox "Es la primera página para este filtro", vbOK + vbInformation, ThisWorkbook.Name
        Exit Sub
    End If
    '
    '   Decrementamos  la página
    '
    mModel.PaginaActual = mPagActual - 1
    '
    '   Se invoca al metodo buscar del modelo
    '   y si encuentra elementos se visualizan
    '   Si no se da un mensaje
    '
    If mModel.SearchSorteos() Then
        '
        '   Inicializa la rejilla y visualiza los sorteos encontrados
        '
        mView.ClearGrid
        mView.AddSorteosToGrid mModel
    Else
        MsgBox "No se han encontrado Sorteos con los filtros indicados", vbOK + vbInformation, ThisWorkbook.Name
    End If
  
  On Error GoTo 0
BuscarPrevPage__CleanExit:
    Exit Sub
BuscarPrevPage_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    Call HandleException(ErrNumber, ErrDescription, "SorteoController.BuscarPrevPage", ErrSource)
    Err.Raise ErrNumber, "SorteoController.BuscarPrevPage", ErrDescription
End Sub
'------------------------------------------------------------------------------*
' Procedimiento  : BuscarLastPage
' Fecha          :
' Propósito      : Ir a la última página de la selección de datos
'------------------------------------------------------------------------------*
Public Sub BuscarLastPage()
  On Error GoTo BuscarLastPage_Error
    '
    '   Invoca al modelo que busque Sorteos con el filtro definido en la interfaz
    '
    Set mModel = mView.GetFiltroBusqueda
    '
    '   Obtenemos la página actual
    '
    mPagActual = mView.PaginaActual
    '
    '   Comprobamos el final
    '
    If mPagActual >= mView.TotalPaginas Then
        '
        '   Mensaje y salir
        '
        MsgBox "Es la última página para este filtro", vbOK + vbInformation, ThisWorkbook.Name
        Exit Sub
    End If
    '
    '   Vamos a la última página
    '
    mModel.PaginaActual = mView.TotalPaginas
    '
    '   Se invoca al metodo buscar del modelo
    '   y si encuentra elementos se visualizan
    '   Si no se da un mensaje
    '
    If mModel.SearchSorteos() Then
        '
        '   Inicializa la rejilla y visualiza los sorteos encontrados
        '
        mView.ClearGrid
        mView.AddSorteosToGrid mModel
    Else
        MsgBox "No se han encontrado Sorteos con los filtros indicados", vbOK + vbInformation, ThisWorkbook.Name
    End If
  
  On Error GoTo 0
BuscarLastPage__CleanExit:
    Exit Sub
BuscarLastPage_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    Call HandleException(ErrNumber, ErrDescription, "SorteoController.BuscarLastPage", ErrSource)
    Err.Raise ErrNumber, "SorteoController.BuscarLastPage", ErrDescription
End Sub
'------------------------------------------------------------------------------*
' Procedimiento  : EditarPorId
' Fecha          :
' Propósito      : Editar un registro seleccionado en la rejilla
' Parámetros     : IdSorteo Identificador único del sorteo
'                  datRow   Fila de la celda seleccionada
'------------------------------------------------------------------------------*
Public Sub EditarPorId(ByVal IdSorteo As Integer, datRow As Integer)
  On Error GoTo EditarPorId_Error
    '
    '   Visualizamos la fila
    '
    mView.SetFilaRegistro datRow
    '
    '   Invoca al modelo que carge el registro solicitado
    '
    mModel.GetSorteoRecord IdSorteo
    '
    '   Pasamos el modelo con el nuevo objeto a la vista
    '
    mView.DisplaySorteoRecord mModel
    '
    '   Seleccionamos la hoja de edición
    '
    ' ThisWorkbook.Sheets("Editar").Activate
    ' RAnge("C5").Select
  
  On Error GoTo 0
EditarPorId__CleanExit:
    Exit Sub
EditarPorId_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    Call HandleException(ErrNumber, ErrDescription, "SorteoController.EditarPorId", ErrSource)
    Err.Raise ErrNumber, "SorteoController.EditarPorId", ErrDescription
End Sub
'------------------------------------------------------------------------------*
' Procedimiento  : SetJuego
' Fecha          :
' Propósito      : Configura la vista para el juego
'------------------------------------------------------------------------------*
Public Sub SetJuego(ByVal mJuego As String)
    '
    '   Trasladamos a la vista la configuración del juego
    '
    mView.SetDisplayJuego mJuego
End Sub
'------------------------------------------------------------------------------*
' Procedimiento  : ResetFiltro
' Fecha          : lu., 05/ago/2019 23:48:47
' Propósito      : Resetea las condiciones de búsqueda
'------------------------------------------------------------------------------*
Public Sub ResetFiltro()
    '
    '   Inicializamos filtros
    '
    mView.ClearFiltros
    '
    '   Inicializamos rejilla
    '
    mView.ClearGrid
End Sub
' *===========(EOF):SorteoController.cls



