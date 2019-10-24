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
'
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
' Fecha          : sá., 11/may/2019 18:39:22
' Propósito      : Invoca al modelo para que guarde el Sorteo de la interfaz
'------------------------------------------------------------------------------*
'
Public Sub Guardar()
'    '
'    '   Invoca al modelo que guarde el Sorteo que devuelve la interfaz
'    '
'    If mModel.GuardarSorteoRecord(mView.GetDisplaySorteo) Then
'        '
'        '   Mensaje confirmación
'        MsgBox "El registro ha sido guardado correctamente.", vbOK + vbInformation, ThisWorkbook.Name
'        Exit Sub
'
'    End If
End Sub
'------------------------------------------------------------------------------*
' Procedimiento  : Nuevo
' Fecha          : sá., 11/may/2019 18:39:22
' Propósito      : Invoca al modelo para que cree un nuevo Sorteo
'------------------------------------------------------------------------------*
Public Sub Nuevo()
'    '
'    '   Inicializamos la rejilla de consulta
'    '
'    mSelectView.ClearGrid
'    mSelectView.ClearFiltros
'    '
'    '   Inicializamos el area de trabajo
'    '
'    mEditView.ClearSorteoDisplay True
'    '
'    '   Invoca al modelo que cree un nuevo sorteo
'    '
'    mModel.NuevoSorteoRecord
'    '
'    '   Pasamos el modelo con el nuevo objeto a la vista
'    '
'    mEditView.DisplaySorteoRecord mModel
'    '
'    '   Seleccionamos la hoja de edición
'    '
'    ThisWorkbook.Sheets("Editar").Activate
'    RAnge("C5").Select
End Sub
'------------------------------------------------------------------------------*
' Procedimiento  : Eliminar
' Fecha          : sá., 11/may/2019 18:39:22
' Propósito      : Elimina el sorteo actual
'------------------------------------------------------------------------------*
Public Sub Eliminar(ByVal IdSorteo As Integer)
'    '
'    '   Solicitar confirmación
'    '
'    If MsgBox("¿Está seguro que quiere elminar este registro?", vbQuestion + vbYesNo, _
'                ThisWorkbook.Name) = vbYes Then
'
'        '
'        '   Invoca al modelo que elimine el Sorteo
'        '
'        mModel.EliminarSorteoRecord IdSorteo
'        '
'        '
'        '
'        mEditView.ClearSorteoDisplay False
'    End If
End Sub
'------------------------------------------------------------------------------*
' Procedimiento  : Siguiente
' Fecha          : sá., 11/may/2019 19:30:38
' Propósito      : Visualiza el sorteo siguiente al actual
'------------------------------------------------------------------------------*
Public Sub Siguiente(ByVal IdSorteo As Integer)
'    '
'    '   Invoca al modelo que carge el siguiente registro
'    '
'    If mModel.GetNextSorteoRecord(IdSorteo) Then
'        '
'        '   Pasamos el modelo con el nuevo objeto a la vista
'        '
'        mEditView.DisplaySorteoRecord mModel
'    Else
'        '
'        '   Mensaje y salir
'        '
'        MsgBox "Es el último sorteo.", vbOK + vbInformation, ThisWorkbook.Name
'        Exit Sub
'    End If
End Sub
'------------------------------------------------------------------------------*
' Procedimiento  : Anterior
' Fecha          : sá., 11/may/2019 19:30:04
' Propósito      : Visualiza el sorteo anterior al Id actual
'------------------------------------------------------------------------------*
Public Sub Anterior(ByVal IdSorteo As Integer)
'    '
'    '   Invoca al modelo que carge el anterior registro
'    '
'    If mModel.GetPrevSorteoRecord(IdSorteo) Then
'        '
'        '   Pasamos el modelo con el nuevo objeto a la vista
'        '
'        mEditView.DisplaySorteoRecord mModel
'    Else
'        '
'        '   Mensaje y salir
'        '
'        MsgBox "Es el primer sorteo.", vbOK + vbInformation, ThisWorkbook.Name
'        Exit Sub
'    End If
End Sub
'------------------------------------------------------------------------------*
' Procedimiento  : IrAlPrimero
' Fecha          : sá., 11/may/2019 18:39:22
' Propósito      : Visualiza el primer Sorteo
'------------------------------------------------------------------------------*
Public Sub IrAlPrimero()
'    '
'    '   Invoca al modelo que carge el primer registro
'    '
'    mModel.GetFirstSorteo
'    '
'    '   Pasamos el modelo con el nuevo objeto a la vista
'    '
'    mEditView.DisplaySorteoRecord mModel
'
End Sub
'------------------------------------------------------------------------------*
' Procedimiento  : IrAlUltimo
' Fecha          : sá., 11/may/2019 19:28:53
' Propósito      : Visualizar el último Sorteo
'------------------------------------------------------------------------------*
Public Sub IrAlUltimo()
'    '
'    '   Invoca al modelo que carge el ultimo registro
'    '
'    mModel.GetLastSorteo
'    '
'    '   Pasamos el modelo con el nuevo objeto a la vista
'    '
'    mEditView.DisplaySorteoRecord mModel
'
End Sub
'------------------------------------------------------------------------------*
' Procedimiento  : Buscar
' Fecha          : sá., 11/may/2019 18:39:22
' Propósito      : Buscar sorteos con las condiciones del filtro
'------------------------------------------------------------------------------*
Public Sub Buscar()
'    '
'    '   Invoca al modelo que busque Sorteos con el filtro definido en la interfaz
'    '
'    Set mModel = mSelectView.GetFiltroBusqueda
'    '
'    '   Indicamos la primera página de la busqueda
'    '
'    mModel.PaginaActual = 1
'    '   Se invoca al metodo buscar del modelo
'    '   y si encuentra elementos se visualizan
'    '   Si no se da un mensaje
'    '
'    If mModel.SearchSorteos() Then
'        '
'        '   Inicializa la rejilla y visualiza los sorteos encontrados
'        '
'        mSelectView.ClearGrid
'        mSelectView.AddSorteosToGrid mModel
'    Else
'        MsgBox "No se han encontrado Sorteos con los filtros indicados", vbOK + vbInformation, ThisWorkbook.Name
'    End If
End Sub
'------------------------------------------------------------------------------*
' Procedimiento  : BuscarFirstPage
' Fecha          : mi., 24/jul/2019 12:18:11
' Propósito      : Ir a la primera página de la selección realizada
'------------------------------------------------------------------------------*
Public Sub BuscarFirstPage()
'    '
'    '   Invoca al modelo que busque Sorteos con el filtro definido en la interfaz
'    '
'    Set mModel = mSelectView.GetFiltroBusqueda
'    '
'    '   Obtenemos la página actual
'    '
'    mPagActual = mSelectView.PaginaActual
'    '
'    '   Comprobamos el Inicio
'    '
'    If mPagActual = 1 Then
'        '
'        '   Mensaje y salir
'        '
'        MsgBox "Es la primera página para este filtro", vbOK + vbInformation, ThisWorkbook.Name
'        Exit Sub
'    End If
'    '
'    '   Vamos a la primera página
'    '
'    mModel.PaginaActual = 1
'    '
'    '   Se invoca al metodo buscar del modelo
'    '   y si encuentra elementos se visualizan
'    '   Si no se da un mensaje
'    '
'    If mModel.SearchSorteos() Then
'        '
'        '   Inicializa la rejilla y visualiza los sorteos encontrados
'        '
'        mSelectView.ClearGrid
'        mSelectView.AddSorteosToGrid mModel
'    Else
'        MsgBox "No se han encontrado Sorteos con los filtros indicados", vbOK + vbInformation, ThisWorkbook.Name
'    End If
End Sub
'------------------------------------------------------------------------------*
' Procedimiento  : BuscarNextPage
' Fecha          : mi., 24/jul/2019 12:16:53
' Propósito      : Presenta la siguiente página a la visualizada
'------------------------------------------------------------------------------*
Public Sub BuscarNextPage()
'    '
'    '   Invoca al modelo que busque Sorteos con el filtro definido en la interfaz
'    '
'    Set mModel = mSelectView.GetFiltroBusqueda
'    '
'    '   Obtenemos la página actual
'    '
'    mPagActual = mSelectView.PaginaActual
'    '
'    '   Comprobamos el final
'    '
'    If mPagActual >= mSelectView.TotalPaginas Then
'        '
'        '   Mensaje y salir
'        '
'        MsgBox "Es la última página para este filtro", vbOK + vbInformation, ThisWorkbook.Name
'        Exit Sub
'    End If
'    '
'    '   Incrementamos la página
'    '
'    mModel.PaginaActual = mPagActual + 1
'    '
'    '   Se invoca al metodo buscar del modelo
'    '   y si encuentra elementos se visualizan
'    '   Si no se da un mensaje
'    '
'    If mModel.SearchSorteos() Then
'        '
'        '   Inicializa la rejilla y visualiza los sorteos encontrados
'        '
'        mSelectView.ClearGrid
'        mSelectView.AddSorteosToGrid mModel
'    Else
'        MsgBox "No se han encontrado Sorteos con los filtros indicados", vbOK + vbInformation, ThisWorkbook.Name
'    End If

End Sub
'------------------------------------------------------------------------------*
' Procedimiento  : BuscarPrevPage
' Fecha          : mi., 24/jul/2019 12:28:34
' Propósito      : Ir a la página anterior de la selección
'------------------------------------------------------------------------------*
Public Sub BuscarPrevPage()
'    '
'    '   Invoca al modelo que busque Sorteos con el filtro definido en la interfaz
'    '
'    Set mModel = mSelectView.GetFiltroBusqueda
'    '
'    '   Obtenemos la página actual
'    '
'    mPagActual = mSelectView.PaginaActual
'    '
'    '   Comprobamos el Inicio
'    '
'    If mPagActual = 1 Then
'        '
'        '   Mensaje y salir
'        '
'        MsgBox "Es la primera página para este filtro", vbOK + vbInformation, ThisWorkbook.Name
'        Exit Sub
'    End If
'    '
'    '   Decrementamos  la página
'    '
'    mModel.PaginaActual = mPagActual - 1
'    '
'    '   Se invoca al metodo buscar del modelo
'    '   y si encuentra elementos se visualizan
'    '   Si no se da un mensaje
'    '
'    If mModel.SearchSorteos() Then
'        '
'        '   Inicializa la rejilla y visualiza los sorteos encontrados
'        '
'        mSelectView.ClearGrid
'        mSelectView.AddSorteosToGrid mModel
'    Else
'        MsgBox "No se han encontrado Sorteos con los filtros indicados", vbOK + vbInformation, ThisWorkbook.Name
'    End If

End Sub
'------------------------------------------------------------------------------*
' Procedimiento  : BuscarLastPage
' Fecha          : mi., 24/jul/2019 12:21:46
' Propósito      : Ir a la última página de la selección de datos
'------------------------------------------------------------------------------*
Public Sub BuscarLastPage()
'    '
'    '   Invoca al modelo que busque Sorteos con el filtro definido en la interfaz
'    '
'    Set mModel = mSelectView.GetFiltroBusqueda
'    '
'    '   Obtenemos la página actual
'    '
'    mPagActual = mSelectView.PaginaActual
'    '
'    '   Comprobamos el final
'    '
'    If mPagActual >= mSelectView.TotalPaginas Then
'        '
'        '   Mensaje y salir
'        '
'        MsgBox "Es la última página para este filtro", vbOK + vbInformation, ThisWorkbook.Name
'        Exit Sub
'    End If
'    '
'    '   Vamos a la última página
'    '
'    mModel.PaginaActual = mSelectView.TotalPaginas
'    '
'    '   Se invoca al metodo buscar del modelo
'    '   y si encuentra elementos se visualizan
'    '   Si no se da un mensaje
'    '
'    If mModel.SearchSorteos() Then
'        '
'        '   Inicializa la rejilla y visualiza los sorteos encontrados
'        '
'        mSelectView.ClearGrid
'        mSelectView.AddSorteosToGrid mModel
'    Else
'        MsgBox "No se han encontrado Sorteos con los filtros indicados", vbOK + vbInformation, ThisWorkbook.Name
'    End If
'

End Sub
'------------------------------------------------------------------------------*
' Procedimiento  : EditarPorId
' Fecha          : mi., 24/jul/2019 12:15:41
' Propósito      : Editar un registro seleccionado en la rejilla
' Parámetros     : IdSorteo Identificador único del sorteo
'                  datRow   Fila de la celda seleccionada
'------------------------------------------------------------------------------*
Public Sub EditarPorId(ByVal IdSorteo As Integer, datRow As Integer)
'    '
'    '   Visualizamos la fila
'    '
'    mSelectView.SetFilaRegistro datRow
'    '
'    '   Invoca al modelo que carge el registro solicitado
'    '
'    mModel.GetSorteoRecord IdSorteo
'    '
'    '   Pasamos el modelo con el nuevo objeto a la vista
'    '
'    mEditView.DisplaySorteoRecord mModel
'    '
'    '   Seleccionamos la hoja de edición
'    '
'    ThisWorkbook.Sheets("Editar").Activate
'    RAnge("C5").Select
End Sub
'------------------------------------------------------------------------------*
' Procedimiento  : SetJuego
' Fecha          : sá., 11/may/2019 18:39:22
' Propósito      : Configura la vista para el juego
'------------------------------------------------------------------------------*
Public Sub SetJuego(ByVal mJuego As String)
'    '
'    '   Trasladamos a la vista la configuración del juego
'    '
'    mEditView.SetDisplayJuego mJuego
End Sub
'------------------------------------------------------------------------------*
' Procedimiento  : ResetFiltro
' Fecha          : lu., 05/ago/2019 23:48:47
' Propósito      : Resetea las condiciones de búsqueda
'------------------------------------------------------------------------------*
Public Sub ResetFiltro()
'    '
'    '   Inicializamos filtros
'    '
'    mSelectView.ClearFiltros
'    '
'    '   Inicializamos rejilla
'    '
'    mSelectView.ClearGrid
End Sub
' *===========(EOF):SorteoController.cls



