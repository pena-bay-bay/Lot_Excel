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
' *     Creaci�n   : mi., 17/abr/2019 22:59:32
' *     Versi�n    : 1.0
' *     Prop�sito  : Controlador de sorteo
' *
' *============================================================================*
Option Explicit
Option Base 0
'
'--- Variables Privadas -------------------------------------------------------*
Private mModel          As SorteoModel                  ' Modelo de negocio
Private mEditView       As SorteoEditarView             ' CU Editar Sorteo
Private mSelectView     As SorteoSeleccionarView        ' CU Seleccionar Sorteo
Private mPagActual      As Integer                      ' P�gina Actual de datos

'--- M�todos Privados ---------------------------------------------------------*
'
'
'
Private Sub Class_Initialize()
    Set mModel = New SorteoModel
    Set mEditView = New SorteoEditarView
    Set mSelectView = New SorteoSeleccionarView
End Sub
'
'
'
Private Sub Class_Terminate()
    Set mModel = Nothing
    Set mEditView = Nothing
    Set mSelectView = Nothing
End Sub
'-- Propiedades
'--- M�todos P�blicos ---------------------------------------------------------*
'------------------------------------------------------------------------------*
' Procedimiento  : Guardar
' Fecha          : s�., 11/may/2019 18:39:22
' Prop�sito      : Invoca al modelo para que guarde el Sorteo de la interfaz
'------------------------------------------------------------------------------*
'
Public Sub Guardar()
    '
    '   Invoca al modelo que guarde el Sorteo que devuelve la interfaz
    '
    If mModel.GuardarSorteoRecord(mEditView.GetDisplaySorteo) Then
        '
        '   Mensaje confirmaci�n
        MsgBox "El registro ha sido guardado correctamente.", vbOK + vbInformation, ThisWorkbook.Name
        Exit Sub

    End If
End Sub
'------------------------------------------------------------------------------*
' Procedimiento  : Nuevo
' Fecha          : s�., 11/may/2019 18:39:22
' Prop�sito      : Invoca al modelo para que cree un nuevo Sorteo
'------------------------------------------------------------------------------*
Public Sub Nuevo()
    '
    '   Inicializamos la rejilla de consulta
    '
    mSelectView.ClearGrid
    mSelectView.ClearFiltros
    '
    '   Inicializamos el area de trabajo
    '
    mEditView.ClearSorteoDisplay True
    '
    '   Invoca al modelo que cree un nuevo sorteo
    '
    mModel.NuevoSorteoRecord
    '
    '   Pasamos el modelo con el nuevo objeto a la vista
    '
    mEditView.DisplaySorteoRecord mModel
    '
    '   Seleccionamos la hoja de edici�n
    '
    ThisWorkbook.Sheets("Editar").Activate
    Range("C5").Select
End Sub
'------------------------------------------------------------------------------*
' Procedimiento  : Eliminar
' Fecha          : s�., 11/may/2019 18:39:22
' Prop�sito      : Elimina el sorteo actual
'------------------------------------------------------------------------------*
Public Sub Eliminar(ByVal IdSorteo As Integer)
    '
    '   Solicitar confirmaci�n
    '
    If MsgBox("�Est� seguro que quiere elminar este registro?", vbQuestion + vbYesNo, _
                ThisWorkbook.Name) = vbYes Then
                
        '
        '   Invoca al modelo que elimine el Sorteo
        '
        mModel.EliminarSorteoRecord IdSorteo
        '
        '
        '
        mEditView.ClearSorteoDisplay False
    End If
End Sub
'------------------------------------------------------------------------------*
' Procedimiento  : Siguiente
' Fecha          : s�., 11/may/2019 19:30:38
' Prop�sito      : Visualiza el sorteo siguiente al actual
'------------------------------------------------------------------------------*
Public Sub Siguiente(ByVal IdSorteo As Integer)
    '
    '   Invoca al modelo que carge el siguiente registro
    '
    If mModel.GetNextSorteoRecord(IdSorteo) Then
        '
        '   Pasamos el modelo con el nuevo objeto a la vista
        '
        mEditView.DisplaySorteoRecord mModel
    Else
        '
        '   Mensaje y salir
        '
        MsgBox "Es el �ltimo sorteo.", vbOK + vbInformation, ThisWorkbook.Name
        Exit Sub
    End If
End Sub
'------------------------------------------------------------------------------*
' Procedimiento  : Anterior
' Fecha          : s�., 11/may/2019 19:30:04
' Prop�sito      : Visualiza el sorteo anterior al Id actual
'------------------------------------------------------------------------------*
Public Sub Anterior(ByVal IdSorteo As Integer)
    '
    '   Invoca al modelo que carge el anterior registro
    '
    If mModel.GetPrevSorteoRecord(IdSorteo) Then
        '
        '   Pasamos el modelo con el nuevo objeto a la vista
        '
        mEditView.DisplaySorteoRecord mModel
    Else
        '
        '   Mensaje y salir
        '
        MsgBox "Es el primer sorteo.", vbOK + vbInformation, ThisWorkbook.Name
        Exit Sub
    End If
End Sub
'------------------------------------------------------------------------------*
' Procedimiento  : IrAlPrimero
' Fecha          : s�., 11/may/2019 18:39:22
' Prop�sito      : Visualiza el primer Sorteo
'------------------------------------------------------------------------------*
Public Sub IrAlPrimero()
    '
    '   Invoca al modelo que carge el primer registro
    '
    mModel.GetFirstSorteo
    '
    '   Pasamos el modelo con el nuevo objeto a la vista
    '
    mEditView.DisplaySorteoRecord mModel
    
End Sub
'------------------------------------------------------------------------------*
' Procedimiento  : IrAlUltimo
' Fecha          : s�., 11/may/2019 19:28:53
' Prop�sito      : Visualizar el �ltimo Sorteo
'------------------------------------------------------------------------------*
Public Sub IrAlUltimo()
    '
    '   Invoca al modelo que carge el ultimo registro
    '
    mModel.GetLastSorteo
    '
    '   Pasamos el modelo con el nuevo objeto a la vista
    '
    mEditView.DisplaySorteoRecord mModel

End Sub
'------------------------------------------------------------------------------*
' Procedimiento  : Buscar
' Fecha          : s�., 11/may/2019 18:39:22
' Prop�sito      : Buscar sorteos con las condiciones del filtro
'------------------------------------------------------------------------------*
Public Sub Buscar()
    '
    '   Invoca al modelo que busque Sorteos con el filtro definido en la interfaz
    '
    Set mModel = mSelectView.GetFiltroBusqueda
    '
    '   Indicamos la primera p�gina de la busqueda
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
        mSelectView.ClearGrid
        mSelectView.AddSorteosToGrid mModel
    Else
        MsgBox "No se han encontrado Sorteos con los filtros indicados", vbOK + vbInformation, ThisWorkbook.Name
    End If
End Sub
'------------------------------------------------------------------------------*
' Procedimiento  : BuscarFirstPage
' Fecha          : mi., 24/jul/2019 12:18:11
' Prop�sito      : Ir a la primera p�gina de la selecci�n realizada
'------------------------------------------------------------------------------*
Public Sub BuscarFirstPage()
    '
    '   Invoca al modelo que busque Sorteos con el filtro definido en la interfaz
    '
    Set mModel = mSelectView.GetFiltroBusqueda
    '
    '   Obtenemos la p�gina actual
    '
    mPagActual = mSelectView.PaginaActual
    '
    '   Comprobamos el Inicio
    '
    If mPagActual = 1 Then
        '
        '   Mensaje y salir
        '
        MsgBox "Es la primera p�gina para este filtro", vbOK + vbInformation, ThisWorkbook.Name
        Exit Sub
    End If
    '
    '   Vamos a la primera p�gina
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
        mSelectView.ClearGrid
        mSelectView.AddSorteosToGrid mModel
    Else
        MsgBox "No se han encontrado Sorteos con los filtros indicados", vbOK + vbInformation, ThisWorkbook.Name
    End If
End Sub
'------------------------------------------------------------------------------*
' Procedimiento  : BuscarNextPage
' Fecha          : mi., 24/jul/2019 12:16:53
' Prop�sito      : Presenta la siguiente p�gina a la visualizada
'------------------------------------------------------------------------------*
Public Sub BuscarNextPage()
    '
    '   Invoca al modelo que busque Sorteos con el filtro definido en la interfaz
    '
    Set mModel = mSelectView.GetFiltroBusqueda
    '
    '   Obtenemos la p�gina actual
    '
    mPagActual = mSelectView.PaginaActual
    '
    '   Comprobamos el final
    '
    If mPagActual >= mSelectView.TotalPaginas Then
        '
        '   Mensaje y salir
        '
        MsgBox "Es la �ltima p�gina para este filtro", vbOK + vbInformation, ThisWorkbook.Name
        Exit Sub
    End If
    '
    '   Incrementamos la p�gina
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
        mSelectView.ClearGrid
        mSelectView.AddSorteosToGrid mModel
    Else
        MsgBox "No se han encontrado Sorteos con los filtros indicados", vbOK + vbInformation, ThisWorkbook.Name
    End If

End Sub
'------------------------------------------------------------------------------*
' Procedimiento  : BuscarPrevPage
' Fecha          : mi., 24/jul/2019 12:28:34
' Prop�sito      : Ir a la p�gina anterior de la selecci�n
'------------------------------------------------------------------------------*
Public Sub BuscarPrevPage()
    '
    '   Invoca al modelo que busque Sorteos con el filtro definido en la interfaz
    '
    Set mModel = mSelectView.GetFiltroBusqueda
    '
    '   Obtenemos la p�gina actual
    '
    mPagActual = mSelectView.PaginaActual
    '
    '   Comprobamos el Inicio
    '
    If mPagActual = 1 Then
        '
        '   Mensaje y salir
        '
        MsgBox "Es la primera p�gina para este filtro", vbOK + vbInformation, ThisWorkbook.Name
        Exit Sub
    End If
    '
    '   Decrementamos  la p�gina
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
        mSelectView.ClearGrid
        mSelectView.AddSorteosToGrid mModel
    Else
        MsgBox "No se han encontrado Sorteos con los filtros indicados", vbOK + vbInformation, ThisWorkbook.Name
    End If

End Sub
'------------------------------------------------------------------------------*
' Procedimiento  : BuscarLastPage
' Fecha          : mi., 24/jul/2019 12:21:46
' Prop�sito      : Ir a la �ltima p�gina de la selecci�n de datos
'------------------------------------------------------------------------------*
Public Sub BuscarLastPage()
    '
    '   Invoca al modelo que busque Sorteos con el filtro definido en la interfaz
    '
    Set mModel = mSelectView.GetFiltroBusqueda
    '
    '   Obtenemos la p�gina actual
    '
    mPagActual = mSelectView.PaginaActual
    '
    '   Comprobamos el final
    '
    If mPagActual >= mSelectView.TotalPaginas Then
        '
        '   Mensaje y salir
        '
        MsgBox "Es la �ltima p�gina para este filtro", vbOK + vbInformation, ThisWorkbook.Name
        Exit Sub
    End If
    '
    '   Vamos a la �ltima p�gina
    '
    mModel.PaginaActual = mSelectView.TotalPaginas
    '
    '   Se invoca al metodo buscar del modelo
    '   y si encuentra elementos se visualizan
    '   Si no se da un mensaje
    '
    If mModel.SearchSorteos() Then
        '
        '   Inicializa la rejilla y visualiza los sorteos encontrados
        '
        mSelectView.ClearGrid
        mSelectView.AddSorteosToGrid mModel
    Else
        MsgBox "No se han encontrado Sorteos con los filtros indicados", vbOK + vbInformation, ThisWorkbook.Name
    End If


End Sub
'------------------------------------------------------------------------------*
' Procedimiento  : EditarPorId
' Fecha          : mi., 24/jul/2019 12:15:41
' Prop�sito      : Editar un registro seleccionado en la rejilla
' Par�metros     : IdSorteo Identificador �nico del sorteo
'                  datRow   Fila de la celda seleccionada
'------------------------------------------------------------------------------*
Public Sub EditarPorId(ByVal IdSorteo As Integer, datRow As Integer)
    '
    '   Visualizamos la fila
    '
    mSelectView.SetFilaRegistro datRow
    '
    '   Invoca al modelo que carge el registro solicitado
    '
    mModel.GetSorteoRecord IdSorteo
    '
    '   Pasamos el modelo con el nuevo objeto a la vista
    '
    mEditView.DisplaySorteoRecord mModel
    '
    '   Seleccionamos la hoja de edici�n
    '
    ThisWorkbook.Sheets("Editar").Activate
    Range("C5").Select
End Sub
'------------------------------------------------------------------------------*
' Procedimiento  : SetJuego
' Fecha          : s�., 11/may/2019 18:39:22
' Prop�sito      : Configura la vista para el juego
'------------------------------------------------------------------------------*
Public Sub SetJuego(ByVal mJuego As String)
    '
    '   Trasladamos a la vista la configuraci�n del juego
    '
    mEditView.SetDisplayJuego mJuego
End Sub
'------------------------------------------------------------------------------*
' Procedimiento  : ResetFiltro
' Fecha          : lu., 05/ago/2019 23:48:47
' Prop�sito      : Resetea las condiciones de b�squeda
'------------------------------------------------------------------------------*
Public Sub ResetFiltro()
    '
    '   Inicializamos filtros
    '
    mSelectView.ClearFiltros
    '
    '   Inicializamos rejilla
    '
    mSelectView.ClearGrid
End Sub
' *===========(EOF):SorteoController.cls

