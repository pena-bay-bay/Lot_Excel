Attribute VB_Name = "CU_DefinirSorteo"
' *============================================================================*
' *
' *     Fichero    : CU_DefinirSorteo.cls
' *
' *     Autor      : Carlos Almela Baeza
' *     Creación   : mi., 17/abr/2019 22:44:55
' *     Versión    : 1.0
' *     Propósito  : Caso de Uso Definir sorteo.
' *
' *============================================================================*
Option Explicit
Option Base 0
'
'   Constantes
'
Public Const ERR_TODO = 999
Public Const MSG_TODO = "Rutina pendiente de codificar."
Public Const LT_EUROMILLON As String = "Euromillon"
Public Const LT_ESTRELLAS As String = "Estrellas"
Public Const LT_GORDO As String = "Gordo Primitiva"
Public Const LT_BONOLOTO As String = "Bonoloto"
Public Const LT_PRIMITIVA As String = "Primitiva"
Public Const LT_CLAVE As String = "Clave"
Public Const LT_COMPLEMENTARIO As String = "Complementario"
Public Const LT_REINTEGRO As String = "Reintegro"
Public Const LP_PREMIOS_EURO As String = "25.326.022,00;253.763,79;35.462,69;" & _
                                         "3.097,48;141,67;102,16;44,39;20,82;13" & _
                                         ",76;10,27;11,31;8,00;4,00"
Public Const LP_PREMIOS_GORDO As String = "5.438.778,89;165.842,81;7.911,24;" & _
                                         "105,37;28,33;9,17;4,89;3,00;1,50"
Public Const LP_PREMIOS_BONO As String = "981.440,37;47.921,93;895,74;37,09;4,00;0,50"
Public Const LP_PREMIOS_PRIMI As String = "13.085.952,17;1.598.135,28;59.930" & _
                                          ",07;1.438,76;51,58;8,00;1,00"

'--- Variables Privadas -------------------------------------------------------*
Private Const hjEditarSorteo As String = "Editar"
Private Const hjSeleccionarSorteo As String = "Consultar"
Private Const rgAreaEdicion As String = "A3:I23"

'--- Variables Publicas -------------------------------------------------------*

Public oSorteoEditarView        As SorteoEditarView
Public oSorteoSeleccionarView   As SorteoSeleccionarView



' *============================================================================*
'--- Area de testing ----------------------------------------------------------*
'------------------------------------------------------------------------------*
' Procedimiento  : SorteoControllerTest
' Fecha          : ma., 28/may/2019 20:49:40
' Propósito      : Pruebas unitarias de la clase SorteoController
'------------------------------------------------------------------------------*
'
Private Sub SorteoControllerTest()
    Dim mCtrl As SorteoController
    
    
  On Error GoTo SorteoControllerTest_Error
    Debug.Print "------------------------------"
    Debug.Print "Testing clase SorteoController"
    Debug.Print "=> Init"
    Set mCtrl = New SorteoController
    '
    '   1.- metodo Anterior
    '
    mCtrl.Anterior 2333
    Debug.Print "   mCtrl.Anterior (2333)= " & 2333
    '
    '   2.- Metodo Buscar
    '
    mCtrl.Buscar
    Debug.Print "   mCtrl.Buscar           "
    '
    '   3.- Metodo BuscarFirstPage
    '
    mCtrl.BuscarFirstPage
    '
    '   3.- Metodo BuscarFirstPage
    '
    mCtrl.BuscarLastPage
    '
    '   3.- Metodo BuscarFirstPage
    '
    mCtrl.BuscarNextPage
    '
    '   3.- Metodo BuscarFirstPage
    '
    mCtrl.BuscarPrevPage
    '
    '   3.- Metodo BuscarFirstPage
    '
    mCtrl.EditarPorId 2333, 12
    '
    '   3.- Metodo BuscarFirstPage
    '
    mCtrl.Eliminar 2333
    '
    '   3.- Metodo BuscarFirstPage
    '
    mCtrl.Guardar
    '
    '   3.- Metodo BuscarFirstPage
    '
    mCtrl.IrAlPrimero
    '
    '   3.- Metodo BuscarFirstPage
    '
    mCtrl.IrAlUltimo
    '
    '   3.- Metodo BuscarFirstPage
    '
    mCtrl.Nuevo
    '
    '   3.- Metodo BuscarFirstPage
    '
    mCtrl.SetJuego LT_EUROMILLON
    '
    '   3.- Metodo BuscarFirstPage
    '
    mCtrl.Siguiente 2333
    '
  On Error GoTo 0
SorteoControllerTest__CleanExit:
    Exit Sub
            
SorteoControllerTest_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    '   Audita el error
    Call HandleException(ErrNumber, ErrDescription, "CU_DefinirSorteo.SorteoControllerTest", ErrSource)
    '   Informa del error
    Call MsgBox(ErrDescription, vbCritical Or vbSystemModal, ThisWorkbook.Name)
    Call Trace("CERRAR")
End Sub

'------------------------------------------------------------------------------*
' Procedimiento  : SorteoModelTest
' Fecha          : ma., 28/may/2019 20:50:28
' Propósito      : Pruebas unitarias de la clase SorteoModel
'------------------------------------------------------------------------------*
'
Private Sub SorteoModelTest()
    Dim oModel   As SorteoModel
    Dim mId      As Integer
    Dim mIdFirst As Integer
    Dim mIdLast  As Integer
    Dim mResult  As Variant

  On Error GoTo SorteoModelTest_Error
    Debug.Print "------------------------------"
    Debug.Print "Testing clase SorteoModel"
    Debug.Print "=> Init"
    Set oModel = New SorteoModel
    '
    '   1.- Test metodo NuevoSorteoRecord
    '
    oModel.NuevoSorteoRecord
    mId = oModel.IdSelected
    Debug.Print "   NuevoSorteoRecord (Id)= " & mId
    With oModel
        .Juego = LT_BONOLOTO
        .CombinacionGanadora = "45-12-9-27-15-3"
        .Complementario = 14
        .FechaSorteo = Format(Now, "dd/MM/yyyy")
        .DiaSemana = UCase(Left(Format(Now, "ddd"), 1))
        .NumSorteo = "2019/008"
        .OrdenAparicion = "Si"
        .Reintegro = 8
        .Semana = 21
    End With
    '
    '   2.- Test metodo GuardarSorteoRecord
    '
    mResult = oModel.GuardarSorteoRecord(oModel)
    Debug.Print "   GuardarSorteoRecord => " & mResult
    '
    '   3.- Test metodo GetFirstSorteo
    '
    oModel.GetFirstSorteo
    Debug.Print "   GetFirstSorteo (mId) => " & oModel.IdSelected
    mIdFirst = oModel.IdSelected
    '
    '   4.- Test metodo GetLastSorteo
    '
    oModel.GetLastSorteo
    Debug.Print "   GetLastSorteo (mId) => " & oModel.IdSelected
    mIdLast = oModel.IdSelected
    '
    '   5.- Test metodo GetSorteoRecord
    '
    mId = mIdFirst + Int((mIdLast - mIdFirst) / 2)
    mResult = oModel.GetSorteoRecord(mId)
    Debug.Print "   GetSorteoRecord (mId) => " & oModel.IdSelected & " Combinacion: " & oModel.CombinacionGanadora
    '
    '   6.- Test metodo GetNextSorteoRecord
    '
    mResult = oModel.GetNextSorteoRecord(mId)
    Debug.Print "   GetNextSorteoRecord (mId) => " & oModel.IdSelected & " Combinacion: " & oModel.CombinacionGanadora
    '
    '   7.- Test metodo GetPrevSorteoRecord
    '
    mResult = oModel.GetPrevSorteoRecord(mId)
    Debug.Print "   GetPrevSorteoRecord (mId) => " & oModel.IdSelected & " Combinacion: " & oModel.CombinacionGanadora
    '
    '   8.- Test metodo EliminarSorteoRecord
    '
    mResult = oModel.EliminarSorteoRecord(mIdLast)
    Debug.Print "   EliminarSorteoRecord => " & mResult
    '
    '   9.- Text metodo SearchSorteos por fecha
    '
    oModel.FechaSorteo = #3/20/2019#
    oModel.Juego = Empty
    oModel.LineasPorPagina = 7
    oModel.PaginaActual = 1
    mResult = oModel.SearchSorteos
    Debug.Print "   SearchSorteos (por fecha)=> " & mResult & " Registros: " & oModel.TotalRegistros
    '
    '   10.- Text metodo SearchSorteos por periodo
    '
    oModel.FechaInicio = #3/22/2019#
    oModel.FechaFin = #3/22/2019#
    oModel.Juego = Empty
    oModel.FechaSorteo = 0
    oModel.LineasPorPagina = 7
    oModel.PaginaActual = 1
    mResult = oModel.SearchSorteos
    Debug.Print "   SearchSorteos (por periodo)=> " & mResult & " Registros: " & oModel.TotalRegistros
    '
    '   11.- Text metodo SearchSorteos por juego
    '
    oModel.Juego = LT_BONOLOTO
    oModel.FechaInicio = Empty
    oModel.FechaFin = Empty
    oModel.FechaSorteo = 0
    oModel.LineasPorPagina = 7
    oModel.PaginaActual = 1
    mResult = oModel.SearchSorteos
    Debug.Print "   SearchSorteos (por juego)=> " & mResult & " Registros: " & oModel.TotalRegistros
    '
    '   12.- Text metodo SearchSorteos sin filtros
    '
    oModel.FechaSorteo = 0
    oModel.Juego = Empty
    oModel.FechaInicio = Empty
    oModel.FechaFin = Empty
    oModel.LineasPorPagina = 7
    oModel.PaginaActual = 2
    mResult = oModel.SearchSorteos
    Debug.Print "   SearchSorteos (sin filtro)=> " & mResult & " Registros: " & oModel.TotalRegistros
    '
    '
    '
    Debug.Print "=> Finish"
    Debug.Print "------------------------------"
    
  On Error GoTo 0
SorteoModelTest__CleanExit:
    Exit Sub
            
SorteoModelTest_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    '   Audita el error
    Call HandleException(ErrNumber, ErrDescription, "CU_DefinirSorteo.SorteoModelTest", ErrSource)
    '   Informa del error
    Call MsgBox(ErrDescription, vbCritical Or vbSystemModal, ThisWorkbook.Name)
    Call Trace("CERRAR")
    
End Sub
'------------------------------------------------------------------------------*
' Procedimiento  : SorteoEditarViewTest
' Fecha          : ma., 28/may/2019 20:50:28
' Propósito      : Pruebas unitarias de la clase SorteoEditarView
'------------------------------------------------------------------------------*
'
Private Sub SorteoEditarViewTest()
    Dim mView As SorteoEditarView
    Dim mModel As SorteoModel
  
  On Error GoTo SorteoEditarViewTest_Error
    Debug.Print "------------------------------"
    Debug.Print "Testing clase SorteoEditarView"
    Debug.Print "=> Init"
    
    Set mModel = New SorteoModel
    Set mView = New SorteoEditarView
    '
    '   1.- Prueba unitaria método ClearSorteoDisplay
    '
    mView.ClearSorteoDisplay False
    Debug.Print "  * ClearSorteoDisplay Ok"
    '
    '   2.- Prueba unitaria DisplaySorteoRecord
    '
    With mModel
        .IdSelected = 2329
        .Juego = LT_PRIMITIVA
        .NumSorteo = "2019/065"
        .FechaSorteo = "16/03/2019"
        .DiaSemana = "L"
        .Semana = 11
        .OrdenAparicion = "Si"
        .CombinacionGanadora = "9 - 14 - 5 - 49 - 32 - 4"
        .Complementario = 30
        .Reintegro = 7
    End With
    mView.DisplaySorteoRecord mModel
    Debug.Print "  * DisplaySorteoRecord Ok"
    '
    '   3.- Prueba unitaria GetDisplaySorteo
    '
    Set mModel = mView.GetDisplaySorteo
    Debug.Print "  * GetDisplaySorteo Ok" & mModel.CombinacionGanadora
    '
    '   4.- Prueba unitaria SetDisplayJuego
    '
    mView.ClearSorteoDisplay True
    mView.SetDisplayJuego LT_BONOLOTO
    Debug.Print "  * SetDisplayJuego Ok " & LT_BONOLOTO
    mView.SetDisplayJuego LT_PRIMITIVA
    Debug.Print "  * SetDisplayJuego Ok " & LT_PRIMITIVA
    mView.SetDisplayJuego LT_EUROMILLON
    Debug.Print "  * SetDisplayJuego Ok " & LT_EUROMILLON
    mView.SetDisplayJuego LT_GORDO
    Debug.Print "  * SetDisplayJuego Ok " & LT_GORDO
    '
    '
    '
    Debug.Print "=> Finish"
    Debug.Print "------------------------------"
  On Error GoTo 0
SorteoEditarViewTest__CleanExit:
    Exit Sub
            
SorteoEditarViewTest_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    '   Audita el error
    Call HandleException(ErrNumber, ErrDescription, "CU_DefinirSorteo.SorteoEditarViewTest", ErrSource)
    '   Informa del error
    Call MsgBox(ErrDescription, vbCritical Or vbSystemModal, ThisWorkbook.Name)
    Call Trace("CERRAR")
End Sub

'------------------------------------------------------------------------------*
' Procedimiento  : SorteoSeleccionarViewTest
' Fecha          : ma., 28/may/2019 20:50:28
' Propósito      : Pruebas unitarias de la clase SorteoSeleccionarView
'------------------------------------------------------------------------------*
'
Private Sub SorteoSeleccionarViewTest()
    Dim mView As SorteoSeleccionarView
    Dim mModel As SorteoModel
    
  On Error GoTo SorteoSeleccionarViewTest_Error
    Debug.Print "------------------------------"
    Debug.Print "Testing clase SorteoSeleccionarView"
    Debug.Print "=> Init"
    '
    '   Creamos objetos a probar
    '
    Set mModel = New SorteoModel
    Set mView = New SorteoSeleccionarView
    '
    '   1.- Prueba unitaria método ClearFiltros
    '
    mView.ClearFiltros
    Debug.Print "  * ClearFiltros Ok"
    '
    '   2.- Prueba unitaria método ClearGrid
    '
    mView.ClearGrid
    Debug.Print "  * ClearGrid Ok"
    '
    '   3.- Prueba unitaria GetFiltroBusqueda
    '
    ThisWorkbook.Sheets("Consultar").Range("C5") = LT_BONOLOTO
    Set mModel = mView.GetFiltroBusqueda
    Debug.Print "  * GetFiltroBusqueda Ok =>" & mModel.Juego
    '
    '   4.- Prueba unitaria Pagina Actual
    '
    ThisWorkbook.Sheets("Consultar").Range("B19") = "Página:5/9"
    Debug.Print "  * Pagina Actual Ok (5)=>" & mView.PaginaActual
    '
    '   5.- Prueba unitaria Total Paginas
    '
    Debug.Print "  * Total Paginas Ok (9)=>" & mView.TotalPaginas
    '
    '   6.- Prueba unitaria AddSorteosToGrid
    '
    ThisWorkbook.Sheets("Consultar").Range("C5") = LT_BONOLOTO
    Set mModel = mView.GetFiltroBusqueda
    mModel.PaginaActual = 1
    mModel.SearchSorteos
    mView.AddSorteosToGrid mModel
    Debug.Print "  * Pagina Actual Ok (6)=>" & UBound(mModel.ResultadosSearch, 1)
  
  On Error GoTo 0
SorteoSeleccionarViewTest__CleanExit:
    Exit Sub
            
SorteoSeleccionarViewTest_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    '   Audita el error
    Call HandleException(ErrNumber, ErrDescription, "CU_DefinirSorteo.SorteoSeleccionarViewTest", ErrSource)
    '   Informa del error
    Call MsgBox(ErrDescription, vbCritical Or vbSystemModal, ThisWorkbook.Name)
    Call Trace("CERRAR")
End Sub

'' *===========(EOF): CU_DefinirSorteo
