Attribute VB_Name = "Lot_PqtSugerenciasTesting"
' *============================================================================*
' *
' *     Fichero    : Lot_PqtSugerenciasTesting.bas
' *
' *     Autor      : Carlos Almela Baeza
' *     Creación   : ma., 28/ene/2020 20:32:45
' *     Versión    : 1.0
' *     Propósito  : Módulo de pruebas de las clases del paquete Sugerencias
' *                - Metodo
' *                - Metodos
' *                - MetodoEngine
' *                - MetodoModel
' *                - MetodoController
' *                - MetodoSelectView
' *                - MetodoEditView
' *============================================================================*
Option Explicit
Option Base 0
'
'--- Variables Privadas -------------------------------------------------------*
Dim mObj As Metodo
Dim mCol As Metodos
Dim mMdl As MetodoModel
Dim mCtrl As MetodoController
'Dim mSelView as MetodoSelectView
'Dim mEdiView as MetodoEditView
'--- Constantes ---------------------------------------------------------------*
'--- Mensajes -----------------------------------------------------------------*
'--- Errores ------------------------------------------------------------------*
'--- Propiedades --------------------------------------------------------------*
'--- Métodos Privados ---------------------------------------------------------*
'------------------------------------------------------------------------------*
' Procedimiento  : MetodoTest
' Fecha          : sá., 07/mar/2020 23:53:15
' Propósito      : Pruebas unitarias clase Metodo
'------------------------------------------------------------------------------*
Private Sub MetodoTest()
  On Error GoTo MetodoTest_Error
    '
    '   Case error 1: Objeto vacio
    '
    Set mObj = New Metodo
    PrintMetodo mObj
    '
    '   Case error 2: Metodo Sin definir
    '
    Set mObj = New Metodo
    With mObj
        .TipoProcedimiento = mtdSinDefinir
        .CriteriosAgrupacion = grpDecenas
        .CriteriosOrdenacion = ordDesviacion
        .DiasAnalisis = 90
        .ModalidadJuego = LP_LB_6_49
    End With
    PrintMetodo mObj
    '
    '   Case error 3: Metodo aleatorio
    '
    Set mObj = New Metodo
    With mObj
        .TipoProcedimiento = mtdAleatorio
        .ModalidadJuego = LP_LB_6_49
    End With
    PrintMetodo mObj
    '
    '   Case error 4: Metodo Bombo
    '
        Set mObj = New Metodo
    With mObj
        .TipoProcedimiento = mtdBombo
        .ModalidadJuego = LP_LB_6_49
    End With
    PrintMetodo mObj
    '
    '   Case error 5: mtdEstadistico
    '
    Set mObj = New Metodo
    With mObj
        .TipoProcedimiento = mtdEstadistico
        .CriteriosOrdenacion = ordAusencia
        .Orden = True
        .NumeroSorteos = 90
        .ModalidadJuego = LP_LB_6_49
        .SentidoOrdenacion = True
    End With
    PrintMetodo mObj
    '
    '   Case error 6: mtdEstaDosVariables
    '
    Set mObj = New Metodo
    With mObj
        .TipoProcedimiento = mtdSinDefinir
        .CriteriosAgrupacion = grpDecenas
        .CriteriosOrdenacion = ordDesviacion
        .DiasAnalisis = 90
        .ModalidadJuego = LP_LB_6_49
    End With
    PrintMetodo mObj
    '
    '   Case error 7: mtdAlgoritmoAG
    '   Case error 8: mtdRedNeuronal
    '   Case error 9: mtdEstadCombinacion
    '
    
  On Error GoTo 0
MetodoTest__CleanExit:
    Exit Sub
MetodoTest_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    '   Audita el error
    Call HandleException(ErrNumber, ErrDescription, "Lot_PqtSugerenciasTesting.MetodoTest", ErrSource)
    '   Informa del error
    Call MsgBox(ErrDescription, vbCritical Or vbSystemModal, ThisWorkbook.Name)
    Call Trace("CERRAR")
End Sub
'------------------------------------------------------------------------------*
' Procedimiento  : MetodoTest
' Fecha          : sá., 07/mar/2020 23:53:15
' Propósito      : Pruebas unitarias clase Metodos
'------------------------------------------------------------------------------*
Private Sub MetodosTest()
  On Error GoTo MetodosTest_Error
  On Error GoTo 0
MetodosTest__CleanExit:
    Exit Sub
MetodosTest_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    '   Audita el error
    Call HandleException(ErrNumber, ErrDescription, "Lot_PqtSugerenciasTesting.MetodosTest", ErrSource)
    '   Informa del error
    Call MsgBox(ErrDescription, vbCritical Or vbSystemModal, ThisWorkbook.Name)
    Call Trace("CERRAR")
End Sub
'------------------------------------------------------------------------------*
' Procedimiento  : MetodoModelTest
' Fecha          : sá., 07/mar/2020 23:53:15
' Propósito      : Pruebas unitarias clase MetodoModel
'------------------------------------------------------------------------------*
Private Sub MetodoModelTest()
    'Dim mMdl As MetodoModel
  On Error GoTo MetodoModelTest_Error
    '
    '   Creación del Objeto
    '
    Set mMdl = New MetodoModel
    PrintMetodoModel mMdl
    '
    '   Creación nuevo Metodo
    '
    With mMdl.Metodo
        .TipoProcedimiento = mtdAleatorio
        .CriteriosOrdenacion = ordProbabilidad
        .DiasAnalisis = 90
        .ModalidadJuego = LP_LB_6_49
    End With
    PrintMetodoModel mMdl
    If mMdl.Add Then
        Debug.Print "Método añandido Id (" & CStr(mMdl.Metodo.Id) & ")"
    End If
    '
    '   Busqueda con exito
    '
    If mMdl.GetRecord(1) Then
        PrintMetodo mMdl.Metodo
    Else
        Debug.Print "#Error MetodoModel.GetRecord()"
    End If
    '
    '   Busqueda sin exito
    '
    On Error Resume Next
    If mMdl.GetRecord(15) Then
        If Err.Number <> 0 Then
            Debug.Print "#Error: " & Err.Number & " - " & Err.Description
        Else
            PrintMetodoModel mMdl
        End If
    End If
'    mMdl.Save
'    mMdl.Del
'    mMdl.GetFirst
'    mMdl.GetNext
'    mMdl.GetLast
'    mMdl.GetNext
'    mMdl.GetPrev
'    mMdl.Search
  On Error GoTo 0
MetodoModelTest__CleanExit:
    Exit Sub
MetodoModelTest_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    Call HandleException(ErrNumber, ErrDescription, "Lot_PqtSugerenciasTesting.MetodoModelTest", ErrSource)
    Call MsgBox(ErrDescription, vbCritical Or vbSystemModal, ThisWorkbook.Name)
    Call Trace("CERRAR")
End Sub
'------------------------------------------------------------------------------*
' Procedimiento  : MetodoControllerTest
' Fecha          : sá., 07/mar/2020 23:53:15
' Propósito      : Pruebas unitarias clase MetodoController
'------------------------------------------------------------------------------*
Private Sub MetodoControllerTest()
  On Error GoTo MetodoControllerTest_Error
  On Error GoTo 0
MetodoControllerTest__CleanExit:
    Exit Sub
MetodoControllerTest_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    '   Audita el error
    Call HandleException(ErrNumber, ErrDescription, "Lot_PqtSugerenciasTesting.MetodosTest", ErrSource)
    '   Informa del error
    Call MsgBox(ErrDescription, vbCritical Or vbSystemModal, ThisWorkbook.Name)
    Call Trace("CERRAR")
End Sub
'------------------------------------------------------------------------------*
' Procedimiento  : MetodoSelectViewTest
' Fecha          : sá., 07/mar/2020 23:53:15
' Propósito      : Pruebas unitarias clase MetodoSelectView
'------------------------------------------------------------------------------*
Private Sub MetodoSelectViewTest()
  On Error GoTo MetodoSelectViewTest_Error
  On Error GoTo 0
MetodoSelectViewTest__CleanExit:
  Exit Sub
MetodoSelectViewTest_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    '   Audita el error
    Call HandleException(ErrNumber, ErrDescription, "Lot_PqtSugerenciasTesting.MetodosTest", ErrSource)
    '   Informa del error
    Call MsgBox(ErrDescription, vbCritical Or vbSystemModal, ThisWorkbook.Name)
    Call Trace("CERRAR")
End Sub
'------------------------------------------------------------------------------*
' Procedimiento  : MetodoEditViewTest
' Fecha          : sá., 07/mar/2020 23:53:15
' Propósito      : Pruebas unitarias clase MetodoEditView
'------------------------------------------------------------------------------*
Private Sub MetodoEditViewTest()
  On Error GoTo MetodoEditViewTest_Error
  
  On Error GoTo 0
MetodoEditViewTest__CleanExit:
    Exit Sub
MetodoEditViewTest_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    '   Audita el error
    Call HandleException(ErrNumber, ErrDescription, "Lot_PqtSugerenciasTesting.MetodosTest", ErrSource)
    '   Informa del error
    Call MsgBox(ErrDescription, vbCritical Or vbSystemModal, ThisWorkbook.Name)
    Call Trace("CERRAR")
End Sub
'------------------------------------------------------------------------------*
' Procedimiento  : PrintMetodo
' Fecha          : sá., 07/mar/2020 23:53:15
' Propósito      : Visualizar el objeto
'------------------------------------------------------------------------------*
Private Sub PrintMetodo(mObj As Metodo)
     Debug.Print "==> Pruebas Metodo"
     '-> Propiedades
     Debug.Print vbTab & "CriteriosAgrupacion        = " & mObj.CriteriosAgrupacion
     Debug.Print vbTab & "CriteriosOrdenacion        = " & mObj.CriteriosOrdenacion
     Debug.Print vbTab & "DiasAnalisis               = " & mObj.DiasAnalisis
     Debug.Print vbTab & "FechaAlta                  = " & mObj.EntidadNegocio.FechaAlta
     Debug.Print vbTab & "Id                         = " & mObj.Id
     Debug.Print vbTab & "ModalidadJuego             = " & mObj.ModalidadJuego
     Debug.Print vbTab & "NumeroSorteos              = " & mObj.NumeroSorteos
     Debug.Print vbTab & "Orden                      = " & mObj.Orden
     Debug.Print vbTab & "Pronosticos                = " & mObj.Pronosticos
     Debug.Print vbTab & "SentidoOrdenacion          = " & mObj.SentidoOrdenacion
     Debug.Print vbTab & "TipoMuestra                = " & mObj.TipoMuestra
     Debug.Print vbTab & "TipoProcedimiento          = " & mObj.TipoProcedimiento
     '-> Metodos
     Debug.Print vbTab & "AgrupacionToString()       = " & mObj.AgrupacionToString()
     Debug.Print vbTab & "OrdenacionToString()       = " & mObj.OrdenacionToString()
     Debug.Print vbTab & "TipoProcedimientoTostring()= " & mObj.TipoProcedimientoTostring()
     Debug.Print vbTab & "ToString()                 = " & mObj.ToString()
     Debug.Print vbTab & "EsValido()                 = " & mObj.EsValido()
End Sub
'------------------------------------------------------------------------------*
' Procedimiento  : PrintMetodoModel
' Fecha          : mi., 11/mar/2020 19:45:28
' Propósito      : Visualizar el objeto MetodoModel
'------------------------------------------------------------------------------*
Private Sub PrintMetodoModel(mObj As MetodoModel)
    Debug.Print "==> Pruebas MetodoModel"
    '-> Propiedades
    Debug.Print vbTab & "LinePerPage    = " & mObj.LinePerPage
    Debug.Print vbTab & "Metodo         = " & mObj.Metodo.ToString()
    Debug.Print vbTab & "Metodos.Count  = " & "#TODO" 'mObj.Metodos.Count
    Debug.Print vbTab & "TotalPages     = " & mObj.TotalPages
    Debug.Print vbTab & "TotalRecords   = " & mObj.TotalRecords
End Sub
' *===========(EOF): Lot_PqtSugerenciasTesting.bas
