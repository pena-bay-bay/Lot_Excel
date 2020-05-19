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
' *                - MetodoModel
' *                - MetodoController
' *                - FiltroCombinacion
' *                - FiltrosCombinacion
' *                - Sugerencia
' *                - Sugerencias
' *                - RealizarSugerencia
' *============================================================================*
Option Explicit
Option Base 0
'
'--- Variables Privadas -------------------------------------------------------*
Dim mObj            As Metodo
Dim mCol            As Metodos
Dim mMdl            As MetodoModel
Dim mCtrl           As MetodoController
Dim i               As Integer




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
        .CriteriosAgrupacion = grpPeso
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
    '   Case error 7: Probar validación de filtros
    '
    Set mObj = New Metodo
    With mObj
        .TipoProcedimiento = mtdEstadCombinacion
        .CriteriosOrdenacion = ordDesviacion
        .TipoMuestra = False
        .DiasAnalisis = 90
        .ModalidadJuego = LP_LB_6_49
    End With
    PrintMetodo mObj
    '
    '   Case error 8: Probar validación de Tipo Muestra Dias
    '
    Set mObj = New Metodo
    With mObj
        .TipoProcedimiento = mtdEstadistico
        .CriteriosOrdenacion = ordDesviacion
        .TipoMuestra = False
        .DiasAnalisis = 0
        .ModalidadJuego = LP_LB_6_49
    End With
    PrintMetodo mObj
    '
    '   Case error 9: Probar validación de Tipo Muestra Registros
    '
    Set mObj = New Metodo
    With mObj
        .TipoProcedimiento = mtdEstadistico
        .CriteriosOrdenacion = ordDesviacion
        .TipoMuestra = True
        .NumeroSorteos = 0
        .ModalidadJuego = LP_LB_6_49
    End With
    PrintMetodo mObj
    
    
  On Error GoTo 0
MetodoTest__CleanExit:
    Exit Sub
MetodoTest_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    Call HandleException(ErrNumber, ErrDescription, "Lot_PqtSugerenciasTesting.MetodoTest", ErrSource)
    Call MsgBox(ErrDescription, vbCritical Or vbSystemModal, ThisWorkbook.Name)
    Call Trace("CERRAR")
End Sub





'------------------------------------------------------------------------------*
' Procedimiento  : MetodosTest
' Fecha          : lu., 23/mar/2020 20:17:14
' Propósito      : Pruebas unitarias clase Metodos
'------------------------------------------------------------------------------*
Private Sub MetodosTest()
    Dim mObj        As Metodos
    Dim mMtd        As Metodo
  On Error GoTo MetodosTest_Error
    '
    '   TestUnit 01: Objeto en vacio
    '
    Set mObj = New Metodos
    PrintMetodos mObj
    '
    '   TestUnit 02: Metodo ADD
    '
    Set mMtd = New Metodo
    With mMtd
        .Id = 564
        .TipoProcedimiento = mtdSinDefinir
        .CriteriosAgrupacion = grpDecenas
        .CriteriosOrdenacion = ordDesviacion
        .DiasAnalisis = 90
        .ModalidadJuego = LP_LB_6_49
    End With
    mObj.Add mMtd
    PrintMetodos mObj
    '
    '   TestUnit 03: Colección de 3 objetos
    '
    Set mMtd = New Metodo
    With mMtd
        .Id = 565
        .TipoProcedimiento = mtdBomboCargado
        .CriteriosAgrupacion = grpParidad
        .CriteriosOrdenacion = ordAusencia
        .DiasAnalisis = 90
        .ModalidadJuego = GP_5_54
    End With
    mObj.Add mMtd
    Set mMtd = New Metodo
    With mMtd
        .Id = 566
        .TipoProcedimiento = mtdEstadCombinacion
        .CriteriosAgrupacion = grpPeso
        .CriteriosOrdenacion = ordFrecuencia
        .DiasAnalisis = 90
        .ModalidadJuego = EU_5_50
    End With
    mObj.Add mMtd
    PrintMetodos mObj
    '
    '   TestUnit 04: Propiedad Count
    '
    Debug.Print "Numero de metodos( 3 ) = " & mObj.Count
    '
    '   TestUnit 05: Items
    '
    Debug.Print "Numero de Items ( 3 ) = " & mObj.Items.Count
    '
    '   TestUnit 06: MarkForDelete
    '
    mObj.MarkForDelete (2)
    PrintMetodos mObj
    '
    '   TestUnit 07: Undelete
    '
    mObj.Undelete (2)
    PrintMetodos mObj
    '
    '   TestUnit 08: Delete
    '
    mObj.Delete 1
    PrintMetodos mObj
    '
    '   TestUnit 09: Clear
    '
    mObj.Clear
    PrintMetodos mObj
    '
    '
  On Error GoTo 0
MetodosTest__CleanExit:
    Exit Sub
MetodosTest_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    Call HandleException(ErrNumber, ErrDescription, "Lot_PqtSugerenciasTesting.MetodosTest", ErrSource)
    Call MsgBox(ErrDescription, vbCritical Or vbSystemModal, ThisWorkbook.Name)
    Call Trace("CERRAR")
End Sub






'------------------------------------------------------------------------------*
' Procedimiento  : MetodoModelTest
' Fecha          : sá., 07/mar/2020 23:53:15
' Propósito      : Pruebas unitarias clase MetodoModel
'------------------------------------------------------------------------------*
Private Sub MetodoModelTest()
  On Error GoTo MetodoModelTest_Error
    '
    '   TestUnit 01: Creación del Objeto
    '
    Set mMdl = New MetodoModel
    PrintMetodoModel mMdl
    '
    '   TestUnit 02: Creación nuevo Metodo
    '
    With mMdl.Metodo
        .TipoProcedimiento = mtdAleatorio
        .Pronosticos = 6
        .ModalidadJuego = LP_LB_6_49
        .CriteriosOrdenacion = ordProbabilidad
        .SentidoOrdenacion = False
        .CriteriosAgrupacion = grpTerminacion
        .TipoMuestra = True
        .NumeroSorteos = 50
        .DiasAnalisis = 90
    End With
    PrintMetodoModel mMdl
    If mMdl.Add Then
        Debug.Print "Método añandido Id (" & CStr(mMdl.Metodo.Id) & ")"
    End If
    '
    '   TestUnit 03: Busqueda con exito
    '
    If mMdl.GetRecord(1) Then
        PrintMetodo mMdl.Metodo
    Else
        Debug.Print "#Error MetodoModel.GetRecord()"
    End If
    '
    '   TestUnit 04: Busqueda sin exito
    '
    On Error Resume Next
    If mMdl.GetRecord(15) Then
        If Err.Number <> 0 Then
            Debug.Print "#Error: " & Err.Number & " - " & Err.Description
        Else
            PrintMetodoModel mMdl
        End If
    End If
    On Error GoTo MetodoModelTest_Error
    '
    '   TestUnit 05: Save
    '
    If mMdl.GetRecord(1) Then
        With mMdl.Metodo
            .CriteriosAgrupacion = grpPeso
            .CriteriosOrdenacion = ordFrecuencia
            .DiasAnalisis = 20
        End With
        If mMdl.Save Then
            PrintMetodo mMdl.Metodo
        Else
            Debug.Print "#Error MetodoModel.Save()"
        End If
    End If
    '
    '   TestUnit 06: mMdl.Del
    '
    If mMdl.Del(1) Then
        Debug.Print "TestUnit 06: mMdl.Del OK"
    Else
        Debug.Print "#Error MetodoModel.Del()"
    End If
    '
    '   Crear 10 registros
    '
    '
    With mMdl.Metodo
        .TipoProcedimiento = mtdAleatorio
        .Pronosticos = 6
        .ModalidadJuego = LP_LB_6_49
        .CriteriosOrdenacion = ordProbabilidad
        .SentidoOrdenacion = False
        .CriteriosAgrupacion = grpTerminacion
        .TipoMuestra = True
        .NumeroSorteos = 50
        .DiasAnalisis = 90
    End With
    For i = 0 To 10
        mMdl.Add
    Next i
    '
    '   TestUnit 07: mMdl.GetFirst
    '
    If mMdl.GetFirst Then
        PrintMetodo mMdl.Metodo
    Else
        Debug.Print "#Error MetodoModel.GetFirst()"
    End If
    '
    '   TestUnit 08: mMdl.GetNext
    '
    If mMdl.GetNext(1) Then
        PrintMetodo mMdl.Metodo
    Else
        Debug.Print "#Error MetodoModel.GetNext()"
    End If
    '
    '   TestUnit 09: mMdl.GetLast
    '
    If mMdl.GetLast Then
        PrintMetodo mMdl.Metodo
    Else
        Debug.Print "#Error MetodoModel.GetLast()"
    End If
    '
    '   TestUnit 10: mMdl.GetPrev
    '
    If mMdl.GetPrev(5) Then
        PrintMetodo mMdl.Metodo
    Else
        Debug.Print "#Error MetodoModel.GetPrev()"
    End If
    '
    '   TestUnit 11: mMdl.Search
    '
    mMdl.LinePerPage = 5
    If mMdl.Search(2) Then
        PrintMetodoModel mMdl
    Else
        Debug.Print "#Error MetodoModel.Search(2)"
    End If
   
    
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
    Dim mFrm As frmMetodoSelectView
    
  On Error GoTo MetodoControllerTest_Error
    '
    '   1.-  Metodo
    '
    Set mCtrl = New MetodoController
    
    
    
    
    
  On Error GoTo 0
MetodoControllerTest__CleanExit:
    Exit Sub
MetodoControllerTest_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    Call HandleException(ErrNumber, ErrDescription, "Lot_PqtSugerenciasTesting.MetodosTest", ErrSource)
    Call MsgBox(ErrDescription, vbCritical Or vbSystemModal, ThisWorkbook.Name)
    Call Trace("CERRAR")
End Sub




'------------------------------------------------------------------------------*
' Procedimiento  : FiltroCombinacionTest
' Fecha          : sá., 07/mar/2020 23:53:15
' Propósito      : Pruebas unitarias clase MetodoController
'------------------------------------------------------------------------------*
Private Sub FiltroCombinacionTest()
    Dim mObj As FiltroCombinacion
    Dim mComb As Combinacion
    Dim mResBool As Boolean
  On Error GoTo FiltroCombinacionTest_Error
    '
    '   UnitTest 1.- Objeto Vacio
    '
    Set mObj = New FiltroCombinacion
    PrintFiltroCombinacion mObj
    '
    '   UnitTest 2.- Objeto definido
    '
    With mObj
        .TipoFiltro = tfConsecutivos
        .FilterValue = "3/2/0"
        .MultiplesFiltros = False
    End With
    PrintFiltroCombinacion mObj
    '
    '   UnitTest 3.- Evaluar una combinación
    '
    Set mComb = New Combinacion
    mComb.Texto = "12-17-19-20-35-36"
    With mObj
        .TipoFiltro = tfConsecutivos
        .FilterValue = "2"
        .MultiplesFiltros = False
    End With
    PrintFiltroCombinacion mObj
    mResBool = mObj.EvaluarCombinacion(mComb)
    Debug.Print "FiltroCombinación : " & mComb.Texto & " -->" & mResBool
    '
    '   UnitTest 4.- Agregar un valor para el analisis OR
    '
    Set mObj = New FiltroCombinacion
    With mObj
        .TipoFiltro = tfAltoBajo
        .FilterValue = "3/3"
        .AddFilterValue "4/2"
    End With
    PrintFiltroCombinacion mObj
    '
    '   UnitTest 5.- Parse de un filtro
    '
    Set mObj = New FiltroCombinacion
    mObj.Parse "(1)Paridad:[3/3,4/2,2/4]"
    PrintFiltroCombinacion mObj
    '
    '   UnitTest 6.- Probar una combinación con varios valores
    '
    Set mComb = New Combinacion
    mComb.Texto = "12-17-19-20-35-36"
    mResBool = mObj.EvaluarCombinacion(mComb)
    Debug.Print "FiltroCombinación : " & mComb.Texto & " -->" & mResBool
    
    
    
  On Error GoTo 0
FiltroCombinacionTest__CleanExit:
    Exit Sub
FiltroCombinacionTest_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    Call HandleException(ErrNumber, ErrDescription, "Lot_PqtSugerenciasTesting.FiltroCombinacionTest", ErrSource)
    Call MsgBox(ErrDescription, vbCritical Or vbSystemModal, ThisWorkbook.Name)
    Call Trace("CERRAR")
End Sub


'------------------------------------------------------------------------------*
' Procedimiento  : FiltrosCombinacionTest
' Fecha          :
' Propósito      : Pruebas unitarias colección FiltrosCombinacion
'------------------------------------------------------------------------------*
Private Sub FiltrosCombinacionTest()
    Dim mObj  As FiltrosCombinacion
    Dim mFltr As FiltroCombinacion
    
  On Error GoTo FiltrosCombinacionTest_Error
    '
    '   UnitTest 1.- Objeto Vacio
    '
    Set mObj = New FiltrosCombinacion
    PrintFiltrosCombinacion mObj
    
    '
    '   UnitTest 2.- Agregar un filtro
    '
    Set mFltr = New FiltroCombinacion
    With mFltr
        .TipoFiltro = tfSuma
        .FilterValue = "150..200"
    End With
    mObj.Add mFltr
    PrintFiltrosCombinacion mObj
    '
    '   UnitTest 3.- Agregar dos filtros distintos
    '
    Set mObj = New FiltrosCombinacion
    Set mFltr = New FiltroCombinacion
    With mFltr
        .TipoFiltro = tfParidad
        .FilterValue = "3/3"
    End With
    mObj.Add mFltr
    Set mFltr = New FiltroCombinacion
    With mFltr
        .TipoFiltro = tfAltoBajo
        .FilterValue = "4/2"
    End With
    mObj.Add mFltr
    PrintFiltrosCombinacion mObj
    '
    '   UnitTest 4.- Agregar dos filtros iguales
    '
    Set mObj = New FiltrosCombinacion
    Set mFltr = New FiltroCombinacion
    With mFltr
        .TipoFiltro = tfParidad
        .FilterValue = "3/3"
    End With
    mObj.Add mFltr
    Set mFltr = New FiltroCombinacion
    With mFltr
        .TipoFiltro = tfParidad
        .FilterValue = "4/2"
    End With
    mObj.Add mFltr
    PrintFiltrosCombinacion mObj
    '
    '   UnitTest 5.- Parse de un filtro
    '
    Set mObj = New FiltrosCombinacion
    mObj.Parse "(6)Suma:[150..200]"
    PrintFiltrosCombinacion mObj
    '
    '   UnitTest 6.- Parse de dos filtros
    '
    Set mObj = New FiltrosCombinacion
    mObj.Parse "(2)Peso:[3/3]|(1)Paridad:[5/1]"
    PrintFiltrosCombinacion mObj


  On Error GoTo 0
FiltrosCombinacionTest__CleanExit:
    Exit Sub
FiltrosCombinacionTest_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    Call HandleException(ErrNumber, ErrDescription, "Lot_PqtSugerenciasTesting.FiltrosCombinacionTest", ErrSource)
    Call MsgBox(ErrDescription, vbCritical Or vbSystemModal, ThisWorkbook.Name)
    Call Trace("CERRAR")
End Sub


'------------------------------------------------------------------------------*
' Procedimiento  : RealizarSugerenciaTest
' Fecha          :
' Propósito      : Pruebas del Caso de Uso Realizar Sugerencia
'------------------------------------------------------------------------------*
Private Sub SugerenciaTest()
    Dim mObj As Sugerencia
  On Error GoTo SugerenciaTest_Error
    '
    '   UnitTest 1.- Objeto Vacio
    '
    Set mObj = New Sugerencia
    PrintSugerencia mObj
    '
    '   UnitTest 2.- Objeto con datos
    '
    With mObj
        .Combinacion.Texto = "15-48-07-13-22-31"
        .Entidad.Id = 23
        .Metodo.TipoProcedimiento = mtdAleatorio
        .Metodo.Pronosticos = 6
        .Metodo.CriteriosAgrupacion = grpDecenas
        .Metodo.CriteriosOrdenacion = ordDesviacion
        .Metodo.TipoMuestra = True
        .Metodo.NumeroSorteos = 50
        .Modalidad = LP_LB_6_49
        .Parametros.FechaAnalisis = Date
    End With
    PrintSugerencia mObj
    '
    '   UnitTest 3.- Objeto No valido, metodo no valido
    '
    Set mObj = New Sugerencia
    With mObj
        .Combinacion.Texto = "05-14-18-22-47-31"
        .Entidad.Id = 12
        .Metodo.TipoProcedimiento = mtdSinDefinir
        .Metodo.Pronosticos = 6
        .Metodo.TipoMuestra = True
        .Metodo.NumeroSorteos = 50
        .Modalidad = LP_LB_6_49
        .Parametros.FechaAnalisis = Date
    End With
    If Not mObj.IsValid Then
        Debug.Print " La sugerencia no es valida "
    End If
    PrintSugerencia mObj
    
    '
    '   UnitTest 4.- Objeto No valido, falta fecha análisis
    '
    '
    '   UnitTest 5.- Objeto Valido
    '
    
    On Error GoTo 0
SugerenciaTest__CleanExit:
    Exit Sub
SugerenciaTest_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    Call HandleException(ErrNumber, ErrDescription, "Lot_PqtSugerenciasTesting.SugerenciaTest", ErrSource)
    Call MsgBox(ErrDescription, vbCritical Or vbSystemModal, ThisWorkbook.Name)
    Call Trace("CERRAR")
End Sub




'------------------------------------------------------------------------------*
' Procedimiento  : RealizarSugerenciaTest
' Fecha          :
' Propósito      : Pruebas del Caso de Uso Realizar Sugerencia
'------------------------------------------------------------------------------*
Private Sub RealizarSugerenciaTest()
    Dim mObj As RealizarSugerencia
    Dim mSug As Sugerencia
    Dim mMtd As Metodo
    
  On Error GoTo RealizarSugerenciaTest_Error
    '
    '   UnitTest 1.- Objeto Vacio
    '
    Set mObj = New RealizarSugerencia
    PrintRealizarSugerencia mObj
    '
    '   UnitTest 2.- Sugerencia metodo aleatorio
    '
    Set mMtd = New Metodo
    With mMtd
        .TipoProcedimiento = mtdAleatorio
        .ModalidadJuego = LP_LB_6_49
        .Pronosticos = 6
    End With
    Set mSug = mObj.GetSugerencia(mMtd, #5/4/2020#)
    PrintSugerencia mSug
    '
    PrintRealizarSugerencia mObj
    
    '
    '   UnitTest 3.- Sugerencia metodo Bombo
    '
    '
    '   UnitTest 4.- Sugerencia metodo BomboCargado
    '
    '
    '   UnitTest 5.- Sugerencia metodo Estadistica
    '
    '
    '   UnitTest 6.- Sugerencia metodo Estadistica Combinación
    '
    '
    '   UnitTest 7.- Sugerencia metodo Aleatorio con filtros
    '
    '
    '   UnitTest 8.- Sugerencia metodo Estadistica con filtros
    '
    
    
    
    
    
    
    
    
    
    Err.Raise ERR_TODO, "Lot_PqtSugerenciasTesting.RealizarSugerenciaTest", MSG_TODO
    
  
  On Error GoTo 0
RealizarSugerenciaTest__CleanExit:
    Exit Sub
RealizarSugerenciaTest_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    Call HandleException(ErrNumber, ErrDescription, "Lot_PqtSugerenciasTesting.RealizarSugerenciaTest", ErrSource)
    Call MsgBox(ErrDescription, vbCritical Or vbSystemModal, ThisWorkbook.Name)
    Call Trace("CERRAR")
End Sub


'------------------------------------------------------------------------------*
' Procedimiento  : PrintSugerencia
' Fecha          :
' Propósito      : Visualizar el objeto RealizarSugerencia
'------------------------------------------------------------------------------*
Private Sub PrintSugerencia(mObj As Sugerencia)
    Debug.Print "==> Pruebas Sugerencia"
    '-> Propiedades
    Debug.Print vbTab & "Combinacion       = " & mObj.Combinacion.Texto
    Debug.Print vbTab & "FechaAlta         = " & mObj.Entidad.FechaAlta
    Debug.Print vbTab & "FechaModificacion = " & mObj.Entidad.FechaModificacion
    Debug.Print vbTab & "Id                = " & mObj.Entidad.Id
    Debug.Print vbTab & "Metodo            = " & mObj.Metodo.ToString()
    Debug.Print vbTab & "Modalidad         = " & mObj.Modalidad
    Debug.Print vbTab & "Parametros        = " & mObj.Parametros.ToString()
'    '-> Metodos
    Debug.Print vbTab & "MensajeError()    = " & mObj.MensajeError()
    Debug.Print vbTab & "ToString()        = " & mObj.ToString()
    Debug.Print vbTab & "IsValid()         = " & mObj.IsValid()
End Sub


'------------------------------------------------------------------------------*
' Procedimiento  : PrintRealizarSugerencia
' Fecha          :
' Propósito      : Visualizar el objeto RealizarSugerencia
'------------------------------------------------------------------------------*
Private Sub PrintRealizarSugerencia(mObj As RealizarSugerencia)
    Debug.Print "==> Pruebas RealizarSugerencia"
'    '-> Propiedades
'    Debug.Print vbTab & "FilterValue         =" & mObj.
'    Debug.Print vbTab & "MultiplesFiltros    =" & mObj.MultiplesFiltros
'    Debug.Print vbTab & "NameFiltro          =" & mObj.NameFiltro
'    Debug.Print vbTab & "TipoFiltro          =" & mObj.TipoFiltro
'    '-> Metodos
'    Debug.Print vbTab & "AddFilterValue()    =" & "#Metodo" 'mObj.AddFilterValue
'    Debug.Print vbTab & "EvaluarCombinacion()=" & "#Metodo" 'mObj.EvaluarCombinacion
'    Debug.Print vbTab & "GetValoresFiltros() =" & "#Metodo" 'mObj.GetValoresFiltros
'    Debug.Print vbTab & "Parse()             =" & "#Metodo" 'mObj.Parse
'    Debug.Print vbTab & "ToString()          =" & mObj.ToString
End Sub



'------------------------------------------------------------------------------*
' Procedimiento  : PrintFiltroCombinacion
' Fecha          : sá., 07/mar/2020 23:53:15
' Propósito      : Visualizar el objeto
'------------------------------------------------------------------------------*
Private Sub PrintFiltroCombinacion(mObj As FiltroCombinacion)
    Debug.Print "==> Pruebas FiltroCombinacion"
    '-> Propiedades
    Debug.Print vbTab & "FilterValue         =" & mObj.FilterValue
    Debug.Print vbTab & "MultiplesFiltros    =" & mObj.MultiplesFiltros
    Debug.Print vbTab & "NameFiltro          =" & mObj.NameFiltro
    Debug.Print vbTab & "TipoFiltro          =" & mObj.TipoFiltro
    '-> Metodos
    Debug.Print vbTab & "AddFilterValue()    =" & "#Metodo" 'mObj.AddFilterValue
    Debug.Print vbTab & "EvaluarCombinacion()=" & "#Metodo" 'mObj.EvaluarCombinacion
    Debug.Print vbTab & "GetValoresFiltros() =" & "#Metodo" 'mObj.GetValoresFiltros
    Debug.Print vbTab & "Parse()             =" & "#Metodo" 'mObj.Parse
    Debug.Print vbTab & "ToString()          =" & mObj.ToString
End Sub



'------------------------------------------------------------------------------*
' Procedimiento  : PrintFiltrosCombinacion
' Fecha          : sá., 07/mar/2020 23:53:15
' Propósito      : Visualizar el objeto
'------------------------------------------------------------------------------*
Private Sub PrintFiltrosCombinacion(mObj As FiltrosCombinacion)
    Debug.Print "==> Pruebas FiltrosCombinacion"
    '-> Propiedades
    Debug.Print vbTab & "Count      =" & mObj.Count
    Debug.Print vbTab & "Items      =" & mObj.Items.Count
    '-> Metodos
    Debug.Print vbTab & "Add()      =" & "#Metodo" 'mObj.Add
    Debug.Print vbTab & "Parse()    =" & "#Metodo" 'mObj.Parse
    Debug.Print vbTab & "ToString() =" & mObj.ToString
End Sub



'------------------------------------------------------------------------------*
' Procedimiento  : PrintMetodos
' Fecha          : lu., 23/mar/2020 20:19:04
' Propósito      : Visualizar el objeto
'------------------------------------------------------------------------------*
Private Sub PrintMetodos(mObj As Metodos)
    Dim mMtd As Metodo
    Debug.Print "==> Pruebas Metodos"
    '-> Propiedades
    Debug.Print vbTab & "Add                 =" & "#Metodo" 'mObj.Add
    Debug.Print vbTab & "Clear               =" & "#Metodo" 'mObj.Clear
    Debug.Print vbTab & "Count               =" & mObj.Count
    Debug.Print vbTab & "Delete              =" & "#Metodo" 'mObj.Delete
    Debug.Print vbTab & "Items.Count         =" & mObj.Items.Count
    Debug.Print vbTab & "MarkForDelete       =" & "#Metodo" 'mObj.MarkForDelete
    Debug.Print vbTab & "Undelete            =" & "#Metodo" 'mObj.Undelete
    For Each mMtd In mObj.Items
        Debug.Print vbTab & mMtd.ToString
    Next mMtd
End Sub




'------------------------------------------------------------------------------*
' Procedimiento  : PrintMetodo
' Fecha          : sá., 07/mar/2020 23:53:15
' Propósito      : Visualizar el objeto
'------------------------------------------------------------------------------*
Private Sub PrintMetodo(mObj As Metodo)
     Debug.Print "==> Pruebas Metodo"
     '-> Propiedades
     Debug.Print vbTab & "Id                         = " & mObj.Id
     Debug.Print vbTab & "TipoProcedimiento          = " & mObj.TipoProcedimiento
     Debug.Print vbTab & "Pronosticos                = " & mObj.Pronosticos
     Debug.Print vbTab & "ModalidadJuego             = " & mObj.ModalidadJuego
     Debug.Print vbTab & "CriteriosOrdenacion        = " & mObj.CriteriosOrdenacion
     Debug.Print vbTab & "SentidoOrdenacion          = " & mObj.SentidoOrdenacion
     Debug.Print vbTab & "CriteriosAgrupacion        = " & mObj.CriteriosAgrupacion
     Debug.Print vbTab & "TipoMuestra                = " & mObj.TipoMuestra
     Debug.Print vbTab & "NumeroSorteos              = " & mObj.NumeroSorteos
     Debug.Print vbTab & "DiasAnalisis               = " & mObj.DiasAnalisis
     Debug.Print vbTab & "FechaAlta                  = " & mObj.EntidadNegocio.FechaAlta
     Debug.Print vbTab & "FechaModificacion          = " & mObj.EntidadNegocio.FechaModificacion
     '-> Metodos
     Debug.Print vbTab & "AgrupacionToString()       = " & mObj.AgrupacionToString()
     Debug.Print vbTab & "OrdenacionToString()       = " & mObj.OrdenacionToString()
     Debug.Print vbTab & "TipoProcedimientoTostring()= " & mObj.TipoProcedimientoTostring()
     Debug.Print vbTab & "ToString()                 = " & mObj.ToString()
     Debug.Print vbTab & "IsValid()                  = " & mObj.IsValid()
     Debug.Print vbTab & "GetMessage()               = " & mObj.GetMessage()
End Sub




'------------------------------------------------------------------------------*
' Procedimiento  : PrintMetodoModel
' Fecha          : mi., 11/mar/2020 19:45:28
' Propósito      : Visualizar el objeto MetodoModel
'------------------------------------------------------------------------------*
Private Sub PrintMetodoModel(mObj As MetodoModel)
    Debug.Print "==> Pruebas MetodoModel"
    '-> Propiedades
    Debug.Print vbTab & "LinePerPage   = " & mObj.LinePerPage
    Debug.Print vbTab & "Metodo        = " & mObj.Metodo.ToString()
    Debug.Print vbTab & "Metodos.Count = " & mObj.Metodos.Count
    Debug.Print vbTab & "TotalPages    = " & mObj.TotalPages
    Debug.Print vbTab & "TotalRecords  = " & mObj.TotalRecords
    Debug.Print vbTab & "CurrentPage   = " & mObj.CurrentPage
End Sub


' *===========(EOF): Lot_PqtSugerenciasTesting.bas
