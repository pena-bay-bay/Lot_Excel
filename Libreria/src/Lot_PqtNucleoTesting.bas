Attribute VB_Name = "Lot_PqtNucleoTesting"
' *============================================================================*
' *
' *     Fichero    : Lot_PqtNucleoTesting.bas
' *
' *     Autor      : Carlos Almela Baeza
' *     Creación   : sáb, 01/nov/2014 21:01:24
' *     Versión    : 1.0
' *     Propósito  : Colección de pruebas unitarias de las clases del paquete
' *                  Nucleo:
' *                    - Periodo
' *                    - Parametro
' *                    - Parametros
' *                    - ParametroEngine
' *                    - DataBaseExcel
' *                    - EntidadNegocio
' *
' *============================================================================*
Option Explicit
Option Base 0
'
'--- Variables Privadas -------------------------------------------------------*
Private mDBase As DataBaseExcel
'---------------------------------------------------------------------------------------
' Procedure : PqtNucleoTest
' Author    : CHARLY
' Date      : ma., 16/jul/2019 13:44:37
' Purpose   : Realizar el conjunto de pruebas del paquete nucleo
'---------------------------------------------------------------------------------------
'
Public Sub PqtNucleoTest()
    PeriodoTest
    ParametroTest
    ParametrosTest
    ParametrosEngineTest
    DataBaseExcelTest
    DataBaseExcelTest2
    EntidadNegocioTest
End Sub
'---------------------------------------------------------------------------------------
' Procedure : PeriodoTest
' Author    : CHARLY
' Date      : sáb, 01/nov/2014 21:01:47
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub PeriodoTest()
    Dim Obj         As Periodo
    Dim dFIni       As Date
    Dim dFFin       As Date
    Dim mLista      As Variant
    Dim frm         As frmSelPeriodo
    Dim cboPrueba   As Object
    
  On Error GoTo PeriodoTest_Error
    '
    '
    dFIni = #5/1/2017#
    dFFin = #7/6/2017#
    Set Obj = New Periodo
    '
    Obj.Init dFIni, dFFin
    '
    '
    PrintPeriodo Obj
    '
    '
    Obj.Tipo_Fecha = ctLoQueVadeAño
    '
    '
    PrintPeriodo Obj
    '
    '   Creamos un formulario contenedor de controles para referenciar un Combo
    '
    Set frm = New frmSelPeriodo
    Set cboPrueba = frm.cboPerMuestra
     
    mLista = Array(ctPersonalizadas, ctSemanaPasada, ctSemanaActual, ctMesActual, ctHoy, ctAyer, ctLoQueVadeMes, _
                                      ctLoQueVadeSemana)
 
    Obj.CargaCombo cboPrueba, mLista
    
    frm.Show
    
  On Error GoTo 0
PeriodoTest__CleanExit:
    Exit Sub
            
PeriodoTest_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    '   Audita el error
    Call HandleException(ErrNumber, ErrDescription, "Lot_PqtNucleoTesting.PeriodoTest", ErrSource)
    '   Informa del error
    Call MsgBox(ErrDescription, vbCritical Or vbSystemModal, ThisWorkbook.Name)
    Call Trace("CERRAR")
End Sub
'---------------------------------------------------------------------------------------
' Procedure : ParametrosTest
' Author    : CHARLY
' Date      : sáb, 01/nov/2014 21:01:53
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub ParametrosTest()
    Dim mObj        As Parametros
    Dim mPar        As Parametro
    
  On Error GoTo ParametrosTest_Error
    '
    '   Prueba método Add
    '
    Set mPar = New Parametro
    With mPar
        .Id = 563
        .Nombre = "MI_PRUEBA"
        .Orden = 1
        .Tipo = parTexto
        .Descripcion = "Descripción de la variable MI_PRUEBA"
        .Valor = "Valor de Prueba"
    End With
    Set mObj = New Parametros
    mObj.Add mPar
    '
    '   Segundo Parametro
    '
    Set mPar = New Parametro
    With mPar
        .Id = 562
        .Nombre = "MI_PRUEBA_2"
        .Orden = 1
        .Tipo = parTexto
        .Descripcion = "Descripción de la variable MI_PRUEBA 2"
        .Valor = "Valor de Prueba 2"
    End With
    mObj.Add mPar
    '
    '   Visualiza la colección
    '
    PrintParametros mObj
    '
    '   Prueba propiedad Count
    '
    Debug.Print "=> Propiedad Count (2) => " & mObj.Count
    '
    '   Prueba método Items
    '
    Debug.Print "=> Prueba Items"
    For Each mPar In mObj.Items
        Debug.Print vbTab & "* (" & mPar.Id & ") Valor=>" & mPar.Valor
    Next mPar
    '
    '   Prueba método MarkForDelete
    '
    mObj.MarkForDelete 1
    Debug.Print "=> Prueba MarkForDelete"
    Debug.Print vbTab & "* (" & mObj.Items(1).Id & ") Valor=> " & mObj.Items(1).EntidadNegocio.MarkForDelete
    '
    '   Prueba método Undelete
    '
    mObj.Undelete 1
    Debug.Print "=> Prueba Undelete"
    Debug.Print vbTab & "* (" & mObj.Items(1).Id & ") Valor=> " & mObj.Items(1).EntidadNegocio.MarkForDelete
    '
    '   Activamos el control del error porque queremos desmarcar un elemento inexistente
    '
    On Error Resume Next
    mObj.Undelete 5
    If Err.Number > 0 Then
        Debug.Print "#Err:" & Err.Number & "-" & Err.Description
    End If
    On Error GoTo 0
    '
    '   Prueba método Delete
    '
    mObj.Delete 1
    '
    Debug.Print "=> Metodo Delete (1) => " & mObj.Count
    '
    '   Activamos el control del error porque queremos borrar un elemento inexistente
    '
    On Error Resume Next
    mObj.Delete 5
    If Err.Number > 0 Then
        Debug.Print "#Err:" & Err.Number & "-" & Err.Description
    End If
    On Error GoTo 0
    '
    '   Prueba metodo Clear
    '
    Debug.Print "=> Prueba Clear"
    mObj.Clear
    PrintParametros mObj
    
  On Error GoTo 0
ParametrosTest__CleanExit:
    Exit Sub
            
ParametrosTest_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    '   Audita el error
    Call HandleException(ErrNumber, ErrDescription, "Lot_PqtNucleoTesting.ParametrosTest", ErrSource)
    '   Informa del error
    Call MsgBox(ErrDescription, vbCritical Or vbSystemModal, Application.Caption)
    Call Trace("CERRAR")
End Sub

'---------------------------------------------------------------------------------------
' Procedure : ParametroTest
' Author    : CHARLY
' Date      : sáb, 01/nov/2014 21:02:02
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub ParametroTest()
    Dim oPar As Parametro
    Dim iPar As Integer
    Dim dPar As Date
    Dim bPar As Boolean
    Dim pPar As Double
  
  On Error GoTo ParametroTest_Error
    
    
    Set oPar = New Parametro
    With oPar
        .Descripcion = "Esta es una variable de prueba"
        .Nombre = "VARIABLE"
        .Tipo = parTexto
        .Valor = "Ejemplo"
    End With
    PrintParametro oPar
    '
    ' Prueba entero
    '
    iPar = 3294
    oPar.Valor = iPar
    PrintParametro oPar
    '
    ' prueba fecha
    '
    dPar = #1/1/2014#
    oPar.Valor = dPar
    oPar.Tipo = parFecha
    PrintParametro oPar
    '
    ' prueba Doble
    '
    pPar = 12536.254
    oPar.Valor = pPar
    oPar.Tipo = parDecimalPrecision
    PrintParametro oPar
  
  On Error GoTo 0
ParametroTest__CleanExit:
    Exit Sub
            
ParametroTest_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    '   Audita el error
    Call HandleException(ErrNumber, ErrDescription, "Lot_PqtNucleoTesting.ParametroTest", ErrSource)
    '   Informa del error
    Call MsgBox(ErrDescription, vbCritical Or vbSystemModal, ThisWorkbook.Name)
    Call Trace("CERRAR")
End Sub

'---------------------------------------------------------------------------------------
' Procedure : ParametrosEngineTest
' Author    : CHARLY
' Date      : sáb, 01/nov/2014 21:02:04
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub ParametrosEngineTest()
    Dim mEng    As ParametroEngine
    Dim mCol    As Parametros
    Dim mObj    As Parametro
    
  On Error GoTo ParametrosEngineTest_Error
    '
    '   Prueba GetNewParametro
    '
    Set mEng = New ParametroEngine
    '
    Set mObj = mEng.GetNewParametro
    '
    PrintParametro mObj
    '----------------------------------------------------
    '
    '   Prueba SetParametro
    '
    With mObj
        .Descripcion = "Variable de prueba EngineTest"
        .Nombre = "MIVARIABLE"
        .Tipo = parEntero
        .Valor = 10
    End With
    '
    PrintParametro mObj
    '
    mEng.SetParametro mObj
    '----------------------------------------------------
    '
    '   Modificar datos del parametro
    '
    Set mObj = mEng.GetParametroById(mObj.Id)
        
    With mObj
        .EntidadNegocio.IsNew = False
        .Descripcion = "Variable de prueba EngineTest Modificado"
        .Nombre = "MIVARIABLE2"
        .Valor = 365
    End With
    '
    mEng.SetParametro mObj
    '----------------------------------------------------
    '
    '   Modificar datos del parametro por Nombre
    '
    Set mCol = mEng.GetParametroByName("MIVARIABLE2")
    With mCol.Items(1)
        .EntidadNegocio.IsNew = False
        .Descripcion = "Variable de prueba EngineTest Modificado"
        .Nombre = "MIVARIABLE2"
        .Valor = 365
    End With
    
    mEng.SetParametro mObj
    
    '----------------------------------------------------
    '
    '   Prueba de borrar parametro
    '
    mObj.EntidadNegocio.IsNew = False
    mObj.EntidadNegocio.MarkForDelete = True
    '
    mEng.SetParametro mObj
    '----------------------------------------------------
    '
    '   Prueba de GetTiposParametros
    '
    Dim mVar As Variant
    mVar = mEng.GetTipoParametros()
    Debug.Print "Nombres de Tipo de Variables: " & mVar(0) & ", " & mVar(1) & "..."
    
    '----------------------------------------------------
    '
    '   Manejo de colecciones
    '   Creamos 5 objetos
    '   Los modificamos
    '   y actualizamos la colección
    '
    
    
    
    
    
    
    
    '----------------------------------------------------
    '
    '  Cerramos la Base de datos
    '
    Set mDBase = New DataBaseExcel
    mDBase.Abrir
    mDBase.Cerrar
    Set mDBase = Nothing
  On Error GoTo 0
ParametrosEngineTest__CleanExit:
    Exit Sub
        
ParametrosEngineTest_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    '   Audita el error
    Call HandleException(ErrNumber, ErrDescription, "Lot_PqtNucleoTesting.ParametrosEngineTest", ErrSource)
    '   Informa del error
    Call MsgBox(ErrDescription, vbCritical Or vbSystemModal, ThisWorkbook.Name)
    Call Trace("CERRAR")
End Sub



'------------------------------------------------------------------------------*
' Procedimiento  : PrintEntidadNegocio
' Fecha          : 21/sep/2018
' Propósito      : Visualiza las propiedades y metodos de la clase EntidadNegocio
' Parámetros     : EntidadNegocio
'------------------------------------------------------------------------------*
'
Private Sub PrintEntidadNegocio(Obj As EntidadNegocio)
    Debug.Print "==> EntidadNegocio "
    Debug.Print vbTab & "ClassStorage  = " & Obj.ClassStorage
    Debug.Print vbTab & "FechaAlta     = " & Obj.FechaAlta
    Debug.Print vbTab & "FechaBaja     = " & Obj.FechaBaja
    Debug.Print vbTab & "FechaModificacion  = " & Obj.FechaModificacion
    Debug.Print vbTab & "ID            = " & Obj.Id
    Debug.Print vbTab & "IsDirty       = " & Obj.IsDirty
    Debug.Print vbTab & "IsNew         = " & Obj.IsNew
    Debug.Print vbTab & "MarkForDelete = " & Obj.MarkForDelete
    Debug.Print vbTab & "Origen        = " & Obj.Origen
    Debug.Print vbTab & "Situacion     = " & Obj.Situacion
End Sub

'---------------------------------------------------------------------------------------
' Procedure : PintarPeriodo
' Author    : CHARLY
' Date      : sáb, 01/nov/2014 21:02:14
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub PrintPeriodo(Obj As Periodo)
    Debug.Print "==> Periodo "
    Debug.Print vbTab & "Dias          = " & Obj.Dias
    Debug.Print vbTab & "Fecha Final   = " & Obj.FechaFinal
    Debug.Print vbTab & "Fecha Inicial = " & Obj.FechaInicial
    Debug.Print vbTab & "Texto         = " & Obj.Texto
    Debug.Print vbTab & "Tipo Fecha    = " & Obj.Tipo_Fecha
End Sub

'---------------------------------------------------------------------------------------
' Procedure : PrintParametro
' Author    : CHARLY
' Date      : sáb, 01/nov/2014 21:02:19
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub PrintParametro(Obj As Parametro)
    Debug.Print "==> Parametro "
    Debug.Print vbTab & "Descripcion     = " & Obj.Descripcion
    Debug.Print vbTab & "EntidadNegocio  = " & Obj.EntidadNegocio.ClassStorage
    Debug.Print vbTab & "Fecha Alta      = " & Obj.FechaAlta
    Debug.Print vbTab & "Fecha Modif.    = " & Obj.FechaModificacion
    Debug.Print vbTab & "Id              = " & Obj.Id
    Debug.Print vbTab & "Nombre          = " & Obj.Nombre
    Debug.Print vbTab & "Orden           = " & Obj.Orden
    Debug.Print vbTab & "Tipo            = " & Obj.Tipo
    Debug.Print vbTab & "Valor           = " & Obj.Valor
    Debug.Print vbTab & "ToString()      = " & Obj.ToString
    Debug.Print vbTab & "TipoToString()  = " & Obj.TipoToString
End Sub
'---------------------------------------------------------------------------------------
' Procedure : PrintParametros
' Author    : Charly
' Date      : 21/12/2018
' Purpose   : Probar la colección  Parametros
'---------------------------------------------------------------------------------------
'
Private Sub PrintParametros(cObj As Parametros)
    Dim mObj As Parametro
    
    Debug.Print "==> Parametros "
    Debug.Print vbTab & "Count           = " & cObj.Count
    For Each mObj In cObj.Items
        Debug.Print vbTab & mObj.ToString
    Next mObj
End Sub
'---------------------------------------------------------------------------------------
' Procedure : ParametrosMetodoTest
' Author    : Charly
' Date      : 19/03/2012
' Purpose   : Probar la clase ParametrosMetodo
'---------------------------------------------------------------------------------------
'
'Private Sub ParametrosMetodoTest()
'    Dim m_objParMetodo As ParametrosMetodo
'
'    Set m_objParMetodo = New ParametrosMetodo
'
'
'    With m_objParMetodo
'        .CriteriosAgrupacion = grpDecenas
'        .CriteriosOrdenacion = ordProbabilidad
'        .DiasAnalisis = 45
'        .ID = 1
'        .ModalidadJuego = LP_LB_6_49
'        .NumeroSorteos = 40
'        .Orden = 1
'        .Pronosticos = 6
'        .SentidoOrdenacion = True
'    End With
'
'    Debug.Print "==> Pruebas ParametrosMetodoTest"
'    Debug.Print "Id                       = " & m_objParMetodo.ID
'    Debug.Print "Juego                    = " & m_objParMetodo.ModalidadJuego
'    Debug.Print "Criterio Ordenación      = " & m_objParMetodo.CriteriosOrdenacion
'    Debug.Print "Criterio Agrupación      = " & m_objParMetodo.CriteriosAgrupacion
'    Debug.Print "Dias de Analisis         = " & m_objParMetodo.DiasAnalisis
'    Debug.Print "Numero de Sorteos        = " & m_objParMetodo.NumeroSorteos
'    Debug.Print "Orden                    = " & m_objParMetodo.Orden
'    Debug.Print "Pronosticos              = " & m_objParMetodo.Pronosticos
'    Debug.Print "Sentido de la Ordenación = " & m_objParMetodo.SentidoOrdenacion
'    Debug.Print "OrdenacionToString()     = " & m_objParMetodo.OrdenacionToString()
'    Debug.Print "AgrupacionToString()     = " & m_objParMetodo.AgrupacionToString()
'    Debug.Print "ToString()               = " & m_objParMetodo.ToString()
'
'End Sub


'---------------------------------------------------------------------------------------
' Procedure : MetodoTest
' Author    : Charly
' Date      : 19/03/2012
' Purpose   :
'---------------------------------------------------------------------------------------
'
'Public Sub MetodoTest()
'    Dim m_objMetodo As metodo
'
'    Set m_objMetodo = New metodo
'
'    With m_objMetodo
'        .TipoProcedimiento = mtdEstadistico
'        .EntidadNegocio.FechaModificacion = Date
'        .EsMultiple = False
'        .Parametros.CriteriosAgrupacion = grpParidad
'        .Parametros.CriteriosOrdenacion = ordDesviacion
'        .Parametros.DiasAnalisis = 42
'        .Parametros.SentidoOrdenacion = True
'        .TipoMuestra = True
'    End With
'
'
'    Debug.Print "==> Pruebas Metodo"
'    Debug.Print "ClassStorage           =" & m_objMetodo.EntidadNegocio.ClassStorage
'    Debug.Print "FechaAlta              =" & m_objMetodo.EntidadNegocio.FechaAlta
'    Debug.Print "FechaBaja              =" & m_objMetodo.EntidadNegocio.FechaBaja
'    Debug.Print "FechaModificacion      =" & m_objMetodo.EntidadNegocio.FechaModificacion
'    Debug.Print "Id                     =" & m_objMetodo.EntidadNegocio.ID
'    Debug.Print "IsDirty                =" & m_objMetodo.EntidadNegocio.IsDirty
'    Debug.Print "IsNew                  =" & m_objMetodo.EntidadNegocio.IsNew
'    Debug.Print "MarkForDelete          =" & m_objMetodo.EntidadNegocio.MarkForDelete
'    Debug.Print "Origen                 =" & m_objMetodo.EntidadNegocio.Origen
'    Debug.Print "Situacion              =" & m_objMetodo.EntidadNegocio.Situacion
'    Debug.Print "EsMultiple             =" & m_objMetodo.EsMultiple
'    Debug.Print "AgrupacionToString     =" & m_objMetodo.Parametros.AgrupacionToString
'    Debug.Print "CriteriosAgrupacion    =" & m_objMetodo.Parametros.CriteriosAgrupacion
'    Debug.Print "CriteriosOrdenacion    =" & m_objMetodo.Parametros.CriteriosOrdenacion
'    Debug.Print "DiasAnalisis           =" & m_objMetodo.Parametros.DiasAnalisis
'    Debug.Print "Id                     =" & m_objMetodo.Parametros.ID
'    Debug.Print "ModalidadJuego         =" & m_objMetodo.Parametros.ModalidadJuego
'    Debug.Print "NumeroSorteos          =" & m_objMetodo.Parametros.NumeroSorteos
'    Debug.Print "Orden                  =" & m_objMetodo.Parametros.Orden
'    Debug.Print "OrdenacionToString     =" & m_objMetodo.Parametros.OrdenacionToString
'    Debug.Print "Pronosticos            =" & m_objMetodo.Parametros.Pronosticos
'    Debug.Print "SentidoOrdenacion      =" & m_objMetodo.Parametros.SentidoOrdenacion
'    Debug.Print "ToString               =" & m_objMetodo.Parametros.ToString
'    Debug.Print "TipoMuestra            =" & m_objMetodo.TipoMuestra
'    Debug.Print "TipoProcedimiento      =" & m_objMetodo.TipoProcedimiento
'
'End Sub

'------------------------------------------------------------------------------*
' Procedimiento  : DataBaseExcelTest2
' Fecha          : 11/12/2018
' Propósito      : Pruebas unitarias de la clase DataBaseExcel Modificaciones
'------------------------------------------------------------------------------*
'
Public Sub DataBaseExcelTest2()
    Dim mObj        As DataBaseExcel
    Dim mRango      As Range
  On Error GoTo DataBaseExcelTest2_Error
    '
    '   Creamos la base de datos
    '
    Set mObj = New DataBaseExcel
    '
    '   Abrir la base de datos
    '
    mObj.Abrir
    '
    '   GetLastRow
    '
    Set mRango = mObj.GetLastRow(tblBonoloto)
    PrintRango mRango
    '
    Set mRango = mObj.GetLastRow(tblGordo)
    PrintRango mRango
    '
    Set mRango = mObj.GetLastRow(tblParametros)
    PrintRango mRango
    '
    '   InsertRow
    '
    With mRango
        .Cells(1, 1).Value = 3                        'N
        .Cells(1, 2).Value = "Prueba3"                'Nombre
        .Cells(1, 3).Value = 1                        'Orden
        .Cells(1, 4).Value = "Prueba3"                'Valor
        .Cells(1, 5).Value = 1                        'Tipo
        .Cells(1, 6).Value = "Cocepto de Prueba3"     'Concepto
        .Cells(1, 7).Value = Now                      'FechaAlta
        .Cells(1, 8).Value = Now                      'FechaModificacion
    End With
    '
    '   Select Id no existente
    '
    Set mRango = mObj.GetRowById(5, tblParametros)
    If mObj.ErrNumber <> 0 Then
        Debug.Print "#err Select NO Id: " & mObj.ErrNumber & "-" & mObj.ErrDescription & "(" _
                    & mObj.ErrProcces & ")"
    End If
    '
    '   UpdateRow
    '
    Set mRango = mObj.GetRowById(3, tblParametros)
    With mRango
        .Cells(1, 1).Value = 1                          'N
        .Cells(1, 4).Value = "Texto Actualizado " & Now  'Valor
        .Cells(1, 6).Value = "Prueba Actualización"     'Concepto
        .Cells(1, 8).Value = Now                      'FechaModificacion
    End With
    '
    '   Select por Columna
    '
    Set mRango = mObj.GetRowByColumn("Prueba3", 1, tblParametros)
    PrintRango mRango
    
    '
    '   DeleteRow
    '
    mObj.DeleteRow 1, tblParametros
    '
    '   SelectApuestasByBoleto
    '
    Dim mCol As Collection
    Dim mTup As TuplaAparicion
    Set mCol = mObj.SelectApuestasByBoleto(5, tblApuestas)
    If mObj.ErrNumber = 0 Then
        For Each mTup In mCol
            Debug.Print vbTab & "Registro: " & mTup.NumeroRegistro
        Next mTup
    Else
        Debug.Print "#(" & mObj.ErrNumber & ") - " & mObj.ErrDescription & " in " & mObj.ErrProcces
    End If
    Set mCol = mObj.SelectApuestasByBoleto(7, tblApuestas)
    If mObj.ErrNumber = 0 Then
        For Each mTup In mCol
            Debug.Print vbTab & "Registro: " & mTup.NumeroRegistro
        Next mTup
    Else
        Debug.Print "#(" & mObj.ErrNumber & ") - " & mObj.ErrDescription & " in " & mObj.ErrProcces
    End If
    '
    '   TODO: Buscar numeros en los sorteos
    '
    '
    '
    '   Cerramos el libro de datos
    '
    mObj.Cerrar
    '
    '
  On Error GoTo 0
DataBaseExcelTest2__CleanExit:
    Exit Sub
            
DataBaseExcelTest2_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    '   Audita el error
    Call HandleException(ErrNumber, ErrDescription, "Lot_PqtNucleoTesting.DataBaseExcelTest2", ErrSource)
    '   Informa del error
    Call MsgBox(ErrDescription, vbCritical Or vbSystemModal, Application.Caption)
    Call Trace("CERRAR")
End Sub

'------------------------------------------------------------------------------*
' Procedimiento  : DataBaseExcelTest
' Fecha          : 28/05/2018
' Propósito      : Pruebas unitarias de la clase DataBaseExcel Consulta
'------------------------------------------------------------------------------*
'
Public Sub DataBaseExcelTest()
    Dim mObj        As DataBaseExcel
    Dim mRango      As Range
    Dim mPer        As Periodo
    Dim mId         As Integer
    
 On Error GoTo DataBaseExcelTest_Error
    '
    '   Creamos la base de datos
    '
    Set mObj = New DataBaseExcel
    '
    '   Abrir la base de datos
    '
    mObj.Abrir
    '
    '   Get Last Id
    '
    mId = mObj.GetLastID(tblGordo)
    Debug.Print "Último Id de Gordo : " & mId
    '
    mId = mObj.GetLastID(tblPrimitiva)
    Debug.Print "Último Id de Primitiva : " & mId
    '
    mId = mObj.GetLastID(tblBonoloto)
    Debug.Print "Último Id de Bonoloto : " & mId
    '
    '   GetRowById
    '
    mId = 6794
    Set mRango = mObj.GetRowById(mId, tblBonoloto)
    '   Fecha = 22/11/2018 Bonoloto
    PrintRango mRango
    '
    mId = 3167
    Set mRango = mObj.GetRowById(mId, tblPrimitiva)
    '   Fecha = 01/11/2018 Primitiva
    PrintRango mRango
    '
    mId = 1106
    Set mRango = mObj.GetRowById(mId, tblEuromillon)
    '   Fecha = 26/10/2018 Euromillon
    PrintRango mRango
    '
    mId = 1096
    Set mRango = mObj.GetRowById(mId, tblGordo)
    '   Fecha = 04/11/2018 Gordo
    PrintRango mRango
    '
    '   No encontrado
    mId = 9999
    On Error Resume Next
    Set mRango = mObj.GetRowById(mId, tblGordo)
    Debug.Print Err.Number & " - " & Err.Description
    On Error GoTo 0
    
    '
    '   SelectByFechas
    '
    '   Creamos el periodo
    Set mPer = New Periodo
    mPer.FechaInicial = #1/5/2018#
    mPer.FechaFinal = #1/15/2018#
    Set mRango = mObj.SelectByFechas(mPer.FechaInicial, mPer.FechaFinal, tblBonoloto)
    '   Fechas (5/1/2018 al 15/01/2018 Bonoloto
    PrintRango mRango
    '
    '
    mPer.FechaInicial = #1/5/2018#
    mPer.FechaFinal = #1/16/2018#
    Set mRango = mObj.SelectByFechas(mPer.FechaInicial, mPer.FechaFinal, tblEuromillon)
    '   Fechas (5/1/2018 al 16/01/2018 Euromillon
    PrintRango mRango
    '
    '
    mPer.FechaInicial = #1/7/2018#
    mPer.FechaFinal = #1/28/2018#
    Set mRango = mObj.SelectByFechas(mPer.FechaInicial, mPer.FechaFinal, tblGordo)
    '   Fechas (7/1/2018 al 28/1/2018 Gordo Primitiva
    PrintRango mRango
    '
    '   Rango no encontrado
    mPer.FechaInicial = #10/1/2018#
    mPer.FechaFinal = #10/25/2018#
    Set mRango = mObj.SelectByFechas(mPer.FechaInicial, mPer.FechaFinal, tblPrimitiva)
    '   Fechas (1/10/2018 al 25/10/2018 Primitiva
    If mRango Is Nothing Then
        Debug.Print Err.Number & " - " & Err.Description
    Else
        PrintRango mRango
    End If
    '
    '   SelectById
    '
    Set mRango = mObj.SelectByIds(6797, 6800, tblBonoloto)
    '   Fechas (26/11/2018 al 29/11/2018 Bonoloto
    PrintRango mRango
    '
    Set mRango = mObj.SelectByIds(1109, 1112, tblEuromillon)
    '   Fechas (06/11/2018 al 16/11/2018 Euromillon
    PrintRango mRango
    '
    Set mRango = mObj.SelectByIds(1094, 1098, tblGordo)
    '   Fechas (21/10/2018 al 18/11/2018 Gordo Primitiva
    PrintRango mRango
    '
    '   Rango no encontrado
    Set mRango = mObj.SelectByIds(5690, 6008, tblPrimitiva)
    '   Fechas (registros no existen en Primitiva
    If mRango Is Nothing Then
        Debug.Print Err.Number & " - " & Err.Description
    Else
        PrintRango mRango
    End If
    
    '
    '   Cerramos el libro de datos
    '
    mObj.Cerrar
    '
    '
  On Error GoTo 0
DataBaseExcelTest__CleanExit:
    Exit Sub
            
DataBaseExcelTest_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    '   Audita el error
    Call HandleException(ErrNumber, ErrDescription, "Lot_PqtNucleoTesting.DataBaseExcelTest", ErrSource)
    '   Informa del error
    Call MsgBox(ErrDescription, vbCritical Or vbSystemModal, Application.Caption)
    Call Trace("CERRAR")
End Sub

'------------------------------------------------------------------------------*
' Procedimiento  : PrintRango
' Fecha          : 09/12/2018
' Propósito      : Imprime un rango definido
'------------------------------------------------------------------------------*
'
Private Sub PrintRango(mObj As Range)
    Dim mCelda As Range
    Debug.Print "==> Visualizar un rango de datos"
    If mObj Is Nothing Then
        Debug.Print "Rango = #Null"
        Exit Sub
    End If
    Debug.Print "Address                =" & mObj.Address
    Debug.Print "Column                 =" & mObj.Column
    Debug.Print "Count                  =" & mObj.Count
    Debug.Print "Height                 =" & mObj.Height
    Debug.Print "Left                   =" & mObj.Left
    Debug.Print "Row                    =" & mObj.Row
    For Each mCelda In mObj.Cells
        Debug.Print vbTab & "-" & mCelda.Address & "= " & mCelda.Value
    Next mCelda
End Sub
'------------------------------------------------------------------------------*
' Procedimiento  : EntidadNegocioTest
' Fecha          : 21/sep/2018
' Propósito      : Pruebas unitarias de la clase EntidadNegocio
'------------------------------------------------------------------------------*
Private Sub EntidadNegocioTest()
    Dim Obj         As EntidadNegocio
    
  On Error GoTo EntidadNegocioTest_Error
    '
    '
    Set Obj = New EntidadNegocio
    With Obj
        .ClassStorage = True
        .FechaModificacion = Now()
        .IsDirty = True
        .MarkForDelete = True
        .Origen = 1
        .Situacion = 1
    End With
    '
    '
    PrintEntidadNegocio Obj
    '
    '
    '
    Obj.FechaModificacion = Now()
    PrintEntidadNegocio Obj
    
  On Error GoTo 0
EntidadNegocioTest__CleanExit:
    Exit Sub
EntidadNegocioTest_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    '   Audita el error
    Call HandleException(ErrNumber, ErrDescription, "[File].EntidadNegocioTest", ErrSource)
    '   Informa del error
    Call MsgBox(ErrDescription, vbCritical Or vbSystemModal, ThisWorkbook.Name)
    Call Trace("CERRAR")
End Sub


' *===========(EOF): Lot_PqtNucleoTesting.bas
