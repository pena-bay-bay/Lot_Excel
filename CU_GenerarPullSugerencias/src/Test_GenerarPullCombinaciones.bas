Attribute VB_Name = "Test_GenerarPullCombinaciones"
' *============================================================================*
' *
' *     Fichero    : Test_GenerarPullCombinaciones.bas
' *
' *     Autor      : Carlos Almela Baeza
' *     Creación   : lun, 14/dic/2020 23:42:04
' *     Versión    : 1.0
' *     Propósito  :
' *
' *============================================================================*
Option Explicit
Option Base 0

'------------------------------------------------------------------------------*
' Procedimiento  : FrmProgresoTesting
' Fecha          : lu., 21/dic/2020 18:44:47
' Propósito      : Testing del formulario de progreso
'------------------------------------------------------------------------------*
Private Sub FrmProgresoTesting()
    Dim ofrm    As frmProgreso
    Dim i       As Long
 On Error GoTo FrmProgresoTesting_Error
    '
    '   Creamos el objeto proceso
    '
    Set ofrm = New frmProgreso
    '
    '   Establecemos parametros del bucle
    '
    With ofrm
        .Fase = "Generacion de combinaciones"
        .Maximo = 500000
    End With
    '
    '   Visualizamos Formulario
    '
    ofrm.Show
    '
    '   Bucle de prueba
    '
    For i = 1 To ofrm.Maximo
        ofrm.Valor = i
    Next i
    '
    '   Mostramos resumen
    '
    ofrm.DisProceso
    '
    '   Destruimos elobjeto
    '
    Set ofrm = Nothing
    
  On Error GoTo 0
FrmProgresoTesting__CleanExit:
    Exit Sub
            
FrmProgresoTesting_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    Call HandleException(ErrNumber, ErrDescription, "Testing.FrmProgresoTesting", ErrSource)
    Call MsgBox(ErrDescription, vbCritical Or vbSystemModal, ThisWorkbook.Name)
    Call Trace("CERRAR")
End Sub




'------------------------------------------------------------------------------*
' Procedimiento  : GenPullCombinacionesViewTesting
' Fecha          : lu., 21/dic/2020 18:44:47
' Propósito      : Testing de la vista del proceso
'------------------------------------------------------------------------------*
Private Sub GenPullCombinacionesViewTesting()
    Dim mView       As GenPullCombinacionesView
    Dim mModel      As GenPullCombinacionesModel
    Dim mArray      As Variant
    
  On Error GoTo GenPullCombinacionesViewTesting_Error
    '
    '   Creamos el Objeto
    '
    Set mView = New GenPullCombinacionesView
    '
    '   Visualizamos propiedades
    '
    Print_GenPullCombinacionesView mView
    '
    '   Caso de prueba 1: Metodo BorrarFiltros
    '
    Debug.Print " Método => BorrarFiltros()"
    mView.BorrarFiltros
    If mView.TotalFiltros <> 0 Then
        Debug.Print "   #Error Borrar filtros no  es 0 =>" & CInt(mView.TotalFiltros)
    Else
        Debug.Print "   Prueba BorrarFiltros correcta"
    End If
    '
    '   Caso de prueba 2: Metodo AgregarFiltro
    '
    Debug.Print " Método => AgregarFiltro()"
    mView.AgregarFiltro
    If mView.TotalFiltros <> 1 Then
        Debug.Print "   #Error Agregar filtros :" & CInt(mView.TotalFiltros)
    Else
        Debug.Print "   Prueba AgregarFiltro correcta"
    End If
    '
    '   Caso de prueba 3: Metodo ClearSalida
    '
    Debug.Print " Método => ClearSalida()"
    mView.ClearSalida
    If mView.TotalCombinaciones <> 0 Then
        Debug.Print "   #Error Clear Salida"
    Else
        Debug.Print "   Prueba ClearSalida correcta"
    End If
    '
    '   Caso de prueba 4: Metodo ClearSalidaEvaluacion
    '
    Debug.Print " Método => ClearSalidaEvaluacion()"
    mView.ClearSalidaEvaluacion
    If mView.CombinacionesEvaluadas <> 0 Then
        Debug.Print "   #Error Clear ClearSalidaEvaluacion"
    Else
        Debug.Print "   Prueba ClearSalidaEvaluacion correcta"
    End If
    '
    '   Caso de prueba 5: Metodo ClearSalidaFiltros
    '
    Debug.Print " Método => ClearSalidaFiltros()"
    mView.ClearSalidaFiltros
    If mView.CombinacionesFiltradas <> 0 Then
        Debug.Print "   #Error Clear ClearSalidaFiltros"
    Else
        Debug.Print "   Prueba ClearSalidaFiltros correcta"
    End If
    '
    '   Caso de prueba 6: Metodo ClearSalidaFiltros
    '
    Debug.Print " Método => GetParametrosProceso()"
    Set mModel = mView.GetParametrosProceso()
    If mModel.IsValid() Then
        Debug.Print "   Prueba GetParametrosProceso correcta"
    Else
        Debug.Print "   #Error GetParametrosProceso"
    End If
    '
    '   Caso de prueba 7: Metodo GetSorteos
    '
    Debug.Print " Método => GetSorteos()"
    mArray = mView.GetSorteos()
    If UBound(mArray) > 0 Then
        Debug.Print "   Prueba GetSorteos correcta"
    Else
        Debug.Print "   #Error GetSorteos"
    End If
    '
    '   Caso de prueba 8: Metodo SetFiltros
    '
    Debug.Print " Método => SetFiltros()"
    mArray = Array("6/0", "5/1", "4/2", "3/3", "2/4", "1/5", "0/6")
    mView.SetFiltros mArray
    If IsEmpty(mView.ValorFiltro) Then
        Debug.Print "   Prueba SetFiltros correcta"
    Else
        Debug.Print "   #Error SetFiltros"
    End If
  
  On Error GoTo 0
GenPullCombinacionesViewTesting__CleanExit:
    Exit Sub
GenPullCombinacionesViewTesting_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    Call HandleException(ErrNumber, ErrDescription, "Testing.GenPullCombinacionesViewTesting", ErrSource)
    Call MsgBox(ErrDescription, vbCritical Or vbSystemModal, ThisWorkbook.Name)
    Call Trace("CERRAR")
End Sub





'------------------------------------------------------------------------------*
' Procedimiento  : GenPullCombinacionesModelTesting
' Fecha          : ju., 04/feb/2021 16:29:35
' Propósito      : Testing del modelo del proceso
'------------------------------------------------------------------------------*
Private Sub GenPullCombinacionesModelTesting()
    Dim mModel      As GenPullCombinacionesModel
  On Error GoTo GenPullCombinacionesViewTesting_Error
    '
    '   Creamos el Objeto
    '
    Set mModel = New GenPullCombinacionesModel
    '
    '   Visualizamos propiedades
    '
    Print_GenPullCombinacionesModel mModel
    '
    '   Caso de prueba 1: Metodo BorrarFiltros
    '
    Debug.Print " Método ComprobarCombinaciones => ()"
    Debug.Print " Método EvaluarCombinaciones => ()"
    Debug.Print " Método FiltrarCombinaciones => ()"
    Debug.Print " Método GenerarCombinaciones => ()"
    Debug.Print " GetFiltrosOf => ()"
    Debug.Print " GetMessage => ()"
    Debug.Print " IsValid => ()"

    
'        mModel.ComprobarCombinaciones
'        mModel.EvaluarCombinaciones
'        mModel.FiltrarCombinaciones
'        mModel.GenerarCombinaciones
'        mModel.GetFiltrosOf
'        mModel.GetMessage
'        mModel.IsValid
    
  On Error GoTo 0
GenPullCombinacionesViewTesting__CleanExit:
    Exit Sub
GenPullCombinacionesViewTesting_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    Call HandleException(ErrNumber, ErrDescription, "Testing.GenPullCombinacionesModelTesting", ErrSource)
    Call MsgBox(ErrDescription, vbCritical Or vbSystemModal, ThisWorkbook.Name)
    Call Trace("CERRAR")
End Sub


'---------------------------------------------------------------------------------------
' Procedure : Print_GenPullCombinacionesView
' Author    : Charly
' Date      : ju., 28/ene/2021 20:06:32
' Purpose   : Visualiza las propiedades de la clase GenPullCombinacionesView
'---------------------------------------------------------------------------------------
'
Private Sub Print_GenPullCombinacionesView(obj As GenPullCombinacionesView)
    Debug.Print "==> Pruebas GenPullCombinacionesViewTest"
    Debug.Print vbTab & "CombinacionesEvaluadas         =" & obj.CombinacionesEvaluadas
    Debug.Print vbTab & "CombinacionesFiltradas         =" & obj.CombinacionesFiltradas
    Debug.Print vbTab & "CombinacionesGeneradas         =" & obj.CombinacionesGeneradas
    Debug.Print vbTab & "NumSugerencias                 =" & obj.NumSugerencias
    Debug.Print vbTab & "RangoSalida.Address            =" & obj.RangoSalida.Address
    Debug.Print vbTab & "RangoSalidaComprobacion.Address=" & obj.RangoSalidaComprobacion.Address
    Debug.Print vbTab & "RangoSalidaEvaluadas.Address   =" & obj.RangoSalidaEvaluadas.Address
    Debug.Print vbTab & "RangoSalidaFiltros.Address     =" & obj.RangoSalidaFiltros.Address
    Debug.Print vbTab & "RegistrosComprobados           =" & obj.RegistrosComprobados
    Debug.Print vbTab & "RegistrosPremiados             =" & obj.RegistrosPremiados
    Debug.Print vbTab & "TipoFiltro                     =" & obj.TipoFiltro
    Debug.Print vbTab & "TotalCombinaciones             =" & obj.TotalCombinaciones
    Debug.Print vbTab & "TotalNumeros                   =" & obj.TotalNumeros
    Debug.Print vbTab & "TotalFiltros                   =" & obj.TotalFiltros
    Debug.Print vbTab & "TotalCoste                     =" & obj.TotalCoste
    Debug.Print vbTab & "TotalImporte                   =" & obj.TotalImporte
    Debug.Print vbTab & "ValorFiltro                    =" & obj.ValorFiltro
End Sub




'---------------------------------------------------------------------------------------
' Procedure : Print_GenPullCombinacionesModel
' Author    : Charly
' Date      : ju., 04/feb/2021 16:27:21
' Purpose   : Visualiza las propiedades de la clase GenPullCombinacionesModel
'---------------------------------------------------------------------------------------
'
Private Sub Print_GenPullCombinacionesModel(obj As GenPullCombinacionesModel)
    Debug.Print "==> Pruebas Print_GenPullCombinacionesModel"
    Debug.Print vbTab & "CombinacionGanadora         =" & obj.CombinacionGanadora
    Debug.Print vbTab & "FaseProceso                 =" & obj.FaseProceso
    Debug.Print vbTab & "Filtros                     =" & obj.Filtros
    Debug.Print vbTab & "MatrizNumeros               =" & obj.MatrizNumeros
    Debug.Print vbTab & "NumerosSugerencia           =" & obj.NumerosSugerencia
    Debug.Print vbTab & "Sorteos                     =" & obj.Sorteos
    Debug.Print vbTab & "TotalCombinaciones          =" & obj.TotalCombinaciones
    Debug.Print vbTab & "TotalNumerosCombinar        =" & obj.TotalNumerosCombinar
End Sub
'' *===========(EOF): Test_GenerarPullCombinaciones.bas
