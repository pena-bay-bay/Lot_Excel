Attribute VB_Name = "Lot_12_TestMetodo01"
' *============================================================================*
' *
' *     Fichero    : Lot_12_TestMetodo01.bas
' *
' *     Autor      : Carlos Almela Baeza
' *     Creación   : mie, 04/nov/2020 20:02:45
' *     Versión    : 1.0
' *     Propósito  : Probar un método de estadistica fijando una fecha
' *                  Para una visualización del método en el tiempo se podría
' *                  ejecutar la estadística desde una fecha hasta el número
' *                  de días que se quiere comprobar y ver como el método
' *                  acierta en el tiempo. Por ejemplo fijamos una fecha
' *                  (01/08/2020) y establecemos un método:
' *                  Probabilidad 45 días descendente, y una apuesta de
' *                  8 números; pues en un bucle desde la fecha + 7 días
' *                  para tener estadística se calcula una sugerencia cada
' *                  día y se comprueban los aciertos a lo largo del tiempo.
' *                  Empezaríamos 7 días más tarde del 1 de ago. (8-ago, si
' *                  hay sorteo sino uno más) calculamos la estadística desde
' *                  el 1-ago al 7-ago y deducimos una sugerencia de 8 números
' *                  y comprobamos con el resultado del día 8-ago, así hasta
' *                  llegar a una estadística de 45 registros, y revisar
' *                  aciertos en el tiempo.
' *
' *============================================================================*
Option Explicit
Option Base 0

'--- Variables Privadas -------------------------------------------------------*
Private NumSorteosPremiados         As Integer
Private FechaFinProceso             As Date
Private mInfo                       As InfoSorteo
Private mDb                         As BdDatos
Private mMetodo                     As Metodo
Private mParamMuestra               As ParametrosMuestra
Private mSorteo                     As Sorteo
Private mMuestra                    As Muestra
Private mCalSuge                    As RealizarSugerencia
Private mSuge                       As Sugerencia

'--- Constantes ---------------------------------------------------------------*
                                    ' Fecha 01/08/2020
                                    ' Fecha 14/09/2020
                                    ' Fecha 19/11/2020
Private Const FECHA_INICIO          As Date = #11/19/2020#
Private Const REGISTROS_OFFSET      As Integer = 7
Private Const COLOR_STDBOLAS        As Integer = ordProbabilidad
Private Const NUM_PRONOSTICOS       As Integer = 6


'--- Mensajes -----------------------------------------------------------------*
'--- Errores ------------------------------------------------------------------*
'--- Métodos Privados ---------------------------------------------------------*


'------------------------------------------------------------------------------*
' Procedimiento  : GetMuestraProceso
' Fecha          : 05/11/2020
' Propósito      : Obtenemos la muestra del periodo de análisis
'------------------------------------------------------------------------------*
Private Function GetMuestraProceso(NewData As ParametrosMuestra) As Muestra
    Dim objMuestra As Muestra
    Dim m_objRg As Range
    
  On Error GoTo GetMuestraProceso_Error
    Set objMuestra = New Muestra
    '
    '       Calcula la Muestra
    '
    '   obtiene el rango con los datos comprendido entre las dos fechas
    '
    Set m_objRg = mDb.GetSorteosInFechas(NewData.PeriodoDatos)
    '
    '   se lo pasa al constructor de la clase y obtiene las estadisticas para cada bola
    '
    Set objMuestra.ParametrosMuestra = NewData
    Select Case JUEGO_DEFECTO
        Case LoteriaPrimitiva, Bonoloto:
            objMuestra.Constructor m_objRg, ModalidadJuego.LP_LB_6_49
        
        Case GordoPrimitiva:
            objMuestra.Constructor m_objRg, ModalidadJuego.GP_5_54
        
        Case Euromillones:
            objMuestra.Constructor m_objRg, ModalidadJuego.EU_5_50
    End Select
    '
    '   Devolvemos la muestra
    '
    Set GetMuestraProceso = objMuestra
    
  On Error GoTo 0
GetMuestraProceso__CleanExit:
    Exit Function
            
GetMuestraProceso_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    Call HandleException(ErrNumber, ErrDescription, _
                "Lot_12_TestMetodo01.GetMuestraProceso", ErrSource)
    Err.Raise ErrNumber, "Lot_12_TestMetodo01.GetMuestraProceso", ErrDescription
End Function



'------------------------------------------------------------------------------*
' Procedimiento  : GetParametrosMuestra
' Fecha          : 04/11/2020
' Propósito      : Devuelve los parámetros de la  muestra
'------------------------------------------------------------------------------*
Private Function GetParametrosMuestra(NewData As Date)
    Dim mObj As ParametrosMuestra
    
  On Error GoTo GetParametrosMuestra_Error
    Set mObj = New ParametrosMuestra
    With mObj
        .TipoMuestra = False   ' Por días
        .FechaAnalisis = NewData
        .FechaFinal = mInfo.GetAnteriorSorteo(NewData)
        .FechaInicial = FECHA_INICIO
    End With
    Set GetParametrosMuestra = mObj
  
  On Error GoTo 0
GetParametrosMuestra__CleanExit:
    Exit Function
GetParametrosMuestra_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    Call HandleException(ErrNumber, ErrDescription, _
                "Lot_12_TestMetodo01.GetParametrosMuestra", ErrSource)
    Err.Raise ErrNumber, "Lot_12_TestMetodo01.GetParametrosMuestra", ErrDescription
End Function




'------------------------------------------------------------------------------*
' Procedimiento  : VisualizarLiterales
' Fecha          : 04/11/2020
' Propósito      : Prepara la hoja de salida para el proceso
'------------------------------------------------------------------------------*
Private Sub VisualizarLiterales()
    Dim i As Integer
    Dim mVar As Variant

  On Error GoTo VisualizarLiterales_Error
    '
    '   Borramos la hoja de salida
    '
    Borra_Salida
    '
    '   Literales del proceso
    '
    Range("A1").Activate
    ActiveCell.Value = "Calculo de la Fecha Fija Optima de Muestra"
    ActiveCell.Font.Bold = True
    '
    '   Literales verticales
    '       Fecha de inicio 01/09/2020
    '       Offset          7
    '       Fecha de Fin    04/11/2020
    '       Metodo prueba
    '
    mVar = Split("Fecha de Inicio;Offset;Fecha de Fin;Método Sugerencia;Sorteos Acertados", ";")
    For i = 0 To UBound(mVar)
        ActiveCell.Offset(i + 2, 0).Value = mVar(i)
        ActiveCell.Offset(i + 2, 0).Font.Bold = True
    Next i
    '
    '       Literales horizontales
    '
    Range("D2").Activate
    '
    '
    '
    Select Case JUEGO_DEFECTO
        Case Bonoloto, LoteriaPrimitiva:
            mVar = Split("N;Fecha;N1;N2;N3;N4;N5;N6;C;Sugerencia;Aciertos;Importe", ";")
        Case GordoPrimitiva:
            mVar = Split("N;Fecha;N1;N2;N3;N4;N5;R;Sugerencia;Aciertos;Importe", ";")
        Case Euromillones:
            mVar = Split("N;Fecha;N1;N2;N3;N4;N5;E1;E2;Sugerencia;Aciertos;Importe", ";")
    End Select
    For i = 0 To UBound(mVar)
        ActiveCell.Offset(0, i).Value = mVar(i)
        ActiveCell.Offset(0, i).Font.Bold = True
    Next i
    
  On Error GoTo 0
VisualizarLiterales__CleanExit:
    Exit Sub
            
VisualizarLiterales_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    Call HandleException(ErrNumber, ErrDescription, _
                "Lot_12_TestMetodo01.VisualizarLiterales", ErrSource)
    Err.Raise ErrNumber, "Lot_12_TestMetodo01.VisualizarLiterales", ErrDescription
End Sub



'------------------------------------------------------------------------------*
' Procedimiento  : VisualizarParametros
' Fecha          : 06/11/2020
' Propósito      : Visualiza los parametros del proceso
'------------------------------------------------------------------------------*
Private Sub VisualizarParametros()
    Range("B3").Activate
    ActiveCell.Offset(0, 0).Value = FECHA_INICIO
    ActiveCell.Offset(1, 0).Value = REGISTROS_OFFSET
    ActiveCell.Offset(2, 0).Value = FechaFinProceso
    ActiveCell.Offset(3, 0).Value = mMetodo.ToString()
End Sub




'------------------------------------------------------------------------------*
' Procedimiento  : VisualizaSorteo
' Fecha          : 06/11/2020
' Propósito      : Visualiza el resultado del sorteo
'------------------------------------------------------------------------------*
Private Sub VisualizaSorteo(NewSorteo As Sorteo, NewMuestra As Muestra, _
                            nRow As Integer, NewSuge As Sugerencia, _
                            NewAciertos As String, NewImporte As Currency)
    Dim oNum As Numero
    Dim oBola As Bola
    Dim j As Integer
    Dim mColorIndex As Integer
  
  On Error GoTo VisualizaSorteo_Error
    Range("D3").Activate
    
    ActiveCell.Offset(nRow, 0).Value = nRow + 1
    ActiveCell.Offset(nRow, 1).Value = NewSorteo.Fecha
    Select Case JUEGO_DEFECTO
        Case Bonoloto, LoteriaPrimitiva:
            For j = 1 To 6
                Set oNum = NewSorteo.Combinacion.Numeros(j)
                Set oBola = NewMuestra.Get_Bola(oNum.Valor)
                Select Case COLOR_STDBOLAS
                    Case ordProbabilidad: mColorIndex = oBola.Color_Probabilidad
                    Case ordProbTiempoMedio: mColorIndex = oBola.Color_Tiempo_Medio
                    Case ordFrecuencia: mColorIndex = oBola.Color_Frecuencias
                End Select
                With ActiveCell.Offset(nRow, j + 1)
                    .Value = oNum.Valor
                    .NumberFormat = "00"
                    .Interior.ColorIndex = mColorIndex
                End With
            Next j
            oNum.Valor = NewSorteo.Complementario
            Set oBola = NewMuestra.Get_Bola(oNum.Valor)
            Select Case COLOR_STDBOLAS
                Case ordProbabilidad: mColorIndex = oBola.Color_Probabilidad
                Case ordProbTiempoMedio: mColorIndex = oBola.Color_Tiempo_Medio
                Case ordFrecuencia: mColorIndex = oBola.Color_Frecuencias
            End Select
            With ActiveCell.Offset(nRow, j + 1)
                .Value = oNum.Valor
                .NumberFormat = "00"
                .Interior.ColorIndex = mColorIndex
            End With
            
        Case Euromillones:
            For j = 1 To 5
                Set oNum = NewSorteo.Combinacion.Numeros(j)
                Set oBola = NewMuestra.Get_Bola(oNum.Valor)
                Select Case COLOR_STDBOLAS
                    Case ordProbabilidad: mColorIndex = oBola.Color_Probabilidad
                    Case ordProbTiempoMedio: mColorIndex = oBola.Color_Tiempo_Medio
                    Case ordFrecuencia: mColorIndex = oBola.Color_Frecuencias
                End Select
                With ActiveCell.Offset(nRow, j + 1)
                    .Value = oNum.Valor
                    .NumberFormat = "00"
                    .Interior.ColorIndex = mColorIndex
                End With
            Next j
        Case GordoPrimitiva:
            For j = 1 To 5
                Set oNum = NewSorteo.Combinacion.Numeros(j)
                Set oBola = NewMuestra.Get_Bola(oNum.Valor)
                Select Case COLOR_STDBOLAS
                    Case ordProbabilidad: mColorIndex = oBola.Color_Probabilidad
                    Case ordProbTiempoMedio: mColorIndex = oBola.Color_Tiempo_Medio
                    Case ordFrecuencia: mColorIndex = oBola.Color_Frecuencias
                End Select
                With ActiveCell.Offset(nRow, j + 1)
                    .Value = oNum.Valor
                    .NumberFormat = "00"
                    .Interior.ColorIndex = mColorIndex
                End With
            Next j
    
    End Select
    j = j + 2
    '
    '
    '
    ActiveCell.Offset(nRow, j).Value = NewSuge.ToString()
    ActiveCell.Offset(nRow, j + 1).Value = NewAciertos
    ActiveCell.Offset(nRow, j + 2).Value = NewImporte
    
    
On Error GoTo 0
VisualizaSorteo__CleanExit:
    Exit Sub
            
VisualizaSorteo_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    Call HandleException(ErrNumber, ErrDescription, _
                "Lot_12_TestMetodo01.VisualizaSorteo", ErrSource)
    Err.Raise ErrNumber, "Lot_12_TestMetodo01.VisualizaSorteo", ErrDescription
End Sub



'--- Métodos Públicos ---------------------------------------------------------*
'------------------------------------------------------------------------------*
' Procedimiento  : CalFechaOptima
' Fecha          : 04/11/2020
' Propósito      : Calcular la fecha fija óptima para la estadistica fija
'------------------------------------------------------------------------------*
Public Sub CalFechaOptima()
    Dim mFecIni As Date
    Dim mFecFin As Date
    Dim mRango As Range
    Dim mPeriodo As Periodo
    Dim mFila As Range
    Dim mRow As Integer
    Dim mCheck      As ComprobarSugerencia
    Dim mAciertos   As String
    Dim mImporteAciertos As Currency
    
  On Error GoTo CalFechaOptima_Error
    '
    '   Creamos los objetos del proceso
    '
    Set mInfo = New InfoSorteo
    Set mDb = New BdDatos
    Set mPeriodo = New Periodo
    Set mSorteo = New Sorteo
    Set mMetodo = New Metodo
    Set mParamMuestra = New ParametrosMuestra
    Set mSuge = New Sugerencia
    Set mCalSuge = New RealizarSugerencia
    '
    '   Desactiva la presentación
    '
    CALCULOOFF
    '
    '   Creamos el comprobador de Sugerencias
    '
    Set mCheck = New ComprobarSugerencia
    '
    '   Visualiza los literales del proceso
    '
    VisualizarLiterales
    '
    '   Calculamos el periodo de tratamiento
    '
    mFecIni = FECHA_INICIO + REGISTROS_OFFSET
    FechaFinProceso = mDb.UltimoResultado
    If Not mInfo.EsFechaSorteo(mFecIni) Then
        mFecIni = mInfo.GetAnteriorSorteo(mFecIni)
    End If
    mPeriodo.FechaInicial = mFecIni
    mPeriodo.FechaFinal = FechaFinProceso
    '
    '   Definimos el metodo de sugerencia
    '
    With mMetodo
        .TipoProcedimiento = mtdEstadistico
        .CriteriosOrdenacion = ordProbabilidad
        .SentidoOrdenacion = False
        .TipoMuestra = False
        .DiasAnalisis = 0
        .ModalidadJuego = LP_LB_6_49
        .Pronosticos = NUM_PRONOSTICOS
    End With
    '
    '
    '
    VisualizarParametros
    '
    '   Obtenemos el rango de análisis
    '
    Set mRango = mDb.GetSorteosInFechas(mPeriodo)
    '
    '
    '
    mRow = 0
    For Each mFila In mRango.Rows
        '
        '   Establecemos el sorteo
        '
        mSorteo.Constructor mFila
        '
        '   Obtenemos la muestra para este sorteo como Fecha Analisis
        '
        Set mParamMuestra = GetParametrosMuestra(mSorteo.Fecha)
        '
        '   Calculamos parametros muestra
        '
        Set mMuestra = GetMuestraProceso(mParamMuestra)
        '
        '   Obtenemos la sugerencia para esta fecha
        '
        mMetodo.DiasAnalisis = mParamMuestra.DiasAnalisis
        Set mSuge = mCalSuge.GetSugerencia(mMetodo, mSorteo.Fecha)
        '
        '   Verificamos la sugerencia con el sorteo
        '
        Set mCheck.Sorteo = mSorteo
        mCheck.ComprobarSugerencia mSuge
        If mCheck.NumerosAcertados > 0 Then
            If mCheck.EstaPremiada Then
                mAciertos = mCheck.CategoriaPremioTxt
                mImporteAciertos = mCheck.ImporteApuesta
            Else
                mAciertos = mCheck.NumerosAcertados
                mImporteAciertos = 0
            End If
        Else
            mAciertos = ""
            mImporteAciertos = 0
        End If
        '
        '
        '   Visualizamos Sorteo, sugerencia y aciertos
        '
        VisualizaSorteo mSorteo, mMuestra, mRow, mSuge, mAciertos, mImporteAciertos
        '
        '   Totalizamos sorteos
        '
        mRow = mRow + 1
    Next mFila
    
    Cells.Select                            'Selecciona todas las celdas de la hoja
    Cells.EntireColumn.AutoFit              'Autoajusta el tamaño de las columnas
    '
    '  Activa la presentación
    '
    CALCULOON
  On Error GoTo 0
CalFechaOptima__CleanExit:
    Exit Sub
    
CalFechaOptima_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    Call HandleException(ErrNumber, ErrDescription, "Lot_12_TestMetodo01.CalFechaOptima", ErrSource)
    Call MsgBox(ErrDescription, vbCritical Or vbSystemModal, ThisWorkbook.Name)
    Call Trace("CERRAR")
End Sub

' *===========(EOF): Lot_12_TestMetodo01.bas
