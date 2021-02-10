Attribute VB_Name = "Lot_12_TestMetodo01"
' *============================================================================*
' *
' *     Fichero    : Lot_12_TestMetodo01.bas
' *
' *     Autor      : Carlos Almela Baeza
' *     Creaci�n   : mie, 04/nov/2020 20:02:45
' *     Versi�n    : 1.0
' *     Prop�sito  : Probar un m�todo de estadistica fijando una fecha
' *                  Para una visualizaci�n del m�todo en el tiempo se podr�a
' *                  ejecutar la estad�stica desde una fecha hasta el n�mero
' *                  de d�as que se quiere comprobar y ver como el m�todo
' *                  acierta en el tiempo. Por ejemplo fijamos una fecha
' *                  (01/08/2020) y establecemos un m�todo:
' *                  Probabilidad 45 d�as descendente, y una apuesta de
' *                  8 n�meros; pues en un bucle desde la fecha + 7 d�as
' *                  para tener estad�stica se calcula una sugerencia cada
' *                  d�a y se comprueban los aciertos a lo largo del tiempo.
' *                  Empezar�amos 7 d�as m�s tarde del 1 de ago. (8-ago, si
' *                  hay sorteo sino uno m�s) calculamos la estad�stica desde
' *                  el 1-ago al 7-ago y deducimos una sugerencia de 8 n�meros
' *                  y comprobamos con el resultado del d�a 8-ago, as� hasta
' *                  llegar a una estad�stica de 45 registros, y revisar
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


'--- Constantes ---------------------------------------------------------------*
Private Const FECHA_INICIO          As Date = #9/1/2020#
Private Const REGISTROS_OFFSET      As Integer = 7

'--- Mensajes -----------------------------------------------------------------*
'--- Errores ------------------------------------------------------------------*
'--- M�todos Privados ---------------------------------------------------------*



'------------------------------------------------------------------------------*
' Procedimiento  : GetMuestraProceso
' Fecha          : 05/11/2020
' Prop�sito      : Obtenemos la muestra del periodo de an�lisis
'------------------------------------------------------------------------------*
Private Function GetMuestraProceso(NewData As ParametrosMuestra) As ParametrosMuestra
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
    Exit Sub
            
GetMuestraProceso_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    Call HandleException(ErrNumber, ErrDescription, _
                "Lot_12_TestMetodo01.GetMuestraProceso", ErrSource)
    Err.Raise ErrNumber, "Lot_12_TestMetodo01.GetMuestraProceso", ErrDescription
End Function

'------------------------------------------------------------------------------*
' Procedimiento  : VisualizarLiterales
' Fecha          : 04/11/2020
' Prop�sito      : Prepara la hoja de salida para el proceso
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
    mVar = Split("Fecha de Inicio;Offset;Fecha de Fin;M�todo Sugerencia;Sorteos Acertados", ";")
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




'--- M�todos P�blicos ---------------------------------------------------------*
'------------------------------------------------------------------------------*
' Procedimiento  : CalFechaOptima
' Fecha          : 04/11/2020
' Prop�sito      : Calcular la fecha fija �ptima para la estadistica fija
'------------------------------------------------------------------------------*
Public Sub CalFechaOptima()
    Dim mFecIni As Date
    Dim mFecFin As Date
    Dim mRango As Range
    Dim mPeriodo As Periodo
    Dim mFila As Range
    
  On Error GoTo CalFechaOptima_Error
    '
    '   Creamos los objetos del proceso
    '
    Set mInfo = New InfoSorteo
    Set mDb = New BdDatos
    Set mPeriodo = New Periodo
    Set mSorteo = New Sorteo
    '
    '   Desactiva la presentaci�n
    '
    CALCULOOFF
    '
    '   Visualiza los literales del proceso
    '
    VisualizarLiterales
    '
    '   Calculamos los parametros de la muestra
    '
    mFecIni = FECHA_INICIO
    mFecFin = FECHA_INICIO + REGISTROS_OFFSET
    FechaFinProceso = mDb.UltimoResultado
    
    If Not mInfo.EsFechaSorteo(mFecIni) Then
        mFecIni = mInfo.GetAnteriorSorteo(mFecIni)
    End If
    If Not mInfo.EsFechaSorteo(mFecFin) Then
        mFecFin = mInfo.GetProximoSorteo(mFecFin)
    End If
    
    Set mParamMuestra = New ParametrosMuestra
    With mParamMuestra
        .FechaAnalisis = mFecFin
        .FechaFinal = mInfo.GetAnteriorSorteo(mFecFin)
        .FechaInicial = mFecIni
        .TipoMuestra = True
    End With
    mPeriodo.FechaInicial = mFecIni
    mPeriodo.FechaFinal = FechaFinProceso
    '
    '   Obtenemos el rango de an�lisis
    '
    Set mRango = mDb.GetApuestaInFechas(mPeriodo)
    '
    '
    '
    For Each mFila In mRango.Rows
        '
        '   Establecemos el sorteo
        '
        mSorteo.Constructor mFila
        '
        '   Calculamos parametros muestra
        '
        Set mMuestra = GetMuestraProceso(mParamMuestra)
        '
        '
        '
    
        ' Visualiza Sorteo
    
    
    
    Next mFila
    '
    '
    '  Activa la presentaci�n
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
