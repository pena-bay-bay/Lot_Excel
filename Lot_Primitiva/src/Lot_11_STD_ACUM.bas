Attribute VB_Name = "Lot_11_STD_ACUM"
' *============================================================================*
' *
' *     Fichero    : Lot_11_STD_ACUM.bas
' *
' *     Autor      : Carlos Almela Baeza
' *     Creación   : sá., 09/nov/2019 12:16:39
' *     Versión    : 1.0
' *     Propósito  : Calcular los datos estadisticos de un periodo de sorteos
' *
' *============================================================================*
Option Explicit
Option Base 0

'--- Variables Privadas -------------------------------------------------------*
Public mRow As Integer      ' Variable global de la fila activa

'------------------------------------------------------------------------------*
' Procedimiento  : CalcularStdAcum
' Fecha          : sá., 09/nov/2019 12:20:58
' Propósito      : Obtener las estadisticas de un periodo y los aciertos
'------------------------------------------------------------------------------*
'
Public Sub CalcularStdAcum()
    Dim mFecDesde       As Date
    Dim mFecHasta       As Date
    Dim mFec            As Date
    Dim mSorteos        As Integer
    Dim oParMuestra     As ParametrosMuestra
    Dim oMuestra        As Muestra
    Dim oInfo           As InfoSorteo
    Dim oSorteo         As Sorteo
    Dim oEngSort        As SorteoEngine
    Dim rgDatos         As Range
    Dim mDB             As New BdDatos
    
  On Error GoTo CalcularStdAcum_Error
    '
    '   Definir parametros de la rutina
    '
    'mFecDesde = #11/23/2019#
    mFecHasta = #3/7/2020#
    mSorteos = 90
    mFecDesde = mFecHasta - ((90 / 2) * 7)
    
    Set oInfo = New InfoSorteo
    Set oEngSort = New SorteoEngine
    oInfo.Constructor LoteriaPrimitiva
    mFecDesde = oInfo.GetAnteriorSorteo(mFecDesde)
    '
    '   Borra la hoja de salida
    '
    Borra_Salida
    Application.ScreenUpdating = False                          'Desactiva el reflejo de pantalla
    '
    '   escribir cabeceras
    '
    PonCabCalcularStdAcum
    '
    '   Crear parametros de la muestra
    '
    Set oParMuestra = New ParametrosMuestra
    With oParMuestra
        .Juego = LoteriaPrimitiva
        .FechaAnalisis = mFecDesde
        .FechaFinal = oInfo.GetAnteriorSorteo(mFecDesde)
        .NumeroSorteos = mSorteos
    End With
    '
    '   Establecemos la fila
    '
    mRow = 1
    '
    '   Bucle de fechas desde hasta
    '
    For mFec = mFecDesde To mFecHasta
        '
        '   Si la fecha no es de un sorteo saltamos a la siguiente
        '
        If oInfo.EsFechaSorteo(mFec) Then
            '
            '   Actualizamos parametros muestra con la nueva fecha
            '
            oParMuestra.FechaAnalisis = mFec
            oParMuestra.FechaFinal = oInfo.GetAnteriorSorteo(mFec)
            '
            '   Obtenemos la estadistica para la fecha
            '
            Set rgDatos = mDB.Resultados_Fechas(oParMuestra.FechaInicial, _
                                                oParMuestra.FechaFinal)
            '
            '   se lo pasa al constructor de la clase y obtiene las estadisticas para cada bola
            '
            Set oMuestra = New Muestra
            Set oMuestra.ParametrosMuestra = oParMuestra
            oMuestra.Constructor rgDatos, LoteriaPrimitiva
            '
            '   Obtenemos el sorteo de la fecha
            '
            Set oSorteo = oEngSort.GetSorteoByFecha(mFec)
            '
            '   Visualizamos la muestra
            '
            If Not (oSorteo Is Nothing) Then
                DisMuestraFecha oMuestra, oSorteo
            End If
        End If
    Next mFec
    '
    '
    '
    Cells.Select                'Selecciona todas las celdas de la hoja
    Cells.EntireColumn.AutoFit  'Autoajusta el tamaño de las columnas
    
    Range("A1").Select          'Se posiciona en la celda del primer número
    Selection.AutoFilter        'Crea un autofiltro
    Application.ScreenUpdating = True
   
  On Error GoTo 0
    Exit Sub
CalcularStdAcum_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    Call HandleException(ErrNumber, ErrDescription, "Lot_11_STD_ACUM.CalcularStdAcum")
    Call MsgBox(ErrDescription, vbError Or vbSystemModal, NOMBRE_APLICACION)
    Call Trace("CERRAR")
End Sub
'------------------------------------------------------------------------------*
' Procedimiento  : PonCabCalcularStdAcum
' Fecha          : sá., 09/nov/2019 12:19:01
' Propósito      : Escribir en la hoja de salida la cabecera de los datos
'------------------------------------------------------------------------------*
'
Private Sub PonCabCalcularStdAcum()
    Dim mStr As String
    Dim mMtr As Variant
    Dim i As Integer
  
  On Error GoTo PonCabCalcularStdAcum_Error
    mStr = "Id;Fecha;Numero;Apariciones;Ausencias;Prob;Prob Tiempo;" _
         & "Prob Frecuencias;Tiempo;Desv;Moda;Max;Min;Terminación;" _
         & "Decena;Paridad;Peso;C.Ausencias;Acierto"
    mMtr = Split(mStr, ";")
    Range("A1").Activate
    For i = 0 To UBound(mMtr)
        With ActiveCell.Offset(0, i)
            .Value = mMtr(i)
            .Font.Bold = True
        End With
    Next i
  
  On Error GoTo 0
    Exit Sub
PonCabCalcularStdAcum_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    Call HandleException(ErrNumber, ErrDescription, "Lot_11_STD_ACUM.PonCabCalcularStdAcum")
    Err.Raise ErrNumber, "Lot_11_STD_ACUM.PonCabCalcularStdAcum", ErrDescription
End Sub

'------------------------------------------------------------------------------*
' Procedimiento  : DisMuestraFecha
' Fecha          : sá., 09/nov/2019 12:19:47
' Propósito      : Visualizar los datos de la estadistica en las celdas
' Parámetros     : Muestra estadistica y sorteo del dia
'------------------------------------------------------------------------------*
'
Private Sub DisMuestraFecha(datMuestra As Muestra, datSorteo As Sorteo)
    Dim i           As Integer
    Dim Num         As Integer
    Dim oBola       As bola
    
  On Error GoTo DisMuestraFecha_Error
    
    Range("A1").Activate

    For i = 1 To 49
        ActiveCell.Offset(mRow, 0).Value = mRow
        '
        '
        '   Obtiene la bola de trabajo de la muestra
        '
        Set oBola = datMuestra.Get_Bola(i)
        '
        '   Guarda la fecha del sorteo
        '
        ActiveCell.Offset(mRow, 1).Value = datSorteo.Fecha
        ActiveCell.Offset(mRow, 1).NumberFormat = "dd/mm/yyyy"
        '
        'escribe en la fila correspondiente la informacion
        'de la bola formateando la celda quecontiene la informacion
        '
        ActiveCell.Offset(mRow, 2).Value = oBola.Numero.Valor
        ActiveCell.Offset(mRow, 2).NumberFormat = "0"
        
        ActiveCell.Offset(mRow, 3).Value = oBola.Apariciones
        ActiveCell.Offset(mRow, 3).NumberFormat = "0"
        
        ActiveCell.Offset(mRow, 4).Value = oBola.Ausencias
        ActiveCell.Offset(mRow, 4).NumberFormat = "0"
        
        ActiveCell.Offset(mRow, 5).Value = oBola.Probabilidad
        ActiveCell.Offset(mRow, 5).NumberFormat = "0.000%"
        
        ActiveCell.Offset(mRow, 6).Value = oBola.Prob_TiempoMedio
        ActiveCell.Offset(mRow, 6).NumberFormat = "0.000%"
              
        ActiveCell.Offset(mRow, 7).Value = oBola.Prob_Frecuencia
        ActiveCell.Offset(mRow, 7).NumberFormat = "0.000%"
        
        ActiveCell.Offset(mRow, 8).Value = oBola.Tiempo_Medio
        ActiveCell.Offset(mRow, 8).NumberFormat = "0"
        
        ActiveCell.Offset(mRow, 9).Value = oBola.Desviacion_Tm
        ActiveCell.Offset(mRow, 9).NumberFormat = "0"
        
        ActiveCell.Offset(mRow, 10).Value = oBola.Moda
        ActiveCell.Offset(mRow, 10).NumberFormat = "0"
        
        ActiveCell.Offset(mRow, 11).Value = oBola.Maximo_Tm
        ActiveCell.Offset(mRow, 11).NumberFormat = "0"
        
        ActiveCell.Offset(mRow, 12).Value = oBola.Minimo_Tm
        ActiveCell.Offset(mRow, 12).NumberFormat = "0"
        
        ActiveCell.Offset(mRow, 13).Value = oBola.Numero.Terminacion
        ActiveCell.Offset(mRow, 14).Value = oBola.Numero.Decena
        ActiveCell.Offset(mRow, 15).Value = oBola.Numero.Paridad
        ActiveCell.Offset(mRow, 16).Value = oBola.Numero.Peso
        ActiveCell.Offset(mRow, 17).Value = oBola.Clase_Ausencias
        
        Num = oBola.Numero.Valor
        If datSorteo.Combinacion.Contiene(Num) Then
            ActiveCell.Offset(mRow, 18).Value = 1
        Else
            ActiveCell.Offset(mRow, 18).Value = 0
        End If
        mRow = mRow + 1
    Next i

  On Error GoTo 0
    Exit Sub
DisMuestraFecha_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    Call HandleException(ErrNumber, ErrDescription, "Lot_11_STD_ACUM.DisMuestraFecha")
    Err.Raise ErrNumber, "Lot_11_STD_ACUM.DisMuestraFecha", ErrDescription
End Sub
