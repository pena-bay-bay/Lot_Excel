Attribute VB_Name = "Lot_03_ObtenerEstadisticas"
' *============================================================================*
' *
' *     Fichero    : Lot_03_ObtenerEstadisticas
' *
' *     Autor      : Carlos Almela Baeza
' *     Creación   : mar, 17/01/2012 23:45
' *     Versión    : 1.0
' *     Propósito  : Obtener estadisticas de un periodo definido
' *
' *
' *============================================================================*
Option Explicit



' *============================================================================*
' *     Procedure  : btn_Obtener_Estadisticas
' *     Fichero    : Lot_ObtenerEstadisticas
' *     Autor      : Carlos Almela Baeza
' *     Creacion   : sáb, 14/01/2012 19:12
' *     Asunto     :
' *============================================================================*
'
Public Sub btn_Obtener_Estadisticas()
    Dim objMuestra          As New Muestra            ' objeto muestra
    Dim m_objRg             As Range                  ' rango de datos
    Dim m_objBd             As New BdDatos            ' base de datos
    Dim mFrm                As frmMuestra             ' Formulario de párametros de la muestra
    Dim m_objParMuestra     As ParametrosMuestra      ' Parametros de la muestra
  On Error GoTo btn_Obtener_Estadisticas_Error

    Borra_Salida                                      ' Borrar la hoja de salida
          
    Set mFrm = New frmMuestra                         ' Crea un formulario de captura de la muestra
    
    mFrm.Tag = ESTADO_INICIAL                         ' inicializa el estado
    
    Do While mFrm.Tag <> BOTON_CERRAR                 ' Bucle hasta pulsar Ejecutar o Salir
        '
        ' Se inicializa el boton cerrar para salir del bucle
        mFrm.Tag = BOTON_CERRAR
        '
        ' Se muestra el formulario y queda a la espera de funciones
        ' pulsando el botón ejecutar
        mFrm.Show
        '
        ' Se bifurca la función
        '
        Select Case mFrm.Tag
                                                     ' El usuario ha cerrado el
            Case ""                                  ' cuadro de dialogo con la [X]
                mFrm.Tag = BOTON_CERRAR
                Exit Sub
                
            Case EJECUTAR                            ' Se ha pulsado el botón ejecutar
                Set m_objParMuestra = mFrm.ParMuestra
                mFrm.Tag = BOTON_CERRAR
                   
            Case BOTON_CERRAR
                Exit Sub
        
        End Select
    Loop
     
    Set mFrm = Nothing                               ' Elimina el formulario
    '
    '       Calcula la Muestra
    '
    '   obtiene el rango con los datos comprendido entre las dos fechas
    '
    Set m_objRg = m_objBd.GetSorteosInFechas(m_objParMuestra.PeriodoDatos)
    '
    '   se lo pasa al constructor de la clase y obtiene las estadisticas para cada bola
    '
    Set objMuestra.ParametrosMuestra = m_objParMuestra
    Select Case JUEGO_DEFECTO
        Case LoteriaPrimitiva, Bonoloto:
            objMuestra.Constructor m_objRg, ModalidadJuego.LP_LB_6_49
        
        Case GordoPrimitiva:
            objMuestra.Constructor m_objRg, ModalidadJuego.GP_5_54
        
        Case Euromillones:
            objMuestra.Constructor m_objRg, ModalidadJuego.EU_5_50
            
    End Select
    '
    '
    '       Pinta la muestra en la hoja
    '
    Pintar_Muestra objMuestra
    '
    '       Se posiciona en la celda A9
    '
    Range("A9").Select
    
btn_Obtener_Estadisticas_CleanExit:
  On Error GoTo 0
    Exit Sub
btn_Obtener_Estadisticas_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    Call HandleException(ErrNumber, ErrDescription, "Lot_ObtenerEstadisticas.btn_Obtener_Estadisticas")
    Call MsgBox(ErrDescription, vbError Or vbSystemModal, NOMBRE_APLICACION)
    Call Trace("CERRAR")
End Sub





' *============================================================================*
' *     Procedure  : Pintar_Muestra
' *     Fichero    : Lot_ObtenerEstadisticas
' *     Autor      : Carlos Almela Baeza
' *     Creacion   : jue, 09/02/2012 23:41
' *     Asunto     :
' *============================================================================*
'
Public Sub Pintar_Muestra(ByRef objMuestra As Muestra)
    Dim m_objPar            As ParametrosMuestra
    Dim i                   As Integer
    Dim m_objBola           As Bola
    Dim mSort               As Sorteo
    Dim mEngSort            As SorteoEngine
    Dim mMin                As Integer
    Dim mMax                As Integer
    
  On Error GoTo Pintar_Muestra_Error
    '
    '   escribe los textos de la estadistica
    '
    Range("A1").Activate
    ActiveCell.Value = "Estadisticas sobre números "
    ActiveCell.Font.Bold = True
    ActiveCell.Offset(1, 0).Value = "Fecha Analisis"
    ActiveCell.Offset(2, 0).Value = "Fecha de inicio"
    ActiveCell.Offset(3, 0).Value = "Fecha de Fin"
    ActiveCell.Offset(4, 0).Value = "Dias Analizados"
    ActiveCell.Offset(5, 0).Value = "Numero de Sorteos "
    ActiveCell.Offset(6, 0).Value = "Total Numeros"
    '
    '   Texto de la fila
    '
    ActiveCell.Offset(8, 0).Value = "Numero"
    ActiveCell.Offset(8, 1).Value = "Apariciones"
    ActiveCell.Offset(8, 2).Value = "Ausencias"
    ActiveCell.Offset(8, 3).Value = "Prob"
    ActiveCell.Offset(8, 4).Value = "Prob Tiempo"
    ActiveCell.Offset(8, 5).Value = "Prob Frecuencias"
    ActiveCell.Offset(8, 6).Value = "Tiempo"
    ActiveCell.Offset(8, 7).Value = "Desv"
    ActiveCell.Offset(8, 8).Value = "Moda"
    ActiveCell.Offset(8, 9).Value = "Max"
    ActiveCell.Offset(8, 10).Value = "Min"
    ActiveCell.Offset(8, 11).Value = "Ultima Fecha"
    ActiveCell.Offset(8, 12).Value = "Proxima Fecha"
    ActiveCell.Offset(8, 13).Value = "Terminación"
    ActiveCell.Offset(8, 14).Value = "Decena"
    ActiveCell.Offset(8, 15).Value = "Paridad"
    ActiveCell.Offset(8, 16).Value = "Peso"
    ActiveCell.Offset(8, 17).Value = "Tendencia"
    ActiveCell.Offset(8, 18).Value = "C.Ausencias"
    ActiveCell.Offset(8, 19).Value = "V.Homogeneo"
    
    '   Obtiene parametros de la muestra
    '
    Set m_objPar = objMuestra.ParametrosMuestra
    '
    '   Visualiza Parámetros de la muestra
    '
    Range("B2").Activate
    ActiveCell.Value = m_objPar.FechaAnalisis
    ActiveCell.Offset(1, 0).Value = m_objPar.FechaInicial
    ActiveCell.Offset(2, 0).Value = m_objPar.FechaFinal
    ActiveCell.Offset(3, 0).Value = objMuestra.Total_Dias
    ActiveCell.Offset(4, 0).Value = m_objPar.NumeroSorteos
    ActiveCell.Offset(5, 0).Value = objMuestra.Total_Numeros
    '
    '   Obtiene el sorteo
    '
    Set mEngSort = New SorteoEngine
    Set mSort = mEngSort.GetSorteoByFecha(m_objPar.FechaAnalisis)
    '
    '   Matriz de información estadistica
    '
    Range("A9").Activate    'Se posiciona en la celda A9
    '
    '   Determinamos bolas del juego
    '
    Select Case JUEGO_DEFECTO
        Case Bonoloto, LoteriaPrimitiva:
            mMin = 1
            mMax = 49
        Case GordoPrimitiva:
            mMin = 1
            mMax = 54
        Case Euromillones:
            mMin = 1
            mMax = 50
    End Select
    '
    '   para cada bola
    '
    For i = mMin To mMax
        '
        '
        '   Obtiene la bola de trabajo de la muestra
        '
        Set m_objBola = objMuestra.Get_Bola(i)
                
        'escribe en la fila correspondiente la informacion
        'de la bola; coloreando y formateando la celda que
        'contiene la informacion
        ActiveCell.Offset(i, 0).Value = m_objBola.Numero.Valor
        ActiveCell.Offset(i, 0).NumberFormat = "00"
        If Not (mSort Is Nothing) Then
            If mSort.Combinacion.Contiene(m_objBola.Numero.Valor) Then
                    ActiveCell.Offset(i, 0).Interior.ColorIndex = COLOR_VERDE
            End If
            If mSort.Complementario = m_objBola.Numero.Valor Then
                    ActiveCell.Offset(i, 0).Interior.ColorIndex = COLOR_NUMCOMPLE
            End If
        End If
        ActiveCell.Offset(i, 1).Value = m_objBola.Apariciones
        ActiveCell.Offset(i, 1).NumberFormat = "0"
        
        ActiveCell.Offset(i, 2).Value = m_objBola.Ausencias
        ActiveCell.Offset(i, 2).NumberFormat = "0"
        
        ActiveCell.Offset(i, 3).Value = m_objBola.Probabilidad
        ActiveCell.Offset(i, 3).NumberFormat = "0.000%"
        
        ActiveCell.Offset(i, 4).Value = m_objBola.Prob_TiempoMedio
        ActiveCell.Offset(i, 4).NumberFormat = "0.000%"
              
        ActiveCell.Offset(i, 5).Value = m_objBola.Prob_Frecuencia
        ActiveCell.Offset(i, 5).NumberFormat = "0.000%"
        
        ActiveCell.Offset(i, 6).Value = m_objBola.Tiempo_Medio
        ActiveCell.Offset(i, 6).NumberFormat = "0.00"
        
        ActiveCell.Offset(i, 7).Value = m_objBola.Desviacion_Tm
        ActiveCell.Offset(i, 7).NumberFormat = "0.00"
        
        ActiveCell.Offset(i, 8).Value = m_objBola.Moda
        ActiveCell.Offset(i, 8).NumberFormat = "0"
        
        ActiveCell.Offset(i, 9).Value = m_objBola.Maximo_Tm
        ActiveCell.Offset(i, 9).NumberFormat = "0"
        
        ActiveCell.Offset(i, 10).Value = m_objBola.Minimo_Tm
        ActiveCell.Offset(i, 10).NumberFormat = "0"
        
        ActiveCell.Offset(i, 11).Value = m_objBola.Ultima_Fecha
        ActiveCell.Offset(i, 11).NumberFormat = "dd/mm/yyyy"
        
        ActiveCell.Offset(i, 12).Value = m_objBola.ProximaFecha
        ActiveCell.Offset(i, 12).NumberFormat = "dd/mm/yyyy"
        
        ActiveCell.Offset(i, 13).Value = m_objBola.Numero.Terminacion
        ActiveCell.Offset(i, 14).Value = m_objBola.Numero.Decena
        ActiveCell.Offset(i, 15).Value = m_objBola.Numero.Paridad
        ActiveCell.Offset(i, 16).Value = m_objBola.Numero.Peso
        ActiveCell.Offset(i, 17).Value = m_objBola.Tendencia
        ActiveCell.Offset(i, 18).Value = m_objBola.Clase_Ausencias
        ActiveCell.Offset(i, 19).Value = m_objBola.ValorHomogeneo
        ActiveCell.Offset(i, 19).NumberFormat = "0.000"
    Next i
    
    'Colorea los rangos de probabilidades
    Select Case JUEGO_DEFECTO
        Case Bonoloto, LoteriaPrimitiva:
            'Colorea los rangos de probabilidades
            Colorear_Matriz Range("D10:D58"), True         'Probabilidad del n?mero
            Colorear_Matriz Range("E10:E58"), True         'Tiempo Medio
            Colorear_Matriz Range("F10:F58"), True         'Probabilidad de la frecuencia
            Colorear_Matriz Range("M10:M58"), False        'Pr?xima fecha
            Colorear_Matriz Range("g10:g58"), False        'Tiempo medio
            Colorear_Matriz Range("H10:H58"), False        'Desviacion
            Colorear_Matriz Range("i10:i58"), False        'Moda
        Case GordoPrimitiva:
            'Colorea los rangos de probabilidades
            Colorear_Matriz Range("D10:D63"), True         'Probabilidad del n?mero
            Colorear_Matriz Range("E10:E63"), True         'Tiempo Medio
            Colorear_Matriz Range("F10:F63"), True         'Probabilidad de la frecuencia
            Colorear_Matriz Range("M10:M63"), False        'Pr?xima fecha
            Colorear_Matriz Range("g10:g63"), False        'Tiempo medio
            Colorear_Matriz Range("H10:H63"), False        'Desviacion
            Colorear_Matriz Range("i10:i63"), False        'Moda
        Case Euromillones:
            'Colorea los rangos de probabilidades
            Colorear_Matriz Range("D10:D59"), True         'Probabilidad del n?mero
            Colorear_Matriz Range("E10:E59"), True         'Tiempo Medio
            Colorear_Matriz Range("F10:F59"), True         'Probabilidad de la frecuencia
            Colorear_Matriz Range("M10:M59"), False        'Pr?xima fecha
            Colorear_Matriz Range("g10:g59"), False        'Tiempo medio
            Colorear_Matriz Range("H10:H59"), False        'Desviacion
            Colorear_Matriz Range("i10:i59"), False        'Moda
    End Select
    '
    '
    '
    Cells.Select                'Selecciona todas las celdas de la hoja
    Cells.EntireColumn.AutoFit  'Autoajusta el tamaño de las columnas
    
    Range("A9").Select          'Se posiciona en la celda del primer número
    Selection.AutoFilter        'Crea un autofiltro

Pintar_Muestra_CleanExit:
  On Error GoTo 0
    Exit Sub

Pintar_Muestra_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    Call HandleException(ErrNumber, ErrDescription, "Lot_ObtenerEstadisticas.Pintar_Muestra")
    Err.Raise ErrNumber, ErrSource, ErrDescription
End Sub

' *===========(EOF): Lot_03_ObtenerEstadisticas
