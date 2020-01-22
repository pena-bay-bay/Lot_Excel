Attribute VB_Name = "Lot_09_ComprobarMetodo"
    ' Module    : Lot_ComprobarMetodo
' Author    : Charly
' Date      : 11/03/2012
' Purpose   : Comprueba un método por separado en la muestra
'---------------------------------------------------------------------------------------

Option Explicit

'---------------------------------------------------------------------------------------
' Procedure : btn_ComprobarMetodo
' Author    : Charly
' Date      : 25/03/2012
' Purpose   : Caso de Uso que prueba un método de sugerencia con los sorteos
'             CU_RealizarPruebaMetodos
'---------------------------------------------------------------------------------------
'
Public Sub btn_ComprobarMetodo()
    Dim m_frm           As frmProcesaMetodo     'Formulario de configuración del metodo
    Dim objMetodo       As metodo               'Objeto Metodo configurado
    Dim objFechas       As Periodo              'Periodo de sorteos
    
  On Error GoTo btn_ComprobarMetodo_Error
    '
    '   Iniciamos el formulario de captura de parametros del Metodo
    '   Parte del CU_DefinirParametrosMetodo
    '
    Set m_frm = New frmProcesaMetodo
    '
    '   Borra el contenido de la hoja de salida
    '
    Borra_Salida
    '
    '   Inicializa el estado del formulario
    '
    m_frm.Tag = ESTADO_INICIAL
    '
    '   Bucle realizar hasta que no se haya pulsado el BOTON cerrar
    '
    Do While m_frm.Tag <> BOTON_CERRAR
        '
        '    Se inicializa el boton cerrar para salir del bucle
        '
        m_frm.Tag = BOTON_CERRAR
        '
        '   Se muestra el formulario y queda a la espera de funciones
        '   pulsando el botón ejecutar
        m_frm.Show
        '
        '   Se bifurca la función
        '
        Select Case m_frm.Tag
                                    ' El usuario ha cerrado el
            Case ""                 ' cuadro de dialogo con la [X]
                m_frm.Tag = BOTON_CERRAR
            
            Case EJECUTAR           ' Se ha pulsado el botón ejecutar
                 '
                 '  Borra el contenido de la hoja de salida para multiples ejecuciones
                 '
                 Borra_Salida
                 '
                 '  Obtenemos el método y el periodo de prueba del metodo
                 '
                 Set objMetodo = m_frm.metodo
                 Set objFechas = m_frm.PeriodoSorteos
                 '
                 '  Preparamos la hoja de salida con los literales del proceso
                 '
                 pintar_ComprobarMetodo objMetodo, objFechas
                 '
                 '  Ejecutamos el procedimiento
                 '
                 cmd_ComprobarMetodo objMetodo, objFechas
                               
        End Select
        '
        '   Repetimos el bucle
        '
    Loop
    '
    '   Configura la hoja de salida adecuando el tamaño de las letras y posicionandose en la celda activa
    '
    Cells.Select                                ' Selecciona todas las celdas de la hoja
    Cells.EntireColumn.AutoFit                  ' Autoajusta el tamaño de las columnas
    Range("A1").Select                          ' Se posiciona en la celda A1
    
   On Error GoTo 0
   Exit Sub

btn_ComprobarMetodo_Error:
   Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
   ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
   '   Audita el error
   Call HandleException(ErrNumber, ErrDescription, "Lot_ComprobarMetodo.btn_ComprobarMetodo")
   '   Informa del error
   Call MsgBox(ErrDescription, vbError Or vbSystemModal, NOMBRE_APLICACION)
   Call Trace("CERRAR")
End Sub

'---------------------------------------------------------------------------------------
' Procedure : cmd_ComprobarMetodo
' Author    : Charly
' Date      : 10/04/2012
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub cmd_ComprobarMetodo(datMetodo As metodo, datPeriodo As Periodo)
    'Dim mFecha          As Date                     ' Fecha de analisis
    Dim objMuestra      As Muestra                  ' Muestra del analisis
    Dim objSorteo       As Sorteo                   ' Sorteo a comprobar
    Dim objSugerencia   As Sugerencia               ' Sugerencia
    Dim x               As Integer                  ' Coordenada Columnas
    Dim y               As Integer                  ' Coordenada Filas
    Dim DB              As New BdDatos              ' Base de datos
    Dim objRango        As Range                    ' Rango de resultados
    Dim objFila         As Range                    ' Fila de resultados
    Dim objCUDefSug     As CU_DefinirSugerencia     ' CU Definir Sugerencia
    
  On Error GoTo cmd_ComprobarMetodo_Error
    '
    '   obtiene el rango con los datos comprendido entre las dos fechas
    '
    Set objRango = DB.Resultados_Fechas(datPeriodo.FechaInicial, datPeriodo.FechaFinal)
    '
    '   Nos posicionamos en la celda de inicio de escritura e inicializamos coordenadas
    '
    Range("D3").Activate
    x = 0
    y = 0
    '
    '   Inicializamos el sorteo, la sugerencia,
    '   la muestra y el CU
    '
    Set objSorteo = New Sorteo
    Set objSugerencia = New Sugerencia
    Set objMuestra = New Muestra
    Set objCUDefSug = New CU_DefinirSugerencia
    '
    '   Recorremos el rango de sorteos
    '
    For Each objFila In objRango.Rows
        '
        '   Componemos el sorteo
        '
        objSorteo.Constructor objFila
        '
        '   Calcula los parametros de la muestra
        '
        Set objMuestra = GetMuestra(objSorteo.Fecha, datMetodo)
        '
        '   Obtiene la sugerencia
        '
        Set objSugerencia = objCUDefSug.GetSugerencia(datMetodo, objMuestra)
        '
        '   Pinta el sorteo
        '
        pintar_Sorteo ActiveCell.Offset(x, 0), objSorteo, objMuestra, datMetodo.Parametros.CriteriosOrdenacion
        '
        '   Pinta la sugerencia
        '
        pintar_Sugerencia ActiveCell.Offset(x, 8), objSugerencia, objSorteo
        '
        '
        '
        x = x + 1
    Next objFila

   On Error GoTo 0
   Exit Sub

cmd_ComprobarMetodo_Error:
     Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
   ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
   '   Audita el error
   Call HandleException(ErrNumber, ErrDescription, "Lot_ComprobarMetodo.cmd_ComprobarMetodo")
   '   Lanza el Error
   Err.Raise ErrNumber, ErrSource, ErrDescription
End Sub

'---------------------------------------------------------------------------------------
' Procedure : pintar_ComprobarMetodo
' Author    : Charly
' Date      : 10/04/2012
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub pintar_ComprobarMetodo(datMetodo As metodo, datPeriodo As Periodo)
    Dim i As Integer                    ' Indice de columnas
    
  On Error GoTo pintar_ComprobarMetodo_Error
    '
    '   Nos posicionamos en la celda A1 de la hoja salida
    '
    Range("A1").Activate
    '
    '   Visualizamos los parametros del metodo
    '
    ActiveCell.Value = "Procedimiento Probar Metodo"
    ActiveCell.Font.Bold = True                         ' En negrita
    '
    '   Literales del método
    '
    ActiveCell.Offset(1, 0).Value = "Procedimiento"
    ActiveCell.Offset(2, 0).Value = "Ordenacion"
    ActiveCell.Offset(3, 0).Value = "Criterio ordenacion"
    ActiveCell.Offset(4, 0).Value = "Agrupacion"
    ActiveCell.Offset(5, 0).Value = "Dias Muestra"
    ActiveCell.Offset(6, 0).Value = "Registros Muestra"
    ActiveCell.Offset(7, 0).Value = "Pronósticos"
    '
    '   Parametros del método
    '
    ActiveCell.Offset(1, 1).Value = datMetodo.TipoProcedimientoTostring()
    ActiveCell.Offset(2, 1).Value = datMetodo.Parametros.OrdenacionToString()
    ActiveCell.Offset(3, 1).Value = IIf(datMetodo.Parametros.SentidoOrdenacion, "Ascendente", "Descendente")
    ActiveCell.Offset(4, 1).Value = datMetodo.Parametros.AgrupacionToString()
    ActiveCell.Offset(5, 1).Value = datMetodo.Parametros.DiasAnalisis
    ActiveCell.Offset(6, 1).Value = datMetodo.Parametros.NumeroSorteos
    ActiveCell.Offset(7, 1).Value = datMetodo.Parametros.Pronosticos

    '
    '   Literales del periodo
    '
    Range("A10").Activate
    ActiveCell.Value = "Rango de Sorteos"
    ActiveCell.Font.Bold = True
    ActiveCell.Offset(1, 0).Value = "Fecha Final"
    ActiveCell.Offset(2, 0).Value = "Fecha Inicial"
    ActiveCell.Offset(3, 0).Value = "Dias"
     
    '
    '   Valores del periodo
    '
    ActiveCell.Offset(1, 1).Value = datPeriodo.FechaFinal
    ActiveCell.Offset(1, 1).NumberFormat = "ddd, dd/mm/yyyy"        'Formato de fecha lun, 01/05/2012
    ActiveCell.Offset(2, 1).Value = datPeriodo.FechaInicial
    ActiveCell.Offset(2, 1).NumberFormat = "ddd, dd/mm/yyyy"
    ActiveCell.Offset(3, 1).Value = datPeriodo.Dias


    '
    '   Literales de las Columnas del proceso
    '
    Range("D2").Activate
    ActiveCell.Offset(0, 0).Value = "F.Sorteo"
    For i = 1 To 6
        ActiveCell.Offset(0, i).Value = "N" + CStr(i)
    Next i
    ActiveCell.Offset(0, i).Value = "C"
    i = i + 1
    ActiveCell.Offset(0, i).Value = "_"
    '
    '   Datos sugerencia
    '
    Range("L2").Activate
    
    For i = 1 To datMetodo.Parametros.Pronosticos
        ActiveCell.Offset(0, i).Value = "P" + CStr(i)
    Next i
    '
    '   Aciertos
    '
    ActiveCell.Offset(0, i).Value = "A"
    i = i + 1
    ActiveCell.Offset(0, i).Value = "Premio"
    '
    '   Configuramos el tamaño de las columnas
    '
    Cells.Select                                ' Selecciona todas las celdas de la hoja
    Cells.EntireColumn.AutoFit                  ' Autoajusta el tamaño de las columnas
   
   On Error GoTo 0
   Exit Sub

pintar_ComprobarMetodo_Error:
     Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
   ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
   '   Audita el error
   Call HandleException(ErrNumber, ErrDescription, "Lot_ComprobarMetodo.pintar_ComprobarMetodo")
   '   Lanza el Error
   Err.Raise ErrNumber, ErrSource, ErrDescription
End Sub

'---------------------------------------------------------------------------------------
' Procedure : pintar_Sugerencia
' Author    : Charly
' Date      : 14/04/2012
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub pintar_Sugerencia(datCelda As Range, _
                              datSugerencia As Sugerencia, _
                              datSorteo As Sorteo)
    
    Dim i           As Integer
    'Dim mNum        As Numero
    Dim mBola       As bola
    Dim iNum        As Integer
    Dim iColor      As Long
    Dim m_oCU       As New CU_ValidarSugerencia     ' define y crea un caso de uso validar sugerencia
    
  On Error GoTo pintar_Sugerencia_Error
    '
    '   Bucle de extracción de los números de la sugerencia
    '
    For i = 1 To datSugerencia.Combinacion.Count
        '
        '   Establecemos el Numero
        '
        Set mBola = datSugerencia.Bolas.Item(i)
        '
        '   Ponemos el valor
        '
        iNum = mBola.Numero.Valor
        datCelda.Offset(0, i).Value = iNum
        '
        '   Colorear la celda del color de la ordenación
        '
        Select Case datSugerencia.metodo.Parametros.CriteriosOrdenacion
            Case ordFrecuencia
                iColor = mBola.Color_Frecuencias
            Case ordProbTiempoMedio
                iColor = mBola.Color_Tiempo_Medio
            Case Else
                iColor = mBola.Color_Probabilidad
        End Select
        datCelda.Offset(0, i).Interior.ColorIndex = iColor
        '
        '    si existe en sorteo se pone en negrita
        '
        If (datSorteo.Combinacion.Contiene(iNum)) Then
            datCelda.Offset(0, i).Font.Bold = True
        End If
    Next i
    '
    '   Obtener el Numero de aciertos
    '
    datCelda.Offset(0, i).Value = m_oCU.GetAciertos(datSugerencia, datSorteo)
    i = i + 1
    '
    '   Obtiene el premio
    '
    datCelda.Offset(0, i).Value = m_oCU.GetPremio(datSugerencia, datSorteo)

   On Error GoTo 0
   Exit Sub

pintar_Sugerencia_Error:
     Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
   ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
   '   Audita el error
   Call HandleException(ErrNumber, ErrDescription, "Lot_ComprobarMetodo.pintar_Sugerencia")
   '   Lanza el Error
   Err.Raise ErrNumber, ErrSource, ErrDescription
End Sub

'---------------------------------------------------------------------------------------
' Procedure : pintar_Sorteo
' Author    : Charly
' Date      : 14/04/2012
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub pintar_Sorteo(datCelda As Range, _
                          datSorteo As Sorteo, _
                          datMuestra As Muestra, _
                          datOrden As TipoOrdenacion)
                          
    Dim i               As Integer      ' Indice
    Dim mNum            As Numero       ' Numero
    Dim mBola           As bola
    
  On Error GoTo pintar_Sorteo_Error
    '
    '   Coloca la fecha en la celda en formato lun, 1/05/2012
    '
    datCelda.Value = datSorteo.Fecha
    datCelda.NumberFormat = "ddd, dd/mm/yyyy"
    '
    '   Para los Numeros del sorteo y el complementario
    '
    For i = 1 To 7
        '
        '   Obtiene el Numero de la combinación del sorteo
        '
        Set mNum = datSorteo.Combinacion.Numeros(i)
        '
        '   Obtiene la bola de la muestra
        '
        Set mBola = datMuestra.Get_Bola(mNum.Valor)
        '
        '   Establece el valor en la celda
        '
        datCelda.Offset(0, i).Value = mNum.Valor
        '
        '   Colorea la celda con el  criterio de la ordenación
        '
        
        datCelda.Offset(0, i).Interior.ColorIndex = mBola.Color_Probabilidad
    Next i

   On Error GoTo 0
   Exit Sub

pintar_Sorteo_Error:
     Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
   ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
   '   Audita el error
   Call HandleException(ErrNumber, ErrDescription, "Lot_ComprobarMetodo.pintar_Sorteo")
   '   Lanza el Error
   Err.Raise ErrNumber, ErrSource, ErrDescription
End Sub

'---------------------------------------------------------------------------------------
' Procedure : GetMuestra
' Author    : Charly
' Date      : 14/04/2012
' Purpose   : Configura una muestra de bolas para un rango de fechas definido en el
'             metodo
'---------------------------------------------------------------------------------------
'
Private Function GetMuestra(datFeAnalisis As Date, datMetodo As metodo) As Muestra
    Dim objParam            As ParametrosMuestra        ' Parametros de muestra
    Dim objMuestra          As Muestra                  ' Muestra de sorteos
    Dim objBd               As New BdDatos              ' Base de datos
    Dim objRg               As Range                    ' Rango de sorteos
  
  On Error GoTo GetMuestra_Error
    '
    '   Calculamos parametros de la muestra
    '
    Set objParam = New ParametrosMuestra
    objParam.FechaAnalisis = datFeAnalisis
    objParam.FechaFinal = datFeAnalisis - 1
    '
    '   Si el tipo de muestra se basa en dias o registros
    '
    If datMetodo.TipoMuestra Then
        '
        '   Calcula las fechas del rango
        '
        objParam.FechaInicial = objParam.FechaFinal - datMetodo.Parametros.DiasAnalisis
    Else
        '
        '   Calcula los registros
        '
        objParam.NumeroSorteos = datMetodo.Parametros.NumeroSorteos
    End If
    '
    '   obtiene el rango con los datos comprendido entre las dos fechas
    '
    Set objRg = objBd.Resultados_Fechas(objParam.FechaInicial, _
                                         objParam.FechaFinal)
    '
    '   Se crea la muestra
    '
    Set objMuestra = New Muestra
    '
    '   El rango de sorteos se lo pasamos al constructor de la clase y
    '   obtiene las estadisticas para cada bola
    '
    Set objMuestra.ParametrosMuestra = objParam
    objMuestra.Constructor objRg, JUEGO_DEFECTO
    '
    '   Se devuelve la muestra
    '
    Set GetMuestra = objMuestra

   On Error GoTo 0
   Exit Function

GetMuestra_Error:
     Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
   ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
   '   Audita el error
   Call HandleException(ErrNumber, ErrDescription, "Lot_ComprobarMetodo.GetMuestra")
   '   Lanza el Error
   Err.Raise ErrNumber, ErrSource, ErrDescription
End Function



