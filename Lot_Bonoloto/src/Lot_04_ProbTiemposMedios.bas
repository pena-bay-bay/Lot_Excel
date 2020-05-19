Attribute VB_Name = "Lot_04_ProbTiemposMedios"
' *============================================================================*
' *
' *     Fichero    : Lot_ProbTiemposMedios
' *
' *     Autor      : Carlos Almela Baeza
' *     Creación   : dom, 12/02/2012 23:22
' *     Versión    : 1.0
' *     Propósito  : Obtiene la estadistica detallada para todos los números
' *                  o un conjunto de ellos
' *
' *============================================================================*
Option Explicit
'
'   Variables Privadas
'
Dim oMuestra            As Muestra              ' Muestra a dibujar
Dim iNumero             As Integer              ' Numero a dibujar
Dim oSorteo             As Sorteo               ' Sorteo a dibujar
Dim dFechaSorteo        As Date                 ' Fecha de Sorteo
Dim oNumero             As Numero               ' Numero a analizar
Dim oCombinacion        As Combinacion          ' Conjunto de números
Dim oSorteoEngine       As SorteoEngine         ' Motor de sorteos
Dim iCol                As Integer              ' Coordenada de columnas
Dim iRow                As Integer              ' Coordenada de Filas
Dim N                   As Integer              ' Entero
Dim rgCelda             As Range                ' Celda activa

'---------------------------------------------------------------------------------------
' Procedimiento : btn_Prob_TiemposMedios
' Creación      : 9-may-2007 15:15
' Autor         : Carlos Almela Baeza
' Objeto        :
'---------------------------------------------------------------------------------------
'
Public Sub btn_Prob_TiemposMedios()
  
  On Error GoTo btn_Prob_TiemposMedios_Error
    '
    '   Borra Hoja de Salida
    '
    Borra_Salida
    
    frmEstadisticaNumero.Tag = ESTADO_INICIAL  ' Se asigna el estado inicial a la etiqueta
                                            ' del formulario
                                            ' Mientras la etiqueta del formulario no tenga
                                            ' el indicador de BOTON_CERRAR se ejecutara
                                            ' el bucle de proceso
    Do While frmEstadisticaNumero.Tag <> BOTON_CERRAR
        
        ' Se inicializa el boton cerrar para salir del bucle
        frmEstadisticaNumero.Tag = BOTON_CERRAR
        
        ' Se muestra el formulario y queda a la espera de funciones
        ' pulsando el botón ejecutar
        frmEstadisticaNumero.Show vbModal
        
        'Se bifurca la función
        Select Case frmEstadisticaNumero.Tag
                                    ' El usuario ha cerrado el
            Case ""                 ' cuadro de dialogo con la [X]
                frmEstadisticaNumero.Tag = BOTON_CERRAR
            
            Case EJECUTAR
                Application.ScreenUpdating = False
                '
                '   Obtiene los datos del formulario
                '
                Set oMuestra = frmEstadisticaNumero.MuestraCalculo
                dFechaSorteo = frmEstadisticaNumero.FechaSorteo
                Set oCombinacion = frmEstadisticaNumero.Combinacion
                '
                '   En función del tipo de proceso realiza una función u otra
                '
                Select Case frmEstadisticaNumero.TipoProceso
                    Case 1: ' Todos los Numeros
                        cmd_CalculaTodosProb oMuestra
                          
                         
                    Case 2: ' Los Numeros de un sorteo
                        cmd_NumerosSorteo oMuestra, dFechaSorteo
                         
                    Case 3: ' Un conjunto de números
                        cmd_CalculaCombinacion oMuestra, oCombinacion
                        
                End Select
                
                Application.ScreenUpdating = True
        End Select
    Loop


btn_Prob_TiemposMedios_CleanExit:
    Set frmEstadisticaNumero = Nothing
   On Error GoTo 0
    Exit Sub

btn_Prob_TiemposMedios_Error:
    

    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    '   Audita el error
    Call HandleException(ErrNumber, ErrDescription, "Lot_ProbTiemposMedios.btn_Prob_TiemposMedios")
    '   Informa del error
    Call MsgBox(ErrDescription, vbError Or vbSystemModal, NOMBRE_APLICACION)
    '   cierra el log
    Trace ("CERRAR")
End Sub

' *============================================================================*
' *     Procedure  : cmd_CalculaTodosProb
' *     Fichero    : Lot_ProbTiemposMedios
' *     Autor      : Carlos Almela Baeza
' *     Creacion   : dom, 12/02/2012 20:35
' *     Asunto     :
' *============================================================================*
'
Private Sub cmd_CalculaTodosProb(obj_muestra As Muestra)
    
  On Error GoTo cmd_CalculaTodosProb_Error
    '
    '   inicializa coordenadas x=1, y=1
    '
    iCol = 1:    iRow = 1
    '
    '   Borra la hoja de salida
    '
    Borra_Salida
    '
    '   Nos posicionamos en A1
    '
    Range("A1").Activate
    '
    '   Situamos el zoom de la ventana a 90%
    '
    ActiveWindow.Zoom = 90          'situar la ventana
    '
    '   Título de la hoja
    '
    ActiveCell.Value = "Frecuencias de todos los Números"
    '
    '   En negrita
    '
    ActiveCell.Font.Bold = True
    '
    '   Para los 49 Numeros
    '
    For N = 1 To 49
        '
        '   Posiciona el cursor en la celda correspondiente
        '
        Set rgCelda = Range(Cells(iRow, iCol), Cells(iRow, iCol))
        '
        '   Dibuja la información del Numero N
        '
        cmd_CalculaUnNumero obj_muestra, N, rgCelda
        '
        '   Recalcula coordenadas para el siguiente número
        '         x = x + 15
        '         y = y + 100
        '   cada 10 Numeros x = 1
        '
        If (N Mod 10) = 0 Then
            iCol = 1
            iRow = iRow + 100
        Else
            iCol = iCol + 15
        End If
    Next N
    
    Cells.Select                                ' Selecciona todas las celdas de la hoja
    Cells.EntireColumn.AutoFit                  ' Autoajusta el tamaño de las columnas

cmd_CalculaTodosProb_CleanExit:
   On Error GoTo 0
    Exit Sub

cmd_CalculaTodosProb_Error:

    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    '   Audita el error
    Call HandleException(ErrNumber, ErrDescription, "Lot_ProbTiemposMedios.cmd_CalculaTodosProb")
    '   Lanza el Error
    Err.Raise ErrNumber, ErrSource, ErrDescription
End Sub

' *============================================================================*
' *     Procedure  : cmd_CalculaUnNumero
' *     Fichero    : Lot_ProbTiemposMedios
' *     Autor      : Carlos Almela Baeza
' *     Creacion   : dom, 12/02/2012 20:35
' *     Asunto     :
' *============================================================================*
'
Private Sub cmd_CalculaUnNumero(obj_muestra As Muestra, obj_N As Integer, obj_rango As Range)
    Dim objPar              As ParametrosMuestra
    Dim oBola               As New bola
    
  On Error GoTo cmd_CalculaUnNumero_Error
    '
    '   Dibuja los Textos
    '
    Print_Texto obj_rango
    '
    '   Obtiene los parametros de la muestra
    '
    Set objPar = obj_muestra.ParametrosMuestra
    '
    '   Obtiene los datos de la bola según el número
    '
    Set oBola = obj_muestra.Get_Bola(obj_N)
    '
    '   Dibuja los valores de la bola
    '
    PrintBola obj_rango, oBola, objPar
    '
    '   Dibuja las fechas de aparición de la bola
    '
    PrintFechas obj_rango, oBola, objPar
    '
    '   Calcula las frecuencias de aparición
    '
    PrintFrecuencias obj_rango, oBola, objPar
    
cmd_CalculaUnNumero_CleanExit:
   On Error GoTo 0
    Exit Sub

cmd_CalculaUnNumero_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    '   Audita el error
    Call HandleException(ErrNumber, ErrDescription, "Lot_ProbTiemposMedios.cmd_CalculaUnNumero")
    '   Lanza el Error
    Err.Raise ErrNumber, ErrSource, ErrDescription
End Sub

' *============================================================================*
' *     Procedure  : cmd_NumerosSorteo
' *     Fichero    : Lot_ProbTiemposMedios
' *     Autor      : Carlos Almela Baeza
' *     Creacion   : sab, 03/02/2018
' *     Asunto     : Visualiza los Numeros de un sorteo
' *============================================================================*
'
Private Sub cmd_NumerosSorteo(obj_muestra As Muestra, obj_Fecha As Date)

On Error GoTo cmd_NumerosSorteo_Error
    '
    '   inicializa coordenadas x=1, y=1
    '
    iCol = 1:    iRow = 1
    '
    '   Borra la hoja de salida
    '
    Borra_Salida
    '
    '   Nos posicionamos en A1
    '
    Range("A1").Activate
    '
    '   Situamos el zoom de la ventana a 90%
    '
    ActiveWindow.Zoom = 90          'situar la ventana
    '
    '   Título de la hoja
    '
    ActiveCell.Value = "Frecuencias de un Sorteo"
    '
    '   En negrita
    '
    ActiveCell.Font.Bold = True
    '
    '   Obtenemos el sorteo de la fecha
    '
    Set oSorteoEngine = New SorteoEngine
    
    Set oSorteo = oSorteoEngine.GetSorteoByFecha(obj_Fecha)
    '
    '   Bucle de selección de cada Numero del sorteo
    '
    For N = 1 To oSorteo.Combinacion.Count
        '
        '   Obtenemos el Numero iesimo
        '
        Set oNumero = oSorteo.Combinacion.Numeros(N)
        '
        '   Para el valor del Numero
        '
        iNumero = oNumero.Valor
        '
        '   posicion
        '
        Set rgCelda = Range(Cells(iRow, iCol), Cells(iRow, iCol))
        '
        '   Visualiza la información de un Numero
        '
        cmd_CalculaUnNumero obj_muestra, iNumero, rgCelda
        '
        '   Calculamos coordenada siguiente Numero
        '
        If (N Mod 2) = 0 Then
            iCol = 1
            iRow = iRow + 120
        Else
            iCol = iCol + 15
        End If
    Next N

cmd_NumerosSorteo_CleanExit:
   On Error GoTo 0
    Exit Sub

cmd_NumerosSorteo_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    '   Audita el error
    Call HandleException(ErrNumber, ErrDescription, "Lot_ProbTiemposMedios.cmd_NumerosSorteo")
    '   Lanza el Error
    Err.Raise ErrNumber, ErrSource, ErrDescription
End Sub

' *============================================================================*
' *     Procedure  : cmd_CalculaUnNumero
' *     Fichero    : Lot_ProbTiemposMedios
' *     Autor      : Carlos Almela Baeza
' *     Creacion   : dom, 12/02/2012 20:35
' *     Asunto     :
' *============================================================================*
'
Private Sub cmd_CalculaCombinacion(obj_muestra As Muestra, obj_Combinacion As Combinacion)
    Dim objPar              As ParametrosMuestra
    Dim oBola               As New bola
    Dim oNum                As New Numero
    
  On Error GoTo cmd_CalculaCombinacion_Error
    '
    '   inicializa coordenadas x=1, y=1
    '
    iCol = 1:    iRow = 1:  N = 0
    '
    '   Borra la hoja de salida
    '
    Borra_Salida
    '
    '   Nos posicionamos en A1
    '
    Range("A1").Activate
    '
    '   Situamos el zoom de la ventana a 90%
    '
    ActiveWindow.Zoom = 90          'situar la ventana
    '
    '   Título de la hoja
    '
    ActiveCell.Value = "Frecuencias de Numeros"
    '
    '   En negrita
    '
    ActiveCell.Font.Bold = True
    '
    '   Para cada bola en la combinación
    '
    For Each oNum In obj_Combinacion.Numeros
        '
        '   nesima bola
        '
        N = N + 1
        '
        '   posicion
        '
        Set rgCelda = Range(Cells(iRow, iCol), Cells(iRow, iCol))
        '
        '   Visualiza la información de un Numero
        '
        cmd_CalculaUnNumero obj_muestra, oNum.Valor, rgCelda
        '
        '   Calculamos coordenada siguiente Numero
        '
        If (N Mod 2) = 0 Then
            iCol = 1
            iRow = iRow + 120
        Else
            iCol = iCol + 15
        End If
    Next oNum
    
cmd_CalculaCombinacion_CleanExit:
   On Error GoTo 0
    Exit Sub

cmd_CalculaCombinacion_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    '   Audita el error
    Call HandleException(ErrNumber, ErrDescription, "Lot_ProbTiemposMedios.cmd_CalculaCombinacion")
    '   Lanza el Error
    Err.Raise ErrNumber, ErrSource, ErrDescription

End Sub

' *============================================================================*
' *     Procedure  : Print_Texto
' *     Fichero    : Lot_ProbTiemposMedios
' *     Autor      : Carlos Almela Baeza
' *     Creacion   : dom, 12/02/2012 23:08
' *     Asunto     : Imprime las etiquetas de la estadistica
' *============================================================================*
'
Private Sub Print_Texto(oCelda As Range)
  
  On Error GoTo Print_Texto_Error
    '
    '   Se posiciona en la celda activa
    '
    oCelda.Activate
    '
    '
    '
    With ActiveCell.Offset(1, 0)
        .Value = "Parametros"
        .Interior.ColorIndex = 15
        .Interior.Pattern = xlSolid
        .HorizontalAlignment = xlCenter
    End With
    With ActiveCell.Offset(1, 1)
        .Value = "Valor"
        .Interior.ColorIndex = 15
        .Interior.Pattern = xlSolid
        .HorizontalAlignment = xlCenter
    End With
    
    ActiveCell.Offset(2, 0).Value = "Análisis Inicio Periodo"
    ActiveCell.Offset(3, 0).Value = "         Fin Periodo"
    ActiveCell.Offset(4, 0).Value = "Fecha de prevision"
    ActiveCell.Offset(6, 0).Value = "Numero"
    ActiveCell.Offset(7, 0).Value = "Apariciones"
    ActiveCell.Offset(8, 0).Value = "Ausencias"
    ActiveCell.Offset(9, 0).Value = "Probabilidad"
    ActiveCell.Offset(10, 0).Value = "Prob.Tiempo Medio"
    ActiveCell.Offset(11, 0).Value = "Prob.Frecuencias"
    ActiveCell.Offset(12, 0).Value = "Tiempo medio "
    ActiveCell.Offset(13, 0).Value = "Desviación"
    ActiveCell.Offset(14, 0).Value = "Máximo"
    ActiveCell.Offset(15, 0).Value = "Mínimo"
    ActiveCell.Offset(16, 0).Value = "Ultima Fecha"
    ActiveCell.Offset(17, 0).Value = "Próxima Fecha"
    ActiveCell.Offset(18, 0).Value = "Terminación"
    ActiveCell.Offset(19, 0).Value = "Decena"
    ActiveCell.Offset(20, 0).Value = "Paridad"
    ActiveCell.Offset(21, 0).Value = "Peso"
    ActiveCell.Offset(22, 0).Value = "Tendencia"
    'ActiveCell.Offset(23, 0).Value = "Prob.Tendencia"

Print_Texto_CleanExit:
   On Error GoTo 0
    Exit Sub

Print_Texto_Error:
    

    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    '   Audita el error
    Call HandleException(ErrNumber, ErrDescription, "Lot_ProbTiemposMedios.Print_Texto")
    '   Lanza el Error
    Err.Raise ErrNumber, ErrSource, ErrDescription
End Sub

' *============================================================================*
' *     Procedure  : PrintBola
' *     Fichero    : Lot_ProbTiemposMedios
' *     Autor      : Carlos Almela Baeza
' *     Creacion   : dom, 12/02/2012 20:37
' *     Asunto     :
' *============================================================================*
'
Private Sub PrintBola(oCelda As Range, oBola As bola, oPar As ParametrosMuestra)
    
  On Error GoTo PrintBola_Error
    
    oCelda.Offset(2, 1).Activate
    ActiveCell.Value = oPar.FechaInicial
    ActiveCell.NumberFormat = "dd/mmm/yyyy"
    ActiveCell.Offset(1, 0).Value = oPar.FechaFinal
    ActiveCell.Offset(1, 0).NumberFormat = "dd/mmm/yyyy"
    ActiveCell.Offset(2, 0).Value = oPar.FechaAnalisis
    ActiveCell.Offset(2, 0).NumberFormat = "dd/mmm/yyyy"
    oCelda.Offset(6, 1).Activate
    ActiveCell.Offset(0, 0).Value = oBola.Numero.Valor
    ActiveCell.Offset(1, 0).Value = oBola.Apariciones
    ActiveCell.Offset(2, 0).Value = oBola.Ausencias
    With ActiveCell.Offset(3, 0)
        .Value = oBola.Probabilidad
        .NumberFormat = "0.000%"
        .Interior.ColorIndex = oBola.Color_Probabilidad
    End With
    With ActiveCell.Offset(4, 0)
        .Value = oBola.Prob_TiempoMedio
        .NumberFormat = "0.000%"
        .Interior.ColorIndex = oBola.Color_Tiempo_Medio
    End With
    With ActiveCell.Offset(5, 0)
        .Value = oBola.Prob_Frecuencia
        .NumberFormat = "0.000%"
        .Interior.ColorIndex = oBola.Color_Frecuencias
    End With
    ActiveCell.Offset(6, 0).Value = oBola.Tiempo_Medio
    ActiveCell.Offset(6, 0).NumberFormat = "0.00"
    ActiveCell.Offset(7, 0).Value = oBola.Desviacion_Tm
    ActiveCell.Offset(8, 0).Value = oBola.Maximo_Tm
    ActiveCell.Offset(9, 0).Value = oBola.Minimo_Tm
    ActiveCell.Offset(10, 0).Value = oBola.Ultima_Fecha
    ActiveCell.Offset(10, 0).NumberFormat = "dd/mmm/yyyy"
    ActiveCell.Offset(11, 0).Value = oBola.ProximaFecha
    ActiveCell.Offset(11, 0).NumberFormat = "dd/mmm/yyyy"
    ActiveCell.Offset(12, 0).Value = oBola.Numero.Terminacion
    ActiveCell.Offset(13, 0).Value = oBola.Numero.Decena
    ActiveCell.Offset(14, 0).Value = oBola.Numero.Paridad
    ActiveCell.Offset(15, 0).Value = oBola.Numero.Peso
    ActiveCell.Offset(16, 0).Value = oBola.Tendencia

PrintBola_CleanExit:
   On Error GoTo 0
    Exit Sub

PrintBola_Error:
    

    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    '   Audita el error
    Call HandleException(ErrNumber, ErrDescription, "Lot_ProbTiemposMedios.PrintBola")
    '   Lanza el Error
    Err.Raise ErrNumber, ErrSource, ErrDescription
End Sub

' *============================================================================*
' *     Procedure  : PrintFechas
' *     Fichero    : Lot_ProbTiemposMedios
' *     Autor      : Carlos Almela Baeza
' *     Creacion   : mié, 30/11/2011
' *     Asunto     :
' *============================================================================*
'
Private Sub PrintFechas(oCelda As Range, _
                        oBola As bola, _
                        oPar As ParametrosMuestra)
    Dim xFila           As Integer
    Dim xCol            As Integer
    Dim sAddressX       As String
    Dim sAddressY       As String
    Dim Rg_Salida       As Range
    Dim txt_rango       As String
    Dim iColGraph       As Integer
    Dim iRowGraph       As Integer
    Dim Index           As Integer
    Dim nameObjet       As String
    Dim iNum            As Integer
    Dim j               As Integer
  On Error GoTo PrintFechas_Error
    
    oCelda.Offset(39, 0).Activate
    xFila = ActiveCell.Row
    xCol = ActiveCell.Column
    sAddressX = ActiveCell.Address
    For j = 0 To oBola.Apariciones - 1 Step 1
            Cells(xFila + j, xCol).Value = oBola.Matriz_Fechas(j)
            Cells(xFila + j, xCol).NumberFormat = "dd-mmm"
            If j > 0 Then
                Cells(xFila + j, xCol + 1).Value = oBola.Matriz_Apariciones(j - 1)
                sAddressY = Cells(xFila + j, xCol + 1).Address
            End If
    Next j
    '
    '
    '
    Cells(xFila + j, xCol).Value = oPar.FechaAnalisis
    Cells(xFila + j, xCol).NumberFormat = "dd-mmm"
    Cells(xFila + j, xCol + 1).Value = oBola.Ausencias
    sAddressY = Cells(xFila + j, xCol + 1).Address
    '
    '
    '
    'Insertar Gráfico
    iNum = oBola.Numero.Valor
    nameObjet = "Salida" & CStr(iNum)
    txt_rango = sAddressX & ":" & sAddressY
    Set Rg_Salida = Range(txt_rango)
    iColGraph = oCelda.Column + 4
    iRowGraph = oCelda.Row
    Rg_Salida.Select
    ActiveSheet.Shapes.AddChart.Select
    Index = ActiveSheet.Shapes.Count
    ActiveSheet.Shapes(Index).Select
    With ActiveChart
        .SetSourceData Source:=Rg_Salida, PlotBy:=xlColumns
        .ChartType = xlLineMarkers
        .HasTitle = True
        .ChartTitle.Characters.Text = "Apariciones Nº ( " & CStr(iNum) & " )"
        .Axes(xlCategory, xlPrimary).HasTitle = False
        .Axes(xlCategory, xlPrimary).TickLabels.NumberFormat = "[$-C0A]d-mmm;@"
        .Axes(xlValue, xlPrimary).HasTitle = False
        .HasLegend = False
        .HasDataTable = False
    End With
    ActiveSheet.Shapes(Index).Left = ActiveSheet.Columns(iColGraph).Left
    ActiveSheet.Shapes(Index).Top = ActiveSheet.Rows(iRowGraph).Top

PrintFechas_CleanExit:
   On Error GoTo 0
    Exit Sub

PrintFechas_Error:
    

    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    '   Audita el error
    Call HandleException(ErrNumber, ErrDescription, "Lot_ProbTiemposMedios.PrintFechas")
    '   Lanza el Error
    Err.Raise ErrNumber, ErrSource, ErrDescription
End Sub

' *============================================================================*
' *     Procedure  : PrintFrecuencias
' *     Fichero    : Lot_ProbTiemposMedios
' *     Autor      : Carlos Almela Baeza
' *     Creacion   : dom, 12/02/2012 23:29
' *     Asunto     : Dibuja la tabla de frecuencias de aparición del número
' *============================================================================*
'
Private Sub PrintFrecuencias(oCelda As Range, _
                             oBola As bola, _
                             oPar As ParametrosMuestra)
    Dim m_frec          As Variant
    Dim xFila           As Integer
    Dim xCol            As Integer
    Dim sAddressX       As String
    Dim sAddressY       As String
    Dim Rg_Salida       As Range
    Dim Rg_Values       As Range
    Dim txt_rango       As String
    Dim txt_rgFrec      As String
    Dim iColGraph       As Integer
    Dim iRowGraph       As Integer
    Dim Index           As Integer
    Dim nameObjet       As String
    Dim iNum            As Integer
    Dim j               As Integer
    
    
  On Error GoTo PrintFrecuencias_Error
    '
    '   Se desplaza el cursor (X=39, Y=3)
    '
    oCelda.Offset(39, 3).Activate
    '
    '   Se obtiene la fila de la celda activa
    '
    xFila = ActiveCell.Row
    '
    '   Se obtiene la columna de la celda activa
    '
    xCol = ActiveCell.Column
    '
    '
    '
    sAddressX = ActiveCell.Address
    '
    '   Se obtiene la matriz de frecuencias de la Bola
    '
    m_frec = oBola.Frecuencias
    '
    '   Bucle para cada frecuencia
    '
    For j = 0 To UBound(Rango_Frecuencias)
            '
            '   Desplazamiento 1 fila
            '
            Cells(xFila + j, xCol).Value = Rango_Frecuencias(j)
            Cells(xFila + j, xCol + 1).Value = m_frec(j + 1, 1)
    Next j
    '
    '   Calcula el rango de datos
    '
    j = UBound(Rango_Frecuencias)
    sAddressX = Cells(xFila, xCol + 1).Address
    sAddressY = Cells(xFila + j, xCol + 1).Address
    txt_rango = sAddressX & ":" & sAddressY
    Set Rg_Salida = Range(txt_rango)
    '
    '   Calcula el rango de valores
    '
    sAddressX = Cells(xFila, xCol).Address
    sAddressY = Cells(xFila + j, xCol).Address
    txt_rgFrec = sAddressX & ":" & sAddressY
    Set Rg_Values = Range(txt_rgFrec)
    '
    '   Posición del grafico
    '
    iColGraph = oCelda.Column + 4
    iRowGraph = oCelda.Row + 19
    
    '
    '   Obtiene el Numero y asigna nombre al gráfico
    '
    iNum = oBola.Numero.Valor
    nameObjet = "SalidaF" & CStr(iNum)
    '
    '
    '
    'Insertar Gráfico
    Rg_Salida.Select
    ActiveSheet.Shapes.AddChart.Select
    Index = ActiveSheet.Shapes.Count
    ActiveSheet.Shapes(Index).Select
    With ActiveChart
        .SetSourceData Source:=Rg_Salida, PlotBy:=xlColumns
        .SeriesCollection(1).XValues = Rg_Values
        .ChartType = xlColumnClustered
        .HasTitle = True
        .ChartTitle.Characters.Text = "Frecuencias Nº ( " & CStr(iNum) & " )"
        .Axes(xlCategory, xlPrimary).HasTitle = False
        .Axes(xlValue, xlPrimary).HasTitle = False
        .HasAxis(xlCategory, xlPrimary) = True
        .HasAxis(xlValue, xlPrimary) = True
        .Axes(xlValue).HasMajorGridlines = True
        .HasLegend = False
        .HasDataTable = False
    End With
    '
    '   Posiciona el gráfico debajo del anterior
    '
    ActiveSheet.Shapes(Index).Left = ActiveSheet.Columns(iColGraph).Left
    ActiveSheet.Shapes(Index).Top = ActiveSheet.Rows(iRowGraph).Top

PrintFrecuencias_CleanExit:
   On Error GoTo 0
    Exit Sub

PrintFrecuencias_Error:
    

    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    '   Audita el error
    Call HandleException(ErrNumber, ErrDescription, "Lot_ProbTiemposMedios.PrintFrecuencias")
    '   Lanza el Error
    Err.Raise ErrNumber, ErrSource, ErrDescription
End Sub

