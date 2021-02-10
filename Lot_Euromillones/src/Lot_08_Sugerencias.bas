Attribute VB_Name = "Lot_08_Sugerencias"
' *============================================================================*
' *
' *     Fichero    : Lot_08_Sugerencias.mod
' *
' *     Autor      : Carlos Almela Baeza
' *     Creación   : lun, 10/10/2011 23:26
' *     Modificado :
' *     Versión    : 1.1
' *     Propósito  : Generar un conjunto de sugerencias atendiendo a
' *                  los métodos seleccionados
' *
' *============================================================================*
Option Explicit
Option Base 0

'--- Variables Privadas -------------------------------------------------------*
Private mColId        As Collection          ' Colección de metodos Seleccionados
Private mModel        As MetodoModel         ' Modelo del objeto Metodo
Private mCtrl         As MetodoController    ' Controlador del objeto Metodo
Private mCalSuge      As RealizarSugerencia  ' Clase generadora de sugerencias
Private frmSugerencia As frmMetodoSelectView ' Formulario de selección
Private oSorteo       As Sorteo              ' Sorteo de la fecha de análisis
Private oEng          As SorteoEngine        ' Motor de gestión de Sorteos
Private i             As Integer             ' Contador
Private mCurrentId    As Integer             ' Id de método actual
Private mSuge         As Sugerencia          ' Sugerencia actual
Private mLinea        As Integer             ' linea de visualización sugerencia
Private mCurrentPage  As Integer             ' Contador de páginas
Private mMtdo         As Metodo              ' Current Metodo
Private mTotalPaginas As Integer             ' Total de páginas
'--- Constantes ---------------------------------------------------------------*
'--- Mensajes -----------------------------------------------------------------*
'--- Errores ------------------------------------------------------------------*
'--- Propiedades --------------------------------------------------------------*
'--- Métodos Privados ---------------------------------------------------------*

' *============================================================================*
' *     Procedure  : btn_SugerirApuestas
' *     Fichero    : Lot_Sugerencias
' *     Autor      : Carlos Almela Baeza
' *     Creacion   : lun, 10/10/2011
' *     Asunto     :
' *============================================================================*
'
Public Sub btn_SugerirApuestas()
  On Error GoTo btn_SugerirApuestas_Error
    '
    '   Borra el contenido de la hoja de salida
    '
    Borra_Salida
    Application.ScreenUpdating = False       'Desactiva el refresco de pantalla
    '
    '   Crea un formulario de captura de métodos
    '
    Set frmSugerencia = New frmMetodoSelectView
    '
    '   inicializa el estado
    '
    frmSugerencia.Tag = ESTADO_INICIAL
    '
    '   Selecciona parametros del proceso
    '
    Do While frmSugerencia.Tag <> BOTON_CERRAR
    
        ' Se inicializa el boton cerrar para salir del bucle
        frmSugerencia.Tag = BOTON_CERRAR
        
        ' Se muestra el formulario y queda a la espera de funciones
        ' pulsando el botón ejecutar
        frmSugerencia.Show
       
        'Se bifurca la función
        Select Case frmSugerencia.Tag
                                    ' El usuario ha cerrado el
            Case ""                 ' cuadro de dialogo con la [X]
                frmSugerencia.Tag = BOTON_CERRAR
            
            Case EJECUTAR           ' Se ha pulsado el botón ejecutar
                '
                '   Creamos  la colección de Metodos seleccionados
                '
                Set mColId = New Collection
                '
                '
                '
                If Not (frmSugerencia.AllMetodosSelected) Then
                    '
                    '   Establecemos los métodos seleccionados
                    '
                    Set mColId = frmSugerencia.SelectedIds
                End If
                '
                '   Invocamos a la rutina de visualización de resultados
                '
                VisualizarSugerencias mColId, _
                                      frmSugerencia.AllMetodosSelected, _
                                      frmSugerencia.FechaAnalisis, _
                                      frmSugerencia.Pronosticos
                '
                '   Cerramos el formulario
                '
                frmSugerencia.Tag = BOTON_CERRAR
        End Select
    Loop

    Application.ScreenUpdating = True       'Desactiva el refresco de pantalla
btn_SugerirApuestas_CleanExit:
   On Error GoTo 0
   Exit Sub

btn_SugerirApuestas_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    Call HandleException(ErrNumber, ErrDescription, "Lot_Sugerencias.btn_SugerirApuestas")
    Call MsgBox(ErrDescription, vbError Or vbSystemModal, NOMBRE_APLICACION)
    Call Trace("CERRAR")
End Sub





'------------------------------------------------------------------------------*
' Procedimiento  : VisualizarSugerencias
' Fecha          :
' Propósito      : Visualizar las sugerencias seleccionadas en la hoja Salida
' Parámetros     : Metodos seleccionados, Fecha Analisis, Indicador de todos
'                  los métodos
'------------------------------------------------------------------------------*
Private Sub VisualizarSugerencias(datSelMetodos As Collection, _
                                  datAllMetodos As Boolean, _
                                  datFechaAnalisis As Date, _
                                  datPronosticos As Integer)
    
  On Error GoTo VisualizarSugerencias_Error
    '
    '   Comprobar si hay sorteo para la fecha elegida
    '
    Set oEng = New SorteoEngine
    Set oSorteo = oEng.GetSorteoByFecha(datFechaAnalisis)
    '
    '   Preparar literales de salida
    '
    DisLiterales oSorteo, datFechaAnalisis
    '
    '   Establecemos los objetos necesarios para el proceso
    '
    Set mModel = New MetodoModel
    Set mCtrl = New MetodoController
    Set mCalSuge = New RealizarSugerencia
    '
    '   Inicializamos la linea de visualización de la sugerencia
    '
    mLinea = 0
    '
    '   Bucle para todos los métodos
    '
    If datSelMetodos.Count > 0 Then
        '
        '   Para cada método seleccionado
        '
        For i = 1 To datSelMetodos.Count
            '
            '   Obtenemos el Id actual
            '
            mCurrentId = datSelMetodos.Item(i)
            '
            '   Si encuentra el método
            '
            If mModel.GetRecord(mCurrentId) Then
                '
                '   Asignamos pronósticos
                '
                mModel.Metodo.Pronosticos = datPronosticos
                '
                '   Creamos una sugerencia
                '
                Set mSuge = mCalSuge.GetSugerencia(mModel.Metodo, datFechaAnalisis)
                '
                '   Si la sugerencia es válida
                '
                If Not (mSuge Is Nothing) Then
                    '
                    '   Visualizamos la sugerencia en la hoja de salids
                    '
                    DisSugerencia mSuge, oSorteo, mLinea
                    '
                    '   Incrementamos la linea
                    '
                    mLinea = mLinea + 1
                End If
            End If
        Next i
    Else
        '
        '
        '
        mCtrl.SetLinePerPage 10
        '
        '   Cargamos la primera página
        '
        Set mModel = mCtrl.GoPageNumber(1)
        '
        '   Si tenemos registros
        '
        If mModel.TotalRecords > 0 Then
            mTotalPaginas = mModel.TotalPages
            mCurrentPage = 1
            Do
                If mCurrentPage > 1 Then
                    Set mModel = mCtrl.GoPageNumber(mCurrentPage)
                End If
                For Each mMtdo In mModel.Metodos.Items
                    '
                    '   Asignamos pronósticos
                    '
                    mMtdo.Pronosticos = datPronosticos
                    '
                    '   Creamos una sugerencia
                    '
                    Set mSuge = mCalSuge.GetSugerencia(mMtdo, datFechaAnalisis)
                    '
                    '   Si la sugerencia es válida
                    '
                    If Not (mSuge Is Nothing) Then
                        '
                        '   Visualizamos la sugerencia en la hoja de salids
                        '
                        DisSugerencia mSuge, oSorteo, mLinea
                        '
                        '   Incrementamos la linea
                        '
                        mLinea = mLinea + 1
                    End If
                Next mMtdo
                mCurrentPage = mCurrentPage + 1
            Loop Until mCurrentPage > mTotalPaginas
        End If
    End If

VisualizarSugerencias_CleanExit:
   On Error GoTo 0
   Exit Sub

VisualizarSugerencias_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    Call HandleException(ErrNumber, ErrDescription, "Lot_Sugerencias.VisualizarSugerencias")
    Err.Raise ErrNumber, "Lot_Sugerencias.VisualizarSugerencias", ErrDescription
End Sub




'------------------------------------------------------------------------------*
' Procedimiento  : DisLiterales
' Fecha          : mi., 13/may/2020 19:23:05
' Propósito      : Configura la hoja para la información de Salida del proceso
'------------------------------------------------------------------------------*
Private Sub DisLiterales(datSorteo As Sorteo, datFechaAnalisis As Date)
    Dim c As Integer
    Dim n As Integer
    Dim mVar As Variant
    
  On Error GoTo DisLiterales_Error
    '
    '   Titulo
    '
    Range("A1").Activate
    ActiveCell.Value = "Sugerencias"
    '
    '   Detalles sorteo
    '
    Range("A3").Activate
    ActiveCell.Value = "Fecha de Análisis"
    ActiveCell.Offset(0, 1).Value = datFechaAnalisis
    ActiveCell.Offset(0, 1).NumberFormat = FMT_FECHA
    
    If datSorteo Is Nothing Then
        ActiveCell.Offset(1, 0).Value = "Numero de Sorteo"
        ActiveCell.Offset(1, 1).Value = "Sin Sorteo"
    Else
        ActiveCell.Offset(1, 0).Value = "Numero de Sorteo"
        ActiveCell.Offset(1, 1).Value = datSorteo.NumeroSorteo
        ActiveCell.Offset(2, 0).Value = "Concurso"
        ActiveCell.Offset(2, 1).Value = IIf(datSorteo.Juego = Bonoloto, "BL", "LP")
        ActiveCell.Offset(3, 0).Value = "Dia"
        ActiveCell.Offset(3, 1).Value = datSorteo.Dia
        ActiveCell.Offset(4, 0).Value = "Semana"
        ActiveCell.Offset(4, 1).Value = datSorteo.Semana
    End If
    '
    '   Combinación
    '
    If Not (datSorteo Is Nothing) Then
        Range("D2").Activate
        ActiveCell.Offset(1, 0).Value = "Combinación:"
        c = datSorteo.Complementario
        For i = 1 To datSorteo.Combinacion.Count
            n = datSorteo.Combinacion.Numeros(i).Valor
            If n <> c Then
                ActiveCell.Offset(0, i).Value = "N" & i
                ActiveCell.Offset(1, i).Value = n
                ActiveCell.Offset(1, i).Interior.ColorIndex = COLOR_VERDE
            End If
        Next i
        If datSorteo.Juego = Bonoloto Or _
        datSorteo.Juego = LoteriaPrimitiva Then
            ActiveCell.Offset(0, i).Value = "C"
            ActiveCell.Offset(1, i).Value = c
            ActiveCell.Offset(1, i).Interior.ColorIndex = COLOR_AMARILLO
        End If
    End If
    '
    '   Metodo
    '
    Range("E5").Activate
    mVar = Split("Id;N1;N2;N3;N4;N5;N6;N7;N8;N9;N10;N11;Aciertos;Parametros", ";")
    For i = 0 To UBound(mVar)
        ActiveCell.Offset(0, i).Value = mVar(i)
        ActiveCell.Offset(0, i).Font.Bold = True
    Next i
    Range("D6").Activate
    ActiveCell.Value = "Metodos"
    '
    '   Colores y Ajustes
    '
    '
    '   Titulo
    '
    Range("A1:R1").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
        .Style = "Título"
    End With
    Selection.Merge
    '
    '   Detalles proceso
    '
    If Not (datSorteo Is Nothing) Then
        '
        '   Label Detalles: Purpura, Enfasis 4, Claro 60%
        '
        With Range("A3:A7").Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .ThemeColor = xlThemeColorAccent4
            .TintAndShade = 0.599993896298105
            .PatternTintAndShade = 0
        End With
        '
        '  Text Detalles:  Purpura, Enfasis 4, Claro 80%
        '
        With Range("B3:B7").Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .ThemeColor = xlThemeColorAccent4
            .TintAndShade = 0.799981688894314
            .PatternTintAndShade = 0
        End With
        '
        '   Label Sorteo: Purpura, Enfasis 4, Claro 60%
        '
        With Range("D2:K2").Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .ThemeColor = xlThemeColorAccent4
            .TintAndShade = 0.599993896298105
            .PatternTintAndShade = 0
        End With
        '
        '   Text Sorteo: Purpura, Enfasis 4, Claro 80%
        '
        With Range("D3").Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .ThemeColor = xlThemeColorAccent4
            .TintAndShade = 0.799981688894314
            .PatternTintAndShade = 0
        End With
    Else
        '
        '   Label Detalles: Purpura, Enfasis 4, Claro 60%
        '
        With Range("A3:A4").Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .ThemeColor = xlThemeColorAccent4
            .TintAndShade = 0.599993896298105
            .PatternTintAndShade = 0
        End With
        '
        '   Text Detalles:  Purpura, Enfasis 4, Claro 80%
        '
        With Range("B3:B4").Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .ThemeColor = xlThemeColorAccent4
            .TintAndShade = 0.799981688894314
            .PatternTintAndShade = 0
        End With
    End If
    '
    '   Cabecera Metodo
    '
    With Range("D5:R5").Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent2
        .TintAndShade = 0.599993896298105
        .PatternTintAndShade = 0
    End With
    With Range("D6").Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent2
        .TintAndShade = 0.599993896298105
        .PatternTintAndShade = 0
    End With
    '
    '   Autoajustar celdas
    '
    Cells.Select                            'Selecciona todas las celdas de la hoja
    Cells.EntireColumn.AutoFit              'Autoajusta el tamaño de las columnas
    Range("A1").Select                      'Se posiciona el cursor en la celda A1
    
DisLiterales_CleanExit:
    On Error GoTo 0
    Exit Sub
DisLiterales_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    Call HandleException(ErrNumber, ErrDescription, "Lot_Sugerencias.DisLiterales")
    Err.Raise ErrNumber, "Lot_Sugerencias.DisLiterales", ErrDescription
End Sub

'------------------------------------------------------------------------------*
' Procedimiento  : DisLiterales
' Fecha          : mi., 13/may/2020 19:23:05
' Propósito      : Configura la hoja para la información de Salida del proceso
'------------------------------------------------------------------------------*
Private Sub DisSugerencia(datSuge As Sugerencia, _
                          datSorteo As Sorteo, _
                          datLinea As Integer)
    Dim j           As Integer
    Dim n           As Integer
    Dim mNum        As Numero
    Dim mCuValidar  As CU_ValidarSugerencia
  
  On Error GoTo DisSugerencia_Error
    '
    '   Esquina superior izquierda del listado
    '
    Range("E6").Activate
    '
    '
    ActiveCell.Offset(datLinea, 0).Value = "#" & CStr(datSuge.Metodo.Id)
    With ActiveCell.Offset(datLinea, 0).Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent2
        .TintAndShade = 0.799981688894314
        .PatternTintAndShade = 0
    End With
    '
    '   Ordenamos la Suge
    '
    datSuge.Sort
    '
    '
    '
    If Not (datSorteo Is Nothing) Then
        '
        '   Bucle combinación
        '
        For j = 1 To datSuge.Combinacion.Count
            Set mNum = datSuge.Combinacion.Numeros(j)
            n = mNum.Valor
            ActiveCell.Offset(datLinea, j).Value = n
            If datSorteo.Complementario = mNum.Valor Then
                ActiveCell.Offset(datLinea, j).Interior.ColorIndex = COLOR_AMARILLO
            ElseIf datSorteo.Combinacion.Contiene(n) Then
                ActiveCell.Offset(datLinea, j).Interior.ColorIndex = COLOR_VERDE_CLARO
            End If
        Next j
        '
        '   Comprobamos la sugerencia
        '
        Set mCuValidar = New CU_ValidarSugerencia
        ActiveCell.Offset(datLinea, 12).Value = mCuValidar.GetPremio(datSuge, datSorteo)
        With ActiveCell.Offset(datLinea, 12).Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .ThemeColor = xlThemeColorAccent2
            .TintAndShade = 0.799981688894314
            .PatternTintAndShade = 0
        End With
    Else
        For j = 1 To datSuge.Combinacion.Count
            Set mNum = datSuge.Combinacion.Numeros(j)
            n = mNum.Valor
            ActiveCell.Offset(datLinea, j).Value = n
        Next j
    End If
    '
    '   Metodo aplicado
    '
    ActiveCell.Offset(datLinea, 13).Value = datSuge.Metodo.ToString
DisSugerencia_CleanExit:
    On Error GoTo 0
    Exit Sub
DisSugerencia_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    Call HandleException(ErrNumber, ErrDescription, "Lot_Sugerencias.DisSugerencia")
    Err.Raise ErrNumber, "Lot_Sugerencias.DisSugerencia", ErrDescription
End Sub

' *===========(EOF): Lot_08_Sugerencias.mod
