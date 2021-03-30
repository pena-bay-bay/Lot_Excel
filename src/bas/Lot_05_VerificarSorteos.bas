Attribute VB_Name = "Lot_05_VerificarSorteos"
' *============================================================================*
' *
' *     Fichero    : Lot_05_VerificarSorteos.bas
' *
' *     Autor      : Carlos Almela Baeza
' *     Creación   : mié, 17/sep/2014 00:01:44
' *     Revisión   : mi., 28/oct/2020 20:12:29
' *     Versión    : 1.0
' *     Propósito  : Proceso de análisis de los sorteos aparecidos en un
' *                  período de tiempo.
' *                  Se visualiza la probabilidad de los Numeros y las
' *                  caracteristicas de la combinación (paridad, Terminaciones,
' *                  consecutivos, etc)
' *============================================================================*

Option Explicit
Option Base 0


'---------------------------------------------------------------------------------------
' Procedure : btn_VerificarSorteos
' Author    : CHARLY
' Date      : mié, 17/sep/2014 00:04:57
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Sub btn_VerificarSorteos()
    Dim oFrm        As frmSelPeriodo
  
   On Error GoTo btn_VerificarSorteos_Error
    '
    '   Desactiva la presentación
    '
    CALCULOOFF
    '
    '   Creamos el formulario del periodo de tiempo
    '
    Set oFrm = New frmSelPeriodo
    '
    '   Localizar criterios de consulta
    '
    oFrm.Tag = ESTADO_INICIAL
    '
    '   Bucle de control del proceso
    '
    Do While oFrm.Tag <> BOTON_CERRAR
        '
        ' Se inicializa el boton cerrar para salir del bucle
        oFrm.Tag = BOTON_CERRAR
        '
        ' Se muestra el formulario y queda a la espera de funciones
        ' pulsando el botón ejecutar
        oFrm.Show vbModal
        '
        'Se bifurca la función
        Select Case oFrm.Tag
                                    ' El usuario ha cerrado el
            Case ""                 ' cuadro de dialogo con la [X]
                oFrm.Tag = BOTON_CERRAR
            '
            '   Selecciona el botón [Aceptar]
            '
            Case EJECUTAR
                VerificarSorteo oFrm.Periodo
                oFrm.Tag = BOTON_CERRAR
        End Select
    Loop
    '
    '  Se elimina de la memoria el formulario
    '
    Set oFrm = Nothing
    '
    '
    '
    Ir_A_Hoja ("Salida")
    '
    '  Activa la presentación
    '
    CALCULOON
            
   On Error GoTo 0
    Exit Sub
            
btn_VerificarSorteos_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    Call HandleException(ErrNumber, ErrDescription, "Lot_02_VerificarSorteos.btn_VerificarSorteos", ErrSource)
    Call MsgBox(ErrDescription, vbError Or vbSystemModal, NOMBRE_APLICACION)
    Call Trace("CERRAR")
End Sub


'---------------------------------------------------------------------------------------
' Procedure : VerificarSorteo
' Author    : CHARLY
' Date      : mié, 17/sep/2014 00:09:12
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub VerificarSorteo(vNewValue As Periodo)
    Dim oParMuestra     As ParametrosMuestra
    Dim oMuestra        As Muestra
    Dim oBola           As Bola
    Dim oInfo           As InfoSorteo
    Dim oSorteo         As Sorteo
    Dim oNum            As Numero
    Dim rgFila          As Range
    Dim rgDatos         As Range
    Dim mDb             As New BdDatos
    Dim i               As Integer
    Dim j               As Integer
   On Error GoTo VerificarSorteo_Error
    '
    '   Borra la hoja de salida
    '
    Borra_Salida
    '
    '   Escribe los textos de la salida
    '
    PonTextosCabecera
    '
    '   Calcula la muestra estadistica utilizando
    '   Fecha Inicial del periodo como Fecha análisis
    '
    Set oParMuestra = New ParametrosMuestra
    Set oInfo = New InfoSorteo
    With oParMuestra
        .Juego = JUEGO_DEFECTO
        .TipoMuestra = True
        .FechaAnalisis = vNewValue.FechaInicial
        .FechaFinal = oInfo.GetAnteriorSorteo(vNewValue.FechaInicial)
        .NumeroSorteos = 90
    End With
    '
    '   Visualiza los valores del proceso
    '
    PonValoresCabecera vNewValue, oParMuestra
    '
    '
    '   obtiene el rango con los datos comprendido entre las dos fechas
    '
    Set rgDatos = mDb.GetSorteosInFechas(vNewValue)
    'Nos posicionamos en la celda de inicio de escritura
    Range("D3").Activate                        'Se posiciona el cursor en la
                                                'celda D3

    i = 0
    Set oSorteo = New Sorteo
    For Each rgFila In rgDatos.Rows
        
        'Componemos el resultado
        oSorteo.Constructor rgFila
        '
        '   Obtenemos la muestra para la fecha de sorteo
        '
        With oParMuestra
            .FechaAnalisis = oSorteo.Fecha
            .FechaFinal = oInfo.GetAnteriorSorteo(oSorteo.Fecha)
            .NumeroSorteos = 90
        End With
        '
        '   Calculamos la estadistica
        '
        Set oMuestra = GetMuestra(oParMuestra)
        '
        '   Escribimos el resultado
        '
        ActiveCell.Offset(i, 0).Value = oSorteo.Fecha
        ActiveCell.Offset(i, 1).Value = oSorteo.Semana
        
        
        Select Case JUEGO_DEFECTO
            Case Bonoloto, LoteriaPrimitiva
                '
                '   Coloreamos la Combinación ganadora con la estadistica del muestreo
                '
                For j = 1 To 6
                    Set oNum = oSorteo.Combinacion.Numeros(j)
                    Set oBola = oMuestra.Get_Bola(oNum.Valor)
                    With ActiveCell.Offset(i, j + 1)
                        .Value = oNum.Valor
                        .NumberFormat = "00"
                        .Interior.ColorIndex = oBola.Color_Probabilidad
                    End With
                Next j
                j = j + 1
                oNum.Valor = oSorteo.Complementario
                Set oBola = oMuestra.Get_Bola(oNum.Valor)
                With ActiveCell.Offset(i, j)
                    .Value = oNum.Valor
                    .NumberFormat = "00"
                    .Interior.ColorIndex = oBola.Color_Probabilidad
                End With
                j = j + 1
                ActiveCell.Offset(i, j) = oSorteo.Reintegro
                j = j + 1
            
            Case GordoPrimitiva
                For j = 1 To 5
                    Set oNum = oSorteo.Combinacion.Numeros(j)
                    Set oBola = oMuestra.Get_Bola(oNum.Valor)
                    With ActiveCell.Offset(i, j + 1)
                        .Value = oNum.Valor
                        .NumberFormat = "00"
                        .Interior.ColorIndex = oBola.Color_Probabilidad
                    End With
                Next j
                ' #TODO: Obtener la estadistica de reintegros
                j = j + 1
                ActiveCell.Offset(i, j) = oSorteo.Reintegro
                j = j + 1
            
            Case Euromillones
                For j = 1 To 5
                    Set oNum = oSorteo.Combinacion.Numeros(j)
                    Set oBola = oMuestra.Get_Bola(oNum.Valor)
                    With ActiveCell.Offset(i, j + 1)
                        .Value = oNum.Valor
                        .NumberFormat = "00"
                        .Interior.ColorIndex = oBola.Color_Probabilidad
                    End With
                Next j
                For j = 1 To 2
                    Set oNum = oSorteo.Estrellas.Numeros(j)
'                     #TODO: Obtener la bola de la estadistica de estrellas
'                    Set oBola = oMuestra.Get_Bola(oNum.Valor)
                    With ActiveCell.Offset(i, j + 6)
                        .Value = oNum.Valor
                        .NumberFormat = "00"
'                     #TODO: Obtener el color de la bola de la estadistica de estrellas
'                        .Interior.ColorIndex = oBola.Color_Probabilidad
                    End With
                Next j
        End Select
        '
        '   Escribimos caracteristicas del sorteo
        '
        ActiveCell.Offset(i, j + 1).Value = "'" & oSorteo.Combinacion.FormulaParidad
        ActiveCell.Offset(i, j + 2).Value = "'" & oSorteo.Combinacion.FormulaAltoBajo
        ActiveCell.Offset(i, j + 3).Value = "'" & oSorteo.Combinacion.FormulaDecenas
        ActiveCell.Offset(i, j + 4).Value = "'" & oSorteo.Combinacion.FormulaSeptenas
        ActiveCell.Offset(i, j + 5).Value = "'" & oSorteo.Combinacion.FormulaTerminaciones
        ActiveCell.Offset(i, j + 6).Value = "'" & oSorteo.Combinacion.FormulaConsecutivos
        ActiveCell.Offset(i, j + 7).Value = oSorteo.Combinacion.Suma
                
        i = i + 1                           ' Incrementa la Fila
    Next rgFila

    Cells.Select                            'Selecciona todas las celdas de la hoja
    Cells.EntireColumn.AutoFit              'Autoajusta el tamaño de las columnas
    
    Range("A1").Activate                    'Se posiciona el cursor en la celda A1
   On Error GoTo 0
    Exit Sub
            
VerificarSorteo_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    Call HandleException(ErrNumber, ErrDescription, "Lot_02_VerificarSorteos.VerificarSorteo", ErrSource)
    Err.Raise ErrNumber, "Lot_02_VerificarSorteos.VerificarSorteo", ErrDescription
End Sub




'---------------------------------------------------------------------------------------
' Procedimiento : PonTextosCabecera
' Creación      : 12-dic-2006 23:45
' Autor         : Carlos Almela Baeza
' Objeto        : Imprime los textos de la hoja de salida
'---------------------------------------------------------------------------------------
'
Private Sub PonTextosCabecera()
    Dim sRango As String
    Dim sRangoA As String
    
    Range("A1").Activate
    ActiveCell.Value = "Comprobacion de resultados"
    ActiveCell.Font.Bold = True
    ActiveCell.Offset(1, 0).Value = "Fecha Final"
    ActiveCell.Offset(2, 0).Value = "Fecha Inicial"
    ActiveCell.Offset(4, 0).Value = "Fecha Analisis"
    ActiveCell.Offset(5, 0).Value = "Fin Muestra"
    ActiveCell.Offset(6, 0).Value = "Inicio Muestra"
    ActiveCell.Offset(7, 0).Value = "Dias Analizados"
    ActiveCell.Offset(8, 0).Value = "Numero de Sorteos "

    
'*------------------------
    Range("D1").Activate
    ActiveCell.Value = "Resultados"
    
    Select Case JUEGO_DEFECTO
        Case Bonoloto, LoteriaPrimitiva:
            ActiveCell.Offset(1, 0).Value = "Fecha"
            ActiveCell.Offset(1, 1).Value = "Sem"
            ActiveCell.Offset(1, 2).Value = "N1"
            ActiveCell.Offset(1, 3).Value = "N2"
            ActiveCell.Offset(1, 4).Value = "N3"
            ActiveCell.Offset(1, 5).Value = "N4"
            ActiveCell.Offset(1, 6).Value = "N5"
            ActiveCell.Offset(1, 7).Value = "N6"
            ActiveCell.Offset(1, 8).Value = "C"
            ActiveCell.Offset(1, 9).Value = "R"
            ActiveCell.Offset(1, 10).Value = "_"
            Range("D1:N1").Select
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
            End With
            Selection.Merge
            sRango = "O1"
            sRangoA = "O1:U1"
            
        Case Euromillones:
            ActiveCell.Offset(1, 0).Value = "Fecha"
            ActiveCell.Offset(1, 1).Value = "Sem"
            ActiveCell.Offset(1, 2).Value = "N1"
            ActiveCell.Offset(1, 3).Value = "N2"
            ActiveCell.Offset(1, 4).Value = "N3"
            ActiveCell.Offset(1, 5).Value = "N4"
            ActiveCell.Offset(1, 6).Value = "N5"
            ActiveCell.Offset(1, 7).Value = "E1"
            ActiveCell.Offset(1, 8).Value = "E2"
            ActiveCell.Offset(1, 9).Value = "_"
            Range("D1:M1").Select
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
            End With
            Selection.Merge
            sRango = "N1"
            sRangoA = "N1:T1"

        
        Case GordoPrimitiva:
            ActiveCell.Offset(1, 0).Value = "Fecha"
            ActiveCell.Offset(1, 1).Value = "Sem"
            ActiveCell.Offset(1, 2).Value = "N1"
            ActiveCell.Offset(1, 3).Value = "N2"
            ActiveCell.Offset(1, 4).Value = "N3"
            ActiveCell.Offset(1, 5).Value = "N4"
            ActiveCell.Offset(1, 6).Value = "N5"
            ActiveCell.Offset(1, 7).Value = "R"
            ActiveCell.Offset(1, 8).Value = "_"
            Range("D1:L1").Select
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
            End With
            Selection.Merge
            sRango = "M1"
            sRangoA = "M1:S1"
            
    End Select
    
'*------------------------
    Range(sRango).Activate
    ActiveCell.Value = "Formulas Combinacion"
    ActiveCell.Offset(1, 0).Value = "Paridad"
    ActiveCell.Offset(1, 1).Value = "Peso"
    ActiveCell.Offset(1, 2).Value = "Decena"
    ActiveCell.Offset(1, 3).Value = "Septena"
    ActiveCell.Offset(1, 4).Value = "Terminaciones"
    ActiveCell.Offset(1, 5).Value = "Consecutivos"
    ActiveCell.Offset(1, 6).Value = "Suma"

'*-----------------------|
    Range(sRangoA).Select
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
    End With
    Selection.Merge

End Sub


Private Sub PonValoresCabecera(vNewPeriodo As Periodo, vNewParam As ParametrosMuestra)
    Range("b2").Activate
    ActiveCell.Value = vNewPeriodo.FechaFinal
    ActiveCell.Offset(1, 0).Value = vNewPeriodo.FechaInicial
    ActiveCell.Offset(3, 0).Value = vNewParam.FechaAnalisis
    ActiveCell.Offset(4, 0).Value = vNewParam.FechaFinal
    ActiveCell.Offset(5, 0).Value = vNewParam.FechaInicial
    ActiveCell.Offset(6, 0).Value = vNewParam.DiasAnalisis
    ActiveCell.Offset(7, 0).Value = vNewParam.NumeroSorteos
End Sub


Private Function GetMuestra(vNewValue As ParametrosMuestra) As Muestra
    Dim oMuestra        As Muestra
    Dim rgDatos         As Range
    Dim mDb             As New BdDatos
  On Error GoTo GetMuestra_Error
    '
    '   Calcula la muestra para colorear los números
    '
    '   obtiene el rango con los datos comprendido entre las dos fechas
    '
    Set rgDatos = mDb.GetSorteosInFechas(vNewValue.PeriodoDatos)
                                        
    '
    '   se lo pasa al constructor de la clase y obtiene las estadisticas para cada bola
    '
    Set oMuestra = New Muestra
    Set oMuestra.ParametrosMuestra = vNewValue
    
    Select Case JUEGO_DEFECTO
        Case LoteriaPrimitiva, Bonoloto:
            oMuestra.Constructor rgDatos, ModalidadJuego.LP_LB_6_49
        
        Case GordoPrimitiva:
            oMuestra.Constructor rgDatos, ModalidadJuego.GP_5_54
        
        Case Euromillones:
            oMuestra.Constructor rgDatos, ModalidadJuego.EU_5_50
    End Select
    '
    '   Devolvemos la muestra
    '
    Set GetMuestra = oMuestra
    
  On Error GoTo 0
GetMuestra__CleanExit:
    Exit Function
    
GetMuestra_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    Call HandleException(ErrNumber, ErrDescription, "Lot_05_VerificarSorteos.GetMuestra", ErrSource)
    Err.Raise ErrNumber, "Lot_05_VerificarSorteos.GetMuestra", ErrDescription
End Function

' *===========(EOF): Lot_05_VerificarSorteos.bas
