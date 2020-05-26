Attribute VB_Name = "Lot_05_VerificarSorteos"
'---------------------------------------------------------------------------------------
' Module    : Lot_02_VerificarSorteos
' Author    : CHARLY
' Date      : mié, 17/sep/2014 00:01:44
' Purpose   : Proceso de análisis de los sorteos aparecidos en un período de tiempo
'             Se visualiza la probabilidad de los Numeros y las caracteristicas de la
'             combinación (paridad, Terminaciones, consecutivos, etc)
'---------------------------------------------------------------------------------------
'
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
            
                
            
   On Error GoTo 0
       Exit Sub
            
btn_VerificarSorteos_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    Call HandleException(ErrNumber, ErrDescription, "Lot_02_VerificarSorteos.btn_VerificarSorteos", ErrSource)
   '   Informa del error
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
    Dim oBola           As bola
    Dim oInfo           As InfoSorteo
    Dim oSorteo         As Sorteo
    Dim oNum            As Numero
    Dim rgDatos         As Range
    Dim rgFila          As Range
    Dim mDB             As New BdDatos
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
        .FechaAnalisis = vNewValue.FechaInicial
        .FechaFinal = oInfo.GetAnteriorSorteo(vNewValue.FechaInicial)
        .NumeroSorteos = 100
    End With
    '
    '   Visualiza los valores del proceso
    '
    PonValoresCabecera vNewValue, oParMuestra
    '
    '   Calcula la muestra para colorear los números
    '
    '   obtiene el rango con los datos comprendido entre las dos fechas
    '
    Set rgDatos = mDB.Resultados_Fechas(oParMuestra.FechaInicial, _
                                        oParMuestra.FechaFinal)
    '
    '   se lo pasa al constructor de la clase y obtiene las estadisticas para cada bola
    '
    Set oMuestra = New Muestra
    Set oMuestra.ParametrosMuestra = oParMuestra
    oMuestra.Constructor rgDatos, JUEGO_DEFECTO
    '
    '   Comprueba que las fechas sean de sorteo
    '
    If Not oInfo.EsFechaSorteo(vNewValue.FechaInicial) Then
        vNewValue.FechaInicial = oInfo.GetAnteriorSorteo(vNewValue.FechaInicial)
    End If
    If Not oInfo.EsFechaSorteo(vNewValue.FechaFinal) Then
        vNewValue.FechaFinal = oInfo.GetAnteriorSorteo(vNewValue.FechaFinal)
    End If
    '
    '
    '   obtiene el rango con los datos comprendido entre las dos fechas
    '
    Set rgDatos = mDB.Resultados_Fechas(vNewValue.FechaInicial, vNewValue.FechaFinal)
        
    'Nos posicionamos en la celda de inicio de escritura
    Range("D3").Activate                        'Se posiciona el cursor en la
                                                'celda D3

    i = 0
    Set oSorteo = New Sorteo
    For Each rgFila In rgDatos.Rows
    
        'Componemos el resultado
        oSorteo.Constructor rgFila
        '
        '   Escribimos el resultado
        '
        ActiveCell.Offset(i, 0).Value = oSorteo.Fecha
        ActiveCell.Offset(i, 1).Value = oSorteo.Dia
            
        'Coloreamos la Combinación ganadora con la estadistica del muestreo
        For j = 1 To 7
            Set oNum = oSorteo.Combinacion.Numeros(j)
            Set oBola = oMuestra.Get_Bola(oNum.Valor)
            With ActiveCell.Offset(i, j + 1)
                .Value = oNum.Valor
                .NumberFormat = "00"
                .Interior.ColorIndex = oBola.Color_Probabilidad
            End With
        Next j
        '
        '   Escribimos caracteristicas del sorteo
        '
        ActiveCell.Offset(i, 10).Value = "'" & oSorteo.Combinacion.FormulaParidad
        ActiveCell.Offset(i, 11).Value = "'" & oSorteo.Combinacion.FormulaAltoBajo
        ActiveCell.Offset(i, 12).Value = "'" & oSorteo.Combinacion.FormulaDecenas
        ActiveCell.Offset(i, 13).Value = "'" & oSorteo.Combinacion.FormulaTerminaciones
        ActiveCell.Offset(i, 14).Value = "'" & oSorteo.Combinacion.FormulaConsecutivos
        ActiveCell.Offset(i, 15).Value = oSorteo.Combinacion.Suma
        ActiveCell.Offset(i, 16).Value = oSorteo.Combinacion.Producto
                
        i = i + 1                           ' Incrementa la Fila
    Next rgFila

    Cells.Select                            'Selecciona todas las celdas de la hoja
    Cells.EntireColumn.AutoFit              'Autoajusta el tamaño de las columnas
    
    Range("A1").Activate                    'Se posiciona el cursor en la celda A1
    'PonCabecera
    '
    '  Obtiene una coleccion de sorteos para el periodo
    '  Obtiene los parametros de la muestra de evaluacion
    '  Obtiene la muestra para estos parametros
    '  Prepara la hoja de salida.
    '        Borrado y literales
    '  Para cada sorteo en la coleccion
    '     pinta los Numeros y colorea segun la
    '     pinta los datos de la combinación
    '     agrega a la estadistica los datos de la combinación
    '  siguiente sorteo
    '  Pinta la estadistica de combinaciones
    '       elemento n/total combinaciones
           
   On Error GoTo 0
       Exit Sub
            
VerificarSorteo_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    Call HandleException(ErrNumber, ErrDescription, "Lot_02_VerificarSorteos.VerificarSorteo", ErrSource)
    '   Lanza el error
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
'
' TODO: visualizar segun tipo de juego
    ActiveCell.Value = "Resultados"
    ActiveCell.Offset(1, 0).Value = "Fecha"
    ActiveCell.Offset(1, 1).Value = "Sem"
    ActiveCell.Offset(1, 2).Value = "N1"
    ActiveCell.Offset(1, 3).Value = "N2"
    ActiveCell.Offset(1, 4).Value = "N3"
    ActiveCell.Offset(1, 5).Value = "N4"
    ActiveCell.Offset(1, 6).Value = "N5"
    ActiveCell.Offset(1, 7).Value = "N6"
    ActiveCell.Offset(1, 8).Value = "C"
    ActiveCell.Offset(1, 9).Value = "_"
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

'*------------------------
    Range("N1").Activate
    ActiveCell.Value = "Formulas Combinacion"
    ActiveCell.Offset(1, 0).Value = "Paridad"
    ActiveCell.Offset(1, 1).Value = "Peso"
    ActiveCell.Offset(1, 2).Value = "Decena"
    ActiveCell.Offset(1, 3).Value = "Terminaciones"
    ActiveCell.Offset(1, 4).Value = "Consecutivos"
    ActiveCell.Offset(1, 5).Value = "Suma"
    ActiveCell.Offset(1, 6).Value = "Producto"

'*-----------------------|
    Range("N1:Y1").Select
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
