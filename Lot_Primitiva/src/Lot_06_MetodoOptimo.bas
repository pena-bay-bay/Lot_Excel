Attribute VB_Name = "Lot_06_MetodoOptimo"
' *============================================================================*
' *
' *     Fichero    : Lot_MetodoOptimo
' *
' *     Tipo       : Módulo
' *     Autor      : CAB3780Y
' *     Creacion   : jue, 08/05/2008  15:03
' *     Version    : 1.0
' *     Asunto     : Simulación de Métodos hasta obtener los parámetros del
' *                  método que obtienen mejor resultado
' *============================================================================*
Option Explicit
Private DB As New BdDatos

Type rango
    ElementoInferior As Integer
    ElementoSuperior As Integer
End Type

Type IntervaloTiempo
    FechaInicial As Date
    FechaFinal As Date
End Type
'
'
'
Public Sub btn_Metodo_Optimo()

On Error GoTo btn_Metodo_Optimo_Error

    Borra_Salida                    'Borra el contenido de la hoja de salida
    
    Pintar_Literales_01             'Rellena Literales
    
    Set frmMetodoOptimo = New frmMetodoOptimo 'Crea un formulario de captura de métodos
    
    frmMetodoOptimo.Tag = ESTADO_INICIAL 'inicializa el estado
    
    Do While frmMetodoOptimo.Tag <> BOTON_CERRAR
    
        ' Se inicializa el boton cerrar para salir del bucle
        frmMetodoOptimo.Tag = BOTON_CERRAR
        
        ' Se muestra el formulario y queda a la espera de funciones
        ' pulsando el botón ejecutar
        frmMetodoOptimo.Show
       
        'Se bifurca la función
        Select Case frmMetodoOptimo.Tag
                                    ' El usuario ha cerrado el
            Case ""                 ' cuadro de dialogo con la [X]
                frmMetodoOptimo.Tag = BOTON_CERRAR
            
            Case EJECUTAR               ' Se ha pulsado el botón ejecutar
                 Borra_Salida           ' Borra el contenido de la hoja de salida
                 Pintar_Literales_01    ' Rellena Literales

                 cmd_MetodoOptimo frmMetodoOptimo.Parametros
                               
            'Case SUGERIR           ' Propone una apuesta para los diversos métodos
                   
        End Select
    Loop

On Error GoTo 0
   Exit Sub

btn_Metodo_Optimo_Error:
    Set frmMetodoOptimo = Nothing
    MsgBox "Error " & Err.Number & " (" & Err.Description & _
    ") in procedure btn_Metodo_Optimo of Módulo Lot_MetodoOptimo"

End Sub



'Objetivo final
'   Fecha, {sorteo} {apuesta} Aciertos, parametroA, parametroB, ParametroC
'
'   Método de la Decena
'        Rango de Estadística { n dias [7..n] para la muestra} ó [n..m] tal que N > M
'        Rango de retardo     { 0..15 dias ó de 0 a n}
'        Rango de prueba      { Fecha fin - dias prueba}
'
'Para cada dia desde Fecha ini hasta Fecha Fin
'     Obtener día del sorteo
'     iniciar parámetros Apuesta guardada, metodo apuesta
'     para Rango días de muestra desde n hasta m
'         Para gango de Retraso de dias desde 0 hasta n
'             Obtener Muestra
'             Obtener Apuesta
'             Enfrentar Sorteo a apuesta
'             Si Aciertos apuesta > Aciertos apuesta guardada
'                Apuesta guardada = Apuesta
'                Metodo guardado = parámetros Apuesta
'             Fin Si
'        Siguiente Retraso
'    Siguiente IntervaloMuestra
'    Pintar Sorteo, apuesta seleccionada, método
'Siguiente dia
'
Private Sub cmd_MetodoOptimo(vNewValue As ParametrosSimulacion)
    Dim i As Integer, j As Integer, h As Integer
    Dim m_array_Metodos() As Variant
    Dim m_res               As New Resultado    'Objeto Combinación Ganadora
    Dim m_metodo            As New MetodoOld    'Objeto Metodo de Sugerencia que aglutina
                                                'los parametros del metodo
    Dim m_ms                As New Muestra      'Objeto que representa la muestra de analisis
    Dim m_Bd                As New BdDatos      'base de datos
    Dim m_fei               As Date             'Fecha de inicio del rango de analisis
    Dim m_fef               As Date             'Fecha de fin del rango de analisis
    Dim rg_datos            As Range            'Rango de Combinaciones a analizar
    Dim rg_muestra          As Range            'Rango de la muestra
    Dim Fila                As Range            'Fila de resultados para analizar
    Dim m_CU                As CU_DefinirApuesta 'Proceso de obtención de la apuesta
    Dim m_CU2               As CU_ComprobarApuesta 'Proceso de obtención de la apuesta
    Dim m_apta              As ApuestaOld       ' Apuesta
    Dim m_array             As Variant          ' Declara una matriz
    Dim m_conta_dias        As Integer          ' contador de dias de apuesta
    Dim iFecha              As Date
    Dim mRgMuestra          As IntervaloTiempo
    Dim mRgDias             As rango
    Dim mRgRetraso          As rango
    Dim mxAciertos          As Integer
    Dim miAciertos          As Integer
    Dim mxMetodo            As String
    Dim mxMuestra           As Muestra
    Dim mRow                As Integer
    Dim celda               As Range
    Dim m_apuesta_optima    As Apuesta
    Dim mMetodo             As ParametrosMetodoOld
   
   On Error GoTo cmd_MetodoOptimo_Error
    
    Set mMetodo = vNewValue.Metodos(1)
    
    'Mostrar parametros del metodo
    Pintar_Parametros vNewValue
    
    If (mMetodo.DiasMuestra <= 4) Then
        Err.Raise -999, "cmd_MetodoOptimo", "Los días de la Muestra del analisis debe ser superior a 4."
    End If
    '
    '   Definir Intervalos de Pruebas
    mRgDias.ElementoInferior = 4                                               ' Rango mínimo de dias a analizar
    mRgDias.ElementoSuperior = mMetodo.DiasMuestra                             ' Rango de dias a analizar
    mRgRetraso.ElementoInferior = 0
    mRgRetraso.ElementoSuperior = mMetodo.DiasRetardo

    '   Definir Metodo a probar
    Set m_metodo = New MetodoOld
    m_metodo.Pronosticos = 7
    m_array_Metodos = vNewValue.ArrayMetodos
    Range("D3").Activate
    Set m_CU = New CU_DefinirApuesta                                            ' Declara el CU Definir una apuesta
    Set m_CU2 = New CU_ComprobarApuesta                                         ' Declara el CU Comprobar una apuesta
    Set m_apuesta_optima = New ApuestaOld                                          ' Declara la apuesta optima
    
    Debug.Print "Inicio => " & Time
    For iFecha = vNewValue.RangoAnalisis.FechaInicial To vNewValue.RangoAnalisis.FechaFinal Step 1
    
        If Application.WorksheetFunction.Weekday(iFecha) <> 1 Then          ' No es domingo
            Set m_res = m_Bd.Get_Resultado(iFecha)

            mxAciertos = 0
            Set m_CU2.Resultado = m_res
            
            For i = mRgDias.ElementoInferior To mRgDias.ElementoSuperior Step 1
                For j = mRgRetraso.ElementoInferior To mRgRetraso.ElementoSuperior Step 1
                    '
                    '   Obtenemos la muestra
                    '   Fecha final de la muestra = Fecha de analisis - dias de retardo - 1 dia
                    mRgMuestra.FechaFinal = iFecha - j - 1
                    '
                    '   Fecha inicial de la muestra = fecha final - dias de Muestra
                    mRgMuestra.FechaInicial = mRgMuestra.FechaFinal - i
                    Set rg_muestra = m_Bd.Resultados_Fechas(mRgMuestra.FechaInicial, mRgMuestra.FechaFinal)
                    
                    m_metodo.Fecha_Evaluacion = iFecha
                    m_ms.Constructor rg_muestra, JUEGO_DEFECTO
                    '
                    '   Prueba los métodos seleccionados
                    '
                    For h = 0 To UBound(m_array_Metodos) Step 1
                        m_metodo.Tipo_Metodo = m_array_Metodos(h)  'Solo metodos de control
                        Set m_CU.metodo = m_metodo          ' Asigna el método de obtención
                        Set m_CU.Muestra = m_ms             ' Asigna la muestra de números
                        Set m_apta = m_CU.Get_Apuesta       ' Obtiene la apuesta del CU
                                                            ' Averiguamos los aciertos
                        Set m_CU2.Apuesta = m_apta
                        m_apta.metodo.Dias_Proceso = i
                        m_apta.metodo.Dias_Retraso = j
                        m_apta.aciertos = m_CU2.Get_Aciertos(False)
                        If (m_apta.aciertos >= 3) Then
                            Set celda = Range("D3").Offset(mRow, 0)
                            Pintar_resultado celda, m_res, m_apta, m_apta.aciertos, m_apta.metodo.Nombre_metodo, m_ms
                            mRow = mRow + 1
                        End If
    '                    Debug.Print "Prueba: => " & iFecha & ", I=> " & i; ", J=> " & j; ", Aciertos => " & m_cu2.Get_Aciertos(False)
                    Next h
                Next j
            Next i
'            Set celda = Range("D3").Offset(mRow, 0)
'            Pintar_resultado celda, m_res, m_apuesta_optima, mxAciertos, mxMetodo, mxMuestra
'            mRow = mRow + 1
        End If
    Next iFecha
    Debug.Print "Fin => " & Time
       
    Cells.Select                            'Selecciona todas las celdas de la hoja
    Cells.EntireColumn.AutoFit              'Autoajusta el tamaño de las columnas
    
   On Error GoTo 0
   Exit Sub

cmd_MetodoOptimo_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & _
        ") in procedure cmd_MetodoOptimo of Módulo Lot_MetodoOptimo"

End Sub

Private Sub Pintar_Literales_01()

   On Error GoTo print_literales_Error
    Range("A1").Activate

    ActiveCell.Value = "Método óptimo"
    ActiveCell.Font.Bold = True
    ActiveCell.Offset(1, 0).Value = "Fecha Inicial"
    ActiveCell.Offset(2, 0).Value = "Fecha Final"
    ActiveCell.Offset(3, 0).Value = "Pronósticos"
    ActiveCell.Offset(4, 0).Value = "Sorteos de Análisis"
    ActiveCell.Offset(5, 0).Value = "Días Retardo"
    ActiveCell.Offset(6, 0).Value = "Tipo Método"
    
'
''*-----------------------|
    Range("D2").Select                        ' Se posiciona en el inicio del rótulo
    ActiveCell.Value = "Fecha"
    ActiveCell.Offset(0, 1).Value = "Día"
    ActiveCell.Offset(0, 2).Value = "N1"
    ActiveCell.Offset(0, 3).Value = "N2"
    ActiveCell.Offset(0, 4).Value = "N3"
    ActiveCell.Offset(0, 5).Value = "N4"
    ActiveCell.Offset(0, 6).Value = "N5"
    ActiveCell.Offset(0, 7).Value = "N6"
    ActiveCell.Offset(0, 8).Value = "C"
    ActiveCell.Offset(0, 9).Value = "_"
    ActiveCell.Offset(0, 10).Value = "Apuesta"
    ActiveCell.Offset(0, 11).Value = "Aciertos"
    ActiveCell.Offset(0, 12).Value = "Dias"
    ActiveCell.Offset(0, 13).Value = "Retardo"
    ActiveCell.Offset(0, 14).Value = "Método"
    
     Range("D2:r2").Select
    With Selection                              ' con la seleccion se modifican propiedades
        .HorizontalAlignment = xlCenter         ' Alineación horizontal centrada
        .VerticalAlignment = xlBottom           ' Alineación vertical bajo
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
'    Selection.Merge                             ' Mezcla las celdas
        
    Cells.Select                                ' Selecciona todas las celdas de la hoja
    Cells.EntireColumn.AutoFit                  ' Autoajusta el tamaño de las columnas
   
   On Error GoTo 0
   Exit Sub

print_literales_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure print_literales of Módulo Pintar_Literales_01"

End Sub

Private Sub Pintar_Parametros(vNewValue As ParametrosSimulacion)
    Dim mMtd As ParametrosMetodoOld
    Set mMtd = vNewValue.Metodos(1)
    Range("B1").Activate
    ActiveCell.Offset(1, 0).Value = vNewValue.RangoAnalisis.FechaInicial
    ActiveCell.Offset(2, 0).Value = vNewValue.RangoAnalisis.FechaFinal
    ActiveCell.Offset(3, 0).Value = vNewValue.Pronosticos
    ActiveCell.Offset(4, 0).Value = mMtd.DiasMuestra
    ActiveCell.Offset(5, 0).Value = mMtd.DiasRetardo
    'ActiveCell.Offset(6, 0).Value = mMtd.DiasRetardo

End Sub

Private Sub Pintar_resultado(rango As Range, parmRes As Resultado, _
                             paramApuesta As ApuestaOld, aciertos As Integer, _
                             metodo As String, parmMuestra As Muestra)
    rango.Offset(0, 0).Value = parmRes.Fecha
    rango.Offset(0, 1).Value = parmRes.Dia
    rango.Offset(0, 2).Value = parmRes.Numeros(0)
    Colorea_Celda rango.Offset(0, 2), parmRes.Numeros(0), parmMuestra, paramApuesta.metodo
    rango.Offset(0, 3).Value = parmRes.Numeros(1)
    Colorea_Celda rango.Offset(0, 3), parmRes.Numeros(1), parmMuestra, paramApuesta.metodo
    rango.Offset(0, 4).Value = parmRes.Numeros(2)
    Colorea_Celda rango.Offset(0, 4), parmRes.Numeros(2), parmMuestra, paramApuesta.metodo
    rango.Offset(0, 5).Value = parmRes.Numeros(3)
    Colorea_Celda rango.Offset(0, 5), parmRes.Numeros(3), parmMuestra, paramApuesta.metodo
    rango.Offset(0, 6).Value = parmRes.Numeros(4)
    Colorea_Celda rango.Offset(0, 6), parmRes.Numeros(4), parmMuestra, paramApuesta.metodo
    rango.Offset(0, 7).Value = parmRes.Numeros(5)
    Colorea_Celda rango.Offset(0, 7), parmRes.Numeros(5), parmMuestra, paramApuesta.metodo
    rango.Offset(0, 8).Value = parmRes.Complementario
    Colorea_Celda rango.Offset(0, 8), parmRes.Complementario, parmMuestra, paramApuesta.metodo
    
    rango.Offset(0, 10).Value = paramApuesta.Texto
    rango.Offset(0, 11).Value = aciertos
    rango.Offset(0, 12).Value = paramApuesta.metodo.Dias_Proceso
    rango.Offset(0, 13).Value = paramApuesta.metodo.Dias_Retraso
    rango.Offset(0, 14).Value = metodo
End Sub


