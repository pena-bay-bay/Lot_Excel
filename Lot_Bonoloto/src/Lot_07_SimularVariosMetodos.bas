Attribute VB_Name = "Lot_07_SimularVariosMetodos"
' *============================================================================*
' *
' *     Fichero    : Lot_SimularVariosMetodos
' *
' *     Autor      : Carlos Almela Baeza
' *     Creación   : vie, 18/03/2011 20:29
' *     Versión    : 1.0
' *     Propósito  :
' *
' *
' *============================================================================*
Option Explicit
Private DB As New BdDatos


' *============================================================================*
' *     Procedure  : btn_SimularVariosMetodos
' *     Fichero    : Lot_SimularVariosMetodos
' *     Autor      : Carlos Almela Baeza
' *     Creacion   : vie, 18/03/2011
' *     Asunto     :
' *============================================================================*
'
Public Sub btn_SimularVariosMetodos()
On Error GoTo btn_SimularVariosMetodos_error

    Borra_Salida                    'Borra el contenido de la hoja de salida
    
    Pintar_Textos                   'Rellena Literales
    
    Set frmMetodos = New frmMetodos 'Crea un formulario de captura de métodos
    
    frmMetodos.Tag = ESTADO_INICIAL 'inicializa el estado
    
    Do While frmMetodos.Tag <> BOTON_CERRAR
    
        ' Se inicializa el boton cerrar para salir del bucle
        frmMetodos.Tag = BOTON_CERRAR
        
        ' Se muestra el formulario y queda a la espera de funciones
        ' pulsando el botón ejecutar
        frmMetodos.Show
       
        'Se bifurca la función
        Select Case frmMetodos.Tag
                                    ' El usuario ha cerrado el
            Case ""                 ' cuadro de dialogo con la [X]
                frmMetodos.Tag = BOTON_CERRAR
            
            Case EJECUTAR           ' Se ha pulsado el botón ejecutar
                 Borra_Salida       ' Borra el contenido de la hoja de salida
                 Pintar_Textos      ' Rellena Literales
                 Pintar_Parametros frmMetodos.Parametros
                 cmd_Ejecutar frmMetodos.Parametros
                               
            Case SIMULAR_METODOS    ' Propone una apuesta para los diversos métodos
                 Borra_Salida       ' Borra el contenido de la hoja de salida
                 pintar_Sugerencia frmMetodos.Parametros  ' Rellena Literales
                 cmd_Simular frmMetodos.Parametros
                   
        End Select
    Loop

On Error GoTo 0
   Exit Sub

btn_SimularVariosMetodos_error:
    Set frmMetodos = Nothing
    MsgBox "Error " & Err.Number & " (" & Err.Description & _
    ") in procedure btn_MultiplesMetodos of Módulo Lot_SimularVariosMetodos"

End Sub

Private Sub pintar_Sugerencia(vNewData As ParametrosSimulacion)
    Dim i As Integer
    
On Error GoTo pintar_Sugerencia_Error
    Range("A1").Activate
    
'Parámetros de la sugerencia
    ActiveCell.Value = "Sugerencia Múltiple"
    ActiveCell.Font.Bold = True
    ActiveCell.Offset(1, 0).Value = "Fecha de Sugerencia"
    ActiveCell.Offset(2, 0).Value = "Métodos"
    ActiveCell.Offset(3, 0).Value = "Pronósticos"
    
    ActiveCell.Offset(1, 1).Value = vNewData.FechaFinal
    ActiveCell.Offset(1, 1).NumberFormat = "ddd, dd/mm/yyyy"
    ActiveCell.Offset(2, 1).Value = vNewData.NumMetodos
    ActiveCell.Offset(3, 1).Value = vNewData.Pronosticos

' Literales de las Columnas
    Range("D2").Activate
    ActiveCell.Offset(0, 0).Value = "Descripcion Método"
    For i = 1 To vNewData.Pronosticos
        ActiveCell.Offset(0, i).Value = "N" + CStr(i)
    Next i
    
    Cells.Select                                ' Selecciona todas las celdas de la hoja
    Cells.EntireColumn.AutoFit                  ' Autoajusta el tamaño de las columnas
   
   On Error GoTo 0
   Exit Sub

pintar_Sugerencia_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & _
    ") in procedure Pintar_Sugerencia of Módulo Lot_SimularVariosMetodos"

End Sub
'
'
'
'
Private Sub cmd_Simular(vNewData As ParametrosSimulacion)
    Dim objParMuestra As ParametrosMuestra
    Dim mMtd As ParametrosMetodoOld, mMetodo As MetodoOld
    Dim mMuestra As Muestra, mRango As Range, mApuesta As Apuesta
    Dim mPer As Periodo, i As Integer, n As Variant, j As Integer
    Dim m_CU As CU_DefinirApuesta
    Dim m_array As Variant          ' Declara una matriz
    Dim mMuestraColor As Muestra
On Error GoTo cmd_Simular_Error

    '
    '       Calcular el metodo de coloreo con 45 días
    '
    Set mRango = DB.Resultados_Fechas(vNewData.FechaFinal - 45, vNewData.FechaFinal)
    Set mMuestraColor = New Muestra
    mMuestraColor.Constructor mRango, JUEGO_DEFECTO
    '
    '
    '
    Set mMuestra = New Muestra
    Set objParMuestra = New ParametrosMuestra
    Set m_CU = New CU_DefinirApuesta    ' Declara el CU Definir una apuesta
    Set mPer = New Periodo
    Set mMetodo = New MetodoOld
    Range("D3").Activate
    i = 0
    For Each mMtd In vNewData.Metodos
            '
            '
            '
            ActiveCell.Offset(i, 0).Value = mMtd.ToString()
            mPer.FechaFinal = vNewData.FechaFinal - 1 - mMtd.DiasRetardo
            mPer.FechaInicial = mPer.FechaFinal - mMtd.DiasMuestra
            objParMuestra.FechaAnalisis = vNewData.FechaFinal
            objParMuestra.FechaFinal = mPer.FechaFinal
            objParMuestra.FechaInicial = mPer.FechaInicial
         
            Set mRango = DB.Resultados_Fechas(mPer.FechaInicial, mPer.FechaFinal)
            Set objParMuestra = New ParametrosMuestra
            mMuestra.Constructor mRango, JUEGO_DEFECTO
            mMetodo.Tipo_Metodo = mMtd.Ordenacion
            mMetodo.Pronosticos = vNewData.Pronosticos
            Set m_CU.metodo = mMetodo          ' Asigna el método de obtención
            Set m_CU.Muestra = mMuestra        ' Asigna la muestra de números
            Set mApuesta = m_CU.Get_Apuesta    ' Obtiene la apuesta del CU
            m_array = mApuesta.Pronosticos     ' Obtiene una matriz con los pronosticos
            j = 1
            For Each n In m_array              ' Para cada número en la matríz
                                               ' Colorea la celda con el color de la
                                               ' probabilidad
            Colorea_CeldaProb ActiveCell.Offset(i, j), n, mMuestraColor
            j = j + 1                       ' Incrementa la columna
        Next n                              ' Siguiente número
        i = i + 1
    Next mMtd
    
    Cells.Select                            'Selecciona todas las celdas de la hoja
    Cells.EntireColumn.AutoFit              'Autoajusta el tamaño de las columnas


On Error GoTo 0
    Exit Sub

cmd_Simular_Error:
        MsgBox "Error " & Err.Number & " (" & Err.Description & _
        ") in procedure cmd_Simular of Módulo Lot_SimularVariosMetodos"
End Sub

Private Sub Pintar_Textos()
    '       Parámetros del proceso
    Range("A1").Activate
    ActiveCell.Value = "Métodos Múltiples"
    ActiveCell.Font.Bold = True
    ActiveCell.Offset(1, 0).Value = "Fecha inicial"
    ActiveCell.Offset(2, 0).Value = "Fecha final"
    ActiveCell.Offset(3, 0).Value = "Dias Analizados"
    ActiveCell.Offset(4, 0).Value = "Pronosticos"
    ActiveCell.Offset(5, 0).Value = "Total metodos"
    ActiveCell.Offset(6, 0).Value = "Colores Sorteo"
    
    '       Datos
    Range("D2").Activate
    ActiveCell.Value = "Fecha"
    ActiveCell.Offset(0, 1).Value = "Día"
    ActiveCell.Offset(0, 2).Value = "N1"
    ActiveCell.Offset(0, 3).Value = "N2"
    ActiveCell.Offset(0, 4).Value = "N3"
    ActiveCell.Offset(0, 5).Value = "N4"
    ActiveCell.Offset(0, 6).Value = "N5"
    ActiveCell.Offset(0, 7).Value = "N6"
    ActiveCell.Offset(0, 8).Value = "C"
    ActiveCell.Offset(0, 9).Value = "Total"
End Sub

Private Sub Pintar_Parametros(vNewData As ParametrosSimulacion)
    Dim mPar As ParametrosMetodo
    Dim i As Integer

On Error GoTo Pintar_Parametros_Error

    Range("B2").Activate
    ActiveCell.Value = vNewData.FechaInicial
    ActiveCell.Offset(1, 0).Value = vNewData.FechaFinal
    ActiveCell.Offset(2, 0).Value = vNewData.Dias
    ActiveCell.Offset(3, 0).Value = vNewData.Pronosticos
    ActiveCell.Offset(4, 0).Value = vNewData.NumMetodos
    
    Range("N2").Activate
    i = 0
    For Each mPar In vNewData.Metodos
        ActiveCell.Offset(0, i).Value = "M" + CStr(mPar.Id)
        ActiveCell.Offset(0, i).AddComment
        ActiveCell.Offset(0, i).Comment.Text Text:=mPar.ToString
        i = i + 1
    Next mPar
    
    ActiveCell.Offset(0, i).Value = "Coste"
    i = i + 1
    ActiveCell.Offset(0, i).Value = "Premio"
    
On Error GoTo 0
   Exit Sub

Pintar_Parametros_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & _
    ") in procedure Pintar_Parametros of Módulo Modulo1"
End Sub

Private Sub cmd_Ejecutar(vNewData As ParametrosSimulacion)
    Dim rg_datos As Range, Fila As Range, mRangoSorteo As Range
    Dim i As Integer, j As Integer, mPremio As String, mTotal As Integer
    Dim mSorteo As New Resultado, mApuesta As ApuestaOld, mMetodo As MetodoOld
    Dim mMtd As ParametrosMetodoOld, mMuestra As Muestra, mImporte As Currency
    Dim mCUComprobar As CU_ComprobarApuesta, mCuSuger As CU_DefinirApuesta
    Dim objParMuestra As ParametrosMuestra
    
On Error GoTo cmd_Ejecutar_error
    
    'obtiene el rango con los datos comprendido entre las dos fechas
    Set rg_datos = DB.Resultados_Fechas(vNewData.FechaInicial, vNewData.FechaFinal)
   
   'Nos posicionamos en la celda de inicio de escritura
    Range("D3").Activate                'Se posiciona el cursor en la
                                        'celda D3
    i = 0                               'inicializa el contador de filas
    
    'Recorre las celdas del rango de analisis
    For Each Fila In rg_datos.Rows
            
            'Componemos el resultado
            mSorteo.Constructor Fila
            
            'Obtiene la muestra para los colores del sorteo
            ' TO DO
            '
            '
            '   Crea los objetos de los casos de uso
            Set mCuSuger = New CU_DefinirApuesta
            Set mMetodo = New metodo
            Set mMuestra = New Muestra
            Set mCUComprobar = New CU_ComprobarApuesta
            Set objParMuestra = New ParametrosMuestra
            '
            '       Pintar_Resultado
            ActiveCell.Offset(i, 0).Value = mSorteo.Fecha
            ActiveCell.Offset(i, 1).Value = mSorteo.Dia
            ActiveCell.Offset(i, 2).Value = mSorteo.Numeros(0)
            ActiveCell.Offset(i, 3).Value = mSorteo.Numeros(1)
            ActiveCell.Offset(i, 4).Value = mSorteo.Numeros(2)
            ActiveCell.Offset(i, 5).Value = mSorteo.Numeros(3)
            ActiveCell.Offset(i, 6).Value = mSorteo.Numeros(4)
            ActiveCell.Offset(i, 7).Value = mSorteo.Numeros(5)
            ActiveCell.Offset(i, 8).Value = mSorteo.Complementario
            
            j = 10                          'Inicializa el contador de columnas
                                            'Para cada método en la coleccion
            mTotal = 0                      ' Inicializa contador de Métodos acertados
            mImporte = 0                    ' Inicializa el acumulado de premios
                                            ' Para cada método en la colección
            For Each mMtd In vNewData.Metodos
                    Debug.Print mMtd.ToString
                    '
                    '   Configura el método
                    '
                    mMetodo.Pronosticos = vNewData.Pronosticos
                    mMetodo.Fecha_Evaluacion = mSorteo.Fecha
                    mMetodo.Fecha_Final = mSorteo.Fecha - 1 - mMtd.DiasRetardo
                    mMetodo.Fecha_Inicial = mMetodo.Fecha_Final - mMtd.DiasMuestra
                    mMetodo.Tipo_Metodo = mMtd.Ordenacion
                    '
                    '   Define los parametros de la muestra
                    '
                    objParMuestra.FechaAnalisis = mMetodo.Fecha_Evaluacion
                    objParMuestra.FechaFinal = mMetodo.Fecha_Final
                    objParMuestra.FechaInicial = mMetodo.Fecha_Inicial
                    'objParMuestra.ModalidadJuego = LP_LB_6_49
                    '
                    '   Define la muestra
                    Set mRangoSorteo = DB.Resultados_Fechas(mMetodo.Fecha_Inicial, mMetodo.Fecha_Final)
                    Set mMuestra.ParametrosMuestra = objParMuestra
                    mMuestra.Constructor mRangoSorteo, JUEGO_DEFECTO
                    '
                    '   Carga el Definir apuesta
                    Set mCuSuger.metodo = mMetodo
                    Set mCuSuger.Muestra = mMuestra
                    Set mApuesta = mCuSuger.Get_Apuesta
                    '
                    '
                    '   Comprobar apuesta
                    Set mCUComprobar.Apuesta = mApuesta
                    Set mCUComprobar.Resultado = mSorteo
                    '
                    '
                    mPremio = mCUComprobar.Get_Premio()
                    If mPremio = "" Then
                        '
                        '   Obtiene el total de números acertados con el complementario
                        ActiveCell.Offset(i, j).Value = mCUComprobar.Get_Aciertos(Mas_Complementario:=True)
                    Else
                        '
                        '       Acumula un método con premio
                        mTotal = mTotal + 1
                        '
                        '       Obtiene la categoría del premio
                        ActiveCell.Offset(i, j).Value = mCUComprobar.Get_Premio()
                        ColoreaCelda ActiveCell.Offset(i, j), COLOR_VERDE
                        '
                        '       Obtiene el premio de la categoría
'                        mImporte = mImporte + mCUComprobar.Get_ImportePremio()
                        mImporte = mImporte + mCUComprobar.GetImporteEsperado()
                    End If
                '
                '       Incrementa la columna
                j = j + 1
            Next mMtd
            '
            '       Visualiza el total de métodos acertados
            ActiveCell.Offset(i, 9).Value = mTotal
            '
            '       Calcula el coste de las apuestas realizadas
            ActiveCell.Offset(i, j).Value = Cal_Coste(mSorteo.Dia, vNewData.Pronosticos, vNewData.Metodos.Count)
            '
            '       Visualiza el importe de los premios
            ActiveCell.Offset(i, j + 1).Value = mImporte
            '
            '       Incrementa la fila
            i = i + 1
    Next Fila

    Cells.Select                            'Selecciona todas las celdas de la hoja
    Cells.EntireColumn.AutoFit              'Autoajusta el tamaño de las columnas

On Error GoTo 0
    Exit Sub

cmd_Ejecutar_error:
        MsgBox "Error " & Err.Number & " (" & Err.Description & _
        ") in procedure cmd_Ejecutar of Módulo Módulo1"

End Sub
    
Private Function Cal_Coste(mDia As String, mPronosticos As Integer, mNumMetodos As Integer) As Currency
    Dim mCoste As Currency
    Dim Total_Apuestas As Integer
    mCoste = 0
    
    Select Case mPronosticos
        Case 5: Total_Apuestas = 44
        Case 6: Total_Apuestas = 1
        Case 7: Total_Apuestas = 7
        Case 8: Total_Apuestas = 28
        Case 9: Total_Apuestas = 84
        Case 10: Total_Apuestas = 210
        Case 11: Total_Apuestas = 462
        Case Else: Total_Apuestas = 0
    End Select

    Select Case mDia
        Case "L", "M", "X", "V":
            mCoste = Total_Apuestas * 0.5
        Case "J", "S":
            mCoste = Total_Apuestas * 1
        Case "D":
            mCoste = Total_Apuestas * 1.5
    End Select
    Cal_Coste = mCoste * mNumMetodos
End Function

