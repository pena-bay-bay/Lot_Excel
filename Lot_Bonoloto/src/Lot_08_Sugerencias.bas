Attribute VB_Name = "Lot_08_Sugerencias"
' *============================================================================*
' *
' *     Fichero    : Lot_Sugerencias
' *
' *     Autor      : Carlos Almela Baeza
' *     Creación   : lun, 10/10/2011 23:26
' *     Versión    : 1.0
' *     Propósito  :
' *
' *
' *============================================================================*
Private DB As New BdDatos
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

        Borra_Salida                    'Borra el contenido de la hoja de salida
        
        Pintar_Textos                   'Rellena Literales
        
        Set frmSugerencia = New frmSugerencia 'Crea un formulario de captura de métodos
        
        frmSugerencia.Tag = ESTADO_INICIAL 'inicializa el estado
        
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
                     Borra_Salida       ' Borra el contenido de la hoja de salida
'                     pintar_Sugerencia frmSugerencia.Parametros  ' Rellena Literales
'                     cmd_Simular frmSugerencia.Parametros
                       
            End Select
        Loop

btn_SugerirApuestas_CleanExit:
   On Error GoTo 0
    Exit Sub

btn_SugerirApuestas_Error:
    

    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    '   Audita el error
    Call HandleException(ErrNumber, ErrDescription, "Lot_Sugerencias.btn_SugerirApuestas")
    '   Informa del error
    Call MsgBox(ErrDescription, vbError Or vbSystemModal, NOMBRE_APLICACION)
End Sub
' *============================================================================*
' *     Procedure  : Pintar_Sugerencia
' *     Fichero    : Lot_Sugerencias
' *     Autor      : Carlos Almela Baeza
' *     Creacion   : lun, 10/10/2011
' *     Asunto     :
' *============================================================================*
'
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

Pintar_Sugerencia_CleanExit:
   On Error GoTo 0
    Exit Sub

pintar_Sugerencia_Error:
    

    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    '   Audita el error
    Call HandleException(ErrNumber, ErrDescription, "Lot_Sugerencias.Pintar_Sugerencia")
    '   Lanza el Error
    Err.Raise ErrNumber, ErrSource, ErrDescription
End Sub
' *============================================================================*
' *     Procedure  : cmd_Simular
' *     Fichero    : Lot_Sugerencias
' *     Autor      : Carlos Almela Baeza
' *     Creacion   : lun, 10/10/2011
' *     Asunto     :
' *============================================================================*
'
Private Sub cmd_Simular(vNewData As ParametrosSimulacion)
        Dim objParMuestra As ParametrosMuestra
        Dim mMtd As ParametrosMetodoOld, mMetodo As MetodoOld
        Dim mMuestra As Muestra, mRango As Range, mApuesta As Apuesta
        Dim mPer As Periodo, i As Integer, n As Variant, j As Integer
        Dim m_CU As CU_DefinirApuesta
        Dim m_array As Variant          ' Declara una matriz
        Dim mMuestraColor As Muestra
        '
        '       Calcular el metodo de coloreo con 45 días
        '
  On Error GoTo cmd_Simular_Error

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
        Set mMetodo = New metodo
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
            Set mMuestra.ParametrosMuestra = objParMuestra
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

cmd_Simular_CleanExit:
   On Error GoTo 0
    Exit Sub

cmd_Simular_Error:
    

    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    '   Audita el error
    Call HandleException(ErrNumber, ErrDescription, "Lot_Sugerencias.cmd_Simular")
    '   Lanza el Error
    Err.Raise ErrNumber, ErrSource, ErrDescription
End Sub

' *============================================================================*
' *     Procedure  : Pintar_Textos
' *     Fichero    : Lot_Sugerencias
' *     Autor      : Carlos Almela Baeza
' *     Creacion   : lun, 10/10/2011
' *     Asunto     :
' *============================================================================*
'
Private Sub Pintar_Textos()
    '       Parámetros del proceso
  On Error GoTo Pintar_Textos_Error

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

Pintar_Textos_CleanExit:
   On Error GoTo 0
    Exit Sub

Pintar_Textos_Error:
    

    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    '   Audita el error
    Call HandleException(ErrNumber, ErrDescription, "Lot_Sugerencias.Pintar_Textos")
    '   Lanza el Error
    Err.Raise ErrNumber, ErrSource, ErrDescription
End Sub

' *============================================================================*
' *     Procedure  : Pintar_Parametros
' *     Fichero    : Lot_Sugerencias
' *     Autor      : Carlos Almela Baeza
' *     Creacion   : lun, 10/10/2011
' *     Asunto     :
' *============================================================================*
'
'Private Sub Pintar_Parametros(vNewData As ParametrosSimulacion)
'        Dim mpar As ParametrosMetodo
'        Dim I As Integer
'  On Error GoTo Pintar_Parametros_Error
'
'        Range("B2").Activate
'        ActiveCell.Value = vNewData.FechaInicial
'        ActiveCell.Offset(1, 0).Value = vNewData.FechaFinal
'        ActiveCell.Offset(2, 0).Value = vNewData.dias
'        ActiveCell.Offset(3, 0).Value = vNewData.Pronosticos
'        ActiveCell.Offset(4, 0).Value = vNewData.NumMetodos
'
'        Range("N2").Activate
'        I = 0
'        For Each mpar In vNewData.Metodos
'            ActiveCell.Offset(0, I).Value = "M" + CStr(mpar.Id)
'            ActiveCell.Offset(0, I).AddComment
'            ActiveCell.Offset(0, I).Comment.Text Text:=mpar.ToString
'            I = I + 1
'        Next mpar
'
'        ActiveCell.Offset(0, I).Value = "Coste"
'        I = I + 1
'        ActiveCell.Offset(0, I).Value = "Premio"
'
'Pintar_Parametros_CleanExit:
'   On Error GoTo 0
'    Exit Sub
'
'Pintar_Parametros_Error:
'
'
'    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
'    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
'    '   Audita el error
'    Call HandleException(ErrNumber, ErrDescription, "Lot_Sugerencias.Pintar_Parametros")
'    '   Lanza el Error
'    Err.Raise ErrNumber, ErrSource, ErrDescription
'
'End Sub

  
