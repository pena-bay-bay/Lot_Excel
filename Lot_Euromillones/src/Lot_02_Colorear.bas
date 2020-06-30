Attribute VB_Name = "Lot_02_Colorear"
'---------------------------------------------------------------------------------------
' Modulo    : Lot_Colorear.bas
' Creado    : 16/03/2007  22:14
' Autor     : Carlos Almela Baeza
' Version   : 1.0.1 20/03/2007 9:50
' Objeto    : Funciones que colorean los resultados
'---------------------------------------------------------------------------------------
Option Explicit
Option Base 0

'*-----------------| OBJETOS |-----------------------------+
Private DB                      As New BdDatos          'Objeto Base de Datos
Private m_array                 As Variant
Private m_valor                 As Integer
Private m_rgFila                As Range
Private m_rgDatos               As Range
Private i                       As Integer
Private m_res                   As Sorteo
Private color                   As Integer
Private ColIni                  As Integer
Private ColFin                  As Integer

'---------------------------------------------------------------------------------------
' Procedure : cmd_Colorear
' DateTime  : 12/jun/2007 23:42
' Author    : Carlos Almela Baeza
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Sub btn_Colorear()
    
  On Error GoTo btn_Colorear_Error
    '
    ' Nos colocamos en la página de resultados
    '
    DB.Ir_A_Hoja "Resultados"
    '
    '   obtiene el rango de Resultados  e inicializa el color
    '
    Set m_rgDatos = DB.RangoResultados
    '
    '   Función colorear un rango
    '
    ColoreaCelda m_rgDatos, xlColorIndexAutomatic
    '
    '   Presenta formulario de Tipo de Funcionalidad
    '
    frmRealizarColoreado.Tag = ESTADO_INICIAL
    '
    '
    '
    Do While frmRealizarColoreado.Tag <> BOTON_CERRAR
        ' Se inicializa el boton cerrar para salir del bucle
        frmRealizarColoreado.Tag = BOTON_CERRAR
        
        ' Se muestra el formulario y queda a la espera de funciones
        ' pulsando el botón ejecutar
        frmRealizarColoreado.Show
       
        'Se bifurca la función
        Select Case frmRealizarColoreado.Tag
                                    ' El usuario ha cerrado el
            Case ""                 ' cuadro de dialogo con la [X]
                frmRealizarColoreado.Tag = BOTON_CERRAR
            
            Case COLOREAR_CARACTERISTICAS
                cmd_color_caracteristicas frmRealizarColoreado.Tipo_Caracteristica
            
            Case COLOREAR_UNAFECHA
                cmd_color_fecha frmRealizarColoreado.Fecha_Sorteo
            
            Case COLOREAR_NumeroS
                cmd_color_combinacion frmRealizarColoreado.TextCombinacion
        
        End Select
    Loop

btn_Colorear_CleanExit:
   On Error GoTo 0
    Exit Sub

btn_Colorear_Error:
    

    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    '   Audita el error
    Call HandleException(ErrNumber, ErrDescription, "Lot_Colorear.btn_Colorear")
    '   Informa del error
    Call MsgBox(ErrDescription, vbError Or vbSystemModal, NOMBRE_APLICACION)
    Call Trace("CERRAR")
End Sub

'---------------------------------------------------------------------------------------
' Procedure : cmd_color_fecha
' DateTime  : 12/jun/2007 23:42
' Author    : Carlos Almela Baeza
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub cmd_color_fecha(vNewData As Date)

   On Error GoTo cmd_color_fecha_error
    
    'obtiene el resultado de la fecha
     Set m_res = DB.Get_Resultado(vNewData)
     If m_res Is Nothing Then
        Err.Raise 100, "Lot_02_Colorear.cmd_color_fecha", "No existe sorteo para esta fecha"
        Exit Sub
     End If
     m_array = m_res.Combinacion.Numeros
     
    'obtiene el rango de datos
    'y lo colorea de blanco
    Set m_rgDatos = DB.RangoResultados
    ColoreaCelda m_rgDatos, xlColorIndexAutomatic
    
    ' para cada fila(resultado) en el rango de Datos
    For Each m_rgFila In m_rgDatos.Rows
            
            'Comprueba los Numeros que se encuentran
            'entre las columnas E y K
            'y colorea de anaranjado si lo encuentra
             For i = 6 To 12
                color = 0
                If (IsNumeric(m_rgFila.Cells(1, i).Value)) Then
                    m_valor = m_rgFila.Cells(1, i).Value
                Else
                    m_valor = 0
                End If
                Select Case m_valor
                    Case m_array(0): color = COLOR_TERMINACION8
                    Case m_array(1): color = COLOR_TERMINACION1
                    Case m_array(2): color = COLOR_TERMINACION2
                    Case m_array(3): color = COLOR_TERMINACION3
                    Case m_array(4): color = COLOR_TERMINACION4
                    Case m_array(5): color = COLOR_TERMINACION5
                    Case m_res.Complementario: color = COLOR_TERMINACION6
                End Select
                If color > 0 Then
                        ColoreaCelda m_rgFila.Cells(1, i), color
                End If
            Next i
    Next m_rgFila
    

   On Error GoTo 0
   Exit Sub

cmd_color_fecha_error:

   Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
   ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
   '   Audita el error
   Call HandleException(ErrNumber, ErrDescription, ErrSource, "Lot_02_Colorear.cmd_color_fecha")
   '   Lanza el error
   Err.Raise ErrNumber, "Lot_02_Colorear.cmd_color_fecha", ErrDescription

End Sub

'---------------------------------------------------------------------------------------
' Procedure : cmd_color_combinacion
' DateTime  : 12/jun/2007 23:42
' Author    : Carlos Almela Baeza
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub cmd_color_combinacion(vNewData As Apuesta)
    Dim j As Integer
    Dim h As Integer
    Dim m_num As New Numero
    
   On Error GoTo cmd_color_combinacion_Error

      
    'obtiene el rango de datos
    Set m_rgDatos = DB.RangoResultados
    ColoreaCelda m_rgDatos, xlColorIndexNone
    
    Select Case JUEGO_DEFECTO
        Case Bonoloto, LoteriaPrimitiva:
            ColIni = 6
            ColFin = 12
        Case gordoPrimitiva:
            ColIni = 7
            ColFin = 11
        Case Euromillones:
            ColIni = 7
            ColFin = 11
    End Select
    
    'Realizamos un bucle para todas las combinacionesxt h
     For Each m_rgFila In m_rgDatos.Rows
        ' Inicializar contador de Numeros encontrados
        j = 0
        'Comprueba los Numeros que se encuentran
        'entre las columnas E y K
        'y colorea de anaranjado si lo encuentra
        For i = ColIni To ColFin
            ' Inicializamos el color del número
            color = -1
            'Obtenemos el número de la celda
            m_num.Valor = m_rgFila.Cells(1, i).Value
            If (m_num.Valor > 0 And m_num.Valor < 50) Then
                'Comprueba el valor con los Numeros de la combinacion
                If vNewData.Combinacion.Contiene(m_num.Valor) Then
                    h = m_num.Terminacion
                        
                        Select Case h
                            Case 0: color = COLOR_TERMINACION0
                            Case 1: color = COLOR_TERMINACION1
                            Case 2: color = COLOR_TERMINACION2
                            Case 3: color = COLOR_TERMINACION3
                            Case 4: color = COLOR_TERMINACION4
                            Case 5: color = COLOR_TERMINACION5
                            Case 6: color = COLOR_TERMINACION6
                            Case 7: color = COLOR_TERMINACION7
                            Case 8: color = COLOR_TERMINACION8
                            Case 9: color = COLOR_TERMINACION9
                        End Select
                        j = j + 1
                End If
                
            End If
            'si el número es distinto de -1 la celda debe ser coloreada
            If (color > -1) Then
                ColoreaCelda m_rgFila.Cells(1, i), color
            End If
        Next i
    Next m_rgFila
    

   On Error GoTo 0
   Exit Sub

cmd_color_combinacion_Error:

   Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
   ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
   '   Audita el error
   Call HandleException(ErrNumber, ErrDescription, ErrSource, "Lot_02_Colorear.cmd_color_combinacion")
   '   Lanza el error
   Err.Raise ErrNumber, "Lot_02_Colorear.cmd_color_combinacion", ErrDescription
    
End Sub

'---------------------------------------------------------------------------------------
' Procedure : cmd_color_caracteristicas
' DateTime  : 12/jun/2007 23:42
' Author    : Carlos Almela Baeza
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub cmd_color_caracteristicas(vNewData As Integer)
'    Dim m_rgFila            As Range
'    Dim m_rgDatos           As Range
    Dim oSorteo             As Sorteo
'    Dim i                   As Integer
    Dim m_num               As New Numero
'    Dim color               As Integer
   
   On Error GoTo cmd_color_caracteristicas_Error
    '
    '   Elimina los colores de los resultados
    '
    Set m_rgDatos = DB.RangoResultados
    ColoreaCelda m_rgDatos, xlColorIndexNone
    
    Select Case JUEGO_DEFECTO
        Case Bonoloto, LoteriaPrimitiva:
            ColIni = 6
            ColFin = 12
        Case gordoPrimitiva:
            ColIni = 7
            ColFin = 11
        Case Euromillones:
            ColIni = 7
            ColFin = 11
    End Select

    '
    '   Creamos el objeto Sorteo
    '
    Set oSorteo = New Sorteo
    '
    '   Para cada fila en el rango de datos
    '
     For Each m_rgFila In m_rgDatos.Rows
        '
        '
        '
        oSorteo.Constructor m_rgFila
        '
        '   Verifica los Numeros situados entre las
        '   columnas D (6) y J(12)
        '
        For i = ColIni To ColFin
            '
            '   Inicializamos el color del número
            '
            color = -1
            '
            '   Obtenemos el número de la celda
            '
            m_num.Valor = m_rgFila.Cells(1, i).Value
            '
            '   Segun el tipo de modalidad seleccionada se llama a una función
            '
            Select Case vNewData
                Case 1: color = get_color_paridad(m_num)
                Case 2: color = get_color_peso(m_num)
                Case 3: color = get_color_decena(m_num)
                Case 4: color = get_color_terminacion(m_num)
                Case 5: color = get_color_continuo(m_num, oSorteo.Combinacion)
                
'                color = get_color_continuo(m_num.Valor, m_num_ant, m_num_sig, m_num_comp)
'                    m_num_ant = m_num.Valor
                    
            End Select
                                            
            'si el número es distinto de -1 la celda debe ser coloreada
            If (color > -1) Then
                ColoreaCelda m_rgFila.Cells(1, i), color
            End If
            '
            '   Siguiente Numero
            '
        Next i
        '
        '   Siguiente fila
        '
    Next m_rgFila
    
   On Error GoTo 0
   Exit Sub

cmd_color_caracteristicas_Error:

   Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
   ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
   '   Audita el error
   Call HandleException(ErrNumber, ErrDescription, ErrSource, "Lot_02_Colorear.cmd_color_caracteristicas")
   '   Lanza el error
   Err.Raise ErrNumber, "Lot_02_Colorear.cmd_color_caracteristicas", ErrDescription

End Sub


'---------------------------------------------------------------------------------------
' Procedure : get_color_continuo
' DateTime  : 12/jun/2007 23:42
' Author    : Carlos Almela Baeza
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Function get_color_continuo(vDataNum As Numero, vDataCol As Combinacion) As Integer
    Dim mNumerosCol()       As Integer
    Dim mColoresCol()       As Integer
    Dim mNum                As Numero
    Dim mSelColor           As Integer
    Dim i                   As Integer
    Dim j                   As Integer
    Dim mDif                As Integer
    Dim mDifAnt             As Integer
    
  On Error GoTo get_color_continuo_Error
    '
    '   Redimensionamos la matriz de números y colores
    '
    ReDim mNumerosCol(vDataCol.Count - 1)
    ReDim mColoresCol(vDataCol.Count - 1)
    '
    '   Obtenemos los números de la combinación
    '
    i = 0
    For Each mNum In vDataCol.Numeros
        '
        '   Nos guardamos el valor del Numero
        '
        mNumerosCol(i) = mNum.Valor
        mColoresCol(i) = 0
        '
        '   Incrementamos el indice
        '
        i = i + 1
    Next mNum
    '
    '   Ordenamos de forma ascendente
    '
    Ordenar mNumerosCol, True
    '
    '   Iniciamos el selector de color
    '
    mSelColor = 0
    '
    '   Marcamos los Numeros consecutivos de la combinación
    '
    For i = 0 To (UBound(mNumerosCol) - 1)
        '
        '   Calculamos la diferencia entre números
        '
        mDif = mNumerosCol(i + 1) - mNumerosCol(i)
        '
        '   Si la diferencia es 1 (son consecutivos)
        '
        If mDif = 1 Then
            '
            '   Si es el primer consecutivo inicializamos el color
            '
            If mSelColor = 0 Then
                mSelColor = 1
            End If
            '
            '   Seleccionamos el color para el Numero
            '
            mColoresCol(i) = mSelColor
            '
            '   Guardamos la diferencia
            '
            mDifAnt = mDif
        End If
        '
        '   Si la diferencia actual es distinta de 1
        '   y la diferencia anterior era 1 entonces este es
        '   el siguiente Numero consecutivo y cambiamos de color
        '
        If mDif > 1 And mDifAnt = 1 Then
            '
            '   marcamos el Numero
            '
            mColoresCol(i) = mSelColor
            mSelColor = mSelColor + 1
            mDifAnt = mDif
        End If
    Next i
    '
    '   El último Numero no se marca, comprobamos
    '   si era consecutivo
    '
    If mDifAnt = 1 Then
        mColoresCol(i) = mSelColor
    End If
    '
    '   Buscamos el Numero en la combinación
    '
    For i = 0 To UBound(mNumerosCol)
        '
        '   si es el número
        '
        If vDataNum.Valor = mNumerosCol(i) Then
            '
            '   Seleccionamos el valor del color
            '
            mSelColor = mColoresCol(i)
            Exit For
        End If
    Next i
    '
    '   Asignamos color
    '
    Select Case mSelColor
        Case 0: get_color_continuo = -1
        Case 1: get_color_continuo = COLOR_AÑIL
        Case 2: get_color_continuo = COLOR_AMARILLO
        Case 3: get_color_continuo = COLOR_ANARANJADO
    End Select
    '
    '   TODO: Trasladar la rutina a Combinación para ejecutarla una sola vez
    '         y color del Numero consecutivo al Numero
    '
    On Error GoTo 0
    Exit Function
    
get_color_continuo_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    '   Audita el error
    Call HandleException(ErrNumber, ErrDescription, "Lot_02_Colorear.get_color_continuo", ErrSource)
    '   Lanza el Error
    Err.Raise ErrNumber, "Lot_02_Colorear.get_color_continuo", ErrDescription
End Function

'---------------------------------------------------------------------------------------
' Procedure : get_color_paridad
' DateTime  : 12/jun/2007 23:42
' Author    : Carlos Almela Baeza
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Function get_color_paridad(vNewData As Numero) As Integer
    If (vNewData.Paridad = LT_PAR) Then
        get_color_paridad = COLOR_ANARANJADO
    Else
        get_color_paridad = COLOR_ROJO
    End If
End Function

'---------------------------------------------------------------------------------------
' Procedure : get_color_peso
' DateTime  : 12/jun/2007 23:42
' Author    : Carlos Almela Baeza
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Function get_color_peso(vNewData As Numero) As Integer
    If (vNewData.Peso = LT_BAJO) Then
        get_color_peso = COLOR_ANARANJADO
    Else
        get_color_peso = COLOR_ROJO
    End If
End Function

'---------------------------------------------------------------------------------------
' Procedure : get_color_decena
' DateTime  : 12/jun/2007 23:42
' Author    : Carlos Almela Baeza
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Function get_color_decena(vNewData As Numero) As Integer
    Select Case vNewData.Decena
        Case 1: get_color_decena = COLOR_DECENA1
        Case 2: get_color_decena = COLOR_DECENA2
        Case 3: get_color_decena = COLOR_DECENA3
        Case 4: get_color_decena = COLOR_DECENA4
        Case 5: get_color_decena = COLOR_DECENA5
    End Select
End Function

'---------------------------------------------------------------------------------------
' Procedure : get_color_terminacion
' DateTime  : 12/jun/2007 23:42
' Author    : Carlos Almela Baeza
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Function get_color_terminacion(vNewData As Numero) As Integer
     Select Case vNewData.Terminacion
        Case 0: get_color_terminacion = COLOR_TERMINACION0
        Case 1: get_color_terminacion = COLOR_TERMINACION1
        Case 2: get_color_terminacion = COLOR_TERMINACION2
        Case 3: get_color_terminacion = COLOR_TERMINACION3
        Case 4: get_color_terminacion = COLOR_TERMINACION4
        Case 5: get_color_terminacion = COLOR_TERMINACION5
        Case 6: get_color_terminacion = COLOR_TERMINACION6
        Case 7: get_color_terminacion = COLOR_TERMINACION7
        Case 8: get_color_terminacion = COLOR_TERMINACION8
        Case 9: get_color_terminacion = COLOR_TERMINACION9
    End Select
End Function

'---------------------------------------------------------------------------------------
' Procedure : Colorear_Matriz
' DateTime  : 14/ago/2007 08:14
' Author    : Carlos Almela Baeza
' Purpose   : Colorea un rango de datos con 6 colores según su valor
' Parameters: rgColumna - Rango de celdas a colorear (columna)
'             OrderBy - Orden de colorear: True  Ascendente
'                                          False Descendente
'---------------------------------------------------------------------------------------
'
Public Sub Colorear_Matriz(rgColumna As Range, Optional OrderBy As Boolean)
    Dim m_result As Variant             'Vector de resultados
    Dim m_datos As Variant              'Vector de datos
    Dim celda As Range                  'Celda a colorear
    Dim i As Integer                    'Indice Vector
    '
    '   Si el Orden de colorear no está definido asumimos Ascendente
    '
    If IsMissing(OrderBy) Then
        OrderBy = True                    ' Orden Ascendente menos a mas
    End If
    '
    '   Redimensionamos una matriz con los datos de la columna
    '
    ReDim m_datos(rgColumna.Cells.Count - 1)
    '
    '   Extraemos los valores de la columna
    '
    m_datos = rgColumna.Value
    '
    '   Obtenemos los colores para cada valor
    '
    m_result = Asignar_colores(m_datos, OrderBy)
    '
    '   Para cada celda en el rango
    '
    i = 0
    For Each celda In rgColumna
        '
        '   Coloreamos la celda con el color de la matriz
        '
        DestacarRango celda, CInt(m_result(i, 2))
        i = i + 1
    Next celda
    
End Sub
'---------------------------------------------------------------------------------------
' Procedure : Asignar_colores
' DateTime  : 14/ago/2007 08:22
' Author    : Carlos Almela Baeza
' Purpose   : Determina los colores asignados a cada valor
'---------------------------------------------------------------------------------------
'
Private Function Asignar_colores(vMatrizDatos As Variant, OrderBy As Boolean) As Variant
    Dim m_max           As Double       ' Elemento Maximo de la matriz
    Dim m_min           As Double       ' Elemento Minimo
    Dim m_dif           As Double
    Dim m_color         As Integer      '
    Dim m_vDatosColores() As Variant
    Dim m_valores       As Integer
    Dim i As Integer
    Dim j As Integer
    '
    '
    '   Redimensiona la matriz de colores con el número total de valores
    '
    m_valores = (UBound(vMatrizDatos) - LBound(vMatrizDatos)) + 1
    ReDim m_vDatosColores(m_valores, 2)
    '
    '    Calcular máximos y mínimos
    '
    m_max = 0
    m_min = 99999999999#
    For i = LBound(vMatrizDatos) To UBound(vMatrizDatos)
        If vMatrizDatos(i, 1) > m_max Then m_max = vMatrizDatos(i, 1)
        If vMatrizDatos(i, 1) < m_min Then m_min = vMatrizDatos(i, 1)
    Next i
    '
    '   Calculamos el diferencial para 6 tramos entre el máx y el mín
    '
    m_dif = (m_max - m_min) / 6
    '
    '
    '
    j = LBound(m_vDatosColores)
    '
    '   Si el orden es ascendente los rangos minimos son azules
    '
    If OrderBy Then
        '
        '       Asignamos colores segun el rango,
        '       para cada valor de la matriz de datos
        '
        For i = LBound(vMatrizDatos) To UBound(vMatrizDatos)
            Select Case (vMatrizDatos(i, 1))
                Case Is > (m_dif * 5) + m_min:  m_color = COLOR_ROJO
                Case Is > (m_dif * 4) + m_min:  m_color = COLOR_MARRON
                Case Is > (m_dif * 3) + m_min:  m_color = COLOR_AMARILLO
                Case Is > (m_dif * 2) + m_min:  m_color = COLOR_VERDE_CLARO
                Case Is > (m_dif * 1) + m_min:  m_color = COLOR_AÑIL
                Case Else:                      m_color = COLOR_AZUL_OSCURO
            End Select
            m_vDatosColores(j, 1) = vMatrizDatos(i, 1)      'Valor
            m_vDatosColores(j, 2) = m_color                 'Color asignado
            j = j + 1
        Next
    Else
    '
    '   Si el orden es descendente los rangos maximos son azules
    '
        For i = LBound(vMatrizDatos) To UBound(vMatrizDatos)
            Select Case (vMatrizDatos(i, 1))
                Case Is > (m_dif * 5) + m_min:  m_color = COLOR_AZUL_OSCURO
                Case Is > (m_dif * 4) + m_min:  m_color = COLOR_AÑIL
                Case Is > (m_dif * 3) + m_min:  m_color = COLOR_VERDE_CLARO
                Case Is > (m_dif * 2) + m_min:  m_color = COLOR_AMARILLO
                Case Is > (m_dif * 1) + m_min:  m_color = COLOR_MARRON
                Case Else:                      m_color = COLOR_ROJO
            End Select
            m_vDatosColores(j, 1) = vMatrizDatos(i, 1)      'Valor
            m_vDatosColores(j, 2) = m_color                 'Color asignado
            j = j + 1
        Next
    End If
    '
    '   Devolvemos la matriz
    '
    Asignar_colores = m_vDatosColores                   'Se devuelve la matriz de colores
End Function
'---------------------------------------------------------------------------------------
' Procedimiento : DestacarRango
' Creación      : 11/07/2006 23:28
' Autor         : Carlos Almela Baeza
' Objeto        : Esta rutina aplica una sombra a un rango determinado
'---------------------------------------------------------------------------------------
'
Public Sub DestacarRango(rng As Range, color As Integer)
    On Error Resume Next
    rng.Interior.ColorIndex = color
    If (rng.Interior.ColorIndex = COLOR_ROJO) Then rng.Font.ColorIndex = COLOR_AMARILLO
End Sub

'---------------------------------------------------------------------------------------
' Procedure : get_color_array
' DateTime  : 15/ago/2007 19:59
' Author    : Carlos Almela Baeza
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Function get_color_array(vDatos As Variant, n As Integer) As Integer
    Dim m_datos  As Variant             'Ordenación por número
    Dim m_result As Variant             'Vector de resultados
    Dim j As Integer
    Dim i As Integer
   On Error GoTo get_color_array_Error
    If n = 0 Or n > 49 Then
        get_color_array = xlNone
        Exit Function
    End If
    m_datos = vDatos
    m_result = Asignar_colores(m_datos, True)
    j = LBound(vDatos)
    For i = LBound(vDatos) To UBound(vDatos)
        If (vDatos(i, 0) = n) Then
            Exit For
        End If
        j = j + 1
    Next i
        
    get_color_array = m_result(j + 1, 2)

   On Error GoTo 0
   Exit Function

get_color_array_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure get_color_array of Módulo Lot_Colorear"
End Function

