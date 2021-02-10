Attribute VB_Name = "Lot_Funciones"
' *============================================================================*
' *
' *     Fichero    : Lot_Funciones.bas
' *
' *     Autor      : Carlos Almela Baeza
' *     Creación   : ju., 08/ago/2019 19:08:13
' *     Versión    : 2.0
' *     Propósito  : Recopila las funciones mas utilizadas en la aplicación
' *
' *============================================================================*
Option Explicit
Option Base 0
'---------------------------------------------------------------------------------------
' Procedimiento : Colorea_Celda
' Creación      : 16-dic-2006 21:14
' Autor         : Carlos Almela Baeza
' Objeto        : Colorea la celda según el número y la muestra a aplicar
'---------------------------------------------------------------------------------------
'
'Public Sub Colorea_Celda(celda As Range, _
'                          Numero As Variant, _
'                          ByRef objMuestra As Muestra, _
'                          objMetodo As metodo)
'    Dim i        As Integer                     'Número en formato Entero
'    Dim m_iColor As Integer                     'Color de la celda
'   On Error GoTo Colorea_Celda_Error
'    i = CInt(Numero)                            'Obtiene el entero del número
'    m_iColor = xlNone                           'Inicializa el color a automático
'
'    'Selección de la matriz de números según el método
'
'
'    Select Case (objMetodo.Parametros.CriteriosOrdenacion)
'
''        Case ordSinDefinir
'
''        Case ordProbabilidad
'
'        Case ordProbTiempoMedio:
'            m_iColor = get_color_array(objMuestra.Matriz_ProbTiempos, i)
'
'        Case ordFrecuencia:
'            m_iColor = get_color_array(objMuestra.Matriz_ProbFrecuencias, i)
'
''        Case ordAusencia
'
''        Case ordTiempoMedio
'
''        Case ordDesviacion
'
''        Case ordProximaFecha
'
''        Case ordModa
'
'        Case Else: m_iColor = get_color_array(objMuestra.Matriz_Probabilidades, i)
'    End Select
'
'    celda.Value = Numero                       'Asigna el Numero a la celda
'    DestacarRango celda, m_iColor              'Colorea la celda
'
'   On Error GoTo 0
'   Exit Sub
'
'Colorea_Celda_Error:
'
'    MsgBox "Error " & Err.Number & " (" & Err.Description & _
'        ") in procedure Colorea_Celda of Módulo Lot_VerSorteos"
'
'End Sub
'---------------------------------------------------------------------------------------
' Procedimiento : Colorea_Celda
' Creación      : 16-dic-2006 21:14
' Autor         : Carlos Almela Baeza
' Objeto        : Colorea la celda según el número y la muestra a aplicar
'---------------------------------------------------------------------------------------
'
'Public Sub Colorea_CeldaProb(celda As Range, _
'                          Numero As Variant, _
'                          ByRef objMuestra As Muestra)
'
'    Dim i        As Integer                     'Número en formato Entero
'    Dim m_iColor As Integer                     'Color de la celda
'   On Error GoTo Colorea_Celda_Error
'    i = CInt(Numero)                            'Obtiene el entero del número
'    m_iColor = xlNone                           'Inicializa el color a automático
'
'    m_iColor = get_color_array(objMuestra.Matriz_Probabilidades, i)
'
'    celda.Value = Numero                       'Asigna el Numero a la celda
'    DestacarRango celda, m_iColor              'Colorea la celda
'
'   On Error GoTo 0
'   Exit Sub
'
'Colorea_Celda_Error:
'
'    MsgBox "Error " & Err.Number & " (" & Err.Description & _
'        ") in procedure Colorea_Celda of Módulo Lot_VerSorteos"
'
'End Sub

'---------------------------------------------------------------------------------------
' Procedimiento : ColoreaCelda
' Creación      : 09-nov-2006 10:09
' Autor         : Carlos Almela Baeza
' Objeto        : Colorea una celda o rango de celdas
'---------------------------------------------------------------------------------------
'
Public Sub ColoreaCelda(cldTarget As Range, color As Integer)
        If (color = COLOR_TERMINACION0) Then
            cldTarget.Font.ColorIndex = COLOR_NUMCOMPLE
        Else
            cldTarget.Font.ColorIndex = xlColorIndexAutomatic
        End If
        cldTarget.Interior.ColorIndex = color
End Sub

'Public Function GetFechaRegistro(datRegistro As Integer, Optional datJuego As Juego = PrimitivaBonoloto)
'    Dim maxRegistro     As Integer
'    Dim maxFecha        As Date
'    Dim m_dtFecha       As Date
'    Dim mDB             As BdDatos
'    '
'    ' Obtengo los datos del último registro
'    '
'    Set mDB = New BdDatos
'
'    maxFecha = mDB.UltimoResultado
'    maxRegistro = mDB.UltimoRegistro
'    '
'    ' Si es inferior al máximo
'    '
'    Select Case datRegistro
'    Case Is = maxRegistro
'        m_dtFecha = maxFecha
'    Case Else
'        m_dtFecha = mDB.GetSimulacionFecha(datRegistro)
'    End Select
'    GetFechaRegistro = m_dtFecha
'End Function


'Public Function GetRegistroFecha(datFecha As Date, Optional datJuego As Juego = PrimitivaBonoloto)
'    Dim mRes            As Resultado
'    Dim maxRegistro     As Integer
'    Dim maxFecha        As Date
'    Dim mIDifDias       As Integer
'    Dim iDiaSem         As Integer
'    Dim mDB             As BdDatos
'
'
'    Set mDB = New BdDatos
'
'    maxFecha = mDB.UltimoResultado
'    maxRegistro = mDB.UltimoRegistro
'
'    If (datFecha < maxFecha) Then
'        mRes = mDB.Get_Resultado(datFecha)
'        GetRegistroFecha = mRes
'    Else
'        mIDifDias = datFecha - maxFecha
'    End If
'    iDiaSem = Weekday(datFecha, vbMonday)
'    Select Case datJuego
'    ' L, M, X, J, V, S
'    Case PrimitivaBonoloto
'    ' L, M, X, V
'    Case Juego.Bonoloto
'    ' J, S
'    Case Juego.LoteriaPrimitiva
'    ' M, V
'    Case Juego.Euromillones
'    ' D
'    Case Juego.gordoPrimitiva
'
'    End Select
'
'End Function

Public Function GetModa(datValores As Variant) As Double
    Dim mResult As Double
    On Error Resume Next
    mResult = Application.WorksheetFunction.Mode(datValores)
    If Err.Number <> 0 Then
        mResult = Application.WorksheetFunction.Median(datValores)
    End If
    GetModa = mResult
End Function

'---------------------------------------------------------------------------------------
' Procedure : Version_Libreria
' Author    : CHARLY
' Date      : sáb, 14/01/2012 19:50
' Purpose   : Visualiza la version de las macros
'---------------------------------------------------------------------------------------
'
Public Sub Version_Libreria()
    Dim Version As String
    Version = " La versión de la librería es la:" & vbCrLf _
              & vbTab & lotVersion & vbCrLf _
              & "de fecha " & vbTab & lotFeVersion
    MsgBox Version, vbApplicationModal + vbInformation + vbOKOnly, "Librería de Funciones de la Loteria"
End Sub

'---------------------------------------------------------------------------------------
' Modulo    : fn_collections
' Creado    : 25/05/2004 23:04
' Autor     : Carlos Almela Baeza
' Version   : 1.0.0 08-dic-2006 20:55
' Objeto    : Módulo de manejo de colecciones y arrays
'---------------------------------------------------------------------------------------
'
'---------------------------------------------------------------------------------------
' Procedimiento : ExistenElem
' Creación      : 12-nov-2006 00:38
' Autor         : Carlos Almela Baeza
' Objeto        :
'---------------------------------------------------------------------------------------
'
Public Function ExistenElem(col As Collection, Clave As String) As Boolean
    Dim prueba As Variant
    On Error Resume Next
    prueba = col.Item(Clave)
    ExistenElem = (Err <> 5)
End Function
'---------------------------------------------------------------------------------------
' Procedimiento : EliminarTodosElementos
' Creación      : 12-nov-2006 00:38
' Autor         : Carlos Almela Baeza
' Objeto        :
'---------------------------------------------------------------------------------------
'
Public Sub EliminarTodosElementos(col As Collection)
    Do While col.Count
        col.Remove 1
    Loop
End Sub
'---------------------------------------------------------------------------------------
' Procedimiento : SustituirElem
' Creación      : 12-nov-2006 00:37
' Autor         : Carlos Almela Baeza
' Objeto        :
'---------------------------------------------------------------------------------------
'
Public Sub SustituirElem(col As Collection, indice As Variant, nuevoValor As Variant)
    col.Remove indice
    If VarType(indice) = vbString Then
        col.Add nuevoValor, indice
    Else
        If indice > col.Count Then
            col.Add nuevoValor, , col.Count
        Else
            col.Add nuevoValor, , indice
        End If
    End If
End Sub
'---------------------------------------------------------------------------------------
' Procedimiento : ShellSortAny
' Creación      : 12-nov-2006 00:37
' Autor         : Carlos Almela Baeza
' Objeto        :
'---------------------------------------------------------------------------------------
'
Public Sub ShellSortAny(arr As Variant, numEls As Long, descendente As Boolean)
    Dim indice As Long, indice2 As Long, primerElem As Long
    Dim distancia As Long, Valor As Variant
    
   On Error GoTo ShellSortAny_Error
    
    ' salir si no es un array
    If VarType(arr) < vbArray Then Exit Sub
    
    primerElem = LBound(arr)
    
    ' encontrar el mejor valor para distancia
    Do
        distancia = distancia * 3 + 1
    Loop Until distancia > numEls
    
    ' ordenar el array
    Do
        distancia = distancia / 3
        For indice = distancia + primerElem To numEls + primerElem - 1
            Valor = arr(indice)
            indice2 = indice
            Do While (arr(indice2 - distancia) > Valor) Xor descendente
                arr(indice2) = arr(indice2 - distancia)
                indice2 = indice2 - distancia
                If indice2 - distancia < primerElem Then Exit Do
            Loop
            arr(indice2) = Valor
        Next
    Loop Until distancia = 3

   On Error GoTo 0
   Exit Sub

ShellSortAny_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & _
        ") in procedure ShellSortAny of Módulo fn_collections"
End Sub

'---------------------------------------------------------------------------------------
' Procedimiento : Ordenar
' Creación      : 30-Oct-2002 00:40
' Autor         : Carlos Almela Baeza
' Objeto        : Ordenar una matriz de una sola dimensión
'---------------------------------------------------------------------------------------
'
Public Sub Ordenar(ByRef matriz As Variant, Ascendente As Boolean)
    Dim tsOrdenado  As Boolean          'indicador de matriz ordenada
    Dim MxLimite    As Integer          'Limite máximo del bucle
    Dim TmpDato     As Variant          'Dato temporal para el intercambio
    Dim i           As Integer          'Indice
    
    'Si no es un elemento array sale de la rutina
    If Not IsArray(matriz) Then Exit Sub
    
    'obtiene el número máximo de elementos de la matriz
    MxLimite = UBound(matriz)
    
    'Bucle de ordenación, se realiza hasta que esté ordenada
    Do
        tsOrdenado = True               'Se parte de matriz ordenada
        For i = 0 To MxLimite - 1       'Se revisa cada elemento con el siguiente
            If Ascendente Then
                If matriz(i) > matriz(i + 1) _
                And (matriz(i + 1) <> 0) Then
                    tsOrdenado = False
                    TmpDato = matriz(i)         'Guardamos la posicion iesima
                    matriz(i) = matriz(i + 1)   'pasamos la posision siguiente a la
                                                'iesima
                    matriz(i + 1) = TmpDato     'pasamos el dato guardado a la siguiente
                End If
            Else
                If matriz(i) < matriz(i + 1) _
                And (matriz(i + 1) <> 0) Then
                    tsOrdenado = False
                    TmpDato = matriz(i)         'Guardamos la posicion iesima
                    matriz(i) = matriz(i + 1)   'pasamos la posision siguiente a la
                                                'iesima
                    matriz(i + 1) = TmpDato     'pasamos el dato guardado a la siguiente
                End If
            End If
        Next i
    Loop Until tsOrdenado
End Sub
'------------------------------------------------------------------------------*
'Función     : Ordenar2
'Fecha       : 28-Nov-1999
'Parametros  : Matriz de dos dimensiones
'Descripción : Ordena la matriz de paso
'                     x ->
'              matriz (0, 0) (0, 1)
'             y  |    (1, 0) (1, 1)
'                v    (2, 0) (2, 1)
'------------------------------------------------------------------------------*
Public Sub Ordenar2(ByRef matriz As Variant, Optional columna As Integer = 2, Optional Ascendente As Boolean = True)
    Dim tsOrdenado As Boolean
    Dim limiteY As Integer
    Dim limiteX As Integer
    Dim TmpDato() As Variant
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    
   On Error GoTo Ordenar2_Error

    limiteY = UBound(matriz, 1)
    limiteX = UBound(matriz, 2)
    ReDim TmpDato(limiteX)
    j = IIf(columna = 1, 0, 1)
    If Not IsArray(matriz) Then Exit Sub
    
    Do
        tsOrdenado = True
        For i = 0 To limiteY - 1
            If Ascendente Then
                If (matriz(i, j) < matriz(i + 1, j)) Then
                    tsOrdenado = False
                    For k = 0 To limiteX
                      TmpDato(k) = matriz(i, k)
                      matriz(i, k) = matriz(i + 1, k)
                      matriz(i + 1, k) = TmpDato(k)
                    Next k
                End If
            Else
                If (matriz(i, j) > matriz(i + 1, j)) Then
                    tsOrdenado = False
                    For k = 0 To limiteX
                      TmpDato(k) = matriz(i, k)
                      matriz(i, k) = matriz(i + 1, k)
                      matriz(i + 1, k) = TmpDato(k)
                    Next k
                End If
            End If
        Next i
    Loop Until tsOrdenado

   On Error GoTo 0
   Exit Sub

Ordenar2_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Ordenar2 of Módulo fn_collections"
End Sub





