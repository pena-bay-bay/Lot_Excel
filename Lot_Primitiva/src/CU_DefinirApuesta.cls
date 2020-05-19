VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CU_DefinirApuesta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'---------------------------------------------------------------------------------------
' Module    : CU_DefinirApuesta
' DateTime  : 12/ago/2007 19:13
' Author    : Carlos Almela Baeza
' Purpose   : Caso de uso que realiza el proceso de sugerir una apuestas
'---------------------------------------------------------------------------------------
Private m_objMuestra As Muestra
Private m_objMetodo As MetodoOld
'
'
Private Sub Class_Terminate()
    Set m_objMuestra = Nothing
    Set m_objMetodo = Nothing
End Sub
'
'
'
Public Property Get Muestra() As Muestra
    Set Muestra = m_objMuestra
End Property
Public Property Set Muestra(objMuestra As Muestra)
    Set m_objMuestra = objMuestra
End Property
'
'
'
Public Property Get metodo() As MetodoOld
    Set metodo = m_objMetodo
End Property
Public Property Set metodo(objMetodo As MetodoOld)
    Set m_objMetodo = objMetodo
End Property

'---------------------------------------------------------------------------------------
' Procedimiento : Get_Apuesta
' Creación      : 17-dic-2006 00:04
' Autor         : Carlos Almela Baeza
' Objeto        :
'---------------------------------------------------------------------------------------
'
Public Function Get_Apuesta() As ApuestaOld
    Dim objResult As ApuestaOld
  
   On Error GoTo Get_Apuesta_Error

    If m_objMetodo Is Nothing Then Err.Raise "falta el método"
    If m_objMuestra Is Nothing Then Err.Raise "falta la muestra"
    
    Set objResult = New Apuesta
    
    Select Case m_objMetodo.Tipo_Metodo
        Case ctAleatorio: Set objResult = metodo_Aleatorio()
'        Case ctBombo: Set objResult = metodo_Bombo()
        Case ctBombo: Set objResult = metodo_Aleatorio()
        Case ctProbabilidad: Set objResult = metodo_Apariciones()
        Case ctDecena: Set objResult = metodo_Decenas()
        Case ctTiempoMedio: Set objResult = metodo_tiempoMedio()
        Case ctFrecuencias: Set objResult = metodo_frecuencias()
        Case ctDecFrecuencias: Set objResult = metodo_decfrecuencias()
        Case ctFechaProxima: Set objResult = metodo_ultimaFecha()
        Case ctDecTiempoMedio: Set objResult = metodo_decTiempoMedio()
        Case ctAusencias: Set objResult = metodo_Ausencias()
    
    End Select
    Set objResult.metodo = m_objMetodo
    Set Get_Apuesta = objResult
    Set objResult = Nothing

   On Error GoTo 0
   Exit Function

Get_Apuesta_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & _
            ") in procedure Get_Apuesta of Módulo de clase CU_DefinirApuesta"
End Function
'
'
'
Private Function metodo_Bombo() As ApuestaOld
    Set objResult = New ApuestaOld
    Dim myBombo As bombo
    Set myBombo = New bombo
    With myBombo
        .bolas = 49
        .Numero_giros = 10
        .tipo_giro = lotGIROS
        .cargar (m_objMuestra.Matriz_Probabilidades)
        .girar
    End With
    
    Dim a_res() As Integer
    Dim nElem  As Long
    nElem = CLng(m_objMetodo.Pronosticos - 1)
    ReDim a_res(nElem)
    For i = 0 To nElem
            Set myBola = myBombo.extraccion
            a_res(i) = myBola.Valor
    Next i
    
    Ordenar a_res, True
    For i = 0 To UBound(a_res)
        
        Set objPronostico = New Pronostico
        Set objPronostico.metodo = m_objMetodo
        Set objPronostico.bola = m_objMuestra.Get_Bola(a_res(i))
        objResult.Add objPronostico
    Next i
    
    metodo_Bombo = objResult
End Function
'
'
'
Private Function metodo_Decenas() As ApuestaOld
    Dim a_mat(4) As Variant         'Bolas ordenadas por probabilidad
    Dim a_dec As Variant            'Orden de decenas
    Dim m_inum As Integer           'Numero seleccionado
    Dim objResult As ApuestaOld     ' Objeto Apuesta
    Dim objPronostico As Pronostico ' Objeto Pronostico
    Dim i As Integer                    ' Numero a tratar
    Dim j As Integer                    ' Indicador de la decena
    Dim x As Integer                    ' coordenada de decena más probable
    Dim y As Integer                    ' coordenada de nivel de decena más probable
    
    
   On Error GoTo metodo_Decenas_Error

   
    Set objResult = New ApuestaOld
    
    For i = 0 To 4
        a_mat(i) = m_objMuestra.Matriz_Decenas(i + 1, 0)
    Next i
    a_dec = m_objMuestra.Orden_Decenas
    
    x = 0
    y = 0
    'Bucle hasta completar el número de pronósticos
    For i = 0 To m_objMetodo.Pronosticos - 1
        ' Variación de la matriz de decenas
        ' Decenas      y   .   .   ->
        '  x   0..9    9   1   4   6   3   8   2   7   5
        '  .   10..19  16  17  19  11  15  12  13  14  18  10
        '  .   20..29  24  29  22  23  28  20  27  21  25  26
        '  |   30..39  35  36  30  34  33  39  32  37  31  38
        '  v   40..59  41  44  47  40  49  42  43  46  48  45
        '
        j = a_dec(x, 0) - 1         ' Obtiene el índice de la decena más probable
        m_inum = a_mat(j)(y, 0)     ' Obtiene el iesimo número dentro de la decena
                                    ' segun su probabilidad
        
        Set objPronostico = New Pronostico
        Set objPronostico.metodo = m_objMetodo
        Set objPronostico.bola = m_objMuestra.Get_Bola(m_inum)
        objResult.Combinacion.Add objPronostico.bola.Numero
        x = x + 1
        If (x = 4) Then             '
            x = 0
            y = y + 1
        End If
    Next i
    
    Set metodo_Decenas = objResult
   
   On Error GoTo 0
   Exit Function

metodo_Decenas_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & _
            ") in procedure metodo_Decenas of Módulo de clase CU_DefinirApuesta"
End Function
'---------------------------------------------------------------------------------------
' Procedimiento : metodo_Aleatorio
' Creación      : 27-feb-2007 23:39
' Autor         : Carlos Almela Baeza
' Objeto        :
'---------------------------------------------------------------------------------------
'
Private Function metodo_Aleatorio() As ApuestaOld
    Dim objResult As ApuestaOld
    Dim objPronostico As Pronostico
    Dim m_inum As Integer
   
   On Error GoTo metodo_Aleatorio_Error

    Set objResult = New ApuestaOld
    Do Until objResult.Total_Pronosticos = m_objMetodo.Pronosticos
        m_inum = Round(48 * Rnd()) + 1
        If Not objResult.Existe(m_inum) Then
            Set objPronostico = New Pronostico
            Set objPronostico.metodo = m_objMetodo
            Set objPronostico.bola = m_objMuestra.Get_Bola(m_inum)
            objResult.Combinacion.Add objPronostico.bola.Numero
        End If
    Loop
    Set metodo_Aleatorio = objResult

   On Error GoTo 0
   Exit Function

metodo_Aleatorio_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & _
    ") in procedure metodo_Aleatorio of Módulo de clase CU_DefinirApuesta"
End Function
'---------------------------------------------------------------------------------------
' Procedimiento : metodo_Apariciones
' Creación      : 27-feb-2007 23:40
' Autor         : Carlos Almela Baeza
' Objeto        :
'---------------------------------------------------------------------------------------
'
Private Function metodo_Apariciones() As ApuestaOld
    Dim a_mat As Variant
    Dim m_inum As Integer
    Dim objResult As ApuestaOld
    Dim objPronostico As Pronostico
   On Error GoTo metodo_Apariciones_Error

    Set objResult = New ApuestaOld
    
    a_mat = m_objMuestra.Matriz_Probabilidades
   
    For i = 0 To m_objMetodo.Pronosticos - 1
            m_inum = a_mat(i, 0)
            Set objPronostico = New Pronostico
            Set objPronostico.metodo = m_objMetodo
            Set objPronostico.bola = m_objMuestra.Get_Bola(m_inum)
            objResult.Combinacion.Add objPronostico.bola.Numero
    Next i
    Set metodo_Apariciones = objResult

   On Error GoTo 0
   Exit Function

metodo_Apariciones_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & _
    ") in procedure metodo_Apariciones of Módulo de clase CU_DefinirApuesta"
End Function

'---------------------------------------------------------------------------------------
' Procedimiento : metodo_tiempoMedio
' Creación      : 27-feb-2007 23:40
' Autor         : Carlos Almela Baeza
' Objeto        :
'---------------------------------------------------------------------------------------
'
Private Function metodo_tiempoMedio() As ApuestaOld
'    Dim a_mat(48, 1) As Double
'    Dim B_mat As Variant
    Dim a_mat As Variant
    Dim m_inum As Integer
    Dim objResult As ApuestaOld
      Dim objPronostico As Pronostico
   On Error GoTo metodo_tiempoMedio_Error

    Set objResult = New ApuestaOld
    a_mat = m_objMuestra.Matriz_ProbTiempos
    For i = 0 To m_objMetodo.Pronosticos - 1
            m_inum = a_mat(i, 0)
            Set objPronostico = New Pronostico
            Set objPronostico.metodo = m_objMetodo
            Set objPronostico.bola = m_objMuestra.Get_Bola(m_inum)
            objResult.Combinacion.Add objPronostico.bola.Numero
    Next i
    Set metodo_tiempoMedio = objResult

   On Error GoTo 0
   Exit Function

metodo_tiempoMedio_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & _
        ") in procedure metodo_tiempoMedio of Módulo de clase CU_DefinirApuesta"

End Function
'
'
'
'
Private Function metodo_frecuencias() As ApuestaOld
    Dim a_mat As Variant
    Dim m_inum As Integer
    Dim objResult As ApuestaOld
     Dim objPronostico As Pronostico
    Set objResult = New ApuestaOld
    
    a_mat = m_objMuestra.Matriz_ProbFrecuencias
   
    For i = 0 To m_objMetodo.Pronosticos - 1
            m_inum = a_mat(i, 0)
            Set objPronostico = New Pronostico
            Set objPronostico.metodo = m_objMetodo
            Set objPronostico.bola = m_objMuestra.Get_Bola(m_inum)
            objResult.Combinacion.Add objPronostico.bola.Numero
    Next i
    Set metodo_frecuencias = objResult
End Function
'
'
'
'
Private Function metodo_decTiempoMedio() As ApuestaOld
    Dim a_mat(4) As Variant             'Bolas ordenadas por fercuencias
    Dim a_dec As Variant                'Orden de decenas
    Dim m_inum As Integer               'Numero seleccionado
    Dim objResult As ApuestaOld            ' Objeto Apuesta
    Dim objPronostico As Pronostico     ' Objeto Pronostico
    Dim i As Integer                    ' Numero a tratar
    Dim j As Integer                    ' Indicador de la decena
    Dim x As Integer                    ' coordenada de decena más probable
    Dim y As Integer                    ' coordenada de nivel de decena más probable
    
    Set objResult = New ApuestaOld
    
    For i = 0 To 4
        a_mat(i) = m_objMuestra.Matriz_Decenas(i + 1, 1)
    Next i
    a_dec = m_objMuestra.Orden_Decenas
    
    x = 0
    y = 0
    'Bucle hasta completar el número de pronósticos
    For i = 0 To m_objMetodo.Pronosticos - 1
        ' Variación de la matriz de decenas
        ' Decenas      y   .   .   ->
        '  x   0..9    9   1   4   6   3   8   2   7   5
        '  .   10..19  16  17  19  11  15  12  13  14  18  10
        '  .   20..29  24  29  22  23  28  20  27  21  25  26
        '  |   30..39  35  36  30  34  33  39  32  37  31  38
        '  v   40..59  41  44  47  40  49  42  43  46  48  45
        '
        j = a_dec(x, 0) - 1         ' Obtiene el índice de la decena más probable
        m_inum = a_mat(j)(y, 0)     ' Obtiene el iesimo número dentro de la decena
                                    ' segun su probabilidad
        
        Set objPronostico = New Pronostico
        Set objPronostico.metodo = m_objMetodo
        Set objPronostico.bola = m_objMuestra.Get_Bola(m_inum)
        objResult.Combinacion.Add objPronostico.bola.Numero
        x = x + 1
        If (x > 4) Then
            x = 0
            y = y + 1
        End If
    Next i
    
    Set metodo_decTiempoMedio = objResult

End Function
'
'
'
'
Private Function metodo_decfrecuencias() As ApuestaOld
    Dim a_mat(4) As Variant             'Bolas ordenadas por fercuencias
    Dim a_dec As Variant                'Orden de decenas
    Dim m_inum As Integer               'Numero seleccionado
    Dim objResult As ApuestaOld            ' Objeto Apuesta
    Dim objPronostico As Pronostico     ' Objeto Pronostico
    Dim i As Integer                    ' Numero a tratar
    Dim j As Integer                    ' Indicador de la decena
    Dim x As Integer                    ' coordenada de decena más probable
    Dim y As Integer                    ' coordenada de nivel de decena más probable
    
    Set objResult = New ApuestaOld
    
    For i = 0 To 4
        a_mat(i) = m_objMuestra.Matriz_Decenas(i + 1, 2)
    Next i
    a_dec = m_objMuestra.Orden_Decenas
    
    x = 0
    y = 0
    'Bucle hasta completar el número de pronósticos
    For i = 0 To m_objMetodo.Pronosticos - 1
        ' Variación de la matriz de decenas
        ' Decenas      y   .   .   ->
        '  x   0..9    9   1   4   6   3   8   2   7   5
        '  .   10..19  16  17  19  11  15  12  13  14  18  10
        '  .   20..29  24  29  22  23  28  20  27  21  25  26
        '  |   30..39  35  36  30  34  33  39  32  37  31  38
        '  v   40..59  41  44  47  40  49  42  43  46  48  45
        '
        j = a_dec(x, 0) - 1         ' Obtiene el índice de la decena más probable
        m_inum = a_mat(j)(y, 0)     ' Obtiene el iesimo número dentro de la decena
                                    ' segun su probabilidad
        
        Set objPronostico = New Pronostico
        Set objPronostico.metodo = m_objMetodo
        Set objPronostico.bola = m_objMuestra.Get_Bola(m_inum)
        objResult.Combinacion.Add objPronostico.bola.Numero
        x = x + 1
        If (x > 4) Then
            x = 0
            y = y + 1
        End If
    Next i
    
    Set metodo_decfrecuencias = objResult
   

End Function
'
'
'
'
'
Private Function metodo_ultimaFecha() As ApuestaOld
    Dim a_Vec(48, 1) As Double
    Dim m_bola As bola
    Dim x As Integer
    Dim i As Integer
    Dim objResult As ApuestaOld            ' Objeto Apuesta
    Dim objPronostico As Pronostico     ' Objeto Pronostico
    
    On Error GoTo metodo_ultimaFecha_Error:
    x = 0
    Set objResult = New ApuestaOld
    
    For i = 1 To 49
        Set m_bola = m_objMuestra.Get_Bola(i)
        a_Vec(i - 1, 0) = CDbl(i)
        a_Vec(i - 1, 1) = m_bola.ProximaFecha
        Set m_bola = Nothing
    Next i
    
    Ordenar2 a_Vec, 2, False
    
    For i = 0 To m_objMetodo.Pronosticos - 1
        x = a_Vec(i, 0)
        Set objPronostico = New Pronostico
        Set objPronostico.metodo = m_objMetodo
        Set objPronostico.bola = m_objMuestra.Get_Bola(x)
        objResult.Combinacion.Add objPronostico.bola.Numero
    Next i
   
    Set metodo_ultimaFecha = objResult
    
metodo_ultimaFecha_exit:
        Exit Function
        
metodo_ultimaFecha_Error:
        Dim lErrNumber              As Long
        Dim strErrSource            As String
        Dim strErrDescription       As String
        lErrNumber = Err.Number
        strErrDescription = Err.Description
        strErrSource = "metodo_ultimaFecha"
        Debug.Print lErrNumber, strErrDescription, strErrSource, CU_DefinirApuesta
End Function
'
'
'
'
'

Private Function metodo_Septenas() As ApuestaOld
    Dim a_mat(7) As Variant         'Bolas ordenadas por probabilidad
    Dim a_sep As Variant            'Orden de septenas
    Dim m_inum As Integer           'Numero seleccionado
    Dim objResult As ApuestaOld        ' Objeto Apuesta
    Dim objPronostico As Pronostico ' Objeto Pronostico
    Dim i As Integer                    ' Numero a tratar
    Dim j As Integer                    ' Indicador de la decena
    Dim x As Integer                    ' coordenada de decena más probable
    Dim y As Integer                    ' coordenada de nivel de decena más probable
    
    
   On Error GoTo metodo_Septenas_Error

   
    Set objResult = New ApuestaOld
    
    For i = 0 To 7
'        a_mat(i) = m_objMuestra.Matriz_Septenas(i + 1, 0)
    Next i
'    a_sep = m_objMuestra.Orden_Septenas
    
    x = 0
    y = 0
    'Bucle hasta completar el número de pronósticos
    For i = 0 To m_objMetodo.Pronosticos - 1
        ' Variación de la matriz de decenas
        ' Decenas      y   .   .   ->
        '  x   00..07   9   1   4   6   3   8
        '  .   08..14  16  17  19  11  15  12
        '  .   15..21  24  29  22  23  28  20
        '  |   22..28  35  36  30  34  33  39
        '  v   29..35  41  44  47  40  49  42  43  46  48  45
        '
        j = a_sep(x, 0) - 1         ' Obtiene el índice de la decena más probable
        m_inum = a_mat(j)(y, 0)     ' Obtiene el iesimo número dentro de la decena
                                    ' segun su probabilidad
        
        Set objPronostico = New Pronostico
        Set objPronostico.metodo = m_objMetodo
        Set objPronostico.bola = m_objMuestra.Get_Bola(m_inum)
        objResult.Combinacion.Add objPronostico.bola.Numero
        x = x + 1
        If (x = 4) Then             '
            x = 0
            y = y + 1
        End If
    Next i
    
    Set metodo_Septenas = objResult
   
   On Error GoTo 0
   Exit Function

metodo_Septenas_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & _
            ") in procedure metodo_Septenas of Módulo de clase CU_DefinirApuesta"
End Function


Private Function metodo_Ausencias()

    Dim a_Vec(49, 1) As Double
    Dim m_bola As bola
    Dim x As Integer
    Dim i As Integer
    Dim y As Integer
    Dim z As Integer
    Dim objResult As ApuestaOld            ' Objeto Apuesta
    Dim objPronostico As Pronostico     ' Objeto Pronostico
    Dim a_VecRangos(49, 2) As Double
    x = 0
    Set objResult = New ApuestaOld
    
    For i = 1 To 49
        Set m_bola = m_objMuestra.Get_Bola(i)
        a_Vec(i - 1, 0) = CDbl(i)
        a_Vec(i - 1, 1) = m_bola.Ausencias
        Set m_bola = Nothing
    Next i
    
    Ordenar2 a_Vec, 2
    '
    '   Matriz ordenada de Mas a menor ausencias
    '   Trocear matriz en tres vectores de trea áreas
    '
    x = 0: y = 0: z = 0
    For i = 0 To 48
        Select Case a_Vec(i, 1)
        Case 0 To 3
            a_VecRangos(x, 0) = a_Vec(i, 0)
            x = x + 1
        Case 4 To 10
            a_VecRangos(y, 1) = a_Vec(i, 0)
            y = y + 1
        Case Else
            a_VecRangos(z, 2) = a_Vec(i, 0)
            z = z + 1
        
        End Select
    Next i
    
    x = 0: y = 0: z = 0
    For i = 0 To m_objMetodo.Pronosticos - 1
        Do Until x > 0
            x = a_VecRangos(y, z)
            z = z + 1
            If z = 3 Then
                z = 0
                y = y + 1
            End If
        Loop
        Set objPronostico = New Pronostico
        Set objPronostico.metodo = m_objMetodo
        Set objPronostico.bola = m_objMuestra.Get_Bola(x)
        objResult.Combinacion.Add objPronostico.bola.Numero
        x = 0
    Next i
   
    Set metodo_Ausencias = objResult

End Function


