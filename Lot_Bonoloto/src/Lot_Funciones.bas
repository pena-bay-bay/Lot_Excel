Attribute VB_Name = "Lot_Funciones"
'
'       Módulo de funciones comunes a la librería LOT_Libreria
'
Option Explicit

'---------------------------------------------------------------------------------------
' Procedimiento : Colorea_Celda
' Creación      : 16-dic-2006 21:14
' Autor         : Carlos Almela Baeza
' Objeto        : Colorea la celda según el número y la muestra a aplicar
'---------------------------------------------------------------------------------------
'
Public Sub Colorea_Celda(celda As Range, _
                          Numero As Variant, _
                          ByRef objMuestra As Muestra, _
                          objMetodo As Metodo)
    Dim i        As Integer                     'Número en formato Entero
    Dim m_iColor As Integer                     'Color de la celda
   On Error GoTo Colorea_Celda_Error
    i = CInt(Numero)                            'Obtiene el entero del número
    m_iColor = xlNone                           'Inicializa el color a automático
    
    'Selección de la matriz de números según el método
    
    
    Select Case (objMetodo.CriteriosOrdenacion)

'        Case ordSinDefinir

'        Case ordProbabilidad

        Case ordProbTiempoMedio:
            m_iColor = get_color_array(objMuestra.Matriz_ProbTiempos, i)

        Case ordFrecuencia:
            m_iColor = get_color_array(objMuestra.Matriz_ProbFrecuencias, i)
        
'        Case ordAusencia

'        Case ordTiempoMedio

'        Case ordDesviacion

'        Case ordProximaFecha

'        Case ordModa

        Case Else: m_iColor = get_color_array(objMuestra.Matriz_Probabilidades, i)
    End Select
    
    celda.Value = Numero                       'Asigna el Numero a la celda
    DestacarRango celda, m_iColor              'Colorea la celda

   On Error GoTo 0
   Exit Sub

Colorea_Celda_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & _
        ") in procedure Colorea_Celda of Módulo Lot_VerSorteos"
    
End Sub
'---------------------------------------------------------------------------------------
' Procedimiento : Colorea_Celda
' Creación      : 16-dic-2006 21:14
' Autor         : Carlos Almela Baeza
' Objeto        : Colorea la celda según el número y la muestra a aplicar
'---------------------------------------------------------------------------------------
'
Public Sub Colorea_CeldaProb(celda As Range, _
                          Numero As Variant, _
                          ByRef objMuestra As Muestra)
                          
    Dim i        As Integer                     'Número en formato Entero
    Dim m_iColor As Integer                     'Color de la celda
   On Error GoTo Colorea_Celda_Error
    i = CInt(Numero)                            'Obtiene el entero del número
    m_iColor = xlNone                           'Inicializa el color a automático
    
    m_iColor = get_color_array(objMuestra.Matriz_Probabilidades, i)
    
    celda.Value = Numero                       'Asigna el Numero a la celda
    DestacarRango celda, m_iColor              'Colorea la celda

   On Error GoTo 0
   Exit Sub

Colorea_Celda_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & _
        ") in procedure Colorea_Celda of Módulo Lot_VerSorteos"
    
End Sub

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

Public Function GetFechaRegistro(datRegistro As Integer, Optional datJuego As Juego = PrimitivaBonoloto)
    Dim maxRegistro     As Integer
    Dim maxFecha        As Date
    Dim m_dtFecha       As Date
    Dim mDB             As BdDatos
    '
    ' Obtengo los datos del último registro
    '
    Set mDB = New BdDatos
    
    maxFecha = mDB.UltimoResultado
    maxRegistro = mDB.UltimoRegistro
    '
    ' Si es inferior al máximo
    '
    Select Case datRegistro
    Case Is = maxRegistro
        m_dtFecha = maxFecha
    Case Else
        m_dtFecha = mDB.GetSimulacionFecha(datRegistro)
    End Select
    GetFechaRegistro = m_dtFecha
End Function


Public Function GetRegistroFecha(datFecha As Date, Optional datJuego As Juego = PrimitivaBonoloto)
    Dim mRes            As Resultado
    Dim maxRegistro     As Integer
    Dim maxFecha        As Date
    Dim mIDifDias       As Integer
    Dim iDiaSem         As Integer
    Dim mDB             As BdDatos
    
    
    Set mDB = New BdDatos
    
    maxFecha = mDB.UltimoResultado
    maxRegistro = mDB.UltimoRegistro
    
    If (datFecha < maxFecha) Then
        mRes = mDB.Get_Resultado(datFecha)
        GetRegistroFecha = mRes
    Else
        mIDifDias = datFecha - maxFecha
    End If
    iDiaSem = Weekday(datFecha, vbMonday)
    Select Case datJuego
    ' L, M, X, J, V, S
    Case PrimitivaBonoloto
    ' L, M, X, V
    Case Juego.Bonoloto
    ' J, S
    Case Juego.LoteriaPrimitiva
    ' M, V
    Case Juego.Euromillones
    ' D
    Case Juego.gordoPrimitiva
    
    End Select

End Function

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




