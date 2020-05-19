VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Muestra"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'---------------------------------------------------------------------------------------
' Modulo    : Muestra
' Creado    : 03-may-2006 22:41
' Autor     : Carlos Almela Baeza
' Version   : 1.1.01 15/12/2006 8:56
' Objeto    :
'---------------------------------------------------------------------------------------
Option Explicit
Option Base 0

Private mParametrosMuestra  As ParametrosMuestra        ' Parametros de la muestra
Private m_col                   As Collection               ' Colección de bolas
Private m_bol                   As bola                     ' Bola de trabajo
Private mRegTotal               As Integer                  ' Total de registros
Private mTotalNumeros           As Integer                  ' Total de números analizados
Private m_Estrellas             As Collection               ' Colección de Estrellas
Private m_Reintegros            As Collection               ' Colección de Reintegros

'---------------------------------------------------------------------------------------
' Procedimiento : Class_Initialize
' Creación      : 12-nov-2006 00:31
' Autor         : Carlos Almela Baeza
' Objeto        :
'---------------------------------------------------------------------------------------
'
Private Sub Class_Initialize()
    Set m_col = New Collection                              ' Se crea una nueva colección
    Set m_Estrellas = New Collection                        ' Se crea una colección de Estrellas
    Set mParametrosMuestra = New ParametrosMuestra      ' Se inicia la clase de parametros
    Set m_bol = New bola                                    ' Se inicializa la Bola
                                                            ' Se define la matriz de frecuencias
'    Rango_Frecuencias = Array(1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 20, 50, 100)
End Sub

Private Sub Class_Terminate()
    Set m_col = Nothing                         'Se elimina la colección de números
    Set mParametrosMuestra = Nothing        'Se elimina los parametros de la muestra
    Set m_bol = Nothing                         'Se elimina
End Sub

'---------------------------------------------------------------------------------------
' Procedimiento : Constructor
' Creación      : 12-nov-2006 00:23
' Autor         : Carlos Almela Baeza
' Objeto        :
'---------------------------------------------------------------------------------------
Public Sub Constructor(rg_muestra As Range, vNewJuego As Juego)
        Dim i               As Integer
        Dim mStrKey         As String
        Dim rgRow           As Range
        Dim objSorteo       As New Sorteo
        Dim objNum          As New Numero
        Dim nRegAnalisis    As Integer
        Dim oInfoJuego      As InfoSorteo
  On Error GoTo Constructor_Error
        '
        '   Borra la coleccion de Numeros
        '
        If m_col.Count > 0 Then
            EliminarTodosElementos m_col
        End If
        mRegTotal = 0
        '
        '   Se crea la colección de bolas en vacio
        '   TODO: En función del juego se crearán 49, 50, 54 y 10 estrellas o clave
        '
        For i = 1 To 49
            Set m_bol = New bola
            With m_bol
                .Constructor (i)
            End With
            mStrKey = CStr(m_bol.Numero.Valor)
            m_col.Add m_bol, mStrKey
        Next i
        '
        '   Se actualizan las bolas con la información de la muestra de sorteos
        '
        For Each rgRow In rg_muestra.Rows
            objSorteo.Constructor rgRow
            '
            '  para cada Numero en la combinación
            '
            For Each objNum In objSorteo.Combinacion.Numeros
                '
                '   Obtiene la clave del Numero
                '
                mStrKey = CStr(objNum.Valor)
                '
                '   Con la bola del Numero actualiza propiedades
                '
                With m_col.Item(mStrKey)
                    '
                    '   Agrega Numero de registro
                    '
                    i = .colRegistros.Count
                    If i = 0 Then
                        Set .colRegistros = New Collection
                        .colRegistros.Add objSorteo.Registro, CStr(i)
                    Else
                        .colRegistros.Add objSorteo.Registro, CStr(i), , i
                    End If
                    '
                    '  Agrega fecha de aparicion
                    '
                    i = .colFechasAparicion.Count
                    If i = 0 Then
                        Set .colFechasAparicion = New Collection
                        .colFechasAparicion.Add objSorteo.Fecha, CStr(i)
                    Else
                        .colFechasAparicion.Add objSorteo.Fecha, CStr(i), , i
                    End If
                    mRegTotal = mRegTotal + 1
                End With
            Next objNum
        Next rgRow
        '
        '
        '
        'nRegAnalisis = oInfoJuego.
        nRegAnalisis = mParametrosMuestra.ResgistroAnalisis
        '
        '  actualiza las estadisticas de todas las bolas
        '
        For i = 1 To 49
            mStrKey = CStr(i)
            m_col.Item(mStrKey).Actualizar (nRegAnalisis)
            m_col.Item(mStrKey).Total = mRegTotal
            
        Next i
        '
        '   Calcula los colores de cada bola
        '
        CalcularColores
        '
        '   Calcula valor Homogeneo
        '
        CalcularValorHomogeneo
        
Constructor_CleanExit:
   On Error GoTo 0
    Exit Sub

Constructor_Error:
    

    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    '   Audita el error
    Call HandleException(ErrNumber, ErrDescription, "Muestra.Constructor")
    '   Lanza el Error
    Err.Raise ErrNumber, ErrSource, ErrDescription
End Sub


'---------------------------------------------------------------------------------------
' Procedure : Asignar_colores
' DateTime  : 14/ago/2007 08:22
' Author    : Carlos Almela Baeza
' Purpose   : Determina los colores asignados a cada valor
'---------------------------------------------------------------------------------------
'
Private Function AsignarColores(vMatrizDatos As Variant) As Variant
    Dim m_max As Double                     'Elemento Maximo de la matriz
    Dim m_min As Double
    Dim m_dif As Double
    Dim m_color As Integer
    Dim m_vDatosColores() As Variant
    Dim m_valores As Integer
    Dim i As Integer
    Dim j As Integer
    'Redimensiona la matriz de colores con el número total de valores
    m_valores = (UBound(vMatrizDatos) - LBound(vMatrizDatos)) + 1
    ReDim m_vDatosColores(m_valores, 2)
    
    'Calcular máximos y mínimos
    m_max = 0
    m_min = 99999999999#
    For i = LBound(vMatrizDatos) To UBound(vMatrizDatos)
        If vMatrizDatos(i, 1) > m_max Then m_max = vMatrizDatos(i, 1)
        If vMatrizDatos(i, 1) < m_min Then m_min = vMatrizDatos(i, 1)
    Next i
    
    'Calculamos el diferencial del rango para 6 intervalos
    m_dif = (m_max - m_min) / 6
    j = LBound(m_vDatosColores)
    'Asignamos colores segun el rango, para cada valor de la matriz de datos
    For i = LBound(vMatrizDatos) To UBound(vMatrizDatos)
        Select Case (vMatrizDatos(i, 1))
            Case Is > (m_dif * 5) + m_min:  m_color = COLOR_ROJO
            Case Is > (m_dif * 4) + m_min:  m_color = COLOR_MARRON
            Case Is > (m_dif * 3) + m_min:  m_color = COLOR_AMARILLO
            Case Is > (m_dif * 2) + m_min:  m_color = COLOR_VERDE_CLARO
            Case Is > (m_dif * 1) + m_min:  m_color = COLOR_AÑIL
            Case Else:                      m_color = COLOR_AZUL_OSCURO
        End Select
        m_vDatosColores(j, 0) = vMatrizDatos(i, 0)      'Numero
        m_vDatosColores(j, 1) = vMatrizDatos(i, 1)      'Valor
        m_vDatosColores(j, 2) = m_color                 'Color asignado
        j = j + 1
    Next
    AsignarColores = m_vDatosColores                   'Se devuelve la matriz de colores
End Function

'---------------------------------------------------------------------------------------
' Procedimiento : Get_Bola
' Creación      : 05-dic-2006 20:17
' Autor         : Carlos Almela Baeza
' Objeto        :
'---------------------------------------------------------------------------------------
'
Public Function Get_Bola(Numero As Integer) As bola
    Dim objResult As bola
    Dim strNumero As String
  On Error GoTo Get_Bola_Error
    strNumero = CStr(Numero)
    Set objResult = m_col.Item(strNumero)
    Set Get_Bola = objResult

Get_Bola_CleanExit:
   On Error GoTo 0
    Exit Function

Get_Bola_Error:

    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    '   Audita el error
    Call HandleException(ErrNumber, ErrDescription, "Muestra.Get_Bola")
    '   Lanza el Error
    Err.Raise ErrNumber, ErrSource, ErrDescription
End Function

'---------------------------------------------------------------------------------------
' Procedure : Matriz_Probabilidades
' DateTime  : 14/ago/2007 22:41
' Author    : Carlos Almela Baeza
' Purpose   : Obtiene la matriz de los números con su probabilidad
'---------------------------------------------------------------------------------------
'
Public Property Get Matriz_Probabilidades() As Variant
    Dim m_aProb() As Double             'Matriz de probabilidades
    Dim i As Integer                    'Contador
    Dim bola As bola                    'Bola
   On Error GoTo Matriz_Probabilidades_Error

    'Redimensiona la matriz con los números de la colección
    ReDim m_aProb(m_col.Count - 1, 1)
    
    'Obtiene el dato de probabilidad para cada bola
    i = 0
    For Each bola In m_col
        m_aProb(i, 0) = bola.Numero.Valor
        m_aProb(i, 1) = bola.Probabilidad
        i = i + 1
    Next bola
    
    'Ordena una matriz de dos dimensiones por la 2 dimensión
    Ordenar2 m_aProb
    
    Matriz_Probabilidades = m_aProb

   On Error GoTo 0
   Exit Property

Matriz_Probabilidades_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Matriz_Probabilidades of Módulo de clase Muestra"
End Property

'---------------------------------------------------------------------------------------
' Procedimiento : Total_Numeros
' Creación      : 12-nov-2006 00:23
' Autor         : Carlos Almela Baeza
' Objeto        :
'---------------------------------------------------------------------------------------
'
Public Property Get Total_Numeros() As Integer
    Total_Numeros = mRegTotal
End Property

'---------------------------------------------------------------------------------------
' Procedimiento : Total_Dias
' Creación      : 12-nov-2006 00:29
' Autor         : Carlos Almela Baeza
' Objeto        :
'---------------------------------------------------------------------------------------
'
Public Property Get Total_Dias() As Integer
    Total_Dias = mParametrosMuestra.DiasAnalisis
End Property

'---------------------------------------------------------------------------------------
' Procedimiento : Total_Semanas
' Creación      : 12-nov-2006 00:30
' Autor         : Carlos Almela Baeza
' Objeto        :
'---------------------------------------------------------------------------------------
'
Public Property Get Total_Semanas() As Integer
    Total_Semanas = Fix(mParametrosMuestra.DiasAnalisis / 7)
End Property

'---------------------------------------------------------------------------------------
' Procedimiento : Matriz_Ausencias
' Creación      : 08-dic-2006 20:54
' Autor         : Carlos Almela Baeza
' Objeto        :
'---------------------------------------------------------------------------------------
'
Public Property Get Matriz_Ausencias() As Variant
    Dim m_aProb() As Double
    Dim i As Integer
    Dim bola As bola
  On Error GoTo Matriz_Ausencias_Error

    ReDim m_aProb(m_col.Count, 1)
    i = 0
    For Each bola In m_col
        m_aProb(i, 0) = bola.Numero.Valor
        m_aProb(i, 1) = bola.Ausencias
        i = i + 1
    Next bola
    
    Ordenar2 m_aProb
    
    Matriz_Ausencias = m_aProb

Matriz_Ausencias_CleanExit:
   On Error GoTo 0
    Exit Property

Matriz_Ausencias_Error:
    

    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    '   Audita el error
    Call HandleException(ErrNumber, ErrDescription, "Muestra.Matriz_Ausencias")
    '   Lanza el Error
    Err.Raise ErrNumber, ErrSource, ErrDescription
End Property

' *============================================================================*
' *     Procedure  : Matriz_ProbTiempos
' *     Fichero    : Muestra
' *     Autor      : Carlos Almela Baeza
' *     Creacion   : sáb, 14/01/2012 19:06
' *     Asunto     :
' *============================================================================*
'
Public Property Get Matriz_ProbTiempos() As Variant
        Dim m_aProb() As Double
        Dim i As Integer
        Dim bola As bola
  On Error GoTo Matriz_ProbTiempos_Error

        ReDim m_aProb(m_col.Count - 1, 1)
        i = 0
        For Each bola In m_col
            m_aProb(i, 0) = bola.Numero.Valor
            m_aProb(i, 1) = bola.Prob_TiempoMedio
            i = i + 1
        Next bola
        Ordenar2 m_aProb
        Matriz_ProbTiempos = m_aProb

Matriz_ProbTiempos_CleanExit:
   On Error GoTo 0
    Exit Property

Matriz_ProbTiempos_Error:
    

    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    '   Audita el error
    Call HandleException(ErrNumber, ErrDescription, "Muestra.Matriz_ProbTiempos")
    '   Lanza el Error
    Err.Raise ErrNumber, ErrSource, ErrDescription
End Property

'---------------------------------------------------------------------------------------
' Procedure : Matriz_ProbFrecuencias
' DateTime  : 12/ago/2007 00:26
' Author    : Carlos Almela Baeza
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get Matriz_ProbFrecuencias() As Variant
        Dim m_aProb() As Double
        Dim i As Integer
        Dim bola As bola
  On Error GoTo Matriz_ProbFrecuencias_Error

        ReDim m_aProb(m_col.Count - 1, 1)
        i = 0
        For Each bola In m_col
            m_aProb(i, 0) = bola.Numero.Valor
            m_aProb(i, 1) = bola.Prob_Frecuencia
            i = i + 1
        Next bola
        
        Ordenar2 m_aProb
        
        Matriz_ProbFrecuencias = m_aProb

Matriz_ProbFrecuencias_CleanExit:
   On Error GoTo 0
    Exit Property

Matriz_ProbFrecuencias_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    '   Audita el error
    Call HandleException(ErrNumber, ErrDescription, "Muestra.Matriz_ProbFrecuencias")
    '   Lanza el Error
    Err.Raise ErrNumber, ErrSource, ErrDescription
End Property

'---------------------------------------------------------------------------------------
' Procedimiento : Matriz_Decenas
' Creación      : 08-dic-2006 20:54
' Autor         : Carlos Almela Baeza
' Objeto        : Tipo_prob { 0 = Probabilidad, 1 = tiempo_medio, 2 = Frecuencia}
'---------------------------------------------------------------------------------------
'
Public Property Get Matriz_Decenas(Numero_Decena As Integer, tipo_prob As Integer) As Variant
    Dim m_aProb(10, 1) As Double
    Dim i As Integer
    Dim bola As bola
   On Error GoTo Matriz_Decenas_Error
    
    i = 0
    For Each bola In m_col
        If bola.Numero.Decena = Numero_Decena Then
            m_aProb(i, 0) = bola.Numero.Valor
            If tipo_prob = 1 Then
                m_aProb(i, 1) = bola.Prob_TiempoMedio
            ElseIf tipo_prob = 2 Then
                m_aProb(i, 1) = bola.Prob_Frecuencia
            Else
                m_aProb(i, 1) = bola.Probabilidad
            End If
            i = i + 1
        End If
    Next bola

'    Matriz_Decenas = m_vMatriz_Decenas
    Ordenar2 m_aProb
    Matriz_Decenas = m_aProb

   On Error GoTo 0
   Exit Property

Matriz_Decenas_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & _
        ") in procedure Matriz_Decenas of Módulo de clase Muestra"

End Property

' *============================================================================*
' *     Procedure  : Orden_Decenas
' *     Fichero    : Muestra
' *     Autor      : Carlos Almela Baeza
' *     Creacion   : sáb, 14/01/2012 19:07
' *     Asunto     :
' *============================================================================*
'
Public Property Get Orden_Decenas() As Variant
    Dim m_aProb(4, 1) As Integer
    Dim j As Integer
    Dim i As Integer
    Dim objBola As bola
    
  On Error GoTo Orden_Decenas_Error

    For Each objBola In m_col
        j = objBola.Numero.Decena
        i = j - 1
        m_aProb(i, 0) = j
        m_aProb(i, 1) = m_aProb(i, 1) + objBola.Apariciones

    Next objBola

    Ordenar2 m_aProb, 2
    Orden_Decenas = m_aProb

Orden_Decenas_CleanExit:
   On Error GoTo 0
    Exit Property

Orden_Decenas_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    '   Audita el error
    Call HandleException(ErrNumber, ErrDescription, "Muestra.Orden_Decenas")
    '   Lanza el Error
    Err.Raise ErrNumber, ErrSource, ErrDescription
End Property

'---------------------------------------------------------------------------------------
' Procedimiento : Matriz_ParesImpares
' Creación      : 08-dic-2006 20:54
' Autor         : Carlos Almela Baeza
' Objeto        :
'---------------------------------------------------------------------------------------
'
Public Property Get Matriz_ParesImpares() As Variant
    Dim m_aProb(2, 1) As Variant
    Dim bola As bola
  On Error GoTo Matriz_ParesImpares_Error

    m_aProb(0, 0) = LT_PAR
    m_aProb(1, 0) = LT_IMPAR
    
    For Each bola In m_col
         If bola.Numero.Paridad = LT_PAR Then
                m_aProb(0, 1) = m_aProb(0, 1) + bola.Apariciones
         End If
         If bola.Numero.Paridad = LT_IMPAR Then
                m_aProb(1, 1) = m_aProb(1, 1) + bola.Apariciones
         End If
    Next bola

    Ordenar2 m_aProb
    Matriz_ParesImpares = m_aProb

Matriz_ParesImpares_CleanExit:
   On Error GoTo 0
    Exit Property

Matriz_ParesImpares_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    '   Audita el error
    Call HandleException(ErrNumber, ErrDescription, "Muestra.Matriz_ParesImpares")
    '   Lanza el Error
    Err.Raise ErrNumber, ErrSource, ErrDescription
End Property

'---------------------------------------------------------------------------------------
' Procedimiento : Matriz_AltosBajos
' Creación      : 08-dic-2006 20:54
' Autor         : Carlos Almela Baeza
' Objeto        :
'---------------------------------------------------------------------------------------
'
Public Property Get Matriz_AltosBajos() As Variant
    Dim m_aProb(2, 1) As Variant
    Dim bola As bola
  On Error GoTo Matriz_AltosBajos_Error

    m_aProb(0, 0) = LT_ALTO
    m_aProb(1, 0) = LT_BAJO
    
    For Each bola In m_col
         If bola.Numero.Peso = LT_ALTO Then
                m_aProb(0, 1) = m_aProb(0, 1) + bola.Apariciones
         End If
         If bola.Numero.Peso = LT_BAJO Then
                m_aProb(1, 1) = m_aProb(1, 1) + bola.Apariciones
         End If
    Next bola

    Matriz_AltosBajos = m_aProb

Matriz_AltosBajos_CleanExit:
   On Error GoTo 0
    Exit Property

Matriz_AltosBajos_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    '   Audita el error
    Call HandleException(ErrNumber, ErrDescription, "Muestra.Matriz_AltosBajos")
    '   Lanza el Error
    Err.Raise ErrNumber, ErrSource, ErrDescription
End Property

'---------------------------------------------------------------------------------------
' Procedimiento : Matriz_Terminaciones
' Creación      : 08-dic-2006 20:54
' Autor         : Carlos Almela Baeza
' Objeto        :
'---------------------------------------------------------------------------------------
'
Public Property Get Matriz_Terminaciones(Terminacion As Integer) As Variant
    Dim m_aProb(5, 1) As Double
    Dim i As Integer
    Dim bola As bola
  On Error GoTo Matriz_Terminaciones_Error

    i = 0
    For Each bola In m_col
        If bola.Numero.Terminacion = Terminacion Then
            m_aProb(i, 0) = bola.Numero.Valor
            m_aProb(i, 1) = bola.Probabilidad
            i = i + 1
        End If
    Next bola
    Ordenar2 m_aProb
    Matriz_Terminaciones = m_aProb

Matriz_Terminaciones_CleanExit:
   On Error GoTo 0
    Exit Property

Matriz_Terminaciones_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    '   Audita el error
    Call HandleException(ErrNumber, ErrDescription, "Muestra.Matriz_Terminaciones")
    '   Lanza el Error
    Err.Raise ErrNumber, ErrSource, ErrDescription
End Property

' *============================================================================*
' *     Procedure  : ParametrosMuestra
' *     Fichero    : Muestra
' *     Autor      : Carlos Almela Baeza
' *     Creacion   : sáb, 21/01/2012 22:43
' *     Asunto     :
' *============================================================================*
'
Public Property Get ParametrosMuestra() As ParametrosMuestra

  On Error GoTo ParametrosMuestra_Error

    Set ParametrosMuestra = mParametrosMuestra

ParametrosMuestra_CleanExit:
   On Error GoTo 0
    Exit Property

ParametrosMuestra_Error:
    

    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    '   Audita el error
    Call HandleException(ErrNumber, ErrDescription, "Muestra.ParametrosMuestra")
    '   Lanza el Error
    Err.Raise ErrNumber, ErrSource, ErrDescription
End Property

' *============================================================================*
' *     Procedure  : ParametrosMuestra
' *     Fichero    : Muestra
' *     Autor      : Carlos Almela Baeza
' *     Creacion   : sáb, 21/01/2012 22:43
' *     Asunto     :
' *============================================================================*
'
Public Property Set ParametrosMuestra(objParametrosMuestra As ParametrosMuestra)

  On Error GoTo ParametrosMuestra_Error

    Set mParametrosMuestra = objParametrosMuestra

ParametrosMuestra_CleanExit:
   On Error GoTo 0
    Exit Property

ParametrosMuestra_Error:
    

    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    '   Audita el error
    Call HandleException(ErrNumber, ErrDescription, "Muestra.ParametrosMuestra")
    '   Lanza el Error
    Err.Raise ErrNumber, ErrSource, ErrDescription
End Property



' *============================================================================*
' *     Procedure  : CalcularColores
' *     Fichero    : Muestra
' *     Autor      : Carlos Almela Baeza
' *     Creacion   : mié, 15/02/2012 00:00
' *     Asunto     :
' *============================================================================*
'
Public Sub CalcularColores()
    Dim mMatrizValoresA() As Variant
    Dim mMatrizValoresB() As Variant
    Dim mMatrizValoresC() As Variant
    Dim mColoresA As Variant
    Dim mColoresB As Variant
    Dim mColoresC As Variant
    Dim i As Integer
    Dim mBola As bola
  On Error GoTo CalcularColores_Error
    
    ReDim mMatrizValoresA(m_col.Count - 1, 0 To 1)
    ReDim mMatrizValoresB(m_col.Count - 1, 0 To 1)
    ReDim mMatrizValoresC(m_col.Count - 1, 0 To 1)
    
    i = 0
    For Each mBola In m_col
        mMatrizValoresA(i, 1) = mBola.Probabilidad
        mMatrizValoresB(i, 1) = mBola.Prob_TiempoMedio
        mMatrizValoresC(i, 1) = mBola.Prob_Frecuencia
        mMatrizValoresA(i, 0) = mBola.Numero.Valor
        mMatrizValoresB(i, 0) = mBola.Numero.Valor
        mMatrizValoresC(i, 0) = mBola.Numero.Valor
        
        i = i + 1
    Next mBola
    mColoresA = AsignarColores(mMatrizValoresA)
    mColoresB = AsignarColores(mMatrizValoresB)
    mColoresC = AsignarColores(mMatrizValoresC)
    i = 0
    For Each mBola In m_col
        mBola.Color_Probabilidad = mColoresA(i, 2)
        mBola.Color_Tiempo_Medio = mColoresB(i, 2)
        mBola.Color_Frecuencias = mColoresC(i, 2)
        i = i + 1
    Next mBola

CalcularColores_CleanExit:
   On Error GoTo 0
    Exit Sub

CalcularColores_Error:
    

    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    '   Audita el error
    Call HandleException(ErrNumber, ErrDescription, "Muestra.CalcularColores")
    '   Lanza el Error
    Err.Raise ErrNumber, ErrSource, ErrDescription
End Sub


' *============================================================================*
' *     Procedure  : GetMatrizBolas
' *     Fichero    : Muestra
' *     Autor      : Carlos Almela Baeza
' *     Creacion   : jue, 30/04/2009
' *     Asunto     :
' *============================================================================*
'
Public Function GetMatrizBolas(vOrden As TipoOrdenacion, _
                               Optional datOrden As Boolean = True) As Variant
    Dim bola        As bola                 ' Instancia Bola
    Dim m_aProb()   As Variant              ' Matriz de datos
    Dim i           As Integer              ' índice
  On Error GoTo GetMatrizBolas_Error
    '
    '   Redimensionamos la matriz al Numero de bolas con dos dimensiones
    '   (49, 2)
    '
    ReDim m_aProb(m_col.Count - 1, 1)
    '
    '   Para cada bola en la coleccion
    '
    For Each bola In m_col
        '
        '   Asignamos el valor a la coordenada 1 (0)
        '
        m_aProb(i, 0) = bola.Numero.Valor
        '
        '   Seleccionamos el dato según el tipo
        '
        Select Case vOrden
            Case ordProbabilidad:
                m_aProb(i, 1) = bola.Probabilidad

            Case ordProbTiempoMedio
                m_aProb(i, 1) = bola.Prob_TiempoMedio

            Case ordFrecuencia
                m_aProb(i, 1) = bola.Prob_Frecuencia

            Case ordAusencia
                m_aProb(i, 1) = bola.Ausencias

            Case ordTiempoMedio
                m_aProb(i, 1) = bola.Tiempo_Medio

            Case ordDesviacion
                m_aProb(i, 1) = bola.Desviacion_Tm

            Case ordProximaFecha
                m_aProb(i, 1) = bola.ProximaFecha

            Case ordModa
                m_aProb(i, 1) = bola.Moda

       End Select
       '
       '    Incrementamos I
       '
        i = i + 1
    Next bola
    '
    '   Ordenamos la matriz segun el criterio agregado
    '
    Ordenar2 m_aProb, 2, datOrden
    
    GetMatrizBolas = m_aProb

   On Error GoTo 0
   Exit Function

GetMatrizBolas_Error:
     Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
   ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
   '   Audita el error
   Call HandleException(ErrNumber, ErrDescription, "Muestra.GetMatrizBolas")
   '   Lanza el Error
   Err.Raise ErrNumber, ErrSource, ErrDescription
End Function


Public Property Get Juego() As Juego
    Juego = mParametrosMuestra.Juego
End Property
'------------------------------------------------------------------------------*
' Procedimiento  : CalcularValorHomogeneo
' Fecha          : 07/mar/2020
' Propósito      : Calcular el dato homogeneizado de probabilidades
'------------------------------------------------------------------------------*
Private Sub CalcularValorHomogeneo()
    Dim mProbArray      As Variant
    Dim mTMArray        As Variant
    Dim mFrecArray      As Variant
    Dim i               As Integer
    Dim mBola           As bola
    Dim mMuProb         As Double
    Dim mMuTm           As Double
    Dim mMuFrec         As Double
    Dim mDesvProb       As Double
    Dim mDesvTm         As Double
    Dim mDesvFrec       As Double
    Dim mValorH         As Double
    Dim mTemp           As Double
  On Error GoTo CalcularValorHomogeneo_Error
    '
    '   Redimensionamos variables
    '
    ReDim mProbArray(m_col.Count - 1)
    ReDim mTMArray(m_col.Count - 1)
    ReDim mFrecArray(m_col.Count - 1)
    '
    '   Obtenemos matriz de valores de las bolas
    '
    i = 0
    For Each mBola In m_col
        mProbArray(i) = mBola.Probabilidad
        mTMArray(i) = mBola.Prob_TiempoMedio
        mFrecArray(i) = mBola.Prob_Frecuencia
        i = i + 1
    Next mBola
    '
    '   Calculamos media y desviación tipica
    '
    With Application.WorksheetFunction
        mMuProb = .Average(mProbArray)
        mMuTm = .Average(mTMArray)
        mMuFrec = .Average(mFrecArray)
        mDesvProb = .AveDev(mProbArray)
        mDesvTm = .AveDev(mTMArray)
        mDesvFrec = .AveDev(mFrecArray)
    End With
    '
    '   Calculamos el valor homogeneo
    '
    For i = 1 To m_col.Count
        mTemp = 0
        mTemp = m_col.Item(i).Probabilidad - mMuProb
        mTemp = mTemp / mDesvProb
        mValorH = mTemp
        mTemp = m_col.Item(i).Prob_TiempoMedio - mMuTm
        mTemp = mTemp / mDesvTm
        mValorH = mValorH + mTemp
        mTemp = m_col.Item(i).Prob_Frecuencia - mMuFrec
        mTemp = mTemp / mDesvFrec
        mValorH = mValorH + mTemp
        mValorH = mValorH / 3
        m_col.Item(i).ValorHomogeneo = mValorH
    Next i

  On Error GoTo 0
CalcularValorHomogeneo__CleanExit:
    Exit Sub
            
CalcularValorHomogeneo_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    '   Audita el error
    Call HandleException(ErrNumber, ErrDescription, "Muestra.CalcularValorHomogeneo", ErrSource)
    '   Lanza el Error
    Err.Raise ErrNumber, "FilMuestrae.CalcularValorHomogeneo", ErrDescription
End Sub
