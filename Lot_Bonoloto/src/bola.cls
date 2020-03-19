VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "bola"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'---------------------------------------------------------------------------------------
' Modulo    : Bola
' Creado    : 06-may-2004 23:05
' Autor     : Carlos Almela Baeza
' Version   : 1.0.2 09-dic-2006 20:49
' Objeto    : Esta clase contiene las propiedades probabilisticas
'             de una bola que simula un Numero en el bombo
'---------------------------------------------------------------------------------------
Option Explicit
Private m_Numero                As Numero
Private m_apariciones           As Integer
Private m_total                 As Integer
Private m_ausencias             As Integer
Private m_iModa                 As Integer
Private m_iMediana              As Integer
Private m_dTiempo_Medio         As Double
Private m_iMaximo_Tm            As Integer
Private m_iMinimo_Tm            As Integer
Private m_dDesviacion_Tm        As Double
Private m_dProb_TiempoMedio     As Double
Private m_dtUltima_Fecha        As Date
Private m_frecuencias           As Variant
Private m_dtFechaAnalisis       As Date
Private m_dtProximaFecha        As Date
Private m_iUltimoRegistro       As Integer
Private m_vMatriz_Apariciones   As Variant
Private m_vMatriz_Fechas        As Variant
Private m_iColor_Probabilidad   As Integer
Private m_iColor_Tiempo_Medio   As Integer
Private m_iColor_Frecuencias    As Integer
Private mValorHomogeneo         As Double
'---------------------------------------------------------------------------------------
Public colRegistros             As Collection   ' Colección de Numeros de registros
Public colFechasAparicion       As Collection   ' Colección de fechas de aparicion
Public colTiemposMedios         As Collection   ' Colección de Numero de registros entre
                                                ' apariciones


'---------------------------------------------------------------------------------------
' Procedimiento : Class_Initialize
' Creación      : 12-nov-2006 00:15
' Autor         : Carlos Almela Baeza
' Objeto        :
'---------------------------------------------------------------------------------------
'
Private Sub Class_Initialize()
    Set m_Numero = New Numero
    Set colRegistros = New Collection
    Set colFechasAparicion = New Collection
    Set colTiemposMedios = New Collection
'   Rango_Frecuencias = Array(1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 20, 50, 100)
    Rango_Frecuencias = Array(3, 6, 9, 15, 25, 50)
    m_apariciones = -1
    m_dProb_TiempoMedio = -1
    mValorHomogeneo = 0
End Sub

' *============================================================================*
' *     Procedure  : Class_Terminate
' *     Fichero    : bola
' *     Autor      : Carlos Almela Baeza
' *     Creacion   : dom, 12/02/2012 18:08
' *     Asunto     :
' *============================================================================*
'
Private Sub Class_Terminate()
    Set m_Numero = Nothing
    Set colRegistros = Nothing
    Set colFechasAparicion = Nothing
    Set colTiemposMedios = New Collection
End Sub

'---------------------------------------------------------------------------------------
' Procedimiento : Constructor
' Creación      : 12-nov-2006 00:14
' Autor         : Carlos Almela Baeza
' Objeto        :
'---------------------------------------------------------------------------------------
'
Public Sub Constructor(vNewValor As Integer)
    m_Numero.Valor = vNewValor
End Sub

' *============================================================================*
' *     Procedure  : Actualizar
' *     Fichero    : bola
' *     Autor      : Carlos Almela Baeza
' *     Creacion   : dom, 05/02/2012 19:47
' *     Asunto     :
' *============================================================================*
'
Public Sub Actualizar(datRegAnalisis As Integer)
    Dim mStrKey         As String
    Dim i               As Integer
    Dim j               As Integer
    Dim a               As Variant
    Dim Dias            As Integer
    Dim vValores        As Variant
    Dim mu              As Double
    Dim Sigma           As Double
    Dim x               As Integer

    
  On Error GoTo Actualizar_Error

    j = 0: i = 0
    For Each a In colRegistros
        i = a
        If j = 0 Then
            j = i
        Else
            Dias = i - j
            mStrKey = CStr(colTiemposMedios.Count)
            If colTiemposMedios.Count <> 0 Then
                colTiemposMedios.Add Dias, mStrKey, , colTiemposMedios.Count
            Else
                colTiemposMedios.Add Dias, mStrKey
            End If
            j = i
        End If
    Next a
    '
    '   Ausencias
    '
    m_iUltimoRegistro = i
    m_ausencias = datRegAnalisis - m_iUltimoRegistro
    '
    '   Fecha última aparición
    '
    m_apariciones = colFechasAparicion.Count
    Select Case m_apariciones
    Case 0:
        '
        '   Buscar la última fecha de aparición
        '
        '   m_dtUltima_Fecha =
        ReDim vValores(1) As Integer
        vValores(0) = 0
    Case 1:
        m_dtUltima_Fecha = colFechasAparicion.Item(m_apariciones)
        ReDim vValores(1) As Integer
        vValores(0) = 0
    Case Else
        m_dtUltima_Fecha = colFechasAparicion.Item(m_apariciones)
        '
        '   Traspasar a matriz valores
        '
        ReDim vValores(colTiemposMedios.Count - 1) As Integer
        i = 0
        For Each a In colTiemposMedios
            vValores(i) = a
            i = i + 1
        Next a
    End Select
    '
    '   contar frecuencias
    '
    
    With Application.WorksheetFunction
        m_dTiempo_Medio = .Average(vValores)
        m_iMaximo_Tm = .Max(vValores)
        m_iMinimo_Tm = .Min(vValores)
        m_dDesviacion_Tm = .AveDev(vValores)
        m_iModa = GetModa(vValores)
        m_iMediana = .Median(vValores)
        m_frecuencias = .Frequency(vValores, Rango_Frecuencias)
    End With
    '
    '   Probabilidad de tiempo medio
    '
    If m_dProb_TiempoMedio = -1 Then
            mu = m_dTiempo_Medio
            Sigma = m_dDesviacion_Tm
            x = m_ausencias
            If Sigma > 0 And mu > 0 Then
                m_dProb_TiempoMedio = (1 / (Sigma * Sqr(2 * PI))) * (Exp(-0.5 * ((x - mu / Sigma) ^ 2)))
            Else
                m_dProb_TiempoMedio = 0
            End If
    End If


Actualizar_CleanExit:
   On Error GoTo 0
    Exit Sub

Actualizar_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    '   Audita el error
    Call HandleException(ErrNumber, ErrDescription, "bola.Actualizar")
    '   Lanza el Error
    Err.Raise ErrNumber, ErrSource, ErrDescription
  
End Sub

'---------------------------------------------------------------------------------------
' Procedimiento : Apariciones
' Creación      : 12-nov-2006 00:14
' Autor         : Carlos Almela Baeza
' Objeto        :
'---------------------------------------------------------------------------------------
'
Public Property Get Apariciones() As Integer
    If m_apariciones = -1 Then
        m_apariciones = colFechasAparicion.Count
    End If
    Apariciones = m_apariciones
End Property

Public Property Get Frecuencias() As Variant
    Frecuencias = m_frecuencias
End Property

' *============================================================================*
' *     Procedure  : Prob_Frecuencia
' *     Fichero    : bola
' *     Autor      : Carlos Almela Baeza
' *     Creacion   : jue, 09/02/2012 23:43
' *     Asunto     :
' *============================================================================*
'
Public Property Get Prob_Frecuencia() As Double
    Dim m_frec As Integer
'    Dim m_aus As Integer
    Dim i As Integer
    '
    '   Si no ha salido la probabilidad es 0 o no hay frecuencias
    '
    If m_apariciones = 0 _
    Or IsEmpty(m_frecuencias) Then
        Prob_Frecuencia = 0
        Exit Property
    End If
    '
    '   Localizamos el número de apariciones para un intervalo
    '   de ausencias
    '
    m_frec = -1
    For i = 0 To UBound(Rango_Frecuencias) - 1
        If m_ausencias >= Rango_Frecuencias(i) _
        And m_ausencias < Rango_Frecuencias(i + 1) Then
            m_frec = m_frecuencias(i + 1, 1)
            Exit For
        End If
    Next i
    '
    '   Si no se encuentran las ausencias se toma como frecuencia el último valor
    '
    If m_frec = -1 And _
    Not IsEmpty(m_frecuencias) Then
        i = UBound(Rango_Frecuencias)
        m_frec = m_frecuencias(i + 1, 1)
    End If
    '
    '   La probabilidad de la frecuencia es el número de apariciones para esas ausencias
    '   entre el total de apariciones
    '
    Prob_Frecuencia = m_frec / m_apariciones
End Property

'---------------------------------------------------------------------------------------
' Procedimiento : Total
' Creación      : 12-nov-2006 00:14
' Autor         : Carlos Almela Baeza
' Objeto        :
'---------------------------------------------------------------------------------------
'
Public Property Get Total() As Integer
    Total = m_total
End Property

Public Property Let Total(ByVal vNewValue As Integer)
    m_total = vNewValue
End Property

'---------------------------------------------------------------------------------------
' Procedimiento : probabilidad
' Creación      : 12-nov-2006 00:13
' Autor         : Carlos Almela Baeza
' Objeto        :
'---------------------------------------------------------------------------------------
'
Public Property Get Probabilidad() As Double
    If (m_total > 0) Then
        Probabilidad = m_apariciones / m_total
    Else
        Probabilidad = 0
    End If
End Property

'---------------------------------------------------------------------------------------
' Procedimiento : Numero
' Creación      : 12-nov-2006 00:13
' Autor         : Carlos Almela Baeza
' Objeto        :
'---------------------------------------------------------------------------------------
'
Public Property Get Numero() As Numero
    Set Numero = m_Numero
End Property

Public Property Let Numero(ByVal vNewValue As Numero)
    m_Numero = vNewValue
End Property

'---------------------------------------------------------------------------------------
' Procedimiento : Ausencias
' Creación      : 12-nov-2006 00:13
' Autor         : Carlos Almela Baeza
' Objeto        :
'---------------------------------------------------------------------------------------
'
Public Property Get Ausencias() As Integer
    Ausencias = m_ausencias
End Property

Public Property Get Clase_Ausencias() As String
    Select Case m_ausencias
        Case 1, 2, 3: Clase_Ausencias = "A"
        Case 4 To 9: Clase_Ausencias = "B"
        Case Else: Clase_Ausencias = "C"
    End Select
End Property

'---------------------------------------------------------------------------------------
' Procedimiento : Color Probabilidad
' Creación      : 12-nov-2006 23:38
' Autor         : Carlos Almela Baeza
' Objeto        :
'---------------------------------------------------------------------------------------
'
Public Property Get Color_Probabilidad() As Integer
    Color_Probabilidad = m_iColor_Probabilidad
End Property

Public Property Let Color_Probabilidad(ByVal vNewValor As Integer)
    m_iColor_Probabilidad = vNewValor
End Property

'---------------------------------------------------------------------------------------
' Procedimiento : Color Frecuencias
' Creación      : 12-nov-2006 23:38
' Autor         : Carlos Almela Baeza
' Objeto        :
'---------------------------------------------------------------------------------------
'
Public Property Get Color_Frecuencias() As Integer
    Color_Frecuencias = m_iColor_Frecuencias
End Property

Public Property Let Color_Frecuencias(ByVal vNewValor As Integer)
    m_iColor_Frecuencias = vNewValor
End Property

'---------------------------------------------------------------------------------------
' Procedimiento : Color_Tiempo_Medio
' Creación      : 08-dic-2006 20:51
' Autor         : Carlos Almela Baeza
' Objeto        :
'---------------------------------------------------------------------------------------
'
Public Property Get Color_Tiempo_Medio() As Integer
    Color_Tiempo_Medio = m_iColor_Tiempo_Medio
End Property

Public Property Let Color_Tiempo_Medio(ByVal iColor_Tiempo_Medio As Integer)
    m_iColor_Tiempo_Medio = iColor_Tiempo_Medio
End Property

'---------------------------------------------------------------------------------------
' Procedimiento : Matriz_Apariciones
' Creación      : 08-dic-2006 20:52
' Autor         : Carlos Almela Baeza
' Objeto        :
'---------------------------------------------------------------------------------------
'
Public Property Get Matriz_Apariciones() As Variant
    Dim a               As Variant
    Dim i               As Integer
    If IsEmpty(m_vMatriz_Apariciones) Then
        ReDim m_vMatriz_Apariciones(colTiemposMedios.Count - 1)
        i = 0
        For Each a In Me.colTiemposMedios
              m_vMatriz_Apariciones(i) = a
              i = i + 1
        Next a
    End If
    Matriz_Apariciones = m_vMatriz_Apariciones
End Property

Public Property Get Matriz_Fechas() As Variant
    Dim a               As Variant
    Dim i               As Integer
    If IsEmpty(m_vMatriz_Fechas) Then
        ReDim m_vMatriz_Fechas(colFechasAparicion.Count - 1)
        i = 0
        For Each a In Me.colFechasAparicion
              m_vMatriz_Fechas(i) = a
              i = i + 1
        Next a
    End If

    Matriz_Fechas = m_vMatriz_Fechas
End Property

'---------------------------------------------------------------------------------------
' Procedimiento : Tiempo_Medio
' Creación      : 08-dic-2006 20:52
' Autor         : Carlos Almela Baeza
' Objeto        :
'---------------------------------------------------------------------------------------
'
Public Property Get Tiempo_Medio() As Double
    Tiempo_Medio = m_dTiempo_Medio
End Property

'---------------------------------------------------------------------------------------
' Procedimiento : Maximo_Tm
' Creación      : 08-dic-2006 20:52
' Autor         : Carlos Almela Baeza
' Objeto        :
'---------------------------------------------------------------------------------------
'
Public Property Get Maximo_Tm() As Integer
    Maximo_Tm = m_iMaximo_Tm
End Property

'---------------------------------------------------------------------------------------
' Procedimiento : Minimo_Tm
' Creación      : 08-dic-2006 20:52
' Autor         : Carlos Almela Baeza
' Objeto        :
'---------------------------------------------------------------------------------------
'
Public Property Get Minimo_Tm() As Integer
    Minimo_Tm = m_iMinimo_Tm
End Property

Public Property Get Mediana() As Integer
    Mediana = m_iMediana
End Property
'---------------------------------------------------------------------------------------
' Procedimiento : Desviacion_Tm
' Creación      : 08-dic-2006 20:52
' Autor         : Carlos Almela Baeza
' Objeto        :
'---------------------------------------------------------------------------------------
'
Public Property Get Desviacion_Tm() As Double
    Desviacion_Tm = m_dDesviacion_Tm
End Property

'---------------------------------------------------------------------------------------
' Procedimiento : Prob_TiempoMedio
' Creación      : 08-dic-2006 20:52
' Autor         : Carlos Almela Baeza
' Objeto        :
'---------------------------------------------------------------------------------------
'
Public Property Get Prob_TiempoMedio() As Double
    Prob_TiempoMedio = m_dProb_TiempoMedio
End Property

'---------------------------------------------------------------------------------------
' Procedimiento : Ultima_Fecha
' Creación      : 08-dic-2006 20:52
' Autor         : Carlos Almela Baeza
' Objeto        :
'---------------------------------------------------------------------------------------
'
Public Property Get Ultima_Fecha() As Date
    Ultima_Fecha = m_dtUltima_Fecha
End Property


' *============================================================================*
' *     Procedure  : Moda
' *     Fichero    : bola
' *     Autor      : Carlos Almela Baeza
' *     Creacion   : jue, 09/02/2012 23:44
' *     Asunto     :
' *============================================================================*
'
Public Property Get Moda() As Integer
    Moda = m_iModa
End Property
'------------------------------------------------------------------------------*
' Procedimiento  : ValorHomogeneo
' Fecha          : dd/MMM/yyyy
' Propósito      : Incluir el valor homogeneo de la colección para la bola
' Parámetros     : Almacena el cálculo de la homogeneización de Probabilidad
'                  + TiempoMedio + Frecuencia
'------------------------------------------------------------------------------*
Public Property Get ValorHomogeneo() As Double
    ValorHomogeneo = mValorHomogeneo
End Property

Public Property Let ValorHomogeneo(vNewData As Double)
    mValorHomogeneo = vNewData
End Property
' *============================================================================*
' *     Procedure  : ProximaFecha
' *     Fichero    : bola
' *     Autor      : Carlos Almela Baeza
' *     Creacion   : jue, 09/02/2012 23:44
' *     Asunto     :
' *============================================================================*
'
Public Property Get ProximaFecha() As Date
    Dim mx_Frec     As Integer
    Dim i           As Integer
    Dim mx_dias     As Integer
    Dim mn_I        As Integer
    Dim mx_I        As Integer
    Dim mInfo       As InfoSorteo
    Dim mDias       As Integer

  On Error GoTo ProximaFecha_Error
    '
    '   Si no hemos calculado la fecha
    '
    If m_dtProximaFecha = 0 Then
        '
        '   Inicializamos el número de sorteos a calcular
        '
        mDias = 0
        ' Evaluamos si ausencias es mayor o menor que Tiempo medio
        ' Si es menor, próxima aparicion es igual a la ultima aparición mas tiempo
        ' medio
        If m_ausencias < m_dTiempo_Medio Then
            mDias = CInt(m_dTiempo_Medio)
        Else
            ' Si la ausencia es mayor que el Tiempo Medio
            ' Localizamos el rango de frecuencias con mas probabilidad
            '
            ' Primero localiza la ausencia en el rango de frecuencias
            For i = 0 To UBound(Rango_Frecuencias) - 1
                If Rango_Frecuencias(i) >= m_ausencias Then
                    mn_I = i
                    Exit For
                End If
            Next i
            '   Despues busca el rango de mayor apariciones
            '   al finalizar suma ese rango al último registro
            '   y se va a localizar una fecha a DB
            mx_I = UBound(Rango_Frecuencias) + 1
            '
            ' Si hay frecuenciasRango_Frecuencias
            If Not (IsEmpty(m_frecuencias)) Then
                '
                '   y la ausencia es inferior al max de las frecuencias
                '
                If (mn_I < mx_I) Then
                    '
                    '   Localiza desde las ausencias hasta el máximo de frecuencias
                    '   la max frecuencia
                    For i = mn_I To mx_I
                        If (m_frecuencias(i + 1, 1) > mx_Frec) Then
                            mx_Frec = m_frecuencias(i + 1, 1)
                            mDias = Rango_Frecuencias(i)
                        End If
                    Next i
                    If mDias = 0 Then
                        If m_iMaximo_Tm >= m_ausencias Then
                            mDias = m_iMaximo_Tm
                         Else
                            mDias = m_ausencias
                        End If
                    End If
                End If
            Else
                '
                '   Si no hay frecuencias asumimos el máximo del numero si es mayor que las ausencias
                '
                If m_iMaximo_Tm > m_ausencias Then
                    mDias = m_iMaximo_Tm
                Else
                    mDias = m_ausencias
                End If
            End If
        End If
    End If
    '
    '   Sumamos los dias a la última fecha
    '
    If mDias = 0 Then
        m_dtProximaFecha = MX_FECHA
    Else
        Set mInfo = New InfoSorteo
        m_dtProximaFecha = mInfo.AddDiasSorteo(m_dtUltima_Fecha, mDias)
    End If
    
    ProximaFecha = m_dtProximaFecha

ProximaFecha_CleanExit:
   On Error GoTo 0
    Exit Property

ProximaFecha_Error:

    Call HandleException(Err.Number, Err.Description, "bola.ProximaFecha")

End Property


Public Property Let Fecha_Analisis(ByVal dtFecha_Analisis As Date)
    m_dtFechaAnalisis = dtFecha_Analisis
End Property

Public Property Get Tendencia() As String
    Dim mx_corto As Integer, mx_medio As Integer, mx_largo As Integer
    Dim mT_corto As Integer, mT_medio As Integer, mT_largo As Integer
    Dim i As Integer, mx_I
    Dim K1 As Integer, K2 As Integer
    
    mx_corto = 0: mx_medio = 0: mx_largo = 0
    mT_corto = 0: mT_medio = 0: mT_largo = 0
        
    K1 = 4: K2 = 11
    
    If (IsEmpty(Frecuencias)) Then
        Tendencia = "No definida"
        Exit Property
    End If
    mx_I = UBound(m_frecuencias)
    For i = 2 To mx_I
        If (i <= K1) Then
             If (m_frecuencias(i, 1) > mx_corto) Then
                mx_corto = m_frecuencias(i, 1)
             End If
             mT_corto = mT_corto + m_frecuencias(i, 1)
             
        ElseIf (i > K1) And (i <= K2) Then
             If (m_frecuencias(i, 1) > mx_medio) Then
                mx_medio = m_frecuencias(i, 1)
             End If
             mT_medio = mT_medio + m_frecuencias(i, 1)
        Else
             If (m_frecuencias(i, 1) > mx_largo) Then
                mx_largo = m_frecuencias(i, 1)
             End If
             mT_largo = mT_largo + m_frecuencias(i, 1)
        End If
    Next i
    
    
    If (mx_corto >= mx_medio) And (mx_corto >= mx_largo) Then
        Tendencia = "Corto"
    ElseIf (mx_medio >= mx_corto) And (mx_medio >= mx_largo) Then
        Tendencia = "Medio"
    Else
        Tendencia = "Largo"
    End If
End Property


Public Property Get UltimoRegistro() As Integer
    UltimoRegistro = m_iUltimoRegistro
End Property
