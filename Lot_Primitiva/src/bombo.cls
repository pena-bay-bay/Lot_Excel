VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "bombo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'/**
' *     File        bombo.cls
' *     Created     25/05/2004 21:52
' *     Version     1.1.02
' *     Modified    04/06/2006 11:08
' *     Author      Carlos Almela Baeza
' *     Subject     Clase bombo, simula un bombo de la lotería
' *                 primitiva
' **/
'-----------------------------------------------------------*
'
'//Variables privadas
Private m_iextracciones As Integer          'Numero de extracciones realizadas
Private m_objbola_extraida As bola          'Bola extraida
Private m_itipo_giro As Integer             'tipo de giro, por tiempo o vueltas
Private m_dtiempo_giro As Double            'tiempo de giro
Private m_lNumero_giros As Long             '
Private m_num_bolas As Integer
Private m_bolas As bolas
Private m_modo As Integer
Private Const m_cargado = True
Private Const m_version = "1.1.01"
'------------------------------------------------------*
Private Sub Class_Initialize()
    m_iextracciones = 0
    m_num_bolas = 0
    tipo_giro = 0
    tiempo_giro = 0
    m_lNumero_giros = 1
    Set m_bolas = New bolas
    Debug.Print "Inicio de clase Bombo (" & m_version & ")"
End Sub
Private Sub Class_Terminate()
'    Debug.Print "Finalización de la clase" & "bombo"
    Set m_objbola_extraida = Nothing
End Sub
'------------------------------------------------------*
'---------------------------------------------------------------------------------------
' Procedimiento : cargar
' Creación      : 14-sep-2006 23:37
' Autor         : Carlos Almela Baeza
' Objeto        :
'---------------------------------------------------------------------------------------
'
Public Sub cargar(Optional a_Prob_Datos As Variant)
    Dim i As Integer
    Dim b As bola
    'si me vienen datos trato los datos y asigno a bolas y
    'pongo el modo en bombo_cargado
   On Error GoTo cargar_Error

    If (IsMissing(a_Prob_Datos)) Then
        m_modo = Not m_cargado
    Else
        m_modo = m_cargado
    End If
    
    For i = 0 To m_num_bolas - 1
        Set b = New bola
        If m_modo = m_cargado Then
            b.Numero.Valor = a_Prob_Datos(i, 0)
'            b.Orden = I
'            b.Probabilidad = a_Prob_Datos(I, 1)
        Else
            b.Numero.Valor = i
'            b.Orden = I
        End If
        m_bolas.Add b
    Next i
    m_iextracciones = 0

   On Error GoTo 0
   Exit Sub

cargar_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cargar of Módulo de clase bombo"
End Sub
'------------------------------------------------------*
'---------------------------------------------------------------------------------------
' Procedimiento : girar
' Creación      : 14-sep-2006 23:37
' Autor         : Carlos Almela Baeza
' Objeto        :
'---------------------------------------------------------------------------------------
'
Public Sub girar()
    Dim m_Ti As Double
    Dim m_Ni As Long
   On Error GoTo girar_Error

    If Situacion = lotVACIO Then Exit Sub
    m_Ni = 0
    Select Case Me.tipo_giro
        Case lotGIROS:
            m_Ti = Time
            For i = 1 To Me.Numero_giros
                realizar_giro
                m_Ni = m_Ni + 1
            Next i
            Me.tiempo_giro = Time() - m_Ti
        Case lotTIEMPO:
            m_time = Time() + Me.tiempo_giro
            m_Ni = 0
            While Time() < m_time
                realizar_giro
                m_Ni = m_Ni + 1
            Wend
     End Select
    Me.Numero_giros = m_Ni

   On Error GoTo 0
   Exit Sub

girar_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure girar of Módulo de clase bombo"
End Sub
'/**
' * Function    extraccion
' *
' * @Numero     Devuelve un objeto Numero  entre 1 y
' *             el número de bolas del bombo
' **/
'-----------------------------------------------------------*
Function extraccion() As bola
    'i = aleat()
    i = Rueda()
    If IsObject(bola_extraida) Then Set bola_extraida = Nothing
    
    Set bola_extraida = m_bolas.Item(i)
    m_bolas.Remove i
    
    m_iextracciones = m_iextracciones + 1
    Set extraccion = bola_extraida
End Function
'/**
' * Property    Item
' *
' * @Object     Devuelve una bola según la clave
' **/
Public Property Get Item(vNewValue) As bola
    Set Item = m_bolas.Item(vNewValue)
End Property
' * Property    bolas
' *
' * @int        Numero de bolas que contiene el bombo
' **/
Public Property Get bolas() As Integer
    m_num_bolas = m_bolas.Count
    bolas = m_num_bolas
End Property

Public Property Let bolas(ByVal vNewValue As Integer)
    m_num_bolas = vNewValue
End Property
'/**
' * Function    aleat
' *
' * @int        Devuelve un entero aleatoriamente entre 1 y
' *             el número de bolas del bombo
' **/
'-----------------------------------------------------------*
Private Function aleat() As Integer
    Static b_rand As Boolean
    If Not b_rand Then          'La primera vez que se ejecuta
        b_rand = True           'la función RND se ceba la
        Randomize               'la semilla
    End If
    aleat = Int((Me.bolas * Rnd) + 1)
End Function

Private Sub realizar_giro()
    Dim Ni As Integer, Nf As Integer
    Dim m_bolI, m_bolF As bola
    Dim m_col As Collection
    
   On Error GoTo realizar_giro_Error

    Ni = aleat()
    While (Ni <> Nf) And (Nf = 0)
        Nf = aleat()
    Wend
     
     
    ' intercambio de bolas
    Set m_bolI = m_bolas.Item(Ni)
    Set m_bolF = m_bolas.Item(Nf)
    Set m_col = m_bolas.Coleccion
'    Debug.Print "Total Bolas : " & m_col.Count
'    Debug.Print "Bola inicial: " & Ni
'    Debug.Print "Bola final: " & Nf
'    If Ni = m_col.Count Then
'        Debug.Print "ahora falla"
'    End If
    
    SustituirElem m_col, Ni, m_bolF
    SustituirElem m_col, Nf, m_bolI
    Set m_bolas.Coleccion = m_col

   On Error GoTo 0
   Exit Sub

realizar_giro_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & _
        ") in procedure realizar_giro of Módulo de clase bombo"
End Sub


Public Property Get Situacion() As Integer

    If (m_bolas.Count = 0) Then
        Situacion = lotVACIO
    Else
        Situacion = lotCARGADO
    End If
End Property

'---------------------------------------------------------------------------------------
' Procedimiento : extracciones
' Creación      : 14-sep-2006 23:11
' Autor         : Carlos Almela Baeza
' Objeto        :
'---------------------------------------------------------------------------------------
'
Public Property Get extracciones() As Integer

    extracciones = m_iextracciones

End Property


'---------------------------------------------------------------------------------------
' Procedimiento : bola_extraida
' Creación      : 14-sep-2006 23:11
' Autor         : Carlos Almela Baeza
' Objeto        :
'---------------------------------------------------------------------------------------
'
Public Property Get bola_extraida() As bola

    Set bola_extraida = m_objbola_extraida

End Property

Public Property Set bola_extraida(objbola_extraida As bola)

    Set m_objbola_extraida = objbola_extraida

End Property

'---------------------------------------------------------------------------------------
' Procedimiento : tipo_giro
' Creación      : 14-sep-2006 23:11
' Autor         : Carlos Almela Baeza
' Objeto        :
'---------------------------------------------------------------------------------------
'
Public Property Get tipo_giro() As Integer

    tipo_giro = m_itipo_giro

End Property

Public Property Let tipo_giro(ByVal itipo_giro As Integer)

    m_itipo_giro = itipo_giro

End Property

'---------------------------------------------------------------------------------------
' Procedimiento : tiempo_giro
' Creación      : 14-sep-2006 23:11
' Autor         : Carlos Almela Baeza
' Objeto        :
'---------------------------------------------------------------------------------------
'
Public Property Get tiempo_giro() As Double

    tiempo_giro = m_dtiempo_giro

End Property

Public Property Let tiempo_giro(ByVal dtiempo_giro As Double)

    m_dtiempo_giro = dtiempo_giro

End Property

'---------------------------------------------------------------------------------------
' Procedimiento : Numero_giros
' Creación      : 14-sep-2006 23:10
' Autor         : Carlos Almela Baeza
' Objeto        :
'---------------------------------------------------------------------------------------
'
Public Property Get Numero_giros() As Long

    Numero_giros = m_lNumero_giros

End Property

Public Property Let Numero_giros(ByVal lNumero_giros As Long)

    m_lNumero_giros = lNumero_giros

End Property
Public Property Get Coleccion() As Collection
    Set Coleccion = m_bolas.Coleccion
End Property


Private Function Rueda() As Integer
    Dim m_col As New Collection
    Dim m_bol As New bola
    Dim i As Integer
    Dim m_dProb() As Double
    Dim m_tot As Double

    Set m_col = m_bolas.Coleccion
    ReDim m_dProb(m_col.Count - 1)
    Rueda = 0
    i = 0
    For Each m_bol In m_col
        m_dProb(i) = m_bol.Probabilidad
        i = i + 1
        m_tot = m_tot + m_bol.Probabilidad
    Next m_bol

    ' Reducir a 100
    For i = 1 To UBound(m_dProb)
        m_dProb(i) = (m_dProb(i) / m_tot) * 100
    Next i
    Randomize
    j = Int(100 * Rnd)
    Suma = 0
    For i = 1 To (UBound(m_dProb) - 1)
        Suma = Suma + m_dProb(i)
        If (Suma > j) Then
            Rueda = i
            Exit For
        End If
    Next i
    ' Si no se selecciona ningún número es el último
    If Rueda = 0 Then
        Rueda = m_col.Count
    End If
End Function


