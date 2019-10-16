VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Periodo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' *============================================================================*
' *
' *     Fichero    : Periodo.cls
' *
' *     Autor      : Carlos Almela Baeza
' *     Creación   : 22/04/2007 21:31
' *     Versión    : 1.0
' *     Propósito  : Clase que representa un intervalo de tiempo entre dos fechas
' *
' *============================================================================*
Option Explicit
Option Base 0

'
'--- Variables Privadas -------------------------------------------------------*
Private m_dtFecha_Inicial           As Date
Private m_dtFecha_Final             As Date
Private m_vFechas_Personalizadas    As Variant
Private m_VId_FechasPersonalizadas  As Variant
Private m_iTipo_Fecha               As Integer
Private mvarTexto                   As String
Private m_sTipoFecha                As String

'--- Constantes ---------------------------------------------------------------*
Private Const MinDate = #1/1/1900#
Private Const MaxDate = #12/31/2999#
Private Const lt0 = "Personalizadas"
Private Const lt1 = "Semana Anterior"
Private Const lt2 = "Quincena Anterior"
Private Const lt3 = "Mes Anterior"
Private Const lt4 = "Año Anterior"
Private Const lt5 = "Semana Actual"
Private Const lt6 = "Quincena Actual"
Private Const lt7 = "Mes Actual"
Private Const lt8 = "Año Actual"
Private Const lt9 = "Lo que va de semana"
Private Const lt10 = "Lo que va de mes"
Private Const lt11 = "Lo que va de año"
Private Const lt12 = "Lo que va de Trimestre"
Private Const lt13 = "Ultima Semana"
Private Const lt14 = "Ultima Quincena"
Private Const lt15 = "Ultimo Mes"
Private Const lt16 = "Ultimo Trimestre"
Private Const lt17 = "Ultimo Año"
Private Const lt18 = "Hasta Hoy"
Private Const lt19 = "Hoy"
Private Const lt20 = "Ayer"
Private Const lt21 = "Mañana"
Private Const lt_1 = "Sin Definir"
'--- Mensajes -----------------------------------------------------------------*
'--- Errores ------------------------------------------------------------------*
'--- Metodos Privados ---------------------------------------------------------*
'------------------------------------------------------------------------------*
' Procedimiento  : Class_Initialize
' Fecha          : 22/04/2007 21:36
' Propósito      : Inicialización de las propiedades internas de la clase
'------------------------------------------------------------------------------*
Private Sub Class_Initialize()
    m_vFechas_Personalizadas = Array(lt0, lt1, lt2, lt3, lt4, lt5, lt6, _
                                     lt7, lt8, lt9, lt10, lt11, lt12, lt13, _
                                     lt14, lt15, lt16, lt17, lt18, lt19, lt20, lt21)
    m_VId_FechasPersonalizadas = Array(0, 1, 2, 3, 4, 5, 6, 7, 8, 9, _
                                10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21)
    Tipo_Fecha = ctSinDefinir
End Sub
'------------------------------------------------------------------------------*
' Procedimiento  : fnUltimoDia
' Fecha          : 22/04/2007 21:36
' Propósito      : Devuelve el ultimo dia del mes de la fecha
' Parámetros     : Fecha a procesar
' Retorno        : (Int) último día del mes de la fecha
'------------------------------------------------------------------------------*
Private Function fnUltimoDia(tmpFecha As Date) As Integer
    Dim tmpMes As Integer
    Dim tmpBisiesto As Boolean
    Dim i As Integer
    
    i = Year(tmpFecha) Mod 4
    If i = 0 Then
        tmpBisiesto = True
    Else
        tmpBisiesto = False
    End If
    tmpMes = Month(tmpFecha)
    Select Case tmpMes
        Case 1, 3, 5, 7, 8, 10, 12: fnUltimoDia = 31
        Case 2
            If tmpBisiesto Then
                fnUltimoDia = 29
            Else
                fnUltimoDia = 28
            End If
                
        Case 4, 6, 9, 11: fnUltimoDia = 30
    End Select
End Function

'
'----| fnRestaMes (  fecha ) : fecha  Restar un mes a la fecha
'
'------------------------------------------------------------------------------*
' Procedimiento  : fnRestaMes
' Fecha          : 22/04/2007 21:36
' Propósito      : Resta un mes a la fecha
' Parámetros     : Fecha a procesar
' Retorno        : (Date) Fecha correspondiente a un mes menos
'------------------------------------------------------------------------------*
Private Function fnRestaMes(vNewValue As Date) As Date
    Dim mMes As Integer, mDia As Integer
    Dim mFecha As Date
    If Month(vNewValue) = 1 Then
        fnRestaMes = CDate(Day(vNewValue) & "/" & 12 & "/" _
                 & Year(vNewValue) - 1)
    Else
        mMes = Month(vNewValue) - 1
        mFecha = CDate(1 & "/" & mMes & "/" & Year(vNewValue))
        mDia = fnUltimoDia(mFecha)
        fnRestaMes = CDate(mDia & "/" & mMes & "/" & Year(vNewValue))
    End If
End Function

'--- Metodos Publicos ---------------------------------------------------------*
'------------------------------------------------------------------------------*
' Procedimiento  : Init
' Fecha          : 22/04/2007 21:36
' Propósito      : Inicializa el periodo con dos fechas
' Parámetros     : vNewFechaInicial fecha inicial
'                  vNewFechaFinal fecha final
'------------------------------------------------------------------------------*
Public Sub Init(vNewFechaInicial As Date, vNewFechaFinal As Date)
    m_dtFecha_Inicial = vNewFechaInicial
    m_dtFecha_Final = vNewFechaFinal
    ' Comprobar si las fechas encajan en los periodos definidos
    
End Sub
'------------------------------------------------------------------------------*
' Procedimiento  : Contiene
' Fecha          : 22/04/2007 21:36
' Propósito      : Devuelve true si la fecha indicada está contenida en el
'                  periodo
' Parámetros     : vNewValue Fecha de comprobación
'------------------------------------------------------------------------------*
Public Function Contiene(vNewValue As Date) As Boolean
    If (m_dtFecha_Inicial <= vNewValue) And _
       (m_dtFecha_Final >= vNewValue) Then
        Contiene = True
    Else
        Contiene = False
    End If
End Function

'------------------------------------------------------------------------------*
' Procedimiento  : CargaTabla
' Fecha          : 22/04/2007 21:36
' Propósito      : Carga una list box con los periodos establecidos
' Parámetros     : cmbLista Objeto de tipo Control ComboBox
'------------------------------------------------------------------------------*
Public Sub CargaTabla(cmbLista As Object)
    CargaCombo cmbLista, m_VId_FechasPersonalizadas
End Sub

'---------------------------------------------------------------------------------------
' Procedure : CargaCombo
' Author    : Charly
' Date      : 10/11/2013
' Purpose   : Carga una lista de selección con unos parámetros determinados
'---------------------------------------------------------------------------------------
'
Public Sub CargaCombo(cboLista As Object, mElementos As Variant)
    Dim i As Integer
    Dim literal As String
  On Error GoTo CargaCombo_Error
    '
    '  Si el tipo de objeto no es un control salir
    '
    If Not (TypeOf cboLista Is Control) Then
        Return
    End If
    '
    ' Inicializamos el Combo
    '
    cboLista.Clear
    '
    ' Para cada elemento en la matriz
    '
    For i = 0 To UBound(mElementos)
        '
        ' Obtenemos el literal
        '
        literal = m_vFechas_Personalizadas(mElementos(i))
        '
        ' Se añade a la colección
        '
        cboLista.AddItem literal
        
    Next i
    'Inicializar el control
    cboLista.ListIndex = -1

   On Error GoTo 0
   Exit Sub

CargaCombo_Error:
     Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
   ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
   '   Audita el error
   Call HandleException(ErrNumber, ErrDescription, "Periodo.CargaCombo")
   '   Lanza el Error
   Err.Raise ErrNumber, ErrSource, ErrDescription
End Sub

'--- Propiedades --------------------------------------------------------------*
'------------------------------------------------------------------------------*
' Procedimiento  : FechaInicial
' Fecha          : 22/04/2007 21:36
' Propósito      : Fecha inicial del periodo
'------------------------------------------------------------------------------*
Public Property Get FechaInicial() As Date
    FechaInicial = m_dtFecha_Inicial
End Property

Public Property Let FechaInicial(ByVal dtFecha_Inicial As Date)
    m_dtFecha_Inicial = dtFecha_Inicial
End Property

'------------------------------------------------------------------------------*
' Procedimiento  : FechaFinal
' Fecha          : 22/04/2007 21:36
' Propósito      : Fecha final del periodo
'------------------------------------------------------------------------------*
Public Property Get FechaFinal() As Date
    FechaFinal = m_dtFecha_Final
End Property

Public Property Let FechaFinal(ByVal dtFecha_Final As Date)
    m_dtFecha_Final = dtFecha_Final
End Property

'------------------------------------------------------------------------------*
' Procedimiento  : Dias
' Fecha          : 22/04/2007 21:36
' Propósito      : Número de días comprendido entre la fecha inicial
'                  y la fecha final
'------------------------------------------------------------------------------*
Public Property Get Dias() As Integer
    If m_dtFecha_Inicial = 0 Then
        Dias = 1
    Else
        Dias = m_dtFecha_Final - m_dtFecha_Inicial + 1
    End If
End Property

'------------------------------------------------------------------------------*
' Procedimiento  : Tipo_Fecha
' Fecha          : 22/04/2007 21:49
' Propósito      : Periodo de fechas tipificados del lenguaje natural,
'                  por ejemplo: hoy, la última semana, etc.
'------------------------------------------------------------------------------*
Public Property Get Tipo_Fecha() As Integer
   Tipo_Fecha = m_iTipo_Fecha
End Property

Public Property Let Tipo_Fecha(ByVal iTipo_Fecha As Integer)
   On Error GoTo Tipo_Fecha_Error
    m_iTipo_Fecha = iTipo_Fecha
      
    Select Case m_iTipo_Fecha
        Case ctPersonalizadas
            mvarTexto = lt0
            
        Case ctSemanaPasada
            mvarTexto = lt1
            m_dtFecha_Final = Date - Weekday(Date, vbMonday)
            m_dtFecha_Inicial = m_dtFecha_Final - 6
        
        Case ctQuincenaPasada
            mvarTexto = lt2
            m_dtFecha_Final = Date - Weekday(Date, vbMonday)
            m_dtFecha_Inicial = m_dtFecha_Final - 13
        
        Case ctMesAnterior
            mvarTexto = lt3
            m_dtFecha_Final = fnRestaMes(Date)
            m_dtFecha_Inicial = CDate("1/" & Month(m_dtFecha_Final) & "/" _
                              & Year(m_dtFecha_Final))
            m_dtFecha_Final = CDate(fnUltimoDia(m_dtFecha_Final) & "/" _
                              & Month(m_dtFecha_Final) & "/" _
                              & Year(m_dtFecha_Final))
        
        Case ctAñoAnterior
            mvarTexto = lt4
            m_dtFecha_Final = CDate("31/12/" & Year(Date) - 1)
            m_dtFecha_Inicial = CDate("1/1/" & Year(Date) - 1)
        
        Case ctSemanaActual
            mvarTexto = lt5
            m_dtFecha_Final = Date + (7 - Weekday(Date, vbMonday))
            m_dtFecha_Inicial = m_dtFecha_Final - 6
        
        Case ctQuincenaActual
            mvarTexto = lt6
            If (Day(Date) > 15) Then
                m_dtFecha_Inicial = CDate("15/" & Month(Date) & "/" _
                               & Year(Date))
                m_dtFecha_Final = CDate(fnUltimoDia(Date) & "/" & Month(Date) & "/" _
                               & Year(Date))
            Else
                m_dtFecha_Inicial = CDate("1/" & Month(Date) & "/" _
                               & Year(Date))
                m_dtFecha_Final = CDate("15/" & Month(Date) & "/" _
                               & Year(Date))
            
            End If
        
        Case ctMesActual
            mvarTexto = lt7
            m_dtFecha_Final = CDate(fnUltimoDia(Date) & "/" _
                              & Month(Date) & "/" _
                              & Year(Date))
            m_dtFecha_Inicial = CDate("1/" & Month(Date) & "/" _
                              & Year(Date))
        
        Case ctAñoActual
            mvarTexto = lt8
            m_dtFecha_Final = CDate("31/12/" & Year(Date))
            m_dtFecha_Inicial = CDate("1/1/" & Year(Date))
        
        
        Case ctLoQueVadeSemana
            mvarTexto = lt9
            m_dtFecha_Final = Date
            m_dtFecha_Inicial = Date - (Weekday(Date, vbMonday) - 1)
        
        Case ctLoQueVadeMes
            mvarTexto = lt10
            m_dtFecha_Final = Date
            m_dtFecha_Inicial = CDate("1/" & Month(Date) & "/" _
                              & Year(Date))
        
        Case ctLoQueVadeAño
            mvarTexto = lt11
            m_dtFecha_Final = Date
            m_dtFecha_Inicial = CDate("1/1/" & Year(Date))
        
        Case ctLoQueVadeTrimestre
            mvarTexto = lt12
            m_dtFecha_Final = Date
            If (Month(Date) >= 9) Then
                m_dtFecha_Inicial = CDate("1/9/" & Year(Date))
            ElseIf (Month(Date) >= 6) Then
                m_dtFecha_Inicial = CDate("1/6/" & Year(Date))
            ElseIf (Month(Date) >= 1) Then
                m_dtFecha_Inicial = CDate("1/1/" & Year(Date))
            End If
            
        Case ctUltimaSemana
            mvarTexto = lt13
            m_dtFecha_Final = Date
            m_dtFecha_Inicial = Date - 6
        
        Case ctUltimaQuincena
            mvarTexto = lt14
            m_dtFecha_Final = Date
            m_dtFecha_Inicial = m_dtFecha_Final - 13
            
        Case ctUltimoMes
            mvarTexto = lt15
            m_dtFecha_Final = Date
            m_dtFecha_Inicial = m_dtFecha_Final - 30
        
        Case ctUltimoTrimestre
            mvarTexto = lt16
            m_dtFecha_Final = Date
            m_dtFecha_Inicial = m_dtFecha_Final - 89
        
        Case ctUltimoAño
            mvarTexto = lt17
            m_dtFecha_Final = Date
            m_dtFecha_Inicial = m_dtFecha_Final - 365
            
        Case ctHastaHoy
            mvarTexto = lt18
            m_dtFecha_Final = Date
            m_dtFecha_Inicial = 0
                    
        Case ctHoy
            mvarTexto = lt19
            m_dtFecha_Final = Date
            m_dtFecha_Inicial = Date
    
        Case ctAyer
            mvarTexto = lt20
            m_dtFecha_Final = Date - 1
            m_dtFecha_Inicial = Date - 1
        
        Case ctMañana
            mvarTexto = lt21
            m_dtFecha_Final = Date + 1
            m_dtFecha_Inicial = Date + 1
        
        Case ctSinDefinir
            mvarTexto = lt_1
            m_dtFecha_Final = MaxDate
            m_dtFecha_Inicial = MinDate

    End Select
   
   On Error GoTo 0
   Exit Property

Tipo_Fecha_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    '   Audita el error
    Call HandleException(ErrNumber, ErrDescription, "Periodo.Tipo_Fecha", ErrSource)
    '   Lanza el Error
    Err.Raise ErrNumber, "Periodo.Tipo_Fecha", ErrDescription
End Property

'------------------------------------------------------------------------------*
' Procedimiento  : Texto
' Fecha          : 22/04/2007 21:49
' Propósito      : Periodo de fechas tipificados del lenguaje natural,
'------------------------------------------------------------------------------*
Public Property Let Texto(ByVal vNewValue As String)
    m_sTipoFecha = vNewValue
    Select Case m_sTipoFecha
        Case lt0: Tipo_Fecha = 0
        Case lt1: Tipo_Fecha = 1
        Case lt2: Tipo_Fecha = 2
        Case lt3: Tipo_Fecha = 3
        Case lt4: Tipo_Fecha = 4
        Case lt5: Tipo_Fecha = 5
        Case lt6: Tipo_Fecha = 6
        Case lt7: Tipo_Fecha = 7
        Case lt8: Tipo_Fecha = 8
        Case lt9: Tipo_Fecha = 9
        Case lt10: Tipo_Fecha = 10
        Case lt11: Tipo_Fecha = 11
        Case lt12: Tipo_Fecha = 12
        Case lt13: Tipo_Fecha = 13
        Case lt14: Tipo_Fecha = 14
        Case lt15: Tipo_Fecha = 15
        Case lt16: Tipo_Fecha = 16
        Case lt17: Tipo_Fecha = 17
        Case lt18: Tipo_Fecha = 18
        Case lt19: Tipo_Fecha = 19
        Case lt20: Tipo_Fecha = 20
        Case lt21: Tipo_Fecha = 21
    End Select
End Property

Public Property Get Texto() As String
    Select Case m_iTipo_Fecha
        Case 0: Texto = lt0
        Case 1: Texto = lt1
        Case 2: Texto = lt2
        Case 3: Texto = lt3
        Case 4: Texto = lt4
        Case 5: Texto = lt5
        Case 6: Texto = lt6
        Case 7: Texto = lt7
        Case 8: Texto = lt8
        Case 9: Texto = lt9
        Case 10: Texto = lt10
        Case 11: Texto = lt11
        Case 12: Texto = lt12
        Case 13: Texto = lt13
        Case 14: Texto = lt14
        Case 15: Texto = lt15
        Case 16: Texto = lt16
        Case 17: Texto = lt17
        Case 18: Texto = lt18
        Case 19: Texto = lt19
        Case 20: Texto = lt20
        Case 21: Texto = lt21
    End Select
End Property


' *===========(EOF): Periodo.cls
