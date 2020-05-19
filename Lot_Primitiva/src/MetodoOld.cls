VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MetodoOld"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'---------------------------------------------------------------------------------------
' Modulo    : Metodo
' Creado    : 16-dic-2006 21:28
' Autor     : Carlos Almela Baeza
' Version   : 1.0.0 Fecha
' Objeto    : Clase que representa el método de obtención de la sugerencia
'---------------------------------------------------------------------------------------
Private m_dtFecha_Inicial As Date
Private m_dtFecha_Final As Date
Private m_dtFecha_Evaluacion As Date
Private m_iTipo_Metodo As Integer
Private m_iPronosticos As Integer
Private m_iDias_Proceso As Integer
Private m_iDias_Retraso As Integer

Const lt_mtd0 = "Aleatorio"
Const lt_mtd1 = "Bombo"
Const lt_mtd2 = "Probabilidad"
Const lt_mtd3 = "Decena"
Const lt_mtd4 = "Tiempo Medio"
Const lt_mtd5 = "Frecuencias"
Const lt_mtd6 = "Decenas Frecuencias"
Const lt_mtd7 = "Proxima Fecha"
Const lt_mtd8 = "Decenas T.Medio"
Const lt_mtd9 = "Ausencias"

Private m_vLiterales_Metodos    As Variant

Public Enum TiposMetodos
    ctAleatorio = 0
    ctBombo = 1
    ctProbabilidad = 2
    ctDecena = 3
    ctTiempoMedio = 4
    ctFrecuencias = 5
    ctDecFrecuencias = 6
    ctFechaProxima = 7
    ctDecTiempoMedio = 8
    ctAusencias = 9
End Enum
'
'----| CargaTabla ( ListControl ) --> Carga una list box con los periodos establecidos
'
Public Sub CargaTabla(cmbLista As Object)

'Si el tipo de objeto no es un control salir
    If Not (TypeOf cmbLista Is Control) Then
        Return
    End If

'Inicializar el control
    cmbLista.Clear
       
'A±adir elementos de la lista
    cmbLista.AddItem lt_mtd0, 0
    cmbLista.AddItem lt_mtd1, 1
    cmbLista.AddItem lt_mtd2, 2
    cmbLista.AddItem lt_mtd3, 3
    cmbLista.AddItem lt_mtd4, 4
    cmbLista.AddItem lt_mtd5, 5
    cmbLista.AddItem lt_mtd6, 6
    cmbLista.AddItem lt_mtd7, 7
    cmbLista.AddItem lt_mtd8, 8
    cmbLista.AddItem lt_mtd9, 9
   
    
'Inicializar el control
    cmbLista.ListIndex = -1

End Sub

'
'       Tipo de Metodo
'
Public Property Get Tipo_Metodo() As Integer
    Tipo_Metodo = m_iTipo_Metodo
End Property

Public Property Let Tipo_Metodo(ByVal vNewData As Integer)
    m_iTipo_Metodo = vNewData
End Property
'
'
'
Public Property Get Fecha_Inicial() As Date
    Fecha_Inicial = m_dtFecha_Inicial
End Property

Public Property Let Fecha_Inicial(ByVal vNewData As Date)
    m_dtFecha_Inicial = vNewData
End Property
'
'
'

Public Property Get Fecha_Final() As Date
    Fecha_Final = m_dtFecha_Final
End Property

Public Property Let Fecha_Final(ByVal vNewData As Date)
    m_dtFecha_Final = vNewData
End Property
'
'
'
Public Property Get Fecha_Evaluacion() As Date
    Fecha_Evaluacion = m_dtFecha_Evaluacion
End Property

Public Property Let Fecha_Evaluacion(ByVal vNewData As Date)
    m_dtFecha_Evaluacion = vNewData
End Property
'
'
'
Public Property Get Pronosticos() As Integer
    Pronosticos = m_iPronosticos
End Property

Public Property Let Pronosticos(ByVal vNewData As Integer)
    m_iPronosticos = vNewData
End Property
'
'
'
Public Property Get Dias_Proceso() As Integer
    Dias_Proceso = m_iDias_Proceso
End Property

Public Property Let Dias_Proceso(ByVal vNewData As Integer)
    m_iDias_Proceso = vNewData
End Property
'
'
'
Public Property Get Nombre_metodo() As String
    Nombre_metodo = m_vLiterales_Metodos(Me.Tipo_Metodo)
End Property

Public Property Let Nombre_metodo(ByVal vNewData As String)
    For i = 0 To UBound(m_vLiterales_Metodos)
        If (UCase(vNewData) = UCase(m_vLiterales_Metodos(i))) Then
            m_iTipo_Metodo = i
            Exit For
        End If
    Next i
End Property
'
'
'
Public Property Get Dias_Retraso() As Integer
    Dias_Retraso = m_iDias_Retraso
End Property

Public Property Let Dias_Retraso(ByVal iDias_Retraso As Integer)
    m_iDias_Retraso = iDias_Retraso
End Property
'
'
'
Public Property Get LiteralesMetodos() As Variant
    LiteralesMetodos = m_vLiterales_Metodos
End Property
'
'
'
Private Sub Class_Initialize()
    m_vLiterales_Metodos = Array(lt_mtd0, lt_mtd1, lt_mtd2, _
                                 lt_mtd3, lt_mtd4, lt_mtd5, _
                                 lt_mtd6, lt_mtd7, lt_mtd8, lt_mtd9)
End Sub
