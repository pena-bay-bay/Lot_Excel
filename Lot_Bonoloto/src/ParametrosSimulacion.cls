VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ParametrosSimulacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Private mArrayMetodos() As Variant
Private mCol As Collection
Private mRangoRevision As Periodo
Private mPronosticos As Integer
'
'
'
'

Private Sub Class_Initialize()
    Set mCol = New Collection
    Set mRangoRevision = New Periodo
End Sub

Private Sub Class_Terminate()
    'destruye la colección cuando se termina la clase
    Set mCol = Nothing
    Set mRangoRevision = Nothing
End Sub


Public Sub Remove(vNewData As ParametrosMetodoOld)
    Dim strKey As String
    strKey = CStr(vNewData.Id)
    mCol.Remove strKey
End Sub

Public Sub Add(vNewData As ParametrosMetodoOld)
    Dim strKey As String
    strKey = CStr(vNewData.Id)
    mCol.Add vNewData, strKey
End Sub

Public Sub Change(vNewData As ParametrosMetodoOld)
    Dim mMtdAntiguo As ParametrosMetodoOld
    Dim i As Integer
    i = vNewData.Id
    Set mMtdAntiguo = mCol.Item(i)
    SustituirElem mCol, i, vNewData
End Sub

Public Property Get NumMetodos() As Integer
    NumMetodos = mCol.Count
End Property

Public Sub Clear()
    EliminarTodosElementos mCol
End Sub

Public Function GetNewMetodo() As ParametrosMetodoOld
    Dim key As Integer
    Dim mMtd As ParametrosMetodoOld
    Dim mParam As ParametrosMetodoOld
    Set mParam = New ParametrosMetodoOld
    key = 0
    For Each mMtd In mCol
        If (mMtd.Id > key) Then
            key = mMtd.Id
        End If
    Next mMtd
    mParam.Id = key + 1
    Set GetNewMetodo = mParam
End Function
'
'
'
Public Property Get RangoAnalisis() As Periodo
    Set RangoAnalisis = mRangoRevision
End Property

Public Property Set RangoAnalisis(ByVal vNewValue As Periodo)
    Set mRangoRevision = vNewValue
End Property

Public Property Get FechaInicial() As Date
    FechaInicial = mRangoRevision.FechaInicial
End Property

Public Property Get FechaFinal() As Date
    FechaFinal = mRangoRevision.FechaFinal
End Property

Public Property Get Dias() As Integer
    Dias = mRangoRevision.Dias
End Property

Public Property Get Metodos() As Variant
    Set Metodos = mCol
End Property

Public Property Get Pronosticos() As Integer
    Pronosticos = mPronosticos
End Property

Public Property Let Pronosticos(ByVal vNewValue As Integer)
    mPronosticos = vNewValue
End Property

Public Property Get ArrayMetodos() As Variant
    ArrayMetodos = mArrayMetodos
End Property

Public Property Let ArrayMetodos(ByVal vNewValue As Variant)
    mArrayMetodos = vNewValue
End Property

