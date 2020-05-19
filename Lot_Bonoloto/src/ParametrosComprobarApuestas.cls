VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ParametrosComprobarApuestas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'---------------------------------------------------------------------------------------
' Module    : ParametrosComprobarApuestas
' Author    : Charly
' Date      : 10/11/2013
' Purpose   : Parametros del proceso Comprobar Apuestas
'---------------------------------------------------------------------------------------

Option Explicit

Public IntervaloFechas As Periodo
Public colApuestas As New Collection
Public ColAciertos As New Collection
Public ColSorteos As New Collection
Public ColEstaditicas As New Collection
Public TipoComparacion As Integer           ' Tipo de comparación:
                                            '   0 Todos los dias,
                                            '   1 Solo vigencia apuesta
Private m_dTotalCostes As Double
Private m_dTotalPremios As Double

'---------------------------------------------------------------------------------------
' Procedure : Class_Initialize
' Author    : Charly
' Date      : 10/11/2013
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub Class_Initialize()
    Set IntervaloFechas = New Periodo
    Set colApuestas = New Collection
    Set ColAciertos = New Collection
    Set ColSorteos = New Collection
    Set ColSorteos = New Collection
    Set ColEstaditicas = New Collection
End Sub

Public Property Get TotalCostes() As Double

    TotalCostes = m_dTotalCostes

End Property

Public Property Get TotalPremios() As Double

    TotalPremios = m_dTotalPremios

End Property

Public Sub InitColAciertos()
    Set ColAciertos = Nothing
    Set ColAciertos = New Collection
End Sub

Public Sub InitColEstadisticas()
    Set ColEstaditicas = Nothing
    Set ColEstaditicas = New Collection
End Sub
