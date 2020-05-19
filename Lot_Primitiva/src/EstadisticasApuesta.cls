VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EstadisticasApuesta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Public IdApuesta As Integer
Public Costes As Double
Public ImportePremios As Double
Public Puntuacion As Integer
Public DiasAciertos As Integer


Private Sub Class_Initialize()
    IdApuesta = 0
    Costes = 0
    ImportePremios = 0
    Puntuacion = 0
    DiasAciertos = 0
End Sub
