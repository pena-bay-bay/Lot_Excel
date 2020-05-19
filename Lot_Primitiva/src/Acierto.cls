VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Acierto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'---------------------------------------------------------------------------------------
' Module    : Acierto
' Author    : CHARLY
' Date      : 01/03/2014
' Purpose   : Contiene los datos de los aciertos y premios de una apuesta
'---------------------------------------------------------------------------------------

Option Explicit
Option Base 0
'
'
Public IdApuesta As Integer
Public Juego As Juego
Public Categoria As CategoriaPremio
Public BolasAcertadas As Integer
Public ApuestasAcertadas As Integer
Public ReintegroAcertado As Boolean
Public EstrellasAcertadas As Integer
Public ImportePremio As Currency

Private Sub Class_Initialize()
    IdApuesta = 0
    Categoria = Ninguna
    BolasAcertadas = 0
    ApuestasAcertadas = 0
    ReintegroAcertado = False
    EstrellasAcertadas = 0
    ImportePremio = 0
End Sub
