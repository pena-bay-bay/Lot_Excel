VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sugerencia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'---------------------------------------------------------------------------------------
' Module    : Sugerencia
' Author    : Charly
' Date      : 13/04/2012
' Purpose   :
'---------------------------------------------------------------------------------------

Option Explicit

Public Entidad As EntidadNegocio
Public metodo As metodo
Public Parametros As ParametrosMuestra
Private m_colBolas As bolas
Private m_eModalidad As ModalidadJuego
Private m_objCombinacion As Combinacion

Private Sub Class_Initialize()
    Set m_colBolas = New bolas
    Set m_objCombinacion = New Combinacion
End Sub

Private Sub Class_Terminate()

    Set m_colBolas = Nothing

    Set m_objCombinacion = Nothing
End Sub


Public Property Get bolas() As bolas

    Set bolas = m_colBolas

End Property

Public Property Set bolas(colBolas As bolas)

    Set m_colBolas = colBolas

End Property

Public Property Get Modalidad() As ModalidadJuego

    Modalidad = m_eModalidad

End Property

Public Property Let Modalidad(ByVal eModalidad As ModalidadJuego)

    m_eModalidad = eModalidad

End Property

Public Property Get Combinacion() As Combinacion

    Set Combinacion = m_objCombinacion

End Property

Public Property Set Combinacion(objCombinacion As Combinacion)

    Set m_objCombinacion = objCombinacion

End Property

Public Function Validar() As Boolean

    Dim bResult As Boolean


    Validar = bResult

End Function

Public Function MensajeError() As String

    Dim sResult As String


    MensajeError = sResult

End Function

Public Function ToString() As String

    Dim sResult As String


    ToString = sResult

End Function

