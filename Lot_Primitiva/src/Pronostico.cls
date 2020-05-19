VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Pronostico"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'---------------------------------------------------------------------------------------
' Modulo    : Pronostico
' Creado    : 25-dic-2006 21:15
' Autor     : Carlos Almela Baeza
' Version   : 1.0.0 Fecha
' Objeto    : Clase que representa un número seleccionado para una apuesta
'---------------------------------------------------------------------------------------
Private m_objBola As bola
Private m_objMetodo As MetodoOld
'---------------------------------------------------------------------------------------
' Procedimiento : Class_Initialize
' Creación      : 12-nov-2006 00:15
' Autor         : Carlos Almela Baeza
' Objeto        :
'---------------------------------------------------------------------------------------
'
Private Sub Class_Initialize()
    Set m_objBola = New bola
    Set m_objMetodo = New MetodoOld
End Sub

Private Sub Class_Terminate()
    Set m_objBola = Nothing
    Set m_objMetodo = Nothing
End Sub

'---------------------------------------------------------------------------------------
' Procedure : Paridad
' DateTime  : 04/03/2007 23:03
' Author    : Carlos Almela Baeza
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get Paridad() As String
    Paridad = m_objBola.Numero.Paridad
End Property

'---------------------------------------------------------------------------------------
' Procedure : Peso
' DateTime  : 04/03/2007 23:03
' Author    : Carlos Almela Baeza
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get Peso() As String
    Peso = m_objBola.Numero.Peso
End Property

'---------------------------------------------------------------------------------------
' Procedure : Decena
' DateTime  : 04/03/2007 23:03
' Author    : Carlos Almela Baeza
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get Decena() As String
    Decena = m_objBola.Numero.Decena
End Property

'---------------------------------------------------------------------------------------
' Procedure : Terminacion
' DateTime  : 04/03/2007 23:04
' Author    : Carlos Almela Baeza
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get Terminacion() As String
    Terminacion = m_objBola.Numero.Terminacion
End Property


'---------------------------------------------------------------------------------------
' Procedure : Valor
' DateTime  : 04/03/2007 23:04
' Author    : Carlos Almela Baeza
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get Valor() As Integer
    Valor = m_objBola.Numero.Valor
End Property

Public Property Let Valor(ByVal iValor As Integer)
    m_objBola.Numero.Valor = iValor
End Property

'---------------------------------------------------------------------------------------
' Procedure : Bola
' DateTime  : 04/03/2007 23:04
' Author    : Carlos Almela Baeza
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get bola() As bola
    Set bola = m_objBola
End Property

Public Property Set bola(vNewValor As bola)
    Set m_objBola = vNewValor
End Property

'---------------------------------------------------------------------------------------
' Procedimiento : Metodo
' Creación      : 25-dic-2006 21:22
' Autor         : Carlos Almela Baeza
' Objeto        :
'---------------------------------------------------------------------------------------
'
Public Property Get metodo() As metodo
    Set metodo = m_objMetodo
End Property

Public Property Set metodo(objMetodo As metodo)
    Set m_objMetodo = objMetodo
End Property
