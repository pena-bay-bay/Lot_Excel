VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Bolas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'/**
' * Clase   Bolas -Coleccion-
' * Created 04/07/2005 23:26
' * version 1.0.02
' **/
Option Explicit

'//Propiedades privadas
Private m_col As Collection
Private m_error As Integer
Private Const m_version = "1.0.02"


'// Constantes
' Constructor
Private Sub Class_Initialize()
    Set m_col = New Collection
    m_error = 0
    Debug.Print "Inicio de clase" & "Bolas" & m_version
End Sub
' Destructor
Private Sub Class_Terminate()
    'destruye la colecci�n cuando se termina la clase
    Set m_col = Nothing
End Sub

'/**
' * Property    error
' *
' * @int        c�digo de error producido en la clase
' **/
Public Property Get Error() As Variant
    Error = m_error
End Property

'/**
' * Property    mensaje_error
' *
' * @string     Cadena de texto con la explicaci�n del error
' **/
Public Property Get Mensaje_error() As Variant
    Select Case m_error
     Case 0:    Mensaje_error = MSG_NOERROR
     Case Else: Mensaje_error = MSG_DESCONOCIDO
    End Select
End Property

'/**
' * Property    Count
' *
' * @long       Numero de elementos de la colecci�n
' **/
Public Property Get Count() As Long
    Count = m_col.Count
End Property

'/**
' * method      Add
' *
' * @object     A�ade una bola a la colecci�n
' * @string     Clave de indexaci�n hash
' **/
Public Sub Add(m_bola As Bola)
     m_col.Add m_bola
End Sub

'/**
' * method      Remove
' *
' * @string     Elimina un objeto de la colecci�n aportando la Clave del objeto
' **/
Public Sub Remove(vntIndexKey As Variant)
    m_col.Remove vntIndexKey
End Sub


Public Function Item(vntIndexKey As Variant) As Bola
    Set Item = m_col.Item(vntIndexKey)
End Function


'---------------------------------------------------------------------------------------
' Procedimiento : Coleccion
' Creaci�n      : 14-sep-2006 22:52
' Autor         : Carlos Almela Baeza
' Objeto        :
'---------------------------------------------------------------------------------------
'
Public Property Get Coleccion() As Collection

    Set Coleccion = m_col

End Property

Public Property Set Coleccion(colColleccion As Collection)

    Set m_col = colColleccion

End Property
'/**
' * Property    NewEnum
' *
' * @Object     Devuelve el siguiente objeto de la colecci�n, se
' *             utiliza en bucles
' **/
Public Property Get NewEnum() As IUnknown
    Set NewEnum = m_col.[_NewEnum]
End Property

Public Property Get Sum_Prob() As Double
    Dim m_sum As Double
    Dim m_bola As Bola
    m_sum = 0
    For Each m_bola In m_col
        m_sum = m_sum + m_bola.Probabilidad
    Next
End Property

