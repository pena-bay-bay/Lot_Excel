VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "bolas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'/**
' * Clase   Bolas -Coleccion-
' * Created 04/07/2005 23:26
' * version 1.0.02
' **/

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
    'destruye la colección cuando se termina la clase
    Set mCol = Nothing
End Sub

'/**
' * Property    error
' *
' * @int        código de error producido en la clase
' **/
Public Property Get Error() As Variant
    Error = m_error
End Property

'/**
' * Property    mensaje_error
' *
' * @string     Cadena de texto con la explicación del error
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
' * @long       Numero de elementos de la colección
' **/
Public Property Get Count() As Long
    Count = m_col.Count
End Property

'/**
' * method      Add
' *
' * @object     Añade una bola a la colección
' * @string     Clave de indexación hash
' **/
Public Sub Add(m_bola As bola)
     m_col.Add m_bola
End Sub

'/**
' * method      Remove
' *
' * @string     Elimina un objeto de la colección aportando la Clave del objeto
' **/
Public Sub Remove(vntIndexKey As Variant)
    m_col.Remove vntIndexKey
End Sub


Public Function Item(vntIndexKey As Variant) As bola
    Set Item = m_col.Item(vntIndexKey)
End Function


'---------------------------------------------------------------------------------------
' Procedimiento : Coleccion
' Creación      : 14-sep-2006 22:52
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
' * @Object     Devuelve el siguiente objeto de la colección, se
' *             utiliza en bucles
' **/
Public Property Get NewEnum() As IUnknown
    Set NewEnum = m_col.[_NewEnum]
End Property

Public Property Get Sum_Prob() As Double
    Dim m_sum As Double
    Dim m_bola As bola
    m_sum = 0
    For Each m_bola In m_col
        m_sum = m_sum + m_bola.Probabilidad
    Next
End Property

