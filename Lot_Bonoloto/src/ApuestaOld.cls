VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ApuestaOld"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
    '---------------------------------------------------------------------------------------
' Modulo    : Apuesta
' Creado    : 16-dic-2006 23:52
' Autor     : Carlos Almela Baeza
' Version   : 1.0.0 10-ago-2007 13:46
' Objeto    : Clase que serializa una apuesta realizada para una fecha
'---------------------------------------------------------------------------------------
Option Explicit

Private m_dtFecha As Date                   ' Fecha para la que se realiza una apuesta
Private m_objCombinacion As Combinacion     ' Combinación de números
Private m_objMetodo As metodo               ' Metodo de obtención de la apuesta
Private m_iId As Integer                    ' Identificador de la apuestas
Private m_iIdBoleto As Integer              ' Identificador del boleto
Private m_iAciertos As Integer              ' Numero de aciertos
Private m_sMetodo As String                 ' Nombre del metodo


'---------------------------------------------------------------------------------------
' Procedure : Class_Initialize
' Author    : Charly
' Date      : 23/03/2012
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub Class_Initialize()
    Set m_objMetodo = New metodo
    Set m_objCombinacion = New Combinacion
End Sub

'---------------------------------------------------------------------------------------
' Procedure : Class_Terminate
' Author    : Charly
' Date      : 23/03/2012
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub Class_Terminate()
    Set m_objMetodo = Nothing
    Set m_objCombinacion = Nothing
End Sub

'---------------------------------------------------------------------------------------
' Procedure : metodo
' Author    : Charly
' Date      : 23/03/2012
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get metodo() As MetodoOld
    Set metodo = m_objMetodo
End Property

Public Property Set metodo(objMetodo As MetodoOld)
    Set m_objMetodo = objMetodo
End Property

'---------------------------------------------------------------------------------------
' Procedure : Numeros
' Author    : Charly
' Date      : 23/03/2012
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get Numeros() As Variant
    Dim m_vNumeros(9) As Integer
    Dim m_num As bola
    Dim i As Integer
    i = 0
    ReDim vNumeros(m_objCombinacion.Count)
    For Each m_num In m_objCombinacion.Numeros
        m_vNumeros(i) = m_num.Numero.Valor
        i = i + 1
    Next m_num
    Numeros = m_vNumeros
End Property

'---------------------------------------------------------------------------------------
' Procedure : Numeros
' Author    : Charly
' Date      : 23/03/2012
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Let Numeros(ByVal vNumeros As Variant)
    Dim m_num As bola
    Dim i As Integer
    If IsArray(vNumeros) Then
        Set m_objCombinacion = Nothing
        For i = 0 To UBound(vNumeros)
            Set m_num = New bola
            m_num.Numero.Valor = vNumeros(i)
            m_objCombinacion.Add m_num.Numero
        Next i
    End If
End Property

'---------------------------------------------------------------------------------------
' Procedure : Id
' Author    : Charly
' Date      : 23/03/2012
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get Id() As Integer
    Id = m_iId
End Property

'---------------------------------------------------------------------------------------
' Procedure : IdBoleto
' Author    : Charly
' Date      : 23/03/2012
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get IdBoleto() As Integer
    IdBoleto = m_iIdBoleto
End Property

'---------------------------------------------------------------------------------------
' Procedure : aciertos
' Author    : Charly
' Date      : 23/03/2012
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get aciertos() As Integer
    aciertos = m_iAciertos
End Property

Public Property Let aciertos(ByVal iAciertos As Integer)
    m_iAciertos = iAciertos
End Property

'---------------------------------------------------------------------------------------
' Procedimiento : Fecha
' Creación      : 17-dic-2006 00:27
' Autor         : Carlos Almela Baeza
' Objeto        : Fecha para la cual se realiza la apuesta
'---------------------------------------------------------------------------------------
'
Public Property Get Fecha() As Date
    Fecha = m_dtFecha
End Property

Public Property Let Fecha(ByVal dtFecha As Date)
    m_dtFecha = dtFecha
End Property

'---------------------------------------------------------------------------------------
' Procedimiento : Total_Pronosticos
' Creación      : 17-dic-2006 00:28
' Autor         : Carlos Almela Baeza
' Objeto        :
'---------------------------------------------------------------------------------------
'
Public Property Get Total_Pronosticos() As Integer
    Total_Pronosticos = m_objCombinacion.Count
End Property

'---------------------------------------------------------------------------------------
' Procedure : Total_Apuestas
' DateTime  : 12/ago/2007 18:55
' Author    : Carlos Almela Baeza
' Purpose   : Total de apuestas que se realizan con estos pronósticos
'---------------------------------------------------------------------------------------
'
Public Property Get Total_Apuestas() As Integer
    Select Case m_objCombinacion.Count
        Case 5: Total_Apuestas = 44
        Case 6: Total_Apuestas = 1
        Case 7: Total_Apuestas = 7
        Case 8: Total_Apuestas = 28
        Case 9: Total_Apuestas = 84
        Case 10: Total_Apuestas = 210
        Case 11: Total_Apuestas = 462
        Case Else: Total_Apuestas = 0
    End Select
End Property

'---------------------------------------------------------------------------------------
' Procedure : Texto
' DateTime  : 12/ago/2007 18:55
' Author    : Carlos Almela Baeza
' Purpose   : Cadena de texto definiendo la apuesta
'---------------------------------------------------------------------------------------
'
Public Property Get Texto() As Variant
    Texto = m_objCombinacion.Texto
End Property

'---------------------------------------------------------------------------------------
' Procedure : Multiple
' DateTime  : 12/ago/2007 18:56
' Author    : Carlos Almela Baeza
' Purpose   : Tipo de apuesta Multiple o sencilla
'---------------------------------------------------------------------------------------
'
Public Property Get Multiple() As Boolean
    If m_objCombinacion.Count = 6 Then
        Multiple = False
    Else
        Multiple = True
    End If
End Property

'---------------------------------------------------------------------------------------
' Procedimiento : Pronosticos
' Creación      : 16-ene-2007 08:03
' Autor         : Carlos Almela Baeza
' Objeto        : Devuelve una matriz ordenada con los números de la
'                 apuesta
'---------------------------------------------------------------------------------------
'
Public Property Get Pronosticos() As Variant
        Dim m_num() As Integer          ' Declara la matriz de enteros
        Dim i As Integer                ' Declara un contador de números
        Dim m_oprt As Pronostico        ' Declara el objeto pronostico
        
        ReDim m_num(m_objCombinacion.Count - 1)    ' Redimensiona la matríz con el total de números
        i = 0                           ' Inicializa contador
        For Each m_oprt In m_objCombinacion.Numeros      ' Para cada pronostico en la coleccion
            m_num(i) = m_oprt.Valor     ' Asigna el valor del número a la matriz
            i = i + 1                   ' incrementa el contador
        Next m_oprt
    
        Ordenar m_num, True             ' Ordena de forma ascendente la matriz de Numeros
        
        Pronosticos = m_num             ' Devuelve la matriz de pronosticos
End Property

'---------------------------------------------------------------------------------------
' Procedure : Formula_Pares
' DateTime  : 12/ago/2007 18:57
' Author    : Carlos Almela Baeza
' Purpose   : Formula de Numeros pares / impares
'---------------------------------------------------------------------------------------
'
Public Property Get Formula_Pares() As String
    Formula_Pares = m_objCombinacion.FormulaParidad
End Property

'---------------------------------------------------------------------------------------
' Procedure : Suma
' DateTime  : 12/ago/2007 18:58
' Author    : Carlos Almela Baeza
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get Suma() As Integer
    Suma = m_objCombinacion.Suma
End Property

'---------------------------------------------------------------------------------------
' Procedure : Producto
' DateTime  : 12/ago/2007 18:59
' Author    : Carlos Almela Baeza
' Purpose   : Producto de los Numeros
'---------------------------------------------------------------------------------------
'
Public Property Get Producto() As Double
    Producto = m_objCombinacion.Producto
End Property

'---------------------------------------------------------------------------------------
' Procedure : Formula_AltosBajos
' DateTime  : 12/ago/2007 18:59
' Author    : Carlos Almela Baeza
' Purpose   : Formula de números Bajos / Altos
'---------------------------------------------------------------------------------------
'
Public Property Get FormulaAltoBajo() As String
    FormulaAltoBajo = m_objCombinacion.FormulaAltoBajo
End Property

'---------------------------------------------------------------------------------------
' Procedure : Formula_Decena
' DateTime  : 12/ago/2007 19:01
' Author    : Carlos Almela Baeza
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get FormulaDecenas() As String
    FormulaDecenas = m_objCombinacion.FormulaDecenas
End Property

'---------------------------------------------------------------------------------------
' Procedure : Formula_Terminaciones
' DateTime  : 12/ago/2007 19:01
' Author    : Carlos Almela Baeza
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get FormulaTerminaciones() As String
    FormulaTerminaciones = m_objCombinacion.FormulaTerminaciones
End Property

'---------------------------------------------------------------------------------------
' Procedure : Existe
' DateTime  : 12/ago/2007 18:56
' Author    : Carlos Almela Baeza
' Purpose   : Determina si un número está en la apuesta
'---------------------------------------------------------------------------------------
'
Public Function Existe(datNumero As Integer) As Boolean
    Existe = m_objCombinacion.Contiene(datNumero)
End Function


Public Sub Constructor(Fila As Range)
    Dim i            As Integer
    Dim m_objNum     As Numero
   
   On Error GoTo Constructor_Error
    With Fila
        m_iId = .Cells(1, 1).Value
        m_iIdBoleto = .Cells(1, 2).Value
        m_dtFecha = CDate(.Cells(1, 3).Value)
        m_objCombinacion.Clear
        For i = 0 To 8
            If IsNumeric(.Cells(1, 4 + i).Value) _
            And Not IsEmpty(.Cells(1, 4 + i).Value) Then
                Set m_objNum = New Numero
                m_objCombinacion.Add (m_objNum)
            End If
        Next i
        m_iAciertos = .Cells(1, 13).Value
        ' premio como columna añadida
        m_sMetodo = .Cells(1, 14).Value
    End With
   On Error GoTo 0
   Exit Sub

Constructor_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Constructor of M¾dulo de clase Apuesta"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : Combinacion
' Author    : Charly
' Date      : 23/03/2012
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get Combinacion() As Combinacion

   On Error GoTo Combinacion_Error

    Set Combinacion = m_objCombinacion

   On Error GoTo 0
   Exit Property

Combinacion_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Combinacion of Módulo de clase Apuesta"

End Property

'---------------------------------------------------------------------------------------
' Procedure : Combinacion
' Author    : Charly
' Date      : 23/03/2012
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Set Combinacion(objCombinacion As Combinacion)

   On Error GoTo Combinacion_Error

    Set m_objCombinacion = objCombinacion

   On Error GoTo 0
   Exit Property

Combinacion_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Combinacion of Módulo de clase Apuesta"

End Property
