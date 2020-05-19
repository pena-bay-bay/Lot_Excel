VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Combinacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'---------------------------------------------------------------------------------------
' Modulo    : Combinación
' Creado    : 20-abr-2008 11:14
' Autor     : Carlos Almela Baeza
' Version   : 1.0.0 20-abr-2008 11:14
' Objeto    : Determina propiedades de una combinación de números
'             de manera que establece criterios de análisis
'             en conjunto
'---------------------------------------------------------------------------------------
Private m_col                       As Collection   ' Colección de Numeros de la apuesta
Private m_sTexto                    As String

'---------------------------------------------------------------------------------------
' Procedure : Texto
' Author    : Charly
' Date      : 23/03/2012
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get Texto() As String
    Dim m_chars As Integer
    Dim m_num As Numero
   On Error GoTo Texto_Error
    Texto = ""
    i = 0
    '
    '   Si no existen datos en la colección
    '
    If m_col.Count = 0 Then
        Texto = "Combinación vacia"
    Else
        For Each m_num In m_col
            Texto = Texto & "-" & Format(m_num.Valor, "00")
        Next m_num
        m_chars = Len(Texto) - 3
        Texto = Right(Texto, m_chars)
    End If

   On Error GoTo 0
   Exit Property

Texto_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Texto of Módulo de clase Combinacion"
End Property

Public Property Let Texto(ByVal sTexto As String)
    Dim m_vNum As Variant
    Dim m_oNum As Numero
    Dim i As Integer
  On Error GoTo Texto_Error

    m_sTexto = sTexto

    m_vNum = Split(sTexto, "-")
    If UBound(m_vNum) = 0 Then
        m_vNum = Split(sTexto, " ")
    End If

    For i = 0 To UBound(m_vNum)
        Set m_oNum = New Numero
        m_oNum.Valor = CInt(m_vNum(i))
        Add m_oNum
    Next i
   On Error GoTo 0
   Exit Property

Texto_Error:
     Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
   ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
   '   Audita el error
   Call HandleException(ErrNumber, ErrDescription, "Combinacion.Texto")
   '   Lanza el Error
   Err.Raise ErrNumber, ErrSource, ErrDescription
End Property

'---------------------------------------------------------------------------------------
' Procedure : Items
' Author    : Charly
' Date      : 23/03/2012
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get Numeros() As Collection
  On Error GoTo Items_Error

    Set Numeros = m_col

   On Error GoTo 0
   Exit Property

Items_Error:
     Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
   ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
   '   Audita el error
   Call HandleException(ErrNumber, ErrDescription, "Combinacion.Items")
   '   Lanza el Error
   Err.Raise ErrNumber, ErrSource, ErrDescription
End Property

'---------------------------------------------------------------------------------------
' Procedure : Numeros
' Author    : Charly
' Date      : 23/03/2012
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Set Numeros(colItems As Collection)

  On Error GoTo Numeros_Error

    Set m_col = colItems

   On Error GoTo 0
   Exit Property

Numeros_Error:
   Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
   ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
   '   Audita el error
   Call HandleException(ErrNumber, ErrDescription, "Combinacion.Items")
   '   Lanza el Error
   Err.Raise ErrNumber, ErrSource, ErrDescription
End Property

'---------------------------------------------------------------------------------------
' Procedure : TextoOrdenado
' Author    : Charly
' Date      : 23/03/2012
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get TextoOrdenado()
    Dim m_num As Numero
    Dim m_chars As Integer
    Dim m_Numeros() As Integer
    Dim i As Integer
   On Error GoTo TextoOrdenado_Error

    ReDim m_Numeros(m_col.Count - 1)
    i = 0
    For Each m_num In m_col
        m_Numeros(i) = m_num.Valor
        i = i + 1
    Next m_num
    
    Ordenar m_Numeros, True
    
    TextoOrdenado = ""
    For i = 0 To UBound(m_Numeros)
        TextoOrdenado = TextoOrdenado & "-" & Format(m_Numeros(i), "00")
    Next i
    m_chars = Len(TextoOrdenado) - 3
    TextoOrdenado = Right(TextoOrdenado, m_chars)

   On Error GoTo 0
   Exit Property

TextoOrdenado_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure TextoOrdenado of Módulo de clase Combinacion"

End Property

'---------------------------------------------------------------------------------------
' Procedure : Suma
' Author    : Charly
' Date      : 23/03/2012
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get Suma() As Long
    Dim m_suma As Long
    Dim m_num As Numero
    m_suma = 0
    For Each m_num In m_col
           m_suma = m_suma + m_num.Valor
    Next m_num
    Suma = m_suma
End Property

'---------------------------------------------------------------------------------------
' Procedure : Count
' Author    : Charly
' Date      : 23/03/2012
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get Count() As Integer
    Count = m_col.Count
End Property

'---------------------------------------------------------------------------------------
' Procedure : Producto
' Author    : Charly
' Date      : 23/03/2012
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get Producto() As Double
    Dim m_prod As Double
    Dim m_num As Numero
    m_prod = 1
    For Each m_num In m_col
           m_prod = m_prod * m_num.Valor
    Next m_num
    Producto = m_prod
End Property


'---------------------------------------------------------------------------------------
' Procedure : FormulaAltoBajo
' Author    : Charly
' Date      : 23/03/2012
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get FormulaAltoBajo() As String
    Dim m_bajo As Integer
    Dim m_alto As Integer
    Dim m_num As Numero

  On Error GoTo FormulaAltoBajo_Error

    m_bajo = 0
    m_alto = 0
    For Each m_num In m_col
        If (m_num.Peso = LT_ALTO) Then
            m_alto = m_alto + 1
        Else
            m_bajo = m_bajo + 1
        End If
    Next m_num
    FormulaAltoBajo = m_bajo & "/" & m_alto

   On Error GoTo 0
   Exit Property

FormulaAltoBajo_Error:
     Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
   ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
   '   Audita el error
   Call HandleException(ErrNumber, ErrDescription, "Combinacion.FormulaAltoBajo")
   '   Lanza el Error
   Err.Raise ErrNumber, ErrSource, ErrDescription
End Property

'---------------------------------------------------------------------------------------
' Procedure : FormulaDecenas
' Author    : Charly
' Date      : 23/03/2012
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get FormulaDecenas() As String
    Dim m_contadores(5, 1) As Integer
    Dim m_decena As Integer
    Dim m_sum As Integer
    Dim m_num As Numero
    
  On Error GoTo FormulaDecenas_Error

    For Each m_num In m_col
        m_decena = m_num.Decena
        m_sum = m_contadores(m_decena, 1)
        m_sum = m_sum + 1
        m_contadores(m_decena, 0) = m_decena
        m_contadores(m_decena, 1) = m_sum
    Next m_num
        
    Ordenar2 m_contadores
        
    FormulaDecenas = m_contadores(0, 1)
    For x = 1 To UBound(m_contadores, 1)
        If (m_contadores(x, 1) > 0) Then
            FormulaDecenas = FormulaDecenas & "/" & m_contadores(x, 1)
        End If
    Next x

   On Error GoTo 0
   Exit Property

FormulaDecenas_Error:
     Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
   ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
   '   Audita el error
   Call HandleException(ErrNumber, ErrDescription, "Combinacion.FormulaDecenas")
   '   Lanza el Error
   Err.Raise ErrNumber, ErrSource, ErrDescription
End Property

'---------------------------------------------------------------------------------------
' Procedure : FormulaTerminaciones
' Author    : Charly
' Date      : 23/03/2012
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get FormulaTerminaciones() As String
    Dim m_contadores(9, 1) As Integer
    Dim m_terminacion As Integer
    Dim m_sum As Integer
    Dim m_num As Numero
    
    For Each m_num In m_col
        m_terminacion = m_num.Terminacion
        m_sum = m_contadores(m_terminacion, 1)
        m_sum = m_sum + 1
        m_contadores(m_terminacion, 0) = m_terminacion
        m_contadores(m_terminacion, 1) = m_sum
    Next m_num
        
    Ordenar2 m_contadores
        
    FormulaTerminaciones = m_contadores(0, 1)
    For x = 1 To UBound(m_contadores, 1)
        If (m_contadores(x, 1) > 0) Then
            FormulaTerminaciones = FormulaTerminaciones & "/" & m_contadores(x, 1)
        End If
    Next x
End Property

'---------------------------------------------------------------------------------------
' Procedure : FormulaParImpar
' Author    : Charly
' Date      : 23/03/2012
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get FormulaParidad() As String
    Dim m_par As Integer
    Dim m_impar As Integer
    Dim m_num As Numero

  On Error GoTo FormulaParImpar_Error

    m_par = 0
    m_impar = 0
    For Each m_num In m_col
        If (m_num.Paridad = LT_PAR) Then
            m_par = m_par + 1
        Else
            m_impar = m_impar + 1
        End If
    Next m_num
    FormulaParidad = m_par & "/" & m_impar

   On Error GoTo 0
   Exit Property

FormulaParImpar_Error:
     Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
   ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
   '   Audita el error
   Call HandleException(ErrNumber, ErrDescription, "Combinacion.FormulaParImpar")
   '   Lanza el Error
   Err.Raise ErrNumber, ErrSource, ErrDescription
End Property

'---------------------------------------------------------------------------------------
' Procedure : FormulaConsecutivos
' Author    : Charly
' Date      : 23/03/2012
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get FormulaConsecutivos() As String
    Dim m_contadores(4) As Integer
    Dim m_Numeros() As Integer
    Dim m_i As Integer
    Dim m_dif As Integer
    Dim m_num As Numero
  On Error GoTo FormulaConsecutivos_Error

    ReDim m_Numeros(m_col.Count - 1)
    m_i = 0
    For Each m_num In m_col
        m_Numeros(m_i) = m_num.Valor
        m_i = m_i + 1
    Next m_num
    
    Ordenar m_Numeros, True
    
    m_i = 0
    For i = 0 To UBound(m_Numeros) - 2
       m_dif = m_Numeros(i + 1) - m_Numeros(i)
       If (m_dif = 1) Then
            m_contadores(m_i) = IIf(m_contadores(m_i) = 0, m_contadores(m_i) + 2, m_contadores(m_i) + 1)
       Else
            If (m_contadores(m_i) > 0) Then
                m_i = m_i + 1
            End If
       End If
    Next i
    
    Ordenar m_contadores, False
    FormulaConsecutivos = ""
    For i = 0 To UBound(m_contadores) - 1
        If (m_contadores(i) > 0) Then
            FormulaConsecutivos = FormulaConsecutivos & "/" & m_contadores(i)
        End If
    Next i
    
    If Len(FormulaConsecutivos) = 0 Then
        FormulaConsecutivos = "No Consecutivos"
    Else
        FormulaConsecutivos = Mid(FormulaConsecutivos, 2, Len(FormulaConsecutivos) - 1)
    End If

   On Error GoTo 0
   Exit Property

FormulaConsecutivos_Error:
     Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
   ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
   '   Audita el error
   Call HandleException(ErrNumber, ErrDescription, "Combinacion.FormulaConsecutivos")
   '   Lanza el Error
   Err.Raise ErrNumber, ErrSource, ErrDescription
End Property

'---------------------------------------------------------------------------------------
' Procedure : FormulaSeptenas
' Author    : Charly
' Date      : 23/03/2012
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get FormulaSeptenas() As String
    Dim m_contadores(7, 1) As Integer
    Dim m_septena As Integer
    Dim m_sum As Integer
    Dim m_num As Numero
    
  On Error GoTo FormulaSeptenas_Error

    For Each m_num In m_col
        m_septena = m_num.Septena
        m_sum = m_contadores(m_septena, 1)
        m_sum = m_sum + 1
        m_contadores(m_septena, 0) = m_septena
        m_contadores(m_septena, 1) = m_sum
    Next m_num
        
    Ordenar2 m_contadores
        
    FormulaSeptenas = m_contadores(0, 1)
    For x = 1 To UBound(m_contadores, 1)
        If (m_contadores(x, 1) > 0) Then
            FormulaSeptenas = FormulaSeptenas & "/" & m_contadores(x, 1)
        End If
    Next x

   On Error GoTo 0
   Exit Property

FormulaSeptenas_Error:
     Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
   ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
   '   Audita el error
   Call HandleException(ErrNumber, ErrDescription, "Combinacion.FormulaSeptenas")
   '   Lanza el Error
   Err.Raise ErrNumber, ErrSource, ErrDescription
End Property

Private Sub Class_Initialize()
    Set m_col = New Collection
End Sub

Private Sub Class_Terminate()
    Set m_col = Nothing
End Sub

'---------------------------------------------------------------------------------------
' Procedure : Clear
' Author    : Charly
' Date      : 23/03/2012
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Sub Clear()
    EliminarTodosElementos m_col
End Sub

Public Sub Add(oNum As Numero)
    Dim strKey  As String
  On Error GoTo Add_Error
  
    strKey = CStr(oNum.Valor)
    If m_col.Count <> 0 Then
        m_col.Add oNum, strKey, , m_col.Count
    Else
        m_col.Add oNum, strKey
    End If

Add_CleanExit:
   On Error GoTo 0
    Exit Sub

Add_Error:
    

    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    '   Audita el error
    Call HandleException(ErrNumber, ErrDescription, "Combinacion.Add")
    '   Lanza el Error
    Err.Raise ErrNumber, ErrSource, ErrDescription

End Sub

'---------------------------------------------------------------------------------------
' Procedure : Delete
' Author    : Charly
' Date      : 23/03/2012
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Sub Delete(key As Integer)
    Dim m_index As Integer
    Dim m_num As Numero
    Index = 0
    For Each m_num In m_col
        If (m_num.Valor = key) Then
            Exit For
        End If
        m_index = m_index + 1
    Next m_num
    m_col.Remove m_index
End Sub

'---------------------------------------------------------------------------------------
' Procedure : Contiene
' Author    : Charly
' Date      : 23/03/2012
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Function Contiene(Numero As Integer) As Boolean
    Dim m_sNum      As String       ' Clave de busqueda
    Dim prueba      As Variant      ' Resultado de la operación
  On Error Resume Next              ' Desactivamos el error
    m_sNum = CStr(Numero)           ' Convertimos el entero en string
                                    ' (clave de la colección)
    prueba = m_col.Item(m_sNum)     ' Intentamos obtener el elemento de
                                    ' la colección con clave m_sNum
    Contiene = (Err <> 5)           ' Si no lo encuentra aparece un error 5
                                    ' que en la expresión es verdadero
End Function



'---------------------------------------------------------------------------------------
' Procedure : EstaOrdenado
' Author    : CHARLY
' Date      : 15/03/2014
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Function EstaOrdenado() As Boolean
    Dim bResult As Boolean
    Dim m_num As Numero
    Dim m_numAnt As Integer

   On Error GoTo EstaOrdenado_Error
    bResult = True
    m_numAnt = 0
    For Each m_num In m_col
        If (m_num.Valor < m_numAnt) Then
            bResult = False
            Exit For
        End If
        m_numAnt = m_num.Valor
    Next m_num

    EstaOrdenado = bResult
   On Error GoTo 0
   Exit Function

EstaOrdenado_Error:
   Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
   ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
   '   Audita el error
   Call HandleException(ErrNumber, ErrDescription, "Combinacion.EstaOrdenado")
   '   Lanza el Error
   Err.Raise ErrNumber, ErrSource, ErrDescription
End Function
