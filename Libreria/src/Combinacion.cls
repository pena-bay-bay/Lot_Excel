VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Combinacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' *============================================================================*
' *
' *     Fichero    : Combinacion.cls
' *
' *     Autor      : Carlos Almela Baeza
' *     Creación   : 20-abr-2008 11:14
' *     Versión    : 1.1
' *     Propósito  : Determina propiedades de una combinación de números
' *                  de manera que establece criterios de análisis en conjunto
' *
' *============================================================================*
Option Explicit
Option Base 0
'
'--- Variables Privadas -------------------------------------------------------*
Private mCol               As Collection   ' Colección de Numeros de la apuesta
Private mNum               As Numero       ' Numero de la colección
Private mError             As Integer      ' Código de Error
'--- Constantes ---------------------------------------------------------------*
Private Const LT_EMPTYCOMBINACION As String = "#Combinación vacía."
Private Const LT_GUION As String = "-"
'--- Mensajes -----------------------------------------------------------------*
Private Const MSG_COMBIVACIA = "Combinación no válida, no hay números."
Private Const MSG_COMBIMENOSNUM = "Combinación no válida para el juego, faltan números."
Private Const MSG_COMBIMASNUM = "Combinación no válida para el juego, sobran números."
Private Const MSG_COMBINUMMAL = "Combinación con un número erróneo para este juego."

'--- Errores ------------------------------------------------------------------*
Private Const ERR_COMBIVACIA = 1
Private Const ERR_COMBIMENOSNUM = 2
Private Const ERR_COMBIMASNUM = 3
Private Const ERR_COMBINUMMAL = 4
'--- Metodos Privados ---------------------------------------------------------*
Private Sub Class_Initialize()
    Set mCol = New Collection
End Sub

Private Sub Class_Terminate()
    Set mCol = Nothing
End Sub

'--- Propiedades --------------------------------------------------------------*
'---------------------------------------------------------------------------------------
' Procedure : Texto
' Author    : Charly
' Date      : 23/03/2012
' Purpose   : Expresión en formato texto de la combinación, números separados por guion
'---------------------------------------------------------------------------------------
'
Public Property Get Texto() As String
    Texto = ToString(False)
End Property

Public Property Let Texto(ByVal sTexto As String)
    Dim mMatriz       As Variant
    Dim i             As Integer
  On Error GoTo Texto_Error
    mMatriz = Split(sTexto, "-")
    If UBound(mMatriz) = 0 Then
        mMatriz = Split(sTexto, " ")
    End If
    '
    '   Borramos la colección actual
    '
    Clear
    '
    '   Agregamos la mueva
    '
    For i = 0 To UBound(mMatriz)
        Set mNum = New Numero
        mNum.Valor = CInt(mMatriz(i))
        Add mNum
    Next i
    
   On Error GoTo 0
   Exit Property

Texto_Error:
     Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
   ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
   Call HandleException(ErrNumber, ErrDescription, "Combinacion.Texto")
   Err.Raise ErrNumber, ErrSource, ErrDescription
End Property

'------------------------------------------------------------------------------*
' Procedimiento  : Numeros
' Fecha          : 23/03/2012
' Propósito      : Colección de Numeros que componen la combinación
'------------------------------------------------------------------------------*
Public Property Get Numeros() As Collection
  On Error GoTo Items_Error

    Set Numeros = mCol

   On Error GoTo 0
   Exit Property

Items_Error:
     Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
   ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
   Call HandleException(ErrNumber, ErrDescription, "Combinacion.Numeros")
   Err.Raise ErrNumber, ErrSource, ErrDescription
End Property


Public Property Set Numeros(colItems As Collection)

  On Error GoTo Numeros_Error

    Set mCol = colItems

   On Error GoTo 0
   Exit Property

Numeros_Error:
   Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
   ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
   Call HandleException(ErrNumber, ErrDescription, "Combinacion.Numeros")
   Err.Raise ErrNumber, ErrSource, ErrDescription
End Property


'------------------------------------------------------------------------------*
' Procedimiento  : Suma
' Fecha          : 23/03/2012
' Propósito      : Devuelve la suma de los numeros entre si
' Retorno        : Sumatorio de Numeros
'------------------------------------------------------------------------------*
Public Property Get Suma() As Long
    Suma = 0
    For Each mNum In mCol
           Suma = Suma + mNum.Valor
    Next mNum
End Property

'------------------------------------------------------------------------------*
' Procedimiento  : Count
' Fecha          : 23/03/2012
' Propósito      : Devuelve cuantos numeros tiene la combinación
' Retorno        :
'------------------------------------------------------------------------------*
Public Property Get Count() As Integer
    Count = mCol.Count
End Property

'------------------------------------------------------------------------------*
' Procedimiento  : Producto
' Fecha          : 23/03/2012
' Propósito      : Devuelve el producto de los numeros entre si
' Retorno        : Producto de todos los números
'------------------------------------------------------------------------------*
'
Public Property Get Producto() As Double
    Producto = 1
    For Each mNum In mCol
           Producto = Producto * mNum.Valor
    Next mNum
End Property

'------------------------------------------------------------------------------*
' Procedimiento  : FormulaAltoBajo
' Fecha          : 23/03/2012
' Propósito      : Devuelve la formula entre numero menores o iguales a 25
'                  y superiores a 25
' Retorno        : Expresión de Altos y Bajos
'------------------------------------------------------------------------------*
'
Public Property Get FormulaAltoBajo() As String
    Dim m_bajo As Integer
    Dim m_alto As Integer
    Dim m_num As Numero

  On Error GoTo FormulaAltoBajo_Error

    m_bajo = 0
    m_alto = 0
    For Each m_num In mCol
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
   Call HandleException(ErrNumber, ErrDescription, "Combinacion.FormulaAltoBajo")
   Err.Raise ErrNumber, ErrSource, ErrDescription
End Property

'------------------------------------------------------------------------------*
' Procedimiento  : FormulaDecenas
' Fecha          : 23/03/2012
' Propósito      : Devuelve la formula de las decenas a las que pertenecen
'                  los numeros
' Retorno        : Expresión con la formula Decena
'------------------------------------------------------------------------------*
'
Public Property Get FormulaDecenas() As String
    Dim m_contadores(6, 1) As Integer
    Dim m_decena As Integer
    Dim m_sum As Integer
    Dim x As Integer
  On Error GoTo FormulaDecenas_Error

    For Each mNum In mCol
        m_decena = mNum.Decena
        m_sum = m_contadores(m_decena, 1)
        m_sum = m_sum + 1
        m_contadores(m_decena, 0) = m_decena
        m_contadores(m_decena, 1) = m_sum
    Next mNum
        
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
   Call HandleException(ErrNumber, ErrDescription, "Combinacion.FormulaDecenas")
   Err.Raise ErrNumber, ErrSource, ErrDescription
End Property
'------------------------------------------------------------------------------*
' Procedimiento  : FormulaTerminaciones
' Fecha          : 23/03/2012
' Propósito      : Devuelve la formula de las terminaciones de los numeros
' Retorno        : Expresión con la formula terminaciones
'------------------------------------------------------------------------------*
'
Public Property Get FormulaTerminaciones() As String
    Dim m_contadores(9, 1) As Integer
    Dim m_terminacion As Integer
    Dim m_sum As Integer
    Dim m_num As Numero
    Dim x As Integer
    For Each m_num In mCol
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
'------------------------------------------------------------------------------*
' Procedimiento  : FormulaParidad
' Fecha          : 23/03/2012
' Propósito      : Devuelve la formula de los numeros agrupados por par e impar
' Retorno        : Expresión de Numeros Pares e impares
'------------------------------------------------------------------------------*
'
Public Property Get FormulaParidad() As String
    Dim m_par As Integer
    Dim m_impar As Integer
    Dim m_num As Numero

  On Error GoTo FormulaParImpar_Error

    m_par = 0
    m_impar = 0
    For Each m_num In mCol
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
   Call HandleException(ErrNumber, ErrDescription, "Combinacion.FormulaParImpar")
   Err.Raise ErrNumber, ErrSource, ErrDescription
End Property
'------------------------------------------------------------------------------*
' Procedimiento  : FormulaConsecutivos
' Fecha          : 23/03/2012
' Propósito      : Devuelve la formula de los numeros agrupados que están consecutivos
' Retorno        : Expresión de consecutivos
'------------------------------------------------------------------------------*
'
Public Property Get FormulaConsecutivos() As String
    Dim m_contadores(4) As Integer
    Dim m_Numeros() As Integer
    Dim m_i As Integer
    Dim m_dif As Integer
    Dim m_num As Numero
    Dim i As Integer
  On Error GoTo FormulaConsecutivos_Error
    '
    '   Comprobamos que no hay numeros
    '
    If mCol.Count = 0 Then
        FormulaConsecutivos = LT_EMPTYCOMBINACION
        Exit Property
    End If
    
    ReDim m_Numeros(mCol.Count - 1)
    m_i = 0
    For Each m_num In mCol
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
            If Len(FormulaConsecutivos) = 0 Then
                FormulaConsecutivos = m_contadores(i)
            Else
                FormulaConsecutivos = FormulaConsecutivos & "/" & m_contadores(i)
            End If
        End If
    Next i
    
    If Len(FormulaConsecutivos) = 0 Then
        FormulaConsecutivos = "No Consecutivos"
    Else
        FormulaConsecutivos = FormulaConsecutivos & "/0"
    End If

   On Error GoTo 0
   Exit Property

FormulaConsecutivos_Error:
     Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
   ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
   Call HandleException(ErrNumber, ErrDescription, "Combinacion.FormulaConsecutivos")
   Err.Raise ErrNumber, ErrSource, ErrDescription
End Property
'------------------------------------------------------------------------------*
' Procedimiento  : FormulaSeptenas
' Fecha          : 23/03/2012
' Propósito      : Define la formula de los numeros agrupados por septenas (de 7 en 7)
' Retorno        : Expresión de agrupaciones en 7
'------------------------------------------------------------------------------*
'
Public Property Get FormulaSeptenas() As String
    Dim m_contadores(8, 1) As Integer
    Dim m_septena As Integer
    Dim m_sum As Integer
    Dim x As Integer
  On Error GoTo FormulaSeptenas_Error

    For Each mNum In mCol
        m_septena = mNum.Septena
        m_sum = m_contadores(m_septena, 1)
        m_sum = m_sum + 1
        m_contadores(m_septena, 0) = m_septena
        m_contadores(m_septena, 1) = m_sum
    Next mNum
        
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
   Call HandleException(ErrNumber, ErrDescription, "Combinacion.FormulaSeptenas")
   Err.Raise ErrNumber, ErrSource, ErrDescription
End Property

'--- Metodos Publicos ---------------------------------------------------------*
'------------------------------------------------------------------------------*
' Procedimiento  : Clear
' Fecha          : 23/03/2012
' Propósito      : Inicializa la combinación sin numeros
'------------------------------------------------------------------------------*
Public Sub Clear()
    Set mCol = Nothing
    Set mCol = New Collection
End Sub

'------------------------------------------------------------------------------*
' Procedimiento  : Add
' Fecha          : 23/03/2012
' Propósito      : Agrega un número a la combinación
'------------------------------------------------------------------------------*
Public Sub Add(oNum As Numero)
    Dim strKey  As String
  On Error GoTo Add_Error
  
    strKey = oNum.ToString
    
    If mCol.Count <> 0 Then
        mCol.Add oNum, strKey, , mCol.Count
    Else
        mCol.Add oNum, strKey
    End If

Add_CleanExit:
   On Error GoTo 0
    Exit Sub

Add_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    Call HandleException(ErrNumber, ErrDescription, "Combinacion.Add")
    Err.Raise ErrNumber, ErrSource, ErrDescription
End Sub

'------------------------------------------------------------------------------*
' Procedimiento  : Delete
' Fecha          : 23/03/2012
' Propósito      : Elimina un número de la combinación
' Parámetros     : númeno a eliminar
'------------------------------------------------------------------------------*
Public Sub Delete(Key As Integer)
    Dim m_index As Integer
    m_index = 0
    For Each mNum In mCol
        If (mNum.Valor = Key) Then
            Exit For
        End If
        m_index = m_index + 1
    Next mNum
    mCol.Remove m_index
End Sub

'------------------------------------------------------------------------------*
' Procedimiento  : Contiene
' Fecha          : 23/03/2012
' Propósito      : Devuelve true si el número está contenido en la combinación
' Parámetros     : Numero a comprobar
' Retorno        : True si la colección contiene el número
'------------------------------------------------------------------------------*
Public Function Contiene(Numero As Integer) As Boolean
    Dim strKey      As String       ' Clave de busqueda
    Dim prueba      As Variant      ' Resultado de la operación
  On Error Resume Next              ' Desactivamos el error
  
    strKey = Format(Numero, "00")   ' Convertimos el entero en string
                                    ' (clave de la colección)
    prueba = mCol.Item(strKey)      ' Intentamos obtener el elemento de
                                    ' la colección con clave strKey
    Contiene = (Err <> 5)           ' Si no lo encuentra aparece un error 5
                                    ' que en la expresión es verdadero
End Function

'------------------------------------------------------------------------------*
' Procedimiento  : EstaOrdenado
' Fecha          : 15/03/2014
' Propósito      : Devuelve true si los números están en orden ascendente
' Parámetros     :
' Retorno        : True si los números están en orden ascendente
'------------------------------------------------------------------------------*
Public Function EstaOrdenado() As Boolean
    Dim bResult As Boolean
    Dim mNumAnt As Integer

 On Error GoTo EstaOrdenado_Error
    bResult = True
    mNumAnt = 0
    For Each mNum In mCol
        If (mNum.Valor < mNumAnt) Then
            bResult = False
            Exit For
        End If
        mNumAnt = mNum.Valor
    Next mNum

    EstaOrdenado = bResult
   On Error GoTo 0
   Exit Function

EstaOrdenado_Error:
   Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
   ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
   Call HandleException(ErrNumber, ErrDescription, "Combinacion.EstaOrdenado")
   Err.Raise ErrNumber, ErrSource, ErrDescription
End Function
'------------------------------------------------------------------------------*
' Procedimiento  : ToString
' Fecha          : 15/08/2018
' Propósito      : Devuelve una cadena de texto con los números editados
' Parámetros     : OrdenAscendente
' Retorno        : Cadena de Texto con los números separados por "-"
'------------------------------------------------------------------------------*
Public Function ToString(Optional OrdenAscendente As Boolean)
    Dim mNumeros()       As Integer
    Dim mNum             As Numero
    Dim i                As Integer
 On Error GoTo ToString_Error
    '
    '   Asigna el parámetro orden si no se pasa
    '
    If IsMissing(OrdenAscendente) Then
        OrdenAscendente = False
    End If
    '
    ' Inicializamos la variable
    '
    ToString = ""
    '
    '   Si no existen datos en la colección
    '
    If mCol.Count = 0 Then
        ToString = LT_EMPTYCOMBINACION
        Exit Function
    End If
    '
    '   Si queremos la salida ordenada se realiza con matriz
    '
    If OrdenAscendente Then
        '
        '  Redimensionamos una matriz con el número de elementos
        '  de la coleccion
        '
        ReDim mNumeros(mCol.Count - 1)
        '
        ' Inicializamos el indice
        '
        i = 0
        '
        ' Extraemos a la matriz los valores de los número
        '
        For Each mNum In mCol
            '
            '   Asignamos el valor al elemento iesimi
            '
            mNumeros(i) = mNum.Valor
            i = i + 1
        Next mNum
        '
        '   Ordenamos la matriz de forma ascendente
        '
        Ordenar mNumeros, True
        '
        ' Componemos la cadena de los numeros separados por guion
        '
        For i = 0 To UBound(mNumeros)
           ToString = ToString & Format(mNumeros(i), "00") & LT_GUION
        Next i
    Else
        For Each mNum In mCol
            ToString = ToString & mNum.ToString & LT_GUION
        Next mNum

    End If
    '
    ' Quitamos el último guión
    '
    ToString = Left(ToString, Len(ToString) - 1)
    
 On Error GoTo 0
   Exit Function

ToString_Error:
   Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
   ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
   Call HandleException(ErrNumber, ErrDescription, "Combinacion.ToString")
   Err.Raise ErrNumber, ErrSource, ErrDescription
End Function
'------------------------------------------------------------------------------*
' Procedimiento  : EsValido
' Fecha          : 15/08/2018
' Propósito      : Pasando el juego devuelve si el numero es valido o no
' Parámetros     : datJuego
'                  datApuesta, indicador para no validar cuantos numeros
' Retorno        : True si los numeros son válidos para el juego indicado
'------------------------------------------------------------------------------*
Public Function EsValido(Optional datTipoJuego As ModalidadJuego, _
                         Optional datApuesta As Boolean)
    Dim mNumBolas   As Integer
  On Error GoTo EsValido_Error
    '
    '   Inicializamos el error
    '
    mError = 0
    '
    ' si no se pasa el juego asumimos bonoloto
    '
    If IsMissing(datTipoJuego) Then
        datTipoJuego = LP_LB_6_49
    End If
    '
    ' Si no se pasa datApuestas asumimos false
    '
    If IsMissing(datApuesta) Then
        datApuesta = False
    End If
    '
    '  Evaluamos si está vacia
    '
    If mCol.Count = 0 Then
        mError = ERR_COMBIVACIA
    End If
    '
    '   Segun el juego
    '
    Select Case datTipoJuego
        Case LP_LB_6_49: mNumBolas = 6
        Case EU_5_50: mNumBolas = 5
        Case GP_5_54: mNumBolas = 5
        Case EU_2_12: mNumBolas = 2
    End Select
    '
    '  comprobamos que las bolas sean las correctas para el juego
    '
    If Not datApuesta Then
        If mCol.Count < mNumBolas Then
            mError = ERR_COMBIMENOSNUM
        ElseIf mCol.Count > mNumBolas Then
            mError = ERR_COMBIMASNUM
        End If
    End If
    '
    '   Comprobamos cada una de las bolas
    '
    For Each mNum In mCol
        If Not mNum.EsValido(datTipoJuego) Then
            mError = ERR_COMBINUMMAL
        End If
    Next mNum
    '
    '
    '
    EsValido = IIf(mError = 0, True, False)
    
  On Error GoTo 0
EsValido__CleanExit:
    Exit Function
            
EsValido_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    Call HandleException(ErrNumber, ErrDescription, "Combinacion.EsValido", ErrSource)
    Err.Raise ErrNumber, "Combinacion.EsValido", ErrDescription
End Function

'------------------------------------------------------------------------------*
' Procedimiento  : GetMensaje
' Fecha          : 15/08/2018
' Propósito      : Devuelve el mensaje de error si la combinación no es válida
' Retorno        : Mensaje de Error
'------------------------------------------------------------------------------*
Public Function GetMensaje() As String
  On Error GoTo GetMensaje_Error
    '
    '   inicializamos el mensaje
    '
    GetMensaje = Empty
    Select Case mError
        Case ERR_COMBIVACIA: GetMensaje = MSG_COMBIVACIA
        Case ERR_COMBIMASNUM: GetMensaje = MSG_COMBIMASNUM
        Case ERR_COMBIMENOSNUM: GetMensaje = MSG_COMBIMENOSNUM
        Case ERR_COMBINUMMAL: GetMensaje = MSG_COMBINUMMAL
    End Select
    
    On Error GoTo 0
GetMensaje__CleanExit:
    Exit Function
            
GetMensaje_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    Call HandleException(ErrNumber, ErrDescription, "Combinacion.GetMensaje", ErrSource)
    Err.Raise ErrNumber, "Combinacion.GetMensaje", ErrDescription
End Function
' *===========(EOF): Combinacion.cls
