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
' *     Creaci�n   : 20-abr-2008 11:14
' *     Versi�n    : 1.1
' *     Prop�sito  : Determina propiedades de una combinaci�n de n�meros
' *                  de manera que establece criterios de an�lisis en conjunto
' *
' *============================================================================*
Option Explicit
Option Base 0
'
'--- Variables Privadas -------------------------------------------------------*
Private mCol               As Collection   ' Colecci�n de Numeros de la apuesta
Private mNum               As Numero       ' Numero de la colecci�n
Private mError             As Integer      ' C�digo de Error
'--- Constantes ---------------------------------------------------------------*
Private Const LT_EMPTYCOMBINACION As String = "#Combinaci�n vac�a."
Private Const LT_GUION As String = "-"
'--- Mensajes -----------------------------------------------------------------*
Private Const MSG_COMBIVACIA = "Combinaci�n no v�lida, no hay n�meros."
Private Const MSG_COMBIMENOSNUM = "Combinaci�n no v�lida para el juego, faltan n�meros."
Private Const MSG_COMBIMASNUM = "Combinaci�n no v�lida para el juego, sobran n�meros."
Private Const MSG_COMBINUMMAL = "Combinaci�n con un n�mero err�neo para este juego."

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
' Purpose   : Expresi�n en formato texto de la combinaci�n, n�meros separados por guion
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
    '   Borramos la colecci�n actual
    '
    Clear
    '
    '   Agregamos la mueva
    '
    For i = 0 To UBound(mMatriz)
        Set mNum = New Numero
        'If CInt(mMatriz(i)) > 0 Then
        If IsNumeric(mMatriz(i)) Then
            mNum.Valor = CInt(mMatriz(i))
            mNum.Orden = i + 1
            Add mNum
        End If
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
' Prop�sito      : Colecci�n de Numeros que componen la combinaci�n
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
' Prop�sito      : Devuelve la suma de los numeros entre si
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
' Prop�sito      : Devuelve cuantos numeros tiene la combinaci�n
' Retorno        :
'------------------------------------------------------------------------------*
Public Property Get Count() As Integer
    Count = mCol.Count
End Property

'------------------------------------------------------------------------------*
' Procedimiento  : Producto
' Fecha          : 23/03/2012
' Prop�sito      : Devuelve el producto de los numeros entre si
' Retorno        : Producto de todos los n�meros
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
' Prop�sito      : Devuelve la formula entre numero menores o iguales a 25
'                  y superiores a 25
' Retorno        : Expresi�n de Altos y Bajos
'------------------------------------------------------------------------------*
'
Public Property Get FormulaAltoBajo() As String
    Dim m_bajo As Integer
    Dim m_alto As Integer

  On Error GoTo FormulaAltoBajo_Error

    m_bajo = 0
    m_alto = 0
    For Each mNum In mCol
        If (mNum.Peso = LT_ALTO) Then
            m_alto = m_alto + 1
        Else
            m_bajo = m_bajo + 1
        End If
    Next mNum
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
' Prop�sito      : Devuelve la formula de las decenas a las que pertenecen
'                  los numeros
' Retorno        : Expresi�n con la formula Decena
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
        
    Ordenar2 m_contadores, , False
        
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
' Prop�sito      : Devuelve la formula de las terminaciones de los numeros
' Retorno        : Expresi�n con la formula terminaciones
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
        
    Ordenar2 m_contadores, , False
        
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
' Prop�sito      : Devuelve la formula de los numeros agrupados por par e impar
' Retorno        : Expresi�n de Numeros Pares e impares
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
' Prop�sito      : Devuelve la formula de los numeros agrupados que est�n consecutivos
' Retorno        : Expresi�n de consecutivos
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
        FormulaConsecutivos = "NoConsecutivos"
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
' Prop�sito      : Define la formula de los numeros agrupados por septenas (de 7 en 7)
' Retorno        : Expresi�n de agrupaciones en 7
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
        
    Ordenar2 m_contadores, , False
        
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
' Prop�sito      : Inicializa la combinaci�n sin numeros
'------------------------------------------------------------------------------*
Public Sub Clear()
    Set mCol = Nothing
    Set mCol = New Collection
End Sub

'------------------------------------------------------------------------------*
' Procedimiento  : Add
' Fecha          : 23/03/2012
' Prop�sito      : Agrega un n�mero a la combinaci�n
'------------------------------------------------------------------------------*
Public Sub Add(oNum As Numero)
    Dim strKey  As String
  On Error GoTo Add_Error
  
    'strKey = oNum.ToString
    strKey = CStr(oNum.Valor)
    If oNum.Orden <= 0 Then
        oNum.Orden = mCol.Count + 1
    End If
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
' Prop�sito      : Elimina un n�mero de la combinaci�n
' Par�metros     : n�meno a eliminar
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
' Prop�sito      : Devuelve true si el n�mero est� contenido en la combinaci�n
' Par�metros     : Numero a comprobar
' Retorno        : True si la colecci�n contiene el n�mero
'------------------------------------------------------------------------------*
Public Function Contiene(Numero As Integer) As Boolean
    Dim strKey      As String       ' Clave de busqueda
    Dim prueba      As Variant      ' Resultado de la operaci�n
  On Error Resume Next              ' Desactivamos el error
  
    'strKey = Format(Numero, "00")  ' Convertimos el entero en string
                                    ' (clave de la colecci�n)
    strKey = CStr(Numero)           ' Convertimos el entero en string
                                    ' (clave de la colecci�n)
    prueba = mCol.Item(strKey)      ' Intentamos obtener el elemento de
                                    ' la colecci�n con clave strKey
    Contiene = (Err <> 5)           ' Si no lo encuentra aparece un error 5
                                    ' que en la expresi�n es verdadero
End Function

'------------------------------------------------------------------------------*
' Procedimiento  : EstaOrdenado
' Fecha          : 15/03/2014
' Prop�sito      : Devuelve true si los n�meros est�n en orden ascendente
' Par�metros     :
' Retorno        : True si los n�meros est�n en orden ascendente
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
' Prop�sito      : Devuelve una cadena de texto con los n�meros editados
' Par�metros     : OrdenAscendente
' Retorno        : Cadena de Texto con los n�meros separados por "-"
'------------------------------------------------------------------------------*
Public Function ToString(Optional OrdenAscendente As Boolean)
    Dim mNumeros()       As Integer
    Dim mNum             As Numero
    Dim i                As Integer
 On Error GoTo ToString_Error
    '
    '   Asigna el par�metro orden si no se pasa
    '
    If IsMissing(OrdenAscendente) Then
        OrdenAscendente = False
    End If
    '
    ' Inicializamos la variable
    '
    ToString = ""
    '
    '   Si no existen datos en la colecci�n
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
        '  Redimensionamos una matriz con el n�mero de elementos
        '  de la coleccion
        '
        ReDim mNumeros(mCol.Count - 1)
        '
        ' Inicializamos el indice
        '
        i = 0
        '
        ' Extraemos a la matriz los valores de los n�mero
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
    ' Quitamos el �ltimo gui�n
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
' Prop�sito      : Pasando el juego devuelve si el numero es valido o no
' Par�metros     : datJuego
'                  datApuesta, indicador para no validar cuantos numeros
' Retorno        : True si los numeros son v�lidos para el juego indicado
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
    '  Evaluamos si est� vacia
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
' Prop�sito      : Devuelve el mensaje de error si la combinaci�n no es v�lida
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

'------------------------------------------------------------------------------*
' Procedimiento  : GetArray
' Fecha          : mi., 24/jun/2020 19:00:41
' Prop�sito      : Devuelve un array con los n�meros de la colecci�n
' Retorno        : Mensaje de Error
'------------------------------------------------------------------------------*
Public Function GetArray() As Variant
    Dim mNumeros()      As Integer
    Dim i               As Integer
    
    If mCol.Count > 0 Then
        ReDim mNumeros(mCol.Count - 1)
    Else
        GetArray = mNumeros
    End If
    '
    ' Inicializamos el indice
    '
    i = 0
    '
    ' Extraemos a la matriz los valores de los n�mero
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
    '   devolvemos la matriz
    '
    GetArray = mNumeros
End Function

' *===========(EOF): Combinacion.cls


