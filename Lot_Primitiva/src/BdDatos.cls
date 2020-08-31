VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BdDatos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' *============================================================================*
' *
' *     Fichero    : BdDatos.cls
' *
' *     Autor      : Carlos Almela Baeza
' *     Creación   : mi., 01/nov/2006 17:16:00
' *     Modificado : mi., 24/jun/2020 19:26:09
' *     Versión    : 1.1
' *     Propósito  : Acceder a los datos de la hoja
' *                  Objeto que abstrae los rangos de datos de la lógica
' *============================================================================*
Option Explicit
Option Base 0
'--- Constantes ---------------------------------------------------------------*
Private Const hjDatos As String = "Resultados"  'Hoja de resultados
Private Const hjBoletos As String = "Boletos"   'Hoja de Boletos
Private Const hjApuesta As String = "Apuestas"  'Hoja de Apuestas

Private Const AdrssResultadosBL As String = "$A$2:$N${0}"
Private Const AdrssResultadosLP As String = "$A$2:$N${0}"
Private Const AdrssResultadosEU As String = "$A$2:$O${0}"
Private Const AdrssResultadosGP As String = "$A$2:$N${0}"

Private Const AdrssBoletosBL As String = "$A$2:$P${0}"
Private Const AdrssBoletosLP As String = "$A$2:$P${0}"
Private Const AdrssBoletosEU As String = "$A$2:$P${0}"
Private Const AdrssBoletosGP As String = "$A$2:$P${0}"

Private Const AdrssApuestasBL As String = "$A$2:$Y${0}"
Private Const AdrssApuestasLP As String = "$A$2:$Y${0}"
Private Const AdrssApuestasEU As String = "$A$2:$Y${0}"
Private Const AdrssApuestasGP As String = "$A$2:$Y${0}"

'
'--- Variables Privadas -------------------------------------------------------*

Private mRgSorteos As Range
Private mRgBoletos As Range
Private mRgApuestas As Range

Private mstrRangoResultados As String       ' Address tabla de Resultados
Private mstrRangoApuestas   As String       ' Address tabla de Apuestas
Private mstrRangoBoletos    As String       ' Address Tabla de Boletos
Private m_iUltimoRegistro   As Integer      ' Último registro
Private m_dtPrimerResultado As Date
Private m_dtUltimoResultado As Date
Private mInfo               As InfoSorteo   ' Información del concurso
'--- Mensajes -----------------------------------------------------------------*
'--- Errores ------------------------------------------------------------------*
'--- Propiedades --------------------------------------------------------------*


'---------------------------------------------------------------------------------------
' Procedimiento : RangoApuestas
' Creación      : 22-dic-2002 19:47
' Autor         : Carlos Almela Baeza
' Objeto        :
'---------------------------------------------------------------------------------------
'
Public Property Get RangoApuestas() As Range
    Set RangoApuestas = mRgApuestas
End Property


'---------------------------------------------------------------------------------------
' Procedimiento : RangoResultados
' Creación      :  22-dic-2002
' Autor         : Carlos Almela Baeza
' Objeto        :
'---------------------------------------------------------------------------------------
'
Public Property Get RangoResultados() As Range
    Set RangoResultados = mRgSorteos
End Property


'---------------------------------------------------------------------------------------
' Procedure : RangoBoletos
' DateTime  : 03/jun/2007 20:43
' Author    : Carlos Almela Baeza
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get RangoBoletos() As Range
    Set RangoBoletos = mRgBoletos
End Property



'---------------------------------------------------------------------------------------
' Procedure : PrimerResultado
' DateTime  : 01/04/2007 21:38
' Author    : Carlos Almela Baeza
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get PrimerResultado() As Date
    If m_dtPrimerResultado = 0 Then
        m_dtPrimerResultado = CDate(mRgSorteos.Offset(0, 2).Resize(1, 1).Value)
    End If
    PrimerResultado = m_dtPrimerResultado
End Property

' *============================================================================*
' *     Procedure  : UltimoRegistro
' *     Fichero    : BdDatos
' *     Autor      : Carlos Almela Baeza
' *     Creacion   : sáb, 14/01/2012 19:59
' *     Asunto     :
' *============================================================================*
'
Public Property Get UltimoRegistro() As Integer
    If m_iUltimoRegistro = 0 Then
        m_iUltimoRegistro = mRgSorteos.Offset(mRgSorteos.Rows.Count - 1, 0).Resize(1, 1).Value
                                                ' Buscamos la última celda de la primera
                                                ' columna y lo convertimos en fecha
    End If
    UltimoRegistro = m_iUltimoRegistro
End Property

'---------------------------------------------------------------------------------------
' Procedure : UltimoResultado
' DateTime  : 01/04/2007 21:38
' Author    : Carlos Almela Baeza
' Purpose   : Obtiene la fecha del último resultado de la base de datos
'---------------------------------------------------------------------------------------
'
Public Property Get UltimoResultado() As Date
    If m_dtUltimoResultado = 0 Then             ' Si no se tiene la fecha definida se
                                                ' calcula
        m_dtUltimoResultado = CDate(mRgSorteos.Offset(mRgSorteos.Rows.Count - 1, 2).Resize(1, 1).Value)
                                                ' Buscamos la última celda de la primera
                                                ' columna y lo convertimos en fecha
    End If
    UltimoResultado = m_dtUltimoResultado       ' Devolvemos la fecha almacenada
End Property

'--- Métodos Privados ---------------------------------------------------------*
'
'
Private Sub Class_Initialize()
    Dim mStr As String
    Dim mRow As Integer
    Dim mRg As Range
    
  On Error GoTo Class_Initialize_Error
    Select Case JUEGO_DEFECTO
        Case Bonoloto:
            '
            '   Sorteos
            '
            Set mRg = Worksheets(hjDatos).Range("A:A")
            mRow = Application.WorksheetFunction.Subtotal(2, mRg) + 1
            mStr = Replace(AdrssResultadosBL, "{0}", CStr(mRow))
            Set mRgSorteos = Worksheets(hjDatos).Range(mStr)
            '
            '   Boletos
            '
            Set mRg = Worksheets(hjBoletos).Range("A:A")
            mRow = Application.WorksheetFunction.Subtotal(2, mRg) + 1
            mStr = Replace(AdrssBoletosBL, "{0}", CStr(mRow))
            Set mRgBoletos = Worksheets(hjBoletos).Range(mStr)
            '
            '   Apuestas
            '
            Set mRg = Worksheets(hjApuesta).Range("A:A")
            mRow = Application.WorksheetFunction.Subtotal(2, mRg) + 1
            mStr = Replace(AdrssApuestasBL, "{0}", CStr(mRow))
            Set mRgApuestas = Worksheets(hjApuesta).Range(mStr)
            
        Case LoteriaPrimitiva:
            '
            '   Sorteos
            '
            Set mRg = Worksheets(hjDatos).Range("A:A")
            mRow = Application.WorksheetFunction.Subtotal(2, mRg) + 1
            mStr = Replace(AdrssResultadosLP, "{0}", CStr(mRow))
            Set mRgSorteos = Worksheets(hjDatos).Range(mStr)
            '
            '   Boletos
            '
            Set mRg = Worksheets(hjBoletos).Range("A:A")
            mRow = Application.WorksheetFunction.Subtotal(2, mRg) + 1
            mStr = Replace(AdrssBoletosLP, "{0}", CStr(mRow))
            Set mRgBoletos = Worksheets(hjBoletos).Range(mStr)
            '
            '   Apuestas
            '
            Set mRg = Worksheets(hjApuesta).Range("A:A")
            mRow = Application.WorksheetFunction.Subtotal(2, mRg) + 1
            mStr = Replace(AdrssApuestasLP, "{0}", CStr(mRow))
            Set mRgApuestas = Worksheets(hjApuesta).Range(mStr)
            
        Case gordoPrimitiva:
            '
            '   Sorteos
            '
            Set mRg = Worksheets(hjDatos).Range("A:A")
            mRow = Application.WorksheetFunction.Subtotal(2, mRg) + 1
            mStr = Replace(AdrssResultadosGP, "{0}", CStr(mRow))
            Set mRgSorteos = Worksheets(hjDatos).Range(mStr)
            '
            '   Boletos
            '
            Set mRg = Worksheets(hjBoletos).Range("A:A")
            mRow = Application.WorksheetFunction.Subtotal(2, mRg) + 1
            mStr = Replace(AdrssBoletosGP, "{0}", CStr(mRow))
            Set mRgBoletos = Worksheets(hjBoletos).Range(mStr)
            '
            '   Apuestas
            '
            Set mRg = Worksheets(hjApuesta).Range("A:A")
            mRow = Application.WorksheetFunction.Subtotal(2, mRg) + 1
            mStr = Replace(AdrssApuestasGP, "{0}", CStr(mRow))
            Set mRgApuestas = Worksheets(hjApuesta).Range(mStr)
            
        Case Euromillones:
            '
            '   Sorteos
            '
            Set mRg = Worksheets(hjDatos).Range("A:A")
            mRow = Application.WorksheetFunction.Subtotal(2, mRg) + 1
            mStr = Replace(AdrssResultadosEU, "{0}", CStr(mRow))
            Set mRgSorteos = Worksheets(hjDatos).Range(mStr)
            '
            '   Boletos
            '
            Set mRg = Worksheets(hjBoletos).Range("A:A")
            mRow = Application.WorksheetFunction.Subtotal(2, mRg) + 1
            mStr = Replace(AdrssResultadosEU, "{0}", CStr(mRow))
            Set mRgBoletos = Worksheets(hjBoletos).Range(mStr)
            '
            '   Apuestas
            '
            Set mRg = Worksheets(hjApuesta).Range("A:A")
            mRow = Application.WorksheetFunction.Subtotal(2, mRg) + 1
            mStr = Replace(AdrssApuestasEU, "{0}", CStr(mRow))
            Set mRgApuestas = Worksheets(hjApuesta).Range(mStr)
            
    End Select
  On Error GoTo 0
Class_Initialize_CleanExit:
    Exit Sub
Class_Initialize_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    Call HandleException(ErrNumber, ErrDescription, "BdDatos.Class_Initialize", ErrSource)
    Err.Raise ErrNumber, ErrSource, ErrDescription
End Sub



Private Sub Class_Terminate()
    Set mRgSorteos = Nothing
    Set mRgBoletos = Nothing
    Set mRgApuestas = Nothing
End Sub




'---------------------------------------------------------------------------------------
' Procedure : GetRowFechaMenorOIgual
' Author    : Charly
' Date      : 16/06/2020
' Purpose   : Devuelve el registro de un sorteo menor o igual a una fecha dada
'---------------------------------------------------------------------------------------
'
Private Function GetRowFechaMenorOIgual(datFecha As Date) As Integer
    Dim rg_res          As Range        'Rango de búsqueda
    Dim res             As Range
    Dim mFind           As Boolean
    Dim mRow            As Integer
  On Error GoTo GetRowFechaMenorOIgual_Error
    '
    '   Definimos un objeto para el control de sorteos
    '
    Set mInfo = New InfoSorteo
    '
    '   Establecemos el rango de búsqueda
    '
    Set rg_res = mRgSorteos
    '
    '   Inicializamos el indicador de fecha encontrada
    '
    mFind = False
    '
    '   Bucle de busqueda hasta que se encuentre
    '
    While (Not mFind) And (datFecha >= PrimerResultado)
        '
        '   Buscamos el dato en el rango
        '
        Set res = rg_res.Find(datFecha, lookat:=xlWhole, LookIn:=xlValues, MatchCase:=True)
        '
        '   Si no hemos encontrado la fecha
        '
        If res Is Nothing Then
            '
            '   Buscamos la fecha anterior al sorteo
            '
            datFecha = mInfo.GetAnteriorSorteo(datFecha)
            
        Else
            '
            '   Establecemos la fila encontrada
            '
            GetRowFechaMenorOIgual = res.Row - 2
            '
            '   establecemos el indicador de encontrado a true
            '
            mFind = True
        End If
    Wend
    
  On Error GoTo 0
GetRowFechaMenorOIgual_CleanExit:
    Exit Function
GetRowFechaMenorOIgual_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    Call HandleException(ErrNumber, ErrDescription, "BdDatos.GetRowFechaMenorOIgual", ErrSource)
    Err.Raise ErrNumber, ErrSource, ErrDescription
End Function




'---------------------------------------------------------------------------------------
' Procedure : GetRowFechaMayorOIgual
' Author    : Charly
' Date      : 16/06/2020
' Purpose   : Devuelve el registro de un sorteo mayor o igual a una fecha dada
'---------------------------------------------------------------------------------------
'
Private Function GetRowFechaMayorOIgual(datFecha As Date) As Integer
    Dim rg_res          As Range        'Rango de búsqueda
    Dim res             As Range
    Dim mFind           As Boolean
    Dim mRow            As Integer
  On Error GoTo GetRowFechaMayorOIgual_Error
    '
    '   Definimos un objeto para el control de sorteos
    '
    Set mInfo = New InfoSorteo
    '
    '   Establecemos el rango de búsqueda
    '
    Set rg_res = mRgSorteos
    '
    '   Inicializamos el indicador de fecha encontrada
    '
    mFind = False
    '
    '   Bucle de busqueda hasta que se encuentre
    '
    While (Not mFind) And (datFecha <= UltimoResultado)
        '
        '   Buscamos el dato en el rango
        '
        Set res = rg_res.Find(datFecha, lookat:=xlWhole, LookIn:=xlValues, MatchCase:=True)
        '
        '   Si no hemos encontrado la fecha
        '
        If res Is Nothing Then
            '
            '   Buscamos la fecha anterior al sorteo
            '
            datFecha = mInfo.GetProximoSorteo(datFecha)
        Else
            '
            '   Establecemos la fila encontrada
            '
            GetRowFechaMayorOIgual = res.Row - 2
            '
            '   establecemos el indicador de encontrado a true
            '
            mFind = True
        End If
    Wend
    
  On Error GoTo 0
GetRowFechaMayorOIgual_CleanExit:
    Exit Function
GetRowFechaMayorOIgual_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    Call HandleException(ErrNumber, ErrDescription, "BdDatos.GetRowFechaMayorOIgual", ErrSource)
    Err.Raise ErrNumber, ErrSource, ErrDescription
End Function


'------------------------------------------------------------------------------*
' Función        : GetTupla
' Fecha          : vi., 03/jul/2020 18:24:19
' Propósito      : Obtiene una tupla de un número asociada a una celda
' Parámetros     : Rango de la celda
' Retorno        : tuplas con la aparición
'------------------------------------------------------------------------------*
'
Private Function GetTupla(vNewValue As Range) As TuplaAparicion
    Dim mTupla As TuplaAparicion
    Dim i As Integer
  
  On Error GoTo GetTupla_Error
    '
    '   Creamos la Tupla
    '
    Set mTupla = New TuplaAparicion
    '
    '   Calculamos el desplazamiento del registro
    '
    i = vNewValue.Row - 2
    '
    '   Cargamos la tupla
    '
    mTupla.NumeroRegistro = CInt(mRgSorteos.Offset(i, 0).Resize(1, 1).Value)
    mTupla.FechaAparicion = CDate(mRgSorteos.Offset(i, 2).Resize(1, 1).Value)
    Select Case JUEGO_DEFECTO
        Case Bonoloto, LoteriaPrimitiva:
            If vNewValue.Column < 13 Then
                mTupla.OrdenAparicion = vNewValue.Column - 5
            Else
                '   columna del reintegro
                '
                mTupla.OrdenAparicion = 1
            End If
        
        Case gordoPrimitiva
            If vNewValue.Column < 12 Then
                mTupla.OrdenAparicion = vNewValue.Column - 6
            Else
                '   columna del reintegro
                '
                mTupla.OrdenAparicion = 1
            End If
        
        Case Euromillones:
            If vNewValue.Column < 12 Then
                mTupla.OrdenAparicion = vNewValue.Column - 6
            Else
                '   columna de las estrellas
                '
                mTupla.OrdenAparicion = vNewValue.Column - 11
            End If
    End Select
    '
    '   Devolvemos la tupla
    '
    Set GetTupla = mTupla

  On Error GoTo 0
    Exit Function
GetTupla_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    Call HandleException(ErrNumber, ErrDescription, "BdDatos.GetTupla")
    Err.Raise ErrNumber, ErrSource, ErrDescription
End Function



'--- Métodos Públicos ---------------------------------------------------------*






' *============================================================================*
' *     Procedure  : GetRegistroFecha
' *     Fichero    : BdDatos
' *     Autor      : Carlos Almela Baeza
' *     Creacion   : sáb, 14/01/2012 20:37
' *     Asunto     :
' *============================================================================*
'
Public Function GetRegistroFecha(datFecha As Date) As Integer
    Dim iResult             As Integer
    Dim m_rgRes             As Range
    Dim m_rgCelda           As Range
    Dim m_vRes              As Variant
    Dim m_rowIni            As Integer
    Dim m_bFound            As Boolean
  On Error GoTo GetRegistroFecha_Error
    '
    '   Si la fecha es mayor que la última se simula el Numero de registro
    '
    If datFecha > Me.UltimoResultado Then
        GetRegistroFecha = 0
        'GetRegistroFecha = GetSimulacionRegistro(datFecha)
        Exit Function
    End If
    '
    '   Establecer el Rango de resultados
    '
    Set m_rgRes = mRgSorteos
    '
    '   Buscamos el registro menor o igual a la fecha
    '
    m_bFound = False
    While Not (m_bFound)
        '
        '   busca en el rango la fecha
        '
        Set m_vRes = m_rgRes.Find(datFecha, lookat:=xlWhole, LookIn:=xlValues, MatchCase:=True)
        '
        '   Si la encuentra sale del bucle
        '
        If Not (m_vRes Is Nothing) Then
            m_rowIni = m_vRes.Row
            m_bFound = True
        Else
            '
            '   Si no la encuentra le resta un dia
            '
            datFecha = datFecha - 1
        End If
    Wend
    '
    '   Localiza la celda A de la fila encontrada, donde esta el Numero de registro
    '
    Set m_rgCelda = m_rgRes.Offset(m_rowIni - 2, 0).Resize(1, 1)

    iResult = m_rgCelda.Value

    GetRegistroFecha = iResult

GetRegistroFecha_CleanExit:
   On Error GoTo 0
    Exit Function
GetRegistroFecha_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    Call HandleException(ErrNumber, ErrDescription, "BdDatos.GetRegistroFecha")
    Err.Raise ErrNumber, ErrSource, ErrDescription
End Function

Private Function GetSimulacionRegistro(datFecha As Date) As Integer
    Dim m_iReg As Integer
    Dim m_dtFechaMin    As Date
    Dim m_dtFecha As Date
    Dim m_DiaSem As Integer
    
    m_dtFechaMin = Me.UltimoResultado
    m_iReg = Me.UltimoRegistro
    For m_dtFecha = m_dtFechaMin To datFecha Step 1
        '
        '   controla solo bono y primi
        '
        m_DiaSem = Application.WorksheetFunction.Weekday(m_dtFecha, 2)
        If m_DiaSem <> 7 Then
            m_iReg = m_iReg + 1
        End If
    Next m_dtFecha
    GetSimulacionRegistro = m_iReg - 1
End Function

Public Function GetSimulacionFecha(datRegistro As Integer) As Date
    Dim m_iReg          As Integer
    Dim m_dtFecha       As Date
    Dim m_DiaSem        As Integer
    
    m_dtFecha = Me.UltimoResultado
    m_iReg = Me.UltimoRegistro
    
    While m_iReg < datRegistro
        m_iReg = m_iReg + 1
        m_dtFecha = m_dtFecha + 1
        m_DiaSem = Application.WorksheetFunction.Weekday(m_dtFecha, 2)
        If m_DiaSem = 7 Then
            m_dtFecha = m_dtFecha + 1
        End If
    Wend
    GetSimulacionFecha = m_dtFecha
End Function


'------------------------------------------------------------------------------*
' Función        : GetBoletoByFecha
' Fecha          : ju., 02/jul/2020 19:16:33
' Propósito      : Obtiene la fila del boleto de una fecha
' Parámetros     : Fecha del boleto
' Retorno        : Rango de excel con los datos
'------------------------------------------------------------------------------*
'
Public Function GetBoletoByFecha(vNewValue As Date) As Range
    Dim mRg As Range            ' Rango de trabajo
    Dim mRes As Range           ' Rango resultado
    Dim mRow As Integer         ' Fila del resultado
    
  On Error GoTo GetBoletoByFecha_Error
    '
    '   Establecemos como rango de trabajo la columna de fechas ($C$C)
    '
    Set mRg = mRgBoletos.Offset(0, 2).Resize(mRgBoletos.Rows.Count, 1)
    '
    '   Se busca la fecha de entrada
    '
    Set mRes = mRg.Find(vNewValue, _
                       lookat:=xlWhole, _
                       LookIn:=xlValues, _
                       MatchCase:=True)
        
    If Not (mRes Is Nothing) Then
        '
        '   A la fila donde localiza la fecha final del período
        '   Se le restan el registro de la cabecera
        '
        mRow = mRes.Row - 2
        '
        '   La primera celda es la fecha de los resultados
        '
        Set mRes = mRgBoletos.Offset(mRow, 0).Resize(1, mRgBoletos.Columns.Count)
    End If
    '
    '   Devolvemos la fila
    '
    Set GetBoletoByFecha = mRes
    
    
GetBoletoByFecha_CleanExit:
  On Error GoTo 0
    Exit Function
GetBoletoByFecha_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    Call HandleException(ErrNumber, ErrDescription, "BdDatos.GetBoletoByFecha")
    Err.Raise ErrNumber, ErrSource, ErrDescription
End Function





'------------------------------------------------------------------------------*
' Función        : GetBoletoInFechas
' Fecha          : ju., 02/jul/2020 19:23:46
' Propósito      : Obtiene las filas de los boletos entre dos fechas
' Parámetros     : Periodo de Fechas
' Retorno        : Rango de excel con los datos
'------------------------------------------------------------------------------*
'
Public Function GetBoletoInFechas(vNewValue As Periodo) As Range
    Dim mRg As Range            ' Rango de trabajo
    Dim mRes As Range           ' Rango resultado
    Dim mRowIni As Integer      ' Fila de la fecha inicial
    Dim mRowFin As Integer      ' Fila de la fecha final
    
  On Error GoTo GetBoletoInFechas_Error
    '
    '   Inicializamos filas
    '
    mRowIni = -1: mRowFin = -1
    '
    '   Establecemos como rango de trabajo la columna de fechas ($C$C)
    '
    Set mRg = mRgBoletos.Offset(0, 2).Resize(mRgBoletos.Rows.Count, 1)
    '
    '   Se busca la fecha inicial de entrada
    '
    Set mRes = mRg.Find(vNewValue.FechaInicial, _
                       lookat:=xlWhole, _
                       LookIn:=xlValues, _
                       MatchCase:=True)
        
    If Not (mRes Is Nothing) Then
        mRowIni = mRes.Row - 2
    End If
    
    Set mRes = mRg.Find(vNewValue.FechaFinal, _
                       lookat:=xlWhole, _
                       LookIn:=xlValues, _
                       MatchCase:=True)
        
    If Not (mRes Is Nothing) Then
        mRowFin = mRes.Row - 2
    End If
    
    If mRowIni >= 0 And mRowFin >= 0 Then
        '
        '   La primera celda es la fecha de los resultados
        '
        Set mRes = mRgBoletos.Offset(mRowIni, 0).Resize((mRowFin - mRowIni) + 1, mRgBoletos.Columns.Count)
    Else
        Set mRes = Nothing
    End If
    '
    '   Devolvemos la fila
    '
    Set GetBoletoInFechas = mRes

GetBoletoInFechas_CleanExit:
  On Error GoTo 0
    Exit Function
GetBoletoInFechas_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    Call HandleException(ErrNumber, ErrDescription, "BdDatos.GetBoletoInFechas")
    Err.Raise ErrNumber, ErrSource, ErrDescription
End Function



'------------------------------------------------------------------------------*
' Función        : GetBoletoById
' Fecha          : ju., 02/jul/2020 19:39:43
' Propósito      : Obtiene el boleto con el Id indicado
' Parámetros     : Identificador del Boleto
' Retorno        : Rango de excel con los datos
'------------------------------------------------------------------------------*
'
Public Function GetBoletoById(vNewValue As Integer) As Range
    Dim mRg As Range            ' Rango de trabajo
    Dim mRes As Range           ' Rango resultado
    Dim mRow As Integer         ' Fila del resultado
    
  On Error GoTo GetBoletoById_Error

    '
    '   Establecemos como rango de trabajo la columna de fechas ($A$A)
    '
    Set mRg = mRgBoletos.Offset(0, 0).Resize(mRgBoletos.Rows.Count, 1)
    '
    '   Se busca la fecha de entrada
    '
    Set mRes = mRg.Find(vNewValue, _
                       lookat:=xlWhole, _
                       LookIn:=xlValues, _
                       MatchCase:=True)
        
    If Not (mRes Is Nothing) Then
        '
        '   A la fila donde localiza la fecha final del período
        '   Se le restan el registro de la cabecera
        '
        mRow = mRes.Row - 2
        '
        '   La primera celda es la fecha de los resultados
        '
        Set mRes = mRgBoletos.Offset(mRow, 0).Resize(1, mRgBoletos.Columns.Count)
    End If
    '
    '   Devolvemos la fila
    '
    Set GetBoletoById = mRes

GetBoletoById_CleanExit:
  On Error GoTo 0
    Exit Function
GetBoletoById_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    Call HandleException(ErrNumber, ErrDescription, "BdDatos.GetBoletoById")
    Err.Raise ErrNumber, ErrSource, ErrDescription
End Function



'------------------------------------------------------------------------------*
' Función        : GetApuestaByFecha
' Fecha          : vi., 03/jul/2020 08:34:35
' Propósito      : Obtiene la apuesta de una fecha
' Parámetros     : Fecha de la apuesta
' Retorno        : Rango de excel con los datos
'------------------------------------------------------------------------------*
'
Public Function GetApuestaByFecha(vNewValue As Date) As Range
    Dim mRg As Range            ' Rango de trabajo
    Dim mRes As Range           ' Rango resultado
    Dim mRow As Integer         ' Fila del resultado
    
  On Error GoTo GetApuestaByFecha_Error
    '
    '   Establecemos como rango de trabajo la columna de fechas ($D$D)
    '
    Set mRg = mRgApuestas.Offset(0, 3).Resize(mRgApuestas.Rows.Count, 1)
    '
    '   Se busca la fecha de entrada
    '
    Set mRes = mRg.Find(vNewValue, _
                       lookat:=xlWhole, _
                       LookIn:=xlValues, _
                       MatchCase:=True)
        
    If Not (mRes Is Nothing) Then
        '
        '   #TODO: Se pueden tener varios registro de una fecha
        '
        '
        '   A la fila donde localiza la fecha final del período
        '   Se le restan el registro de la cabecera
        '
        mRow = mRes.Row - 2
        '
        '   La primera celda es la fecha de los resultados
        '
        Set mRes = mRgApuestas.Offset(mRow, 0).Resize(1, mRgApuestas.Columns.Count)
    End If
    '
    '   Devolvemos la fila
    '
    Set GetApuestaByFecha = mRes
    
    
GetApuestaByFecha_CleanExit:
  On Error GoTo 0
    Exit Function
GetApuestaByFecha_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    Call HandleException(ErrNumber, ErrDescription, "BdDatos.GetApuestaByFecha")
    Err.Raise ErrNumber, ErrSource, ErrDescription
End Function



'------------------------------------------------------------------------------*
' Función        : GetApuestaInFechas
' Fecha          : vi., 03/jul/2020 08:36:49
' Propósito      : Obtiene las apuestas de un periodo de fechas
' Parámetros     : Periodo de busqueda
' Retorno        : Rango de excel con los datos
'------------------------------------------------------------------------------*
'
Public Function GetApuestaInFechas(vNewValue As Periodo) As Range
    Dim mRg As Range            ' Rango de trabajo
    Dim mRes As Range           ' Rango resultado
    Dim mRowIni As Integer      ' Fila de la fecha inicial
    Dim mRowFin As Integer      ' Fila de la fecha final
    Dim mFind                   ' As Boolean
    Dim mTmpRow As Integer      ' Current Fila de la fecha final
    
  On Error GoTo GetApuestaInFechas_Error
    '
    '   Inicializamos filas
    '
    mRowIni = 0: mRowFin = 0
    '
    '   Establecemos como rango de trabajo la columna de fechas ($D$D)
    '
    Set mRg = mRgApuestas.Offset(0, 3).Resize(mRgApuestas.Rows.Count, 1)
    '
    '   Se busca la fecha inicial de entrada
    '
    Set mRes = mRg.Find(vNewValue.FechaInicial, _
                       lookat:=xlWhole, _
                       LookIn:=xlValues, _
                       MatchCase:=True)
            
    If Not (mRes Is Nothing) Then
        mRowIni = mRes.Row - 2
    End If
    
    Set mRes = mRg.Find(vNewValue.FechaFinal, _
                       lookat:=xlWhole, _
                       LookIn:=xlValues, _
                       MatchCase:=True)
        
    If Not (mRes Is Nothing) Then
        '
        '   Guardamos el primer registro
        '
        mTmpRow = mRes.Row
        '
        '   Inicializamos el indicador de bucle
        '
        mFind = False
        '
        '   Buscar mas registros, mientras sea la misma fecha
        '
        While Not mFind
            '
            '   Buscamos siguiente registro
            '
            Set mRes = mRg.FindNext(mRes)
            '
            '   Si no obtenemos rango salimos
            '
            If mRes Is Nothing Then
                mFind = True
            Else
                '
                '   Si la fila es superior a la censada
                '
                If mRes.Row > mTmpRow Then
                    '
                    '   guardamos la última fila
                    '
                    mTmpRow = mRes.Row
                Else
                    '
                    '  Si es menor o igual a la primera salimos
                    '
                    mFind = True
                End If
            End If
        Wend
        '
        '   Calculamos la fila del rango eliminando desplazamientos
        '
        mRowFin = mTmpRow - 2
    End If
    
    If mRowIni > 0 And mRowFin > 0 Then
        '
        '   La primera celda es la fecha de los resultados
        '
        Set mRes = mRgApuestas.Offset(mRowIni, 0).Resize((mRowFin - mRowIni) + 1, mRgApuestas.Columns.Count)
    Else
        Set mRes = Nothing
    End If
    '
    '   Devolvemos la fila
    '
    Set GetApuestaInFechas = mRes

GetApuestaInFechas_CleanExit:
  On Error GoTo 0
    Exit Function
GetApuestaInFechas_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    Call HandleException(ErrNumber, ErrDescription, "BdDatos.GetApuestaInFechas")
    Err.Raise ErrNumber, ErrSource, ErrDescription
End Function



'------------------------------------------------------------------------------*
' Función        : GetApuestaByBoleto
' Fecha          : vi., 03/jul/2020 09:45:25
' Propósito      : Obtiene las apuestas de un boleto
' Parámetros     : Id del Boleto
' Retorno        : Rango de excel con los datos
'------------------------------------------------------------------------------*
'
Public Function GetApuestaByBoleto(vNewValue As Variant) As Range
    Dim mRg As Range            ' Rango de trabajo
    Dim mRes As Range           ' Rango resultado
    Dim mRowI As Integer        ' Fila del primer registro
    Dim mRowF As Integer        ' Fila del ultimo registro
    Dim bFound As Boolean
  On Error GoTo GetApuestaByBoleto_Error
    '
    '   Establecemos como rango de trabajo la columna de IdBoletos ($B$B)
    '
    Set mRg = mRgApuestas.Offset(0, 1).Resize(mRgApuestas.Rows.Count, 1)
    '
    '   Se busca el Id del Boleto  'xlPart
    '
    Set mRes = mRg.Find(vNewValue, _
                       lookat:=xlPart, _
                       LookIn:=xlValues, _
                       MatchCase:=True)
        
    If Not (mRes Is Nothing) Then
        '
        '   Primera fila donde se encuentra el boleto
        '
        mRowI = mRes.Row
        mRowF = mRes.Row
        bFound = False
        '
        '   Bucle de busqueda siguiente
        '
        Do
            '
            '   Buscamos el siguiente registro
            '
            Set mRes = mRg.FindNext(mRes)
            '
            '   Si la fila es la misma salimos del bucle
            '
            If mRes.Row <= mRowF Then
                bFound = True
            Else
                If mRes.Value = vNewValue Then
                    mRowF = mRes.Row
                End If
            End If
        Loop While Not bFound
        '
        '   Calculamos el rango de desplazamiento
        '
        mRowI = mRowI - 2
        mRowF = mRowF - 2
        Set mRes = mRgApuestas.Offset(mRowI, 0).Resize((mRowF - mRowI) + 1, mRgBoletos.Columns.Count)
    End If
    '
    '   Devolvemos la fila
    '
    Set GetApuestaByBoleto = mRes

GetApuestaByBoleto_CleanExit:
  On Error GoTo 0
    Exit Function
GetApuestaByBoleto_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    Call HandleException(ErrNumber, ErrDescription, "BdDatos.GetApuestaByBoleto")
    Err.Raise ErrNumber, ErrSource, ErrDescription
End Function



'------------------------------------------------------------------------------*
' Función        : GetSorteoByFecha
' Fecha          : vi., 03/jul/2020 09:35:23
' Propósito      : Obtiene el sorteo de una fecha dada
' Parámetros     : Fecha de sorteo
' Retorno        : Rango de excel con los datos
'------------------------------------------------------------------------------*
'
Public Function GetSorteoByFecha(vNewValue As Date) As Range
    Dim mRg As Range            ' Rango de trabajo
    Dim mRes As Range           ' Rango resultado
    Dim mRow As Integer         ' Fila del resultado
    
  On Error GoTo GetSorteoByFecha_Error
    '
    '   Establecemos como rango de trabajo la columna de fechas ($C$C)
    '
    Set mRg = mRgSorteos.Offset(0, 2).Resize(mRgSorteos.Rows.Count, 1)
    '
    '   Se busca la fecha de entrada
    '
    Set mRes = mRg.Find(vNewValue, _
                       lookat:=xlWhole, _
                       LookIn:=xlValues, _
                       MatchCase:=True)
        
    If Not (mRes Is Nothing) Then
        '
        '   A la fila donde localiza la fecha final del período
        '   Se le restan el registro de la cabecera
        '
        mRow = mRes.Row - 2
        '
        '   La primera celda es la fecha de los resultados
        '
        Set mRes = mRgSorteos.Offset(mRow, 0).Resize(1, mRgSorteos.Columns.Count)
    End If
    '
    '   Devolvemos la fila
    '
    Set GetSorteoByFecha = mRes
    
    
GetSorteoByFecha_CleanExit:
  On Error GoTo 0
    Exit Function
GetSorteoByFecha_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    Call HandleException(ErrNumber, ErrDescription, "BdDatos.GetSorteoByFecha")
    Err.Raise ErrNumber, ErrSource, ErrDescription
End Function



'------------------------------------------------------------------------------*
' Función        : GetSorteoById
' Fecha          : vi., 03/jul/2020 09:39:37
' Propósito      : Obtiene el sorteo de un Id
' Parámetros     : Id del Sorteo
' Retorno        : Rango de excel con los datos
'------------------------------------------------------------------------------*
'
Public Function GetSorteoById(vNewValue As Integer) As Range
    Dim mRg As Range            ' Rango de trabajo
    Dim mRes As Range           ' Rango resultado
    Dim mRow As Integer         ' Fila del resultado
    
  On Error GoTo GetSorteoById_Error

    '
    '   Establecemos como rango de trabajo la columna de fechas ($A$A)
    '
    Set mRg = mRgSorteos.Offset(0, 0).Resize(mRgSorteos.Rows.Count, 1)
    '
    '   Se busca la fecha de entrada
    '
    Set mRes = mRg.Find(vNewValue, _
                       lookat:=xlWhole, _
                       LookIn:=xlValues, _
                       MatchCase:=True)
        
    If Not (mRes Is Nothing) Then
        '
        '   A la fila donde localiza la fecha final del período
        '   Se le restan el registro de la cabecera
        '
        mRow = mRes.Row - 2
        '
        '   La primera celda es la fecha de los resultados
        '
        Set mRes = mRgSorteos.Offset(mRow, 0).Resize(1, mRgSorteos.Columns.Count)
    End If
    '
    '   Devolvemos la fila
    '
    Set GetSorteoById = mRes

GetSorteoById_CleanExit:
  On Error GoTo 0
    Exit Function
GetSorteoById_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    Call HandleException(ErrNumber, ErrDescription, "BdDatos.GetSorteoById")
    Err.Raise ErrNumber, ErrSource, ErrDescription
End Function




'------------------------------------------------------------------------------*
' Función        : GetSorteosInFechas
' Fecha          : vi., 03/jul/2020 09:42:04
' Propósito      : Obtiene los sorteos de un periodo de fechas
' Parámetros     : periodo de fechas
' Retorno        : Rango de excel con los datos
'------------------------------------------------------------------------------*
'
Public Function GetSorteosInFechas(vNewValue As Periodo) As Range
    Dim mRg As Range            ' Rango de trabajo
    Dim mRes As Range           ' Rango resultado
    Dim mRowIni As Integer      ' Fila de la fecha inicial
    Dim mRowFin As Integer      ' Fila de la fecha final
    
  On Error GoTo GetSorteosInFechas_Error
    '
    '   Inicializamos filas
    '
    mRowIni = 0: mRowFin = 0
    '
    '   Establecemos como rango de trabajo la columna de fechas ($C$C)
    '
    Set mRg = mRgSorteos.Offset(0, 2).Resize(mRgSorteos.Rows.Count, 1)
    '
    '   Se busca la fecha inicial de entrada
    '
    Set mRes = mRg.Find(vNewValue.FechaInicial, _
                       lookat:=xlWhole, _
                       LookIn:=xlValues, _
                       MatchCase:=True)
    '
    '   #TODO: revisar registros duplicados
    '
    If Not (mRes Is Nothing) Then
        mRowIni = mRes.Row - 2
    Else
        mRowIni = GetRowFechaMayorOIgual(vNewValue.FechaInicial)
    End If
    
    Set mRes = mRg.Find(vNewValue.FechaFinal, _
                       lookat:=xlWhole, _
                       LookIn:=xlValues, _
                       MatchCase:=True)
        
    If Not (mRes Is Nothing) Then
        mRowFin = mRes.Row - 2
    Else
        mRowFin = GetRowFechaMenorOIgual(vNewValue.FechaFinal)
    End If
    
    If mRowIni > 0 And mRowFin > 0 Then
        '
        '   La primera celda es la fecha de los resultados
        '
        Set mRes = mRgSorteos.Offset(mRowIni, 0).Resize((mRowFin - mRowIni) + 1, mRgSorteos.Columns.Count)
    Else
        Set mRes = Nothing
    End If
    '
    '   Devolvemos la fila
    '
    Set GetSorteosInFechas = mRes

GetSorteosInFechas_CleanExit:
  On Error GoTo 0
    Exit Function
GetSorteosInFechas_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    Call HandleException(ErrNumber, ErrDescription, "BdDatos.GetSorteosInFechas")
    Err.Raise ErrNumber, ErrSource, ErrDescription
End Function

'------------------------------------------------------------------------------*
' Función        : GetAparicionesNumero
' Fecha          : vi., 03/jul/2020 18:24:19
' Propósito      : Obtiene todas las apariciones de un número para un sorteo
' Parámetros     : Numero del Concurso
' Retorno        : Colección de tuplas con la aparición
'------------------------------------------------------------------------------*
'
Public Function GetAparicionesNumero(vNewValue As Integer) As Variant
    Dim mRg As Range            ' Rango de trabajo
    Dim mRes As Range           ' Rango resultado
    Dim mRowI As Integer        ' Fila del primer registro
    Dim mRowF As Integer        ' Fila del ultimo registro
    Dim bFound As Boolean       ' Indicador de registro encontrado
    Dim mCol As Collection      ' Colección de apariciones
    Dim mTup As TuplaAparicion  ' Tupla de aparición del número
  
  On Error GoTo GetAparicionesNumero_Error
    '
    '   Atendiendo al Juego establecemos el área de los números
    '   Establecemos como rango de trabajo la columna de los numeros ($F$L)
    '
    Select Case JUEGO_DEFECTO
        Case Bonoloto, LoteriaPrimitiva:
            ' ($F:$L)
            Set mRg = mRgSorteos.Offset(0, 5).Resize(mRgSorteos.Rows.Count, 7)
        
        Case gordoPrimitiva, Euromillones:
            ' ($G:$K)
            Set mRg = mRgSorteos.Offset(0, 6).Resize(mRgSorteos.Rows.Count, 5)
    End Select
    '
    '   Creamos la colección
    '
    Set mCol = New Collection
    '
    '   Se busca el Numero
    '
    Set mRes = mRg.Find(vNewValue, _
                       lookat:=xlWhole, _
                       LookIn:=xlValues, _
                       MatchCase:=True)
               
    If Not (mRes Is Nothing) Then
        '
        '   Primera fila donde se encuentra el boleto
        '
        mRowI = mRes.Row
        mRowF = mRes.Row
        bFound = False
        '
        '   Creamos la primera tupla y guardamos
        '
        Set mTup = GetTupla(mRes)
        mCol.Add mTup
        '
        '   Bucle de busqueda siguiente
        '
        Do
            Set mRes = mRg.FindNext(mRes)
            '
            '   Si la fila es la misma o menor salimos del bucle
            '
            If mRes.Row <= mRowF Then
                bFound = True
            Else
                mRowF = mRes.Row
                
                Set mTup = GetTupla(mRes)
                mCol.Add mTup
            End If
        Loop While Not bFound
    End If
    '
    '   Devolvemos la colección de tuplas
    '
    Set GetAparicionesNumero = mCol

GetAparicionesNumero_CleanExit:
  On Error GoTo 0
    Exit Function
GetAparicionesNumero_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    Call HandleException(ErrNumber, ErrDescription, "BdDatos.GetAparicionesNumero")
    Err.Raise ErrNumber, ErrSource, ErrDescription
End Function





'------------------------------------------------------------------------------*
' Función        : GetAparicionesReintegro
' Fecha          : lu., 06/jul/2020 17:25:26
' Propósito      : Obtiene todas las apariciones de un reintegro para un sorteo
' Parámetros     : Reintegro del boleto
' Retorno        : Colección de tuplas con la aparición
'------------------------------------------------------------------------------*
'
Public Function GetAparicionesReintegro(vNewValue As Integer) As Variant
    Dim mRg As Range            ' Rango de trabajo
    Dim mRes As Range           ' Rango resultado
    Dim mRowI As Integer        ' Fila del primer registro
    Dim mRowF As Integer        ' Fila del ultimo registro
    Dim bFound As Boolean       ' Indicador de registro encontrado
    Dim mCol As Collection      ' Colección de apariciones
    Dim mTup As TuplaAparicion  ' Tupla de aparición del número
  
  On Error GoTo GetAparicionesReintegro_Error
    '
    '   Atendiendo al Juego establecemos el área de los números
    '   Establecemos como rango de trabajo la columna de los reintegros
    '
    Select Case JUEGO_DEFECTO
        Case Bonoloto, LoteriaPrimitiva:
            ' ($M:$M)
            Set mRg = mRgSorteos.Offset(0, 12).Resize(mRgSorteos.Rows.Count, 1)
        
        Case gordoPrimitiva, Euromillones:
            ' ($L:$L)
            Set mRg = mRgSorteos.Offset(0, 11).Resize(mRgSorteos.Rows.Count, 1)
    End Select
    '
    '   Creamos la colección
    '
    Set mCol = New Collection
    '
    '   Se busca el Numero
    '
    Set mRes = mRg.Find(vNewValue, _
                       lookat:=xlWhole, _
                       LookIn:=xlValues, _
                       MatchCase:=True)
               
    If Not (mRes Is Nothing) Then
        '
        '   Primera fila donde se encuentra el boleto
        '
        mRowI = mRes.Row
        mRowF = mRes.Row
        bFound = False
        '
        '   Creamos la primera tupla y guardamos
        '
        Set mTup = GetTupla(mRes)
        mCol.Add mTup
        '
        '   Bucle de busqueda siguiente
        '
        Do
            Set mRes = mRg.FindNext(mRes)
            '
            '   Si la fila es la misma o menor salimos del bucle
            '
            If mRes.Row <= mRowF Then
                bFound = True
            Else
                mRowF = mRes.Row
                
                Set mTup = GetTupla(mRes)
                mCol.Add mTup
            End If
        Loop While Not bFound
    End If
    '
    '   Devolvemos la colección de tuplas
    '
    Set GetAparicionesReintegro = mCol

GetAparicionesReintegro_CleanExit:
  On Error GoTo 0
    Exit Function
GetAparicionesReintegro_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    Call HandleException(ErrNumber, ErrDescription, "BdDatos.GetAparicionesReintegro")
    Err.Raise ErrNumber, ErrSource, ErrDescription
End Function



'------------------------------------------------------------------------------*
' Función        : GetAparicionesEstrella
' Fecha          : lu., 06/jul/2020 17:25:26
' Propósito      : Obtiene todas las apariciones de una estrella
' Parámetros     : Estrella del concurso
' Retorno        : Colección de tuplas con la aparición
'------------------------------------------------------------------------------*
'
Public Function GetAparicionesEstrella(vNewValue As Integer) As Variant
    Dim mRg As Range            ' Rango de trabajo
    Dim mRes As Range           ' Rango resultado
    Dim mRowI As Integer        ' Fila del primer registro
    Dim mRowF As Integer        ' Fila del ultimo registro
    Dim bFound As Boolean       ' Indicador de registro encontrado
    Dim mCol As Collection      ' Colección de apariciones
    Dim mTup As TuplaAparicion  ' Tupla de aparición del número
  
  On Error GoTo GetAparicionesEstrella_Error
    '
    '   Creamos la colección
    '
    Set mCol = New Collection
    '
    '   Atendiendo al Juego establecemos el área de los números
    '   Establecemos como rango de trabajo la columna de los reintegros
    '
    Select Case JUEGO_DEFECTO
        Case Bonoloto, LoteriaPrimitiva, gordoPrimitiva:
            '
            '   Sale sin tratamiento porque no hay estrellas en estos juegos
            '
            Set GetAparicionesEstrella = mCol
            Exit Function
        
        Case Euromillones:
            ' ($L:$M)
            Set mRg = mRgSorteos.Offset(0, 12).Resize(mRgSorteos.Rows.Count, 2)
    End Select
    '
    '   Se busca el Numero
    '
    Set mRes = mRg.Find(vNewValue, _
                       lookat:=xlWhole, _
                       LookIn:=xlValues, _
                       MatchCase:=True)
               
    If Not (mRes Is Nothing) Then
        '
        '   Primera fila donde se encuentra el boleto
        '
        mRowI = mRes.Row
        mRowF = mRes.Row
        bFound = False
        '
        '   Creamos la primera tupla y guardamos
        '
        Set mTup = GetTupla(mRes)
        mCol.Add mTup
        '
        '   Bucle de busqueda siguiente
        '
        Do
            Set mRes = mRg.FindNext(mRes)
            '
            '   Si la fila es la misma o menor salimos del bucle
            '
            If mRes.Row <= mRowF Then
                bFound = True
            Else
                mRowF = mRes.Row
                
                Set mTup = GetTupla(mRes)
                mCol.Add mTup
            End If
        Loop While Not bFound
    End If
    '
    '   Devolvemos la colección de tuplas
    '
    Set GetAparicionesEstrella = mCol

GetAparicionesEstrella_CleanExit:
  On Error GoTo 0
    Exit Function
GetAparicionesEstrella_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    Call HandleException(ErrNumber, ErrDescription, "BdDatos.GetAparicionesEstrella")
    Err.Raise ErrNumber, ErrSource, ErrDescription
End Function


' *===========(EOF): BdDatos.cls
