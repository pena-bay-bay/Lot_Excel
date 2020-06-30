VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BdDatos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'---------------------------------------------------------------------------------------
' Modulo    : BdDatos
' Creado    : 01-nov-2006 17:16
' Autor     : Carlos Almela Baeza
' Version   : 1.0.0 Fecha
' Objeto    : Objeto que abstrae los rangos de datos de la lógica
'---------------------------------------------------------------------------------------
Option Explicit

Const hjDatos = "Resultados"        'Hoja de resultados
Const hjBoletos = "Boletos"         'Hoja de Boletos
Const hjApuesta = "Apuestas"        'Hoja de Apuestas


Private mstrRangoResultados As String       ' Address tabla de Resultados
Private mstrRangoApuestas   As String       ' Address tabla de Apuestas
Private mstrRangoBoletos    As String       ' Address Tabla de Boletos
Private m_iUltimoRegistro   As Integer      ' Último registro
Private m_dtPrimerResultado As Date
Private m_dtUltimoResultado As Date
Private mInfo               As InfoSorteo   ' Información del concurso
'---------------------------------------------------------------------------------------
' Procedure : GetFecha
' Author    : CAB3780Y
' Date      : 29/03/2011
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Function GetFecha(mFecha As Date, mReg As Integer) As Date
    Dim rg_res          As Range        'Rango de búsqueda
    Dim res             As Range        'Rango de resultado
    Dim Fila            As Range        'Celda de la fecha
    Dim m_fechaOut      As Date         'Fecha de salida
    Dim m_fila          As Integer      'Fila offset de referencia
   
   On Error GoTo GetFecha_Error

    m_fechaOut = mFecha                 'Se asigna la fecha de entrada
    Set rg_res = Me.RangoResultados     'Obtenemos el rango de resultados
    '
    '   Se busca la fecha de entrada
    '
    Set res = rg_res.Find(mFecha, _
                          LookAt:=xlWhole, _
                          LookIn:=xlValues, _
                          MatchCase:=True)
        
    If Not (res Is Nothing) Then
        '
        '   A la fila donde localiza la fecha final del período
        '   Se le restan el número de registros  y uno para
        '   que sean los registros justos
        '
        m_fila = res.Row - mReg - 1
        '
        '   La primera celda es la fecha de los resultados
        '
        Set Fila = rg_res.Offset(m_fila, 2).Resize(1, 1)
        m_fechaOut = Fila.Value
    End If
    GetFecha = m_fechaOut

   On Error GoTo 0
     Exit Function

GetFecha_Error:

    Dim sNumber As Integer
    Dim sDescription As String
    Dim sSource As String
    With Err
        sNumber = .Number
        sDescription = .Description
        sSource = .Source
    End With
    Call HandleException(sNumber, sDescription, sSource)
    '    Sube el error
    Err.Raise sNumber, sSource, sDescription
   
End Function

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
    Set rg_res = Me.RangoResultados
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
        Set res = rg_res.Find(datFecha, LookAt:=xlWhole, LookIn:=xlValues, MatchCase:=True)
        '
        '   Si no hemos encontrado la fecha
        '
        If res Is Nothing Then
            '
            '   Buscamos la fecha anterior al sorteo
            '
            datFecha = mInfo.GetAnteriorSorteo(datFecha)
            ' datFecha = datFecha - 1
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
    Set rg_res = Me.RangoResultados
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
        Set res = rg_res.Find(datFecha, LookAt:=xlWhole, LookIn:=xlValues, MatchCase:=True)
        '
        '   Si no hemos encontrado la fecha
        '
        If res Is Nothing Then
            '
            '   Buscamos la fecha anterior al sorteo
            '
            datFecha = mInfo.GetProximoSorteo(datFecha)
            'datFecha = datFecha + 1
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

'---------------------------------------------------------------------------------------
' Procedimiento : RangoApuestas
' Creación      : 22-dic-2002 19:47
' Autor         : Carlos Almela Baeza
' Objeto        :
'---------------------------------------------------------------------------------------
'
Public Property Get RangoApuestas() As Range
   On Error GoTo RangoApuestas_Error
    
    set_mstrRangoApuestas
    Set RangoApuestas = Worksheets(hjApuesta).Range(mstrRangoApuestas)
   
   On Error GoTo 0
   Exit Property

RangoApuestas_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & _
        ") in procedure RangoApuestas of Módulo de clase BdDatos"
End Property
'---------------------------------------------------------------------------------------
' Procedimiento : RangoResultados
' Creación      :  22-dic-2002
' Autor         : Carlos Almela Baeza
' Objeto        :
'---------------------------------------------------------------------------------------
'
Public Property Get RangoResultados() As Range
   On Error GoTo RangoResultados_Error

    set_mstrRangoResultados
    Set RangoResultados = Worksheets(hjDatos).Range(mstrRangoResultados)

   On Error GoTo 0
   Exit Property

RangoResultados_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & _
        ") in procedure RangoResultados of Módulo de clase BdDatos"
End Property

'---------------------------------------------------------------------------------------
' Procedure : AddressResultados
' DateTime  : 03/jun/2007 20:45
' Author    : Carlos Almela Baeza
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get AddressResultados() As String
    If (mstrRangoResultados = "") Then
        set_mstrRangoResultados
    End If
    AddressResultados = mstrRangoResultados
End Property

'---------------------------------------------------------------------------------------
' Procedure : AddressApuestas
' DateTime  : 03/jun/2007 20:45
' Author    : Carlos Almela Baeza
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get AddressApuestas() As String
    If (mstrRangoApuestas = "") Then
        set_mstrRangoApuestas
    End If
    AddressApuestas = mstrRangoApuestas
End Property

'---------------------------------------------------------------------------------------
' Procedure : set_mstrRangoResultados
' DateTime  : 03/jun/2007 20:46
' Author    : Carlos Almela Baeza
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub set_mstrRangoResultados()
    Dim rg As Range
    Set rg = Worksheets(hjDatos).Range("A:A")
    mstrRangoResultados = "A2:N" & Application.WorksheetFunction.Subtotal(2, rg) + 1
End Sub

'---------------------------------------------------------------------------------------
' Procedure : set_mstrRangoApuestas
' DateTime  : 03/jun/2007 20:46
' Author    : Carlos Almela Baeza
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub set_mstrRangoApuestas()
    Dim rg As Range
    Set rg = Worksheets(hjApuesta).Range("A:A")
    mstrRangoApuestas = "A2:N" & Application.WorksheetFunction.Subtotal(2, rg) + 1
End Sub
'---------------------------------------------------------------------------------------
' Procedimiento : Get_Resultado
' Creación      : 20-nov-2006 23:25
' Autor         : Carlos Almela Baeza
' Objeto        :
'---------------------------------------------------------------------------------------
'
Public Function Get_Resultado(Fecha As Date) As Sorteo
    Dim objResult As Sorteo
    Dim rg_res As Range
    Dim res As Range
    Dim Fila As Range
    Dim Columnas As Integer
   On Error GoTo Get_Resultado_Error
    Set rg_res = Me.RangoResultados
    Set res = rg_res.Find(Fecha, LookAt:=xlWhole, LookIn:=xlValues, MatchCase:=True)
        
    If Not (res Is Nothing) Then
        Columnas = rg_res.Columns.Count
        Set objResult = New Sorteo
        Set Fila = rg_res.Offset(res.Row - 2, 0).Resize(1, Columnas)
        objResult.Constructor Fila
    End If
    Set Get_Resultado = objResult
    Set objResult = Nothing

   On Error GoTo 0
   Exit Function

Get_Resultado_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & _
        ") in procedure Get_Resultado of Módulo de clase BdDatos"

End Function


'---------------------------------------------------------------------------------------
' Procedure : RangoBoletos
' DateTime  : 03/jun/2007 20:43
' Author    : Carlos Almela Baeza
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get RangoBoletos() As Range
   On Error GoTo RangoBoletos_Error

    set_mstrRangoBoletos
    Set RangoBoletos = Worksheets(hjBoletos).Range(mstrRangoBoletos)

   On Error GoTo 0
   Exit Property

RangoBoletos_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & _
        ") in procedure RangoBoletos of Módulo de clase BdDatos"
End Property


'---------------------------------------------------------------------------------------
' Procedure : set_mstrRangoBoletos
' DateTime  : 03/jun/2007 20:43
' Author    : Carlos Almela Baeza
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub set_mstrRangoBoletos()
    Dim rg As Range
    Set rg = Worksheets(hjBoletos).Range("A:A")
    mstrRangoBoletos = "A2:K" & Application.WorksheetFunction.Subtotal(2, rg) + 1
End Sub

'---------------------------------------------------------------------------------------
' Procedimiento : Ir_A_Hoja
' Creación      : 05-nov-2006 20:38
' Autor         : Carlos Almela Baeza
' Objeto        : Selecciona la hoja del libro donde actua la macro
'---------------------------------------------------------------------------------------
'
Public Sub Ir_A_Hoja(hoja As String)
    Dim Wrk As Workbook
    Dim Pagina As Worksheet
    For Each Wrk In Workbooks
        For Each Pagina In Wrk.Sheets
            If Pagina.Name = hoja Then
                Pagina.Select
                Exit Sub
            End If
        Next Pagina
    Next Wrk
End Sub




'---------------------------------------------------------------------------------------
' Procedimiento : Resultados_Fechas
' Creación      : 27-mar-2007 23:48
' Autor         : Carlos Almela Baeza
' Objeto        :
'---------------------------------------------------------------------------------------
'
Public Function Resultados_Fechas(ByVal Fecha_Inicial As Date, ByVal Fecha_Final As Date) As Range
    Dim res, m_rgResult As Range
    Dim rg_res As Range
    Dim m_rowIni, m_rowFin, Columnas As Integer
    Dim m_filas As Integer
    Dim mInfo As InfoSorteo
    
   On Error GoTo Resultados_Fechas_Error
    '
    '   Comprobamos que sean fechas de sorteo
    '
    Set mInfo = New InfoSorteo
    mInfo.Constructor JUEGO_DEFECTO
        
    If Not mInfo.EsFechaSorteo(Fecha_Inicial) Then   ' Mayor o Igual a la fecha inicial
        Fecha_Inicial = mInfo.GetProximoSorteo(Fecha_Inicial)
    End If
    
    If Not mInfo.EsFechaSorteo(Fecha_Final) Then     ' Menor o Igual a la fecha final
        Fecha_Final = mInfo.GetAnteriorSorteo(Fecha_Final)
    End If
    '
    '   Control de fuera de rango
    '
    If (Fecha_Inicial < PrimerResultado) Then
        Fecha_Inicial = PrimerResultado
    End If
    If (Fecha_Final > UltimoResultado) Then
        Fecha_Final = UltimoResultado
    End If
    '
    '
    '
    m_rowIni = GetRowFechaMayorOIgual(Fecha_Inicial)
'    Set rg_res = Me.RangoResultados
'    Set res = rg_res.Find(Fecha_Inicial, lookat:=xlWhole, LookIn:=xlValues, MatchCase:=True)
'
'    If Not (res Is Nothing) Then
'        m_rowIni = res.Row - 2
'    End If
'    Set rg_res = Me.RangoResultados
'    Set res = rg_res.Find(Fecha_Final, lookat:=xlWhole, LookIn:=xlValues, MatchCase:=True)
'    If Not (res Is Nothing) Then
'        m_rowFin = res.Row - 2
'    End If
    m_rowFin = GetRowFechaMenorOIgual(Fecha_Final)
    
    If IsEmpty(m_rowIni) _
    Or IsEmpty(m_rowFin) Then

        Err.Raise 100, "BdDatos.Resultados_Fechas", _
        " Error en el rango de fechas(" & Fecha_Inicial & " - " & _
                  Fecha_Final & ")  , no encontradas"
    End If
    
    Set rg_res = Me.RangoResultados
    Columnas = rg_res.Columns.Count
    m_filas = rg_res.Rows.Count
    m_filas = (m_rowFin - m_rowIni + 1)
    Set m_rgResult = rg_res.Offset(m_rowIni, 0).Resize(m_filas, Columnas)
    '
    '
    '
    Set Resultados_Fechas = m_rgResult

   On Error GoTo 0
   Exit Function

Resultados_Fechas_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    If ErrNumber <> 100 Then
        Call HandleException(ErrNumber, ErrDescription, ErrSource, _
                            "BdDatos.Resultados_Fechas")
    End If
    Err.Raise ErrNumber, "BdDatos.Resultados_Fechas", ErrDescription
End Function
'---------------------------------------------------------------------------------------
' Procedure : PrimerResultado
' DateTime  : 01/04/2007 21:38
' Author    : Carlos Almela Baeza
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get PrimerResultado() As Date
    Dim rg_res As Range
    
    If m_dtPrimerResultado = 0 Then
        
        Set rg_res = Me.RangoResultados
        m_dtPrimerResultado = CDate(rg_res.Offset(0, 2).Resize(1, 1).Value)
    End If
    PrimerResultado = m_dtPrimerResultado

End Property

'---------------------------------------------------------------------------------------
' Procedure : UltimoResultado
' DateTime  : 01/04/2007 21:38
' Author    : Carlos Almela Baeza
' Purpose   : Obtiene la fecha del último resultado de la base de datos
'---------------------------------------------------------------------------------------
'
Public Property Get UltimoResultado() As Date
    Dim rg_res As Range                     ' Definimos el rango de datos de resultados
    
    If m_dtUltimoResultado = 0 Then             ' Si no se tiene la fecha definida se
                                                ' calcula
        
        Set rg_res = Me.RangoResultados         ' obtenemos el rango de la propiedad de la
                                                ' clase
        m_dtUltimoResultado = CDate(rg_res.Offset(rg_res.Rows.Count - 1, 2).Resize(1, 1).Value)
                                                ' Buscamos la última celda de la primera
                                                ' columna y lo convertimos en fecha
    End If
    UltimoResultado = m_dtUltimoResultado        ' Devolvemos la fecha almacenada
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

  On Error GoTo UltimoRegistro_Error

    If m_iUltimoRegistro = 0 Then
        Dim rg_res As Range                     ' Definimos el rango de datos de resultados
        Set rg_res = Me.RangoResultados         ' obtenemos el rango de la propiedad de la
                                                ' clase
        m_iUltimoRegistro = rg_res.Offset(rg_res.Rows.Count - 1, 0).Resize(1, 1).Value
                                                ' Buscamos la última celda de la primera
                                                ' columna y lo convertimos en fecha
    End If
    UltimoRegistro = m_iUltimoRegistro

UltimoRegistro_CleanExit:
   On Error GoTo 0
    Exit Property

UltimoRegistro_Error:
    

    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    '   Audita el error
    Call HandleException(ErrNumber, ErrDescription, "BdDatos.UltimoRegistro")
    '   Lanza el Error
    Err.Raise ErrNumber, ErrSource, ErrDescription
End Property


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
    Set m_rgRes = Me.RangoResultados
    '
    '   Buscamos el registro menor o igual a la fecha
    '
    m_bFound = False
    While Not (m_bFound)
        '
        '   busca en el rango la fecha
        '
        Set m_vRes = m_rgRes.Find(datFecha, LookAt:=xlWhole, LookIn:=xlValues, MatchCase:=True)
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
    '   Audita el error
    Call HandleException(ErrNumber, ErrDescription, "BdDatos.GetRegistroFecha")
    '   Lanza el Error
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

