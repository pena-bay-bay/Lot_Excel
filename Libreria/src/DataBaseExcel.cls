VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DataBaseExcel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' *============================================================================*
' *
' *     Fichero    : DataBaseExcel.cls
' *
' *     Autor      : Carlos Almela Baeza
' *     Creación   : Mie, 02/may/2018 16:54:26
' *     Versión    : 1.0
' *     Propósito  : Gestionar los libros de datos como una base de datos
' *                  Trabajando con rangos
' *                  Se encarga de abrir los libros de los distintos juegos
' *                  y suministrar datos definiendo los rangos de los datos
' *============================================================================*
Option Explicit
Option Base 0
'--- Constantes ---------------------------------------------------------------*
Private Const NOMBRE_HOJASDATOS = "Bonoloto;Primitiva;Euromillon;Gordo;Parametros;Boletos;Apuestas;Participantes"
'Private Const LIBRO_DATOS = "Lot_Datos.xlsx"
Private Const LIBRO_DATOS = "Lot_Datos_Prueba.xlsx"

'--- Variables Privadas -------------------------------------------------------*
Private mDBIsOpen       As Boolean
Private mWorkbookDatos  As Workbook
Private mHoja           As Worksheet
Private mIdx            As Integer
Private mRutaLibro      As String
Private mListaLibros    As Variant
Private mListaHojas     As Variant
Private mRango          As Range
Private mColumna        As Range
Private mRgFind         As Range
Private mRow            As Integer
Private mRowIni         As Integer
Private mRowFin         As Integer
Private mColumns        As Integer
Private mRows           As Integer
Private mErr            As Integer
Private mErrDes         As String
Private mErrProc        As String

'--- Propiedades --------------------------------------------------------------*
Public Enum NombreTabla
    tblBonoloto = 1
    tblPrimitiva = 2
    tblEuromillon = 3
    tblGordo = 4
    tblParametros = 5
    tblBoleto = 6
    tblApuestas = 7
    tblParticipantes = 8
End Enum

Public Property Get IsOpen() As Boolean
    IsOpen = mDBIsOpen
End Property

Public Property Get ErrNumber() As Integer
    ErrNumber = mErr
End Property

Public Property Get ErrDescription() As String
    ErrDescription = mErrDes
End Property

Public Property Get ErrProcces() As String
    ErrProcces = mErrProc
End Property

'--- Metodos Privados ---------------------------------------------------------*
Private Sub Class_Initialize()
    '   Componer la matriz de libros y la matriz de resultados
    '
    '
    mRutaLibro = ThisWorkbook.Path
    mErr = 0
    mErrDes = ""
    mErrProc = ""
End Sub

Private Sub Class_Terminate()
    Set mWorkbookDatos = Nothing
    Set mHoja = Nothing
End Sub

'------------------------------------------------------------------------------*
' Procedimiento  : GetTabla
' Fecha          : 15/10/2018
' Propósito      : Devolver el Rango perteneciente a la tabla a tratar
' Parámetros     : NombreTabla: Tabla a tratar
' Retorno        :
'------------------------------------------------------------------------------*
'
Private Function GetTabla(parTabla As NombreTabla) As Range
    Dim strNombreTabla As String
    Dim tmp As Variant
On Error GoTo GetTabla_Error
    '
    '  Obtenemos los nombres de las hojas del libro de datos
    '
    tmp = Split(NOMBRE_HOJASDATOS, ";")
    '
    '
    '
    strNombreTabla = tmp(parTabla - 1)
    '
    '   Nos posicionamos en el libro
    '
    Set mWorkbookDatos = Workbooks(LIBRO_DATOS)
    '
    '
    '
    Set GetTabla = mWorkbookDatos.Worksheets(strNombreTabla).Range("A1").CurrentRegion
    
 On Error GoTo 0
GetTabla__CleanExit:
    Exit Function
            
GetTabla_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    '   Audita el error
    Call HandleException(ErrNumber, ErrDescription, "DataBaseExcel.GetTabla", ErrSource)
    '   Lanza el Error
    Err.Raise ErrNumber, "DataBaseExcel.GetTabla", ErrDescription
End Function

'--- Mensajes -----------------------------------------------------------------*
'--- Errores ------------------------------------------------------------------*
'--- Metodos Publicos ---------------------------------------------------------*
'------------------------------------------------------------------------------*
' Procedimiento  : DeleteRow
' Fecha          : 14/10/2018
' Propósito      : Borra una fila de la tabla por su Id
' Parámetros     : Id: Clave del registro, NombreTabla: Tabla a borrar
' Retorno        :
'------------------------------------------------------------------------------*
'
Public Sub DeleteRow(parId As Integer, parTabla As NombreTabla)
 On Error GoTo DeleteRow_Error
    '
    '  Localiza la fila a borrar
    '
    Set mRango = Me.GetRowById(parId, parTabla)
    '
    '   Elimina la fila
    '
    mRango.Delete Shift:=xlShiftUp
 
 On Error GoTo 0
DeleteRow__CleanExit:
    Exit Sub
            
DeleteRow_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    '   Audita el error
    Call HandleException(ErrNumber, ErrDescription, "DataBaseExcel.DeleteRow", ErrSource)
    '   Lanza el Error
    Err.Raise ErrNumber, "DataBaseExcel.DeleteRow", ErrDescription
End Sub

'------------------------------------------------------------------------------*
' Procedimiento  : GetLastID
' Fecha          : 03/06/2018
' Propósito      : Devuelve el último Id de una tabla
' Parámetros     : Nombre de la tabla
' Retorno        : Un identificador de la tabla
'------------------------------------------------------------------------------*
'
Public Function GetLastID(parTabla As NombreTabla) As Integer
    
 On Error GoTo GetLastID_Error
    '
    '   Obtiene el rango de la tabla a consultar
    '
    Set mRango = GetTabla(parTabla)
    '
    '   Obtiene la columna ID que es la primera
    '
    Set mColumna = mRango.Offset(0, 0).Resize(1, 1).EntireColumn
    '
    '   Busca el máximo valor de la fila
    '
    GetLastID = Application.WorksheetFunction.Max(mColumna)
    '
    '   Añade uno al valor de la fila
    '
    GetLastID = GetLastID
 
 
 On Error GoTo 0
GetLastID__CleanExit:
    Exit Function
            
GetLastID_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    '   Audita el error
    Call HandleException(ErrNumber, ErrDescription, "DataBaseExcel.GetLastID", ErrSource)
    '   Lanza el Error
    Err.Raise ErrNumber, "DataBaseExcel.GetLastID", ErrDescription
End Function

'------------------------------------------------------------------------------*
' Procedimiento  : GetRowById
' Fecha          : 03/06/2018
' Propósito      : Devuelve una fila de la tabla por Id
' Parámetros     : Id de la fila, Nombre de la tabla
' Retorno        : Rango con los datos
'------------------------------------------------------------------------------*
'
Public Function GetRowById(vData As Integer, parTabla As NombreTabla) As Range
    
 On Error GoTo GetRowById_Error
    '
    '   Obtiene el rango de la tabla a consultar
    '
    Set mRango = GetTabla(parTabla)
    mColumns = mRango.Columns.Count
    '
    '   Obtiene la columna ID que es la primera
    '
    Set mColumna = mRango.Offset(0, 0).Resize(1, 1).EntireColumn
    '
    '   Busca el código en la tabla
    '
    Set mRgFind = mColumna.Find(vData, LookAt:=xlWhole, LookIn:=xlValues, MatchCase:=True)
    '
    '   Si no lo encuentra el resultado está vacio
    If (mRgFind Is Nothing) Then
        '
        '   Provoca el error de código no encontrado
        '
        On Error Resume Next
        mErr = ERR_REGISTRONOTFOUND
        mErrDes = MSG_REGISTRONOTFOUND
        mErrProc = "GetRowById"
        Err.Raise ERR_REGISTRONOTFOUND, "GetRowById", MSG_REGISTRONOTFOUND & vData
        GoTo GetRowById__CleanExit
    Else
        '
        '   Si lo encuentra calcula la fila y las columnas
        '
        mRow = mRgFind.Row
    End If
    '
    '     Offset( _RowSize_ , _ColumnSize_ )
    '     Resize( _RowSize_ , _ColumnSize_ )
    '
    Set GetRowById = mRango.Offset(mRow - 1, 0).Resize(1, mColumns)

 
 On Error GoTo 0
GetRowById__CleanExit:
    Exit Function
            
GetRowById_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    '   Audita el error
    Call HandleException(ErrNumber, ErrDescription, "DataBaseExcel.GetRowById", ErrSource)
    '   Lanza el Error
    Err.Raise ErrNumber, "DataBaseExcel.GetRowById", ErrDescription
End Function
'------------------------------------------------------------------------------*
' Procedimiento  : GetLastRow
' Fecha          : 03/06/2018
' Propósito      : Devuelve la última fila vacia de la tabla para insertar
' Parámetros     : Nombre de la tabla
' Retorno        : Rango de la última fila
'------------------------------------------------------------------------------*
'
Public Function GetLastRow(parTabla As NombreTabla) As Range
On Error GoTo GetLastRow_Error
    '
    '   Obtenemos el rango de la tabla
    '
    Set mRango = GetTabla(parTabla)
    '
    '   Parametros de la tabla filas y columnas
    '
    mColumns = mRango.Columns.Count
    mRows = mRango.Rows.Count
    '
    '   Devolvemos la última fila
    '
    Set GetLastRow = mRango.Offset(mRows, 0).Resize(1, mColumns)
   
 
 On Error GoTo 0
GetLastRow__CleanExit:
    Exit Function
            
GetLastRow_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    '   Audita el error
    Call HandleException(ErrNumber, ErrDescription, "DataBaseExcel.GetLastRow", ErrSource)
    '   Lanza el Error
    Err.Raise ErrNumber, "DataBaseExcel.GetLastRow", ErrDescription
End Function

'------------------------------------------------------------------------------*
' Procedimiento  : SelectByFechas
' Fecha          : 14/10/2018
' Propósito      : Selecciona un rango de sorteos entre dos fechas
' Parámetros     : FechaInicial, FechaFinal, tabla
' Retorno        : Rango de registros entre las fechas
'------------------------------------------------------------------------------*
'
Public Function SelectByFechas(parFechaIni As Date, parFechaFin As Date, parTabla As NombreTabla) As Range
    Dim mColumn As Integer
    
On Error GoTo SelectByFechas_Error
    '
    '   Obtiene el rango de la tabla a consultar
    '
    Set mRango = GetTabla(parTabla)
    mColumns = mRango.Columns.Count
    mRows = mRango.Rows.Count
    '
    '   Obtiene la columna de fechas que es la tercera
    '
    Select Case parTabla
        Case tblBonoloto, tblPrimitiva, tblGordo, tblEuromillon: mColumn = 2
        Case tblParametros: mColumn = 8
        Case tblBoleto: mColumn = 5
        Case tblApuestas: mColumn = 4
    End Select
    '
    '
    '
    Set mColumna = mRango.Offset(0, 2).Resize(1, 1).EntireColumn
    '
    '   Busca la fecha inicial
    '
    Set mRgFind = mColumna.Find(parFechaIni, LookAt:=xlWhole, LookIn:=xlValues, MatchCase:=True)
    '
    '   Si no lo encuentra el resultado está vacio
    If (mRgFind Is Nothing) Then
        '
        '   Provoca el error de código no encontrado
        '
        On Error Resume Next
        mErr = ERR_REGISTRONOTFOUND
        mErrDes = MSG_REGISTRONOTFOUND
        mErrProc = "SelectByFechas"
        Err.Raise ERR_REGISTRONOTFOUND, "SelectByFechas", MSG_REGISTRONOTFOUND & parFechaIni
        GoTo SelectByFechas__CleanExit
    Else
        '
        '   Si lo encuentra calcula la fila y las columnas
        '
        mRowIni = mRgFind.Row
    End If
    '
    '   Busca la fecha final
    '
    Set mRgFind = mColumna.Find(parFechaFin, LookAt:=xlWhole, LookIn:=xlValues, MatchCase:=True)
    '
    '   Si no lo encuentra el resultado está vacio
    If (mRgFind Is Nothing) Then
        '
        '   Provoca el error de código no encontrado
        '
        On Error Resume Next
        mErr = ERR_REGISTRONOTFOUND
        mErrDes = MSG_REGISTRONOTFOUND
        mErrProc = "SelectByFechas"
        Err.Raise ERR_REGISTRONOTFOUND, "SelectByFechas", MSG_REGISTRONOTFOUND & parFechaIni
        GoTo SelectByFechas__CleanExit
    Else
        '
        '   Si lo encuentra calcula la fila y las columnas
        '
        mRowFin = mRgFind.Row
    End If
    '
    '     Offset( _RowSize_ , _ColumnSize_ )
    '     Resize( _RowSize_ , _ColumnSize_ )
    '
    Set SelectByFechas = mRango.Offset(mRowIni - 1, 0).Resize(mRowFin - mRowIni + 1, mColumns)
 
 
 On Error GoTo 0
SelectByFechas__CleanExit:
    Exit Function
            
SelectByFechas_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    '   Audita el error
    Call HandleException(ErrNumber, ErrDescription, "DataBaseExcel.SelectByFechas", ErrSource)
    '   Lanza el Error
    Err.Raise ErrNumber, "DataBaseExcel.SelectByFechas", ErrDescription

End Function

'------------------------------------------------------------------------------*
' Procedimiento  : SelectByIds
' Fecha          : 14/10/2018
' Propósito      : Devuelve un rango entre dos Ids de registros
' Parámetros     : parIdIni Id inicial del Rango
'                  ParIdFin Id final del rango
'                  parTabla
' Retorno        :
'------------------------------------------------------------------------------*
'
Public Function SelectByIds(parIdIni As Integer, parIdFin As Integer, parTabla As NombreTabla) As Range
On Error GoTo SelectByIds_Error
    '
    '   Obtiene el rango de la tabla a consultar
    '
    Set mRango = GetTabla(parTabla)
    mColumns = mRango.Columns.Count
    '
    '   Obtiene la columna de fechas que es la tercera
    '
    Set mColumna = mRango.Offset(0, 0).Resize(1, 1).EntireColumn
    '
    '   Busca la fecha inicial
    '
    Set mRgFind = mColumna.Find(parIdIni, LookAt:=xlWhole, LookIn:=xlValues, MatchCase:=True)
    '
    '   Si no lo encuentra el resultado está vacio
    If (mRgFind Is Nothing) Then
        '
        '   Provoca el error de código no encontrado
        '
        On Error Resume Next
        mErr = ERR_REGISTRONOTFOUND
        mErrDes = MSG_REGISTRONOTFOUND
        mErrProc = "SelectByIds"
        Err.Raise ERR_REGISTRONOTFOUND, "SelectByIds", MSG_REGISTRONOTFOUND & parIdFin
        GoTo SelectByIds__CleanExit
    Else
        '
        '   Si lo encuentra calcula la fila y las columnas
        '
        mRowIni = mRgFind.Row
    End If
    '
    '   Busca la fecha final
    '
    Set mRgFind = mColumna.Find(parIdFin, LookAt:=xlWhole, LookIn:=xlValues, MatchCase:=True)
    '
    '   Si no lo encuentra el resultado está vacio
    If (mRgFind Is Nothing) Then
        '
        '   Provoca el error de código no encontrado
        '
        On Error Resume Next
        mErr = ERR_REGISTRONOTFOUND
        mErrDes = MSG_REGISTRONOTFOUND
        mErrProc = "SelectByIds"
        Err.Raise ERR_REGISTRONOTFOUND, "SelectByIds", MSG_REGISTRONOTFOUND & parIdFin
        GoTo SelectByIds__CleanExit
    Else
        '
        '   Si lo encuentra calcula la fila y las columnas
        '
        mRowFin = mRgFind.Row
    End If
    '
    '     Offset( _RowSize_ , _ColumnSize_ )
    '     Resize( _RowSize_ , _ColumnSize_ )
    '
    Set SelectByIds = mRango.Offset(mRowIni - 1, 0).Resize(mRowFin - mRowIni + 1, mColumns)
 
 
 On Error GoTo 0
SelectByIds__CleanExit:
    Exit Function
            
SelectByIds_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    '   Audita el error
    Call HandleException(ErrNumber, ErrDescription, "DataBaseExcel.SelectByIds", ErrSource)
    '   Lanza el Error
    Err.Raise ErrNumber, "DataBaseExcel.SelectByIds", ErrDescription

End Function

'------------------------------------------------------------------------------*
' Procedimiento  : Abrir
' Fecha          : 03/06/2018
' Propósito      : Cargar el libro de datos de la aplicación
'------------------------------------------------------------------------------*
'
Public Sub Abrir()
    Dim mTmpPathFile        As String
    
 On Error GoTo Abrir_Error
    '
    '   Comprobar que el libro esta abierto, sino es asi abrirlo
    '
    mDBIsOpen = False
    '
    '   Para cada libro en la colección
    '
    For Each mWorkbookDatos In Application.Workbooks
        '
        '   Si el nombre del libro es el Libro de datos
        '
        If mWorkbookDatos.Name = LIBRO_DATOS Then
            '
            '   Indicador de abierto a true y salimos del
            '   bucle
            '
            mDBIsOpen = True
            Exit For
        End If
    Next mWorkbookDatos
    '
    '   Si el libro no esta abierto
    '
    If mDBIsOpen = False Then
        '
        '   Componemos la ruta del libro de datos
        '
        mTmpPathFile = mRutaLibro & "\" & LIBRO_DATOS
        '
        '   Abrimos el libro
        '
        Set mWorkbookDatos = Workbooks.Open(mTmpPathFile, False)
        '
        '   Indicador de abierto a true
        '
        mDBIsOpen = True
    End If
    
 On Error GoTo 0
Abrir__CleanExit:
    Exit Sub
            
Abrir_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    '   Audita el error
    Call HandleException(ErrNumber, ErrDescription, "DataBaseExcel.Abrir", ErrSource)
    '   Lanza el Error
    Err.Raise ErrNumber, "DataBaseExcel.Abrir", ErrDescription
End Sub
'------------------------------------------------------------------------------*
' Procedimiento  : Cerrar
' Fecha          : 03/06/2018
' Propósito      : Cierra el libro de datos guardandolo
'------------------------------------------------------------------------------*
'
Public Sub Cerrar()
 On Error GoTo Cerrar_Error
   
    '
    '   Comprobar que el libro esta abierto, sino es asi abrirlo
    '
    For Each mWorkbookDatos In Application.Workbooks
        '
        '   Si el libro es el libro de datos
        '
        If mWorkbookDatos.Name = LIBRO_DATOS Then
            '
            '   Salimos del bucle
            '
            Exit For
        End If
    Next mWorkbookDatos
    '
    '   Si el libro esta abierto
    '
    If mDBIsOpen = True Then
        '
        '   Cerramos el libro de datos
        '
        Workbooks(LIBRO_DATOS).Close True
        '
        '   No hay libro abierto
        '
        mDBIsOpen = False
    End If
    
 On Error GoTo 0
Cerrar__CleanExit:
    Exit Sub
            
Cerrar_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    '   Audita el error
    Call HandleException(ErrNumber, ErrDescription, "DataBaseExcel.Cerrar", ErrSource)
    '   Lanza el Error
    Err.Raise ErrNumber, "DataBaseExcel.Cerrar", ErrDescription
End Sub
'------------------------------------------------------------------------------*
' Procedimiento  : GetRowByColumn
' Fecha          : do., 03/feb/2019 22:48:13
' Propósito      : Devuelve una fila de la tabla buscando en la columna enesima
' Parámetros     : Id de la fila, nColumn numero de columna a buscar,
'                  Nombre de la tabla
' Retorno        : Rango con los datos
'------------------------------------------------------------------------------*
'
Public Function GetRowByColumn(vData As Variant, nColumn As Integer, parTabla As NombreTabla) As Range
    
 On Error GoTo GetRowByColumn_Error
    '
    '   Obtiene el rango de la tabla a consultar
    '
    Set mRango = GetTabla(parTabla)
    mColumns = mRango.Columns.Count
    '
    '   Obtiene la columna enesima que es la primera
    '
    Set mColumna = mRango.Offset(0, nColumn).Resize(1, 1).EntireColumn
    '
    '   Busca el código en la tabla
    '
    Set mRgFind = mColumna.Find(vData, LookAt:=xlWhole, LookIn:=xlValues, MatchCase:=True)
    '
    '   Si no lo encuentra el resultado está vacio
    If (mRgFind Is Nothing) Then
        '
        '   Provoca el error de código no encontrado
        '
        On Error Resume Next
        mErr = ERR_REGISTRONOTFOUND
        mErrDes = MSG_REGISTRONOTFOUND
        mErrProc = "GetRowByColumn"
        Err.Raise ERR_REGISTRONOTFOUND, "GetRowByColumn", MSG_REGISTRONOTFOUND & vData
        GoTo GetRowByColumn__CleanExit
    Else
        '
        '   Si lo encuentra calcula la fila y las columnas
        '
        mRow = mRgFind.Row
    End If
    '
    '     Offset( _RowSize_ , _ColumnSize_ )
    '     Resize( _RowSize_ , _ColumnSize_ )
    '
    Set GetRowByColumn = mRango.Offset(mRow - 1, 0).Resize(1, mColumns)

 
 On Error GoTo 0
GetRowByColumn__CleanExit:
    Exit Function
            
GetRowByColumn_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    '   Audita el error
    Call HandleException(ErrNumber, ErrDescription, "DataBaseExcel.GetRowByColumn", ErrSource)
    '   Lanza el Error
    Err.Raise ErrNumber, "DataBaseExcel.GetRowByColumn", ErrDescription
End Function

'------------------------------------------------------------------------------*
' Procedimiento  : GetFirstID
' Fecha          : 03/06/2018
' Propósito      : Devuelve el último Id de una tabla
' Parámetros     : Nombre de la tabla
' Retorno        : Un identificador de la tabla
'------------------------------------------------------------------------------*
'
Public Function GetFirstID(parTabla As NombreTabla) As Integer
    
 On Error GoTo GetFirstID_Error
    '
    '   Obtiene el rango de la tabla a consultar
    '
    Set mRango = GetTabla(parTabla)
    '
    '   Obtiene la columna ID que es la primera
    '
    Set mColumna = mRango.Offset(0, 0).Resize(1, 1).EntireColumn
    '
    '   Busca el minimo valor de la fila
    '
    GetFirstID = Application.WorksheetFunction.Min(mColumna)
 
 On Error GoTo 0
GetFirstID__CleanExit:
    Exit Function
            
GetFirstID_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    '   Audita el error
    Call HandleException(ErrNumber, ErrDescription, "DataBaseExcel.GetFirstID", ErrSource)
    '   Lanza el Error
    Err.Raise ErrNumber, "DataBaseExcel.GetFirstID", ErrDescription
End Function
'------------------------------------------------------------------------------*
' Procedimiento  : SelectApuestasByBoleto
' Fecha          : ma., 08/oct/2019 19:34:20
' Propósito      : Encontrar los registros de apuestas con el mismo Id de Boleto
' Parámetros     : Id Boleto y Nombre de la tabla de apuestas
' Retorno        : Colección de tuplas con los registros
'------------------------------------------------------------------------------*
'
Public Function SelectApuestasByBoleto(vData As Integer, _
                                       parTabla As NombreTabla) As Collection
    Dim mTuplas     As Collection
    Dim mTup        As TuplaAparicion
    Dim mAddressI   As String
    Dim mBuscando   As Boolean
    
 On Error GoTo SelectApuestasByBoleto_Error
    '
    '   Creamos la colección de registros
    '
    Set mTuplas = New Collection
    '
    '   Obtiene el rango de la tabla a consultar
    '
    Set mRango = GetTabla(parTabla)
    '
    '   Obtiene la columna IdBoleto de la tabla Apuestas que es la 2ª
    '
    Set mColumna = mRango.Offset(0, 1).Resize(1, 1).EntireColumn
    '
    '   Busca el la fecha final
    '
    Set mRgFind = mColumna.Find(vData, LookAt:=xlWhole, LookIn:=xlValues, MatchCase:=True)
    '
    '   Si no encontramos nada, provocamos el error
    '
    If mRgFind Is Nothing Then
        mErr = ERR_REGISTRONOTFOUND
        mErrDes = MSG_REGISTRONOTFOUND & vData
        mErrProc = "SelectApuestasByBoleto"
'        Err.Raise ERR_REGISTRONOTFOUND, "SelectApuestasByBoleto", MSG_REGISTRONOTFOUND & vData
    Else
        '
        '   inicializamos campos de error
        ClearError
        '
        '   Inicializamos Control del bucle
        mBuscando = True
        '
        '   Estblecemos fila inicial
        mAddressI = mRgFind.Address
        '
        '   Bucle de busqueda
        While mBuscando
            '
            '   Establezco la fila donde se ha encontrado
            mRow = mRgFind.Row
            '
            '   Creamos la tupla
            Set mTup = New TuplaAparicion
            '
            '   Obtenemos el Id de la fila (posición fila ,0)
            mTup.NumeroRegistro = mRango.Offset(mRow - 1, 0).Resize(1, 1).Value
            '
            '   Agregamos registro a la colección
            If mTuplas.Count = 0 Then
                mTuplas.Add mTup, CStr(mRow)
            Else
                mTuplas.Add mTup, CStr(mRow), , mTuplas.Count
            End If
            '
            '   volvemos a buscar a partir de la ultima dirección
            Set mRgFind = mColumna.FindNext(mRgFind)
            '
            '   Si no encontramos nada mas o la dirección es la inicial
            If (mRgFind Is Nothing) Or (mAddressI = mRgFind.Address) Then
                mBuscando = False
            End If
        Wend
    End If
    '
    '   Devolvemos la colección de registros encontrados
    '
    Set SelectApuestasByBoleto = mTuplas
    
 On Error GoTo 0
SelectApuestasByBoleto__CleanExit:
    Exit Function
            
SelectApuestasByBoleto_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    '   Audita el error
    Call HandleException(ErrNumber, ErrDescription, "DataBaseExcel.SelectApuestasByBoleto", ErrSource)
    '   Lanza el Error
    Err.Raise ErrNumber, "DataBaseExcel.SelectApuestasByBoleto", ErrDescription
End Function
'
'   Inicializa los campos de error
'
Private Sub ClearError()
    mErr = 0
    mErrDes = Empty
    mErrProc = Empty
End Sub
'
'   TODO: Función que devuelve las filas donde se encuentra un número en un rango definido
'         Parametros: Numero, tabla, IdInicial, IdFinal, FndEstrellas, FndReintegro
'         Devuelve: colección de TuplaAparicion
'
' *===========(EOF): DataBaseExcel.cls

