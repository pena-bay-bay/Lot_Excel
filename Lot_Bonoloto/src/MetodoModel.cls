VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MetodoModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' *============================================================================*
' *
' *     Fichero    : MetodoModel.cls
' *
' *     Autor      : Carlos Almela Baeza
' *     Creación   : ma., 10/mar/2020 20:28:54
' *     Versión    : 1.0
' *     Propósito  : Estructura de serialización
' *                  en una hoja Excel de la clase Método de Sugerencia
' *
' *============================================================================*
Option Explicit
Option Base 0
'
'--- Variables Privadas -------------------------------------------------------*
Private mCol                As Metodos      ' Colección de metodos
Private mMetodo             As Metodo       ' Metodo de sugerencia Current
'----Paginado
Private mLinePage           As Integer      ' Lineas por página
Private mCurrentPage        As Integer      ' Página actual
Private mTotalRecord        As Integer      ' Total de registros
Private mTotalPages         As Integer      ' Total de páginas
'----Area de Datos
Private mwsData             As Worksheet    ' hoja de datos del repositorio
Private mrngIds             As Range        ' Rango de los datos
Private mLastRow            As Integer      ' ultima fila del rango
Private mFirstRow           As Integer      ' ultima fila del rango
'--- Constantes ---------------------------------------------------------------*
'--- Mensajes -----------------------------------------------------------------*
'--- Errores ------------------------------------------------------------------*
'--- Propiedades --------------------------------------------------------------*
'------------------------------------------------------------------------------*
' Procedimiento  : Metodos
' Fecha          : mi., 11/mar/2020 19:29:01
' Propósito      : Colección de objetos Metodo
'------------------------------------------------------------------------------*
Public Property Get Metodos() As Metodos
    Set Metodos = mCol
End Property
Public Property Set Metodos(vDatNew As Metodos)
    Set mCol = vDatNew
End Property
'------------------------------------------------------------------------------*
' Procedimiento  : LinePerPage
' Fecha          : mi., 11/mar/2020 19:29:01
' Propósito      : Numero de registros por pagina
'------------------------------------------------------------------------------*
Public Property Get LinePerPage() As Integer
    LinePerPage = mLinePage
End Property
Public Property Let LinePerPage(vDatNew As Integer)
    mLinePage = vDatNew
End Property
'------------------------------------------------------------------------------*
' Procedimiento  : TotalPages
' Fecha          : mi., 11/mar/2020 19:29:01
' Propósito      : Numero total de páginas del repositorio
'------------------------------------------------------------------------------*
Public Property Get TotalPages() As Integer
    TotalPages = mTotalPages
End Property
'------------------------------------------------------------------------------*
' Procedimiento  : TotalRecords
' Fecha          : mi., 11/mar/2020 19:29:01
' Propósito      : Total de registros en la base de datos
'------------------------------------------------------------------------------*
Public Property Get TotalRecords() As Integer
    TotalRecords = mTotalRecord
End Property

'------------------------------------------------------------------------------*
' Procedimiento  : Metodo
' Fecha          : mi., 11/mar/2020 19:29:01
' Propósito      : Objeto serializado Current Metodo
'------------------------------------------------------------------------------*
Public Property Get Metodo() As Metodo
    Set Metodo = mMetodo
End Property
Public Property Set Metodo(vDatNew As Metodo)
    Set mMetodo = vDatNew
End Property

'--- Métodos Privados ---------------------------------------------------------*
'------------------------------------------------------------------------------*
' Procedimiento  : GetRegistro
' Fecha          : lu., 16/mar/2020 18:38:45
' Propósito      : Establece el registro en el modelo
'------------------------------------------------------------------------------*
Private Sub GetRegistro(mRow As Integer)
  On Error GoTo GetRegistro_Error
    With mwsData
        Me.Metodo.Id = .Cells(mRow, 1).Value
        Me.Metodo.CriteriosAgrupacion = .Cells(mRow, 2).Value
        Me.Metodo.CriteriosOrdenacion = .Cells(mRow, 3).Value
        Me.Metodo.DiasAnalisis = .Cells(mRow, 4).Value
        Me.Metodo.EntidadNegocio.FechaAlta = Format(.Cells(mRow, 5).Value, "dd/mm/yyyy")
        Me.Metodo.ModalidadJuego = .Cells(mRow, 6).Value
        Me.Metodo.NumeroSorteos = .Cells(mRow, 7).Value
        Me.Metodo.Orden = .Cells(mRow, 8).Value
        Me.Metodo.Pronosticos = .Cells(mRow, 9).Value
        Me.Metodo.SentidoOrdenacion = .Cells(mRow, 10).Value
        Me.Metodo.TipoMuestra = .Cells(mRow, 10).Value
        Me.Metodo.TipoProcedimiento = .Cells(mRow, 11).Value
    End With
  On Error GoTo 0
GetRegistro___CleanExit:
    Exit Sub
GetRegistro_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    Call HandleException(ErrNumber, ErrDescription, "MetodoModel.GetRegistro", ErrSource)
    Err.Raise ErrNumber, "MetodoModel.GetRegistro", ErrDescription
End Sub '
Private Sub Class_Initialize()
    Dim llastRow As Long            'ultima fila de datos
On Error GoTo Class_Initialize_Error
    '
    '   Definimos el repositorio como la hoja "Data" de este libro
    '
    Set mwsData = ThisWorkbook.Worksheets("Metodos")
    '
    '   Localizamos la última fila de los datos
    '
    llastRow = mwsData.Cells(mwsData.Rows.Count, "A").End(xlUp).Row
    '
    '   Definimos el rango de Id para la columna A
    '
    Set mrngIds = mwsData.Range("A2:A" & llastRow)
    '
    '   Definimos la última fila del rango
    '
    mLastRow = llastRow
    '
    '   Definimos objetos del modelo
    '
    Set mCol = New Metodos
    Set mMetodo = New Metodo
    '
    '   Calculamos total de registros
    '
    mTotalRecord = mrngIds.Row - 1
On Error GoTo 0
Class_Initialize___CleanExit:
    Exit Sub
            
Class_Initialize_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    Call HandleException(ErrNumber, ErrDescription, "MetodoModel.Class_Initialize", ErrSource)
    Err.Raise ErrNumber, "MetodoModel.Class_Initialize", ErrDescription
End Sub

Private Sub Class_Terminate()
    Set mCol = Nothing
    Set mMetodo = Nothing
End Sub
'--- Métodos Públicos ---------------------------------------------------------*
'------------------------------------------------------------------------------*
' Procedimiento  :
' Fecha          :
' Propósito      :
'------------------------------------------------------------------------------*
Public Function Search() As Boolean
  On Error GoTo Search_Error
    '
    '
    Err.Raise ERR_TODO, "MetodoModel.Search", MSG_TODO
    '
  On Error GoTo 0
Search__CleanExit:
    Exit Function
            
Search_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    Call HandleException(ErrNumber, ErrDescription, "MetodoModel.Search", ErrSource)
    Err.Raise ErrNumber, "MetodoModel.Search", ErrDescription
End Function
'------------------------------------------------------------------------------*
' Procedimiento  : Add
' Fecha          : lu., 16/mar/2020 18:39:36
' Propósito      : Agregar un método a la Base de Datos
'------------------------------------------------------------------------------*
Public Function Add() As Boolean
    Dim rngFind As Range           ' Rango de búsqueda
    Dim mFila   As Integer         ' Fila de agregación
    
  On Error GoTo Add_Error
    '
    '   Inicializamos la función
    '
    Add = False
    '
    '   Obtenemos el último registro
    '
    mLastRow = mrngIds.Rows.Count
    If IsEmpty(mrngIds.Value2) Then
        '
        '   Tabla Vacia
        '
        Me.Metodo.Id = 0
        mLastRow = 1
    Else
        If mLastRow = 1 Then
            '
            '   Si el último es tambien el primero
            '
            Me.Metodo.Id = 1
            mLastRow = 2
        Else
            Me.Metodo.Id = mrngIds.Value2(mLastRow, 1)
            '
            '   Incrementamos el total de filas en una
            '
            mLastRow = mLastRow + 1
        End If
    End If
    '
    '   Determinamos el Id del nuevo registro
    '
    Me.Metodo.Id = Me.Metodo.Id + 1
    '
    '   La fila de insercción será el total de filas mas la cabecera
    '
    mFila = mLastRow + 1
    '
    '   Insertamos el registro
    '
    With mwsData
        .Cells(mFila, 1).Value = Me.Metodo.Id
        .Cells(mFila, 2).Value = Me.Metodo.CriteriosAgrupacion
        .Cells(mFila, 3).Value = Me.Metodo.CriteriosOrdenacion
        .Cells(mFila, 4).Value = Me.Metodo.DiasAnalisis
        .Cells(mFila, 5).Value = DateValue(Me.Metodo.EntidadNegocio.FechaAlta)
        .Cells(mFila, 6).Value = Me.Metodo.ModalidadJuego
        .Cells(mFila, 7).Value = Me.Metodo.NumeroSorteos
        .Cells(mFila, 8).Value = Me.Metodo.Orden
        .Cells(mFila, 9).Value = Me.Metodo.Pronosticos
        .Cells(mFila, 10).Value = Me.Metodo.SentidoOrdenacion
        .Cells(mFila, 11).Value = Me.Metodo.TipoMuestra
        .Cells(mFila, 12).Value = Me.Metodo.TipoProcedimiento
    End With
    '
    '   Total de filas
    '
    mTotalRecord = mrngIds.Row - 1
    '
    '   Devolvemos Ok
    '
    Add = True
  On Error GoTo 0
Add__CleanExit:
    Exit Function
            
Add_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    Call HandleException(ErrNumber, ErrDescription, "MetodoModel.Add", ErrSource)
    Err.Raise ErrNumber, "MetodoModel.Add", ErrDescription
End Function
'------------------------------------------------------------------------------*
' Procedimiento  : Save()
' Fecha          : lu., 16/mar/2020 18:29:23
' Propósito      : Guardar el Metodo en el repositorio
'------------------------------------------------------------------------------*
Public Function Save() As Boolean
    Dim rngFind As Range           ' Rango de búsqueda
  On Error GoTo Save_Error
    '
    '   Inicializamos el resultado
    '
    Save = False
    '
    '   Buscamos el Id del registro
    '
    Set rngFind = mrngIds.Find(What:=Metodo.Id, LookIn:=xlValues, LookAt:=xlWhole)
    If Not rngFind Is Nothing Then
        With mwsData
            .Cells(rngFind.Row, 1).Value = Me.Metodo.Id
            .Cells(rngFind.Row, 2).Value = Me.Metodo.CriteriosAgrupacion
            .Cells(rngFind.Row, 3).Value = Me.Metodo.CriteriosOrdenacion
            .Cells(rngFind.Row, 4).Value = Me.Metodo.DiasAnalisis
            .Cells(rngFind.Row, 5).Value = DateValue(Me.Metodo.EntidadNegocio.FechaAlta)
            .Cells(rngFind.Row, 6).Value = Me.Metodo.ModalidadJuego
            .Cells(rngFind.Row, 7).Value = Me.Metodo.NumeroSorteos
            .Cells(rngFind.Row, 8).Value = Me.Metodo.Orden
            .Cells(rngFind.Row, 9).Value = Me.Metodo.Pronosticos
            .Cells(rngFind.Row, 10).Value = Me.Metodo.SentidoOrdenacion
            .Cells(rngFind.Row, 11).Value = Me.Metodo.TipoMuestra
            .Cells(rngFind.Row, 12).Value = Me.Metodo.TipoProcedimiento
        End With
    Else
        '
        '   registro no encontrado
        '
        Err.Raise 100, "Metodo.Save()", "Registro no encontrado"
    End If
    '
    '
    '
    Save = True
    
  On Error GoTo 0
Save__CleanExit:
    Exit Function
            
Save_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    Call HandleException(ErrNumber, ErrDescription, "MetodoModel.Save", ErrSource)
    Err.Raise ErrNumber, "MetodoModel.Save", ErrDescription
End Function
'------------------------------------------------------------------------------*
' Procedimiento  : Del
' Fecha          : lu., 16/mar/2020 18:39:47
' Propósito      : Borrar un método del registro
'------------------------------------------------------------------------------*
Public Function Del(ByVal datId As Integer) As Boolean
    Dim rngFind    As Range
    Dim rngFila    As Range

  On Error GoTo Del_Error
    '
    '   Inicializamos la función
    '
    Del = False
    '
    '   Buscamos el registro
    '
    Set rngFind = mrngIds.Find(What:=datId, LookIn:=xlValues, LookAt:=xlWhole)
    '
    '   Si lo encontramos borramos
    '
    If Not rngFind Is Nothing Then
        '
        '   Definimos un nuevo rango a borrar
        '
        Set rngFila = rngFind.Offset(0, 11)
        '
        '   Eliminamos fila
        '
        rngFila.Delete Shift:=xlShiftUp
        '
        '   Devolvemos el Ok
        '
        Del = True
        Set Metodo = New Metodo
        Set rngFila = Nothing
    Else
        '
        '   Registro no encontrado
        '
        Err.Raise 100, "MetodoModel.Del", "No se encuentra el registro a borrar"
    End If
    Set rngFind = Nothing
     
  On Error GoTo 0
Del__CleanExit:
    Exit Function
            
Del_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    Call HandleException(ErrNumber, ErrDescription, "MetodoModel.Del", ErrSource)
    Err.Raise ErrNumber, "MetodoModel.Del", ErrDescription
End Function
'------------------------------------------------------------------------------*
' Procedimiento  : GetRecord
' Fecha          : lu., 16/mar/2020 18:40:18
' Propósito      : recupera un metodo del repositorio por el Id
'------------------------------------------------------------------------------*
Public Function GetRecord(datId As Integer) As Boolean
    Dim rngFind    As Range
  On Error GoTo GetRecord_Error
    '
    '   Inicializamos función
    '
    GetRecord = False
    '
    '   Buscamos el registro
    '
    Set rngFind = mrngIds.Find(What:=datId, LookIn:=xlValues, LookAt:=xlWhole)
    '
    '   Si lo encontramos cargamos el Objeto
    '
    If Not rngFind Is Nothing Then
        '
        '   Creamos un nuevo Objeto Metodo
        '
        Set Metodo = New Metodo
        '
        '   Cargamos atributos
        '
        GetRegistro rngFind.Row
        '
        '   Devolvemos el Ok
        '
        GetRecord = True
    Else
        '
        '   Registro no encontrado
        '
        Err.Raise -100, "MetodoModel.Del", "No se encuentra el registro a borrar"
    End If
    Set rngFind = Nothing
  
  On Error GoTo 0
GetRecord__CleanExit:
    Exit Function
            
GetRecord_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    Call HandleException(ErrNumber, ErrDescription, "MetodoModel.GetRecord", ErrSource)
    Err.Raise ErrNumber, "MetodoModel.GetRecord", ErrDescription
End Function
'------------------------------------------------------------------------------*
' Procedimiento  :
' Fecha          :
' Propósito      :
'------------------------------------------------------------------------------*
Public Function GetFirst() As Boolean
  On Error GoTo GetFirst_Error
    '
    '
    Err.Raise ERR_TODO, "MetodoModel.GetFirst", MSG_TODO
    '
  On Error GoTo 0
GetFirst__CleanExit:
    Exit Function
            
GetFirst_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    Call HandleException(ErrNumber, ErrDescription, "MetodoModel.GetFirst", ErrSource)
    Err.Raise ErrNumber, "MetodoModel.GetFirst", ErrDescription
End Function
'------------------------------------------------------------------------------*
' Procedimiento  :
' Fecha          :
' Propósito      :
'------------------------------------------------------------------------------*
Public Function GetNext() As Boolean
  On Error GoTo GetNext_Error
    '
    '
    Err.Raise ERR_TODO, "MetodoModel.GetNext", MSG_TODO
    '
  On Error GoTo 0
GetNext__CleanExit:
    Exit Function

GetNext_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    Call HandleException(ErrNumber, ErrDescription, "MetodoModel.GetNext", ErrSource)
    Err.Raise ErrNumber, "MetodoModel.GetNext", ErrDescription
End Function
'------------------------------------------------------------------------------*
' Procedimiento  :
' Fecha          :
' Propósito      :
'------------------------------------------------------------------------------*
Public Function GetPrev() As Boolean
  On Error GoTo GetLast_Error
    '
    '
    Err.Raise ERR_TODO, "MetodoModel.GetLast", MSG_TODO
    '
  On Error GoTo 0
GetLast__CleanExit:
    Exit Function
            
GetLast_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    Call HandleException(ErrNumber, ErrDescription, "MetodoModel.GetLast", ErrSource)
    Err.Raise ErrNumber, "MetodoModel.GetLast", ErrDescription
End Function
'------------------------------------------------------------------------------*
' Procedimiento  :
' Fecha          :
' Propósito      :
'------------------------------------------------------------------------------*
Public Function GetLast() As Boolean
  On Error GoTo GetLast_Error
    '
    '
    Err.Raise ERR_TODO, "MetodoModel.GetLast", MSG_TODO
    '
  On Error GoTo 0
GetLast__CleanExit:
    Exit Function
            
GetLast_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    Call HandleException(ErrNumber, ErrDescription, "MetodoModel.GetLast", ErrSource)
    Err.Raise ErrNumber, "MetodoModel.GetLast", ErrDescription
End Function

'' *===========(EOF): MetodoModel.cls

