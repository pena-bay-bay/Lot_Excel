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
Private mTotalRecord        As Integer      ' Total de registros
Private mCurrentPage        As Integer      ' Página actual
Private mTotalPages         As Integer      ' Total de páginas
'----Area de Datos
Private mwsData             As Worksheet    ' hoja de datos del repositorio
Private mrngIds             As Range        ' Rango de los datos
Private mLastRow            As Integer      ' ultima fila del rango



'--- Propiedades --------------------------------------------------------------*

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
    If mLinePage = 0 Then
        SetRngId
        If mTotalRecord Mod vDatNew = 0 Then
            mTotalPages = Fix(mTotalRecord / vDatNew)
        Else
            mTotalPages = Fix(mTotalRecord / vDatNew) + 1
        End If
    End If
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
    If mTotalRecord = 0 Then
        SetRngId
    End If
    TotalRecords = mTotalRecord
End Property


'------------------------------------------------------------------------------*
' Procedimiento  : CurrentPage
' Fecha          :
' Propósito      : Pagina de datos actual
'------------------------------------------------------------------------------*
Public Property Get CurrentPage() As Integer
    CurrentPage = mCurrentPage
End Property





'--- Métodos Privados ---------------------------------------------------------*

Private Sub Class_Initialize()
    Dim llastRow As Long            'ultima fila de datos
On Error GoTo Class_Initialize_Error
    '
    '   Definimos el repositorio como la hoja "Data" de este libro
    '
    Set mwsData = ThisWorkbook.Worksheets("Metodos")
    '
    '   Establecemos Rango de Clave
    '
    SetRngId
    '
    '   Definimos objetos del modelo
    '
    Set mCol = New Metodos
    Set mMetodo = New Metodo
    '
    '   Inicializamos lineas por pagina
    '
    mLinePage = 0
    '
    '   Inicializamos Pagina actual
    '
    mCurrentPage = 1
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



'------------------------------------------------------------------------------*
' Procedimiento  : GetRegistro
' Fecha          : lu., 16/mar/2020 18:38:45
' Propósito      : Establece el registro en el modelo
'------------------------------------------------------------------------------*
Private Sub GetRegistro(mRow As Integer)
  On Error GoTo GetRegistro_Error
    Set mMetodo = New Metodo
    With mwsData
        If Not IsEmpty(.Cells(mRow, 1).Value) Then
            mMetodo.Id = .Cells(mRow, 1).Value
            mMetodo.TipoProcedimiento = .Cells(mRow, 2).Value
            mMetodo.Pronosticos = .Cells(mRow, 3).Value
            mMetodo.ModalidadJuego = .Cells(mRow, 4).Value
            mMetodo.CriteriosOrdenacion = .Cells(mRow, 5).Value
            mMetodo.SentidoOrdenacion = .Cells(mRow, 6).Value
            mMetodo.CriteriosAgrupacion = .Cells(mRow, 7).Value
            mMetodo.TipoMuestra = .Cells(mRow, 8).Value
            mMetodo.NumeroSorteos = .Cells(mRow, 9).Value
            mMetodo.DiasAnalisis = .Cells(mRow, 10).Value
            mMetodo.Filtros.Parse .Cells(mRow, 11).Value
            mMetodo.EntidadNegocio.FechaAlta = Format(.Cells(mRow, 12).Value, "dd/mm/yyyy hh:nn:ss")
            mMetodo.EntidadNegocio.FechaModificacion = Format(.Cells(mRow, 13).Value, "dd/mm/yyyy hh:nn:ss")
        Else
            mMetodo.Id = 0
            'Err.Raise 100, "Metodo.GetRegistro()", "Registro no encontrado"
        End If
    End With
  On Error GoTo 0
GetRegistro___CleanExit:
    Exit Sub
GetRegistro_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    Call HandleException(ErrNumber, ErrDescription, "MetodoModel.GetRegistro", ErrSource)
    Err.Raise ErrNumber, "MetodoModel.GetRegistro", ErrDescription
End Sub



'------------------------------------------------------------------------------*
' Procedimiento  : SetRngId
' Fecha          : lu., 16/mar/2020 18:38:45
' Propósito      : Establece el rango Clave de la tabla y última fila
'------------------------------------------------------------------------------*
Private Sub SetRngId()
    Dim llastRow As Long            'ultima fila de datos
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
    '   Calculamos el total de registros
    '
    If IsEmpty(mwsData.Cells(2, 1).Value) Then
        mTotalRecord = 0
    Else
        mTotalRecord = mLastRow - 1
    End If
End Sub




'--- Métodos Públicos ---------------------------------------------------------*
'------------------------------------------------------------------------------*
' Procedimiento  : Search
' Fecha          :
' Propósito      : Realizar una busqueda de información según filtros
'                  para los métodos es la página
'------------------------------------------------------------------------------*
Public Function Search(datPage As Integer) As Boolean
    Dim mReg    As Integer
    Dim i       As Integer
    Dim mMtdo   As Metodo
  On Error GoTo Search_Error
    '
    '   Comprobar que hay registros en la base de datos
    '
    If mTotalRecord = 0 Then
        Search = False
        Exit Function
    End If
    '
    '   Comprobamos que la pagina solicitada está en el rango de páginas
    '
    If datPage > TotalPages Then
        datPage = TotalPages
    End If
    '
    '   Si es la primera página
    '
    If datPage = 1 Then
        mReg = 2
    Else
        '
        '   Calculamos la primera fila de la pagina solicitada
        '
        mReg = ((datPage - 1) * LinePerPage) + 2
    End If
    '
    '   Inicializamos la colección
    '
    mCol.Clear
    '
    '   Carga la colección Metodos con la página seleccionada
    '
    For i = 1 To LinePerPage
        '
        '   Cargamos el
        '
        If mReg <= mLastRow Then
            GetRegistro mReg
            '
            '
            Set mMtdo = Me.Metodo
            If Me.Metodo.Id > 0 Then
                '
                '   Guardamos el metodo
                '
                mCol.Add mMtdo
                mReg = mReg + 1
                Set mMtdo = Nothing
            End If
        Else
            Exit For
        End If
    Next i
    '
    '   Si encontramos registros
    '
    If mCol.Count > 0 Then
        mCurrentPage = datPage
        Search = True
    Else
        Search = False
    End If
 
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
' Procedimiento  : SetNew
' Fecha          : ju., 16/abr/2020 23:40:45
' Propósito      : Define un metodo vacio en el modelo
'------------------------------------------------------------------------------*
Public Function SetNew() As Boolean
    Set mMetodo = New Metodo
    mMetodo.Id = TotalRecords + 1
End Function



'------------------------------------------------------------------------------*
' Procedimiento  : Add
' Fecha          : lu., 16/mar/2020 18:39:36
' Propósito      : Agregar un método a la Base de Datos
'------------------------------------------------------------------------------*
Public Function Add() As Boolean
    Dim rngFind     As Range           ' Rango de búsqueda
    Dim mFila       As Integer         ' Fila de agregación
  On Error GoTo Add_Error
    '
    '   Inicializamos la función
    '
    Add = False
    '
    '   Obtenemos el último registro
    '
    SetRngId
    '
    '
    '
    If IsEmpty(mrngIds.Value2) Then
        '
        '   Tabla Vacia
        '
        Me.Metodo.Id = 0
        mFila = mLastRow
    ElseIf IsNumeric(mrngIds.Value2) Then
        '
        '   Si el último es tambien el primero
        '
        Me.Metodo.Id = mrngIds.Value2
        mFila = mLastRow + 1
    Else
        '
        '   Obtenemos el último Id
        '
        Me.Metodo.Id = mrngIds.Value2(mLastRow - 1, 1)
        '
        '   La fila de insercción será el total de filas mas la cabecera
        '
        mFila = mLastRow + 1
    End If
    '
    '   Determinamos el Id del nuevo registro
    '
    Me.Metodo.Id = Me.Metodo.Id + 1
    '
    '   Insertamos el registro
    '
    With mwsData
        .Cells(mFila, 1).Value = Me.Metodo.Id
        .Cells(mFila, 2).Value = Me.Metodo.TipoProcedimiento
        .Cells(mFila, 3).Value = Me.Metodo.Pronosticos
        .Cells(mFila, 4).Value = Me.Metodo.ModalidadJuego
        .Cells(mFila, 5).Value = Me.Metodo.CriteriosOrdenacion
        .Cells(mFila, 6).Value = Me.Metodo.SentidoOrdenacion
        .Cells(mFila, 7).Value = Me.Metodo.CriteriosAgrupacion
        .Cells(mFila, 8).Value = Me.Metodo.TipoMuestra
        .Cells(mFila, 9).Value = Me.Metodo.NumeroSorteos
        .Cells(mFila, 10).Value = Me.Metodo.DiasAnalisis
        .Cells(mFila, 11).Value = Me.Metodo.Filtros.ToString
        .Cells(mFila, 12).Value = DateValue(Me.Metodo.EntidadNegocio.FechaAlta)
        .Cells(mFila, 13).Value = DateValue(Me.Metodo.EntidadNegocio.FechaModificacion)
    End With
    '
    '   Restablecemos el rango de la base de datos
    '
    SetRngId
    '
    '   Calculamos el total de páginas
    '
    If mLinePage <> 0 Then
        If mTotalRecord Mod mLinePage = 0 Then
            mTotalPages = Fix(mTotalRecord / mLinePage)
        Else
            mTotalPages = Fix(mTotalRecord / mLinePage) + 1
        End If
    End If
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
    Set rngFind = mrngIds.Find(What:=Metodo.Id, LookIn:=xlValues, lookat:=xlWhole)
    If Not rngFind Is Nothing Then
        Me.Metodo.EntidadNegocio.FechaModificacion = Date
        With mwsData
            .Cells(rngFind.Row, 1).Value = Me.Metodo.Id
            .Cells(rngFind.Row, 2).Value = Me.Metodo.TipoProcedimiento
            .Cells(rngFind.Row, 3).Value = Me.Metodo.Pronosticos
            .Cells(rngFind.Row, 4).Value = Me.Metodo.ModalidadJuego
            .Cells(rngFind.Row, 5).Value = Me.Metodo.CriteriosOrdenacion
            .Cells(rngFind.Row, 6).Value = Me.Metodo.SentidoOrdenacion
            .Cells(rngFind.Row, 7).Value = Me.Metodo.CriteriosAgrupacion
            .Cells(rngFind.Row, 8).Value = Me.Metodo.TipoMuestra
            .Cells(rngFind.Row, 9).Value = Me.Metodo.NumeroSorteos
            .Cells(rngFind.Row, 10).Value = Me.Metodo.DiasAnalisis
            .Cells(rngFind.Row, 11).Value = Me.Metodo.Filtros.ToString
            .Cells(rngFind.Row, 12).Value = DateValue(Me.Metodo.EntidadNegocio.FechaAlta)
            .Cells(rngFind.Row, 13).Value = DateValue(Me.Metodo.EntidadNegocio.FechaModificacion)
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
    Set rngFind = mrngIds.Find(What:=datId, LookIn:=xlValues, lookat:=xlWhole)
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
    '
    '   establecemos total de registros
    '
     SetRngId
    '
    '   Calculamos el total de páginas
    '
    If mLinePage <> 0 Then
        If mTotalRecord Mod mLinePage = 0 Then
            mTotalPages = Fix(mTotalRecord / mLinePage)
        Else
            mTotalPages = Fix(mTotalRecord / mLinePage) + 1
        End If
    End If

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
    '   Establece el rango de la base de datos
    '
    If mrngIds Is Nothing Then
        '   Establecemos Rango de Clave
        '
        SetRngId
    End If
    '
    '   Buscamos el registro
    '
    Set rngFind = mrngIds.Find(What:=datId, LookIn:=xlValues, lookat:=xlWhole)
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
' Procedimiento  : GetFirst
' Fecha          : ju., 19/mar/2020 18:22:04
' Propósito      : Obtiene el primer elemento
'------------------------------------------------------------------------------*
Public Function GetFirst() As Boolean
    Dim rngFind As Range
  On Error GoTo GetFirst_Error
    '
    '   Inicializamos la función
    '
    GetFirst = False
    '
    '   Localizamos el primer registro
    '
    Set rngFind = mrngIds.Rows(1)
    '
    '   Si encuentra el primer registro
    '
    If Not rngFind Is Nothing Then
        '
        '
        '
        GetRegistro rngFind.Row
        GetFirst = True
    End If
    '
    '   Eliminamos el objeto
    '
    Set rngFind = Nothing
  
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
' Procedimiento  : GetNext
' Fecha          : ju., 19/mar/2020 18:26:28
' Propósito      : Obtiene el siguiente registro al del Id
'------------------------------------------------------------------------------*
Public Function GetNext(ByVal intId As Integer) As Boolean
    Dim rngFind As Range
    Dim mRow    As Integer
  On Error GoTo GetNext_Error
    '
    '   Inicializamos la función
    '
    GetNext = False
    '
    '   Buscamos el Id y si lo encontramos calculamos la siguiente fila
    '
    Set rngFind = mrngIds.Find(What:=intId, LookIn:=xlValues, lookat:=xlWhole)
    If Not rngFind Is Nothing Then
        '
        '   Comprobamos si es el último registro
        '
        If rngFind.Row < mLastRow Then
            mRow = rngFind.Row + 1
            GetRegistro mRow
            GetNext = True
        End If
    End If
    Set rngFind = Nothing
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
' Procedimiento  : GetPrev
' Fecha          : ju., 19/mar/2020 18:32:32
' Propósito      : Obtener el anterior registro al suministrado
'------------------------------------------------------------------------------*
Public Function GetPrev(ByVal intId As Integer) As Boolean
    Dim rngFind As Range
    Dim mRow    As Integer
  On Error GoTo GetLast_Error
    '
    '   Inicilizamos la función
    '
    GetPrev = False
    '
    '   Buscamos el registro suministrado
    '
    Set rngFind = mrngIds.Find(What:=intId, LookIn:=xlValues, lookat:=xlWhole)
    If Not rngFind Is Nothing Then
        '
        '   Comprobamos si es el primer registro
        '
        If rngFind.Row > 2 Then
            mRow = rngFind.Row - 1
            GetRegistro mRow
            GetPrev = True
        End If
    End If
    Set rngFind = Nothing
  
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
' Procedimiento  : GetLast
' Fecha          : ju., 19/mar/2020 19:40:46
' Propósito      : Devuelve el último elemento de la tabla
'------------------------------------------------------------------------------*
Public Function GetLast() As Boolean
  On Error GoTo GetLast_Error
    '
    '   Inicializamos la función
    '
    GetLast = False
    '
    '   Comprobamos que existan registros
    '
    If mLastRow > 0 Then
        '
        '   Obtenemos el último registro
        '
        GetRegistro mLastRow
        GetLast = True
    End If
  
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

