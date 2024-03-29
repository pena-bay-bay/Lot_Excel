VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SorteoModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' *============================================================================*
' *
' *     Fichero    : SorteoModel.cls
' *
' *     Autor      : Carlos Almela Baeza
' *     Creaci�n   : ju., 09/may/2019 00:04:56
' *     Versi�n    : 1.0
' *     Prop�sito  :
' *
' *============================================================================*
Option Explicit
Option Base 0
'
'--- Variables Privadas -------------------------------------------------------*
Private mId             As Integer      ' Identificador �nico del sorteo
Private mJuego          As String       ' Juego al que pertenece el sorteo
Private mNumeroSorteo   As String       ' Numero de sorteo del juego
Private mFecha          As String       ' Fecha del sorteo
Private mDia            As String       ' Dia de la semana
Private mSemana         As Integer      ' Semana del a�o
Private mOrden          As String       ' Orden de aparicion de los datos
Private mCombinacion    As String       ' Combinaci�n ganadora
Private mN1             As Integer      ' Numero 1
Private mN2             As Integer      ' Numero 2
Private mN3             As Integer      ' Numero 3
Private mN4             As Integer      ' Numero 4
Private mN5             As Integer      ' Numero 5
Private mN6             As Integer      ' Numero 6
Private mEstrellas      As String       ' matriz de numeros estrella
Private mE1             As Integer      ' Estrella 1
Private mE2             As Integer      ' Estrella 2
Private mComplementario As Integer      ' Numero complementario
Private mReintegro      As Integer      ' numero del reintegro
Private mPremios        As Variant      ' Matriz de importes de premio
'----Area Busqueda
Private mPeriodo        As Periodo      ' Periodo de Busqueda
'----Area Paginado
Private mTotalPaginas   As Integer      ' Total de paginas de la base de datos
Private mLineasPagina   As Integer      ' Numero de lineas por p�ginas
Private mTotRegistros   As Integer      ' Total registros de la base de datos
Private mPagActual      As Integer      ' P�gina Actual
Private mMatrizResultados As Variant    ' Matriz de resultados de Busqueda
'----Area de Datos
Private mwsData         As Worksheet    ' hoja de datos del repositorio
Private mrngIds         As Range        ' Rango de los datos
Private mLastRow        As Integer      ' ultima fila del rango
Private mFirstRow       As Integer      ' ultima fila del rango
'--- Propiedades --------------------------------------------------------------*
'
'       Propiedades de Sorteo
'
Public Property Get IdSelected() As Integer
    IdSelected = mId
End Property

Public Property Let IdSelected(ByVal vNewValue As Integer)
    mId = vNewValue
End Property
Public Property Get Juego() As String
    Juego = mJuego
End Property

Public Property Let Juego(ByVal vNewValue As String)
    mJuego = vNewValue
    If IsEmpty(mPremios) Then
        SetPremios Empty
    End If
    If IsEmpty(mCombinacion) Then
        SetCombinacion
    End If
End Property

Public Property Get FechaSorteo() As String
    FechaSorteo = mFecha
End Property

Public Property Let FechaSorteo(ByVal vNewValue As String)
    mFecha = vNewValue
End Property
Public Property Get NumSorteo() As Variant
    NumSorteo = mNumeroSorteo
End Property

Public Property Let NumSorteo(ByVal vNewValue As Variant)
    mNumeroSorteo = vNewValue
End Property

Public Property Get DiaSemana() As String
    DiaSemana = mDia
End Property

Public Property Let DiaSemana(ByVal vNewValue As String)
    mDia = vNewValue
End Property

Public Property Get Semana() As Integer
    Semana = mSemana
End Property

Public Property Let Semana(ByVal vNewValue As Integer)
    mSemana = vNewValue
End Property

Public Property Get N1() As Integer
    N1 = mN1
End Property
Public Property Let N1(ByVal vNewValue As Integer)
    mN1 = vNewValue
End Property
Public Property Get N2() As Integer
    N2 = mN2
End Property
Public Property Let N2(ByVal vNewValue As Integer)
    mN2 = vNewValue
End Property
Public Property Get N3() As Integer
    N3 = mN3
End Property
Public Property Let N3(ByVal vNewValue As Integer)
    mN3 = vNewValue
End Property
Public Property Get N4() As Integer
    N4 = mN4
End Property
Public Property Let N4(ByVal vNewValue As Integer)
    mN4 = vNewValue
End Property
Public Property Get N5() As Integer
    N5 = mN5
End Property
Public Property Let N5(ByVal vNewValue As Integer)
    mN5 = vNewValue
End Property
Public Property Get N6() As Integer
    N6 = mN6
End Property
Public Property Let N6(ByVal vNewValue As Integer)
    mN6 = vNewValue
End Property
Public Property Get E1() As Integer
    E1 = mE1
End Property
Public Property Let E1(ByVal vNewValue As Integer)
    mE1 = vNewValue
End Property
Public Property Get E2() As Integer
    E2 = mE2
End Property
Public Property Let E2(ByVal vNewValue As Integer)
    mE2 = vNewValue
End Property
Public Property Get CombinacionGanadora() As String
    SetCombinacion
    CombinacionGanadora = mCombinacion
End Property
Public Property Let CombinacionGanadora(ByVal vNewValue As String)
    Dim mVar As Variant
    mCombinacion = vNewValue
    mVar = Split(mCombinacion, "-")
    If UBound(mVar) = 5 Then
        mN1 = mVar(0)
        mN2 = mVar(1)
        mN3 = mVar(2)
        mN4 = mVar(3)
        mN5 = mVar(4)
        mN6 = mVar(5)
    End If
    If UBound(mVar) = 4 Then
        mN1 = mVar(0)
        mN2 = mVar(1)
        mN3 = mVar(2)
        mN4 = mVar(3)
        mN5 = mVar(4)
    End If
End Property
Public Property Get Complementario() As Integer
    Complementario = mComplementario
End Property
Public Property Let Complementario(ByVal vNewValue As Integer)
    mComplementario = vNewValue
End Property
Public Property Get MatrizPremios() As Variant
    MatrizPremios = mPremios
End Property
Public Property Let MatrizPremios(ByVal vNewValue As Variant)
    mPremios = vNewValue
End Property
Public Property Get Reintegro() As Integer
    Reintegro = mReintegro
End Property
Public Property Let Reintegro(ByVal vNewValue As Integer)
    mReintegro = vNewValue
End Property
Public Property Get Estrellas() As String
    If IsEmpty(mEstrellas) Then
        mEstrellas = "'" & mE1 & "-" & mE2
    End If
    Estrellas = mEstrellas
End Property
Public Property Let Estrellas(ByVal vNewValue As String)
    Dim mVar As Variant
    
    mEstrellas = Replace(vNewValue, "'", "")
    mVar = Split(mEstrellas, "-")
    If UBound(mVar) = 1 Then
        mE1 = mVar(0)
        mE2 = mVar(1)
    End If
End Property
Public Property Get OrdenAparicion() As String
    OrdenAparicion = mOrden
End Property
Public Property Let OrdenAparicion(ByVal vNewValue As String)
    mOrden = vNewValue
End Property
'
'   Campos de Filtrado
'
Public Property Get PeriodoTexto() As String
    PeriodoTexto = mPeriodo.Texto
End Property
Public Property Let PeriodoTexto(ByVal vNewValue As String)
    If mPeriodo Is Nothing Then
        Set mPeriodo = New Periodo
    End If
    If Not IsEmpty(vNewValue) Then
        mPeriodo.Texto = vNewValue
    End If
End Property

Public Property Get FechaInicio() As String
    FechaInicio = Format(mPeriodo.FechaInicial, "dd/mm/yyy")
End Property

Public Property Let FechaInicio(ByVal vNewValue As String)
    If mPeriodo Is Nothing Then
        Set mPeriodo = New Periodo
    End If
    If IsDate(vNewValue) Then
        mPeriodo.FechaInicial = CDate(vNewValue)
    ElseIf Len(vNewValue) = 0 Then
        mPeriodo.Tipo_Fecha = -1
    End If
End Property
Public Property Get FechaFin() As String
    FechaFin = Format(mPeriodo.FechaFinal, "dd/mm/yyy")
End Property
Public Property Let FechaFin(ByVal vNewValue As String)
    If mPeriodo Is Nothing Then
        Set mPeriodo = New Periodo
    End If
    If IsDate(vNewValue) Then
        mPeriodo.FechaFinal = CDate(vNewValue)
    ElseIf Len(vNewValue) = 0 Then
        mPeriodo.Tipo_Fecha = -1
    End If
End Property
Public Property Get LineasPorPagina() As Integer
    LineasPorPagina = mLineasPagina
End Property
Public Property Let LineasPorPagina(ByVal vNewValue As Integer)
    mLineasPagina = vNewValue
End Property
Public Property Get PaginaActual() As Integer
    PaginaActual = mPagActual
End Property
Public Property Let PaginaActual(ByVal vNewValue As Integer)
    mPagActual = vNewValue
End Property
Public Property Get TotalPaginas() As Integer
    TotalPaginas = mTotalPaginas
End Property
Public Property Let TotalPaginas(ByVal vNewValue As Integer)
    mTotalPaginas = vNewValue
End Property
Public Property Get TotalRegistros() As Integer
    TotalRegistros = mTotRegistros
End Property
Public Property Let TotalRegistros(ByVal vNewValue As Integer)
    mTotRegistros = vNewValue
End Property
Public Property Get ResultadosSearch() As Variant
    ResultadosSearch = mMatrizResultados
End Property
'
'
'--- M�todos Privados ---------------------------------------------------------*
Private Sub Class_Initialize()
    Dim llastRow As Long            'ultima fila de datos
    '
    '   Definimos el repositorio como la hoja "Data" de este libro
    '
    Set mwsData = ThisWorkbook.Worksheets("Data")
    '
    '   Localizamos la �ltima fila de los datos
    '
    llastRow = mwsData.Cells(mwsData.Rows.Count, "A").End(xlUp).Row
    '
    '   Definimos el rango de Id para la columna A
    '
    Set mrngIds = mwsData.Range("A2:A" & llastRow)
    '
    '   Definimos la �ltima fila del rango
    '
    mLastRow = llastRow
    '
    '   Definimos Matriz de Resultados
    '
    ReDim mMatrizResultados(7, 7) As String
    '
    '   Inicializamos variables
    '
    Set mPeriodo = New Periodo
End Sub
Private Sub Class_Terminate()
    Set mwsData = Nothing
    Set mrngIds = Nothing
End Sub
'------------------------------------------------------------------------------*
' Procedimiento  : SetPremios
' Fecha          : s�., 01/jun/2019 22:12:44
' Prop�sito      : Establece los premios por defecto de un sorteo
'------------------------------------------------------------------------------*
Private Sub SetPremios(vNewData As String)
    If Len(vNewData) = 0 Then
        Select Case Juego
            Case LT_EUROMILLON
                mPremios = Split(LP_PREMIOS_EURO, ";")
            Case LT_GORDO
                mPremios = Split(LP_PREMIOS_GORDO, ";")
            Case LT_BONOLOTO
                mPremios = Split(LP_PREMIOS_BONO, ";")
            Case LT_PRIMITIVA
                mPremios = Split(LP_PREMIOS_PRIMI, ";")
        End Select
    Else
        mPremios = Split(vNewData, ";")
    End If
End Sub
'------------------------------------------------------------------------------*
' Procedimiento  : SetCombinacion
' Fecha          : s�., 01/jun/2019 22:12:44
' Prop�sito      : Establece el texto de la combinacion
'------------------------------------------------------------------------------*
Private Sub SetCombinacion()
    Select Case Juego
        Case LT_EUROMILLON
            mCombinacion = mN1 & "-" & mN2 & "-" & mN3 & "-" _
                         & mN4 & "-" & mN5
            mEstrellas = mE1 & "-" & mE2
        Case LT_GORDO
            mCombinacion = mN1 & "-" & mN2 & "-" & mN3 & "-" _
                         & mN4 & "-" & mN5
        Case LT_BONOLOTO, LT_PRIMITIVA
            mCombinacion = mN1 & "-" & mN2 & "-" & mN3 & "-" _
                         & mN4 & "-" & mN5 & "-" & mN6
    End Select
    
End Sub
'------------------------------------------------------------------------------*
' Procedimiento  : GetRegistro
' Fecha          : s�., 01/jun/2019 22:12:44
' Prop�sito      : Establece el registro en el modelo
'------------------------------------------------------------------------------*
Private Sub GetRegistro(mRow As Integer)
    With mwsData
        Me.IdSelected = .Cells(mRow, 1).value
        Me.Juego = .Cells(mRow, 2).value
        Me.NumSorteo = .Cells(mRow, 3).value
        Me.FechaSorteo = Format(.Cells(mRow, 4).value, "dd/mm/yyyy")
        Me.DiaSemana = .Cells(mRow, 5).value
        Me.Semana = .Cells(mRow, 6).value
        Me.OrdenAparicion = .Cells(mRow, 7).value
        Me.CombinacionGanadora = .Cells(mRow, 8).value
        Me.Complementario = .Cells(mRow, 9).value
        Me.Estrellas = .Cells(mRow, 10).value
        Me.Reintegro = .Cells(mRow, 11).value
        SetPremios .Cells(mRow, 12).value
    End With
End Sub

'--- M�todos P�blicos ---------------------------------------------------------*
'------------------------------------------------------------------------------*
' Procedimiento  : SearchSorteos
' Fecha          : vi., 28/jun/2019 19:26:59
' Prop�sito      : Busca los sorteos con el criterio seleccionado
'------------------------------------------------------------------------------*
Public Function SearchSorteos() As Boolean
    Dim mJuegoSearch     As String       ' Juego de busqueda
    Dim mDateSearch      As Date         ' Fecha de busqueda
    Dim mResult          As Collection   ' Colecci�n de resultados
    Dim mCelda           As Range        ' Celda de busqueda
    Dim mDatFecha        As Date         ' Fecha del registro
    Dim mDatJuego        As String       ' Juego del registro
    Dim mId              As Variant      ' Id de consulta
    Dim mNreg            As Integer      ' Registro actual
    Dim mIniReg          As Integer      ' Registro desde
    Dim mFinReg          As Integer      ' Registro hasta
    Dim i                As Integer      ' contador
    
  On Error GoTo SearchSorteos_Error
    '
    '   Inicializamos la funci�n No se ha encontrado nada
    '
    SearchSorteos = False
    '
    '   Inicilizamos contadores
    '
    mJuegoSearch = Juego
    mDateSearch = IIf(IsDate(FechaSorteo), FechaSorteo, 0)
    mTotRegistros = 0
    mTotalPaginas = 0
    '
    '   Creamos una colecci�n de resultados
    '
    Set mResult = New Collection
    '
    '   Para cada registro en la tabla
    '
    For Each mCelda In mrngIds.Rows
        '
        '   Establecemos la fecha y el juego a comparar
        '
        mDatFecha = CDate(mCelda.Offset(0, 3).value)
        mDatJuego = mCelda.Offset(0, 1).value
        '
        '   Evaluar el filtro Juego, Fecha y Periodo
        '   o periodo y juego
        '   o solo juego
        '
        If ((mDatJuego = mJuegoSearch) Or (Len(mJuegoSearch) = 0)) And _
        ((mDatFecha = mDateSearch) Or (mDateSearch = 0)) And _
        ((mDatFecha >= mPeriodo.FechaInicial And mDatFecha <= mPeriodo.FechaFinal) Or (mPeriodo.Tipo_Fecha = -1)) Then
            '
            '   Si cumple las condiciones del filtro, guardamos el Id del registro
            '
            mResult.Add Item:=mCelda.value, Key:=CStr(mCelda.value)
        End If
    Next mCelda
    '
    '   Si hemos encontrado coincidencias
    '
    If mResult.Count > 0 Then
        SearchSorteos = True
        '
        '   Calculamos total registros y total p�ginas
        '
        mTotRegistros = mResult.Count
        '
        '   Calculamos el total de p�ginas
        '
        If (mTotRegistros Mod mLineasPagina) = 0 Then
            mTotalPaginas = Int(mTotRegistros / mLineasPagina)
        Else
            mTotalPaginas = Int(mTotRegistros / mLineasPagina) + 1
        End If
        '
        '   Analizamos la pagina solicitada
        '
        Select Case True
            Case mTotalPaginas = 1:             ' Solo tenemos una p�gina
               mIniReg = 1
               mFinReg = mTotRegistros
                
            Case mTotalPaginas > 1 _
             And mPagActual > mTotalPaginas:    ' Mas de una p�gina y la p�gina de busqueda superior a la actual
               mIniReg = (mTotalPaginas - 1) * mLineasPagina
               mFinReg = mTotRegistros
            
            Case mTotalPaginas > 1 _
             And mPagActual = mTotalPaginas:    ' Mas de una p�gina y la p�gina de busqueda es la �ltima
               mIniReg = (mTotalPaginas - 1) * mLineasPagina
               mFinReg = mTotRegistros
            
            Case Else                           ' Pagina de consulta
               mIniReg = (mPagActual - 1) * mLineasPagina + 1
               mFinReg = mIniReg + mLineasPagina - 1
        End Select
        '
        '   Redimensionamos la matriz
        '
        mNreg = mFinReg - mIniReg
        ReDim mMatrizResultados(mNreg, 7) As String
        '
        '    Numero de registros editados
        '
        mNreg = 0
        '
        '   Para ca
        '
        For i = mIniReg To mFinReg
            '
            '   Obtenemos el Id del elemento iesimo
            '
            mId = mResult.Item(i)
            '
            '  Obtenemos el registro
            '
            GetSorteoRecord CInt(mId)
            '
            ' Guarda datos en Matriz
            '
            mMatrizResultados(mNreg, 0) = Me.IdSelected
            mMatrizResultados(mNreg, 1) = Me.Juego
            mMatrizResultados(mNreg, 2) = Me.FechaSorteo
            mMatrizResultados(mNreg, 3) = Me.DiaSemana
            mMatrizResultados(mNreg, 4) = Me.Semana
            Select Case Me.Juego
                Case LT_BONOLOTO, LT_PRIMITIVA
                    mMatrizResultados(mNreg, 5) = Me.CombinacionGanadora & _
                                                  " C-" & Me.Complementario
                    mMatrizResultados(mNreg, 6) = Me.Reintegro
                Case LT_EUROMILLON
                    mMatrizResultados(mNreg, 5) = Me.CombinacionGanadora & _
                                                  " E-" & Me.Estrellas
                Case LT_GORDO
                    mMatrizResultados(mNreg, 5) = Me.CombinacionGanadora
                    mMatrizResultados(mNreg, 6) = Me.Reintegro
            End Select
            '
            '  Incrementa el numero de registros encontratos
            '
            mNreg = mNreg + 1
        Next i
    End If                                  ' If mResult.Count > 0 Then
SearchSorteos__CleanExit:
    Exit Function
            
SearchSorteos_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    '   Audita el error
    Call HandleException(ErrNumber, ErrDescription, "SorteoModel.SearchSorteos", ErrSource)
    '   Lanza el Error
    Err.Raise ErrNumber, "SorteoModel.SearchSorteos", ErrDescription
End Function
'------------------------------------------------------------------------------*
' Procedimiento  : GuardarSorteoRecord
' Fecha          : do., 12/may/2019 21:11:33
' Prop�sito      : Guarda el sorteo en la base de datos
'------------------------------------------------------------------------------*
Public Function GuardarSorteoRecord(ByVal oSorteoModel As SorteoModel) As Boolean
    Dim rngFind As Range           ' Rango de b�squeda
    Dim mTmp    As String          ' Cadena de texto temporal
    
  On Error GoTo GuardarSorteoRecord_Error
    '
    '   TODO Guardar Premio
    '
    GuardarSorteoRecord = False
    Set rngFind = mrngIds.Find(What:=oSorteoModel.IdSelected, LookIn:=xlValues, LookAt:=xlWhole)
    If Not rngFind Is Nothing Then
        With mwsData
            .Cells(rngFind.Row, 1).value = oSorteoModel.IdSelected
            .Cells(rngFind.Row, 2).value = oSorteoModel.Juego
            .Cells(rngFind.Row, 3).value = oSorteoModel.NumSorteo
            .Cells(rngFind.Row, 4).value = DateValue(oSorteoModel.FechaSorteo)
            .Cells(rngFind.Row, 5).value = oSorteoModel.DiaSemana
            .Cells(rngFind.Row, 6).value = oSorteoModel.Semana
            .Cells(rngFind.Row, 7).value = oSorteoModel.OrdenAparicion
            .Cells(rngFind.Row, 8).value = oSorteoModel.CombinacionGanadora
            .Cells(rngFind.Row, 9).value = oSorteoModel.Complementario
            .Cells(rngFind.Row, 10).value = "'" & oSorteoModel.Estrellas
            .Cells(rngFind.Row, 11).value = oSorteoModel.Reintegro
            .Cells(rngFind.Row, 12).value = oSorteoModel.GetPremios
        End With
        GuardarSorteoRecord = True
    Else
        mLastRow = mrngIds.Rows.Count + 2
        With mwsData
            .Cells(mLastRow, 1).value = oSorteoModel.IdSelected
            .Cells(mLastRow, 2).value = oSorteoModel.Juego
            .Cells(mLastRow, 3).value = oSorteoModel.NumSorteo
            .Cells(mLastRow, 4).value = CDate(oSorteoModel.FechaSorteo)
            .Cells(mLastRow, 5).value = oSorteoModel.DiaSemana
            .Cells(mLastRow, 6).value = oSorteoModel.Semana
            .Cells(mLastRow, 7).value = oSorteoModel.OrdenAparicion
            .Cells(mLastRow, 8).value = oSorteoModel.CombinacionGanadora
            .Cells(mLastRow, 9).value = oSorteoModel.Complementario
            .Cells(mLastRow, 10).value = "'" & oSorteoModel.Estrellas
            .Cells(mLastRow, 11).value = oSorteoModel.Reintegro
            .Cells(mLastRow, 12).value = oSorteoModel.GetPremios
        End With
        Set mrngIds = mwsData.Range("A2:A" & mLastRow)
        GuardarSorteoRecord = True
    End If
    Set rngFind = Nothing

  On Error GoTo 0
GuardarSorteoRecord__CleanExit:
    Exit Function
            
GuardarSorteoRecord_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    '   Audita el error
    Call HandleException(ErrNumber, ErrDescription, "SorteoModel.GuardarSorteoRecord", ErrSource)
    '   Lanza el Error
    Err.Raise ErrNumber, "SorteoModel.GuardarSorteoRecord", ErrDescription
End Function
'------------------------------------------------------------------------------*
' Procedimiento  : NuevoSorteoRecord
' Fecha          : do., 12/may/2019 21:11:33
' Prop�sito      : Crea un nuevo registro
'------------------------------------------------------------------------------*
Public Function NuevoSorteoRecord() As Boolean
 
  On Error GoTo NuevoSorteoRecord_Error
    NuevoSorteoRecord = False
    '
    '   Obtenemos el �ltimo registro
    '
    mLastRow = mrngIds.Rows.Count
    Me.IdSelected = mrngIds.Value2(mLastRow, 1)
    '
    '   Incrementamos el total de filas en una
    '
    mLastRow = mLastRow + 1
    '
    '   Determinamos el Id del nuevo registro
    '
    Me.IdSelected = Me.IdSelected + 1
    '
    '   Inicializamos datos del modelo
    '
    Me.Juego = Empty
    Me.FechaSorteo = Empty
    Me.DiaSemana = Empty
    Me.Semana = Empty
    Me.N1 = Empty
    Me.N2 = Empty
    Me.N3 = Empty
    Me.N4 = Empty
    Me.N5 = Empty
    Me.N6 = Empty
    Me.NumSorteo = Empty
    Me.OrdenAparicion = Empty
    Me.Reintegro = Empty
    Me.E1 = Empty
    Me.E2 = Empty
    NuevoSorteoRecord = True

On Error GoTo 0
NuevoSorteoRecord__CleanExit:
    Exit Function
            
NuevoSorteoRecord_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    '   Audita el error
    Call HandleException(ErrNumber, ErrDescription, "SorteoModel.NuevoSorteoRecord", ErrSource)
    '   Lanza el Error
    Err.Raise ErrNumber, "SorteoModel.NuevoSorteoRecord", ErrDescription
End Function
'------------------------------------------------------------------------------*
' Procedimiento  : EliminarSorteoRecord
' Fecha          : do., 12/may/2019 21:11:33
' Prop�sito      : Elimina un registro por el Id
'------------------------------------------------------------------------------*
Public Function EliminarSorteoRecord(ByVal intId As Integer) As Boolean
     Dim rngFind    As Range
     Dim rngFila    As Range
     
  On Error GoTo EliminarSorteoRecord_Error
    EliminarSorteoRecord = False
    
    Set rngFind = mrngIds.Find(What:=intId, LookIn:=xlValues, LookAt:=xlWhole)
    If Not rngFind Is Nothing Then
        Set rngFila = rngFind.Offset(0, 11)
        rngFila.Delete Shift:=xlShiftUp
        EliminarSorteoRecord = True
        Set rngFila = Nothing
     End If
     Set rngFind = Nothing
  
  On Error GoTo 0
EliminarSorteoRecord__CleanExit:
    Exit Function
            
EliminarSorteoRecord_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    '   Audita el error
    Call HandleException(ErrNumber, ErrDescription, "SorteoModel.EliminarSorteoRecord", ErrSource)
    '   Lanza el Error
    Err.Raise ErrNumber, "SorteoModel.EliminarSorteoRecord", ErrDescription
End Function

'------------------------------------------------------------------------------*
' Procedimiento  : GetSorteoRecord
' Fecha          : lu., 27/may/2019 23:46:52
' Prop�sito      : Obtiene el registro con el Id
'------------------------------------------------------------------------------*
Public Function GetSorteoRecord(ByVal intId As Integer) As Boolean
    Dim rngFind    As Range
  On Error GoTo GetSorteoRecord_Error
    GetSorteoRecord = False
    
    Set rngFind = mrngIds.Find(What:=intId, LookIn:=xlValues, LookAt:=xlWhole)
    If Not rngFind Is Nothing Then
        GetRegistro rngFind.Row
        GetSorteoRecord = True
     End If
     Set rngFind = Nothing
  
  On Error GoTo 0
GetSorteoRecord__CleanExit:
    Exit Function
            
GetSorteoRecord_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    '   Audita el error
    Call HandleException(ErrNumber, ErrDescription, "SorteoModel.GetSorteoRecord", ErrSource)
    '   Lanza el Error
    Err.Raise ErrNumber, "SorteoModel.GetSorteoRecord", ErrDescription
End Function
'------------------------------------------------------------------------------*
' Procedimiento  : GetFirstSorteo
' Fecha          : lu., 27/may/2019 23:46:52
' Prop�sito      : Obtiene el primer registro del rango
'------------------------------------------------------------------------------*
Public Function GetFirstSorteo() As Boolean
    Dim rngFind As Range

  On Error GoTo GetFirstSorteo_Error
    GetFirstSorteo = False
    Set rngFind = mrngIds.Rows(1)
    If Not rngFind Is Nothing Then
        GetRegistro rngFind.Row
        GetFirstSorteo = True
    End If
    Set rngFind = Nothing
  
  On Error GoTo 0
GetFirstSorteo__CleanExit:
    Exit Function
            
GetFirstSorteo_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    '   Audita el error
    Call HandleException(ErrNumber, ErrDescription, "SorteoModel.GetFirstSorteo", ErrSource)
    '   Lanza el Error
    Err.Raise ErrNumber, "SorteoModel.GetFirstSorteo", ErrDescription
End Function
'------------------------------------------------------------------------------*
' Procedimiento  : GetNextSorteoRecord
' Fecha          : mi., 29/may/2019 23:34:38
' Prop�sito      : Obtiene el siguiente registro con el Id
'------------------------------------------------------------------------------*
Public Function GetNextSorteoRecord(ByVal intId As Integer) As Boolean
    Dim rngFind As Range
    Dim mRow    As Integer
    
  On Error GoTo GetNextSorteoRecord_Error
    GetNextSorteoRecord = False
    
    Set rngFind = mrngIds.Find(What:=intId, LookIn:=xlValues, LookAt:=xlWhole)
    If Not rngFind Is Nothing Then
        '
        '   Comprobamos si es el �ltimo registro
        '
        If rngFind.Row < mLastRow Then
            mRow = rngFind.Row + 1
            GetRegistro mRow
            GetNextSorteoRecord = True
        Else
            GetNextSorteoRecord = False
        End If
    End If
    Set rngFind = Nothing
  
  On Error GoTo 0
GetNextSorteoRecord__CleanExit:
    Exit Function
            
GetNextSorteoRecord_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    '   Audita el error
    Call HandleException(ErrNumber, ErrDescription, "SorteoModel.GetNextSorteoRecord", ErrSource)
    '   Lanza el Error
    Err.Raise ErrNumber, "SorteoModel.GetNextSorteoRecord", ErrDescription

End Function
'------------------------------------------------------------------------------*
' Procedimiento  : GetPrevSorteoRecord
' Fecha          : mi., 29/may/2019 23:34:38
' Prop�sito      : Obtiene el anterior registro con el Id
'------------------------------------------------------------------------------*
Public Function GetPrevSorteoRecord(ByVal intId As Integer) As Boolean
    Dim rngFind As Range
    Dim mRow    As Integer
    
  On Error GoTo GetPrevSorteoRecord_Error
    GetPrevSorteoRecord = False
    
    Set rngFind = mrngIds.Find(What:=intId, LookIn:=xlValues, LookAt:=xlWhole)
    If Not rngFind Is Nothing Then
        '
        '   Comprobamos si es el primer registro
        '
        If rngFind.Row > 2 Then
            mRow = rngFind.Row - 1
            GetRegistro mRow
            GetPrevSorteoRecord = True
        Else
            GetPrevSorteoRecord = False
        End If
    End If
    Set rngFind = Nothing
  
  On Error GoTo 0
GetPrevSorteoRecord__CleanExit:
    Exit Function
            
GetPrevSorteoRecord_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    '   Audita el error
    Call HandleException(ErrNumber, ErrDescription, "SorteoModel.GetPrevSorteoRecord", ErrSource)
    '   Lanza el Error
    Err.Raise ErrNumber, "SorteoModel.GetPrevSorteoRecord", ErrDescription

End Function

'------------------------------------------------------------------------------*
' Procedimiento  : GetLastSorteo
' Fecha          : mi., 29/may/2019 23:43:05
' Prop�sito      : Obtiene el �ltimo registro
'------------------------------------------------------------------------------*
Public Function GetLastSorteo() As Boolean

  On Error GoTo GetLastSorteo_Error
    GetLastSorteo = False

    If mLastRow > 0 Then
        GetRegistro mLastRow
        GetLastSorteo = True
    End If
  
  On Error GoTo 0
GetLastSorteo__CleanExit:
    Exit Function
            
GetLastSorteo_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    '   Audita el error
    Call HandleException(ErrNumber, ErrDescription, "SorteoModel.GetLastSorteo", ErrSource)
    '   Lanza el Error
    Err.Raise ErrNumber, "SorteoModel.GetLastSorteo", ErrDescription
End Function
'------------------------------------------------------------------------------*
' Procedimiento  : GetPremios
' Fecha          :
' Prop�sito      : Define una variable de texto para la matriz de premios
'------------------------------------------------------------------------------*
Public Function GetPremios() As String
    Dim mTexto As String
    Dim i As Integer
  
  On Error GoTo GetPremios_Error
  
    If IsEmpty(mPremios) Then
        GetPremios = Empty
        Exit Function
    Else
        If Len(mPremios(0)) = 0 Then
            mTexto = "0"
        Else
            mTexto = mPremios(0)
        End If
        For i = 1 To UBound(mPremios)
            If Len(mPremios(i)) = 0 Then
                mTexto = mTexto & ";" & 0
            Else
                mTexto = mTexto & ";" & mPremios(i)
            End If
        Next i
    End If
    GetPremios = mTexto
    
  On Error GoTo 0
GetPremios__CleanExit:
    Exit Function
            
GetPremios_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    '   Audita el error
    Call HandleException(ErrNumber, ErrDescription, "SorteoModel.GetPremios", ErrSource)
    '   Lanza el Error
    Err.Raise ErrNumber, "SorteoModel.GetLastSorteo", ErrDescription
End Function

'' *===========(EOF): SorteoModel.cls
