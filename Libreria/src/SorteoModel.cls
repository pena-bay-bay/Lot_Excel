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
' *     Creación   : ju., 09/may/2019 00:04:56
' *     Versión    : 1.0
' *     Propósito  :
' *
' *============================================================================*
Option Explicit
Option Base 0
'
'--- Variables Privadas -------------------------------------------------------*
Private mId             As Integer      ' Identificador único del sorteo
Private mJuego          As String       ' Juego al que pertenece el sorteo
Private mNumeroSorteo   As String       ' Numero de sorteo del juego
Private mFecha          As String       ' Fecha del sorteo
Private mDia            As String       ' Dia de la semana
Private mSemana         As Integer      ' Semana del año
Private mOrden          As String       ' Orden de aparicion de los datos
Private mCombinacion    As String       ' Combinación ganadora
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
Private mLineasPagina   As Integer      ' Numero de lineas por páginas
Private mTotRegistros   As Integer      ' Total registros de la base de datos
Private mPagActual      As Integer      ' Página Actual
Private mMatrizResultados As Variant    ' Matriz de resultados de Busqueda
'----Area de Datos
Private mwsData         As Worksheet    ' hoja de datos del repositorio
Private mrngIds         As Range        ' Rango de los datos
Private mLastRow        As Integer      ' ultima fila del rango
Private mFirstRow       As Integer      ' ultima fila del rango
'
Private mSort           As Sorteo       ' sorteo en curso
Private mEng            As SorteoEngine ' Motor de Sorteos
Private mTipoJuego      As Juego
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
    Select Case mJuego
        Case LT_BONOLOTO: mTipoJuego = bonoloto
        Case LT_PRIMITIVA: mTipoJuego = LoteriaPrimitiva
        Case LT_EUROMILLON: mTipoJuego = Euromillones
        Case LT_GORDO: mTipoJuego = gordoPrimitiva
    End Select
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
'--- Métodos Privados ---------------------------------------------------------*
Private Sub Class_Initialize()
    '
    '   Definimos Matriz de Resultados
    '
    ReDim mMatrizResultados(7, 7) As String
    '
    '   Inicializamos variables
    '
    Set mPeriodo = New Periodo
End Sub
'------------------------------------------------------------------------------*
' Procedimiento  : SetPremios
' Fecha          : sá., 01/jun/2019 22:12:44
' Propósito      : Establece los premios por defecto de un sorteo
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
' Fecha          : sá., 01/jun/2019 22:12:44
' Propósito      : Establece el texto de la combinacion
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
' Fecha          : sá., 01/jun/2019 22:12:44
' Propósito      : Establece el registro en el modelo
'------------------------------------------------------------------------------*
Private Sub GetRegistro(mRow As Integer)
    With mwsData
        Me.IdSelected = .Cells(mRow, 1).Value
        Me.Juego = .Cells(mRow, 2).Value
        Me.NumSorteo = .Cells(mRow, 3).Value
        Me.FechaSorteo = Format(.Cells(mRow, 4).Value, "dd/mm/yyyy")
        Me.DiaSemana = .Cells(mRow, 5).Value
        Me.Semana = .Cells(mRow, 6).Value
        Me.OrdenAparicion = .Cells(mRow, 7).Value
        Me.CombinacionGanadora = .Cells(mRow, 8).Value
        Me.Complementario = .Cells(mRow, 9).Value
        Me.Estrellas = .Cells(mRow, 10).Value
        Me.Reintegro = .Cells(mRow, 11).Value
        SetPremios .Cells(mRow, 12).Value
    End With
End Sub

'--- Métodos Públicos ---------------------------------------------------------*
'------------------------------------------------------------------------------*
' Procedimiento  : SearchSorteos
' Fecha          : vi., 28/jun/2019 19:26:59
' Propósito      : Busca los sorteos con el criterio seleccionado
'------------------------------------------------------------------------------*
Public Function SearchSorteos() As Boolean
    Dim mJuegoSearch     As String       ' Juego de busqueda
    Dim mDateSearch      As Date         ' Fecha de busqueda
    Dim mResult          As Collection   ' Colección de resultados
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
    '   Inicializamos la función No se ha encontrado nada
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
    '   Creamos una colección de resultados
    '
    Set mResult = New Collection
    '
    '   Para cada registro en la tabla
    '
    For Each mCelda In mrngIds.Rows
        '
        '   Establecemos la fecha y el juego a comparar
        '
        mDatFecha = CDate(mCelda.Offset(0, 3).Value)
        mDatJuego = mCelda.Offset(0, 1).Value
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
            mResult.Add Item:=mCelda.Value, Key:=CStr(mCelda.Value)
        End If
    Next mCelda
    '
    '   Si hemos encontrado coincidencias
    '
    If mResult.Count > 0 Then
        SearchSorteos = True
        '
        '   Calculamos total registros y total páginas
        '
        mTotRegistros = mResult.Count
        '
        '   Calculamos el total de páginas
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
            Case mTotalPaginas = 1:             ' Solo tenemos una página
               mIniReg = 1
               mFinReg = mTotRegistros
                
            Case mTotalPaginas > 1 _
             And mPagActual > mTotalPaginas:    ' Mas de una página y la página de busqueda superior a la actual
               mIniReg = (mTotalPaginas - 1) * mLineasPagina
               mFinReg = mTotRegistros
            
            Case mTotalPaginas > 1 _
             And mPagActual = mTotalPaginas:    ' Mas de una página y la página de busqueda es la última
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
' Propósito      : Guarda el sorteo en la base de datos
'------------------------------------------------------------------------------*
Public Function GuardarSorteoRecord(ByVal oSorteoModel As SorteoModel) As Boolean
    Dim rngFind As Range           ' Rango de búsqueda
    Dim mTmp    As String          ' Cadena de texto temporal
    
  On Error GoTo GuardarSorteoRecord_Error
    '
    '   TODO Guardar Premio
    '
    GuardarSorteoRecord = False
    Set rngFind = mrngIds.Find(What:=oSorteoModel.IdSelected, LookIn:=xlValues, LookAt:=xlWhole)
    If Not rngFind Is Nothing Then
        With mwsData
            .Cells(rngFind.Row, 1).Value = oSorteoModel.IdSelected
            .Cells(rngFind.Row, 2).Value = oSorteoModel.Juego
            .Cells(rngFind.Row, 3).Value = oSorteoModel.NumSorteo
            .Cells(rngFind.Row, 4).Value = DateValue(oSorteoModel.FechaSorteo)
            .Cells(rngFind.Row, 5).Value = oSorteoModel.DiaSemana
            .Cells(rngFind.Row, 6).Value = oSorteoModel.Semana
            .Cells(rngFind.Row, 7).Value = oSorteoModel.OrdenAparicion
            .Cells(rngFind.Row, 8).Value = oSorteoModel.CombinacionGanadora
            .Cells(rngFind.Row, 9).Value = oSorteoModel.Complementario
            .Cells(rngFind.Row, 10).Value = "'" & oSorteoModel.Estrellas
            .Cells(rngFind.Row, 11).Value = oSorteoModel.Reintegro
            .Cells(rngFind.Row, 12).Value = oSorteoModel.GetPremios
        End With
        GuardarSorteoRecord = True
    Else
        mLastRow = mrngIds.Rows.Count + 2
        With mwsData
            .Cells(mLastRow, 1).Value = oSorteoModel.IdSelected
            .Cells(mLastRow, 2).Value = oSorteoModel.Juego
            .Cells(mLastRow, 3).Value = oSorteoModel.NumSorteo
            .Cells(mLastRow, 4).Value = CDate(oSorteoModel.FechaSorteo)
            .Cells(mLastRow, 5).Value = oSorteoModel.DiaSemana
            .Cells(mLastRow, 6).Value = oSorteoModel.Semana
            .Cells(mLastRow, 7).Value = oSorteoModel.OrdenAparicion
            .Cells(mLastRow, 8).Value = oSorteoModel.CombinacionGanadora
            .Cells(mLastRow, 9).Value = oSorteoModel.Complementario
            .Cells(mLastRow, 10).Value = "'" & oSorteoModel.Estrellas
            .Cells(mLastRow, 11).Value = oSorteoModel.Reintegro
            .Cells(mLastRow, 12).Value = oSorteoModel.GetPremios
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
' Propósito      : Crea un nuevo registro
'------------------------------------------------------------------------------*
Public Function NuevoSorteoRecord() As Boolean
 
  On Error GoTo NuevoSorteoRecord_Error
    NuevoSorteoRecord = False
    '
    '   Creamos el motor de busquedas de
    '
    Set mEng = New SorteoEngine
    '
    '   Obtenemos un sorteo nuevo
    '
    Set mSort = mEng.GetNewSorteo(mTipoJuego)
    '
    '   Trasladamos las propiedades del sorteo a las variables del modelo
    '
    With mSort
        Me.IdSelected = .Id
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
    End With
    
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
' Propósito      : Elimina un registro por el Id
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
' Propósito      : Obtiene el registro con el Id
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
' Propósito      : Obtiene el primer registro del rango
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
' Propósito      : Obtiene el siguiente registro con el Id
'------------------------------------------------------------------------------*
Public Function GetNextSorteoRecord(ByVal intId As Integer) As Boolean
    Dim rngFind As Range
    Dim mRow    As Integer
    
  On Error GoTo GetNextSorteoRecord_Error
    GetNextSorteoRecord = False
    
    Set rngFind = mrngIds.Find(What:=intId, LookIn:=xlValues, LookAt:=xlWhole)
    If Not rngFind Is Nothing Then
        '
        '   Comprobamos si es el último registro
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
' Propósito      : Obtiene el anterior registro con el Id
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
' Propósito      : Obtiene el último registro
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
' Propósito      : Define una variable de texto para la matriz de premios
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
