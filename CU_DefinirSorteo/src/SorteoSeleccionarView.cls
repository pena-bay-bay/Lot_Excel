VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SorteoSeleccionarView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' *============================================================================*
' *
' *     Fichero    : SorteoSeleccionarView.cls
' *
' *     Autor      : Carlos Almela Baeza
' *     Creación   : ju., 20/jun/2019 10:57:09
' *     Versión    : 1.0
' *     Propósito  : Interfaz de Usuario Vista Seleccionar Sorteo MVC
' *
' *============================================================================*
Option Explicit
Option Base 0
'--- Variables Privadas -------------------------------------------------------*
Private mwsUserInterface As Worksheet
Private Const LT_PAGINA As String = "Página:"
Private Const LT_ADDRESS As String = "B$1:H$1"
'--- Mensajes -----------------------------------------------------------------*
'--- Errores ------------------------------------------------------------------*
'--- Propiedades --------------------------------------------------------------*
'------------------------------------------------------------------------------*
' Procedimiento  : PaginaActual
' Fecha          : vi., 19/jul/2019 12:35:49
' Propósito      : Obtiene la página visualizada en la rejilla
'------------------------------------------------------------------------------*
'
Public Property Get PaginaActual() As Integer
    Dim mStr      As String
    Dim i         As Integer
    Dim j         As Integer
    
  On Error GoTo PaginaActual_Error
    '
    '   Inicializamos el dato
    '
    PaginaActual = 0
    '
    '   Obtenemos el contenido de la celda
    '
    mStr = mwsUserInterface.Range("B19").value
    '
    '   Obtenemos el inicio de la página
    '
    i = InStr(mStr, ":")
    '
    '   Obtenemos la posición de la barra
    '
    j = InStr(mStr, "/")
    '
    '   Comprobamos que se han encontrado los signos separadores
    '
    If i > 0 And j > 0 Then
        i = i + 1
        j = j - i
        mStr = mId(mStr, i, j)
        If IsNumeric(mStr) Then
            PaginaActual = CInt(mStr)
        Else
            PaginaActual = 1
        End If
    End If
    
  On Error GoTo 0
PaginaActual__CleanExit:
    Exit Property
        
PaginaActual_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    '   Audita el error
    Call HandleException(ErrNumber, ErrDescription, "SorteoSeleccionarView.PaginaActual", ErrSource)
    '   Lanza el Error
    Err.Raise ErrNumber, "SorteoSeleccionarView.PaginaActual", ErrDescription

End Property
'------------------------------------------------------------------------------*
' Procedimiento  : TotalPaginas
' Fecha          : vi., 19/jul/2019 18:35:44
' Propósito      : Obtiene el total de paginas de la consulta
'------------------------------------------------------------------------------*
'
Public Property Get TotalPaginas() As Integer
    Dim mStr      As String
    Dim i         As Integer
    Dim j         As Integer
    
  On Error GoTo TotalPaginas_Error
    '
    '   Inicializamos el dato
    '
    TotalPaginas = 0
    '
    '   Obtenemos el contenido de la celda
    '
    mStr = mwsUserInterface.Range("B19").value
    '
    '   Obtenemos la posición de la barra
    '
    i = InStr(mStr, "/")
    '
    '   Obtenemos el tamaño de la cadena
    '
    j = Len(mStr)
    '
    '   Comprobamos que se han encontrado los signos separadores
    '
    If i > 0 And j > 0 Then
        i = i + 1
        j = j - i + 1
        mStr = mId(mStr, i, j)
        If IsNumeric(mStr) Then
            TotalPaginas = CInt(mStr)
        Else
            TotalPaginas = 1
        End If
    End If
    
  On Error GoTo 0
TotalPaginas__CleanExit:
    Exit Property
        
TotalPaginas_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    '   Audita el error
    Call HandleException(ErrNumber, ErrDescription, "SorteoSeleccionarView.TotalPaginas", ErrSource)
    '   Lanza el Error
    Err.Raise ErrNumber, "SorteoSeleccionarView.TotalPaginas", ErrDescription

End Property

'--- Métodos Privados ---------------------------------------------------------*
Private Sub Class_Initialize()
    Set mwsUserInterface = ThisWorkbook.Sheets("Consultar")
End Sub
Private Sub Class_Terminate()
    Set mwsUserInterface = Nothing
End Sub
'--- Métodos Públicos ---------------------------------------------------------*
'------------------------------------------------------------------------------*
' Procedimiento  : ClearGrid
' Fecha          : ju., 20/jun/2019 10:57:09
' Propósito      : Inicializa la Rejilla
' Parámetros     : N/A
'------------------------------------------------------------------------------*
'
Public Sub ClearGrid()
    Application.ScreenUpdating = False
    With mwsUserInterface.Range("B12:H18")
        .ClearContents
        .Interior.Pattern = xlSolid
        .Interior.PatternColorIndex = xlAutomatic
        .Interior.ThemeColor = xlThemeColorAccent1
        .Interior.TintAndShade = 0.799981688894314
        .Interior.PatternTintAndShade = 0
    End With
    mwsUserInterface.Range("B19").value = LT_PAGINA
    Application.ScreenUpdating = True
End Sub
'------------------------------------------------------------------------------*
' Procedimiento  : ClearFiltros
' Fecha          : ju., 20/jun/2019 11:05:00
' Propósito      : Inicializa el area de los filtros
' Parámetros     : N/A
'------------------------------------------------------------------------------*
'
Public Sub ClearFiltros()
    Application.ScreenUpdating = False
    mwsUserInterface.Range("C5:C9").ClearContents
    Application.ScreenUpdating = True
End Sub
'------------------------------------------------------------------------------*
' Procedimiento  : AddSorteosToGrid
' Fecha          : ju., 20/jun/2019 11:07:29
' Propósito      : Visualizar los registros encontrados en la rejilla
' Parámetros     : oModel Modelo de datos
'------------------------------------------------------------------------------*
'
Public Sub AddSorteosToGrid(oModel As SorteoModel)
    Dim mRow As Integer         ' Fila
    Dim mCol As Integer         ' Columna
    
  On Error GoTo AddSorteoToGrid_Error
    '
    '   Para cada elemento de la matriz del modelo
    '
    For mRow = 0 To UBound(oModel.ResultadosSearch, 1)
        For mCol = 0 To UBound(oModel.ResultadosSearch, 2)
            '
            '   Actualizamos la fila de la rejilla correspondiente
            '
            mwsUserInterface.Range("B12").Offset(mRow, mCol).value = oModel.ResultadosSearch(mRow, mCol)
        Next mCol
    Next mRow
    '
    '   Ubica la página
    '
    mwsUserInterface.Range("B19").value = LT_PAGINA & CStr(oModel.PaginaActual) & "/" & CStr(oModel.TotalPaginas)
  
  On Error GoTo 0
AddSorteoToGrid__CleanExit:
    Exit Sub
        
AddSorteoToGrid_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    '   Audita el error
    Call HandleException(ErrNumber, ErrDescription, "SorteoSeleccionarView.AddSorteosToGrid", ErrSource)
    '   Lanza el Error
    Err.Raise ErrNumber, "SorteoSeleccionarView.AddSorteosToGrid", ErrDescription
End Sub
'------------------------------------------------------------------------------*
' Procedimiento  : GetFiltroBusqueda
' Fecha          :
' Propósito      : Guardar la información de un Nuevo sorteo en el repositorio
' Parametros     : Obtiene la información del filtro
'------------------------------------------------------------------------------*
'
Public Function GetFiltroBusqueda() As SorteoModel
    Dim oModel As SorteoModel
  
  On Error GoTo GetFiltroBusqueda_Error
    '
    '   Creamos el modelo de datos
    '
    Set oModel = New SorteoModel
    '
    '   Asignamos valores de las celdas a las variables del modelo
    '
    With oModel
        If mwsUserInterface.Range("C5").value = "Todos los Juegos" Then
            .Juego = Empty
        Else
            .Juego = mwsUserInterface.Range("C5").value
        End If
        .FechaSorteo = mwsUserInterface.Range("C6").value
        .PeriodoTexto = mwsUserInterface.Range("C7").value
        .FechaInicio = mwsUserInterface.Range("C8").value
        .FechaFin = mwsUserInterface.Range("C9").value
        .LineasPorPagina = 7
    End With
    '
    '   Devolvemos el modelo
    '
    Set GetFiltroBusqueda = oModel
  
  On Error GoTo 0
GetFiltroBusqueda__CleanExit:
    Exit Function
        
GetFiltroBusqueda_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    '   Audita el error
    Call HandleException(ErrNumber, ErrDescription, "SorteoSeleccionarView.GetFiltroBusqueda", ErrSource)
    '   Lanza el Error
    Err.Raise ErrNumber, "SorteoSeleccionarView.GetFiltroBusqueda", ErrDescription
End Function
'------------------------------------------------------------------------------*
' Procedimiento  : SetFilaRegistro
' Fecha          : vi., 26/jul/2019 20:43:55
' Propósito      : Marcar la fila seleccionada
' Parametros     : datRow es la fila de la hoja selecionada
'------------------------------------------------------------------------------*
'
Public Sub SetFilaRegistro(datRow As Integer)
    Dim mAddress  As String
    '
    '   Calculamos el área de la fila
    '
    mAddress = Replace(LT_ADDRESS, "$1", Trim(Str(datRow)))
    Application.ScreenUpdating = False
    '
    '   Inicializamos el fondo de la rejilla
    '
    With mwsUserInterface.Range("B12:H18")
        .Interior.Pattern = xlSolid
        .Interior.PatternColorIndex = xlAutomatic
        .Interior.ThemeColor = xlThemeColorAccent1
        .Interior.TintAndShade = 0.799981688894314
        .Interior.PatternTintAndShade = 0
    End With
    '
    '   Marcamos la fila seleccionada
    '
    With mwsUserInterface.Range(mAddress).Interior
        .Pattern = xlSolid
        .ThemeColor = xlThemeColorAccent6
        .PatternTintAndShade = 0
    End With
    Application.ScreenUpdating = True
End Sub
' *===========(EOF): SorteoSeleccionarView.cls
