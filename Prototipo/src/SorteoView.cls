VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SorteoView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' *============================================================================*
' *
' *     Fichero    : SorteoView.cls
' *
' *     Autor      : Carlos Almela Baeza
' *     Creación   : ma., 22/oct/2019 18:49:43
' *     Versión    : 1.0
' *     Propósito  : Interfaz gráfica del CU Definir Sorteo
' *
' *============================================================================*
Option Explicit
Option Base 0
'
'--- Variables Privadas -------------------------------------------------------*
Private mwsUserInterface As Worksheet

'--- Constantes ---------------------------------------------------------------*
Private Const LT_PAGINA As String = "Página:"
Private Const LT_ADDRESS As String = "T$1:Z$1"
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
    mStr = mwsUserInterface.RAnge("T20").Value
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
    Call HandleException(ErrNumber, ErrDescription, "SorteoView.PaginaActual", ErrSource)
    '   Lanza el Error
    Err.Raise ErrNumber, "SorteoView.PaginaActual", ErrDescription

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
    mStr = mwsUserInterface.RAnge("T20").Value
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
    Call HandleException(ErrNumber, ErrDescription, "SorteoView.TotalPaginas", ErrSource)
    '   Lanza el Error
    Err.Raise ErrNumber, "SorteoView.TotalPaginas", ErrDescription

End Property


'--- Métodos Privados ---------------------------------------------------------*
Private Sub Class_Initialize()
    Set mwsUserInterface = Worksheets("Sorteos")
End Sub

Private Sub Class_Terminate()
    Set mwsUserInterface = Nothing
End Sub
'--- Métodos Públicos ---------------------------------------------------------*
'------------------------------------------------------------------------------*
' Procedimiento  : GetDisplaySorteo
' Fecha          : do., 12/may/2019 21:06:51
' Propósito      : Configura la vista para el juego
'------------------------------------------------------------------------------*
Public Function GetDisplaySorteo() As SorteoModel

End Function
'------------------------------------------------------------------------------*
' Procedimiento  : ClearSorteoDisplay
' Fecha          : do., 12/may/2019 21:06:51
' Propósito      : Limpia los datos del registro del formulario
'------------------------------------------------------------------------------*
Public Sub ClearSorteoDisplay(EsNuevo As Boolean)

End Sub
'------------------------------------------------------------------------------*
' Procedimiento  : DisplaySorteoRecord
' Fecha          : do., 12/may/2019 21:06:51
' Propósito      : Visualiza las propiedades del Sorteo
'------------------------------------------------------------------------------*
Public Sub DisplaySorteoRecord(oModel As SorteoModel)

End Sub
'------------------------------------------------------------------------------*
' Procedimiento  : SetDisplayJuego
' Fecha          : do., 12/may/2019 21:09:15
' Propósito      : Establece los literales según el juego
'------------------------------------------------------------------------------*
Public Sub SetDisplayJuego(ByVal mJuego As String)

End Sub

'------------------------------------------------------------------------------*
' Procedimiento  : ClearGrid
' Fecha          : ju., 20/jun/2019 10:57:09
' Propósito      : Inicializa la Rejilla
' Parámetros     : N/A
'------------------------------------------------------------------------------*
'
Public Sub ClearGrid()
    
    With mwsUserInterface.RAnge("T12:Z29")
        .ClearContents
        .Interior.Pattern = xlSolid
        .Interior.PatternColorIndex = xlAutomatic
        .Interior.ThemeColor = xlThemeColorAccent1
        .Interior.TintAndShade = 0.799981688894314
        .Interior.PatternTintAndShade = 0
    End With
    mwsUserInterface.RAnge("T20").Value = LT_PAGINA
    
End Sub
'------------------------------------------------------------------------------*
' Procedimiento  : ClearFiltros
' Fecha          : ju., 20/jun/2019 11:05:00
' Propósito      : Inicializa el area de los filtros
' Parámetros     : N/A
'------------------------------------------------------------------------------*
'
Public Sub ClearFiltros()
    mwsUserInterface.RAnge("W5:Y9").ClearContents
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
            mwsUserInterface.RAnge("T12").Offset(mRow, mCol).Value = oModel.ResultadosSearch(mRow, mCol)
        Next mCol
    Next mRow
    '
    '   Ubica la página
    '
    mwsUserInterface.RAnge("T20").Value = LT_PAGINA & CStr(oModel.PaginaActual) & "/" & CStr(oModel.TotalPaginas)
  
  On Error GoTo 0
AddSorteoToGrid__CleanExit:
    Exit Sub
        
AddSorteoToGrid_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    '   Audita el error
    Call HandleException(ErrNumber, ErrDescription, "SorteoView.AddSorteosToGrid", ErrSource)
    '   Lanza el Error
    Err.Raise ErrNumber, "SorteoView.AddSorteosToGrid", ErrDescription
End Sub
'------------------------------------------------------------------------------*
' Procedimiento  : GetFiltroBusqueda
' Fecha          : ma., 22/oct/2019 19:03:55
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
    Set oModel = GetSorteoModel
    '
    '   Asignamos valores de las celdas a las variables del modelo
    '
    With oModel
        If mwsUserInterface.RAnge("W5").Value = "Todos los Juegos" Then
            .Juego = Empty
        Else
            .Juego = mwsUserInterface.RAnge("W5").Value
        End If
        .FechaSorteo = mwsUserInterface.RAnge("W6").Value
        .PeriodoTexto = mwsUserInterface.RAnge("W7").Value
        .FechaInicio = mwsUserInterface.RAnge("W8").Value
        .FechaFin = mwsUserInterface.RAnge("W9").Value
        .LineasPorPagina = 8
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
    Call HandleException(ErrNumber, ErrDescription, "SorteoView.GetFiltroBusqueda", ErrSource)
    '   Lanza el Error
    Err.Raise ErrNumber, "SorteoView.GetFiltroBusqueda", ErrDescription
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
    
    '
    '   Inicializamos el fondo de la rejilla
    '
    With mwsUserInterface.RAnge("T12:Z19")
        .Interior.Pattern = xlSolid
        .Interior.PatternColorIndex = xlAutomatic
        .Interior.ThemeColor = xlThemeColorAccent1
        .Interior.TintAndShade = 0.799981688894314
        .Interior.PatternTintAndShade = 0
    End With
    '
    '   Marcamos la fila seleccionada
    '
    With mwsUserInterface.RAnge(mAddress).Interior
        .Pattern = xlSolid
        .ThemeColor = xlThemeColorAccent6
        .PatternTintAndShade = 0
    End With
    
End Sub


'' *===========(EOF): SorteoView.cls


