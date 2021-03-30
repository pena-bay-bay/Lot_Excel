VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GenCombinacionesView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' *============================================================================*
' *
' *     Fichero    : GenCombinacionesView.cls
' *
' *     Autor      : Carlos Almela Baeza
' *     Creación   : mi., 21/ago/2019 17:01:04
' *     Versión    : 1.0
' *     Propósito  : Interfaz del caso de uso que obtiene y visualiza información
' *                  de la hoja de portada
' *============================================================================*
Option Explicit
Option Base 0
'
'--- Variables Privadas -------------------------------------------------------*
Private mwsUserInterface As Worksheet

'--- Constantes ---------------------------------------------------------------*
Private Const CELDA_SUGERENCIAS As String = "B3"
Private Const CELDA_TOTNUMEROS As String = "B4"
Private Const CELDA_TOTCOMBINACIONES As String = "B5"
Private Const CELDA_COMBIGNERADAS As String = "C5"
Private Const CELDA_TIPOFILTRO As String = "B7"
Private Const CELDA_VALORFILTRO As String = "B8"
Private Const CELDA_MATRIZFILTROS As String = "B13"
Private Const CELDA_CATEGORIAFILTERS As String = "B50"
'--- Mensajes -----------------------------------------------------------------*
'--- Errores ------------------------------------------------------------------*

'--- Propiedades --------------------------------------------------------------*
Public Property Get NumSugerencias() As Integer
    If Len(mwsUserInterface.Range(CELDA_SUGERENCIAS).Value) > 0 Then
        NumSugerencias = CInt(mwsUserInterface.Range(CELDA_SUGERENCIAS).Value)
    Else
        NumSugerencias = 0
    End If
End Property

Public Property Let NumSugerencias(ByVal vNewValue As Integer)
    mwsUserInterface.Range(CELDA_SUGERENCIAS).Value = vNewValue
End Property

Public Property Get RangoSalida() As Range
    Set RangoSalida = mwsUserInterface.Range("G3")
End Property

Public Property Get TotalNumeros() As Integer
    If Len(mwsUserInterface.Range(CELDA_TOTNUMEROS).Value) > 0 Then
        TotalNumeros = CInt(mwsUserInterface.Range(CELDA_TOTNUMEROS).Value)
    Else
        TotalNumeros = 0
    End If
End Property
Public Property Get TotalCombinaciones() As Long
    If Len(mwsUserInterface.Range(CELDA_TOTCOMBINACIONES).Value) > 0 Then
        TotalCombinaciones = CLng(mwsUserInterface.Range(CELDA_TOTCOMBINACIONES).Value)
    Else
        TotalCombinaciones = 0
    End If
End Property
Public Property Get CombinacionesGeneradas() As Long
    If Len(mwsUserInterface.Range(CELDA_COMBIGNERADAS).Value) > 0 Then
        CombinacionesGeneradas = CInt(mwsUserInterface.Range(CELDA_COMBIGNERADAS).Value)
    Else
        CombinacionesGeneradas = 0
    End If
End Property

Public Property Get TipoFiltro() As String
    TipoFiltro = mwsUserInterface.Range(CELDA_TIPOFILTRO).Value
End Property
Public Property Get ValorFiltro() As String
    ValorFiltro = mwsUserInterface.Range(CELDA_VALORFILTRO).Value
End Property


'--- Métodos Privados ---------------------------------------------------------*
Private Sub Class_Initialize()
    Set mwsUserInterface = ThisWorkbook.Sheets("Portada")
End Sub

Private Sub Class_Terminate()
    Set mwsUserInterface = Nothing
End Sub
'--- Métodos Públicos ---------------------------------------------------------*
'------------------------------------------------------------------------------*
' Procedimiento  : GetParametrosProceso
' Fecha          : mi., 21/ago/2019 19:16:17
' Propósito      : Devuelve la información de la hoja de portada para el modelo
' Retorno        : Modelo
'------------------------------------------------------------------------------*
Public Function GetParametrosProceso() As GenCombinacionesModel
    Dim mMatriz As Variant
    Dim mFiltros As Variant
    
 On Error GoTo GetParametrosProceso_Error
    '
    '   Creamos el modelo
    '
    Set GetParametrosProceso = New GenCombinacionesModel
    '
    '   Obtenemos los numeros para combinar
    '
    mMatriz = GetNumeros
    '
    '   Obtenemos los filtros
    '
    mFiltros = GetFiltros
    '
    '   Asignamos
    '
    With GetParametrosProceso
        .MatrizNumeros = mMatriz
        .Filtros = mFiltros
        .NumerosSugerencia = Me.NumSugerencias
        .TotalNumerosCombinar = Me.TotalNumeros
        .TotalCombinaciones = Me.TotalCombinaciones
    End With
    '
    '
    '
  On Error GoTo 0
GetParametrosProceso__CleanExit:
    Exit Function
            
GetParametrosProceso_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    Call HandleException(ErrNumber, ErrDescription, "GenCombinacionesView.GetParametrosProceso", ErrSource)
    Err.Raise ErrNumber, "GenCombinacionesView.GetParametrosProceso", ErrDescription
End Function

'------------------------------------------------------------------------------*
' Procedimiento  : ClearSalida
' Fecha          : mi., 21/ago/2019 19:10:10
' Propósito      : Inicializar las celdas de salida de la combinación
'------------------------------------------------------------------------------*
Public Sub ClearSalida()
    
    Application.ScreenUpdating = False
    mwsUserInterface.Range("G3").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearContents
    mwsUserInterface.Range("D5").Select
    Application.ScreenUpdating = True
    
End Sub
'------------------------------------------------------------------------------*
' Procedimiento  : GetFiltros
' Fecha          : mi., 28/ago/2019 18:48:23
' Propósito      : Obtiene la matriz de filtro
'------------------------------------------------------------------------------*
Private Function GetFiltros() As Variant
    Dim mMtrz() As String
    Dim mFiltr  As String
    Dim i       As Integer
    
 On Error GoTo GetFiltros_Error
    '
    '   Inicializamos la Matriz con un elemento vacio
    '
    i = 0
    ReDim Preserve mMtrz(i)
    mMtrz(i) = Empty
    '
    '   Nos posicionamos en la celda de inicio
    '
    mwsUserInterface.Range("B13").Select
    '
    '   Repetir hasta que la celda esté vacia o espacios
    '
    Do
        '
        '   Obtenemos el contenido de la celda
        '
        mFiltr = Selection.Offset(i, 0).Value
        '
        '   Si tiene contenido
        '
        If Len(mFiltr) > 0 Then
            ReDim Preserve mMtrz(i)
            mMtrz(i) = mFiltr
        End If
        i = i + 1
    Loop Until (IsEmpty(mFiltr) Or Len(mFiltr) = 0)
    '
    '
    '
    GetFiltros = mMtrz

  On Error GoTo 0
GetFiltros__CleanExit:
    Exit Function
            
GetFiltros_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    Call HandleException(ErrNumber, ErrDescription, "GenCombinacionesView.GetFiltros", ErrSource)
    Err.Raise ErrNumber, "GenCombinacionesView.GetFiltros", ErrDescription
End Function
'------------------------------------------------------------------------------*
' Procedimiento  : GetNumeros
' Fecha          : mi., 28/ago/2019 19:30:23
' Propósito      : Obtiene la matriz de Numeros
'------------------------------------------------------------------------------*
Private Function GetNumeros() As Variant
    Dim mMtrz() As String
    Dim mNum    As String
    Dim i       As Integer
    
 On Error GoTo GetNumeros_Error
    '
    '   Inicializamos la Matriz con un elemento vacio
    '
    i = 0
    ReDim Preserve mMtrz(i)
    mMtrz(i) = Empty
    '
    '   Nos posicionamos en la celda de inicio
    '
    mwsUserInterface.Range("D7").Select
    '
    '   Repetir hasta que la celda esté vacia o espacios
    '
    Do
        '
        '   Obtenemos el contenido de la celda
        '
        mNum = Selection.Offset(i, 0).Value
        '
        '   Si tiene contenido
        '
        If Len(mNum) > 0 Then
            ReDim Preserve mMtrz(i)
            mMtrz(i) = mNum
        End If
        i = i + 1
    Loop Until (IsEmpty(mNum) Or Len(mNum) = 0)
    '
    '
    '
    GetNumeros = mMtrz

  On Error GoTo 0
GetNumeros__CleanExit:
    Exit Function
            
GetNumeros_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    Call HandleException(ErrNumber, ErrDescription, "GenCombinacionesView.GetNumeros", ErrSource)
    Err.Raise ErrNumber, "GenCombinacionesView.GetNumeros", ErrDescription
End Function
'------------------------------------------------------------------------------*
' Procedimiento  : AgregarFiltro
' Fecha          : do., 01/sep/2019 20:42:52
' Propósito      : Agregar un nuevo filtro a la matriz
'------------------------------------------------------------------------------*
Public Sub AgregarFiltro()
    Dim mFiltro As String
    Dim i As Integer
    Dim mStr As Variant
    
  On Error GoTo AgregarFiltro_Error
    '
    '   Pasamos los valores al modelo y nos devuelve un filtro
    '
    mFiltro = Me.TipoFiltro & "|" & Me.ValorFiltro
    '
    '   Inicializamos
    '
    i = 0
    '
    '   Nos posicionamos en la matriz de filtros
    '
    mwsUserInterface.Range(CELDA_MATRIZFILTROS).Select
    '
    '   Buscamos la primera celda vacia
    '
    Do
        '
        '   Obtenemos el contenido de la celda
        '
        mStr = Selection.Offset(i, 0).Value
        '
        '   Si tiene contenido
        '
        If Len(mStr) > 0 Then
            If mStr = mFiltro Then
                MsgBox "El filtro ya está definido.", vbOKOnly + vbExclamation, "Agregar Filtro Combinacion"
                GoTo AgregarFiltro__CleanExit
            End If
            i = i + 1
        End If
        
    Loop Until (IsEmpty(mStr) Or Len(mStr) = 0)
    '
    '   Si hay más de 10 filtro avisamos
    '
    If i > 11 Then
        MsgBox "El número máximo de filtros admitidos es 10.", vbOKOnly + vbExclamation, "Agregar Filtro Combinacion"
        GoTo AgregarFiltro__CleanExit
    Else
        Selection.Offset(i, 0).Value = mFiltro
    End If
  On Error GoTo 0
AgregarFiltro__CleanExit:
    Exit Sub
            
AgregarFiltro_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    Call HandleException(ErrNumber, ErrDescription, "GenCombinacionesController.AgregarFiltro", ErrSource)
    Err.Raise ErrNumber, "GenCombinacionesController.AgregarFiltro", ErrDescription
End Sub
'------------------------------------------------------------------------------*
' Procedimiento  : BorrarFiltros
' Fecha          : do., 01/sep/2019 20:39:59
' Propósito      : Eliminar el contenido de las celdas donde se agregan filtros
'------------------------------------------------------------------------------*
Public Sub BorrarFiltros()
    Dim i       As Integer
    Dim mStr    As Variant
    
  On Error GoTo BorrarFiltros_Error
    '
    '   Nos posicionamos en la matriz de filtros
    '
    mwsUserInterface.Range(CELDA_MATRIZFILTROS).Select
    i = 0
    '
    '   Repetir hasta que la celda esté vacia o espacios
    '
    Do
        '
        '   Obtenemos el contenido de la celda
        '
        mStr = Selection.Offset(i, 0).Value
        '
        '   Si tiene contenido
        '
        If Len(mStr) > 0 Then
            Selection.Offset(i, 0).ClearContents
        End If
        i = i + 1
    Loop Until (IsEmpty(mStr) Or Len(mStr) = 0)
 
 On Error GoTo 0
BorrarFiltros__CleanExit:
    Exit Sub
            
BorrarFiltros_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    Call HandleException(ErrNumber, ErrDescription, "GenCombinacionesController.BorrarFiltros", ErrSource)
    Err.Raise ErrNumber, "GenCombinacionesController.BorrarFiltros", ErrDescription
End Sub

'------------------------------------------------------------------------------*
' Procedimiento  : SetFiltros
' Fecha          : lu., 02/sep/2019 18:59:43
' Propósito      : Establece los posibles valores para un filtro
'------------------------------------------------------------------------------*
Public Sub SetFiltros(datArray As Variant)
    Dim i As Integer
    Dim mRg As String
    
  On Error GoTo SetFiltros_Error
    '
    '   Borramos los filtros existentes
    '
    mwsUserInterface.Range(CELDA_CATEGORIAFILTERS).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearContents
    '
    '   Nos posicionamos en la primera celda
    '
    mwsUserInterface.Range(CELDA_CATEGORIAFILTERS).Select
    '
    '   Trasladamos la matriz a las celdas
    '
    For i = 0 To UBound(datArray)
        Selection.Offset(i, 0).Value = datArray(i)
    Next i
    '
    '   calculamos el rango de validación
    i = i + 49
    mRg = Replace("=$B$50:$B$#1", "#1", CStr(i))
    '
    '   Actualizamos la validación de la celda
    '
    mwsUserInterface.Range(CELDA_VALORFILTRO).Select
    With Selection.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:=mRg
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With
    '
    '   Inicializamos la celda
    '
    Selection.ClearContents
 On Error GoTo 0
SetFiltros__CleanExit:
    Exit Sub
            
SetFiltros_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    Call HandleException(ErrNumber, ErrDescription, "GenCombinacionesController.SetFiltros", ErrSource)
    Err.Raise ErrNumber, "GenCombinacionesController.SetFiltros", ErrDescription
End Sub
'' *===========(EOF): GenCombinacionesView.cls


