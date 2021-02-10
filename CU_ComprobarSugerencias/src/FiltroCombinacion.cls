VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FiltroCombinacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' *============================================================================*
' *
' *     Fichero    : FiltroCombinacion.cls
' *
' *     Autor      : Carlos Almela Baeza
' *     Creación   : mi., 21/ago/2019 23:54:45
' *     Versión    : 1.0
' *     Propósito  : Objeto Filtro para evaluar combinaciones
' *
' *============================================================================*
Option Explicit
Option Base 0

'--- Variables Privadas -------------------------------------------------------*
Private mTextoTipoFiltro As Variant
Private mRgMinSuma    As Integer
Private mRgMaxSuma    As Integer
Private mNameFiltro   As String
Private mColValues    As Collection
Private mFilterValue  As String

'--- Constantes ---------------------------------------------------------------*
Private Const FLTRS_CONSECUTIVOS_03 As String = "No Consecutivos,3/0,2/0"
Private Const FLTRS_CONSECUTIVOS_04 As String = "No Consecutivos,4/0,3/0,2/0,2/2"
Private Const FLTRS_CONSECUTIVOS_05 As String = "No Consecutivos,5/0,4/0,3/0," & _
                                                "3/2,2/0,2/2/0"
Private Const FLTRS_CONSECUTIVOS_06 As String = "No Consecutivos,6/0,5/0,4/2/0," & _
                                                "4/0,3/0,3/3/0,3/2/0,2/0,2/2/0"
Private Const FLTRS_CONSECUTIVOS_07 As String = "No Consecutivos,7/0,6/0,5/2/0," & _
                                                "5/0,4/3/0,4/2/0,4/0,3/3/0,3/2/2/0," & _
                                                "3/2/0,3/0,2/2/2/0,2/2/0,2/0"
Private Const FLTRS_CONSECUTIVOS_08 As String = "No Consecutivos,8/0,7/0,6/2/0," & _
                                                "6/0,5/3/0,5/2/0,5/0,4/4/0,4/3/0,4" & _
                                                "/2/2/0,4/2/0,4/0,3/3/2/0,3/3/0,3/2/2/0," & _
                                                "3/2/0,3/0,2/2/2/2/0,2/2/2/0,2/2/0,2/0"
Private Const FLTRS_CONSECUTIVOS_09 As String = "9/0,8/0,7/0,7/2/0,6/0,6/2/0,6/3/0," & _
                                                "5/0,5/2/0,5/3/0,5/2/2/0,4/0,4/2/0,4/" & _
                                                "3/0,4/4/0,4/3/2/0,4/2/2/0,3/0,3/2/0,3/3/0" & _
                                                ",3/3/2/0,3/2/2/2/0,2/0,2/2/0,2/2/2/0,2/2/2/2/0"
Private Const FLTRS_CONSECUTIVOS_10 As String = "10/0,9/0,8/0,8/2/0,7/0,7/2/0,7/3/0," & _
                                                "6/0,6/2/0,6/3/0,6/2/2/0,5/0,5/2/0,5/3/0," & _
                                                "5/2/2/0,5/3/2/0,4/0,4/2/0,4/3/0,4/4/0,4/" & _
                                                "4/2/0,4/3/2/0,4/2/2/0,4/3/2/0,4/2/2/2/0,3/" & _
                                                "0,3/2/0,3/3/0,3/3/2/0,3/3/2/2/0,2/0,2/2/0" & _
                                                ",2/2/2/0,2/2/2/2/0,2/2/2/2/2/0"
Private Const FLTRS_CONSECUTIVOS_11 As String = "11/0,10/0,9/0,9/2,8/0,8/3,8/2," & _
                                                "7/0,7/3,7/2,7/2/2,6/0,6/5,6/4," & _
                                                "6/3,6/2,6/3/2,6/2/2,5/0,5/5,5/4," & _
                                                "5/3,5/2,5/4/2,5/3,5/3/2,5/2/2,5/" & _
                                                "2/2/2,4/0,4/4,4/2,4/3,4/4/3,4/4/2," & _
                                                "4/3/2,4/2/2,4/3/2/2,4/2/2/2,3/0,3/2," & _
                                                "3/3,3/3/2,3/3/2/2,3/3/2/2/2,3/2/2/2/2," & _
                                                "2/0,2/2,2/2/2,2/2/2/2,2/2/2/2/2"

Private Const FLTRS_DECENAS_03 As String = "3/0,2/1,1/1/1"
Private Const FLTRS_DECENAS_04 As String = "4/0,3/1,2/2,2/1/1,1/1/1/1"
Private Const FLTRS_DECENAS_05 As String = "5/0,4/1,3/2,3/1/1,2/2/1,2/1/1/1,1/1/1/1/1"
Private Const FLTRS_DECENAS_06 As String = "6/0,5/1,4/2,4/1/1,3/3,3/2/1,3/1/1/1,2/2/2,2/2/1/1,2/1/1/1/1"
Private Const FLTRS_DECENAS_07 As String = "7/0,6/1,5/2,5/1/1,4/3,4/2/1,4/1/1/1,3/3/1,3/2/2,3/2/" & _
                                           "1/1,3/1/1/1/1,2/2/2/1,2/2/1/1/1"
Private Const FLTRS_DECENAS_08 As String = "8/0,7/1,6/2,6/1/1,5/3,5/2/1,5/1/1/1,4/4,4/3/1,4/2/2," & _
                                           "4/2/1/1,4/1/1/1/1,3/3/2,3/3/1/1,3/2/2/1,3/2/1/1/1,3/1/1" & _
                                           "/1/1/1,2/2/2/2,2/2/2/1/1,2/2/1/1/1/1"
Private Const FLTRS_DECENAS_09 As String = "9/0,8/1,7/2,7/1/1,6/3,6/2/1,6/1/1/1,5/4,5/3/1,5/2/2," & _
                                           "5/2/1/1,5/1/1/1/1,4/4/1,4/3/2,4/3/1/1,4/2/2/1,4/2/1/1/1" & _
                                           ",3/3/3,3/3/2/1,3/3/1/1/1,3/2/2/2,3/2/2/1/1,2/2/2/2/1"
Private Const FLTRS_DECENAS_10 As String = "10/0,9/1,8/2,8/1/1,7/3,7/2/1,7/1/1/1,6/4,6/3/1,6/2/2" & _
                                           ",6/2/1/1,6/1/1/1/1,5/5,5/4/1,5/3/2,5/3/1/1,5/2/2/1,5/2" & _
                                           "/1/1/1,4/4/2,4/3/2/1,4/3/1/1/1,4/2/2/2,4/2/2/1/1,3/3/3/1" & _
                                           ",3/3/2/1/1,3/2/2/2/2"
Private Const FLTRS_DECENAS_11 As String = "11/0,10/1,9/2,9/1/1,8/3,8/2/1,8/1/1/1,7/4,7/3/1,7/2/" & _
                                           "2,7/2/1/1,7/1/1/1/1,6/5,6/4/1,6/3/2,6/2/2/1,6/2/1/1/1," & _
                                           "5/5/1,5/4/2,5/4/1/1,5/3/3,5/3/2/1,5/3/1/1/1,5/2/2/2,5/2/2/1/1,4" & _
                                           "/4/3,4/4/2/1,4/4/1/1/1,4/3/3/1,4/3/2/2,4/3/2/1/1,3/3/3/2" & _
                                           ",3/3/3/1/1,3/2/2/2/2"
Private Const FLTRS_SEPTENAS_03 As String = "3/0,2/1,1/1/1"
Private Const FLTRS_SEPTENAS_04 As String = "4/0,3/1,2/2,2/1/1,1/1/1/1"
Private Const FLTRS_SEPTENAS_05 As String = "5/0,4/1,3/2,3/1/1,2/2/1,2/1/1/1,1/1/1/1/1"
Private Const FLTRS_SEPTENAS_06 As String = "6/0,5/1,4/2,4/1/1,3/3,3/2/1,3/1/1/1,2/2/2,2/2/1/1,2" & _
                                            "/1/1/1/1,1/1/1/1/1/1"
Private Const FLTRS_SEPTENAS_07 As String = "7/0,6/1,5/2,5/1/1,4/3,4/2/1,4/1/1/1,3/3/1,3/2/2,3/2" & _
                                            "/1/1,3/1/1/1/1,2/2/2/1,2/2/1/1/1,2/1/1/1/1/1,1/1/1/1/1/1/1"
Private Const FLTRS_SEPTENAS_08 As String = "8/0,7/1,6/2,6/1/1,5/3,5/2/1,5/1/1/1,4/4,4/3/1,4/2/2" & _
                                            ",4/2/1/1,4/1/1/1/1,3/3/2,3/3/1/1,3/2/2/1,3/2/1/1/1,3/1" & _
                                            "/1/1/1/1,2/2/2/2,2/2/2/1/1,22/1/1/1/1,2/1/1/1/1/1/1,1/1/" & _
                                            "1/1/1/1/1/1"
Private Const FLTRS_SEPTENAS_09 As String = "9/0,8/1,7/2,7/1/1,6/3,6/2/1,6/1/1/1,5/4,5/3/1,5/2/2" & _
                                            ",5/1/1/1,4/4/1,4/3/2,4/3/1/1,4/2/2/1,4/2/1/1/1,4/1/1/1" & _
                                            "/1/1,3/3/3,3/3/2/1,3/3/1/1/1,3/2/2/2,3/2/2/1/1,3/2/1/1/" & _
                                            "1/1,3/1/1/1/1/1/1,2/2/2/2/1,2/2/2/1/1/1,2/2/1/1/1/1/1,2/" & _
                                            "1/1/1/1/1/1/1"
Private Const FLTRS_SEPTENAS_10 As String = "10/0,9/1,8/2,8/1/1,7/3,7/2/1,7/1/1/1,6/4,6/3/1,6/2/" & _
                                            "2,6/2/1/1,6/1/1/1/1,5/5,5/4/1,5/3/2,5/3/1/1,5/2/2/2,5" & _
                                            "/2/2/1/1,5/2/1/1/1/,5/1/1/1/1/1/1,4/4/2,4/4/1/1,4/3/2/2," & _
                                            "4/3/2/1/1,4/3/1/1/1/1,4/2/2/2/2,4/2/2/2/1/1,4/2/2/1/1/1/" & _
                                            "1,4/2/1/1/1/1/1/1,4/1/1/1/1/1/1/1/1,3/3/3/1,3/3/2/2,3/3/" & _
                                            "2/1/1,3/3/1/1/1/1,3/2/1/1/1/1/1,3/1/1/1/1/1/1/1,2/2/2/2/2" & _
                                            ",2/2/2/2/1/1,2/2/2/1/1/1/1,2/2/1/1/1/1/1/1"
Private Const FLTRS_SEPTENAS_11 As String = "11/0,10/1,9/2,9/1/18/3,8/2/1,8/1/1/1,7/4,7/3/1,7/2/" & _
                                            "2,7/2/1/1,7/1/1/1/1,6/5,6/4/1,6/3/2,6/3/1/1,6/2/2/1,6" & _
                                            "/2/1/1/1,6/1/1/1/1/1,5/5/1,5/4/2,5/4/1/1,5/3/2,5/3/1/1," & _
                                            "5/2/2/2,5/2/2/1/1,5/2/1/1/1/1,5/1/1/1/1/1/1,4/4/3,4/4/2" & _
                                            "/1,4/4/1/1/1,4/3/3/1/0,4/3/2/2,4/2/2/2/1,4/2/2/1/1/1,4/" & _
                                            "2/1/1/1/1/1,4/1/1/1/1/1/1/1,3/3/3/2,3/3/3/1/1,3/3/2/2/1," & _
                                            "3/3/2/1/1/1,3/3/1/1/1/1/1,3/2/1/1/1/1/1/1,2/2/2/2/2/1,2/2" & _
                                            "/2/2/1/1/1,2/2/2/1/1/1/1/1"
Private Const FLTRS_TERMINACIONES_03 As String = "3/0,2/1,1/1/1"
Private Const FLTRS_TERMINACIONES_04 As String = "4/0,3/1,2/2,2/1/1,1/1/1/1"
Private Const FLTRS_TERMINACIONES_05 As String = "5/0,4/1,3/2,3/1/1,2/2/1,2/1/1/1,1/1/1/1/1"
Private Const FLTRS_TERMINACIONES_06 As String = "6/0,5/1,4/2,4/1/1,3/2/1,3/1/1/1,2/2/2,2/2/1/1,2" & _
                                            "/1/1/1/1,1/1/1/1/1/1"
Private Const FLTRS_TERMINACIONES_07 As String = "7/0,6/1,5/2,5/1/1,4/2/1,4/1/1/1,3/2/2,3/2/1/1,3" & _
                                            "/1/1/1/1,2/1/1/1/1/1,1/1/1/1/1/1/1"
Private Const FLTRS_TERMINACIONES_08 As String = "8/0,7/1,6/2,6/1/1,5/3,5/2/1,5/1/1/1,4/4,4/3/1," & _
                                            "4/2/2,4/2/1/1,4/1/1/1/1,3/2/2/1,3/2/1/1/1,3/1/1/1/1/1,2" & _
                                            "/2/2/2,2/2/2/1/1,2/2/1/1/1/1,2/1/1/1/1/1/1"

Private Const FLTRS_TERMINACIONES_09 As String = "9/0,8/1,7/2,7/1/1,6/3,6/2/1,6/1/1/1,5/4,5/3/1," & _
                                            "5/2/2,5/2/1/1,5/1/1/1/1,4/3/2,4/2/2/1,4/2/1/1/1,4/1/1/1" & _
                                            "/1/1,3/2/2/2,3/2/2/1/1,3/2/1/1/1/1,3/1/1/1/1/1/1,2/2/2/2" & _
                                            "/1,2/2/2/1/1/1,2/2/1/1/1/1/1,2/1/1/1/1/1/1/1,1/1/1/1/1/1/1/1/1"

Private Const FLTRS_TERMINACIONES_10 As String = "10/0,9/1,8/2,8/1/1,7/3,7/2/1,7/1/1/1,6/4,6/3/1," & _
                                            "6/2/2,6/2/1/1,6/1/1/1/1,5/5,5/4/1,5/3/2,5/2/2/1,5/2/1" & _
                                            "/1/1,5/1/1/1/1/1,4/4/2,4/3/3,4/3/2/1,4/2/2/2,4/2/2/1/1," & _
                                            "4/2/1/1/1/1,4/1/1/1/1/1/1,3/3/3/1,3/3/2/1/1,3/2/2/2/1,3" & _
                                            "/2/2/1/1/1,3/2/1/1/1/1/1,3/1/1/1/1/1/1/1,2/2/2/2/2,2/2/2/" & _
                                            "2/1/1,2/2/2/1/1/1/1,2/2/1/1/1/1/1/1,2/1/1/1/1/1/1/1/1,1/1/1/1/1/1/1/1/1/1"

Private Const FLTRS_TERMINACIONES_11 As String = "11/0,10/1,9/2,9/1/1,8/3,8/2/1,8/1/1/1,7/4,7/3/1" & _
                                            ",7/2/2,7/2/1/1,7/1/1/1/1,6/5,6/4/1,6/3/2,6/2/1/1/1,6/" & _
                                            "2/2/1,6/1/1/1/1/1,5/5/1,5/4/2,5/3/3,5/3/2/1,5/2/2/2,5/" & _
                                            "2/2/1/1,5/2/1/1/1/1,5/1/1/1/1/1/1,4/4/3,4/3/2/2,4/3/3/1," & _
                                            "4/3/2/1/1,4/2/2/2/1,4/2/2/1/1/1,4/2/1/1/1/1/1,4/1/1/1/1/" & _
                                            "1/1/1,3/3/3/2,3/3/2/2/1,3/3/2/1/1/1,3/3/1/1/1/1/1,3/2/2/" & _
                                            "2/2,3/2/2/2/1/1,3/2/2/1/1/1/1,3/2/1/1/1/1/1/1,3/1/1/1/1/1" & _
                                            "/1/1/1,2/2/2/2/2/1,2/2/2/2/1/1/1,2/2/2/1/1/1/1/1,2/2/1/1/" & _
                                            "1/1/1/1/1,2/1/1/1/1/1/1/1/1/1"



'--- Propiedades --------------------------------------------------------------*
Public TipoFiltro As TipoFiltro
Public MultiplesFiltros As Boolean

'------------------------------------------------------------------------------*
' Procedimiento  : NameFiltro
' Fecha          : vi., 10/abr/2020 19:39:32
' Propósito      : Nombre del tipo de filtro a aplicar
'------------------------------------------------------------------------------*
Public Property Get NameFiltro() As String
    If Len(mNameFiltro) = 0 And TipoFiltro > 0 Then
        mNameFiltro = mTextoTipoFiltro(TipoFiltro - 1)
    End If
    NameFiltro = mNameFiltro
End Property


Public Property Let NameFiltro(ByVal vNewValue As String)
    mNameFiltro = vNewValue
    TipoFiltro = SetTipoFiltro(mNameFiltro)
End Property

'------------------------------------------------------------------------------*
' Procedimiento  : FilterValue
' Fecha          : vi., 10/abr/2020 19:39:32
' Propósito      : Nombre del tipo de filtro a aplicar
'------------------------------------------------------------------------------*
Public Property Get FilterValue() As String
    Dim mValues As String
    Dim i       As Integer
On Error GoTo gFilterValue_Error
    '
    '   Si hay múltiples valores en un tipo de filtro
    '
    mValues = ""
    i = 1
    If MultiplesFiltros Then
        mValues = ""
        For i = 1 To mColValues.Count Step 1
            If i <> mColValues.Count Then
                mValues = mValues & mColValues.Item(i) & ","
            Else
                mValues = mValues & mColValues.Item(i)
            End If
        Next i
    Else
        If mColValues.Count > 0 Then
            mValues = mColValues.Item(i)
        End If
    End If
    FilterValue = mValues
  On Error GoTo 0
gFilterValue__CleanExit:
    Exit Function
gFilterValue_Error:
       Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    Call HandleException(ErrNumber, ErrDescription, "FiltroCombinacion.FilterValue", ErrSource)
    Err.Raise ErrNumber, "FiltroCombinacion.FilterValue", ErrDescription
End Property

Public Property Let FilterValue(vNewValue As String)
    Dim mVar As Variant
    Dim mStr As String
    Dim i As Integer

On Error GoTo lFilterValue_Error
    mVar = Split(vNewValue, ",")
    For i = 0 To UBound(mVar)
        mStr = CStr(mVar(i))
        Me.AddFilterValue mStr
    Next i
    
  On Error GoTo 0
lFilterValue__CleanExit:
    Exit Property
lFilterValue_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    Call HandleException(ErrNumber, ErrDescription, "FiltroCombinacion.FilterValue", ErrSource)
    Err.Raise ErrNumber, "FiltroCombinacion.FilterValue", ErrDescription
End Property


'--- Métodos Privados ---------------------------------------------------------*
Private Sub Class_Initialize()
    mTextoTipoFiltro = Split(NOMBRES_TIPOS_FILTRO, ";")
    Set mColValues = New Collection
    TipoFiltro = 0
    MultiplesFiltros = False
End Sub

Private Sub Class_Terminate()
    Set mColValues = Nothing
End Sub

'------------------------------------------------------------------------------*
' Procedimiento  : SetTipoFiltro
' Fecha          : lu., 02/sep/2019 13:19:29
' Propósito      : convertir una cadena a Tipo de Filtro
' Parámetros     : Cadena Filtro
'------------------------------------------------------------------------------*
Private Function SetTipoFiltro(ByVal vNewValue As String) As TipoFiltro
    Dim i As Integer
  On Error GoTo SetTipoFiltro_Error
    For i = 0 To UBound(mTextoTipoFiltro)
        If vNewValue = mTextoTipoFiltro(i) Then
            SetTipoFiltro = i + 1
            Exit For
        End If
    Next i
    
  On Error GoTo 0
SetTipoFiltro__CleanExit:
    Exit Function
SetTipoFiltro_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    Call HandleException(ErrNumber, ErrDescription, "FiltroCombinacion.SetTipoFiltro", ErrSource)
    Err.Raise ErrNumber, "FiltroCombinacion.SetTipoFiltro", ErrDescription
End Function


'------------------------------------------------------------------------------*
' Procedimiento  : EvalSuma
' Fecha          : mi., 21/ago/2019 23:56:26
' Propósito      : Devuelve un rango de datos en los que se clasifica la suma
'                  de los numeros
' Parámetros     : Suma de la combinación
'------------------------------------------------------------------------------*
Private Function EvalSuma(vNewValue As Integer) As String
    Dim mVar As Variant
    
  On Error GoTo EvalSuma_Error
    mVar = Split(Me.FilterValue, "..")
    mRgMinSuma = CInt(mVar(0))
    mRgMaxSuma = CInt(mVar(1))
    
    If (mRgMinSuma <= vNewValue) And (vNewValue <= mRgMaxSuma) Then
        EvalSuma = FilterValue
    Else
        EvalSuma = CStr(vNewValue)
    End If
  On Error GoTo 0
EvalSuma__CleanExit:
    Exit Function
    
EvalSuma_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    Call HandleException(ErrNumber, ErrDescription, "FiltroCombinacion.EvalSuma", ErrSource)
    Err.Raise ErrNumber, "FiltroCombinacion.EvalSuma", ErrDescription
End Function



'------------------------------------------------------------------------------*
' Procedimiento  : EvaluarOR
' Fecha          : ma., 03/sep/2019 19:23:24
' Propósito      : Evalua múltiples filtros del mismo tipo
' Parámetros     : datFiltro
' Retorno        : True si el filtro esta contenido en la colección
'------------------------------------------------------------------------------*
Private Function EvaluarOR(datFiltro As String) As Boolean
    Dim mVar As Variant
    Dim mBool As Boolean
    Dim i As Integer
    
    mBool = False
    mVar = Split(Me.FilterValue, ",")
    For i = 0 To UBound(mVar)
        mBool = (datFiltro = mVar(i)) Or mBool
    Next i
    EvaluarOR = mBool
End Function



'--- Métodos Públicos ---------------------------------------------------------*
'------------------------------------------------------------------------------*
' Procedimiento  : Parse
' Fecha          : mi., 21/ago/2019 23:56:26
' Propósito      : convertir una cadena que expresa un filtro en una clase
' Parámetros     : Cadena Filtro
'------------------------------------------------------------------------------*
Public Sub Parse(vNewValue As String)
    Dim mStr As String
    Dim mVar As Variant
    Dim i As Integer
    Dim j As Integer
  On Error GoTo Parse_Error
    '
    '   "(#1)#2:[#3]"
    '
    mStr = Replace(vNewValue, "(", "")
    mStr = Replace(mStr, ")", " ")
    mStr = Replace(mStr, ":", " ")
    mStr = Replace(mStr, "[", "")
    mStr = Replace(mStr, "]", "")
    mVar = Split(mStr, " ")
    TipoFiltro = CInt(mVar(0))
    NameFiltro = Trim(mVar(1))
    FilterValue = CStr(mVar(2))
  
  On Error GoTo 0
Parse__CleanExit:
    Exit Sub
    
Parse_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    Call HandleException(ErrNumber, ErrDescription, "File.Parse", ErrSource)
    Err.Raise ErrNumber, "File.Parse", ErrDescription
End Sub



'------------------------------------------------------------------------------*
' Procedimiento  : EvaluarCombinacion
' Fecha          : mi., 21/ago/2019 23:56:26
' Propósito      : Evalua si la formula de la combinación está contenida en la
'                  colección de filtros
' Parámetros     : Combinación
' Retorno        : True si la combinación cumple los filtros
'------------------------------------------------------------------------------*
Public Function EvaluarCombinacion(vNewValue As Combinacion) As Boolean
    Dim mStr  As String
    
  On Error GoTo EvaluarCombinacion_Error
     
     Select Case TipoFiltro
        Case tfParidad: mStr = vNewValue.FormulaParidad
        Case tfAltoBajo: mStr = vNewValue.FormulaAltoBajo
        Case tfConsecutivos: mStr = vNewValue.FormulaConsecutivos
        Case tfDecenas: mStr = vNewValue.FormulaDecenas
        Case tfSeptenas: mStr = vNewValue.FormulaSeptenas
        Case tfSuma: mStr = EvalSuma(vNewValue.Suma)
        Case tfTerminaciones: mStr = vNewValue.FormulaTerminaciones
     End Select
     
     If Me.MultiplesFiltros Then
        EvaluarCombinacion = EvaluarOR(mStr)
     ElseIf Me.FilterValue = mStr Then
        EvaluarCombinacion = True
     Else
        EvaluarCombinacion = False
     End If
     
  On Error GoTo 0
EvaluarCombinacion__CleanExit:
    Exit Function
    
EvaluarCombinacion_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    Call HandleException(ErrNumber, ErrDescription, "FiltroCombinacion.EvaluarCombinacion", ErrSource)
    Err.Raise ErrNumber, "FiltroCombinacion.EvaluarCombinacion", ErrDescription
End Function





'------------------------------------------------------------------------------*
' Procedimiento  : GetValoresFiltros
' Fecha          : ma., 03/sep/2019 19:23:24
' Propósito      : Generar los posibles valores de un filtro
' Parámetros     : Numero de Sugerencias
' Retorno        : Matriz de filtros para un numero de sugerencias
'------------------------------------------------------------------------------*
Public Function GetValoresFiltros(datSugerencias As Integer) As Variant
    Dim mVar As Variant
    Dim mMtrz() As Variant
    Dim mFilt As String
    Dim k   As Integer
    Dim i   As Integer
    Dim j   As Integer
    
  On Error GoTo GetValoresFiltros_Error
    
    k = 0
    Select Case TipoFiltro
        Case tfParidad, tfAltoBajo:
            For i = datSugerencias To 0 Step -1
                '
                j = datSugerencias - i
                mFilt = i & "/" & j
                '
                ReDim Preserve mMtrz(k)
                mMtrz(k) = mFilt
                k = k + 1
            Next i
            mVar = mMtrz
            
        Case tfConsecutivos:
            Select Case datSugerencias
                Case 3: mVar = Split(FLTRS_CONSECUTIVOS_03, ",")
                Case 4: mVar = Split(FLTRS_CONSECUTIVOS_04, ",")
                Case 5: mVar = Split(FLTRS_CONSECUTIVOS_05, ",")
                Case 6: mVar = Split(FLTRS_CONSECUTIVOS_06, ",")
                Case 7: mVar = Split(FLTRS_CONSECUTIVOS_07, ",")
                Case 8: mVar = Split(FLTRS_CONSECUTIVOS_08, ",")
                Case 9: mVar = Split(FLTRS_CONSECUTIVOS_09, ",")
                Case 10: mVar = Split(FLTRS_CONSECUTIVOS_10, ",")
                Case 11: mVar = Split(FLTRS_CONSECUTIVOS_11, ",")
            End Select
        
        Case tfSuma:
            ' #TODO: Analizar datSugerencias para adecuar a los rangos de suma
            For i = 0 To 300 Step 50
                j = i + 50 - 1
                mFilt = CStr(i) & ".." & CStr(j)
                '
                ReDim Preserve mMtrz(k)
                mMtrz(k) = mFilt
                k = k + 1
            Next i
            mVar = mMtrz
            
        Case tfDecenas:
            Select Case datSugerencias
                Case 3: mVar = Split(FLTRS_DECENAS_03, ",")
                Case 4: mVar = Split(FLTRS_DECENAS_04, ",")
                Case 5: mVar = Split(FLTRS_DECENAS_05, ",")
                Case 6: mVar = Split(FLTRS_DECENAS_06, ",")
                Case 7: mVar = Split(FLTRS_DECENAS_07, ",")
                Case 8: mVar = Split(FLTRS_DECENAS_08, ",")
                Case 9: mVar = Split(FLTRS_DECENAS_09, ",")
                Case 10: mVar = Split(FLTRS_DECENAS_10, ",")
                Case 11: mVar = Split(FLTRS_DECENAS_11, ",")
            End Select
            
        Case tfSeptenas
            Select Case datSugerencias
                Case 3: mVar = Split(FLTRS_SEPTENAS_03, ",")
                Case 4: mVar = Split(FLTRS_SEPTENAS_04, ",")
                Case 5: mVar = Split(FLTRS_SEPTENAS_05, ",")
                Case 6: mVar = Split(FLTRS_SEPTENAS_06, ",")
                Case 7: mVar = Split(FLTRS_SEPTENAS_07, ",")
                Case 8: mVar = Split(FLTRS_SEPTENAS_08, ",")
                Case 9: mVar = Split(FLTRS_SEPTENAS_09, ",")
                Case 10: mVar = Split(FLTRS_SEPTENAS_10, ",")
                Case 11: mVar = Split(FLTRS_SEPTENAS_11, ",")
            End Select
        
        Case tfTerminaciones
            Select Case datSugerencias
                Case 3: mVar = Split(FLTRS_TERMINACIONES_03, ",")
                Case 4: mVar = Split(FLTRS_TERMINACIONES_04, ",")
                Case 5: mVar = Split(FLTRS_TERMINACIONES_05, ",")
                Case 6: mVar = Split(FLTRS_TERMINACIONES_06, ",")
                Case 7: mVar = Split(FLTRS_TERMINACIONES_07, ",")
                Case 8: mVar = Split(FLTRS_TERMINACIONES_08, ",")
                Case 9: mVar = Split(FLTRS_TERMINACIONES_09, ",")
                Case 10: mVar = Split(FLTRS_TERMINACIONES_10, ",")
                Case 11: mVar = Split(FLTRS_TERMINACIONES_11, ",")
            End Select
    End Select
    '
    '
    '
    GetValoresFiltros = mVar
  
  On Error GoTo 0
GetValoresFiltros__CleanExit:
    Exit Function
    
GetValoresFiltros_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    Call HandleException(ErrNumber, ErrDescription, "FiltroCombinacion.GetValoresFiltros", ErrSource)
    Err.Raise ErrNumber, "FiltroCombinacion.GetValoresFiltros", ErrDescription
End Function



'------------------------------------------------------------------------------*
' Procedimiento  : ToString
' Fecha          : vi., 10/abr/2020 19:39:32
' Propósito      : Expresa el filtro en formato cadena
'                      (#tipo) #NombreFiltro: [#Valor,#valor,#valor];
'------------------------------------------------------------------------------*
Public Function ToString() As String
    Dim mStr As String
    Dim mValues As String
    Dim i As Integer
    
On Error GoTo ToString_Error
    '
    '   Establecemos el patrón
    '
    mStr = "(#1)#2:[#3]"
    mStr = Replace(mStr, "#1", CStr(TipoFiltro))
    mStr = Replace(mStr, "#2", NameFiltro)
    mStr = Replace(mStr, "#3", FilterValue)
    '
    '   Devolvemos la cadena
    '
    ToString = mStr
  
  On Error GoTo 0
ToString__CleanExit:
    Exit Function
    
ToString_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    Call HandleException(ErrNumber, ErrDescription, "FiltroCombinacion.ToString", ErrSource)
    Err.Raise ErrNumber, "FiltroCombinacion.ToString", ErrDescription
End Function


Public Sub AddFilterValue(datValue As String)
  On Error GoTo AddFilterValue_Error
    '
    '   Agregamos el valor a la colección
    '
    If mColValues.Count = 0 Then
        mColValues.Add datValue
        MultiplesFiltros = False
    Else
        mColValues.Add datValue, , mColValues.Count
        MultiplesFiltros = True
    End If
  
  On Error GoTo 0
AddFilterValue__CleanExit:
    Exit Sub
    
AddFilterValue_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    Call HandleException(ErrNumber, ErrDescription, "FiltroCombinacion.ToString", ErrSource)
    Err.Raise ErrNumber, "FiltroCombinacion.ToString", ErrDescription
End Sub

' *===========(EOF): FiltroCombinacion.cls

