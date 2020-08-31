VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ParamProceso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' *============================================================================*
' *
' *     Fichero    : ParamProceso.cls
' *
' *     Autor      : Carlos Almela Baeza
' *     Creación   : mar, 01/may/2018 23:44:59
' *     Versión    : 1.0
' *     Propósito  : Clase que almacena el valor de un parámetro de  proceso
' *
' *============================================================================*
'   #TODO: Controlar el tipo de variable con excepciones
'
Option Explicit
Option Base 0
'--- Variables Privadas -------------------------------------------------------*
Private mValor As Variant
Private mNombre As String
Private mConcepto As String
Private mTipoVariable As VbVarType

'--- Propiedades --------------------------------------------------------------*
'------------------------------------------------------------------------------*
' Procedimiento  : TipoParam
' Fecha          : 24/may/2018
' Propósito      : Indicador del tipo de variable: 1-Texto, 2-Numero,
'                  3-Fecha, 4-Doble
'------------------------------------------------------------------------------*
'
Public Property Get TipoParam() As Integer
    Select Case mTipoVariable
        Case vbString: TipoParam = 1
        Case vbInteger: mTipoVariable = 2
        Case vbDate: mTipoVariable = 3
        Case vbDouble: mTipoVariable = 4
    End Select
End Property

Public Property Let TipoParam(vNewValue As Integer)
    Select Case vNewValue
        Case 1: mTipoVariable = vbString
        Case 2: mTipoVariable = vbInteger
        Case 3: mTipoVariable = vbDate
        Case 4: mTipoVariable = vbDouble
    End Select
End Property
'------------------------------------------------------------------------------*
' Procedimiento  : Valor
' Fecha          : 01/may/2018
' Propósito      : Almacenamiento del valor de una variable
'------------------------------------------------------------------------------*
'
Public Property Get Valor() As Variant
    Valor = mValor
End Property

Public Property Let Valor(ByVal vNewValue As Variant)
    '
    '   Evaluar el tipo de variable para
    '
    
    mValor = vNewValue
    mTipoVariable = VarType(vNewValue)
End Property
'------------------------------------------------------------------------------*
' Procedimiento  : Concepto
' Fecha          : 01/may/2018
' Propósito      : Almacenamiento del concepto de una variable
'------------------------------------------------------------------------------*
'
Public Property Get Concepto() As String
    Concepto = mConcepto
End Property

Public Property Let Concepto(ByVal vNewValue As String)
    mConcepto = vNewValue
End Property
'------------------------------------------------------------------------------*
' Procedimiento  : Nombre
' Fecha          : 01/may/2018
' Propósito      : Almacenamiento del Nombre de una variable
'------------------------------------------------------------------------------*
'
Public Property Get Nombre() As String
    Nombre = mNombre
End Property

Public Property Let Nombre(ByVal vNewValue As String)
    mNombre = vNewValue
End Property
'------------------------------------------------------------------------------*
' Procedimiento  : ToString
' Fecha          : 06/may/2018
' Propósito      : Devuelve el contenido de la variable en una cadena de texto
'------------------------------------------------------------------------------*
'
Public Function ToString() As String
    Dim mTmp As String
    
    Select Case mTipoVariable
        Case VbVarType.vbCurrency:      mTmp = Format(mValor, "##,##0.00")
        Case VbVarType.vbDecimal:       mTmp = Format(mValor, "0.00%")
        Case VbVarType.vbDouble:        mTmp = Format(mValor, "###0.00")
        Case VbVarType.vbBoolean:       mTmp = CStr(mValor)
        Case VbVarType.vbInteger:       mTmp = Format(mValor, "#0")
        Case VbVarType.vbDate:          mTmp = Format(mValor, "dd/mm/yyyy hh:mm:ss")
        Case VbVarType.vbNull:          mTmp = "#Null"
        Case VbVarType.vbEmpty:         mTmp = "#Vacio"
        Case Else:                      mTmp = mValor
    End Select
    
    ToString = "[" & mNombre & "]= " & mTmp & " (" & mConcepto & ")"
End Function
'------------------------------------------------------------------------------*
' Procedimiento  : Class_Initialize
' Fecha          : 06/may/2018
' Propósito      : Inicializa el contenido de las variables internas
'------------------------------------------------------------------------------*
'
Private Sub Class_Initialize()
    mValor = Null
    mTipoVariable = VarType(mValor)
    mNombre = "#NO_DEFINIDO"
    mConcepto = "Variable por definir"
End Sub

' *===========(EOF): ParamProceso.cls
