VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Parametro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' *============================================================================*
' *
' *     Fichero    : Parametro
' *
' *     Autor      : Carlos Almela Baeza
' *     Creación   : Mar, 22/07/2014 18:21
' *     Versión    : 1.0
' *     Propósito  : Almacenar parametros de configuración
' *
' *============================================================================*
Option Explicit
Option Base 0
Public EntidadNegocio As EntidadNegocio
Public Tipo As TipoParametro
Public Descripcion As String
Public Orden As Integer
Public Nombre As String
'
'
'  Tipos Enumerados
'
Public Enum TipoParametro
    parTexto = 1
    parEntero = 2
    parFecha = 3
    parHora = 4
    parFechaHora = 5
    parDecimal = 6
    parDecimalPrecision = 7
    parObjeto = 8
End Enum
'
' variables privadas
'
Private mValor As String

Private Sub Class_Initialize()
    Set EntidadNegocio = New EntidadNegocio
    Valor = ""
    Tipo = parTexto
    Descripcion = ""
    Orden = 1
    Nombre = ""
End Sub


Public Property Get Id() As Long
    Id = Me.EntidadNegocio.Id
End Property

Public Property Get FechaModificacion() As Date
    FechaModificacion = Me.EntidadNegocio.FechaModificacion
End Property

Public Property Get FechaAlta() As Date
    FechaAlta = Me.EntidadNegocio.FechaAlta
End Property



Public Function Parse(datRecord As ADODB.Record) As Parametro

End Function

Public Property Get Valor() As Variant
    Dim mDoble As Double
    Select Case Tipo
        Case parTexto:      Valor = mValor
        Case parEntero:     Valor = CInt(mValor)
        Case parHora:       Valor = CDate(mValor)
        Case parFecha:      Valor = CDate(mValor)
        Case parFechaHora:  Valor = CDate(mValor)
        Case parDecimal:    Valor = CDec(mValor)
        Case parDecimalPrecision
            mDoble = CDbl(mValor)
            Valor = mDoble
        Case Else:      Valor = mValor
    End Select
End Property

Public Property Let Valor(ByVal vNewValue As Variant)
    
    Select Case VarType(vNewValue)
        Case vbEmpty, vbNull:
            mValor = ""
        Case vbInteger, vbByte, vbLong:
            mValor = Str(vNewValue)
            Tipo = parEntero
        Case vbLong:
            mValor = Str(vNewValue)
            Tipo = parEntero
        Case vbSingle, vbDecimal:
            mValor = Str(vNewValue)
            Tipo = parDecimal
        Case vbDouble, vbCurrency
            mValor = Str(vNewValue)
            Tipo = parDecimalPrecision
        Case vbBoolean
            mValor = CStr(vNewValue)
            Tipo = parTexto
        Case vbString
            mValor = vNewValue
            Tipo = parTexto
        Case vbVariant, vbDataObject, vbUserDefinedType, vbArray:
            Err.Raise 1001, "Parametro.SetValor", "Variable no seralizable como parametro"
        Case Else
            mValor = vNewValue
            Tipo = parTexto
    End Select
End Property
