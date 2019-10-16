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
' *     Fichero    : Parametro.cls
' *
' *     Autor      : Carlos Almela Baeza
' *     Creación   : sab, 15/Dic/2018 20:52
' *     Versión    : 1.0
' *     Propósito  : Almacenar parametros de configuración
' *
' *============================================================================*
Option Explicit
Option Base 0
'
'--- Variables Privadas -------------------------------------------------------*
Private mValor              As String

'--- Constantes ---------------------------------------------------------------*
Private Const LT_TIPOSVARIABLE As String = "Texto;Entero;Fecha;Hora;Fecha&H" _
                                        & "ora;Decimal;Doble Precisión;Objeto"
Private Const LT_PATRONVARIABLE As String = "#$1 $2 = $3 ($4)"
'--- Mensajes -----------------------------------------------------------------*
'--- Errores ------------------------------------------------------------------*
'--- Propiedades --------------------------------------------------------------*
Public EntidadNegocio       As EntidadNegocio
Public Tipo                 As TipoParametro
Public Descripcion          As String
Public Orden                As Integer
Public Nombre               As String

'------------------------------------------------------------------------------*
' Procedimiento  : TipoParametro
' Fecha          : 15/12/2018
' Propósito      : Enumeración del típo de parámetro
'------------------------------------------------------------------------------*
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

'------------------------------------------------------------------------------*
' Procedimiento  : ID
' Fecha          : 15/12/2018
' Propósito      : Identificador único
'------------------------------------------------------------------------------*
'
Public Property Get Id() As Long
    Id = Me.EntidadNegocio.Id
End Property
Public Property Let Id(parId As Long)
    Me.EntidadNegocio.Id = parId
End Property
'------------------------------------------------------------------------------*
' Procedimiento  : FechaModificacion
' Fecha          : 15/12/2018
' Propósito      : Fecha de modificación del parámetro
'------------------------------------------------------------------------------*
'
Public Property Get FechaModificacion() As Date
    FechaModificacion = Me.EntidadNegocio.FechaModificacion
End Property

Public Property Let FechaModificacion(vNewValue As Date)
    Me.EntidadNegocio.FechaModificacion = vNewValue
End Property
'------------------------------------------------------------------------------*
' Procedimiento  : FechaAlta
' Fecha          : 15/12/2018
' Propósito      : Fecha de registro del parámetro en la Base de datos
'------------------------------------------------------------------------------*
'
Public Property Get FechaAlta() As Date
    FechaAlta = Me.EntidadNegocio.FechaAlta
End Property

Public Property Let FechaAlta(vNewValue As Date)
    Me.EntidadNegocio.FechaAlta = vNewValue
End Property

'------------------------------------------------------------------------------*
' Procedimiento  : Valor
' Fecha          : 15/12/2018
' Propósito      : Valor de la variable
'------------------------------------------------------------------------------*
'
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
            mValor = CStr(vNewValue)
            Tipo = parDecimalPrecision
        Case vbBoolean
            mValor = CStr(vNewValue)
            Tipo = parTexto
        Case vbString
            mValor = vNewValue
            Tipo = parTexto
        Case vbVariant, vbDataObject, vbUserDefinedType, vbArray:
            Err.Raise ERR_VARIABLENOTSERIAL, "Parametro.SetValor", MSG_VARIABLENOTSERIAL
        Case Else
            mValor = vNewValue
            Tipo = parTexto
    End Select
End Property

'--- Metodos Privados ---------------------------------------------------------*
Private Sub Class_Initialize()
    Set EntidadNegocio = New EntidadNegocio
    Valor = ""
    Tipo = parTexto
    Descripcion = ""
    Orden = 1
    Nombre = ""
End Sub

'--- Metodos Publicos ---------------------------------------------------------*
'------------------------------------------------------------------------------*
' Procedimiento  : ToString()
' Fecha          : 15/12/2018
' Propósito      : Devuelve una cadena de texto con el contenido de la clase
' Retorno        : Parametro en formato texto    #Id NOMBRE = Valor ( tipo )
'------------------------------------------------------------------------------*
'
Public Function ToString() As String
    '
    '   Inicializamos la variable
    '
    ToString = ""
    '
    '   Identificador
    '
    ToString = Replace(LT_PATRONVARIABLE, "$1", Me.Id)
    '
    '   Nombre
    '
    ToString = Replace(ToString, "$2", Me.Nombre)
    '
    '   Valor
    '
    ToString = Replace(ToString, "$3", Me.Valor)
    '
    '   Tipo
    '
    ToString = Replace(ToString, "$4", Me.TipoToString)
    
End Function
'------------------------------------------------------------------------------*
' Procedimiento  : TipoToString()
' Fecha          : 15/12/2018
' Propósito      : Devuelve una cadena de texto con el contenido de la clase
' Retorno        : Parametro en formato texto
'------------------------------------------------------------------------------*
'
Public Function TipoToString() As String
    Dim mVar As Variant
    '
    '   Extraemos los nombres del tipo
    '
    mVar = Split(LT_TIPOSVARIABLE, ";")
    '
    '   Segun el tipo devolvemos el nombre
    '
    TipoToString = mVar(Me.Tipo - 1)
   
    
End Function
'' *===========(EOF): Parametro.cls

