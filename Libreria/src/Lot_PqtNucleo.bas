Attribute VB_Name = "Lot_PqtNucleo"
' *============================================================================*
' *
' *     Fichero    : Lot_PqtNucleo.bas
' *
' *     Autor      : Carlos Almela Baeza
' *     Creación   : vi., 21/sep/2018 17:27:36
' *     Versión    : 1.0
' *     Propósito  : Suministrar las clases a otros libros que referencien
' *                  a este.
' *============================================================================*
Option Explicit
Option Base 0

'--- Errores ------------------------------------------------------------------*
Public Const ERR_REGISTRONOTFOUND As Integer = -100
Public Const ERR_VARIABLENOTSERIAL As Integer = 1001
Public Const ERR_DELETEINDEXERROR As Integer = 1002
Public Const ERR_INDEXERROR As Integer = 1003

'--- Mensajes -----------------------------------------------------------------*
Public Const MSG_REGISTRONOTFOUND As String = "No se ha encontrado el registro #"
Public Const MSG_VARIABLENOTSERIAL As String = "Variable no seralizable como parámetro."
Public Const MSG_DELETEINDEXERROR As String = "No se puede eliminar el elemento # de la colección."
Public Const MSG_INDEXERROR As String = "No se puede acceder el elemento # de la colección."

'--- Literales ----------------------------------------------------------------*
Public Const LT_PARAMSINDESCRIPCION As String = "Parametro sin descripción."
Public Const LT_PARAMSINNOMBRE As String = "Sin Nombre."
Public Const LT_NOMBRESTIPOSPARAMETROS As String = "Texto;Entero;Fecha;Hora;Fecha y hora;Decimal;Decimal de Precisión;Objeto"
Public Const LT_PREMIOSBONOLOTO As String = "1,450000;2,100000;3,1000;4,30;5,4;15,0.5"
Public Const LT_PREMIOSPRIMITIVA  As String = "1,1000000;2,75000;3,3000;4,85;5,8;15,1"
Public Const LT_PREMIOSEUROMILLONES  As String = "1,1000000;2,250000;3,60000;4,2000;5,150;6,100;7,30;8,12;9,12;10,12;11,6;12,6;13,4"
Public Const LT_PREMIOSGORDO As String = "1,1438278.77;2,95015.26;3,6167.81;4,193.01;5,41.13;6,14.48;7,6.44;8,3.00;9,1.5"
 
'--- Enumeraciones ------------------------------------------------------------*
'------------------------------------------------------------------------------*
' Procedimiento  : TiposPeriodos
' Fecha          : 22/04/2007 21:36
' Propósito      : Periodos de tiempo utilizados en el lenguaje natural
'------------------------------------------------------------------------------*
Public Enum TiposPeriodos
    ctSinDefinir = -1
    ctPersonalizadas = 0
    ctSemanaPasada = 1
    ctQuincenaPasada = 2
    ctMesAnterior = 3
    ctAñoAnterior = 4
    ctSemanaActual = 5
    ctQuincenaActual = 6
    ctMesActual = 7
    ctAñoActual = 8
    ctLoQueVadeSemana = 9
    ctLoQueVadeMes = 10
    ctLoQueVadeAño = 11
    ctLoQueVadeTrimestre = 12
    ctUltimaSemana = 13
    ctUltimaQuincena = 14
    ctUltimoMes = 15
    ctUltimoTrimestre = 16
    ctUltimoAño = 17
    ctHastaHoy = 18
    ctHoy = 19
    ctAyer = 20
    ctMañana = 21
End Enum
''------------------------------------------------------------------------------*
'' Procedimiento  : GetColParametros
'' Fecha          : 21/sep/2018
'' Propósito      : Suministrar clases de la libreria LotProject
'' Retorno        : ParametrosProceso
''------------------------------------------------------------------------------*
''
'Public Function GetColParametros() As ParametrosProceso
'    Dim mObj As ParametrosProceso
'    Set mObj = New ParametrosProceso
'    Set GetColParametros = mObj
'End Function
''------------------------------------------------------------------------------*
'' Procedimiento  : GetParamProceso
'' Fecha          : 21/sep/2018
'' Propósito      : Suministrar clases de la libreria LotProject
'' Retorno        : ParamProceso
''------------------------------------------------------------------------------*
''
'Public Function GetParamProceso() As ParamProceso
'    Dim mObj As ParamProceso
'    Set mObj = New ParamProceso
'    Set GetParamProceso = mObj
'End Function
''------------------------------------------------------------------------------*
'' Procedimiento  : GetPoblacion
'' Fecha          : 21/sep/2018
'' Propósito      : Suministrar clases de la libreria LotProject
'' Retorno        : Poblacion
''------------------------------------------------------------------------------*
''
'Public Function GetPoblacion() As Poblacion
'    Dim mObj As Poblacion
'    Set mObj = New Poblacion
'    Set GetPoblacion = mObj
'End Function
''------------------------------------------------------------------------------*
'' Procedimiento  : GetIndividuo
'' Fecha          : 21/sep/2018
'' Propósito      : Suministrar clases de la libreria LotProject
'' Retorno        : Individuo
''------------------------------------------------------------------------------*
''
'Public Function GetIndividuo() As Individuo
'    Dim mObj As Individuo
'    Set mObj = New Individuo
'    Set GetIndividuo = mObj
'End Function
''------------------------------------------------------------------------------*
'' Procedimiento  : GetBombo
'' Fecha          : 21/sep/2018
'' Propósito      : Suministrar clases de la libreria LotProject
'' Retorno        : BomboV2
''------------------------------------------------------------------------------*
''
'Public Function GetBombo() As BomboV2
'    Dim mObj As BomboV2
'    Set mObj = New BomboV2
'    Set GetBombo = mObj
'End Function

' *===========(EOF): Lot_PqtNucleo.bas
