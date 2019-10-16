VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BoletosEngine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' *============================================================================*
' *
' *     Fichero    : BoletosEngine.cls
' *
' *     Autor      : Carlos Almela Baeza
' *     Creación   : sá., 10/ago/2019 19:51:15
' *     Versión    : 1.0
' *     Propósito  : Motor
' *
' *============================================================================*
Option Explicit
Option Base 0
'
'--- Variables Privadas -------------------------------------------------------*
Private mDBase      As DataBaseExcel
Private mCol        As Boletos
Private mObj        As Boleto
Private oNum        As Apuesta
Private mRange      As Range
Private mFila       As Range
Private m_eJuego    As Juego
Private mId         As Integer
Private mTabla      As NombreTabla
'--- Constantes ---------------------------------------------------------------*
Private Const ERR_BOLETONOTFOUND = 3001
Private Const ERR_BOLETOIDNOTFOUND = 3002
Private Const ERR_FECHASBOLETONOTFOUND = 3003
Private Const ERR_BOLETODUPLICADO = 3004
'--- Mensajes -----------------------------------------------------------------*
Private Const MSG_BOLETONOTFOUND As String = "Boleto no encontrando"
Private Const MSG_BOLETOIDNOTFOUND As String = "Boleto de Id: #, no encontrado"
Private Const MSG_FECHASBOLETONOTFOUND  As String = "Boleto de fecha:#, no encontrado"
Private Const MSG_BOLETODUPLICADO  As String = "Boleto de fecha: # ya existe"

'--- Errores ------------------------------------------------------------------*
'--- Propiedades --------------------------------------------------------------*
'--- Métodos Privados ---------------------------------------------------------*
Private Sub Class_Initialize()
    '
    '   Creamos y abrimos la base de datos
    '
    Set mDBase = New DataBaseExcel
    mDBase.Abrir
End Sub

Private Sub Class_Terminate()
    '
    '   Eliminamos el objeto
    '
    Set mDBase = Nothing
End Sub
'--- Métodos Públicos ---------------------------------------------------------*
'------------------------------------------------------------------------------*
' Procedimiento  : GetNewBoleto
' Fecha          : dd/MMM/yyyy
' Propósito      :
' Parámetros     :
' Retorno        :
'------------------------------------------------------------------------------*
Public Function GetNewBoleto(datJuego As Juego) As Boleto

End Function
'------------------------------------------------------------------------------*
' Procedimiento  : GetBoletoId
' Fecha          : dd/MMM/yyyy
' Propósito      :
' Parámetros     :
' Retorno        :
'------------------------------------------------------------------------------*
Public Function GetBoletoId(datId As Integer, datJuego As Juego) As Boleto

End Function
'------------------------------------------------------------------------------*
' Procedimiento  : GetBoletoByFecha
' Fecha          : dd/MMM/yyyy
' Propósito      :
' Parámetros     :
' Retorno        :
'------------------------------------------------------------------------------*
Public Function GetBoletoByFecha(datFecha As Date, datJuego As Juego) As Boleto

End Function
'------------------------------------------------------------------------------*
' Procedimiento  : GetBoletosInIds
' Fecha          : dd/MMM/yyyy
' Propósito      :
' Parámetros     :
' Retorno        :
'------------------------------------------------------------------------------*
Public Function GetBoletosInIds(datIni As Integer, datFin As Integer, _
                                datJuego As Juego) As Boletos

End Function
'------------------------------------------------------------------------------*
' Procedimiento  : GetBoletosInFechas
' Fecha          : dd/MMM/yyyy
' Propósito      :
' Parámetros     :
' Retorno        :
'------------------------------------------------------------------------------*
Public Function GetBoletosInFechas(datFechaIni As Date, datFechaFin As Date, _
                                   datJuego As Juego) As Boletos

End Function
'------------------------------------------------------------------------------*
' Procedimiento  : SetBoleto
' Fecha          : dd/MMM/yyyy
' Propósito      :
' Parámetros     :
' Retorno        :
'------------------------------------------------------------------------------*
Public Sub SetBoleto(datBoleto As Boleto)

End Sub
'------------------------------------------------------------------------------*
' Procedimiento  : SetBoletos
' Fecha          : dd/MMM/yyyy
' Propósito      :
' Parámetros     :
' Retorno        :
'------------------------------------------------------------------------------*
Public Sub SetBoletos(datBoletos As Boletos)

End Sub
'------------------------------------------------------------------------------*
' Procedimiento  : GetFechaUltimoBoleto
' Fecha          : dd/MMM/yyyy
' Propósito      :
' Parámetros     :
' Retorno        :
'------------------------------------------------------------------------------*
Public Function GetFechaUltimoBoleto(datJuego As Juego) As Date

End Function
'------------------------------------------------------------------------------*
' Procedimiento  : GetFechaPrimerBoleto
' Fecha          : dd/MMM/yyyy
' Propósito      :
' Parámetros     :
' Retorno        :
'------------------------------------------------------------------------------*
Public Function GetFechaPrimerBoleto(datJuego As Juego) As Date

End Function
'------------------------------------------------------------------------------*
' Procedimiento  : GetIdPrimerBoleto
' Fecha          : dd/MMM/yyyy
' Propósito      :
' Parámetros     :
' Retorno        :
'------------------------------------------------------------------------------*
Public Function GetIdPrimerBoleto(datJuego As Juego) As Integer

End Function
'------------------------------------------------------------------------------*
' Procedimiento  : GetIdUltimoBoleto
' Fecha          : dd/MMM/yyyy
' Propósito      :
' Parámetros     :
' Retorno        :
'------------------------------------------------------------------------------*
Public Function GetIdUltimoBoleto(datJuego As Juego) As Integer

End Function
'' *===========(EOF): BoletosEngine.cls

