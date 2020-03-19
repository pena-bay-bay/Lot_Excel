VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSugerencia 
   Caption         =   "Generador de Sugerencias"
   ClientHeight    =   4785
   ClientLeft      =   45
   ClientTop       =   495
   ClientWidth     =   7815
   OleObjectBlob   =   "frmSugerencia.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "frmSugerencia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' *============================================================================*
' *
' *     Fichero    : frmSugerencia.frm
' *
' *     Autor      : Carlos Almela Baeza
' *     Creación   : ma., 10/mar/2020 14:04:34
' *     Versión    : 1.0
' *     Propósito  : Seleccionar los métodos de sugerencias
' *
' *============================================================================*
Option Explicit
Option Base 0

'
'--- Variables Privadas -------------------------------------------------------*
Private DB                          As New BdDatos        ' Base de datos
Private mJuego                      As Juego
Private mFechaAnalisis              As Date
'--- Constantes ---------------------------------------------------------------*
Private Const LT_NOMETODOS As String = "( No existen métodos de sugerencia )"
'--- Mensajes -----------------------------------------------------------------*
'--- Errores ------------------------------------------------------------------*
'--- Propiedades --------------------------------------------------------------*
'--- Métodos Privados ---------------------------------------------------------*
'------------------------------------------------------------------------------*
' Procedimiento  : UserForm_Initialize
' Fecha          : ma., 10/mar/2020 14:11:34
' Propósito      : Inicializa los objetos del formulario
'------------------------------------------------------------------------------*
Private Sub UserForm_Initialize()
    Dim mVar        As Variant
    Dim i           As Integer
    Dim mHoy        As Date
    Dim mInfo       As InfoSorteo
  On Error GoTo UserForm_Initialize_Error
    '
    '   Cargamos el Combo de juegos
    '
    mVar = Split(NOMBRE_JUEGOS, ";")
    For i = 0 To UBound(mVar)
        cboJuegos.AddItem mVar(i)
    Next i
    '
    '   Establecemos el juego por defecto
    '
    cboJuegos.ListIndex = JUEGO_DEFECTO - 1
    '
    '   Fijamos el control para que no se modifique
    '
    cboJuegos.Enabled = False
    '
    '   Establecemos la fecha por defecto del análisis
    '
    mHoy = Date
    Set mInfo = New InfoSorteo
    mInfo.Constructor JUEGO_DEFECTO
    
    If mHoy = DB.UltimoResultado Then
            mFechaAnalisis = mInfo.GetProximoSorteo(mHoy)
    Else
        If mInfo.EsFechaSorteo(mHoy) Then
            mFechaAnalisis = mHoy
        Else
            mFechaAnalisis = mInfo.GetProximoSorteo(mHoy)
        End If
    End If
    txtFechaAnalisis.Text = Format(mFechaAnalisis, "dd/mm/yyyy")
    '
    '   Buscamos metodos y si no hay cargamos el literal
    '
    lstMetodos.AddItem LT_NOMETODOS
    
  On Error GoTo 0
UserForm_Initialize__CleanExit:
    Exit Sub
UserForm_Initialize_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    Call HandleException(ErrNumber, ErrDescription, "frmSugerencia.UserForm_Initialize", ErrSource)
    Err.Raise ErrNumber, "frmSugerencia.UserForm_Initialize", ErrDescription
End Sub

'------------------------------------------------------------------------------*
' Procedimiento  : cboJuegos_Change
' Fecha          : ma., 10/mar/2020 14:12:30
' Propósito      : Seleccionar el juego de
' Parámetros     :
'------------------------------------------------------------------------------*
'Private Sub cboJuegos_Change()
'
'End Sub

'------------------------------------------------------------------------------*
' Procedimiento  : <<nombre propiedad>>
' Fecha          : dd/MMM/yyyy
' Propósito      :
' Parámetros     :
'------------------------------------------------------------------------------*
Private Sub chkSelAllMetodos_Click()
    '
    '   Si está seleccionado desmarcar los metodos
    '
End Sub

'------------------------------------------------------------------------------*
' Procedimiento  : <<nombre propiedad>>
' Fecha          : dd/MMM/yyyy
' Propósito      :
' Parámetros     :
'------------------------------------------------------------------------------*
Private Sub cmdAgregar_Click()

End Sub

'------------------------------------------------------------------------------*
' Procedimiento  : cmdCancelar_Click
' Fecha          : ma., 10/mar/2020 14:12:30
' Propósito      : Cerrar el formulario
'------------------------------------------------------------------------------*
Private Sub cmdCancelar_Click()
    Me.Tag = BOTON_CERRAR
    Me.Hide
End Sub

'------------------------------------------------------------------------------*
' Procedimiento  : <<nombre propiedad>>
' Fecha          : dd/MMM/yyyy
' Propósito      :
' Parámetros     :
'------------------------------------------------------------------------------*
Private Sub cmdEditar_Click()

End Sub

'------------------------------------------------------------------------------*
' Procedimiento  : <<nombre propiedad>>
' Fecha          : dd/MMM/yyyy
' Propósito      :
' Parámetros     :
'------------------------------------------------------------------------------*
Private Sub cmdEjecutar_Click()
    Me.Tag = EJECUTAR
    Me.Hide
End Sub

'------------------------------------------------------------------------------*
' Procedimiento  : <<nombre propiedad>>
' Fecha          : dd/MMM/yyyy
' Propósito      :
' Parámetros     :
'------------------------------------------------------------------------------*
Private Sub cmdGoFirst_Click()

End Sub

'------------------------------------------------------------------------------*
' Procedimiento  : <<nombre propiedad>>
' Fecha          : dd/MMM/yyyy
' Propósito      :
' Parámetros     :
'------------------------------------------------------------------------------*
Private Sub cmdGoLast_Click()

End Sub

'------------------------------------------------------------------------------*
' Procedimiento  : <<nombre propiedad>>
' Fecha          : dd/MMM/yyyy
' Propósito      :
' Parámetros     :
'------------------------------------------------------------------------------*
Private Sub cmdNextPage_Click()

End Sub

'------------------------------------------------------------------------------*
' Procedimiento  : <<nombre propiedad>>
' Fecha          : dd/MMM/yyyy
' Propósito      :
' Parámetros     :
'------------------------------------------------------------------------------*
Private Sub cmdPrevPage_Click()

End Sub

'------------------------------------------------------------------------------*
' Procedimiento  : <<nombre propiedad>>
' Fecha          : dd/MMM/yyyy
' Propósito      :
' Parámetros     :
'------------------------------------------------------------------------------*
Private Sub lstMetodos_Click()

End Sub

'------------------------------------------------------------------------------*
' Procedimiento  : <<nombre propiedad>>
' Fecha          : dd/MMM/yyyy
' Propósito      :
' Parámetros     :
'------------------------------------------------------------------------------*
Private Sub txtFechaAnalisis_Change()

End Sub

'------------------------------------------------------------------------------*
' Procedimiento  : <<nombre propiedad>>
' Fecha          : dd/MMM/yyyy
' Propósito      :
' Parámetros     :
'------------------------------------------------------------------------------*
Private Sub txtPronosticos_Change()

End Sub

'--- Métodos Públicos ---------------------------------------------------------*
'' *===========(EOF): frmSugerencia.frm

