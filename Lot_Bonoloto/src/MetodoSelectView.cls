VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MetodoSelectView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' *============================================================================*
' *
' *     Fichero    : MetodoSelectView.cls
' *
' *     Autor      : Carlos Almela Baeza
' *     Creación   : ju., 19/mar/2020 13:25:51
' *     Versión    : 1.0
' *     Propósito  : Caso de Uso Seleccionar Metodos
' *
' *============================================================================*
Option Explicit
Option Base 0


'--- Variables Privadas -------------------------------------------------------*
Private mCtrl               As MetodoController
Private mFechaAnalisis      As Date
Private mPronosticos        As Integer
Private mForm               As frmSugerencia

Public Property Get Form() As Variant
    Set Form = mForm
End Property
Public Property Set Form(datNew As Variant)
    Set mForm = datNew
End Property

'
'
'
Public Property Get cboJuego() As String
    
End Property
'
'
'
Public Property Get FechaAnalisis() As String

End Property
Public Property Let FechaAnalisis(datNew As String)

End Property
Public Property Get FechaAnalisisOk() As Boolean

End Property
'
'
'
Public Property Get Pronosticos() As String

End Property
Public Property Let Pronosticos(datNew As String)

End Property
Public Property Get PronosticossOk() As Boolean

End Property

'
'
'
Public Property Get MtdCtrl() As MetodoController
    Set MtdCtrl = mCtrl
End Property
Public Property Set MtdCtrl(datNew As MetodoController)
    Set mCtrl = datNew
End Property
'
'
'
Public Sub AddMetodo()

End Sub
'
'
'
Public Sub EditMetodo()

End Sub

Public Sub InitForm()
    Dim mVar        As Variant
    Dim i           As Integer
    '
    '   Cargamos el Combo de juegos
    '
    mVar = Split(NOMBRE_JUEGOS, ";")
    For i = 0 To UBound(mVar)
        mForm.cboJuegos.AddItem mVar(i)
    Next i
    '
    '   Establecemos el juego por defecto
    '
    mForm.cboJuegos.ListIndex = JUEGO_DEFECTO - 1
    '
    '   Fijamos el control para que no se modifique
    '
    mForm.cboJuegos.Enabled = False
    
End Sub

Private Sub Class_Initialize()
'    Set mForm = New frmSugerencia
End Sub
