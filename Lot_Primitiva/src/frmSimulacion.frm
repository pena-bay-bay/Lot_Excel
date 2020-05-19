VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSimulacion 
   Caption         =   "Formulario de Simulación"
   ClientHeight    =   4245
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6045
   OleObjectBlob   =   "frmSimulacion.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   2  'Centrar en pantalla
End
Attribute VB_Name = "frmSimulacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False







'---------------------------------------------------------------------------------------
' Module    : frmSimulacion
' DateTime  : 09/dic/2007 20:45
' Author    : Carlos Almela Baeza
' Purpose   : Formulario de captura de datos para las simulaciones
'---------------------------------------------------------------------------------------
Private m_metodo        As metodo
'Private m_periodo       As Periodo
Private m_muestra       As Muestra
'Private m_datos         As BdDatos

Public Property Get Metodo_Simulacion() As metodo
    Metodo_Simulacion = m_metodo
End Property

Public Property Get Muestra_Simulacion() As Muestra
    Muestra_Simulacion = m_muestra
End Property

Private Sub cmdCancelar_Click()
    Me.Tag = BOTON_CERRAR               ' Asigna en la etiqueta del formulario
                                        ' la clave de cerrar
    Me.Hide                             ' Oculta el formulario
End Sub

Private Sub cmdEjecutar_Click()
    Me.Tag = EJECUTAR                   ' Si los datos son correctos se asigna
                                        ' la etiqueta EJECUTAR al formulario
    Me.Hide                             ' Oculta el formulario
End Sub
