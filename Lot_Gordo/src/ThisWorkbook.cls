VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
' *============================================================================*
' *
' *     Fichero    : ThisWorkbook.cls
' *
' *     Autor      : Carlos Almela Baeza
' *     Creación   : dom, 11/mar/2018 23:39:00
' *     Versión    : 1.0
' *     Propósito  : Inicializa la aplicación
' *
' *============================================================================*
Option Explicit
Option Base 0

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    On Error Resume Next
    Fin_Aplicacion
End Sub

Public Sub Pintar_Version()
    Version_Libreria
End Sub

Private Sub Workbook_Open()
    On Error Resume Next
    THISLIBRO = ThisWorkbook.Name
    Crear_Barra_Herramientas BARRA_FUNCIONES
    ThisWorkbook.Sheets(1).Select
End Sub

Private Sub Workbook_SheetActivate(ByVal Sh As Object)
    On Error Resume Next
    Application.CommandBars(BARRA_FUNCIONES).Enabled = True
End Sub

Private Sub Workbook_SheetDeactivate(ByVal Sh As Object)
    On Error Resume Next
    Application.CommandBars(BARRA_FUNCIONES).Enabled = False
End Sub

Private Sub Workbook_WindowActivate(ByVal Wn As Window)
    On Error Resume Next
    Application.CommandBars(BARRA_FUNCIONES).Enabled = True
End Sub

Private Sub Workbook_WindowDeactivate(ByVal Wn As Window)
    On Error Resume Next
    Application.CommandBars(BARRA_FUNCIONES).Enabled = False
End Sub

