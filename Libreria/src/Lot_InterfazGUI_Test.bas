Attribute VB_Name = "Lot_InterfazGUI_Test"
' *============================================================================*
' *
' *     Fichero    : Lot_InterfazGUI_Test.bas
' *
' *     Autor      : Carlos Almela Baeza
' *     Creación   : do., 07/abr/2019 19:38:21
' *     Versión    : 1.0
' *     Propósito  : Testear las clases del prototipo
' *
' *============================================================================*
Option Explicit
Option Base 0

'------------------------------------------------------------------------------*
' Procedimiento  : ContextoControllerTest
' Fecha          : 7/abr/2019
' Propósito      : Pruebas unitarias de la clase ContextoController
'------------------------------------------------------------------------------*
'
Public Sub ContextoControllerTest()
    Dim mCtrl    As ContextoController
    
  On Error GoTo ContextoControllerTest_Error
    '
    '
    Err.Raise ERR_TODO, "Lot_InterfazGUI_Test.ContextoControllerTest", MSG_TODO
    '
  On Error GoTo 0
ContextoControllerTest__CleanExit:
    Exit Sub
            
ContextoControllerTest_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    '   Audita el error
    Call HandleException(ErrNumber, ErrDescription, "Lot_InterfazGUI_Test.ContextoControllerTest", ErrSource)
    '   Informa del error
    Call MsgBox(ErrDescription, vbCritical Or vbSystemModal, ThisWorkbook.Name)
    Call Trace("CERRAR")
End Sub


' *===========(EOF): <<nombre fichero>>
