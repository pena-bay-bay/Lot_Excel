VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EstadoAplicacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

'---------------------------------------------------------------------------------------
' Modulo    : EstadoAplicacion
' Creado    : 01-nov-2006 17:13
' Autor     : Carlos Almela Baeza
' Version   : 1.0.0 Fecha
' Objeto    : Esta clase guarda el estado inicial de la aplicación
'             y lo restaura al salir
'---------------------------------------------------------------------------------------

Private Type COMMANDBAR_STATE
    sNombre As String
    bVisible As Boolean
End Type

Dim m_atCommandBars()       As COMMANDBAR_STATE
Dim m_nCalculation As Integer
Dim m_bCellDragAndDrop As Boolean
Dim m_bDisplayFormulaBar As Boolean
Dim m_bDisplayNoteIndicator As Boolean
Dim m_bDisplayStatusBar As Boolean
Dim m_bEditDirectlyInCell As Boolean
Dim m_bTransitionNavigationKeys As Boolean

Public Sub ObtenerEstado()
'----------------------------------------
' guarda el estado actual en variables miembro
'----------------------------------------

    Dim objCommandBar As CommandBar
    Dim nCtr As Integer
    
    With Application
        
        ' guarda el modo de cálculo. Fíjese que hay que abrir
        ' un libro, o si no la propiedad Calculation fallara,
        ' por lo que para evitar esto, abriremos un libro nuevo
        .ScreenUpdating = False
        .Workbooks.Add
        m_nCalculation = .Calculation
        ActiveWorkbook.Close False

        m_bCellDragAndDrop = .CellDragAndDrop
        m_bDisplayFormulaBar = .DisplayFormulaBar
        m_bDisplayNoteIndicator = .DisplayNoteIndicator
        m_bDisplayStatusBar = .DisplayStatusBar
        m_bEditDirectlyInCell = .EditDirectlyInCell
        m_bTransitionNavigationKeys = .TransitionNavigKeys
        
        ' guarda la visibilidad de cada barra de ¾rdenes
        ReDim m_atCommandBars(.CommandBars.Count - 1)
        nCtr = 0
        For Each objCommandBar In .CommandBars
            With m_atCommandBars(nCtr)
                .sNombre = objCommandBar.Name
                .bVisible = objCommandBar.Visible
            End With
            nCtr = nCtr + 1
        Next objCommandBar
        
    End With

End Sub

Public Sub OcultarTodasLasBarrasDeHerramientas()
'----------------------------------------
' oculta todas las barras de ordenes
'----------------------------------------
    Dim objCommandBar As CommandBar
    On Error Resume Next
    For Each objCommandBar In Application.CommandBars
        objCommandBar.Visible = False
    Next objCommandBar
End Sub


Public Sub RestablecerEstado()
'----------------------------------------
' restablece el estado guardado por ObtenerEstado
'----------------------------------------
    Dim nCtr As Integer
    
    On Error Resume Next
    
    With Application
        .ScreenUpdating = False
        .Workbooks.Add
        .Calculation = m_nCalculation
        ActiveWorkbook.Close False

        .CellDragAndDrop = m_bCellDragAndDrop
        .DisplayFormulaBar = m_bDisplayFormulaBar
        .DisplayNoteIndicator = m_bDisplayNoteIndicator
        .DisplayStatusBar = m_bDisplayStatusBar
        .EditDirectlyInCell = m_bEditDirectlyInCell
        .TransitionNavigKeys = m_bTransitionNavigationKeys
    End With

    For nCtr = 0 To UBound(m_atCommandBars)
        With m_atCommandBars(nCtr)
            Application.CommandBars(.sNombre).Visible = .bVisible
        End With
    Next nCtr
    Application.CommandBars("Worksheet Menu Bar").Visible = True
End Sub


'---------------------------------------------------------------------------------------
' Procedimiento : Existe_Barra
' Creación      : 03-ago-2006 23:25
' Autor         : Carlos Almela Baeza
' Objeto        : Funcion que devuelve true si está la barra creada
'---------------------------------------------------------------------------------------
'
Public Function Existe_Barra(Nombre_Barra As String) As Boolean
   
   On Error GoTo Existe_Barra_Error

    Existe_Barra = False
    For Each cmdBar In Application.CommandBars
        If cmdBar.Name = Nombre_Barra Then
                Existe_Barra = True
                Exit For
        End If
    Next cmdBar


   On Error GoTo 0
   Exit Function

Existe_Barra_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & _
        ") in procedure Existe_Barra of Módulo de clase EstadoAplicacion"

End Function


