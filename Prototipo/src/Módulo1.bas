Attribute VB_Name = "Módulo1"
Option Explicit

Sub Oculta()
Attribute Oculta.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro1 Macro
'

'
    Columns("E:P").Select
    Selection.EntireColumn.Hidden = False
    Columns("Q:AB").Select
    Selection.EntireColumn.Hidden = True
    
End Sub

Sub Muestra()
'
' Macro1 Macro
'

'
    Columns("E:P").Select
    Selection.EntireColumn.Hidden = True
    Columns("Q:AB").Select
    Selection.EntireColumn.Hidden = False
    
End Sub
Sub Botones()
Attribute Botones.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Botones Macro
'

'
    ActiveSheet.Shapes.RAnge(Array("Rounded Rectangle 10")).Select
    Selection.ShapeRange(1).TextFrame2.TextRange.Characters.Text = "refrescar"
    With Selection.ShapeRange(1).TextFrame2.TextRange.Characters(1, 9). _
        ParagraphFormat
        .FirstLineIndent = 0
        .Alignment = msoAlignCenter
    End With
    With Selection.ShapeRange(1).TextFrame2.TextRange.Characters(1, 9).Font
        .NameComplexScript = "+mn-cs"
        .NameFarEast = "+mn-ea"
        .Fill.Visible = msoTrue
        .Fill.ForeColor.ObjectThemeColor = msoThemeColorLight1
        .Fill.ForeColor.TintAndShade = 0
        .Fill.ForeColor.Brightness = 0
        .Fill.Transparency = 0
        .Fill.Solid
        .Size = 11
        .Name = "+mn-lt"
    End With
    RAnge("AC15").Select
End Sub

Sub rutina()
    Dim myDocument As Worksheet
    Dim mShape As Shape
    Set myDocument = Worksheets("Apuestas")
    
    Debug.Print "Total de Shapes: " & myDocument.Shapes.Count
    For Each mShape In myDocument.Shapes
        Debug.Print "Shape.ID        : " & mShape.ID
        Debug.Print "Shape.Name      : " & mShape.Name
        Debug.Print "Shape.Type      : " & mShape.Type
        Debug.Print "AlternativeText      : " & mShape.AlternativeText
        Debug.Print
    Next
End Sub
Option Explicit

Sub Imagen17_Haga_clic_en()
'
' Imagen17_Haga_clic_en Macro
'

'
    ActiveSheet.Shapes.RAnge(Array("ImgNext")).Select
    RAnge("AJ16").Select
    Sheets("Apuestas").Select
    RAnge("AK17").Select
    ActiveWorkbook.Save
End Sub
Option Explicit

Sub ElegirLista()
'
' ElegirLista Macro
'

'
    With Selection.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:="=Periodo"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With
End Sub



