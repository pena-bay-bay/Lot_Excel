Attribute VB_Name = "Módulo1"
Option Explicit

Sub CrearBoton()
Attribute CrearBoton.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro1 Macro
'

'
    ActiveSheet.Buttons.Add(620.25, 64.5, 60.75, 30).Select
    Selection.OnAction = "QuitarBoton"
    Selection.Characters.Text = "Metodos"
    Debug.Print "Nombre Boton" & Selection.Name
    With Selection.Characters(Start:=1, Length:=7).Font
        .Name = "Calibri"
        .FontStyle = "Normal"
        .Size = 11
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ColorIndex = 1
    End With

End Sub
Sub QuitarBoton()
Attribute QuitarBoton.VB_ProcData.VB_Invoke_Func = " \n14"
'
' QuitarBoton Macro
'

'
    ActiveSheet.Shapes.Range(Array("Button 1")).Select
    Selection.Delete
    ActiveSheet.Shapes.Range(Array("Button 2")).Select
    Selection.Delete
    Selection.Cut
    Range("M9").Select
    ActiveSheet.Shapes("Button 4").IncrementLeft 54.1666141732
    ActiveSheet.Shapes("Button 4").IncrementTop 45.8333858268
    Range("L5").Select
    ActiveSheet.Shapes.Range(Array("Button 3")).Select
    Range("M5").Select
    ActiveSheet.Shapes.Range(Array("Button 5")).Select
    ActiveSheet.Shapes("Button 5").IncrementLeft -140
    ActiveSheet.Shapes("Button 5").IncrementTop 25.8333070866
    ActiveSheet.Shapes.Range(Array("Button 3")).Select
    ActiveSheet.Shapes("Button 3").IncrementLeft -66.6666141732
    ActiveSheet.Shapes("Button 3").IncrementTop 49.1666929134
    ActiveSheet.Shapes("Button 3").IncrementLeft 26.6666141732
    ActiveSheet.Shapes("Button 3").IncrementTop -48.3333070866
    ActiveSheet.Shapes.Range(Array("Button 4")).Select
    ActiveSheet.Shapes.Range(Array("Button 5")).Select
    Selection.Delete
    Selection.Cut
    ActiveSheet.Shapes.Range(Array("Button 3")).Select
    Selection.Delete
    Selection.Cut
    ActiveSheet.Shapes.Range(Array("Button 4")).Select
    Selection.Delete
    Selection.Cut
End Sub
