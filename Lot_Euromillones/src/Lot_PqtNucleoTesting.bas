Attribute VB_Name = "Lot_PqtNucleoTesting"
'---------------------------------------------------------------------------------------
' Module    : Lot_PqtNucleoTesting
' Author    : CHARLY
' Date      : sáb, 01/nov/2014 21:01:24
' Purpose   :
'---------------------------------------------------------------------------------------
'
Option Explicit
Option Base 0



'---------------------------------------------------------------------------------------
' Procedure : CreatePeriodo
' Author    : CHARLY
' Date      : sáb, 01/nov/2014 21:01:47
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub CreatePeriodo()
    Dim obj As Periodo
    Dim cboPrueba As ComboBox
    Dim frm As frmSelPeriodo
    Dim mLista As Variant
    
    Set obj = New Periodo
    Set frm = New frmSelPeriodo
    Set cboPrueba = frm.cboPerMuestra
    
    mLista = Array(ctPersonalizadas, ctSemanaPasada, ctSemanaActual, ctMesActual, ctHoy, ctAyer, ctLoQueVadeMes, _
                                     ctLoQueVadeSemana)
    
    obj.CargaCombo cboPrueba, mLista
    
    PintarPeriodo obj
    obj.Tipo_Fecha = ctAñoAnterior
    

End Sub


'---------------------------------------------------------------------------------------
' Procedure : BdDatosTest
' Author    : CHARLY
' Date      : mar, 16/jun/2020 13:12:55
' Purpose   : Probar la clase BdDatos
'---------------------------------------------------------------------------------------
'
Private Sub BdDatosTest()
    Dim Bd  As BdDatos
    Dim obj  As Range
    Dim fIni As Date
    Dim fFin As Date
On Error GoTo BdDatosTest_Error
    '
    '  Caso de Prueba 01 Rango de fechas válido
    '
    Set Bd = New BdDatos
    '
    '   Bonoloto y Primitiva
    '
    fIni = #5/28/2020#   ' Sáb
    fFin = #6/13/2020#   ' Sáb
    Set obj = Bd.Resultados_Fechas(fIni, fFin)
    '
    '
    If "$A$1610:$N$1624" = obj.Address And JUEGO_DEFECTO = Bonoloto Then
        Debug.Print ("Rango seleccionado Correcto: " & obj.Address)
    ElseIf "$A$1610:$N$1624" = obj.Address And JUEGO_DEFECTO = LoteriaPrimitiva Then
        Debug.Print ("Rango seleccionado Correcto: " & obj.Address)
    Else
        Debug.Print (" #Error en rango: " & obj.Address)
    End If
    '
    '  Caso de Prueba 02 Rango de fechas NO válido
    '
    fIni = #3/19/2020#   ' Jue (sin Sorteo)
    fFin = #6/13/2020#   ' Sáb
    Set obj = Bd.Resultados_Fechas(fIni, fFin)
    '
    '
    If "$A$1607:$N$1624" = obj.Address And JUEGO_DEFECTO = Bonoloto Then
        Debug.Print ("Rango seleccionado Correcto: " & obj.Address)
    ElseIf "$A$1610:$N$1624" = obj.Address And JUEGO_DEFECTO = LoteriaPrimitiva Then
        Debug.Print ("Rango seleccionado Correcto: " & obj.Address)
    Else
        Debug.Print (" #Error en rango: " & obj.Address)
    End If
    '
    '  Caso de Prueba 03 Rango de fechas NO válido
    '
    fIni = #2/8/2020#    ' Sáb
    fFin = #3/19/2020#   ' Jue Sin Sorteo
    Set obj = Bd.Resultados_Fechas(fIni, fFin)
    '
    '
    If "$A$1576:$N$1606" = obj.Address And JUEGO_DEFECTO = Bonoloto Then
        Debug.Print ("Rango seleccionado Correcto: " & obj.Address)
    ElseIf "$A$1610:$N$1624" = obj.Address And JUEGO_DEFECTO = LoteriaPrimitiva Then
        Debug.Print ("Rango seleccionado Correcto: " & obj.Address)
    Else
        Debug.Print (" #Error en rango: " & obj.Address)
    End If
    
    
  On Error GoTo 0
BdDatosTest_CleanExit:
    Exit Sub
            
BdDatosTest_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    Call HandleException(ErrNumber, ErrDescription, "Lot_PqtNucleoTesting.BdDatosTest_Error", ErrSource)
    Call MsgBox(ErrDescription, vbCritical Or vbSystemModal, Application)
    Call Trace("CERRAR")
End Sub

'---------------------------------------------------------------------------------------
' Procedure : PintarPeriodo
' Author    : CHARLY
' Date      : sáb, 01/nov/2014 21:02:14
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub PintarPeriodo(datPeriodo As Periodo)
    Debug.Print "==> Periodo "
    Debug.Print "Dias          = " & datPeriodo.Dias
    Debug.Print "Fecha Final   = " & datPeriodo.FechaFinal
    Debug.Print "Fecha Inicial = " & datPeriodo.FechaInicial
    Debug.Print "Texto         = " & datPeriodo.Texto
    Debug.Print "Tipo Fecha    = " & datPeriodo.Tipo_Fecha
End Sub



