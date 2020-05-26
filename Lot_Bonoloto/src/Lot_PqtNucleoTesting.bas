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
    Dim Obj As Periodo
    Dim cboPrueba As ComboBox
    Dim frm As frmSelPeriodo
    Dim mLista As Variant
    
    Set Obj = New Periodo
    Set frm = New frmSelPeriodo
    Set cboPrueba = frm.cboPerMuestra
    
    mLista = Array(ctPersonalizadas, ctSemanaPasada, ctSemanaActual, ctMesActual, ctHoy, ctAyer, ctLoQueVadeMes, _
                                     ctLoQueVadeSemana)
    
    Obj.CargaCombo cboPrueba, mLista
    
    PintarPeriodo Obj
    Obj.Tipo_Fecha = ctAñoAnterior
    

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



