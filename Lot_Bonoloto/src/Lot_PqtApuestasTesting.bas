Attribute VB_Name = "Lot_PqtApuestasTesting"
' *============================================================================*
' *
' *     Fichero    : Lot_PqtApuestasTesting.bas
' *
' *     Autor      : Carlos Almela Baeza
' *     Creación   : jue, 29/mar/2018 20:00:15
' *     Versión    : 1.0
' *     Propósito  : Colección de pruebas del paquete apuestas
' *
' *============================================================================*
Option Explicit
Option Base 0

''------------------------------------------------------------------------------*
'' Procedimiento  : MetodoTest
'' Fecha          : 29/mar/2018
'' Propósito      : Pruebas unitarias de la clase metodo
''------------------------------------------------------------------------------*
''
'Private Sub MetodoTest()
'    Dim oMtdo  As metodo
'
'    Set oMtdo = New metodo
'
''    oMtdo.TipoProcedimiento = mtdAlgoritmoAG
'    Print_Metodo oMtdo
'
'End Sub
'
'Private Sub Print_Metodo(Obj As metodo)
'    Debug.Print "==> Pruebas metodo"
'    Debug.Print vbTab & "EntidadNegocio.FechaAlta   =" & Obj.EntidadNegocio.FechaAlta
''    Debug.Print vbTab & "EsMultiple                 =" & Obj.EsMultiple
'    Debug.Print vbTab & "Parametros                 =" & Obj.ToString
'    Debug.Print vbTab & "TipoMuestra                =" & Obj.TipoMuestra
'    Debug.Print vbTab & "TipoProcedimiento          =" & Obj.TipoProcedimiento
'    Debug.Print vbTab & "TipoProcedimientoTostring  =" & Obj.TipoProcedimientoTostring
'End Sub
'---------------------------------------------------------------------------------------
' Procedure : AciertoTest
' Author    : CHARLY
' Date      : sáb, 01/nov/2014 21:01:34
' Purpose   :
'---------------------------------------------------------------------------------------
'
'Private Sub AciertoTest()
'    Dim Obj As Acierto
'    Set Obj = New Acierto
'    With Obj
'        .ApuestasAcertadas = 1
'        .BolasAcertadas = 2
'        .Categoria = Duodecima
'        .EstrellasAcertadas = 2
'        .IdApuesta = 15
'        .ImportePremio = 12
'        .Juego = Bonoloto
'        .ReintegroAcertado = True
'    End With
'    PrintAcierto Obj
'End Sub

'---------------------------------------------------------------------------------------
' Procedure : ApuestaTest
' Author    : CHARLY
' Date      : sáb, 01/nov/2014 21:01:37
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Sub ApuestaTest()
    Dim oApuesta As Apuesta
    Set oApuesta = New Apuesta
    oApuesta.Combinacion.Texto = "1-2-3-4-5-7"
    
    PrintApuesta oApuesta
    
End Sub

'---------------------------------------------------------------------------------------
' Procedure : CU_ComprobarApuestaTest
' Author    : CHARLY
' Date      : 10/05/2014
' Purpose   :
'
'---------------------------------------------------------------------------------------
'
Private Sub CU_ComprobarApuestaTest()
    Dim Obj As CU_ComprobarApuesta
    Dim oSorteo As Sorteo
    Dim oApuesta As Apuesta
   On Error GoTo CU_ComprobarApuestaTest_Error
    Set Obj = New CU_ComprobarApuesta
    Set oSorteo = New Sorteo
    Set oApuesta = New Apuesta
    With oSorteo
        .EntidadNegocio.Id = 1348
        .Juego = Bonoloto
        .Combinacion.Texto = "1-4-9-23-43-44-31"
        .Reintegro = 4
        .Juego = Bonoloto
        .Fecha = #5/5/2014#
        Debug.Print .Complementario
    End With
    With oApuesta
        .Combinacion.Texto = "1-4-10-19-24-29-31-44"
        .EntidadNegocio.Id = 1
    End With
    
    
    
   On Error GoTo 0
   Exit Sub

CU_ComprobarApuestaTest_Error:
   Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
   ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
   '   Audita el error
   Call HandleException(ErrNumber, ErrDescription, "Lot_Testing.CU_ComprobarApuestaTest")
   Call MsgBox(ErrDescription, vbError Or vbSystemModal, NOMBRE_APLICACION)
   Call Trace("CERRAR")
End Sub

'---------------------------------------------------------------------------------------
' Procedure : PrintApuesta
' Author    : CHARLY
' Date      : sáb, 01/nov/2014 21:02:17
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub PrintApuesta(datApuesta As Apuesta)
    Debug.Print "==> Pintar Combinacion"
    Debug.Print "  Combinacion    = " & datApuesta.Combinacion.Texto
    Debug.Print "  Coste          = " & datApuesta.Coste(Bonoloto)
    Debug.Print "  EntidadNegocio = " & datApuesta.EntidadNegocio.Id
    Debug.Print "  EntidadNegocio = " & datApuesta.EntidadNegocio.ClassStorage
    Debug.Print "  EntidadNegocio = " & datApuesta.EntidadNegocio.FechaAlta
    Debug.Print "  EntidadNegocio = " & datApuesta.EntidadNegocio.FechaBaja
    Debug.Print "  EntidadNegocio = " & datApuesta.EntidadNegocio.FechaModificacion
    Debug.Print "  EntidadNegocio = " & datApuesta.EntidadNegocio.IsDirty
    Debug.Print "  EntidadNegocio = " & datApuesta.EntidadNegocio.IsNew
    Debug.Print "  EntidadNegocio = " & datApuesta.EntidadNegocio.MarkForDelete
    Debug.Print "  EntidadNegocio = " & datApuesta.EntidadNegocio.Origen
    Debug.Print "  EntidadNegocio = " & datApuesta.EntidadNegocio.Situacion
    Debug.Print "  EsMultiple     = " & datApuesta.EsMultiple
    Debug.Print "  FechaAlta      = " & datApuesta.FechaAlta
    Debug.Print "  IdBoleto       = " & datApuesta.IdBoleto
    Debug.Print "  Metodo         = " & datApuesta.Metodo
    Debug.Print "  NumeroApuestas = " & datApuesta.NumeroApuestas
    Debug.Print "  Pronosticos    = " & datApuesta.Pronosticos
    Debug.Print "  SeHaJugado     = " & datApuesta.SeHaJugado
    Debug.Print "  Texto          = " & datApuesta.Texto
End Sub
'---------------------------------------------------------------------------------------
' Procedure : PrintAcierto
' Author    : CHARLY
' Date      : sáb, 01/nov/2014 21:02:07
' Purpose   :
'---------------------------------------------------------------------------------------
'
'Private Sub PrintAcierto(oDat As Acierto)
'    Debug.Print "==> Pruebas Acierto"
'    Debug.Print " .ApuestasAcertadas  = " & oDat.ApuestasAcertadas
'    Debug.Print " .BolasAcertadas     = " & oDat.BolasAcertadas
'    Debug.Print " .Categoria          = " & oDat.Categoria
'    Debug.Print " .EstrellasAcertadas = " & oDat.EstrellasAcertadas
'    Debug.Print " .IdApuesta          = " & oDat.IdApuesta
'    Debug.Print " .ImportePremio      = " & oDat.ImportePremio
'    Debug.Print " .Juego              = " & oDat.Juego
'    Debug.Print " .ReintegroAcertado  = " & oDat.ReintegroAcertado
'End Sub

' *===========(EOF): Lot_PqtApuestasTesting.bas
