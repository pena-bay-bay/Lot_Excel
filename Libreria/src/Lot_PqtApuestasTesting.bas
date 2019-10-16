Attribute VB_Name = "Lot_PqtApuestasTesting"
' *============================================================================*
' *
' *     Fichero    : Lot_PqtApuestasTesting.bas
' *
' *     Autor      : Carlos Almela Baeza
' *     Creación   : ju., 08/ago/2019 19:50:21
' *     Versión    : 1.0
' *     Propósito  : Colección de pruebas unitarias de las clases del paquete
' *                  Apuestas:
' *                    - Apuesta
' *                    - Apuetas
' *                    - ApuestaEngine
' *                    - Boleto
' *                    - BoletoEngine
' *                    - Boletos
' *                    - Participante
' *                    - ParticipanteEngine
' *
' *============================================================================*
Option Explicit
Option Base 0
'
'
'
Public Sub PqtApuestasTest()
    BoletoTest
    BoletosTest
    BoletoEngineTest
    ApuestaTest
    ApuestasTest
    ApuestaEngineTest
    ParticipanteTest
    ParticipanteEngineTest
End Sub

'---------------------------------------------------------------------------------------
' Procedure : ParticipanteEngineTest
' Author    : CHARLY
' Date      : mi., 14/ago/2019 11:15:31
' Purpose   : Pruebas unitarias de la clase ParticipanteEngine
'---------------------------------------------------------------------------------------
'
Private Sub ParticipanteEngineTest()
    Dim mPrs As Participante
    
  On Error GoTo ParticipanteEngineTest_Error
    '
    '   1.- Objeto en vacio
    '
    Set mPrs = New Participante
    PrintParticipante mPrs
    '
    '   2.- Objeto Valido
    '
    Set mPrs = New Participante
    With mPrs
        .Apellido1 = "Almela"
        .Apellido2 = "Baeza"
        .CorreoElectronico = "carlosalmela@gmail.com"
        .Id = 2
        .Nombre = "Carlos"
        .Usuario = "Charly"
    End With
    PrintParticipante mPrs
    '
    '   3.- Objeto No valido
    '
    With mPrs
        .Apellido1 = ""
        .Apellido2 = ""
        .CorreoElectronico = ""
        .Id = 2
        .Nombre = ""
        .Usuario = ""
    End With
    PrintParticipante mPrs
    
    Err.Raise ERR_TODO, "Lot_PqtApuestasTesting.ParticipanteEngineTest", MSG_TODO
  
  On Error GoTo 0
ParticipanteEngineTest__CleanExit:
    Exit Sub
           
ParticipanteEngineTest_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    Call HandleException(ErrNumber, ErrDescription, "Lot_PqtApuestasTesting.ParticipanteEngineTest", ErrSource)
    Call MsgBox(ErrDescription, vbCritical Or vbSystemModal, ThisWorkbook.Name)
    Call Trace("CERRAR")
End Sub

'---------------------------------------------------------------------------------------
' Procedure : ParticipanteTest
' Author    : CHARLY
' Date      : mi., 14/ago/2019 11:15:58
' Purpose   : Pruebas unitarias de la clase Participante
'---------------------------------------------------------------------------------------
'
Private Sub ParticipanteTest()
    Dim oPart  As Participante
    
  On Error GoTo ParticipanteTest_Error
    '
    '   1.- Clase en vacio
    '
    Set oPart = New Participante
    PrintParticipante oPart
    '
    '   2.- Clase correcta
    '
    With oPart
        .Id = 3256
        .Apellido1 = "Almela"
        .Apellido2 = "Baeza"
        .CorreoElectronico = "carlosalmela@gmail.com"
        .Nombre = "Carlos"
        .Rol = rolAdministrador
        .Usuario = "CAB3780Y"
    End With
    PrintParticipante oPart
    '
    '   3.- Clase incompleta falta el Usuario y el correo
    '
    Set oPart = New Participante
    With oPart
        .Id = 256
        .Apellido1 = "Almela"
        .Apellido2 = "Baeza"
        .CorreoElectronico = Empty
        .Nombre = Empty
        .Rol = rolAdministrador
        .Usuario = "CAB3780Y"
    End With
    PrintParticipante oPart
    '
    '   4.- TODO: Usuario duplicado
    '
    '
    '
    '   5.- TODO: Email duplicado
    '
  On Error GoTo 0
ParticipanteTest__CleanExit:
    Exit Sub
           
ParticipanteTest_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    Call HandleException(ErrNumber, ErrDescription, "Lot_PqtApuestasTesting.ParticipanteTest", ErrSource)
    Call MsgBox(ErrDescription, vbCritical Or vbSystemModal, ThisWorkbook.Name)
    Call Trace("CERRAR")
End Sub

'---------------------------------------------------------------------------------------
' Procedure : BoletoTest
' Author    : CHARLY
' Date      : ma., 13/ago/2019 11:55:59
' Purpose   : Pruebas unitarias de la clase Boleto
'---------------------------------------------------------------------------------------
'
Private Sub BoletoTest()
    Dim mObj As Boleto
    Dim mApt As Apuesta
    
  On Error GoTo BoletoTest_Error
    '
    '   Cabecera de pruebas
    '
    Debug.Print "Pruebas del Objeto BOLETO ============="
    '
    '   1.- Boleto vacio de Bonoloto
    '
'    Set mObj = New Boleto
'    mObj.Juego = bonoloto
'    PrintBoleto mObj
    '
    '   2.- Boleto vacio de Primitiva
    '
'    Set mObj = New Boleto
'    mObj.Juego = LoteriaPrimitiva
'    PrintBoleto mObj
    '
    '   3.- Boleto vacio de Euromillon
    '
'    Set mObj = New Boleto
'    mObj.Juego = Euromillones
'    PrintBoleto mObj
    '
    '   4.- Boleto vacio de Gordo
    '
'    Set mObj = New Boleto
'    mObj.Juego = gordoPrimitiva
'    PrintBoleto mObj
    '
    '   5.- Boleto Valido de Bonoloto
    '
'    Set mApt = New Apuesta
'    With mApt
'        .Id = 324
'        .Juego = bonoloto
'        .IdBoleto = 5
'        .FechaSorteo = #8/19/2019#
'        .EsMultiple = False
'        .Coste = 0.5
'        .Numeros.Texto = "12-47-36-05-49-22"
'    End With
'    Set mObj = New Boleto
'
'    With mObj
'        .Apuestas.Add mApt
'        .Cadencia = 1     ' TODO: Definir eumeración para la cadencia: diario, semanal, bisemanal
'        .Comentarios = "Prueba de Comentarios"
'        .Coste = 1
'        .DesglosePremios = "R:1,5ª:(1)3"
'        .EsMultiple = False
'        .FechaSorteo = #8/19/2019#
'        .FechaValidez = #8/19/2019#
'        .Id = 5
'        .IdParticipante = 1
'        .ImportePremios = 5
'        .Joker = 0
'        .Juego = bonoloto
'        .Millon = ""
'        .ReintegroClave = 2
'        .Situacion = blBorrador
'    End With
'    Set mApt = New Apuesta
'    With mApt
'        .Id = 325
'        .Juego = bonoloto
'        .IdBoleto = 5
'        .FechaSorteo = #8/19/2019#
'        .EsMultiple = False
'        .Coste = 0.5
'        .Numeros.Texto = "05-24-17-36-09-47"
'    End With
'    mObj.Apuestas.Add mApt
'    PrintBoleto mObj
'
'    '
'    '   6.- Boleto Valido de Primitiva
'    '
'    Set mApt = New Apuesta
'    With mApt
'        .Id = 365
'        .Juego = LoteriaPrimitiva
'        .IdBoleto = 4
'        .FechaSorteo = #5/25/2019#
'        .EsMultiple = False
'        .Coste = 1
'        .Numeros.Texto = "38-04-17-19-36-48"
'    End With
'
'    Set mObj = New Boleto
'    With mObj
'        .Apuestas.Add mApt
'        .Cadencia = 1     ' TODO: Definir eumeración para la cadencia: diario, semanal, bisemanal
'        .Comentarios = "Prueba de Comentarios"
'        .Coste = 1
'        .DesglosePremios = "Sin Premios"
'        .EsMultiple = False
'        .FechaSorteo = #5/25/2019#
'        .FechaValidez = #5/25/2019#
'        .Id = 4
'        .IdParticipante = 1
'        .ImportePremios = 0
'        .Joker = 2665969
'        .Juego = LoteriaPrimitiva
'        .ReintegroClave = 6
'        .Situacion = blBorrador
'    End With
'    PrintBoleto mObj
'    '
'    '   7.- Boleto Valido de Euromillon
'    '
'    Set mApt = New Apuesta
'    With mApt
'        .Id = 23
'        .Juego = Euromillones
'        .IdBoleto = 7
'        .FechaSorteo = #5/24/2019#
'        .EsMultiple = False
'        .Coste = 2.5
'        .Numeros.Texto = "27-46-42-25-39"
'        .Estrellas.Texto = "11-12"
'    End With
'
'    Set mObj = New Boleto
'    With mObj
'        .Apuestas.Add mApt
'        .Cadencia = 1     ' TODO: Definir eumeración para la cadencia: diario, semanal, bisemanal
'        .Comentarios = "Prueba de Comentarios"
'        .Coste = 2.5
'        .DesglosePremios = "Sin Premios"
'        .EsMultiple = False
'        .FechaSorteo = #5/24/2019#
'        .FechaValidez = #5/24/2019#
'        .Id = 7
'        .IdParticipante = 1
'        .ImportePremios = 0
'        .Juego = Euromillones
'        .Millon = "TLG96606"
'        .Situacion = blBorrador
'    End With
'    PrintBoleto mObj
'    '
'    '   8.- Boleto Valido de Gordo
'    '
'    Set mApt = New Apuesta
'    With mApt
'        .Id = 154
'        .Juego = gordoPrimitiva
'        .IdBoleto = 3
'        .FechaSorteo = #4/28/2019#
'        .EsMultiple = False
'        .Coste = 1.5
'        .Numeros.Texto = "02-24-04-05-43"
'    End With
'
'    Set mObj = New Boleto
'    With mObj
'        .Apuestas.Add mApt
'        .Cadencia = 1     ' TODO: Definir eumeración para la cadencia: diario, semanal, bisemanal
'        .Comentarios = "Prueba de Comentarios"
'        .Coste = 1.5
'        .DesglosePremios = "Sin Premios"
'        .EsMultiple = False
'        .FechaSorteo = #4/28/2019#
'        .FechaValidez = #4/28/2019#
'        .Id = 3
'        .IdParticipante = 1
'        .ImportePremios = 0
'        .Juego = gordoPrimitiva
'        .ReintegroClave = 2
'        .Situacion = blBorrador
'    End With
'    PrintBoleto mObj
    '
    '   9.- Boleto NO Valido de Bonoloto
    '
    
    '
    '  10.- Boleto NO Valido de Primitiva
    '
    '
    '  11.- Boleto NO Valido de Euromillon
    '
    '
    '  12.- Boleto NO Valido de Gordo
    '
    '
    '  13.- Boleto Multiple Valido Bonoloto
    '
    '
    '  14.- Boleto Multiple Valido Gordo
    '
    '
    '  15.- Boleto Multiple Valido Euromillon
    '
    '
    '  16.- Boleto Multiple Valido Primitiva
    '
    '
    Err.Raise ERR_TODO, "Lot_PqtApuestasTesting.BoletoTest", MSG_TODO
  
  On Error GoTo 0
BoletoTest__CleanExit:
    Exit Sub
           
BoletoTest_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    Call HandleException(ErrNumber, ErrDescription, "Lot_PqtApuestasTesting.BoletoTest", ErrSource)
    Call MsgBox(ErrDescription, vbCritical Or vbSystemModal, ThisWorkbook.Name)
    Call Trace("CERRAR")
End Sub

'---------------------------------------------------------------------------------------
' Procedure : BoletoTest
' Author    : CHARLY
' Date      : ma., 13/ago/2019 11:59:43
' Purpose   : Pruebas unitarias de la colección Boletos
'---------------------------------------------------------------------------------------
'
Private Sub BoletosTest()
  On Error GoTo BoletosTest_Error
    Err.Raise ERR_TODO, "Lot_PqtApuestasTesting.BoletosTest", MSG_TODO
  
  On Error GoTo 0
BoletosTest__CleanExit:
    Exit Sub
           
BoletosTest_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    Call HandleException(ErrNumber, ErrDescription, "Lot_PqtApuestasTesting.BoletosTest", ErrSource)
    Call MsgBox(ErrDescription, vbCritical Or vbSystemModal, ThisWorkbook.Name)
    Call Trace("CERRAR")
End Sub

'---------------------------------------------------------------------------------------
' Procedure : BoletoEngineTest
' Author    : CHARLY
' Date      : ma., 13/ago/2019 12:03:07
' Purpose   : Pruebas unitarias de la clase BoletoEngine
'---------------------------------------------------------------------------------------
'
Private Sub BoletoEngineTest()
  On Error GoTo BoletoEngineTest_Error
    Err.Raise ERR_TODO, "Lot_PqtApuestasTesting.BoletoEngineTest", MSG_TODO
  
  On Error GoTo 0
BoletoEngineTest__CleanExit:
    Exit Sub
           
BoletoEngineTest_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    Call HandleException(ErrNumber, ErrDescription, "Lot_PqtApuestasTesting.BoletoEngineTest", ErrSource)
    Call MsgBox(ErrDescription, vbCritical Or vbSystemModal, ThisWorkbook.Name)
    Call Trace("CERRAR")
End Sub

'---------------------------------------------------------------------------------------
' Procedure : ApuestaTest
' Author    : CHARLY
' Date      : ma., 13/ago/2019 12:06:09
' Purpose   : Pruebas unitarias de la clase Apuesta
'---------------------------------------------------------------------------------------
'
Private Sub ApuestaTest()
    Dim mObj As Apuesta
    Dim mNum As Numero
  On Error GoTo ApuestaTest_Error
    '
    '   1.- Apuesta vacia                   Bonoloto
    '
    Set mObj = New Apuesta
    mObj.Juego = bonoloto
    PrintApuesta mObj
    
    '
    '   2.- Apuesta valida                  Bonoloto
    '
    With mObj
        .Id = 324
        .IdBoleto = 5
        .FechaSorteo = #8/19/2019#
        .EsMultiple = False
        .Coste = 0.5
        .Numeros.Texto = "12-47-36-05-49-22"
    End With
    PrintApuesta mObj
    '
    '   3.- Apuesta No valida               Bonoloto
    '
    With mObj
        .Id = 324
        .IdBoleto = 5
        .FechaSorteo = #8/19/2019#
        .EsMultiple = True
        .Coste = 0.5
        .Numeros.Texto = "12-47-56-05-49-22"
    End With
    PrintApuesta mObj
    
    
    '
    '   4.- Apuesta vacia                   Primitiva
    '
    Set mObj = New Apuesta
    mObj.Juego = LoteriaPrimitiva
    PrintApuesta mObj
    
    '
    '   5.- Apuesta valida                  Primitiva
    '
    With mObj
        .Id = 1256
        .IdBoleto = 152
        .FechaSorteo = #8/19/2019#
        .EsMultiple = True
        .Coste = 1
        .ImportePremios = 8
        .Numeros.Texto = "12-47-36-5-49-22-15-35-48-7"
    End With
    PrintApuesta mObj
    '
    '   6.- Apuesta No valida               Primitiva
    '
    With mObj
        .Id = 1256
        .IdBoleto = 152
        .FechaSorteo = #8/19/2019#
        .EsMultiple = False
        .Coste = 1
        .ImportePremios = 8
        .Numeros.Texto = "12-47-56-5-49-22-15-35-48-7"
    End With
    PrintApuesta mObj
    '
    '   7.- Apuesta vacia                   Gordo
    '
    Set mObj = New Apuesta
    mObj.Juego = gordoPrimitiva
    PrintApuesta mObj
    '
    '   8.- Apuesta valida                  Gordo
    '
    With mObj
        .Id = 348
        .IdBoleto = 22
        .FechaSorteo = #8/18/2019#
        .EsMultiple = False
        .Coste = 1.5
        .ImportePremios = 0
        .Numeros.Texto = "12-17-43-5-49"
    End With
    PrintApuesta mObj
    
    '
    '   9.- Apuesta No valida               Gordo
    '
    With mObj
        .Id = 348
        .IdBoleto = 22
        .FechaSorteo = #8/18/2019#
        .EsMultiple = False
        .Coste = 1.5
        .ImportePremios = 0
        .Numeros.Texto = "12-17-43-5-85"
    End With
    PrintApuesta mObj
    
    '
    '   10.- Apuesta vacia                  Euromillones
    '
    Set mObj = New Apuesta
    mObj.Juego = Euromillones
    PrintApuesta mObj
    '
    '   11.- Apuesta valida                 Euromillones
    '
    With mObj
        .Id = 2560
        .IdBoleto = 896
        .FechaSorteo = #8/16/2019#
        .EsMultiple = False
        .Coste = 2.5
        .ImportePremios = 0
        .Numeros.Texto = "12-17-43-5-23"
        .Estrellas.Texto = "2-7"
    End With
    PrintApuesta mObj
    '
    '   12.- Apuesta No valida              Euromillones
    '
    With mObj
        .Id = 2560
        .IdBoleto = 896
        .FechaSorteo = #8/16/2019#
        .EsMultiple = False
        .Coste = 2.5
        .ImportePremios = 0
        .Numeros.Texto = "12-54-43-5-23"
        .Estrellas.Texto = "15-7"
    End With
    PrintApuesta mObj
  
  On Error GoTo 0
ApuestaTest__CleanExit:
    Exit Sub
           
ApuestaTest_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    Call HandleException(ErrNumber, ErrDescription, "Lot_PqtApuestasTesting.ApuestaTest", ErrSource)
    Call MsgBox(ErrDescription, vbCritical Or vbSystemModal, ThisWorkbook.Name)
    Call Trace("CERRAR")
End Sub

'---------------------------------------------------------------------------------------
' Procedure : ApuestasTest
' Author    : CHARLY
' Date      : ma., 13/ago/2019 12:03:07
' Purpose   : Pruebas unitarias de la colección Apuestas
'---------------------------------------------------------------------------------------
'
Private Sub ApuestasTest()
    Dim mObj As Apuestas
    Dim mApt As Apuesta
    
  On Error GoTo ApuestaTest_Error
    '
    '   1.- Objeto en Vacio
    '
    Set mObj = New Apuestas
    PrintApuestas mObj
    '
    '   2.- Apuestas de  Bonoloto
    '
    Set mApt = New Apuesta
    With mApt
        .Juego = bonoloto
        .Id = 4512
        .EsMultiple = False
        .FechaSorteo = #5/15/2018#
        .Numeros.Texto = "10-49-15-31-17-7"
        .Coste = 0.5
    End With
    mObj.Add mApt
    '
    '   Segunda Apuesta
    '
    Set mApt = New Apuesta
    With mApt
        .Juego = bonoloto
        .Id = 6631
        .EsMultiple = False
        .FechaSorteo = #5/15/2018#
        .Numeros.Texto = "38-31-45-5-48-13"
        .Coste = 0.5
    End With
    mObj.Add mApt
    PrintApuestas mObj
    '
    '   3.- Apuestas Euromillon
    '
    Set mObj = New Apuestas
    ' primero item
    Set mApt = New Apuesta
    With mApt
        .Juego = Euromillones
        .Numeros.Texto = "5-31-18-21-35"
        .Estrellas.Texto = "6-9"
        .FechaSorteo = #10/26/2018#
        .Id = 1106
    End With
    mObj.Add mApt
    ' segundo item
    Set mApt = New Apuesta
    With mApt
        .Juego = Euromillones
        .Numeros.Texto = "44-27-23-17-43"
        .Estrellas.Texto = "1-12"
        .FechaSorteo = #10/30/2018#
        .Id = 1107
    End With
    mObj.Add mApt
    ' tercer item
    Set mApt = New Apuesta
    With mApt
        .Juego = Euromillones
        .Numeros.Texto = "15-37-5-17-44"
        .Estrellas.Texto = "11-7"
        .FechaSorteo = #11/2/2018#
        .Id = 1108
    End With
    mObj.Add mApt
    PrintApuestas mObj
    '
    '   4.- Prueba propiedad Count
    '
    Debug.Print "=> Propiedad Count (3) => " & mObj.Count
    '
    '   5.- Prueba método Items
    '
    Debug.Print "=> Prueba Items"
    For Each mApt In mObj.Items
        Debug.Print vbTab & "* (" & mApt.Id & ") Apuesta=>" & mApt.ToString
    Next mApt
    '
    '   6.- Prueba método MarkForDelete
    '
    mObj.MarkForDelete 1
    Debug.Print "=> Prueba MarkForDelete"
    Debug.Print vbTab & "* (" & mObj.Items(1).Id & ") Valor=> " & mObj.Items(1).EntidadNegocio.MarkForDelete
    '
    '   7.- Prueba método Undelete
    '
    mObj.Undelete 1
    Debug.Print "=> Prueba Undelete"
    Debug.Print vbTab & "* (" & mObj.Items(1).Id & ") Valor=> " & mObj.Items(1).EntidadNegocio.MarkForDelete
    '
    '   8.- Activamos el control del error porque queremos desmarcar un elemento inexistente
    '
    On Error Resume Next
    mObj.Undelete 5
    If Err.Number > 0 Then
        Debug.Print "#Err:" & Err.Number & "-" & Err.Description
    End If
    On Error GoTo 0
    '
    '   9.- Prueba método Delete
    '
    mObj.Delete 1
    '
    Debug.Print "=> Metodo Delete (2) => " & mObj.Count
    '
    '   10.- Activamos el control del error porque queremos borrar un elemento inexistente
    '
    On Error Resume Next
    mObj.Delete 5
    If Err.Number > 0 Then
        Debug.Print "#Err:" & Err.Number & "-" & Err.Description
    End If
    On Error GoTo 0
    '
    '   11.- Prueba metodo Clear
    '
    Debug.Print "=> Prueba Clear"
    mObj.Clear
    PrintApuestas mObj
 
  
  On Error GoTo 0
ApuestaTest__CleanExit:
    Exit Sub
           
ApuestaTest_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    Call HandleException(ErrNumber, ErrDescription, "Lot_PqtApuestasTesting.ApuestaTest", ErrSource)
    Call MsgBox(ErrDescription, vbCritical Or vbSystemModal, ThisWorkbook.Name)
    Call Trace("CERRAR")
End Sub

'---------------------------------------------------------------------------------------
' Procedure : ApuestaEngineTest
' Author    : CHARLY
' Date      : ma., 13/ago/2019 12:05:36
' Purpose   : Pruebas unitarias de la clase ApuestaEngine
'---------------------------------------------------------------------------------------
'
Private Sub ApuestaEngineTest()
    Dim mEng As ApuestasEngine
    Dim mObj As Apuesta
    Dim mCol As Apuestas
    
  On Error GoTo ApuestaEngineTest_Error
    '
    '   Cabecera de pruebas
    '
    Debug.Print "Pruebas del Motor de Apuetas ============="
    '
    '   Creamos el motor que proporciona Apuetas
    '
    Set mEng = New ApuestasEngine
    '
    '   1.- Prueba metodo GetNewApuesta()
    '
    Debug.Print vbTab & "1.- Nuevo objeto"
'    Set mObj = mEng.GetNewApuesta(bonoloto)
'    PrintApuesta mObj
'    '
'    '   2.- Prueba metodo SetApuesta Bonoloto
'    '
'    Debug.Print vbTab & "2.- Guardar Objeto (primitiva)"
'    With mObj
'        .Coste = 0.5
'        .EsMultiple = False
'        .FechaSorteo = #8/20/2019#
'        .IdBoleto = 1
'        .Numeros.Texto = "05-08-12-21-38-49"
'    End With
'    PrintApuesta mObj
'    mEng.SetApuesta mObj
'    '
'    '   3.- Prueba metodo SetApuesta Primitiva
'    '
'    Debug.Print vbTab & "3.- Guardar Objeto (primitiva)"
'    Set mObj = mEng.GetNewApuesta(LoteriaPrimitiva)
'    With mObj
'        .Coste = 1
'        .ImportePremios = 8
'        .EsMultiple = False
'        .FechaSorteo = #8/20/2019#
'        .IdBoleto = 2
'        .Numeros.Texto = "1-15-23-24-27-43"
'    End With
'    PrintApuesta mObj
'    mEng.SetApuesta mObj
'    '
'    '   4.- Prueba metodo SetApuesta Euromillon
'    '
'    Debug.Print vbTab & "4.- Guardar Objeto (Euromillon)"
'    Set mObj = mEng.GetNewApuesta(Euromillones)
'    With mObj
'        .Coste = 2.5
'        .ImportePremios = 4.58
'        .EsMultiple = False
'        .FechaSorteo = #8/19/2019#
'        .IdBoleto = 5
'        .Numeros.Texto = "1-49-15-8-44"
'        .Estrellas.Texto = "2-11"
'    End With
'    PrintApuesta mObj
'    mEng.SetApuesta mObj
'    '
'    '   5.- Prueba metodo SetApuesta Gordo
'    '
'    Debug.Print vbTab & "4.- Guardar Objeto (Gordo)"
'    Set mObj = mEng.GetNewApuesta(gordoPrimitiva)
'    With mObj
'        .Coste = 1.5
'        .ImportePremios = 0
'        .EsMultiple = False
'        .FechaSorteo = #8/11/2019#
'        .IdBoleto = 8
'        .Numeros.Texto = "2-24-4-5-43"
'    End With
'    PrintApuesta mObj
'    mEng.SetApuesta mObj
'    '
'    '   6.- Prueba metodo SetApuesta Euromillon multiple
'    '
'    Debug.Print vbTab & "6.- Guardar Objeto (Euromillon)"
'    Set mObj = mEng.GetNewApuesta(Euromillones)
'    With mObj
'        .Coste = 315
'        .ImportePremios = 0
'        .EsMultiple = True
'        .FechaSorteo = #8/23/2019#
'        .IdBoleto = 5
'        .Numeros.Texto = "5-13-24-27-34-36-44"
'        .Estrellas.Texto = "4-5-8-10"
'    End With
'    PrintApuesta mObj
'    mEng.SetApuesta mObj
'    '
'    '   7.- Prueba metodo GetApuestaById
'    '
'    Debug.Print vbTab & "7.- Obtener la apuesta Id: 2"
'    Set mObj = mEng.GetApuestaById(2)
'    PrintApuesta mObj
'    '
'    '   8.- Prueba metodo SetApuestas
'    '
'    Debug.Print vbTab & "8.- SetApuestas"
'    Set mObj = mEng.GetNewApuesta(bonoloto)
'    With mObj
'        .Coste = 0.5
'        .ImportePremios = 0
'        .EsMultiple = False
'        .FechaSorteo = #10/12/2019#
'        .IdBoleto = 234
'        .Numeros.Texto = "5-16-24-34-43"
'    End With
'    Set mCol = New Apuestas
'    mCol.Add mObj
'    Set mObj = mEng.GetNewApuesta(bonoloto)
'    With mObj
'        .Id = .Id + 1
'        .Coste = 0.5
'        .ImportePremios = 0
'        .EsMultiple = False
'        .FechaSorteo = #10/12/2019#
'        .IdBoleto = 234
'        .Numeros.Texto = "1-2-19-22-49"
'    End With
'    mCol.Add mObj
'    PrintApuestas mCol
'    mEng.SetApuestas mCol
'    '
'    '   9.- Prueba del método GetApuestasBoleto
'    '
'    Debug.Print vbTab & "9.- GetApuestasBoleto (234)"
'    Set mCol = mEng.GetApuestasBoleto(234)
'    PrintApuestas mCol
'    '
'    '   Boleto no encontrado
'    On Error Resume Next
'    Debug.Print vbTab & "    GetApuestasBoleto (88)"
'    Set mCol = mEng.GetApuestasBoleto(88)
'    If Err.Number <> 0 Then
'        Debug.Print "Apuestas no encontradas: (#" & Err.Number & ") " & Err.Description
'    End If
'    On Error GoTo ApuestaEngineTest_Error
    '
    '  10.- Prueba del método GetApuestasInFechas
    '
    Debug.Print vbTab & "10.- GetApuestasInFechas "
    '
    '  11.- Prueba del método GetApuestasInIds
    '
    '
    '
    '
    Err.Raise ERR_TODO, "Lot_PqtApuestasTesting.ApuestaEngineTest", MSG_TODO
  
  On Error GoTo 0
ApuestaEngineTest__CleanExit:
    Exit Sub
           
ApuestaEngineTest_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    Call HandleException(ErrNumber, ErrDescription, "Lot_PqtApuestasTesting.ApuestaEngineTest", ErrSource)
    Call MsgBox(ErrDescription, vbCritical Or vbSystemModal, ThisWorkbook.Name)
    Call Trace("CERRAR")
End Sub


Private Sub PrintApuesta(Obj As Apuesta)
    Debug.Print "==> Apuesta "
    Debug.Print vbTab & "Coste           = " & Obj.Coste
    Debug.Print vbTab & "EsMultiple      = " & Obj.EsMultiple
    Debug.Print vbTab & "#obj Estrellas  = " & Obj.Estrellas.ToString()
    Debug.Print vbTab & "FechaSorteo     = " & Obj.FechaSorteo
    Debug.Print vbTab & "Id              = " & Obj.Id
    Debug.Print vbTab & "IdBoleto        = " & Obj.IdBoleto
    Debug.Print vbTab & "ImportePremios  = " & Obj.ImportePremios
    Debug.Print vbTab & "Juego           = " & Obj.Juego
    Debug.Print vbTab & "TotalApuestas   = " & Obj.TotalApuestas
    Debug.Print vbTab & "#obj Numeros    = " & Obj.Numeros.ToString()
    Debug.Print vbTab & "EsValido()      = " & Obj.EsValido()
    Debug.Print vbTab & "GetMensaje()    = " & Obj.GetMensaje()
    Debug.Print vbTab & "ToString()      = " & Obj.ToString()
End Sub

Private Sub PrintBoleto(Obj As Boleto)
    Debug.Print "==> Boleto "
    Debug.Print vbTab & "#obj Apuestas   = " & Obj.Apuestas.Count
    Debug.Print vbTab & "Cadencia        = " & Obj.Cadencia
    Debug.Print vbTab & "Comentarios     = " & Obj.Comentarios
    Debug.Print vbTab & "Coste           = " & Obj.Coste
    Debug.Print vbTab & "DesglosePremios = " & Obj.DesglosePremios
    Debug.Print vbTab & "EsMultiple      = " & Obj.EsMultiple
    Debug.Print vbTab & "FechaSorteo     = " & Obj.FechaSorteo
    Debug.Print vbTab & "FechaValidez    = " & Obj.FechaValidez
    Debug.Print vbTab & "Id              = " & Obj.Id
    Debug.Print vbTab & "IdParticipante  = " & Obj.IdParticipante
    Debug.Print vbTab & "ImportePremios  = " & Obj.ImportePremios
    Debug.Print vbTab & "Joker           = " & Obj.Joker
    Debug.Print vbTab & "Juego           = " & Obj.Juego
    Debug.Print vbTab & "Millon          = " & Obj.Millon
    Debug.Print vbTab & "NumeroApuestas  = " & Obj.NumeroApuestas
    Debug.Print vbTab & "ReintegroClave  = " & Obj.ReintegroClave
    Debug.Print vbTab & "Situacion       = " & Obj.Situacion
    Debug.Print vbTab & "EsValido()      = " & Obj.EsValido()
    Debug.Print vbTab & "GetMensaje()    = " & Obj.GetMensaje()
    Debug.Print vbTab & "ToString()      = " & Obj.ToString()
End Sub

Private Sub PrintApuestas(mObj As Apuestas)
    Dim mApt As Apuesta
    Debug.Print "==> Pruebas Apuestas"
    Debug.Print vbTab & "Add                 =" & "#Metodo" 'mObj.Add
    Debug.Print vbTab & "Clear               =" & "#Metodo" 'mObj.Clear
    Debug.Print vbTab & "Count               =" & mObj.Count
    Debug.Print vbTab & "Delete              =" & "#Metodo" 'mObj.Delete
    Debug.Print vbTab & "Items.Count         =" & mObj.Items.Count
    Debug.Print vbTab & "Juego               =" & mObj.Juego
    Debug.Print vbTab & "MarkForDelete       =" & "#Metodo" 'mObj.MarkForDelete
    Debug.Print vbTab & "Undelete            =" & "#Metodo" 'mObj.Undelete
    For Each mApt In mObj.Items
        Debug.Print vbTab & mApt.ToString
    Next mApt
End Sub

Private Sub PrintParticipante(mObj As Participante)
    Debug.Print "==> Pruebas Participante"
    Debug.Print vbTab & "Apellido1           =" & mObj.Apellido1
    Debug.Print vbTab & "Apellido2           =" & mObj.Apellido2
    Debug.Print vbTab & "CorreoElectronico   =" & mObj.CorreoElectronico
    Debug.Print vbTab & "Id                  =" & mObj.Id
    Debug.Print vbTab & "Nombre              =" & mObj.Nombre
    Debug.Print vbTab & "Rol                 =" & mObj.Rol
    Debug.Print vbTab & "Usuario             =" & mObj.Usuario
    Debug.Print vbTab & "ToString            =" & mObj.ToString
    Debug.Print vbTab & "IsValid             =" & mObj.IsValid
    Debug.Print vbTab & "GetMensaje          =" & mObj.GetMensaje
End Sub
