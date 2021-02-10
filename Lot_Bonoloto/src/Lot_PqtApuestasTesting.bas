Attribute VB_Name = "Lot_PqtApuestasTesting"
' *============================================================================*
' *
' *     Fichero    : Lot_PqtApuestasTesting.bas
' *
' *     Autor      : Carlos Almela Baeza
' *     Creación   : jue, 29/mar/2018 20:00:15
' *     Versión    : 1.0
' *     Propósito  : Colección de pruebas del paquete apuestas
' *                  - Boleto
' *                  - Apuesta
' *                  - ComprobarBoletos
' *
' *============================================================================*
Option Explicit
Option Base 0
'--- Variables Privadas -------------------------------------------------------*
Dim mDB         As BdDatos
Dim mFila       As Range
Dim mFini       As Date
Dim mApt        As Apuesta
    
'
'------------------------------------------------------------------------------*
' Procedimiento  : ApuestaTest
' Fecha          : sáb, 01/nov/2014 21:01:37
' Propósito      : Casos de prueba de la clase Apuesta
'------------------------------------------------------------------------------*
'
Private Sub ApuestaTest()
    Dim mObj        As Apuesta
    
  On Error GoTo ApuestaTest_Error
    '
    '   Caso de Prueba 01: Objeto Vacio
    '
    Set mObj = New Apuesta
    PrintApuesta mObj
    '
    '   Caso de Prueba 02: Apuesta de una fecha existente
    '
    Set mDB = New BdDatos
    mFini = #6/30/2020#  ' Martes
    
    Set mFila = mDB.GetApuestaByFecha(mFini)
    If Not (mFila Is Nothing) Then
        mObj.Constructor mFila
        If mObj.Fecha = mFini Then
            PrintApuesta mObj
        Else
            Debug.Print ("#Error en GetApuestaByFecha: " & mObj.Fecha)
        End If
    Else
        Debug.Print ("#Error en rango: oFila is Nothing")
    End If
    '
    '   Caso de Prueba 03: Apuesta de una fecha Inexistente
    '
    mFini = #6/28/2020#   'Domingo
    
    Set mFila = mDB.GetApuestaByFecha(mFini)
    If Not (mFila Is Nothing) Then
        mObj.Constructor mFila
        If mObj.Fecha = mFini Then
            PrintApuesta mObj
        Else
            Debug.Print ("#Error en GetApuestaByFecha: " & mObj.Fecha)
        End If
    Else
        Debug.Print ("#Error en rango: mFila is Nothing")
    End If



  On Error GoTo 0
    Exit Sub
ApuestaTest_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    Call HandleException(ErrNumber, ErrDescription, "Lot_PqtApuestasTesting.ApuestaTest", ErrSource)
    Call MsgBox(ErrDescription, vbCritical Or vbSystemModal, ThisWorkbook.Name)
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
    Debug.Print "  Aciertos                         = " & datApuesta.Aciertos
    Debug.Print "  CategoriaPremio                  = " & datApuesta.CategoriaPremio
    Debug.Print "  Combinacion                      = " & datApuesta.Combinacion.Texto
    Debug.Print "  Coste                            = " & datApuesta.Coste
    Debug.Print "  EntidadNegocio.ID                = " & datApuesta.EntidadNegocio.Id
    Debug.Print "  EntidadNegocio.ClassStorage      = " & datApuesta.EntidadNegocio.ClassStorage
    Debug.Print "  EntidadNegocio.FechaAlta         = " & datApuesta.EntidadNegocio.FechaAlta
    Debug.Print "  EntidadNegocio.FechaBaja         = " & datApuesta.EntidadNegocio.FechaBaja
    Debug.Print "  EntidadNegocio.FechaModificacion = " & datApuesta.EntidadNegocio.FechaModificacion
    Debug.Print "  EntidadNegocio.IsDirty           = " & datApuesta.EntidadNegocio.IsDirty
    Debug.Print "  EntidadNegocio.IsNew             = " & datApuesta.EntidadNegocio.IsNew
    Debug.Print "  EntidadNegocio.MarkForDelete     = " & datApuesta.EntidadNegocio.MarkForDelete
    Debug.Print "  EntidadNegocio.Origen            = " & datApuesta.EntidadNegocio.Origen
    Debug.Print "  EntidadNegocio.Situacion         = " & datApuesta.EntidadNegocio.Situacion
    Debug.Print "  EsMultiple                       = " & datApuesta.EsMultiple
    Debug.Print "  Fecha                            = " & datApuesta.Fecha
    Debug.Print "  FechaFinVigencia                 = " & datApuesta.FechaFinVigencia
    Debug.Print "  IdBoleto                         = " & datApuesta.IdBoleto
    Debug.Print "  Juego                            = " & datApuesta.Juego
    Debug.Print "  Metodo                           = " & datApuesta.Metodo
    Debug.Print "  NumeroApuestas                   = " & datApuesta.NumeroApuestas
    Debug.Print "  Pronosticos                      = " & datApuesta.Pronosticos
    Debug.Print "  SeHaJugado                       = " & datApuesta.SeHaJugado
    Debug.Print "  Semana                           = " & datApuesta.Semana
    Debug.Print "  Texto                            = " & datApuesta.Texto
End Sub


'------------------------------------------------------------------------------*
' Procedimiento  : BoletoTest
' Fecha          : ju., 09/jul/2020 18:21:44
' Propósito      : Casos de prueba de la clase Boleto
'------------------------------------------------------------------------------*
'
Private Sub BoletoTest()
    Dim mObj  As Boleto
    
  On Error GoTo BoletoTest_Error
    '
    '   Caso de Prueba 01: Objeto Vacio
    '
    Set mObj = New Boleto
    PrintBoleto mObj
    '
    '   Caso de Prueba 02: Boleto de una fecha existente
    '
    mFini = #7/3/2020#  'Viernes
    Set mDB = New BdDatos
    Set mFila = mDB.GetBoletoByFecha(mFini)
    If Not (mFila Is Nothing) Then
        mObj.Constructor mFila
        mObj.SetApuestas
        If mObj.FechaValidez = mFini Then
            PrintBoleto mObj
        Else
            Debug.Print ("#Error en BoletoByFecha: " & mObj.FechaValidez)
        End If
    Else
        Debug.Print ("#Error en rango: oFila is Nothing")
    End If
    '
    '   Caso de Prueba 02: Boleto de una fecha existente
    '
    mFini = #7/5/2020#  ' Domingo
    Set mFila = mDB.GetBoletoByFecha(mFini)
    If Not (mFila Is Nothing) Then
        mObj.Constructor mFila
        mObj.SetApuestas
        If mObj.FechaValidez = mFini Then
            PrintBoleto mObj
        Else
            Debug.Print ("#Error en BoletoByFecha: " & mObj.FechaValidez)
        End If
    Else
        Debug.Print ("Correcto fecha no existe: oFila is Nothing")
    End If
    '
    '   Caso de prueba 03: Boleto existente sin apuestas
    '
    mFini = #7/7/2020#  ' Domingo
    Set mFila = mDB.GetBoletoByFecha(mFini)
    If Not (mFila Is Nothing) Then
        mObj.Constructor mFila
        mObj.Id = 10001   ' Boleto inexistente
        mObj.SetApuestas
        If mObj.FechaValidez = mFini Then
            PrintBoleto mObj
        Else
            Debug.Print ("#Error en BoletoByFecha: " & mObj.FechaValidez)
        End If
    Else
        Debug.Print ("Correcto fecha no existe: oFila is Nothing")
    End If
    '
    '   Caso de prueba 04: Boleto existente con multiples apuestas
    '
    mFini = #10/31/2020#  ' Sábado
    Set mFila = mDB.GetBoletoByFecha(mFini)
    If Not (mFila Is Nothing) Then
        mObj.Constructor mFila
        mObj.Id = 251
        mObj.SetApuestas
        If mObj.FechaValidez = mFini Then
            PrintBoleto mObj
        Else
            Debug.Print ("#Error en BoletoByFecha: " & mObj.FechaValidez)
        End If
    Else
        Debug.Print ("Correcto fecha no existe: oFila is Nothing")
    End If
  
  
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



Private Sub PrintBoleto(mObj As Boleto)
    Debug.Print "==> Pintar Boleto"
    Debug.Print "  Apuestas         = " & mObj.Apuestas.Count
    Debug.Print "  CategoriaPremio  = " & mObj.CategoriaPremio
    Debug.Print "  Constructor      = " ' mObj.Constructor
    Debug.Print "  Coste            = " & mObj.Coste
    Debug.Print "  Dia              = " & mObj.Dia
    Debug.Print "  DiasValidez      = " & mObj.DiasValidez
    Debug.Print "  ElMillon         = " & mObj.ElMillon
    Debug.Print "  Fecha_Caducidad  = " & mObj.FechaCaducidad
    Debug.Print "  FechaValidez     = " & mObj.FechaValidez
    Debug.Print "  Id               = " & mObj.Id
    Debug.Print "  Premios          = " & mObj.ImportePremios
    Debug.Print "  Joker            = " & mObj.Joker
    Debug.Print "  Juego            = " & mObj.Juego
    Debug.Print "  JuegoTexto       = " & mObj.JuegoTexto
    Debug.Print "  Multiple         = " & mObj.Multiple
    Debug.Print "  Notas            = " & mObj.Notas
    Debug.Print "  Numero_Apuestas  = " & mObj.NumeroApuestas
    Debug.Print "  Reintegro        = " & mObj.Reintegro
    Debug.Print "  Semana           = " & mObj.Semana
    Debug.Print "  Semanal          = " & mObj.Semanal
End Sub




'------------------------------------------------------------------------------*
' Procedimiento  : ComprobarApuestaTest
' Fecha          : vi., 28/ago/2020 18:11:35
' Propósito      : Casos de prueba de la clase ComprobarBoleto: Apuestas
'------------------------------------------------------------------------------*
'
Private Sub ComprobarApuestaTest()
    Dim mObj As ComprobarBoletos
    
  On Error GoTo ComprobarApuestaTest_Error
    '
    '   Caso de pruebas 01: Objeto vacio
    '
    Set mApt = New Apuesta
    Set mObj = New ComprobarBoletos
    Debug.Print "==> Comprobar Apuesta"
    Debug.Print "    Apuesta Vacia: " & mObj.ComprobarApuesta(mApt, False)
    PrintComprobarBoleto mObj
    '
    '   Comprobación de BONOLOTO
    '
    If JUEGO_DEFECTO = Bonoloto Then
        '
        '   Caso de pruebas 02: Comprobar apuesta simple bonoloto 5ª categoria
        '
        With mApt
            .Combinacion.Texto = "22-19-26-07-30-45"
            .Fecha = #7/7/2020#  ' 7-19-14-26-46-41-32 C-1
            .Juego = Bonoloto
        End With
        Debug.Print vbTab & "Apuesta 5ª: " & mObj.ComprobarApuesta(mApt, False)
        If mObj.CategoriaPremioTxt = "5ª" Then
            Debug.Print vbTab & "OK => Premio 3cg"
        Else
            Debug.Print vbTab & "#Error Premio distinto de 3cg"
        End If
        PrintComprobarBoleto mObj
        '
        '   Caso de pruebas 03: Comprobar apuesta simple bonoloto 1ª categoria
        '
        With mApt
            .Combinacion.Texto = "14-19-26-07-41-46"
            .Fecha = #7/7/2020#  ' 7-19-14-26-46-41 C-32 R-1
            .Juego = Bonoloto
        End With
        Debug.Print vbTab & "Apuesta 1ª: " & mObj.ComprobarApuesta(mApt, False)
        If mObj.CategoriaPremioTxt = "1ª" Then
            Debug.Print vbTab & "OK => Premio 6cg"
        Else
            Debug.Print vbTab & "#Error Premio distinto de 6cg"
        End If
        PrintComprobarBoleto mObj
        '
        '   Caso de pruebas 04: Comprobar apuesta simple bonoloto 2ª categoria
        '
        With mApt
            .Combinacion.Texto = "7-19-14-26-46-32"
            .Fecha = #7/7/2020#  ' 7-19-14-26-46-41 C-32 R-1
            .Juego = Bonoloto
        End With
        Debug.Print vbTab & "Apuesta 2ª: " & mObj.ComprobarApuesta(mApt, False)
        If mObj.CategoriaPremioTxt = "2ª" Then
            Debug.Print vbTab & "OK => Premio 5cg+C"
        Else
            Debug.Print vbTab & "#Error Premio distinto de 5cg+C"
        End If
        PrintComprobarBoleto mObj
        '
        '   Caso de pruebas 05: Comprobar apuesta multiple bonoloto 5ª categoria
        '
        With mApt
            .Combinacion.Texto = "7-19-14-25-31-44-49"
            .Fecha = #7/7/2020#  ' 7-19-14-26-46-41 C-32 R-1
            .Juego = Bonoloto
        End With
        Debug.Print vbTab & "Apuesta 5ª: " & mObj.ComprobarApuesta(mApt, False)
        If mObj.ImporteApuesta = 16 Then
            Debug.Print vbTab & "OK Prueba correcta "
        Else
            Debug.Print vbTab & " #Error, El importe distinto 16"
        End If
        PrintComprobarBoleto mObj
        '
        '   Caso de pruebas 06: Comprobar apuesta multiple bonoloto 1ª categoria
        '
        With mApt
            .Combinacion.Texto = "7-19-14-26-46-41-49"
            .Fecha = #7/7/2020#  ' 7-19-14-26-46-41 C-32 R-1
            .Juego = Bonoloto
        End With
        Debug.Print vbTab & "Apuesta 1ª: " & mObj.ComprobarApuesta(mApt, False)
        If mObj.ImporteApuesta = 579605.48 Then
            Debug.Print vbTab & "OK Prueba correcta "
        Else
            Debug.Print vbTab & " #Error, El importe distinto 579605,48"
        End If
        PrintComprobarBoleto mObj
        '
        '   Caso de pruebas 07: Comprobar apuesta multiple bonoloto 2ª categoria
        '
        With mApt
            .Combinacion.Texto = "7-19-14-26-46-39-32"
            .Fecha = #7/7/2020#  ' 7-19-14-26-46-41 C-32 R-1
            .Juego = Bonoloto
        End With
        Debug.Print vbTab & "Apuesta 2ª: " & mObj.ComprobarApuesta(mApt, False)
        If mObj.ImporteApuesta = 26444.09 Then
            Debug.Print vbTab & "OK Prueba correcta "
        Else
            Debug.Print vbTab & " #Error, El importe distinto 579605,48"
        End If
        PrintComprobarBoleto mObj
    End If
    '
    '   Comprobación de PRIMITIVA
    '
    If JUEGO_DEFECTO = LoteriaPrimitiva Then
        '
        '   Caso de pruebas 08: Comprobar apuesta simple primitiva 5ª categoria
        '
        With mApt
            .Combinacion.Texto = "7-13-43-19-44-12"
            .Fecha = #7/23/2020#  '7-13-43-34-45-36 C-6 R-4
            .Juego = LoteriaPrimitiva
        End With
        Debug.Print vbTab & "Apuesta 5ª: " & mObj.ComprobarApuesta(mApt, False)
        If mObj.CategoriaPremioTxt = "5ª" Then
            Debug.Print vbTab & "OK => Premio 3cg"
        Else
            Debug.Print vbTab & "#Error Premio distinto de 3cg"
        End If
        PrintComprobarBoleto mObj
        '
        '   Caso de pruebas 09: Comprobar apuesta simple primitiva 1ª categoria
        '
        With mApt
            .Combinacion.Texto = "7-13-43-34-45-36"
            .Fecha = #7/23/2020#  '7-13-43-34-45-36 C-6 R-4
            .Juego = LoteriaPrimitiva
        End With
        Debug.Print vbTab & "Apuesta 1ª: " & mObj.ComprobarApuesta(mApt, False)
        If mObj.CategoriaPremioTxt = "1ª" Then
            Debug.Print vbTab & "OK => Premio 6cg"
        Else
            Debug.Print vbTab & "#Error Premio distinto de 6cg"
        End If
        PrintComprobarBoleto mObj
        '
        '   Caso de pruebas 10: Comprobar apuesta simple primitiva 2ª categoria
        '
        With mApt
            .Combinacion.Texto = "7-13-43-34-45-6"
            .Fecha = #7/23/2020#  '7-13-43-34-45-36 C-6 R-4
            .Juego = LoteriaPrimitiva
        End With
        Debug.Print vbTab & "Apuesta 2ª: " & mObj.ComprobarApuesta(mApt, False)
        If mObj.CategoriaPremioTxt = "2ª" Then
            Debug.Print vbTab & "OK => Premio 5cg+C"
        Else
            Debug.Print vbTab & "#Error Premio distinto de 5cg+C"
        End If
        PrintComprobarBoleto mObj
        '
        '   Caso de pruebas 11: Comprobar apuesta multiple primitiva 5ª categoria
        '
        With mApt
            .Combinacion.Texto = "7-13-43-25-31-44-49"
            .Fecha = #7/23/2020#  '7-13-43-34-45-36 C-6 R-4
            .Juego = LoteriaPrimitiva
        End With
        Debug.Print vbTab & "Apuesta 5ª: " & mObj.ComprobarApuesta(mApt, False)
        If mObj.ImporteApuesta = 16 Then
            Debug.Print vbTab & "OK Prueba correcta "
        Else
            Debug.Print vbTab & " #Error, El importe distinto 16"
        End If
        PrintComprobarBoleto mObj
        '
        '   Caso de pruebas 12: Comprobar apuesta multiple primitiva Esp categoria
        '
        With mApt
            .Combinacion.Texto = "7-13-43-34-45-36-6"
            .Fecha = #7/23/2020#  '7-13-43-34-45-36 C-6 R-4
            .Juego = LoteriaPrimitiva
        End With
        Debug.Print vbTab & "Apuesta Esp: " & mObj.ComprobarApuesta(mApt, False)
        If mObj.ImporteApuesta = 579605.48 Then
            Debug.Print vbTab & "OK Prueba correcta "
        Else
            Debug.Print vbTab & " #Error, El importe distinto 579605,48"
        End If
        PrintComprobarBoleto mObj
        '
        '   Caso de pruebas 13: Comprobar apuesta multiple primitiva 2ª categoria
        '
        With mApt
            .Combinacion.Texto = "7-13-43-34-45-38-6"
            .Fecha = #7/23/2020#  '7-13-43-34-45-36 C-6 R-4
            .Juego = LoteriaPrimitiva
        End With
        Debug.Print vbTab & "Apuesta 2ª: " & mObj.ComprobarApuesta(mApt, False)
        If mObj.ImporteApuesta = 26444.09 Then
            Debug.Print vbTab & "OK Prueba correcta "
        Else
            Debug.Print vbTab & " #Error, El importe distinto 579605,48"
        End If
        PrintComprobarBoleto mObj
    End If
    '
    '   Comprobación de EUROMILLONES
    '
    If JUEGO_DEFECTO = Euromillones Then
        '
        '   Caso de pruebas 14: Comprobar apuesta simple Euromillones 13ª 2 + 0
        '
        With mApt
            .Combinacion.Texto = "29-15-43-19-12"
            .Estrellas.Texto = "5-8"
            .Fecha = #7/21/2020#  '29-15-42-14-24 E-2-4 M-CNL38247
            .Juego = Euromillones
        End With
        Debug.Print vbTab & "Apuesta 13ª: " & mObj.ComprobarApuesta(mApt, False)
        If mObj.CategoriaPremioTxt = "13ª" Then
            Debug.Print vbTab & "OK => Premio 2+0"
        Else
            Debug.Print vbTab & "#Error Premio distinto de 2+0"
        End If
        PrintComprobarBoleto mObj
        '
        '   Caso de pruebas 15: Comprobar apuesta simple Euromillones 1ª categoria
        '
        With mApt
            .Combinacion.Texto = "29-15-42-14-24"
            .Estrellas.Texto = "2-4"
            .Fecha = #7/21/2020#  '29-15-42-14-24 E-2-4 M-CNL38247
            .Juego = Euromillones
        End With
        Debug.Print vbTab & "Apuesta 1ª: " & mObj.ComprobarApuesta(mApt, False)
        If mObj.CategoriaPremioTxt = "1ª" Then
            Debug.Print vbTab & "OK => Premio 5+2"
        Else
            Debug.Print vbTab & "#Error Premio distinto de 5+2"
        End If
        PrintComprobarBoleto mObj
        '
        '   Caso de pruebas 16: Comprobar apuesta simple Euromillones 2ª categoria 5+1
        '
        With mApt
            .Combinacion.Texto = "29-15-42-14-24"
            .Estrellas.Texto = "2-8"
            .Fecha = #7/21/2020#  '29-15-42-14-24 E-2-4 M-CNL38247
            .Juego = Euromillones
        End With
        Debug.Print vbTab & "Apuesta 2ª: " & mObj.ComprobarApuesta(mApt, False)
        If mObj.CategoriaPremioTxt = "2ª" Then
            Debug.Print vbTab & "OK => Premio 5+1"
        Else
            Debug.Print vbTab & "#Error Premio distinto de 5+1"
        End If
        PrintComprobarBoleto mObj
        '
        '   Caso de pruebas 17: Comprobar apuesta multiple (6+3) Euromillones 13ª 2+0
        '
        With mApt
            .Combinacion.Texto = "29-15-43-19-12-8"
            .Estrellas.Texto = "1-8-6"
            .Fecha = #7/21/2020#  '29-15-42-14-24 E-2-4 M-CNL38247
            .Juego = Euromillones
        End With
        Debug.Print vbTab & "Apuesta 13ª: " & mObj.ComprobarApuesta(mApt, False)
        If mObj.ImporteApuesta = 48.84 Then
            Debug.Print vbTab & "OK Prueba correcta "
        Else
            Debug.Print vbTab & " #Error, El importe distinto 16"
        End If
        PrintComprobarBoleto mObj
        '
        '   Caso de pruebas 18: Comprobar apuesta multiple Euromillones 1ª
        '
        With mApt
            .Combinacion.Texto = "29-15-42-14-24-8"
            .Estrellas.Texto = "2-8-4"
            .Fecha = #7/21/2020#  '29-15-42-14-24 E-2-4 M-CNL38247
            .Juego = Euromillones
        End With
        Debug.Print vbTab & "Apuesta 1ª: " & mObj.ComprobarApuesta(mApt, False)
        If mObj.ImporteApuesta = 579605.48 Then
            Debug.Print vbTab & "OK Prueba correcta "
        Else
            Debug.Print vbTab & " #Error, El importe distinto 579605,48"
        End If
        PrintComprobarBoleto mObj
        '
        '   Caso de pruebas 19: Comprobar apuesta multiple Euromillones 2ª categoria
        '
        With mApt
            .Combinacion.Texto = "29-15-42-14-5-8"
            .Estrellas.Texto = "2-8-6"
            .Fecha = #7/21/2020#  '29-15-42-14-24 E-2-4 M-CNL38247
            .Juego = Euromillones
        End With
        Debug.Print vbTab & "Apuesta 2ª: " & mObj.ComprobarApuesta(mApt, False)
        If mObj.ImporteApuesta = 26444.09 Then
            Debug.Print vbTab & "OK Prueba correcta "
        Else
            Debug.Print vbTab & " #Error, El importe distinto 579605,48"
        End If
        PrintComprobarBoleto mObj
    End If
    '
    '   Comprobación de GORDO PRIMITIVA
    '
    If JUEGO_DEFECTO = GordoPrimitiva Then
        '
        '   Caso de pruebas 20: Comprobar apuesta simple Gordo 8ª categoria 2+0
        '
        With mApt
            .Combinacion.Texto = "43-30-42-14-5"
            .Fecha = #7/19/2020#  '43-30-49-50-24  R-9
            .Juego = GordoPrimitiva
        End With
        Debug.Print vbTab & "Apuesta 8ª: " & mObj.ComprobarApuesta(mApt, False)
        If mObj.ImporteApuesta = 3 Then
            Debug.Print vbTab & "OK Prueba correcta "
        Else
            Debug.Print vbTab & " #Error, El importe distinto 3"
        End If
        PrintComprobarBoleto mObj
        '
        '   Caso de pruebas 21: Comprobar apuesta simple Gordo 2ª categoria 5+0
        '
        With mApt
            .Combinacion.Texto = "43-30-49-50-24"
            .Fecha = #7/19/2020#  '43-30-49-50-24  R-9
            .Juego = GordoPrimitiva
        End With
        Debug.Print vbTab & "Apuesta 2ª: " & mObj.ComprobarApuesta(mApt, False)
        If mObj.ImporteApuesta = 0 Then
            Debug.Print vbTab & "OK Prueba correcta " ' 1 de 2ª y 5 de 4ª
        Else
            Debug.Print vbTab & " #Error, El importe distinto 0"
        End If
        PrintComprobarBoleto mObj
        '
        '   Caso de pruebas 22: Comprobar apuesta multiple (8) Gordo 8ª 2+0
        '
        With mApt
            .Combinacion.Texto = "43-30-42-14-5-15-54-20"
            .Fecha = #7/19/2020#  '43-30-49-50-24  R-9
            .Juego = GordoPrimitiva
        End With
        Debug.Print vbTab & "Apuesta 8ª: " & mObj.ComprobarApuesta(mApt, False)
        If mObj.ImporteApuesta = 60 Then
            Debug.Print vbTab & "OK Prueba correcta " ' 20 de 3(8ª)
        Else
            Debug.Print vbTab & " #Error, El importe distinto 60"
        End If
        PrintComprobarBoleto mObj
        '
        '   Caso de pruebas 23: Comprobar apuesta multiple(8) Gordo 2ª categoria 5+0
        '
        With mApt
            .Combinacion.Texto = "43-30-49-50-24-15-54-20"
            .Fecha = #7/19/2020#  '43-30-49-50-24  R-9
            .Juego = GordoPrimitiva
        End With
        Debug.Print vbTab & "Apuesta 8ª: " & mObj.ComprobarApuesta(mApt, False)
        If mObj.ImporteApuesta = 4461.45 Then
            Debug.Print vbTab & "OK Prueba correcta " ' 1 de 2ª+ 15 de 4ª+ 30 de 6ª y 10 de 8ª
        Else
            Debug.Print vbTab & " #Error, El importe distinto 4.461,45"
        End If
        PrintComprobarBoleto mObj
        
    End If
  
  On Error GoTo 0
    Exit Sub
ComprobarApuestaTest_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    Call HandleException(ErrNumber, ErrDescription, "Lot_PqtApuestasTesting.ComprobarApuestaTest", ErrSource)
    Call MsgBox(ErrDescription, vbCritical Or vbSystemModal, ThisWorkbook.Name)
    Call Trace("CERRAR")
End Sub


'------------------------------------------------------------------------------*
' Procedimiento  : ComprobarBoletoTest
' Fecha          : #TODO
' Propósito      : Casos de prueba de la clase ComprobarBoleto
'------------------------------------------------------------------------------*
'
Private Sub ComprobarBoletoTest()
    Dim mObj As ComprobarBoletos
    Dim mBlt As Boleto
    Dim fIni As Date
    Dim Db As BdDatos
    Dim oFila As Range
    
  On Error GoTo ComprobarBoletoTest_Error
    '
    '   Caso de pruebas 01: Objeto vacio
    '
    Set mObj = New ComprobarBoletos
    PrintComprobarBoleto mObj
    
    If JUEGO_DEFECTO = Bonoloto Then
        '
        '   Caso de pruebas 02: Boleto simple sin premios
        '
        fIni = #7/3/2020#  'Viernes 03/07/2020  21-39-34-44-25-2 C-22 R-8
        Set Db = New BdDatos
        Set oFila = Db.GetBoletoByFecha(fIni)
        If Not (oFila Is Nothing) Then
            Set mBlt = New Boleto
            mBlt.Constructor oFila
            mBlt.SetApuestas
            If mBlt.FechaValidez = fIni Then
                Debug.Print ("Resultado del Boleto: " & mObj.ComprobarBoleto(mBlt))
                PrintComprobarBoleto mObj
            Else
                Debug.Print ("#Error en BoletoByFecha: " & mBlt.FechaValidez)
            End If
        Else
            Debug.Print ("#Error en rango: oFila is Nothing")
        End If
        '
        '   Caso de pruebas 03: Boleto multiple sin premios
        '
        With mBlt
            ' 20  29  30  32  33  40
            ' 13  18  26  29  34  41
            .Apuestas.Item(1).Combinacion.Texto = "13-18-26-29-34-41-49"
        End With
        Debug.Print ("Resultado del Boleto: " & mObj.ComprobarBoleto(mBlt))
        PrintComprobarBoleto mObj
        '
        '   Caso de pruebas 04: Boleto Simple con premio
        '
        '
        '   Caso de pruebas 05: Boleto Multiple con premio
        '
        '
    End If
    '
    '   Comprobación de GORDO de la Primitiva
    '
    If JUEGO_DEFECTO = GordoPrimitiva Then
        '
        '   Caso de pruebas 02: Boleto simple sin premios
        '
        fIni = #8/30/2020#  'Domingo 30/08/2020  5-35-9-37-46  R:3
        Set Db = New BdDatos
        Set oFila = Db.GetBoletoByFecha(fIni)
        If Not (oFila Is Nothing) Then
            Set mBlt = New Boleto
            mBlt.Constructor oFila
            mBlt.SetApuestas
            If mBlt.FechaValidez = fIni Then
                Debug.Print ("Resultado del Boleto: " & mObj.ComprobarBoleto(mBlt))
                PrintComprobarBoleto mObj
            Else
                Debug.Print ("#Error en BoletoByFecha: " & mBlt.FechaValidez)
            End If
        Else
            Debug.Print ("#Error en rango: oFila is Nothing")
        End If
        '
        '   Caso de pruebas 03: Boleto multiple (8) sin premios
        '
        With mBlt
            .Apuestas.Item(1).Combinacion.Texto = "5-18-26-29-34-41-49-20"
            .Reintegro = 2
        End With
        Debug.Print ("Resultado del Boleto: " & mObj.ComprobarBoleto(mBlt))
        PrintComprobarBoleto mObj
        '
        '   Caso de pruebas 04: Boleto Simple con premio
        '
        With mBlt
            .Apuestas.Item(1).Combinacion.Texto = "5-35-26-29-34"
            .Reintegro = 3
        End With
        Debug.Print ("Resultado del Boleto: " & mObj.ComprobarBoleto(mBlt))
        PrintComprobarBoleto mObj
        '
        '   Caso de pruebas 05: Boleto Multiple(8) con premio (5ª 3+1
        '
        With mBlt
            .Apuestas.Item(1).Combinacion.Texto = "5-35-9-29-34-41-49-20"
            .Reintegro = 3
            .NumeroApuestas = 56
        End With
        Debug.Print ("Resultado del Boleto: " & mObj.ComprobarBoleto(mBlt)) ' Premio 689,70
        If mObj.ImporteBoleto = 689.7 Then
            Debug.Print " OK Importe = 689,70 "
        Else
            Debug.Print " #Error el importe no es igual a 689,70"
        End If
        PrintComprobarBoleto mObj
    End If
    '
    '   Comprobación de Primitiva
    '
    If JUEGO_DEFECTO = LoteriaPrimitiva Then
    
    End If
    '
    '   Comprobación de Euromillones
    '
    If JUEGO_DEFECTO = Euromillones Then
    
    End If
  On Error GoTo 0
ComprobarBoletoTest__CleanExit:
  Exit Sub
ComprobarBoletoTest_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    Call HandleException(ErrNumber, ErrDescription, "Lot_PqtApuestasTesting.ComprobarBoletoTest", ErrSource)
    Call MsgBox(ErrDescription, vbCritical Or vbSystemModal, ThisWorkbook.Name)
    Call Trace("CERRAR")
End Sub




Private Sub PrintComprobarBoleto(mObj As ComprobarBoletos)
    Debug.Print "==> Pintar ComprobarBoletos"
    Debug.Print "  ComprobarApuesta()  = #Method" ' mObj.ComprobarApuesta
    Debug.Print "  ComprobarBoleto()   = #Method" ' mObj.ComprobarBoleto
    Debug.Print "  BolasAcertadas      = " & mObj.BolasAcertadas
    Debug.Print "  CategoriaPremioTxt  = " & mObj.CategoriaPremioTxt
    Debug.Print "  CatPremioApuesta    = " & mObj.CatPremioApuesta
    Debug.Print "  CatPremioBoleto     = " & mObj.CatPremioBoleto
    Debug.Print "  ImporteApuesta      = " & mObj.ImporteApuesta
    Debug.Print "  ImporteBoleto       = " & mObj.ImporteBoleto
    Debug.Print "  NumerosAcertados    = " & mObj.NumerosAcertados
    Debug.Print "  Premio              = " & mObj.Premio.ToString()
    Debug.Print "  Sorteo()            = " & mObj.Sorteo.ToString()
End Sub

' *===========(EOF): Lot_PqtApuestasTesting.bas
