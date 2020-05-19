Attribute VB_Name = "Lot_Testing"
'---------------------------------------------------------------------------------------
' Module    : Lot_Testing
' Author    : CHARLY
' Date      : sáb, 01/nov/2014 21:01:24
' Purpose   :
'---------------------------------------------------------------------------------------
'
Option Explicit
Option Base 0

'---------------------------------------------------------------------------------------
' Procedure : AciertoTest
' Author    : CHARLY
' Date      : sáb, 01/nov/2014 21:01:34
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub AciertoTest()
    Dim obj As Acierto
    Set obj = New Acierto
    With obj
        .ApuestasAcertadas = 1
        .BolasAcertadas = 2
        .Categoria = Duodecima
        .EstrellasAcertadas = 2
        .IdApuesta = 15
        .ImportePremio = 12
        .Juego = Bonoloto
        .ReintegroAcertado = True
    End With
    PrintAcierto obj
End Sub

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
    Dim obj As CU_ComprobarApuesta
    Dim oSorteo As Sorteo
    Dim oApuesta As Apuesta
   On Error GoTo CU_ComprobarApuestaTest_Error
    Set obj = New CU_ComprobarApuesta
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
' Procedure : PremioTest
' Author    : CHARLY
' Date      : sáb, 01/nov/2014 21:01:41
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Sub PremioTest()
    Dim m_obj As Premio
    Set m_obj = New Premio
    m_obj.BolasAcertadas = 6
    m_obj.ComplementarioAcertado = True
    m_obj.ClaveAcertada = True
    m_obj.Pronosticos = 7
    
    Debug.Print "==> Pruebas Premio"
    Debug.Print "Key                     = " & m_obj.key
    Debug.Print "BolasAcertadas          = " & m_obj.BolasAcertadas
    Debug.Print "ComplementarioAcertado  = " & m_obj.ComplementarioAcertado
    Debug.Print "NumeroEstrellasAcertadas= " & m_obj.NumeroEstrellasAcertadas
    Debug.Print "ClaveAcertada           = " & m_obj.ClaveAcertada
    Debug.Print "Pronosticos             = " & m_obj.Pronosticos
    Debug.Print "CategoriaPremio         = " & m_obj.CategoriaPremio
    Debug.Print "FechaSorteo             = " & m_obj.FechaSorteo
    Debug.Print "ModalidadJuego          = " & m_obj.ModalidadJuego
    Debug.Print "LiteralCategoriaPremio  = " & m_obj.LiteralCategoriaPremio
    Debug.Print "GetPremioEsperado()     = " & m_obj.GetPremioEsperado()
    m_obj.BolasAcertadas = 0
    Debug.Print "GetPremioEsperado()     = " & m_obj.GetPremioEsperado()
    m_obj.BolasAcertadas = 3
    Debug.Print "GetPremioEsperado()     = " & m_obj.GetPremioEsperado()
    m_obj.BolasAcertadas = 4
    Debug.Print "GetPremioEsperado()     = " & m_obj.GetPremioEsperado()
    m_obj.BolasAcertadas = 5
    Debug.Print "GetPremioEsperado()     = " & m_obj.GetPremioEsperado()
    m_obj.BolasAcertadas = 6
    Debug.Print "GetPremioEsperado()     = " & m_obj.GetPremioEsperado()
    m_obj.BolasAcertadas = 7
    Debug.Print "GetPremioEsperado()     = " & m_obj.GetPremioEsperado()

End Sub

'---------------------------------------------------------------------------------------
' Procedure : Premio2Test
' Author    : CHARLY
' Date      : sáb, 01/nov/2014 21:01:42
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub Premio2Test()
    Dim obj As Premio2
    Set obj = New Premio2
    Dim obj1 As Premio2
    Set obj1 = New Premio2
    With obj
        .CategoriaPremio = Cuarta
        .Importe = 125
        .Juego = Bonoloto
        .NumeroAcertantesEspaña = 20
        .NumeroAcertantesEuropa = 15
    End With
    PrintPremio2 obj
    'obj.Parse
    obj1.Parse obj.ToString()
    PrintPremio2 obj1
End Sub


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
' Procedure : EjemploColecciones
' Author    : CHARLY
' Date      : sáb, 01/nov/2014 21:01:51
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Sub EjemploColecciones()
    Dim mComb As New Combinacion
    Dim mNum As New Numero
    mNum.Valor = 25: mComb.Add mNum
    Dim mNum1 As New Numero
    mNum1.Valor = 36: mComb.Add mNum1
    Dim mNum2 As New Numero
    mNum2.Valor = 4: mComb.Add mNum2
    Dim mNum3 As New Numero
    mNum3.Valor = 9: mComb.Add mNum3
    Dim mNum4 As New Numero
    mNum4.Valor = 10: mComb.Add mNum4
    Dim mNum5 As New Numero
    mNum5.Valor = 45: mComb.Add mNum5
    
    Debug.Print "Apuesta" & mComb.TextoOrdenado
    Debug.Print "Devuelve el Numero (5) => " & mComb.Contiene(5)
    Debug.Print "Contiene el 9? " & mComb.Contiene(9)
    Debug.Print "EstaOrdenado:" & mComb.EstaOrdenado
End Sub

'---------------------------------------------------------------------------------------
' Procedure : ParametrosTest
' Author    : CHARLY
' Date      : sáb, 01/nov/2014 21:01:53
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub ParametrosTest()
    Dim obj As Parametros
    
    
End Sub

'---------------------------------------------------------------------------------------
' Procedure : ParametroTest
' Author    : CHARLY
' Date      : sáb, 01/nov/2014 21:02:02
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Sub ParametroTest()
    Dim oPar As Parametro
    Dim iPar As Integer
    Dim dPar As Date
    Dim bPar As Boolean
    Dim pPar As Double
    
    
    Set oPar = New Parametro
    With oPar
        .Descripcion = "Esta es una variable de prueba"
        .Nombre = "VARIABLE"
        .Tipo = parTexto
        .Valor = "Ejemplo"
    End With
    PrintParametro oPar
    '
    ' Prueba entero
    '
    iPar = 3294
    oPar.Valor = iPar
    PrintParametro oPar
    '
    ' prueba fecha
    '
    dPar = #1/1/2014#
    oPar.Valor = dPar
    PrintParametro oPar
    '
    ' prueba Doble
    '
    pPar = 12536.254
    oPar.Valor = pPar
    PrintParametro oPar
   
End Sub

'---------------------------------------------------------------------------------------
' Procedure : ParametrosEngineTest
' Author    : CHARLY
' Date      : sáb, 01/nov/2014 21:02:04
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub ParametrosEngineTest()

End Sub

'---------------------------------------------------------------------------------------
' Procedure : PrintAcierto
' Author    : CHARLY
' Date      : sáb, 01/nov/2014 21:02:07
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub PrintAcierto(oDat As Acierto)
    Debug.Print "==> Pruebas Acierto"
    Debug.Print " .ApuestasAcertadas  = " & oDat.ApuestasAcertadas
    Debug.Print " .BolasAcertadas     = " & oDat.BolasAcertadas
    Debug.Print " .Categoria          = " & oDat.Categoria
    Debug.Print " .EstrellasAcertadas = " & oDat.EstrellasAcertadas
    Debug.Print " .IdApuesta          = " & oDat.IdApuesta
    Debug.Print " .ImportePremio      = " & oDat.ImportePremio
    Debug.Print " .Juego              = " & oDat.Juego
    Debug.Print " .ReintegroAcertado  = " & oDat.ReintegroAcertado
End Sub

'---------------------------------------------------------------------------------------
' Procedure : PrintCombinacion
' Author    : CHARLY
' Date      : sáb, 01/nov/2014 21:02:09
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub PrintCombinacion(obj As Combinacion)
    Debug.Print "==> Combinacion"
    Debug.Print " .Count               = " & obj.Count
    Debug.Print " .FormulaAltoBajo     = " & obj.FormulaAltoBajo
    Debug.Print " .FormulaConsecutivos = " & obj.FormulaConsecutivos
    Debug.Print " .FormulaDecenas      = " & obj.FormulaDecenas
    Debug.Print " .FormulaParidad      = " & obj.FormulaParidad
    Debug.Print " .FormulaSeptenas     = " & obj.FormulaSeptenas
    Debug.Print " .FormulaTerminacion  = " & obj.FormulaTerminaciones
    Debug.Print " .Producto            = " & obj.Producto
    Debug.Print " .Suma                = " & obj.Suma
    Debug.Print " .Texto               = " & obj.Texto
    Debug.Print " .TextoOrdenado       = " & obj.TextoOrdenado
    Debug.Print " .EstaOrdenado()      = " & obj.EstaOrdenado
End Sub

'---------------------------------------------------------------------------------------
' Procedure : PrintPremio
' Author    : Charly
' Date      : 17/11/2013
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub PrintPremio(objPremio As Premio)

  On Error GoTo PrintPremio_Error
    Debug.Print " Key                    =>" & objPremio.key & " ==========="
    Debug.Print "   BolasAcertadas       =>" & objPremio.BolasAcertadas
    Debug.Print "   CategoriaPremio      =>" & objPremio.CategoriaPremio
    Debug.Print "   Complementario       =>" & objPremio.ComplementarioAcertado
    Debug.Print "   FechaSorteo          =>" & objPremio.FechaSorteo
    Debug.Print "   GetPremioEsperado    =>" & objPremio.GetPremioEsperado
    Debug.Print "   ModalidadJuego       =>" & objPremio.ModalidadJuego
        

   On Error GoTo 0
   Exit Sub

PrintPremio_Error:
     Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
   ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
   '   Audita el error
   Call HandleException(ErrNumber, ErrDescription, "Lot_01_ComprobarApuestas.PrintPremio")
   '   Lanza el Error
   Err.Raise ErrNumber, ErrSource, ErrDescription
End Sub

'---------------------------------------------------------------------------------------
' Procedure : PrintPremio2
' Author    : CHARLY
' Date      : 06/04/2014
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub PrintPremio2(obj As Premio2)
    Debug.Print " Premio ]================="
    Debug.Print "   CategoriaPremio        =>" & obj.CategoriaPremio
    Debug.Print "   CategoriaTexto         =>" & obj.CategoriaTexto
    Debug.Print "   Importe                =>" & obj.Importe
    Debug.Print "   Juego                  =>" & obj.Juego
    Debug.Print "   Acertantes en Europa   =>" & obj.NumeroAcertantesEuropa
    Debug.Print "   Acertantes en España   =>" & obj.NumeroAcertantesEspaña
    Debug.Print "   EsValido()             =>" & obj.EsValido()
    Debug.Print "   ToString()             =>" & obj.ToString()
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
    Debug.Print "  Metodo         = " & datApuesta.metodo
    Debug.Print "  NumeroApuestas = " & datApuesta.NumeroApuestas
    Debug.Print "  Pronosticos    = " & datApuesta.Pronosticos
    Debug.Print "  SeHaJugado     = " & datApuesta.SeHaJugado
    Debug.Print "  Texto          = " & datApuesta.Texto
End Sub

'---------------------------------------------------------------------------------------
' Procedure : PrintParametro
' Author    : CHARLY
' Date      : sáb, 01/nov/2014 21:02:19
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub PrintParametro(obj As Parametro)
    Debug.Print "==> Parametro "
    Debug.Print "Descripcion     = " & obj.Descripcion
    Debug.Print "EntidadNegocio  = " & obj.EntidadNegocio.ClassStorage
    Debug.Print "Fecha Alta      = " & obj.FechaAlta
    Debug.Print "Fecha Modif.    = " & obj.FechaModificacion
    Debug.Print "Id              = " & obj.Id
    Debug.Print "Nombre          = " & obj.Nombre
    Debug.Print "Orden           = " & obj.Orden
    Debug.Print "Tipo            = " & obj.Tipo
    Debug.Print "Valor           = " & obj.Valor
End Sub

'---------------------------------------------------------------------------------------
' Procedure : ParametrosMetodoTest
' Author    : Charly
' Date      : 19/03/2012
' Purpose   : Probar la clase ParametrosMetodo
'---------------------------------------------------------------------------------------
'
Private Sub ParametrosMetodoTest()
    Dim m_objParMetodo As ParametrosMetodo
    
    Set m_objParMetodo = New ParametrosMetodo
    
    
    With m_objParMetodo
        .CriteriosAgrupacion = grpDecenas
        .CriteriosOrdenacion = ordProbabilidad
        .DiasAnalisis = 45
        .Id = 1
        .ModalidadJuego = LP_LB_6_49
        .NumeroSorteos = 40
        .Orden = 1
        .Pronosticos = 6
        .SentidoOrdenacion = True
    End With

    Debug.Print "==> Pruebas ParametrosMetodoTest"
    Debug.Print "Id                       = " & m_objParMetodo.Id
    Debug.Print "Juego                    = " & m_objParMetodo.ModalidadJuego
    Debug.Print "Criterio Ordenación      = " & m_objParMetodo.CriteriosOrdenacion
    Debug.Print "Criterio Agrupación      = " & m_objParMetodo.CriteriosAgrupacion
    Debug.Print "Dias de Analisis         = " & m_objParMetodo.DiasAnalisis
    Debug.Print "Numero de Sorteos        = " & m_objParMetodo.NumeroSorteos
    Debug.Print "Orden                    = " & m_objParMetodo.Orden
    Debug.Print "Pronosticos              = " & m_objParMetodo.Pronosticos
    Debug.Print "Sentido de la Ordenación = " & m_objParMetodo.SentidoOrdenacion
    Debug.Print "OrdenacionToString()     = " & m_objParMetodo.OrdenacionToString()
    Debug.Print "AgrupacionToString()     = " & m_objParMetodo.AgrupacionToString()
    Debug.Print "ToString()               = " & m_objParMetodo.ToString()

End Sub


'---------------------------------------------------------------------------------------
' Procedure : MetodoTest
' Author    : Charly
' Date      : 19/03/2012
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Sub MetodoTest()
    Dim m_objMetodo As metodo
    
    Set m_objMetodo = New metodo
    
    With m_objMetodo
        .TipoProcedimiento = mtdEstadistico
        .EntidadNegocio.FechaModificacion = Date
        .EsMultiple = False
        .Parametros.CriteriosAgrupacion = grpParidad
        .Parametros.CriteriosOrdenacion = ordDesviacion
        .Parametros.DiasAnalisis = 42
        .Parametros.SentidoOrdenacion = True
        .TipoMuestra = True
    End With
    
    
    Debug.Print "==> Pruebas Metodo"
    Debug.Print "ClassStorage           =" & m_objMetodo.EntidadNegocio.ClassStorage
    Debug.Print "FechaAlta              =" & m_objMetodo.EntidadNegocio.FechaAlta
    Debug.Print "FechaBaja              =" & m_objMetodo.EntidadNegocio.FechaBaja
    Debug.Print "FechaModificacion      =" & m_objMetodo.EntidadNegocio.FechaModificacion
    Debug.Print "Id                     =" & m_objMetodo.EntidadNegocio.Id
    Debug.Print "IsDirty                =" & m_objMetodo.EntidadNegocio.IsDirty
    Debug.Print "IsNew                  =" & m_objMetodo.EntidadNegocio.IsNew
    Debug.Print "MarkForDelete          =" & m_objMetodo.EntidadNegocio.MarkForDelete
    Debug.Print "Origen                 =" & m_objMetodo.EntidadNegocio.Origen
    Debug.Print "Situacion              =" & m_objMetodo.EntidadNegocio.Situacion
    Debug.Print "EsMultiple             =" & m_objMetodo.EsMultiple
    Debug.Print "AgrupacionToString     =" & m_objMetodo.Parametros.AgrupacionToString
    Debug.Print "CriteriosAgrupacion    =" & m_objMetodo.Parametros.CriteriosAgrupacion
    Debug.Print "CriteriosOrdenacion    =" & m_objMetodo.Parametros.CriteriosOrdenacion
    Debug.Print "DiasAnalisis           =" & m_objMetodo.Parametros.DiasAnalisis
    Debug.Print "Id                     =" & m_objMetodo.Parametros.Id
    Debug.Print "ModalidadJuego         =" & m_objMetodo.Parametros.ModalidadJuego
    Debug.Print "NumeroSorteos          =" & m_objMetodo.Parametros.NumeroSorteos
    Debug.Print "Orden                  =" & m_objMetodo.Parametros.Orden
    Debug.Print "OrdenacionToString     =" & m_objMetodo.Parametros.OrdenacionToString
    Debug.Print "Pronosticos            =" & m_objMetodo.Parametros.Pronosticos
    Debug.Print "SentidoOrdenacion      =" & m_objMetodo.Parametros.SentidoOrdenacion
    Debug.Print "ToString               =" & m_objMetodo.Parametros.ToString
    Debug.Print "TipoMuestra            =" & m_objMetodo.TipoMuestra
    Debug.Print "TipoProcedimiento      =" & m_objMetodo.TipoProcedimiento
       
End Sub


Private Sub InfoSorteoTest()
    Dim mInfo As InfoSorteo
    Dim i As Integer
    Dim mFechaI As Date
    Dim mFechaF As Date

    
    Set mInfo = New InfoSorteo
    '
    '  21/5/2014 Miercoles
    '
    mFechaI = #5/21/2014#
    Debug.Print "==> Pruebas InfoSorteo"
    
    For i = 0 To 7
        mFechaI = mFechaI + i
        mInfo.Constructor Bonoloto
        mFechaF = mInfo.GetProximoSorteo(mFechaI)
        Debug.Print "GetProximoSorteo(" & mFechaI & ", Bonoloto) => "; mFechaF
        Debug.Print "EsFechaSorteo (" & mFechaI & ", Bonoloto) => " & mInfo.EsFechaSorteo(mFechaI)
        mInfo.Constructor gordoPrimitiva
        mFechaF = mInfo.GetProximoSorteo(mFechaI)
        Debug.Print "GetProximoSorteo(" & mFechaI & ", GordoPrimitiva) => "; mFechaF
        mInfo.Constructor LoteriaPrimitiva
        mFechaF = mInfo.GetProximoSorteo(mFechaI)
        Debug.Print "GetProximoSorteo(" & mFechaI & ", LoteriaPrimitiva) => "; mFechaF
        mInfo.Constructor Euromillones
        mFechaF = mInfo.GetProximoSorteo(mFechaI)
        
        Debug.Print "GetProximoSorteo(" & mFechaI & ", Euromillones) => "; mFechaF
        Debug.Print "EsFechaSorteo (" & mFechaI & ", Bonoloto) => " & mInfo.EsFechaSorteo(mFechaI)
    Next i
    '
    '   Sorteos entre dos fechas
    '
    mFechaI = #4/26/2015#   'Domingo
    mFechaF = mFechaI
    For i = 1 To 26
        Debug.Print "Sorteos entre" & mFechaI & " y " & mFechaF
        Debug.Print "   ==>" & mInfo.GetSorteosEntreFechas(mFechaI, mFechaF)
        mFechaF = mFechaF + 1
    Next i
    

End Sub
