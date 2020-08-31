Attribute VB_Name = "Lot_PqtNucleoTesting"
' *============================================================================*
' *
' *     Fichero    : Lot_PqtNucleoTesting.bas
' *
' *     Autor      : Carlos Almela Baeza
' *     Creación   : sáb, 01/nov/2014 21:01:24
' *     Revisión   : ju., 02/jul/2020 19:04:47
' *     Versión    : 1.1
' *     Propósito  : Módulo de pruebas de las clases:
' *                  - Periodo
' *                  - BdDatos
' *                  - ParamProceso
' *                  - Numero
' *                  - Combinación
' *
' *============================================================================*
Option Explicit
Option Base 0




'---------------------------------------------------------------------------------------
' Procedure : PeriodoTest
' Author    : CHARLY
' Date      : sáb, 01/nov/2014 21:01:47
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub PeriodoTest()
    Dim obj As Periodo
    Dim cboPrueba As ComboBox
    Dim frm As frmSelPeriodo
    Dim mLista As Variant
    
  On Error GoTo PeriodoTest_Error
    Set obj = New Periodo
    Set frm = New frmSelPeriodo
    Set cboPrueba = frm.cboPerMuestra
    
    mLista = Array(ctPersonalizadas, ctSemanaPasada, ctSemanaActual, ctMesActual, ctHoy, ctAyer, ctLoQueVadeMes, _
                                     ctLoQueVadeSemana)
    
    obj.CargaCombo cboPrueba, mLista
    
    PintarPeriodo obj
    obj.Tipo_Fecha = ctAñoAnterior
    PintarPeriodo obj
    
    obj.Tipo_Fecha = ctLoQueVadeSemana
    PintarPeriodo obj
    
    Set obj = New Periodo
    obj.FechaInicial = #1/5/2020#
    obj.FechaFinal = #12/5/2020#
    PintarPeriodo obj
    
 On Error GoTo 0
BdDatosTest_CleanExit:
    Exit Sub
PeriodoTest_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    Call HandleException(ErrNumber, ErrDescription, "Lot_PqtNucleoTesting.PeriodoTest", ErrSource)
    Call MsgBox(ErrDescription, vbCritical Or vbSystemModal, Application)
    Call Trace("CERRAR")
End Sub




'---------------------------------------------------------------------------------------
' Procedure : BdDatosTest
' Author    : CHARLY
' Date      : mar, 16/jun/2020 13:12:55
' Purpose   : Probar la clase BdDatos
'---------------------------------------------------------------------------------------
'
Private Sub BdDatosTest()
    Dim Bd   As BdDatos         ' Base de Datos
    Dim obj  As Range           ' Rango resultado
    Dim fIni As Date            ' Fecha de inicio
    Dim fFin As Date            ' fecha de Fin
    Dim mPer As Periodo         ' Periodo de fechas
    Dim mCol As Collection      ' Colección de apariciones
    Dim i    As Integer         ' Entero de trabajo
On Error GoTo BdDatosTest_Error
    '
    '  Caso de Prueba 01 Rango de fechas válido
    '
    Debug.Print "#============= TestCase: 01 "
    
    Set Bd = New BdDatos
    '
    '   Bonoloto y Primitiva
    '
    fIni = #5/28/2020#   ' Sáb
    fFin = #6/13/2020#   ' Sáb
    Set mPer = New Periodo
    mPer.FechaInicial = fIni
    mPer.FechaFinal = fFin
    Set obj = Bd.GetSorteosInFechas(mPer)
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
    Debug.Print "#============= TestCase: 02 "
    fIni = #3/19/2020#   ' Jue (sin Sorteo)
    fFin = #6/13/2020#   ' Sáb
    Set mPer = New Periodo
    mPer.FechaInicial = fIni
    mPer.FechaFinal = fFin
    Set obj = Bd.GetSorteosInFechas(mPer)
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
    Debug.Print "#============= TestCase: 03 "
    fIni = #2/8/2020#    ' Sáb
    fFin = #3/19/2020#   ' Jue Sin Sorteo
    Set mPer = New Periodo
    mPer.FechaInicial = fIni
    mPer.FechaFinal = fFin
    Set obj = Bd.GetSorteosInFechas(mPer)
    '
    '
    If "$A$1576:$N$1606" = obj.Address And JUEGO_DEFECTO = Bonoloto Then
        Debug.Print ("Rango seleccionado Correcto: " & obj.Address)
    ElseIf "$A$1610:$N$1624" = obj.Address And JUEGO_DEFECTO = LoteriaPrimitiva Then
        Debug.Print ("Rango seleccionado Correcto: " & obj.Address)
    Else
        Debug.Print (" #Error en rango: " & obj.Address)
    End If
    '
    '   Caso de Prueba 04 Propiedades del objeto
    '
    Debug.Print "#============= TestCase: 04 "
    Debug.Print "Primer Sorteo    : " & Bd.PrimerResultado
    Debug.Print "Ultimo Sorteo    : " & Bd.UltimoResultado
    Debug.Print "Ultimo registro  : " & Bd.UltimoRegistro
    Debug.Print "Address Apuestas : " & Bd.RangoApuestas.Address
    Debug.Print "Address Sorteos  : " & Bd.RangoResultados.Address
    Debug.Print "Address Boletos  : " & Bd.RangoBoletos.Address
    '
    ' Caso de prueba 05 Boleto de una fecha
    '
    Debug.Print "#============= TestCase: 05 "
    fFin = #3/2/2020#
    Set obj = Bd.GetBoletoByFecha(fFin)
    If Not (obj Is Nothing) Then
        If obj.Address = "$A$56:$P$56" Then
            Debug.Print ("Rango seleccionado Correcto: " & obj.Address)
        Else
            Debug.Print ("#Error en rango: " & obj.Address)
        End If
    Else
        Debug.Print ("#Error en rango: obj is Nothing")
    End If
    '
    ' Caso de prueba 06 Boletos entre fechas
    '
    Debug.Print "#============= TestCase: 06 "
    Set mPer = New Periodo
    mPer.FechaInicial = #6/22/2020#
    mPer.FechaFinal = #6/27/2020#
    Set obj = Bd.GetBoletoInFechas(mPer)
    If Not (obj Is Nothing) Then
        If obj.Address = "$A$92:$P$97" Then
            Debug.Print ("Rango seleccionado Correcto: " & obj.Address)
        Else
            Debug.Print ("#Error en rango: " & obj.Address)
        End If
    Else
        Debug.Print ("#Error en rango: obj is Nothing")
    End If
    
    '
    ' Caso de prueba 07 Boleto de un Id
    '
    Debug.Print "#============= TestCase: 07 "
    Set obj = Bd.GetBoletoById(70)
    If Not (obj Is Nothing) Then
        If obj.Address = "$A$71:$P$71" Then
            Debug.Print ("Rango seleccionado Correcto: " & obj.Address)
        Else
            Debug.Print ("#Error en rango: " & obj.Address)
        End If
    Else
        Debug.Print ("#Error en rango: obj is Nothing")
    End If
    '
    ' Caso de prueba 08 Apuestas de un Boleto
    '
    Debug.Print "#============= TestCase: 08 "
    Set obj = Bd.GetApuestaByBoleto(53)
    If Not (obj Is Nothing) Then
        If obj.Address = "$A$106:$P$107" Then
            Debug.Print ("Rango seleccionado Correcto: " & obj.Address)
        Else
            Debug.Print ("#Error en rango: " & obj.Address)
        End If
    Else
        Debug.Print ("#Error en rango: obj is Nothing")
    End If
    '
    ' Caso de prueba 09 Apuestas de una fecha
    '
    Debug.Print "#============= TestCase: 09 "
    fFin = #3/2/2020#
    Set obj = Bd.GetApuestaByFecha(fFin)
    If Not (obj Is Nothing) Then
        If obj.Address = "$A$110:$Y$110" Then
            Debug.Print ("Rango seleccionado Correcto: " & obj.Address)
        Else
            Debug.Print ("#Error en rango: " & obj.Address)
        End If
    Else
        Debug.Print ("#Error en rango: obj is Nothing")
    End If
    '
    ' Caso de prueba 10 Apuestas entre fechas
    '
    Debug.Print "#============= TestCase: 10 "
    mPer.FechaInicial = #8/11/2020#
    mPer.FechaFinal = #8/12/2020#
    Set obj = Bd.GetApuestaInFechas(mPer)
    If Not (obj Is Nothing) Then
        If obj.Address = "$A$268:$Y$271" Then
            Debug.Print ("Rango seleccionado Correcto: " & obj.Address)
        Else
            Debug.Print ("#Error en rango: " & obj.Address)
        End If
    Else
        Debug.Print ("#Error en rango: obj is Nothing")
    End If
    '
    ' Caso de prueba 11 Sorteo de una fecha
    '
    Debug.Print "#============= TestCase: 11 "
    fFin = #6/16/2020#
    Set obj = Bd.GetSorteoByFecha(fFin)
    If Not (obj Is Nothing) Then
        If obj.Address = "$A$1626:$N$1626" Then
            Debug.Print ("Rango seleccionado Correcto: " & obj.Address)
        Else
            Debug.Print ("#Error en rango: " & obj.Address)
        End If
    Else
        Debug.Print ("#Error en rango: obj is Nothing")
    End If
    '
    ' Caso de prueba 12 Sorteo de un Id
    '
    Debug.Print "#============= TestCase: 12 "
    Set obj = Bd.GetSorteoById(2656)
    If Not (obj Is Nothing) Then
        If obj.Address = "$A$1621:$N$1621" Then
            Debug.Print ("Rango seleccionado Correcto: " & obj.Address)
        Else
            Debug.Print ("#Error en rango: " & obj.Address)
        End If
    Else
        Debug.Print ("#Error en rango: obj is Nothing")
    End If
    '
    ' Caso de prueba 13 Sorteo de un Periodo
    '
    Debug.Print "#============= TestCase: 13 "
    mPer.FechaInicial = #6/22/2020#
    mPer.FechaFinal = #6/27/2020#
    Set obj = Bd.GetSorteosInFechas(mPer)
    If Not (obj Is Nothing) Then
        If obj.Address = "$A$1631:$N$1636" Then
            Debug.Print ("Rango seleccionado Correcto: " & obj.Address)
        Else
            Debug.Print ("#Error en rango: " & obj.Address)
        End If
    Else
        Debug.Print ("#Error en rango: obj is Nothing")
    End If
    '
    ' Caso de Prueba 14 Apariciones de 1 numero
    '
    Debug.Print "#============= TestCase: 14 "
    Set mCol = New Collection
    Set mCol = Bd.GetAparicionesNumero(15)
    If Not (mCol Is Nothing) Then
        If mCol.Count = 233 Then
            Debug.Print ("Apariciones totales del N(15): " & mCol.Count)
        Else
            Debug.Print ("#Error en Apariciones: " & mCol.Count)
        End If
    Else
        Debug.Print ("#Error en rango: mCol is Nothing")
    End If
    '
    ' Caso de prueba 15 Apariciones de 1 reintegro entre fechas
    '
    Debug.Print "#============= TestCase: 15 "
    Set mCol = Bd.GetAparicionesReintegro(0)
    If Not (mCol Is Nothing) Then
        If mCol.Count = 164 Then
            Debug.Print ("Apariciones totales del reintegro (0): " & mCol.Count)
        Else
            Debug.Print ("#Error en Apariciones: " & mCol.Count)
        End If
    Else
        Debug.Print ("#Error en rango: mCol is Nothing")
    End If
    '
    ' Caso de prueba 16 Apariciones de 1 estrella entre fechas
    '
    Debug.Print "#============= TestCase: 16 "
    Set mCol = Bd.GetAparicionesEstrella(12)
    If Not (mCol Is Nothing) Then
        If mCol.Count = 62 Then
            Debug.Print ("Apariciones totales de la estrella (12): " & mCol.Count)
        Else
            Debug.Print ("#Error en Apariciones: " & mCol.Count)
        End If
    Else
        Debug.Print ("#Error en rango: mCol is Nothing")
    End If
    '
    ' Caso de prueba 17 Obtener el registro de una fecha de sorteo existente
    '
    Debug.Print "#============= TestCase: 17 "
    fIni = #5/28/2020#   ' Sáb

    i = Bd.GetRegistroFecha(fIni)
    If i = 2645 Then
        Debug.Print ("Registro asociado al sorteo de fecha (" & fIni & ") es igual: " & CStr(i))
    Else
        Debug.Print ("#Error registro erroneo: " & CStr(i))
    End If
    '
    ' Caso de prueba 18 Obtener el registro de una fecha de sorteo inexistente
    '
    Debug.Print "#============= TestCase: 18 "
    fIni = #5/29/2020#   ' Dom
    i = Bd.GetRegistroFecha(fIni)
    If i = 2645 Then
        Debug.Print ("Registro asociado al sorteo de fecha (" & fIni & ") es igual: " & CStr(i))
    Else
        Debug.Print ("#Error registro erroneo: " & CStr(i))
    End If
    
  On Error GoTo 0
BdDatosTest_CleanExit:
    Exit Sub
            
BdDatosTest_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    Call HandleException(ErrNumber, ErrDescription, "Lot_PqtNucleoTesting.BdDatosTest", ErrSource)
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
    Debug.Print "ToString()    = " & datPeriodo.ToString()
End Sub




'------------------------------------------------------------------------------*
' Procedimiento  : NumeroTest
' Fecha          : 29/Abr/2018
' Propósito      : Pruebas Unitarias de la clase Numero
'------------------------------------------------------------------------------*
'
Private Sub NumeroTest()
    Dim obj As Numero
  
  On Error GoTo NumeroTest_Error
    Set obj = New Numero
    '
    '   Numero Válido
    '
    obj.Valor = 5
    PrintNumero obj
    '
    '   Numero no Valido
    '
    Set obj = New Numero
    obj.Valor = 80
    PrintNumero obj
      
  On Error GoTo 0
NumeroTest__CleanExit:
    Exit Sub
            
NumeroTest_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    Call HandleException(ErrNumber, ErrDescription, "Lot_PqtConcursosTesting.NumeroTest", ErrSource)
    Call MsgBox(ErrDescription, vbCritical Or vbSystemModal, Application)
    Call Trace("CERRAR")
End Sub




'------------------------------------------------------------------------------*
' Procedimiento  : CombinacionTest
' Fecha          : 29/Abr/2018
' Propósito      : Pruebas Unitarias de la clase Combinacion
'------------------------------------------------------------------------------*
'
Private Sub CombinacionTest()
    Dim obj     As Combinacion
    Dim oNum    As Numero
  On Error GoTo CombinacionTest_Error
    '
    '   Combinacion Vacia
    '
    Set obj = New Combinacion
    PrintCombinacion obj
    '
    '   Combinacion Valida  28-1-31-25-33-8-7
    '
    Set oNum = New Numero
    oNum.Valor = 28
    obj.Add oNum
    Set oNum = New Numero
    oNum.Valor = 1
    obj.Add oNum
    Set oNum = New Numero
    oNum.Valor = 31
    obj.Add oNum
    Set oNum = New Numero
    oNum.Valor = 25
    obj.Add oNum
    Set oNum = New Numero
    oNum.Valor = 33
    obj.Add oNum
    Set oNum = New Numero
    oNum.Valor = 8
    obj.Add oNum
    Set oNum = New Numero
    oNum.Valor = 7
    obj.Add oNum
    '
    '
    PrintCombinacion obj
    '
    '   Combinación no Valida 56-32-14-7-9
    '
    Set obj = New Combinacion
    Set oNum = New Numero
    oNum.Valor = 56
    obj.Add oNum
    Set oNum = New Numero
    oNum.Valor = 32
    obj.Add oNum
    Set oNum = New Numero
    oNum.Valor = 14
    obj.Add oNum
    Set oNum = New Numero
    oNum.Valor = 7
    obj.Add oNum
    Set oNum = New Numero
    oNum.Valor = 9
    obj.Add oNum
    '
    '
    PrintCombinacion obj
      
  On Error GoTo 0
CombinacionTest__CleanExit:
    Exit Sub
            
CombinacionTest_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    Call HandleException(ErrNumber, ErrDescription, "Lot_PqtConcursosTesting.CombinacionTest", ErrSource)
    Call MsgBox(ErrDescription, vbCritical Or vbSystemModal, Application)
    Call Trace("CERRAR")
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
    
    Debug.Print "Apuesta:" & mComb.ToString(True)
    Debug.Print "Devuelve el Numero (5) => " & mComb.Contiene(5)
    Debug.Print "Contiene el 9? " & mComb.Contiene(9)
    Debug.Print "EstaOrdenado: " & mComb.EstaOrdenado
End Sub



'------------------------------------------------------------------------------*
' Procedimiento  : Print_Numero
' Fecha          : 29/Abr/2018
' Propósito      : Visualiza las propiedades y metodos de la clase Numero
' Parámetros     : Numero
'------------------------------------------------------------------------------*
'
Private Sub PrintNumero(obj As Numero)
    Debug.Print "==> Pruebas Numero"
    Debug.Print vbTab & "Decena       =" & obj.Decena
    Debug.Print vbTab & "EsPar        =" & obj.EsPar
    Debug.Print vbTab & "Orden        =" & obj.Orden
    Debug.Print vbTab & "Paridad      =" & obj.Paridad
    Debug.Print vbTab & "Peso         =" & obj.Peso
    Debug.Print vbTab & "Septena      =" & obj.Septena
    Debug.Print vbTab & "Terminacion  =" & obj.Terminacion
    Debug.Print vbTab & "Valor        =" & obj.Valor
    Debug.Print vbTab & "EsValido     =" & obj.EsValido(JUEGO_DEFECTO)
    Debug.Print vbTab & "GetMensaje   =" & obj.GetMensaje()
    Debug.Print vbTab & "ToString()   =" & obj.ToString()
End Sub




'------------------------------------------------------------------------------*
' Procedimiento  : Print_Combinacion
' Fecha          : 29/Abr/2018
' Propósito      : Visualiza las propiedades y metodos de la clase Combinacion
' Parámetros     : Combinacion
'------------------------------------------------------------------------------*
'
Private Sub PrintCombinacion(obj As Combinacion)
    Debug.Print "==> Pruebas Combinacion"
    Debug.Print vbTab & "Add()               =" & "#Metodo" 'obj.Add
    Debug.Print vbTab & "Clear()             =" & "#Metodo" 'obj.Clear
    Debug.Print vbTab & "Contiene(N)         =" & obj.Contiene(5)
    Debug.Print vbTab & "Count               =" & obj.Count
    Debug.Print vbTab & "Delete()            =" & "#Metodo" 'obj.Delete
    Debug.Print vbTab & "EstaOrdenado()      =" & obj.EstaOrdenado
    Debug.Print vbTab & "Es Valido()         =" & obj.EsValido
    Debug.Print vbTab & "FormulaAltoBajo     =" & obj.FormulaAltoBajo
    Debug.Print vbTab & "FormulaConsecutivos =" & obj.FormulaConsecutivos
    Debug.Print vbTab & "FormulaDecenas      =" & obj.FormulaDecenas
    Debug.Print vbTab & "FormulaParidad      =" & obj.FormulaParidad
    Debug.Print vbTab & "FormulaSeptenas     =" & obj.FormulaSeptenas
    Debug.Print vbTab & "FormulaTerminaciones=" & obj.FormulaTerminaciones
    Debug.Print vbTab & "GetArray()          =" & "#Array"  'obj.GetArray()
    Debug.Print vbTab & "GetMensaje()        =" & obj.GetMensaje
    Debug.Print vbTab & "Numeros             =" & "#Col"    'obj.Numeros
    Debug.Print vbTab & "Producto            =" & obj.Producto
    Debug.Print vbTab & "Suma                =" & obj.Suma
    Debug.Print vbTab & "Texto               =" & obj.Texto
    Debug.Print vbTab & "TextoOrdenado       =" & obj.ToString(True)
End Sub
' *===========(EOF): Lot_PqtNucleoTesting.bas
