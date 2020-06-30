Attribute VB_Name = "Lot_PqtConcursosTesting"
' *============================================================================*
' *
' *     Fichero    : Lot_PqtConcursosTesting.bas
' *
' *     Autor      : Carlos Almela Baeza
' *     Creación   : Dom, 29/Abr/2018 08:42:00
' *     Versión    : 1.0
' *     Propósito  : Colección de pruebas unitarias de las clases del paquete
' *                  Concurso:
' *                    - Numero
' *                    - Combinacion
' *                    - Sorteo
' *                    - InfoSorteo
' *                    - Tarifa
' *                    - Sorteos
' *                    - SorteoEngine
' *                    - Premio
' *                    - Premios
' *
' *============================================================================*
Option Explicit
Option Base 0

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
    '   Audita el error
    Call HandleException(ErrNumber, ErrDescription, "Lot_PqtConcursosTesting.NumeroTest", ErrSource)
    '   Informa del error
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
    '   Combinación no Valida
    '
    Set obj = New Combinacion
      
  On Error GoTo 0
CombinacionTest__CleanExit:
    Exit Sub
            
CombinacionTest_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    '   Audita el error
    Call HandleException(ErrNumber, ErrDescription, "Lot_PqtConcursosTesting.CombinacionTest", ErrSource)
    '   Informa del error
    Call MsgBox(ErrDescription, vbCritical Or vbSystemModal, Application)
    Call Trace("CERRAR")
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
    Debug.Print vbTab & "Error        =" & obj.Error
    Debug.Print vbTab & "EsPar        =" & obj.EsPar
    Debug.Print vbTab & "EsValido     =" & obj.EsValido(JUEGO_DEFECTO)
    Debug.Print vbTab & "GetMensaje   =" & obj.GetMensaje()
    Debug.Print vbTab & "Orden        =" & obj.Orden
    Debug.Print vbTab & "Paridad      =" & obj.Paridad
    Debug.Print vbTab & "Peso         =" & obj.Peso
    Debug.Print vbTab & "Septena      =" & obj.Septena
    Debug.Print vbTab & "Terminacion  =" & obj.Terminacion
    Debug.Print vbTab & "Valor        =" & obj.Valor
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
    Debug.Print vbTab & "Add                 =" & "#Metodo" 'obj.Add
    Debug.Print vbTab & "Clear               =" & "#Metodo" 'obj.Clear
    Debug.Print vbTab & "Contiene            =" & obj.Contiene(5)
    Debug.Print vbTab & "Count               =" & obj.Count
    Debug.Print vbTab & "Delete              =" & "#Metodo" 'obj.Delete
    Debug.Print vbTab & "EstaOrdenado        =" & obj.EstaOrdenado
    Debug.Print vbTab & "FormulaAltoBajo     =" & obj.FormulaAltoBajo
    Debug.Print vbTab & "FormulaConsecutivos =" & obj.FormulaConsecutivos
    Debug.Print vbTab & "FormulaDecenas      =" & obj.FormulaDecenas
    Debug.Print vbTab & "FormulaParidad      =" & obj.FormulaParidad
    Debug.Print vbTab & "FormulaSeptenas     =" & obj.FormulaSeptenas
    Debug.Print vbTab & "FormulaTerminaciones=" & obj.FormulaTerminaciones
    Debug.Print vbTab & "Numeros             =" & "#Col"  'obj.Numeros
    Debug.Print vbTab & "Producto            =" & obj.Producto
    Debug.Print vbTab & "Suma                =" & obj.Suma
    Debug.Print vbTab & "Texto               =" & obj.Texto
    Debug.Print vbTab & "TextoOrdenado       =" & obj.TextoOrdenado
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
' Procedure : PrintCombinacion
' Author    : CHARLY
' Date      : sáb, 01/nov/2014 21:02:09
' Purpose   :
'---------------------------------------------------------------------------------------
'
'Private Sub PrintCombinacion(Obj As Combinacion)
'    Debug.Print "==> Combinacion"
'    Debug.Print " .Count               = " & Obj.Count
'    Debug.Print " .FormulaAltoBajo     = " & Obj.FormulaAltoBajo
'    Debug.Print " .FormulaConsecutivos = " & Obj.FormulaConsecutivos
'    Debug.Print " .FormulaDecenas      = " & Obj.FormulaDecenas
'    Debug.Print " .FormulaParidad      = " & Obj.FormulaParidad
'    Debug.Print " .FormulaSeptenas     = " & Obj.FormulaSeptenas
'    Debug.Print " .FormulaTerminacion  = " & Obj.FormulaTerminaciones
'    Debug.Print " .Producto            = " & Obj.Producto
'    Debug.Print " .Suma                = " & Obj.Suma
'    Debug.Print " .Texto               = " & Obj.Texto
'    Debug.Print " .TextoOrdenado       = " & Obj.TextoOrdenado
'    Debug.Print " .EstaOrdenado()      = " & Obj.EstaOrdenado
'End Sub

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
Private Sub InfoSorteoTest()
    Dim mInfo As InfoSorteo
    Dim i As Integer
    Dim mFechaI As Date
    Dim mFechaF As Date
    Dim mDias   As Integer
    
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
    '
    '   Add dias a un sorteo
    '
    mFechaI = #4/26/2015#   'Domingo
    mDias = 20
    mFechaF = mFechaI
    mInfo.Constructor Bonoloto
    For i = 1 To 7
        Debug.Print "Calculo de sumar " & CStr(mDias) & " sorteos a la fecha " & mFechaF
        Debug.Print "   ==>" & mInfo.AddDiasSorteo(mFechaF, mDias)
        mFechaF = mFechaF + 1
    Next i
    mDias = 7
    mFechaF = mFechaI
    For i = 1 To 7
        Debug.Print "Calculo de sumar " & CStr(mDias) & " sorteos a la fecha " & mFechaF
        Debug.Print "   ==>" & mInfo.AddDiasSorteo(mFechaF, mDias)
        mFechaF = mFechaF + 1
    Next i
    
    mDias = 3
    mFechaF = mFechaI
    For i = 1 To 7
        Debug.Print "Calculo de sumar " & CStr(mDias) & " sorteos a la fecha " & mFechaF
        Debug.Print "   ==>" & mInfo.AddDiasSorteo(mFechaF, mDias)
        mFechaF = mFechaF + 1
    Next i

End Sub

' *===========(EOF): Lot_PqtConcursosTesting.bas
