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
' *                    - InfoSorteo
' *                    - Tarifa (pendiente)
' *                    - Sorteo
' *                    - Sorteos
' *                    - SorteoEngine
' *                    - SorteoModel
' *                    - Premio
' *                    - Premios
' *
' *============================================================================*
Option Explicit
Option Base 0

'------------------------------------------------------------------------------*
' Procedimiento  : NucleoTest
' Fecha          : sá., 16/mar/2019 13:30:45
' Propósito      : Pruebas Unitarias de las clases del paquete
'------------------------------------------------------------------------------*
'
Public Sub PqtConcursosTest()
    NumeroTest
    CombinacionTest
    InfoSorteoTest
    SorteoTest
    SorteosTest
    SorteoEngineTest
    PremioTest
    PremiosTest
End Sub
'------------------------------------------------------------------------------*
' Procedimiento  : NumeroTest
' Fecha          : 29/Abr/2018
' Propósito      : Pruebas Unitarias de la clase Numero
'------------------------------------------------------------------------------*
'
Private Sub NumeroTest()
    Dim Obj As Numero
  
  On Error GoTo NumeroTest_Error
    Set Obj = New Numero
    PrintNumero Obj, LP_LB_6_49
    '
    '   Numero Válido
    '
    Obj.Valor = 5
    PrintNumero Obj, LP_LB_6_49
    '
    '   Numero no Valido para primitiva
    '
    Set Obj = New Numero
    Obj.Valor = 50
    PrintNumero Obj, LP_LB_6_49
    '
    '   Numero no Valido para Euromillon
    '
    Set Obj = New Numero
    Obj.Valor = 53
    PrintNumero Obj, EU_5_50
    '
    '   Numero no Valido para Gordo
    '
    Set Obj = New Numero
    Obj.Valor = 55
    PrintNumero Obj, GP_5_54
    '
    '   Estrella válida para euromillones
    '
    Set Obj = New Numero
    Obj.Valor = 5
    PrintNumero Obj, EU_2_12
    
    '
    '   Estrella NO válida para euromillones
    '
    Set Obj = New Numero
    Obj.Valor = 15
    PrintNumero Obj, EU_2_12
      
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
' Procedimiento  : PrintNumero
' Fecha          : 29/Abr/2018
' Propósito      : Visualiza las propiedades y metodos de la clase Numero
' Parámetros     : Numero
'------------------------------------------------------------------------------*
'
Private Sub PrintNumero(Obj As Numero, datTipoJuego As ModalidadJuego)
    Debug.Print "==> Pruebas Numero"
    Debug.Print vbTab & "Decena       =" & Obj.Decena
    Debug.Print vbTab & "EsPar        =" & Obj.EsPar
    Debug.Print vbTab & "EsValido     =" & Obj.EsValido(datTipoJuego)
    Debug.Print vbTab & "GetMensaje   =" & Obj.GetMensaje()
    Debug.Print vbTab & "ToString     =" & Obj.ToString()
    Debug.Print vbTab & "Orden        =" & Obj.Orden
    Debug.Print vbTab & "Paridad      =" & Obj.Paridad
    Debug.Print vbTab & "Peso         =" & Obj.Peso
    Debug.Print vbTab & "Septena      =" & Obj.Septena
    Debug.Print vbTab & "Terminacion  =" & Obj.Terminacion
    Debug.Print vbTab & "Valor        =" & Obj.Valor
End Sub

'------------------------------------------------------------------------------*
' Procedimiento  : CombinacionTest
' Fecha          : 29/Abr/2018
' Propósito      : Pruebas Unitarias de la clase Combinacion
'------------------------------------------------------------------------------*
'
Private Sub CombinacionTest()
    Dim Obj     As Combinacion
    Dim oNum    As Numero
  On Error GoTo CombinacionTest_Error
    '
    '   Combinacion Vacia
    '
    Set Obj = New Combinacion
    PrintCombinacion Obj
    '
    '   Combinacion Valida  28-1-31-25-33-8-7
    '
    Set oNum = New Numero
    oNum.Valor = 28
    Obj.Add oNum
    Set oNum = New Numero
    oNum.Valor = 1
    Obj.Add oNum
    Set oNum = New Numero
    oNum.Valor = 31
    Obj.Add oNum
    Set oNum = New Numero
    oNum.Valor = 25
    Obj.Add oNum
    Set oNum = New Numero
    oNum.Valor = 33
    Obj.Add oNum
    Set oNum = New Numero
    oNum.Valor = 8
    Obj.Add oNum
    Set oNum = New Numero
    oNum.Valor = 7
    Obj.Add oNum
    '
    '
    PrintCombinacion Obj
    '
    '   Combinación no Valida para euromillones
    '   28-1-56-25-33-8-7
    '
    Set Obj = New Combinacion
    Set oNum = New Numero
    oNum.Valor = 28
    Obj.Add oNum
    Set oNum = New Numero
    oNum.Valor = 1
    Obj.Add oNum
    Set oNum = New Numero
    oNum.Valor = 56
    Obj.Add oNum
    Set oNum = New Numero
    oNum.Valor = 25
    Obj.Add oNum
    Set oNum = New Numero
    oNum.Valor = 33
    Obj.Add oNum
    Set oNum = New Numero
    oNum.Valor = 8
    Obj.Add oNum
    Set oNum = New Numero
    oNum.Valor = 7
    Obj.Add oNum
    '
    '
    PrintCombinacion Obj
    '
    '   Combinación de estrellas
    '       2 - 7
    '
    Set Obj = New Combinacion
    Set oNum = New Numero
    oNum.Valor = 2
    Obj.Add oNum
    Set oNum = New Numero
    oNum.Valor = 7
    Obj.Add oNum
    '
    '
    PrintCombinacion Obj
      
      
      
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



'------------------------------------------------------------------------------*
' Procedimiento  : PrintCombinacion
' Fecha          : 29/Abr/2018
' Propósito      : Visualiza las propiedades y metodos de la clase Combinacion
' Parámetros     : Combinacion
'------------------------------------------------------------------------------*
'
Private Sub PrintCombinacion(Obj As Combinacion)
    Debug.Print "==> Pruebas Combinacion"
    Debug.Print vbTab & "Add()               =" & "#Metodo" 'obj.Add
    Debug.Print vbTab & "Clear()             =" & "#Metodo" 'obj.Clear
    Debug.Print vbTab & "Contiene(5)         =" & Obj.Contiene(5)
    Debug.Print vbTab & "Count               =" & Obj.Count
    Debug.Print vbTab & "Delete()            =" & "#Metodo" 'obj.Delete
    Debug.Print vbTab & "EstaOrdenado        =" & Obj.EstaOrdenado
    Debug.Print vbTab & "EsValido(bonoloto)  =" & Obj.EsValido(Bonoloto)
    Debug.Print vbTab & "FormulaAltoBajo     =" & Obj.FormulaAltoBajo
    Debug.Print vbTab & "FormulaConsecutivos =" & Obj.FormulaConsecutivos
    Debug.Print vbTab & "FormulaDecenas      =" & Obj.FormulaDecenas
    Debug.Print vbTab & "FormulaParidad      =" & Obj.FormulaParidad
    Debug.Print vbTab & "FormulaSeptenas     =" & Obj.FormulaSeptenas
    Debug.Print vbTab & "FormulaTerminaciones=" & Obj.FormulaTerminaciones
    Debug.Print vbTab & "GetMensaje()        =" & Obj.GetMensaje
    Debug.Print vbTab & "Numeros             =" & "#Col"  'obj.Numeros
    Debug.Print vbTab & "Producto            =" & Obj.Producto
    Debug.Print vbTab & "Suma                =" & Obj.Suma
    Debug.Print vbTab & "Texto               =" & Obj.Texto
    Debug.Print vbTab & "ToString(True)      =" & Obj.ToString(True)
End Sub
'------------------------------------------------------------------------------*
' Procedimiento  : SorteoTest
' Fecha          : 17/Jun/2018
' Propósito      : Pruebas unitarias de la clase Sorteo
'------------------------------------------------------------------------------*
Public Sub SorteoTest()
    Dim mObj        As Sorteo
    Dim mComb       As Combinacion
    
 On Error GoTo SorteoTest_Error
    '
    '   1.- Objeto en Vacio
    '
    Set mObj = New Sorteo
    PrintSorteo mObj
    '
    '   2.- Sorteo Bonoloto
    '
    Set mObj = New Sorteo
    With mObj
        .Juego = Bonoloto
        .Dia = "M"
        .Texto = "10-49-15-31-17-7"
        .Complementario = 34
        .Ordenado = True
        .Fecha = #5/15/2018#
        .Id = 4512
        .NumSorteo = "2018/116"
        .Reintegro = 6
    End With
    PrintSorteo mObj
    '
    '   3.- Sorteo Euromillon
    '
    ' 1076    2018/056    13/07/2018  V   28  Si  49  14  4   1   21  2   12
    '
    Set mObj = New Sorteo
    With mObj
        .Juego = Euromillones
        .Dia = "V"
        .Texto = "49-14-4-1-21"
        .Estrellas.Texto = "2-12"
        .Ordenado = True
        .Fecha = #7/13/2018#
        .Id = 1076
        .NumSorteo = "2018/056"
    End With
    PrintSorteo mObj
    '
    '   4.- Sorteo Gordo
    '
    '  1074    2018/022    03/06/2018  D   22  Si  28  38  44  33  5  C 4
    '
    Set mObj = New Sorteo
    With mObj
        .Juego = gordoPrimitiva
        .Dia = "D"
        .Texto = "28-38-44-33-5"
        .Clave = 4
        .Ordenado = True
        .Fecha = #6/3/2018#
        .Id = 1074
        .NumSorteo = "2018/022"
    End With
    PrintSorteo mObj
    '
    '   5.- Sorteo Primitiva
    '
    ' 3132    2018/052    30/06/2018  S   26  Si  44  30  37  16  14  5   23  3
    Set mObj = New Sorteo
    With mObj
        .Juego = LoteriaPrimitiva
        .Dia = "S"
        .Texto = "44-30-37-16-14-5"
        .Complementario = 23
        .Ordenado = True
        .Fecha = #6/30/2018#
        .Id = 3132
        .NumSorteo = "2018/052"
        .Reintegro = 3
    End With
    PrintSorteo mObj
        
 On Error GoTo 0
SorteoTest__CleanExit:
    Exit Sub
            
SorteoTest_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    '   Audita el error
    Call HandleException(ErrNumber, ErrDescription, "Lot_PqtConcursosTesting.SorteoTest", ErrSource)
    '   Informa del error
    Call MsgBox(ErrDescription, vbCritical Or vbSystemModal, ThisWorkbook.Name)
    Call Trace("CERRAR")
End Sub
'------------------------------------------------------------------------------*
' Procedimiento  : PrintSorteo
' Fecha          : 17/Jun/2018
' Propósito      : Visualizar los atributos de la clase sorteo
'------------------------------------------------------------------------------*
Private Sub PrintSorteo(mObj As Sorteo)
    Debug.Print "==> Pruebas Sorteo"
    Debug.Print vbTab & "CombinacionGanadora =" & "#Col " & mObj.CombinacionGanadora.Texto
    Debug.Print vbTab & "Complementario      =" & mObj.Complementario
    Debug.Print vbTab & "Constructor         =" & "#Metodo " 'mObj.Constructor
    Debug.Print vbTab & "Dia                 =" & mObj.Dia
    Debug.Print vbTab & "EntidadNegocio      =" & "#Objeto " & mObj.EntidadNegocio.Id
    Debug.Print vbTab & "EstrellaDos         =" & mObj.EstrellaDos.Valor
    Debug.Print vbTab & "Estrellas           =" & mObj.Estrellas.Texto
    Debug.Print vbTab & "EstrellaUno         =" & mObj.EstrellaUno.Valor
    Debug.Print vbTab & "EsValido            =" & mObj.EsValido
    Debug.Print vbTab & "Existe              =" & "#Metodo " 'mObj.Existe
    Debug.Print vbTab & "Fecha               =" & mObj.Fecha
    Debug.Print vbTab & "GetMensaje          =" & mObj.GetMensaje
    Debug.Print vbTab & "Juego               =" & mObj.Juego
    Debug.Print vbTab & "Ordenado            =" & mObj.Ordenado
    Debug.Print vbTab & "Premios             =" & "#Col " 'mObj.Premios
    Debug.Print vbTab & "ID                  =" & mObj.Id
    Debug.Print vbTab & "Reintegro           =" & mObj.Reintegro
    Debug.Print vbTab & "Semana              =" & mObj.Semana
    Debug.Print vbTab & "Texto               =" & mObj.Texto
    Debug.Print vbTab & "ToString            =" & mObj.ToString
End Sub

'---------------------------------------------------------------------------------------
' Procedure : PremioTest
' Author    : CHARLY
' Date      : sáb, 01/nov/2014 21:01:42
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub PremioTest()
    Dim Obj   As Premio
    Dim Obj1  As Premio
    Dim mPremioString As String
    
 On Error GoTo PremioTest_Error
    '
    '   Definición de Objetos
    '
    Set Obj = New Premio
    '
    '   Objeto por defecto
    '
    PrintPremio Obj
    '
    '   Asignación de propiedades
    '
    With Obj
        .Categoria = Cuarta
        .Importe = 125
        .Juego = Bonoloto
        .NumeroAcertantesEspaña = 20
        .NumeroAcertantesEuropa = 15
        .ImporteDefault = False
    End With
    PrintPremio Obj
    
    '
    '   Prueba de Parser
    '
    Set Obj1 = New Premio
    mPremioString = Obj.ToString
    '
    '   Asignamos al objeto nuevo
    '
    Obj1.Parse mPremioString
    PrintPremio Obj1
 On Error GoTo 0
PremioTest__CleanExit:
    Exit Sub
            
PremioTest_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    Call HandleException(ErrNumber, ErrDescription, "Lot_PqtConcursosTesting.PremioTest", ErrSource)
    Call MsgBox(ErrDescription, vbCritical Or vbSystemModal, ThisWorkbook.Name)
    Call Trace("CERRAR")
End Sub
'---------------------------------------------------------------------------------------
' Procedure : PrintPremio
' Author    : CHARLY
' Date      : 06/04/2014
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub PrintPremio(Obj As Premio)
    Debug.Print " Premio ]================="
    Debug.Print "   CategoriaPremio        =>" & Obj.Categoria
    Debug.Print "   CategoriaTexto         =>" & Obj.CategoriaTexto
    Debug.Print "   Importe                =>" & Obj.Importe
    Debug.Print "   ImporteDefault         =>" & Obj.ImporteDefault
    Debug.Print "   Juego                  =>" & Obj.Juego
    Debug.Print "   Acertantes en Europa   =>" & Obj.NumeroAcertantesEuropa
    Debug.Print "   Acertantes en España   =>" & Obj.NumeroAcertantesEspaña
    Debug.Print "   GetImportePremio(5)    =>" & Obj.GetImportePremio(5)
    Debug.Print "   EsValido()             =>" & Obj.EsValido()
    Debug.Print "   ToString()             =>" & Obj.ToString()
End Sub

'------------------------------------------------------------------------------*
' Procedimiento  : PremiosTest
' Fecha          : 10/feb/2019 00:29:24
' Propósito      : Pruebas unitarias de la clase Premios
'------------------------------------------------------------------------------*
Private Sub PremiosTest()
    Dim mObj    As Premios
    Dim mPrem   As Premio
    
 On Error GoTo PremiosTest_Error
    '
    '   1.- Objeto en Vacio
    '
    Set mObj = New Premios
    PrintPremios mObj
    '
    '   2.- Prueba método SetPremiosDefecto
    '
    mObj.SetPremiosDefecto Bonoloto
    PrintPremios mObj
    '
    '   Euromillon
    mObj.SetPremiosDefecto Euromillones
    PrintPremios mObj
    '
    '   3.- Prueba Clear
    '
    mObj.Clear
    '
    '   4.- Prueba Add
    '
    Set mPrem = New Premio
    With mPrem
        .Juego = LoteriaPrimitiva
        .Categoria = Quinta
        .Importe = 36
    End With
    mObj.Add mPrem
    Set mPrem = New Premio
    With mPrem
        .Juego = LoteriaPrimitiva
        .Categoria = Cuarta
        .Importe = 256.45
    End With
    mObj.Add mPrem
    PrintPremios mObj
    '
    '   Prueba Get Importe
    '
    mObj.Juego = Bonoloto
    mObj.NumerosAcertados = 3
    mObj.PronosticosApostados = 9
    mObj.ReintegroAcertado = True
    PrintPremios mObj
    '
    '  5.- Parse
    '
    '   Definir una cadena con los juegos de bonoloto y cargar premios
    '
    '   7 apuestas para el juego bonoloto y 6 aciertos
    '
    Set mObj = New Premios
    mObj.SetPremiosDefecto Bonoloto
    mObj.NumerosAcertados = 6
    mObj.PronosticosApostados = 7
    mObj.ComplementarioAcertado = False
    mObj.ReintegroAcertado = False
    Debug.Print "GetImporteTotalPremios (456000): " & mObj.GetImporteTotalPremios()
    '
    '   7 apuestas para el juego Gordo y 6 aciertos
    '
    Set mObj = New Premios
    mObj.SetPremiosDefecto Bonoloto
    mObj.NumerosAcertados = 6
    mObj.PronosticosApostados = 7
    mObj.ComplementarioAcertado = False
    mObj.ReintegroAcertado = False
    Debug.Print "GetImporteTotalPremios (456000): " & mObj.GetImporteTotalPremios()
    '
    '   5 apuestas para el juego Euromillon y 6 aciertos
    '
    Set mObj = New Premios
    mObj.SetPremiosDefecto Bonoloto
    mObj.NumerosAcertados = 6
    mObj.PronosticosApostados = 7
    mObj.ComplementarioAcertado = False
    mObj.ReintegroAcertado = False
    Debug.Print "GetImporteTotalPremios (456000): " & mObj.GetImporteTotalPremios()
    '
    '
    '
    Err.Raise ERR_TODO, "Lot_PqtConcursosTesting.PremiosTest", MSG_TODO
    '
 On Error GoTo 0
PremiosTest__CleanExit:
    Exit Sub
            
PremiosTest_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    Call HandleException(ErrNumber, ErrDescription, "Lot_PqtConcursosTesting.PremiosTest", ErrSource)
    Call MsgBox(ErrDescription, vbCritical Or vbSystemModal, ThisWorkbook.Name)
    Call Trace("CERRAR")
End Sub

'------------------------------------------------------------------------------*
' Procedimiento  : PrintPremios
' Fecha          : do., 10/feb/2019 19:55:16
' Propósito      : Imprimir las propiedades y atributos de Premios
'------------------------------------------------------------------------------*
Private Sub PrintPremios(mObj As Premios)
    Dim mPrem As Premio
    Debug.Print "==> Pruebas Premios"
    Debug.Print vbTab & "ComplementarioAcertado =" & mObj.ComplementarioAcertado
    Debug.Print vbTab & "Count                  =" & mObj.Count
    Debug.Print vbTab & "EstrellasAcertadas     =" & mObj.EstrellasAcertadas
    Debug.Print vbTab & "EstrellasApostadas     =" & mObj.EstrellasApostadas
    Debug.Print vbTab & "IdSorteo               =" & mObj.IdSorteo
    Debug.Print vbTab & "Juego                  =" & mObj.Juego
    Debug.Print vbTab & "NumerosAcertados       =" & mObj.NumerosAcertados
    Debug.Print vbTab & "PronosticosApostados   =" & mObj.PronosticosApostados
    Debug.Print vbTab & "ReintegroAcertado      =" & mObj.ReintegroAcertado
    Debug.Print vbTab & "Add()                  =" & "#Metodo" ' mObj.Add
    Debug.Print vbTab & "Clear()                =" & "#Metodo" ' mObj.Clear
    Debug.Print vbTab & "Delete()               =" & "#Metodo" ' mObj.Delete
    Debug.Print vbTab & "GetImporteTotalPremios()=" & mObj.GetImporteTotalPremios
    Debug.Print vbTab & "Items                  =" & mObj.Items.Count
    Debug.Print vbTab & "MarkForDelete()        =" & "#Metodo" ' mObj.MarkForDelete
    Debug.Print vbTab & "MarkSetPremiosDefecto()=" & "#Metodo" ' mObj.MarkSetPremiosDefecto
    Debug.Print vbTab & "Undelete()             =" & "#Metodo" ' mObj.Undelete
    Debug.Print vbTab & "ToString()             =" & mObj.ToString
    For Each mPrem In mObj.Items
        Debug.Print vbTab & mPrem.ToString
    Next mPrem
End Sub

'------------------------------------------------------------------------------*
' Procedimiento  : InfoSorteoTest
' Fecha          : 10/feb/2019 00:29:24
' Propósito      : Pruebas unitarias de la clase Premios
'------------------------------------------------------------------------------*
Private Sub InfoSorteoTest()
    Dim mInfo As InfoSorteo
    Dim i As Integer
    Dim mFechaI As Date
    Dim mFechaF As Date
 
 On Error GoTo InfoSorteoTest_Error


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
        Debug.Print "EsFechaSorteo (" & mFechaI & ", GordoPrimitiva) => " & mInfo.EsFechaSorteo(mFechaI)
        
        mInfo.Constructor LoteriaPrimitiva
        mFechaF = mInfo.GetProximoSorteo(mFechaI)
        Debug.Print "GetProximoSorteo(" & mFechaI & ", LoteriaPrimitiva) => "; mFechaF
        Debug.Print "EsFechaSorteo (" & mFechaI & ", LoteriaPrimitiva) => " & mInfo.EsFechaSorteo(mFechaI)
        
        mInfo.Constructor Euromillones
        mFechaF = mInfo.GetProximoSorteo(mFechaI)
        Debug.Print "GetProximoSorteo(" & mFechaI & ", Euromillones) => "; mFechaF
        Debug.Print "EsFechaSorteo (" & mFechaI & ", Euromillones) => " & mInfo.EsFechaSorteo(mFechaI)
    Next i
    '
    '   Sorteos entre dos fechas
    '
    mFechaI = #4/26/2015#   'Domingo
    mFechaF = mFechaI
    For i = 1 To 26
        Debug.Print "Sorteos entre " & mFechaI & " y " & mFechaF
        Debug.Print "   ==>" & mInfo.GetSorteosEntreFechas(mFechaI, mFechaF)
        mFechaF = mFechaF + 1
    Next i
    
 On Error GoTo 0
InfoSorteoTest__CleanExit:
    Exit Sub
            
InfoSorteoTest_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    Call HandleException(ErrNumber, ErrDescription, "Lot_PqtConcursosTesting.InfoSorteoTest", ErrSource)
    Call MsgBox(ErrDescription, vbCritical Or vbSystemModal, ThisWorkbook.Name)
    Call Trace("CERRAR")
End Sub
'------------------------------------------------------------------------------*
' Procedimiento  : SorteosTest
' Fecha          : 17/Jun/2018
' Propósito      : Pruebas unitarias de la clase Sorteos
'------------------------------------------------------------------------------*
Public Sub SorteosTest()
    Dim mObj        As Sorteos
    Dim mSort       As Sorteo
    
 On Error GoTo SorteosTest_Error
    '
    '   1.- Objeto en Vacio
    '
    Set mObj = New Sorteos
    PrintSorteos mObj
    '
    '   2.- Sorteos de  Bonoloto
    '
    Set mSort = New Sorteo
    With mSort
        .Juego = Bonoloto
        .Dia = "M"
        .Texto = "10-49-15-31-17-7"
        .Complementario = 34
        .Ordenado = True
        .Fecha = #5/15/2018#
        .Id = 4512
        .NumSorteo = "2018/116"
        .Reintegro = 6
    End With
    mObj.Add mSort
    '
    '   Segundo Sorteo
    '
    Set mSort = New Sorteo
    With mSort
        .Juego = Bonoloto
        .Dia = "X"
        .Texto = "38-31-45-5-48-13"
        .Complementario = 44
        .Ordenado = True
        .Fecha = #5/16/2018#
        .Id = 6631
        .NumSorteo = "2018/117"
        .Reintegro = 9
    End With
    mObj.Add mSort
    PrintSorteos mObj
    '
    '   3.- Sorteo Euromillon
    '
    Set mObj = New Sorteos
    ' primero item
    Set mSort = New Sorteo
    With mSort
        .Juego = Euromillones
        .Dia = "V"
        .Texto = "5-31-18-21-35"
        .Estrellas.Texto = "6-9"
        .Ordenado = True
        .Fecha = #10/26/2018#
        .Id = 1106
        .NumSorteo = "2018/086"
    End With
    mObj.Add mSort
    ' segundo item
    Set mSort = New Sorteo
    With mSort
        .Juego = Euromillones
        .Dia = "M"
        .Texto = "44-27-23-17-43"
        .Estrellas.Texto = "1-12"
        .Ordenado = True
        .Fecha = #10/30/2018#
        .Id = 1107
        .NumSorteo = "2018/087"
    End With
    mObj.Add mSort
    ' tercer item
    Set mSort = New Sorteo
    With mSort
        .Juego = Euromillones
        .Dia = "V"
        .Texto = "15-37-5-17-44"
        .Estrellas.Texto = "11-7"
        .Ordenado = True
        .Fecha = #11/2/2018#
        .Id = 1108
        .NumSorteo = "2018/088"
    End With
    mObj.Add mSort
    PrintSorteos mObj
    '
    '   4.- Prueba propiedad Count
    '
    Debug.Print "=> Propiedad Count (3) => " & mObj.Count
    '
    '   5.- Prueba método Items
    '
    Debug.Print "=> Prueba Items"
    For Each mSort In mObj.Items
        Debug.Print vbTab & "* (" & mSort.Id & ") Sorteo=>" & mSort.ToString
    Next mSort
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
    PrintSorteos mObj
 
 On Error GoTo 0
SorteosTest__CleanExit:
    Exit Sub
            
SorteosTest_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    '   Audita el error
    Call HandleException(ErrNumber, ErrDescription, "Lot_PqtConcursosTesting.SorteosTest", ErrSource)
    '   Informa del error
    Call MsgBox(ErrDescription, vbCritical Or vbSystemModal, ThisWorkbook.Name)
    Call Trace("CERRAR")
End Sub

'------------------------------------------------------------------------------*
' Procedimiento  : SorteoEngineTest
' Fecha          : 17/Jun/2018
' Propósito      : Pruebas unitarias de la clase SorteoEngine
'------------------------------------------------------------------------------*
Public Sub SorteoEngineTest()
    Dim mEng        As SorteoEngine
    Dim mCol        As Sorteos
    Dim mObj        As Sorteo
    Dim mDate       As Date
    Dim mDateI      As Date
    Dim mDateF      As Date
    Dim mInt        As Integer
    Dim mIntI       As Integer
    Dim mIntF       As Integer
    Dim mVar        As Variant
    Dim i           As Integer
    
 On Error GoTo SorteoEngineTest_Error
    '
    '   Cabecera de pruebas
    '
    Debug.Print "Pruebas del Motor de Sorteos ============="
    '
    '   Creamos el motor que proporciona sorteos
    '
    Set mEng = New SorteoEngine
    '
    '
    mEng.Juego = Bonoloto
    Debug.Print vbTab & "mEng.Juego = " & mEng.Juego
    '
    '   1.- Prueba GetFechaPrimerSorteo
    '
    mDate = mEng.GetFechaPrimerSorteo(Bonoloto)
    Debug.Print "Primera fecha sorteo Bonoloto= " & Format(mDate, "dd/mmm/yyyy")
    mDate = mEng.GetFechaPrimerSorteo(LoteriaPrimitiva)
    Debug.Print "Primera fecha sorteo Primitiva= " & Format(mDate, "dd/mmm/yyyy")
    mDate = mEng.GetFechaPrimerSorteo(Euromillones)
    Debug.Print "Primera fecha sorteo Euromillones= " & Format(mDate, "dd/mmm/yyyy")
    mDate = mEng.GetFechaPrimerSorteo(gordoPrimitiva)
    Debug.Print "Primera fecha sorteo Gordo = " & Format(mDate, "dd/mmm/yyyy")
    '
    '   2.- Prueba GetFechaUltimoSorteo
    '
    mDate = mEng.GetFechaUltimoSorteo(Bonoloto)
    Debug.Print "Ultima fecha sorteo Bonoloto= " & Format(mDate, "dd/mmm/yyyy")
    mDate = mEng.GetFechaUltimoSorteo(LoteriaPrimitiva)
    Debug.Print "Ultima fecha sorteo Primitiva= " & Format(mDate, "dd/mmm/yyyy")
    mDate = mEng.GetFechaUltimoSorteo(Euromillones)
    Debug.Print "Ultima fecha sorteo Euromillones= " & Format(mDate, "dd/mmm/yyyy")
    mDate = mEng.GetFechaUltimoSorteo(gordoPrimitiva)
    Debug.Print "Ultima fecha sorteo Gordo = " & Format(mDate, "dd/mmm/yyyy")
    '
    '   3.- Prueba GetIdPrimerSorteo
    '
    mInt = mEng.GetIdPrimerSorteo(Bonoloto)
    Debug.Print "Primer Id Bonoloto= " & CStr(mInt)
    mInt = mEng.GetIdPrimerSorteo(LoteriaPrimitiva)
    Debug.Print "Primer Id Primitiva= " & CStr(mInt)
    mInt = mEng.GetIdPrimerSorteo(Euromillones)
    Debug.Print "Primer Id Euromillones= " & CStr(mInt)
    mInt = mEng.GetIdPrimerSorteo(gordoPrimitiva)
    Debug.Print "Primer Id Gordo= " & CStr(mInt)
    '
    '   4.- Prueba GetIdUltimoSorteo
    '
    mInt = mEng.GetIdUltimoSorteo(Bonoloto)
    Debug.Print "Ultimo Id Bonoloto= " & CStr(mInt)
    mInt = mEng.GetIdUltimoSorteo(LoteriaPrimitiva)
    Debug.Print "Ultimo Id Primitiva= " & CStr(mInt)
    mInt = mEng.GetIdUltimoSorteo(Euromillones)
    Debug.Print "Ultimo Id Euromillones= " & CStr(mInt)
    mInt = mEng.GetIdUltimoSorteo(gordoPrimitiva)
    Debug.Print "Ultimo Id Gordo= " & CStr(mInt)
    '
    '   5.- Prueba GetListaJuegos
    '
    mVar = mEng.GetListaJuegos
    For i = 0 To UBound(mVar)
        Debug.Print "Nombre juego(" & CStr(i) & ") = " & mVar(i)
    Next i
    '
    '   6.- Prueba GetNewSorteo
    '
    Set mObj = mEng.GetNewSorteo(Bonoloto)
    PrintSorteo mObj
    Set mObj = mEng.GetNewSorteo(LoteriaPrimitiva)
    PrintSorteo mObj
    Set mObj = mEng.GetNewSorteo(Euromillones)
    PrintSorteo mObj
    Set mObj = mEng.GetNewSorteo(gordoPrimitiva)
    PrintSorteo mObj
    '
    '   6.- Prueba GetSorteoByFecha
    '
    mDate = #4/1/2019#
    Set mObj = mEng.GetSorteoByFecha(mDate, Bonoloto)
    PrintSorteo mObj
    mDate = #3/16/2019#
    Set mObj = mEng.GetSorteoByFecha(mDate, LoteriaPrimitiva)
    PrintSorteo mObj
    mDate = #3/5/2019#
    Set mObj = mEng.GetSorteoByFecha(mDate, Euromillones)
    PrintSorteo mObj
    mDate = #1/6/2019#
    Set mObj = mEng.GetSorteoByFecha(mDate, gordoPrimitiva)
    PrintSorteo mObj
    '
    '   7.- Prueba GetSorteoById
    '
    mInt = 5633   '02/03/2015
    Set mObj = mEng.GetSorteoById(mInt, Bonoloto)
    PrintSorteo mObj
    mInt = 2985   '02/02/2017
    Set mObj = mEng.GetSorteoById(mInt, LoteriaPrimitiva)
    PrintSorteo mObj
    mInt = 1041   '13/03/2018
    Set mObj = mEng.GetSorteoById(mInt, Euromillones)
    PrintSorteo mObj
    mInt = 1050   '17/12/2017
    Set mObj = mEng.GetSorteoById(mInt, gordoPrimitiva)
    PrintSorteo mObj
    '
    '   8.- Prueba GetSorteosInFechas
    '
    mDateI = #6/20/2015#
    mDateF = #6/25/2015#
    Set mCol = mEng.GetSorteosInFechas(mDateI, mDateF, Bonoloto)
    PrintSorteos mCol
    mDateI = #4/21/2018#
    mDateF = #5/12/2018#
    Set mCol = mEng.GetSorteosInFechas(mDateI, mDateF, LoteriaPrimitiva)
    PrintSorteos mCol
    mDateI = #10/2/2018#
    mDateF = #10/23/2018#
    Set mCol = mEng.GetSorteosInFechas(mDateI, mDateF, Euromillones)
    PrintSorteos mCol
    mDateI = #4/1/2018#
    mDateF = #6/24/2018#
    Set mCol = mEng.GetSorteosInFechas(mDateI, mDateF, gordoPrimitiva)
    PrintSorteos mCol
    '
    '   9.- Prueba GetSorteosInIds
    '
    mIntI = 6924
    mIntF = 6928
    Set mCol = mEng.GetSorteosInIds(mIntI, mIntF, Bonoloto)
    PrintSorteos mCol
    mIntI = 3203
    mIntF = 3208
    Set mCol = mEng.GetSorteosInIds(mIntI, mIntF, LoteriaPrimitiva)
    PrintSorteos mCol
    mIntI = 1144
    mIntF = 1149
    Set mCol = mEng.GetSorteosInIds(mIntI, mIntF, Euromillones)
    PrintSorteos mCol
    mIntI = 1111
    mIntF = 1114
    Set mCol = mEng.GetSorteosInIds(mIntI, mIntF, gordoPrimitiva)
    PrintSorteos mCol
    '
    '   10.- Prueba SetSorteo
    '
    Set mObj = New Sorteo
    With mObj
        .Id = 6953
        .Juego = Bonoloto
        .CombinacionGanadora.Texto = "5-14-25-26-36-3"
        .Complementario = 44
        .Dia = "L"
        .Fecha = #5/27/2019#
        .NumSorteo = "2019/126"
        .Reintegro = 5
    End With
    mEng.SetSorteo mObj
    Debug.Print "Sorteo guardado. Sorteo=> " & mObj.ToString()
    '
    '   Primitiva
    '
    Set mObj = New Sorteo
    With mObj
        .Id = 3227
        .Juego = LoteriaPrimitiva
        .CombinacionGanadora.Texto = "4-29-7-36-28-44"
        .Complementario = 5
        .Dia = "J"
        .Fecha = #5/30/2019#
        .NumSorteo = "2019/043"
        .Reintegro = 8
    End With
    mEng.SetSorteo mObj
    Debug.Print "Sorteo guardado. Sorteo=> " & mObj.ToString()
    '
    '
    Set mObj = New Sorteo
    With mObj
        .Id = 1167
        .Juego = Euromillones
        .CombinacionGanadora.Texto = "50-22-7-10-38"
        .Estrellas.Texto = "2-11"
        .Dia = "M"
        .Fecha = #5/28/2019#
        .NumSorteo = "2019/043"
    End With
    mEng.SetSorteo mObj
    Debug.Print "Sorteo guardado. Sorteo=> " & mObj.ToString()
    '
    '
    Set mObj = New Sorteo
    With mObj
        .Id = 1125
        .Juego = gordoPrimitiva
        .CombinacionGanadora.Texto = "16-25-8-9-32"
        .Clave = 4
        .Dia = "D"
        .Fecha = #6/2/2019#
        .NumSorteo = "2019/022"
    End With
    mEng.SetSorteo mObj
    Debug.Print "Sorteo guardado. Sorteo=> " & mObj.ToString()
    '
    '   11.- Prueba SetSo
    '
    Set mCol = New Sorteos
    mCol.Juego = Bonoloto
    Set mObj = New Sorteo
    With mObj
        .Id = 6953
        .Juego = Bonoloto
        .CombinacionGanadora.Texto = "5-14-25-26-36-3"
        .Complementario = 44
        .Dia = "L"
        .Fecha = #5/27/2019#
        .NumSorteo = "2019/126"
        .Reintegro = 5
    End With
    mCol.Add mObj
    '
    '
    '
    Set mObj = mEng.GetSorteoById(6928, Bonoloto)
    mObj.CombinacionGanadora.Texto = "8-9-11-25-28-39"
    mObj.Ordenado = False
    mCol.Add mObj
    '
    '
    '
    Set mObj = New Sorteo
    With mObj
        .Id = 6954
        .Juego = Bonoloto
        .CombinacionGanadora.Texto = "45-33-32-14-22-7"
        .Complementario = 10
        .Dia = "L"
        .Fecha = #5/27/2019#
        .NumSorteo = "2019/126"
        .Reintegro = 8
    End With
    '
    '
    '
    mCol.Add mObj
    '
    '
    Debug.Print "Sorteos a modificar: " & mCol.Count
    mEng.SetSorteos mCol
 
 
 On Error GoTo 0
SorteoEngineTest__CleanExit:
    Exit Sub
            
SorteoEngineTest_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    '   Audita el error
    Call HandleException(ErrNumber, ErrDescription, "Lot_PqtConcursosTesting.SorteoEngineTest", ErrSource)
    '   Informa del error
    Call MsgBox(ErrDescription, vbCritical Or vbSystemModal, ThisWorkbook.Name)
    Call Trace("CERRAR")
End Sub


'------------------------------------------------------------------------------*
' Procedimiento  : PrintSorteos
' Fecha          : 17/Jun/2018
' Propósito      : Visualizar los atributos de la colección  sorteos
'------------------------------------------------------------------------------*
Private Sub PrintSorteos(mObj As Sorteos)
    Dim mSort As Sorteo
    Debug.Print "==> Pruebas Sorteos"
    Debug.Print vbTab & "Add                 =" & "#Metodo" 'mObj.Add
    Debug.Print vbTab & "Clear               =" & "#Metodo" 'mObj.Clear
    Debug.Print vbTab & "Count               =" & mObj.Count
    Debug.Print vbTab & "Delete              =" & "#Metodo" 'mObj.Delete
    Debug.Print vbTab & "Items.Count         =" & mObj.Items.Count
    Debug.Print vbTab & "Juego               =" & mObj.Juego
    Debug.Print vbTab & "MarkForDelete       =" & "#Metodo" 'mObj.MarkForDelete
    Debug.Print vbTab & "Undelete            =" & "#Metodo" 'mObj.Undelete
    For Each mSort In mObj.Items
        Debug.Print vbTab & mSort.ToString
    Next mSort
End Sub
'------------------------------------------------------------------------------*
' Procedimiento  : SorteoModelTest
' Fecha          : ma., 01/oct/2019 09:54:45
' Propósito      : Pruebas unitarias de la clase SorteoModel
'------------------------------------------------------------------------------*
Private Sub SorteoModelTest()
    Dim mObj As SorteoModel
    Dim mId As Integer
    Dim mSort As Sorteo
    
  On Error GoTo SorteoModelTest_Error
    '
    '   1.- Objeto en Vacio
    '
    Set mObj = New SorteoModel
    PrintSorteoModel mObj
    '
    '   2.- Pruebas con Bonoloto
    '
    mObj.Juego = LT_BONOLOTO
    '
    '   2.1.- Obtiene el primer sorteo de Bonoloto
    '
    mObj.GetFirstSorteo
    PrintSorteoModel mObj
    mId = mObj.IdSelected
    '
    '   2.2.- Obtiene el Siguiente Sorteo al actual de Bonoloto
    '
    mObj.GetNextSorteoRecord mId
    PrintSorteoModel mObj
    '
    '   2.3.- Obtiene el Ultmimo Sorteo de Bonoloto
    '
    mObj.GetLastSorteo
    PrintSorteoModel mObj
    mId = mObj.IdSelected
    '
    '   2.4.- Obtiene el Sorteo anterior al último de Bonoloto
    '
    mObj.GetPrevSorteoRecord mId
    PrintSorteoModel mObj
    
'    mObj.SearchSorteos
    
    
    
    
    
    
    Err.Raise ERR_TODO, "Lot_PqtConcursosTesting.SorteoModelTest", MSG_TODO
    
'
'    mObj.GetPremios
'    mObj.GetSorteoRecord
'    mObj.NuevoSorteoRecord
'    mObj.GuardarSorteoRecord
'    mObj.EliminarSorteoRecord
    
    
    
    '
    '   2.- Agregar un Nuevo Sorteo
    '
    mObj.Juego = LT_BONOLOTO
    If mObj.NuevoSorteoRecord Then
        PrintSorteoModel mObj
    Else
        Debug.Print vbTab & "#Error SorteoModel.NuevoSorteoRecord (BONOLOTO)"
    End If
    mObj.Juego = LT_PRIMITIVA
    If mObj.NuevoSorteoRecord Then
        PrintSorteoModel mObj
    Else
        Debug.Print vbTab & "#Error SorteoModel.NuevoSorteoRecord (PRIMITIVA)"
    End If
    mObj.Juego = LT_EUROMILLON
    If mObj.NuevoSorteoRecord Then
        PrintSorteoModel mObj
    Else
        Debug.Print vbTab & "#Error SorteoModel.NuevoSorteoRecord (EUROMILLON)"
    End If
    mObj.Juego = LT_GORDO
    If mObj.NuevoSorteoRecord Then
        PrintSorteoModel mObj
    Else
        Debug.Print vbTab & "#Error SorteoModel.NuevoSorteoRecord (GORDO)"
    End If
    '
    '   3.- Buscar un sorteo por ID
    '
    Set mObj = New SorteoModel
    mObj.Juego = Bonoloto
    If mObj.GetSorteoRecord(7063) Then
        PrintSorteoModel mObj
    Else
        Debug.Print vbTab & "#Error SorteoModel.GetSorteoRecord (BONOLOTO)"
    End If
    Set mObj = New SorteoModel
    mObj.Juego = LoteriaPrimitiva
    If mObj.GetSorteoRecord(3257) Then
        PrintSorteoModel mObj
    Else
        Debug.Print vbTab & "#Error SorteoModel.GetSorteoRecord (PRIMITIVA)"
    End If
    Set mObj = New SorteoModel
    mObj.Juego = Euromillones
    If mObj.GetSorteoRecord(1140) Then
        PrintSorteoModel mObj
    Else
        Debug.Print vbTab & "#Error SorteoModel.GetSorteoRecord (EUROMILLON)"
    End If
    Set mObj = New SorteoModel
    mObj.Juego = gordoPrimitiva
    If mObj.GetSorteoRecord(1133) Then
        PrintSorteoModel mObj
    Else
        Debug.Print vbTab & "#Error SorteoModel.GetSorteoRecord (GORDO)"
    End If
    '
    '   3.- Guardar un Sorteo
    '
      
    
    
    '   4.- Eliminar un sorteo
    '   5.- Buscar un sorteo GetSorteoRecord
    '   5.- SearchSorteos
    '
'    mObj.EliminarSorteoRecord mId
'    mObj.GetFirstSorteo
'    mObj.GetLastSorteo
'    mObj.GetNextSorteoRecord
'    mObj.GetPremios
'    mObj.GetPrevSorteoRecord
'    mObj.GetSorteoRecord
'    mObj.GuardarSorteoRecord
'    mObj.NuevoSorteoRecord
'    mObj.SearchSorteos
    '
    '
    '
    '
    '
    '


 On Error GoTo 0
SorteoModelTest__CleanExit:
    Exit Sub
            
SorteoModelTest_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    '   Audita el error
    Call HandleException(ErrNumber, ErrDescription, "Lot_PqtConcursosTesting.SorteoModelTest", ErrSource)
    '   Informa del error
    Call MsgBox(ErrDescription, vbCritical Or vbSystemModal, ThisWorkbook.Name)
    Call Trace("CERRAR")
End Sub
'------------------------------------------------------------------------------*
' Procedimiento  : PrintSorteoModel
' Fecha          : ma., 01/oct/2019 10:00:56
' Propósito      : Visualizar los atributos del objeto SorteoModel
'------------------------------------------------------------------------------*
Private Sub PrintSorteoModel(mObj As SorteoModel)
    Dim i As Integer
    Dim j As Integer
    Debug.Print "==> Pruebas SorteoModel"
    Debug.Print vbTab & "IdSelected          = " & mObj.IdSelected
    Debug.Print vbTab & "Juego               = " & mObj.Juego
    Debug.Print vbTab & "FechaSorteo         = " & mObj.FechaSorteo
    Debug.Print vbTab & "DiaSemana           = " & mObj.DiaSemana
    Debug.Print vbTab & "NumSorteo           = " & mObj.NumSorteo
    Debug.Print vbTab & "Semana              = " & mObj.Semana
    Debug.Print vbTab & "OrdenAparicion      = " & mObj.OrdenAparicion
    Debug.Print vbTab & "Reintegro           = " & mObj.Reintegro
    Debug.Print vbTab & "CombinacionGanadora = " & mObj.CombinacionGanadora
    Debug.Print vbTab & "Complementario      = " & mObj.Complementario
    Debug.Print vbTab & "N1                  = " & mObj.N1
    Debug.Print vbTab & "N2                  = " & mObj.N2
    Debug.Print vbTab & "N3                  = " & mObj.N3
    Debug.Print vbTab & "N4                  = " & mObj.N4
    Debug.Print vbTab & "N5                  = " & mObj.N5
    Debug.Print vbTab & "N6                  = " & mObj.N6
    Debug.Print vbTab & "Estrellas           = " & mObj.Estrellas
    Debug.Print vbTab & "E1                  = " & mObj.E1
    Debug.Print vbTab & "E2                  = " & mObj.E2
    Debug.Print vbTab & "FechaFin            = " & mObj.FechaFin
    Debug.Print vbTab & "FechaInicio         = " & mObj.FechaInicio
    Debug.Print vbTab & "LineasPorPagina     = " & mObj.LineasPorPagina
    Debug.Print vbTab & "PaginaActual        = " & mObj.PaginaActual
    Debug.Print vbTab & "TotalPaginas        = " & mObj.TotalPaginas
    Debug.Print vbTab & "TotalRegistros      = " & mObj.TotalRegistros
    Debug.Print vbTab & "MatrizPremios       = "
    If IsArray(mObj.MatrizPremios) Then
        For i = 0 To UBound(mObj.MatrizPremios)
            Debug.Print vbTab & "(" & CStr(i) & ")=>" & mObj.MatrizPremios(i)
        Next i
    Else
        Debug.Print vbTab & "MatrizPremios       = " & mObj.MatrizPremios
    End If
    Debug.Print vbTab & "ResultadosSearch    ="
    For i = 0 To UBound(mObj.ResultadosSearch, 1)
        For j = 0 To UBound(mObj.ResultadosSearch, 2)
            Debug.Print vbTab & "(" & CStr(i) & ", " & CStr(j) & ")=>" & mObj.ResultadosSearch(i, j)
        Next j
    Next i
    Debug.Print "================"
End Sub

' *===========(EOF): Lot_PqtConcursosTesting.bas

