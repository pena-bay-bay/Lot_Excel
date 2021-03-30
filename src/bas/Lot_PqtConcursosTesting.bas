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
' *                    - Sorteo
' *                    - Sorteos
' *                    - SorteoEngine
' *                    - Premio
' *                    - Premios
' *                    - InfoSorteo
' *                    - Tarifa
' *
' *============================================================================*
Option Explicit
Option Base 0



'---------------------------------------------------------------------------------------
' Procedure : PremioTest
' Author    : CHARLY
' Date      : sáb, 01/nov/2014 21:01:42
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub PremioTest()
    Dim mObj As Premio
    
  On Error GoTo PremioTest_Error
    '
    '   Caso de pruebas 01: Objeto Vacio
    '
    Set mObj = New Premio
    PrintPremio mObj
    '
    '   Caso de pruebas 02: Premio valido Bonoloto
    '
    With mObj
        .Id = 1
        .CategoriaPremio = Primera
        .Importe = 160000.89
        .NumeroAcertantesEspaña = 152
    End With
    PrintPremio mObj
    '
    '   Caso de pruebas 03: Premio valido Loteria Primitiva
    '
    Set mObj = New Premio
    mObj.Juego = LoteriaPrimitiva
    mObj.UnPack "2,3,13200.61"
    PrintPremio mObj
    '
    '   Caso de pruebas 04: Premio valido Euromillones
    '
    Set mObj = New Premio
    mObj.Juego = Euromillones
    mObj.UnPack "4,206,113.47,1068"
    PrintPremio mObj
    '
    '   Caso de pruebas 05: Premio valido GordoPrimitiva
    '
    Set mObj = New Premio
    mObj.Juego = GordoPrimitiva
    mObj.UnPack "7,164441,3.00"
    PrintPremio mObj
    '
    '   Caso de pruebas 06: Premio No valido
    '
    Set mObj = New Premio
    mObj.UnPack "0,0,0"
    PrintPremio mObj
    '
    '   Caso de pruebas 07: Reintegro
    '
    Set mObj = New Premio
    mObj.UnPack "15,407755,0.50"
    PrintPremio mObj
    '
    '   Caso de pruebas 08: Metodo Parse
    '
    Set mObj = New Premio
    mObj.Parse "Juego: Bonoloto, Categoria: 15 = Reintegro, Importe: 0.5 Euros, Acertantes: 589644"
    PrintPremio mObj
    mObj.Parse "Juego: Euro Millones, Categoria: 4 = 4ª 4 + 2, Importe: 113.47 Euros, Acertantes: 206 Esp y 1068 Eur"
    PrintPremio mObj
    
  On Error GoTo 0
    Exit Sub
PremioTest_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    Call HandleException(ErrNumber, ErrDescription, "Lot_PqtConcursosTesting.PremioTest")
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
Private Sub PrintPremio(obj As Premio)
    Debug.Print " Premio ]================="
    Debug.Print "   CategoriaPremio        =>" & obj.CategoriaPremio
    Debug.Print "   CategoriaTexto         =>" & obj.CategoriaTexto
    Debug.Print "   ID                     =>" & obj.Id
    Debug.Print "   Importe                =>" & obj.Importe
    Debug.Print "   Juego                  =>" & obj.Juego
    Debug.Print "   Acertantes en España   =>" & obj.NumeroAcertantesEspaña
    Debug.Print "   Acertantes en Europa   =>" & obj.NumeroAcertantesEuropa
    Debug.Print "   EsValido()             =>" & obj.EsValido()
    Debug.Print "   ToString()             =>" & obj.ToString()
    Debug.Print "   Pack()                 =>" & obj.Pack()
End Sub




'---------------------------------------------------------------------------------------
' Procedure : PremioTest
' Author    : CHARLY
' Date      :
' Purpose   : Pruebas unitarias de la clase Premios
'---------------------------------------------------------------------------------------
'
Private Sub PremiosTest()
    Dim mObj    As Premios
    Dim mPrem   As Premio
    
  On Error GoTo PremiosTest_Error
    '
    '   Caso de pruebas 01: Objeto Vacio
    '
    Set mObj = New Premios
    PrintPremios mObj
    '
    '   Caso de pruebas 02: metodo parse Bonoloto
    '
    '    1ª (6 Aciertos)         0      0,00 €
    '    2ª (5 Aciertos + C)     0      0,00 €
    '    3ª (5 Aciertos)        68  3.134,92 €
    '    4ª (4 Aciertos)     3.773     29,82 €
    '    5ª (3 Aciertos)    69.648      4,00 €
    '    Reintegro         382.036      0,50 €
    mObj.Parse "1,0,0;2,0,0;3,68,3134.92;4,3773,29.82;5,69648,4;15,382036,.5"
    PrintPremios mObj
    '
    '   Caso de pruebas 03: Parse Primitiva
    '
    'Categorías                 Acertantes  Premios
    'Especial (6 Aciertos + R)  0           0,00 €
    '1ª (6 Aciertos)            0           0,00 €
    '2ª (5 Aciertos + C)        4           50.019,82 €
    '3ª (5 Aciertos)            227         1.909,71 €
    '4ª (4 Aciertos)            12.883      54,36 €
    '5ª (3 Aciertos)            236.342     8,00 €
    'Reintegro                  1.115.279   1,00 €
    Set mObj = New Premios
    mObj.Juego = LoteriaPrimitiva
    mObj.Parse "14,0,0;1,0,0;2,4,50019.82;3,227,1909.71;4,12883,54.36;5,236342,8;15,1115279,1"
    PrintPremios mObj
    '
    '   Caso de pruebas 04: Parse Gordo
    '
    'Categorías  Acertantes  Premios
    '1ª (5 + 1)  0           0,00 €
    '2ª (5 + 0)  2           76.010,48 €
    '3ª (4 + 1)  31          891,62 €
    '4ª (4 + 0)  162         199,05 €
    '5ª (3 + 1)  1.370       26,90 €
    '6ª (3 + 0)  8.563       13,99 €
    '7ª (2 + 1)  20.336      4,53 €
    '8ª (2 + 0)  130.650     3,00 €
    'Reintegro   322.686     1,50 €

    Set mObj = New Premios
    mObj.Juego = GordoPrimitiva
    mObj.Parse "1,0,0;2,2,76010.48;3,31,891.62;4,162,199.05;5,1370,26.90;6,8563,13.99;7,20336,4.53;8,130650,3.00;15,322686,1.50"
    PrintPremios mObj
    '
    '   Caso de pruebas 05: Parse Euromillon
    '
    'Categorías  Acertantes  Premios         Acertantes Europa
    '1ª 5 + 2    0           0,00 €          0
    '2ª 5 + 1    0           0,00 €          0
    '3ª 5 + 0    1           101.330,71 €    6
    '4ª 4 + 2    2           1.888,15 €      19
    '5ª 4 + 1    98          135,70 €        487
    '6ª 3 + 2    190         85,41 €         818
    '7ª 4 + 0    273         43,14 €         1.138
    '8ª 2 + 2    2.755       19,97 €         12.292
    '9ª 3 + 1    4.884       12,57 €         21.782
    '10ª 3 + 0   11.295      9,94 €          51.291
    '11ª 1 + 2   15.012      9,29 €          66.459
    '12ª 2 + 1   72.027      5,97 €          325.980
    '13ª 2 + 0   170.677     4,02 €          778.495
    
    Set mObj = New Premios
    mObj.Juego = Euromillones
    mObj.Parse "1,0,0,0;2,0,0,0;3,1,101330.71,6;4,2,1888.15,19;" & _
               "5,98,135.70,487;6,190,85.41,818;7,273,43.14,113" & _
               "8;8,2755,19.97,12292;9,4884,12.57,21782;10,1129" & _
               "5,9.94,51291;11,15012,9.29,66459;12,72027,5.97," & _
               "325980;13,170677,4.02,778495"
    PrintPremios mObj
    
    '
    '   Caso de pruebas 06: Item
    '
    Debug.Print " Prueba Item ----------"
    Set mPrem = mObj.Items(4)
    PrintPremio mPrem
    '
    '
    '   Caso de pruebas 08: GetPremioByCategoria
    '
    Debug.Print " Prueba GetPremioByCategoria ----------"
    Set mPrem = mObj.GetPremioByCategoria(Duodecima)
    PrintPremio mPrem

  
  On Error GoTo 0
    Exit Sub
PremiosTest_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    Call HandleException(ErrNumber, ErrDescription, "Lot_PqtConcursosTesting.PremiosTest")
    Call MsgBox(ErrDescription, vbCritical Or vbSystemModal, ThisWorkbook.Name)
    Call Trace("CERRAR")
End Sub


'---------------------------------------------------------------------------------------
' Procedure : PrintPremios
' Author    : CHARLY
' Date      : 14/07/2020
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub PrintPremios(obj As Premios)
    Debug.Print " Premios ]================="
    Debug.Print "   IdSorteo               =>" & obj.IdSorteo
    Debug.Print "   Juego                  =>" & obj.Juego
    Debug.Print "   ToString()             =>" & obj.ToString
    Debug.Print "   Count                  =>" & obj.Count
End Sub





'---------------------------------------------------------------------------------------
' Procedure : InfoSorteoTest
' Author    : CHARLY
' Date      : 06/04/2014
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub InfoSorteoTest()
    Dim mInfo As InfoSorteo
    Dim i As Integer
    Dim mFechaI As Date
    Dim mFechaF As Date
    Dim mDias   As Integer
  
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
        mInfo.Constructor GordoPrimitiva
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


  On Error GoTo 0
    Exit Sub
InfoSorteoTest_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    Call HandleException(ErrNumber, ErrDescription, "Lot_PqtConcursosTesting.InfoSorteoTest")
    Call MsgBox(ErrDescription, vbCritical Or vbSystemModal, ThisWorkbook.Name)
    Call Trace("CERRAR")
End Sub




'------------------------------------------------------------------------------*
' Procedimiento  : SorteoEngineTest
' Fecha          : lu., 06/jul/2020 18:34:21
' Propósito      : Casos de prueba de la clase SorteoEngine
'------------------------------------------------------------------------------*
'
Private Sub SorteoEngineTest()
    Dim mObj    As SorteoEngine
    Dim mSorteo As Sorteo
    Dim mFecha  As Date
    
  On Error GoTo SorteoEngineTest_Error
    '
    ' Caso de prueba 01 Obtener un sorteo existente
    '
    Debug.Print "#============= TestCase: 01 "
    Set mObj = New SorteoEngine
    mFecha = #6/15/2020#
    Set mSorteo = mObj.GetSorteoByFecha(mFecha)
    If Not (mSorteo Is Nothing) Then
        If mSorteo.Fecha = mFecha Then
            Debug.Print ("Prueba Correcta, sorteo localizado: " & mSorteo.ToString())
            PrintSorteo mSorteo
        Else
            Debug.Print ("#Error en GetSorteoByFecha : " & mSorteo.ToString())
        End If
    Else
        Debug.Print ("#Error GetSorteoByFecha: mSorteo is Nothing")
    End If
    '
    ' Caso de prueba 02 Obtener un sorteo inexistente
    '
    Debug.Print "#============= TestCase: 02 "
    Set mObj = New SorteoEngine
    mFecha = #1/5/2020#   ' Domingo
    Set mSorteo = mObj.GetSorteoByFecha(mFecha)
    If Not (mSorteo Is Nothing) Then
        If mSorteo.Fecha = mFecha Then
            Debug.Print ("Prueba Correcta, sorteo localizado: " & mSorteo.ToString())
            PrintSorteo mSorteo
        Else
            Debug.Print ("#Error en GetSorteoByFecha : " & mSorteo.ToString())
        End If
    Else
        Debug.Print ("#Error GetSorteoByFecha: mSorteo is Nothing")
    End If
    '
    ' Caso de prueba 03 Obtener un sorteo existente con Premios
    '
    Debug.Print "#============= TestCase: 03 "
    Set mObj = New SorteoEngine
    mFecha = #7/7/2020#     ' Martes
    Set mSorteo = mObj.GetSorteoByFecha(mFecha)
    If Not (mSorteo Is Nothing) Then
        If mSorteo.Fecha = mFecha Then
            Debug.Print ("Prueba Correcta, sorteo localizado: " & mSorteo.ToString())
            PrintSorteo mSorteo
        Else
            Debug.Print ("#Error en GetSorteoByFecha : " & mSorteo.ToString())
        End If
    Else
        Debug.Print ("#Error GetSorteoByFecha: mSorteo is Nothing")
    End If
    
    
    
  On Error GoTo 0
SorteoEngineTest_CleanExit:
    Exit Sub
SorteoEngineTest_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    Call HandleException(ErrNumber, ErrDescription, "Lot_PqtConcursosTesting.SorteoEngineTest", ErrSource)
    Call MsgBox(ErrDescription, vbCritical Or vbSystemModal, Application)
    Call Trace("CERRAR")
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
        .Combinacion.Texto = "10-49-15-31-17-7"
        .Complementario = 34
        .Ordenado = True
        .Fecha = #5/15/2018#
        .Id = 4512
        .NumeroSorteo = "2018/116"
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
        .Combinacion.Texto = "49-14-4-1-21"
        .Estrellas.Texto = "2-12"
        .Ordenado = True
        .Fecha = #7/13/2018#
        .Id = 1076
        .NumeroSorteo = "2018/056"
    End With
    PrintSorteo mObj
    '
    '   4.- Sorteo Gordo
    '
    '  1074    2018/022    03/06/2018  D   22  Si  28  38  44  33  5  C 4
    '
    Set mObj = New Sorteo
    With mObj
        .Juego = GordoPrimitiva
        .Dia = "D"
        .Combinacion.Texto = "28-38-44-33-5"
        .Reintegro = 4
        .Ordenado = True
        .Fecha = #6/3/2018#
        .Id = 1074
        .NumeroSorteo = "2018/022"
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
        .Combinacion.Texto = "44-30-37-16-14-5"
        .Complementario = 23
        .Ordenado = True
        .Fecha = #6/30/2018#
        .Id = 3132
        .NumeroSorteo = "2018/052"
        .Reintegro = 3
    End With
    PrintSorteo mObj
        
 On Error GoTo 0
SorteoTest__CleanExit:
    Exit Sub
            
SorteoTest_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    Call HandleException(ErrNumber, ErrDescription, "Lot_PqtConcursosTesting.SorteoTest", ErrSource)
    Call MsgBox(ErrDescription, vbCritical Or vbSystemModal, ThisWorkbook.Name)
    Call Trace("CERRAR")
End Sub

'---------------------------------------------------------------------------------------
' Procedure : PintarSorteo
' Author    : Charly
' Date      :
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub PrintSorteo(oSorteo As Sorteo)
    Debug.Print "==> Sorteo "
    Debug.Print "Combinacion          = " & oSorteo.Combinacion.Texto
    Debug.Print "Complementario       = " & oSorteo.Complementario
    Debug.Print "Dia                  = " & oSorteo.Dia
    Debug.Print "Constructor          = " ' oSorteo.Constructor()
    Debug.Print "ElMillon             = " & oSorteo.ElMillon
    Debug.Print "Id                   = " & oSorteo.Id
    Debug.Print "EstrellaDos          = " & oSorteo.EstrellaDos.Valor
    Debug.Print "Estrellas            = " & oSorteo.Estrellas.Texto
    Debug.Print "EstrellaUno          = " & oSorteo.EstrellaUno.Valor
    Debug.Print "Fecha                = " & oSorteo.Fecha
    Debug.Print "GetMensaje()         = " & oSorteo.GetMensaje()
    Debug.Print "ID                   = " & oSorteo.Id
    Debug.Print "ImporteBote          = " & oSorteo.ImporteBote
    Debug.Print "ImporteVenta         = " & oSorteo.ImporteVenta
    Debug.Print "Joker                = " & oSorteo.Joker
    Debug.Print "Juego                = " & oSorteo.Juego
    Debug.Print "NumeroSorteo         = " & oSorteo.NumeroSorteo
    Debug.Print "Ordenado             = " & oSorteo.Ordenado
    Debug.Print "Premios.ToString()   = " & oSorteo.Premios.ToString()
    Debug.Print "Reintegro            = " & oSorteo.Reintegro
    Debug.Print "Semana               = " & oSorteo.Semana
    Debug.Print "ToString()           = " & oSorteo.ToString()
End Sub

' *===========(EOF): Lot_PqtConcursosTesting.bas

