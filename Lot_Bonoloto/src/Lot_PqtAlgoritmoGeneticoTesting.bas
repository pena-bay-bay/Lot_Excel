Attribute VB_Name = "Lot_PqtAlgoritmoGeneticoTesting"
' *============================================================================*
' *
' *     Fichero    : Lot_PqtAlgoritmoGeneticoTesting.bas
' *
' *     Autor      : Carlos Almela Baeza
' *     Creaci�n   : Mar, 01/May/2018 230:38:25
' *     Versi�n    : 1.0
' *     Prop�sito  : Pruebas unitarias de las clases del paquete Algoritmo
' *                  Genetico
' *
' *============================================================================*
Option Explicit
Option Base 0
'------------------------------------------------------------------------------*
' Procedimiento  : ParamProcesoTest
' Fecha          : 25/may/2018
' Prop�sito      :
' Par�metros     :
' Retorno        :
'------------------------------------------------------------------------------*
Public Sub ParamProcesoTest()
    Dim mObj As ParamProceso
    Dim mTmpInt As Integer
    Dim mTmpDate As Date
    Dim mTmpString As String
    Dim mTmpDouble As Double
    Dim mTmpCurrency As Currency
    
    mTmpInt = 58
    mTmpDate = Now()
    mTmpDouble = 0.34569
    mTmpCurrency = 1256.58
    
    Set mObj = New ParamProceso
    Debug.Print mObj.ToString()
    With mObj
        .Valor = mTmpInt
        .Nombre = "NUMERO"
        .Concepto = "Parametro entero concepto"
    End With
    Debug.Print mObj.ToString()
    ' PrintParam
    With mObj
        .Valor = mTmpDate
        .Nombre = "FECHA_HOY"
        .Concepto = "Parametro fecha ejemplo"
    End With
    Debug.Print mObj.ToString()
    With mObj
        .Valor = mTmpDouble
        .Nombre = "DECIMAL"
        .Concepto = "Parametro decimal ejemplo"
    End With
    Debug.Print mObj.ToString()
    With mObj
        .Valor = mTmpCurrency
        .Nombre = "MONEDA"
        .Concepto = "Parametro moneda ejemplo"
    End With
    Debug.Print mObj.ToString()
       With mObj
        .Valor = True
        .Nombre = "BOOLEAN"
        .Concepto = "Parametro Boolean ejemplo"
    End With
    Debug.Print mObj.ToString()
    'TODO: Probar tipo de variable
End Sub

'------------------------------------------------------------------------------*
' Procedimiento  : IndividuoTest
' Fecha          : 25/may/2018
' Prop�sito      :
'------------------------------------------------------------------------------*
Public Sub IndividuoTest()
    Dim mObj As Individuo
    Dim mComb As Combinacion
 On Error GoTo IndividuoTest_Error
    '
    '   Creado en vacio
    '
    Set mObj = New Individuo
    Debug.Print "Objeto vacio : " & mObj.ToString
    Set mObj = Nothing
    '
    '   Individuo de juego 6/49
    '
    Set mObj = New Individuo
    Set mComb = New Combinacion
    mComb.Texto = "12-25-16-3-1-45"
    With mObj
        Set .Genoma = mComb
        .IdPoblacion = "POB20180525T230545"
        .Juego = Bonoloto
        .Mutado = True
    End With
    Debug.Print "Objeto 6/49 : " & mObj.ToString
    Set mObj = Nothing
    '
    '   Individuo de juego  6/54
    '
    '
    '   Individuo de juego  6/50
    '
    '
    '
    '
    
 On Error GoTo 0
IndividuoTest__CleanExit:
    Exit Sub
    
IndividuoTest_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    '   Audita el error
    Call HandleException(ErrNumber, ErrDescription, "Lot_AlgoritmoGeneticoTesting.IndividuoTest", ErrSource)
    '   Informa del error
    Call MsgBox(ErrDescription, vbCritical Or vbSystemModal, Application.Caption)
    Call Trace("CERRAR")
End Sub
'------------------------------------------------------------------------------*
' Procedimiento  : PoblacionTest
' Fecha          : 25/may/2018
' Prop�sito      :
'------------------------------------------------------------------------------*
Public Sub PoblacionTest()
    Dim mObj As Poblacion
    Dim mIdv As Individuo
    Dim mComb As Combinacion
    
 On Error GoTo PoblacionTest_Error
    '
    '   Creado en vacio
    '
    Set mObj = New Poblacion
    Debug.Print "Objeto vacio : " & mObj.ToString
    Set mObj = Nothing
    '
    '   Agregar un individuo a la poblaci�n
    '
    Set mIdv = New Individuo
    Set mComb = New Combinacion
    mComb.Texto = "12-25-16-3-1-45"
    With mIdv
        Set .Genoma = mComb
        .Juego = Bonoloto
        .Mutado = True
        .Fitness = 2300
    End With
    Set mObj = New Poblacion
    With mObj
        .Generacion = 1
        .Juego = Bonoloto
    End With
    mObj.Add mIdv
    Debug.Print "Objeto 6/49 : " & mObj.ToString
    '
    '   Agrega otro individuo
    '
    Set mIdv = New Individuo
    Set mComb = New Combinacion
    mComb.Texto = "36-16-49-8-47-22"
    With mIdv
        Set .Genoma = mComb
        .Juego = Bonoloto
        .Fitness = 145
    End With
    mObj.Add mIdv
    Debug.Print "Objeto 6/49 : " & mObj.ToString
    '
    '   Obtiene el individuo iesimo
    '
    '   Ordena tres individuos por fitness
    '
    '   Inicializa la poblaci�n
    '
    '
 On Error GoTo 0
PoblacionTest__CleanExit:
    Exit Sub
    
PoblacionTest_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    '   Audita el error
    Call HandleException(ErrNumber, ErrDescription, "Lot_AlgoritmoGeneticoTesting.PoblacionTest", ErrSource)
    '   Informa del error
    Call MsgBox(ErrDescription, vbCritical Or vbSystemModal, Application.Caption)
    Call Trace("CERRAR")
End Sub
'------------------------------------------------------------------------------*
' Procedimiento  : ParametrosProcesoTest
' Fecha          : 25/may/2018
' Prop�sito      : Realizar las pruebas unitarias de la clase Parametros Proceso
'------------------------------------------------------------------------------*
Public Sub ParametrosProcesoTest()
    Dim mPar As ParamProceso
    Dim mObj As ParametrosProceso
    
 On Error GoTo ParametrosProcesoTest_Error
    '
    '   Objeto Vacio
    '
    Set mObj = New ParametrosProceso
    Print_ParametrosProceso mObj
    '
    '   Agregar una varible
    '
    Set mPar = New ParamProceso
    With mPar
        .Concepto = "Variable de prueba String"
        .Nombre = "VARTEXTO"
        .Valor = "Texto de Prueba"
    End With
    mObj.Add mPar
    Print_ParametrosProceso mObj
    Set mPar = Nothing
    '
    '   Agregar una segunda varible
    '
    Set mPar = New ParamProceso
    With mPar
        .Concepto = "Variable de prueba Entero"
        .Nombre = "VARENTERO"
        .Valor = 1254
    End With
    mObj.Add mPar
    Print_ParametrosProceso mObj
    Set mPar = Nothing
    '
    '   Agregar una tercera varible
    '
    Set mPar = New ParamProceso
    With mPar
        .Concepto = "Variable de prueba Fecha"
        .Nombre = "VARFECHA"
        .Valor = #5/1/2018#
    End With
    mObj.Add mPar
    Print_ParametrosProceso mObj
    Set mPar = Nothing
    '
    '   Probar referencia a Items
    '
    Debug.Print "Probar variable (2) :" & mObj.Items(2).Valor
    '
    '   Probar m�todo GetVarible
    '
    Set mPar = mObj.GetVariable("VARENTERO")
    Debug.Print "Probar valor variable('VARENTERO') : " & mPar.ToString
    '
    '   Probar metodo Delete (3)
    '
    mObj.Delete mPar
    Print_ParametrosProceso mObj
    '
    '   Probar propiedad COUNT
    '
    Debug.Print "Probar Propiedad COUNT: " & mObj.Count
    '
    '   Probar m�todo Clear
    '
    mObj.Clear
    Print_ParametrosProceso mObj
    Set mPar = Nothing
    Set mObj = Nothing
 On Error GoTo 0
ParametrosProcesoTest__CleanExit:
    Exit Sub
    
ParametrosProcesoTest_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    '   Audita el error
    Call HandleException(ErrNumber, ErrDescription, "Lot_AlgoritmoGeneticoTesting.ParametrosProcesoTest", ErrSource)
    '   Informa del error
    Call MsgBox(ErrDescription, vbCritical Or vbSystemModal, Application.Caption)
    Call Trace("CERRAR")
End Sub
'------------------------------------------------------------------------------*
' Procedimiento  : Print_ParametrosProceso
' Fecha          : 25/may/2018
' Prop�sito      : Visualizar la clase ParametrosProceso
'------------------------------------------------------------------------------*
Private Sub Print_ParametrosProceso(Obj As ParametrosProceso)
Debug.Print "==> Pruebas ParametrosProceso"
    Debug.Print vbTab & "Count                      =" & Obj.Count
    Debug.Print vbTab & "Items.Count                =" & Obj.Items.Count
    Debug.Print vbTab & "Add                        = Obj.Add"
    Debug.Print vbTab & "Clear                      = Obj.Clear"
    Debug.Print vbTab & "Delete                     = Obj.Delete"
    Debug.Print vbTab & "GetVariable                = Obj.GetVariable"
End Sub

' *===========(EOF): Lot_AlgGen_InterfazUI.bas

