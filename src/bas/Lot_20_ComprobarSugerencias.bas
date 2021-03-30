Attribute VB_Name = "Lot_20_ComprobarSugerencias"
' *============================================================================*
' *
' *     Fichero    : Lot_20_ComprobarSugerencias.bas
' *
' *     Autor      : Carlos Almela Baeza
' *     Creación   : mie, 02/dic/2020 19:34:26
' *     Versión    : 1.0
' *     Propósito  : Realizar la comprobación masiva de numeros indexada
' *
' *============================================================================*
Option Explicit
Option Base 0
'
'--- Variables Privadas -------------------------------------------------------*
Dim mArrayNumeros As Variant
Private mRango        As Range
Private mFila         As Range
Private mSorteo       As Sorteo
Private mApuesta      As Apuesta
Private mCmpBoleto    As ComprobarBoletos

'--- Constantes ---------------------------------------------------------------*
Public Const AREA_SORTEO As String = "Entrada!G4:U4"
Public Const AREA_NUMEROS As String = "Entrada!B5:C34"

'--- Mensajes -----------------------------------------------------------------*
Private Const MSG_MATRIZERROR As String = "Error con la matriz de conversión"

'--- Errores ------------------------------------------------------------------*
Private Const ERR_MATRIZERROR As Long = 10001


'--- Métodos Privados ---------------------------------------------------------*
'------------------------------------------------------------------------------*
' Procedimiento  : getMatrizNumeros
' Fecha          : 02/dic/2020
' Propósito      : Obtiene los números indexados en una matriz de 2 x n
'------------------------------------------------------------------------------*
Private Function getMatrizNumeros() As Variant
    Dim i As Integer
 On Error GoTo getMatrizNumeros_Error
    '
    '   Obtenemos el array
    '
    mArrayNumeros = Range(AREA_NUMEROS).Value2
    '
    '   Verificar si tiene todos los numeros o sobran redimensionar
    '
    For i = 1 To UBound(mArrayNumeros)
        If mArrayNumeros(i, 1) = i - 1 Then
            Err.Raise ERR_MATRIZERROR, "getMatrizNumeros", MSG_MATRIZERROR
        End If
        If Not IsNumeric(mArrayNumeros(i, 2)) Then
            Err.Raise ERR_MATRIZERROR, "getMatrizNumeros", MSG_MATRIZERROR
        ElseIf (CInt(mArrayNumeros(i, 2)) < 1) Or _
               (CInt(mArrayNumeros(i, 2)) > 49) Then
            Err.Raise ERR_MATRIZERROR, "getMatrizNumeros", MSG_MATRIZERROR
        End If
        
    Next i
    '
    '   Devolvemos la matriz
    '
    getMatrizNumeros = mArrayNumeros
    
    
  On Error GoTo 0
getMatrizNumeros__CleanExit:
    Exit Function
getMatrizNumeros_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    Call HandleException(ErrNumber, ErrDescription, "Lot_20_ComprobarSugerencias.getMatrizNumeros", ErrSource)
    Err.Raise ErrNumber, "Lot_20_ComprobarSugerencias.getMatrizNumeros", ErrDescription
End Function

'------------------------------------------------------------------------------*
' Procedimiento  : getMatrizNumeros
' Fecha          : 02/dic/2020
' Propósito      : Obtiene los números indexados en una matriz de 2 x n
'------------------------------------------------------------------------------*
Private Function getSorteo() As Sorteo
    Dim mRg As Range
  On Error GoTo getSorteo_Error
    '
    '   Creamos la variable sorteo
    '
    Set mSorteo = New Sorteo
    '
    '   Obtenemos el rango de la información
    '
    Set mRg = Range(AREA_SORTEO)
    '
    '   Cargamos el Sorteo
    '
    mSorteo.Constructor mRg
    '
    '   Devolvemos el sorteo
    '
    Set getSorteo = mSorteo
  On Error GoTo 0
getSorteo__CleanExit:
    Exit Function
getSorteo_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    Call HandleException(ErrNumber, ErrDescription, "Lot_20_ComprobarSugerencias.getSorteo", ErrSource)
    Err.Raise ErrNumber, "Lot_20_ComprobarSugerencias.getSorteo", ErrDescription
End Function

'------------------------------------------------------------------------------*
' Procedimiento  : getApuesta
' Fecha          : 03/dic/2020
' Propósito      : Obtiene una apuesta válida a partir de una fila números
'------------------------------------------------------------------------------*
Private Function getApuesta(newMtzNumeros As Variant, newFila As Range) As Apuesta
    Dim i       As Integer
    Dim j       As Variant
    Dim mNum    As Numero
  On Error GoTo getApuesta_Error
    '
    '   Creamos la variable apuesta
    '
    Set mApuesta = New Apuesta
    '
    '  Recorremos el rango para obtener los indices de la apuesta
    '
    For i = 1 To 6 Step 1
        '
        '   Obtenemos el indice
        j = newFila.Value2(1, i)
        '
        '   Si es numérico resolvemos el numero
        '
        If IsNumeric(j) Then
            '
            '   Creamos el numero
            '
            Set mNum = New Numero
            '
            '   Obtenemos el número de la matriz
            '
            mNum.Valor = newMtzNumeros(j, 2)
            '
            '   Agregamos numero a la apuesta
            '
            mApuesta.Combinacion.Add mNum
        End If
    Next i
    '
    '   Devolvemos la apuesta
    '
    Set getApuesta = mApuesta
    
  On Error GoTo 0
getApuesta__CleanExit:
    Exit Function
getApuesta_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    Call HandleException(ErrNumber, ErrDescription, "Lot_20_ComprobarSugerencias.getApuesta", ErrSource)
    Err.Raise ErrNumber, "Lot_20_ComprobarSugerencias.getApuesta", ErrDescription
End Function



'--- Métodos Públicos ---------------------------------------------------------*
'------------------------------------------------------------------------------*
' Procedimiento  : Btn_Ejecutar
' Fecha          : 02/dic/2020
' Propósito      : Ejecutar la rutina de verificación
'------------------------------------------------------------------------------*
Public Sub Btn_Ejecutar()
    Dim n As Integer
    Dim m As Integer
    Dim mRes As String
  On Error GoTo Btn_Ejecutar_Error
    Ir_A_Hoja "Salida"
    CALCULOOFF
    '
    '   Caturar los numeros en una matriz
    '
    mArrayNumeros = getMatrizNumeros()
    '
    '   Carga sorteo para comprobar
    '
    Set mSorteo = getSorteo()
    '
    '   Obtenemos el rango de trabajo, eliminando la cabecera
    '
    Set mRango = Range("A2").CurrentRegion
    n = mRango.Rows.Count - 1
    m = mRango.Columns.Count
    Set mRango = mRango.Offset(1, 0).Resize(n, m)
    '
    '   Borra la salida del proceso
    '
    mRango.Offset(0, 6).Resize(n, 3).Delete
    '
    '   Nos posicionamos en la primera celda
    '
    Range("A1").Activate
    '
    '   Creamos el objeto comprobador del sorteo y establecemos
    '   el sorteo a comprobar
    '
    Set mCmpBoleto = New ComprobarBoletos
    Set mCmpBoleto.Sorteo = mSorteo
    '
    '   Recorremos el bucle
    '
    For Each mFila In mRango.Rows
        '
        '   Obtener sugerencia en forma de apuesta
        '   #TODO: cambiar apuesta por sugerencia que no tiene fecha
        '
        Set mApuesta = getApuesta(mArrayNumeros, mFila)
        mApuesta.Fecha = mSorteo.Fecha
        '
        '   Obtenemos el resultado
        '
        mRes = mCmpBoleto.ComprobarApuesta(mApuesta, False)
        '
        '   Si se ha acertado bolas
        '
        If mCmpBoleto.BolasAcertadas > 0 Then
            ActiveCell.Offset(mFila.Row - 1, 6).Value = mCmpBoleto.CategoriaPremioTxt
            If mCmpBoleto.CatPremioApuesta <> Ninguna Then
                ActiveCell.Offset(mFila.Row - 1, 8).Value = mCmpBoleto.ImporteApuesta
            End If
        End If
        ActiveCell.Offset(mFila.Row - 1, 7).Value = mApuesta.Coste
    Next mFila
    '
    '
    '
    CALCULOON
    
Btn_Ejecutar_Exit:
  On Error GoTo 0
    Exit Sub
Btn_Ejecutar_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    Call HandleException(ErrNumber, ErrDescription, "Lot_20_ComprobarSugerencias.Btn_Ejecutar")
    Call MsgBox(ErrDescription, vbError Or vbSystemModal, NOMBRE_APLICACION)
    Call Trace("CERRAR")
    CALCULOON
End Sub
' *===========(EOF): Lot_20_ComprobarSugerencias.bas
