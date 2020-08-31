Attribute VB_Name = "Lot_01_ComprobarBoletos"
' *============================================================================*
' *
' *     Fichero    : Lot_06_ComprobarBoletos.bas
' *
' *     Autor      : Carlos Almela Baeza
' *     Creación   : 06/09/2014 19:05
' *     Revisión   : ma., 30/jun/2020 19:48:29
' *     Versión    : 1.0
' *     Propósito  : Caso de Uso Comprobar Boletos Apostados.
' *                  Comprueba los pronosticos acertados
' *============================================================================*
Option Explicit
Option Base 0

'--- Variables Privadas -------------------------------------------------------*
Private oFrm            As frmSelPeriodo    ' Formulario de captura de fechas
Private mDatos          As BdDatos          ' Base de datos
Private oRangoDatos     As Range            ' Rango de datos a analizar
Private oFila           As Range            ' Current Fila
Private mFecha          As Date             ' Fecha
Private mFechaSorteo    As Date             ' Fecha de Sorteo
Private oSorteo         As Sorteo           ' Sorteo a comprobar
Private oApuesta        As Apuesta          ' Apuestas
Private oSorteoEngine   As SorteoEngine     ' Motor de sorteos
Private i               As Integer          ' Contador
Private iNumero         As Integer          ' Numero
Private oBoleto         As Boleto           ' Objeto Boleto
Private oInfo           As InfoSorteo       ' Información del sorteo
Private oCheckBoleto    As ComprobarBoletos ' Comprobador de apuestas
Private sCategoriaPrm   As String           ' Categoria del premio
'--- Constantes ---------------------------------------------------------------*
'--- Mensajes -----------------------------------------------------------------*
'--- Errores ------------------------------------------------------------------*
'------------------------------------------------------------------------------*
' Procedure : btn_ComprobarBoletos
' Author    : CHARLY
' Date      : vie, 12/sep/2014 21:01:36
' Purpose   :
'------------------------------------------------------------------------------*
'
Public Sub btn_ComprobarBoletos()
  
  On Error GoTo btn_ComprobarBoletos_Error
    '
    '   Desactiva la presentación
    '
    CALCULOOFF
    '
    '   Crea el formulario de fecha
    '
    Set oFrm = New frmSelPeriodo
    '
    ' Localizar criterios de consulta
    '
    oFrm.Tag = ESTADO_INICIAL
    '
    '  Bucle de control del proceso
    '
    Do While oFrm.Tag <> BOTON_CERRAR
        
        ' Se inicializa el boton cerrar para salir del bucle
        oFrm.Tag = BOTON_CERRAR
        
        ' Se muestra el formulario y queda a la espera de funciones
        ' pulsando el botón ejecutar
        oFrm.Show vbModal
            
        'Se bifurca la función
        Select Case oFrm.Tag
                                    ' El usuario ha cerrado el
            Case ""                 ' cuadro de dialogo con la [X]
                oFrm.Tag = BOTON_CERRAR
            
            Case EJECUTAR
                '
                '   Ajusta el periodo a fechas de sorteo
                '
                Set oInfo = New InfoSorteo
                If Not oInfo.EsFechaSorteo(oFrm.Periodo.FechaInicial) Then
                    oFrm.Periodo.FechaInicial = oInfo.GetProximoSorteo(oFrm.Periodo.FechaInicial)
                End If
                If Not oInfo.EsFechaSorteo(oFrm.Periodo.FechaFinal) Then
                    oFrm.Periodo.FechaFinal = oInfo.GetAnteriorSorteo(oFrm.Periodo.FechaFinal)
                End If
                '
                '   Comprobamos la información de los Boletos
                '
                ComprobarBoletos oFrm.Periodo
                '
                '   Comprobamos la información de las apuestas
                '
                ComprobarApuestas oFrm.Periodo
                '
                '   Salimos del bucle
                '
                oFrm.Tag = BOTON_CERRAR
        End Select
    Loop
    '
    '  Se elimina de la memoria el formulario
    '
    Set oFrm = Nothing
    '
    '
    '
    Ir_A_Hoja ("Apuestas")
    '
    '  Activa la presentación
    '
    CALCULOON
    
   On Error GoTo 0
   Exit Sub
btn_ComprobarBoletos_Error:
   Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
   ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
   Call HandleException(ErrNumber, ErrDescription, ErrSource, "Lot_06_ComprobarBoletos.btn_ComprobarBoletos")
   Call MsgBox(ErrDescription, vbError Or vbSystemModal, NOMBRE_APLICACION)
   Call Trace("CERRAR")
End Sub




'---------------------------------------------------------------------------------------
' Procedure : ComprobarBoletos
' Author    : CHARLY
' Date      : mi., 12/ago/2020 17:02:06
' Purpose   : Comprueba los registros de Boletos de un periodo definido
'---------------------------------------------------------------------------------------
'
Private Sub ComprobarBoletos(vNewValue As Periodo)
   On Error GoTo ComprobarBoletos_Error
    '
    '   Definimos los objetos del proceso
    '
    Set mDatos = New BdDatos
    '
    '   Ajustamos la fecha final al último registro de datos
    '
    If vNewValue.FechaFinal > mDatos.UltimoResultado Then
        vNewValue.FechaFinal = mDatos.UltimoResultado
    End If
    '
    '   Obtenemos el rango de datos para el periodo
    '
    Set oRangoDatos = mDatos.GetBoletoInFechas(vNewValue)
    '
    '   Si no hay datos para el periodo informamos.
    '
    If oRangoDatos Is Nothing Then
        MsgBox "No existen datos para el periodo solicitado.", vbExclamation + vbOKOnly, "ComprobarBoleto"
        Exit Sub
    End If
    '
    ' Creamos el objeto Boleto
    '
    Set oBoleto = New Boleto
    Set oCheckBoleto = New ComprobarBoletos
    '
    '   Para cada Boleto
    '
    For Each oFila In oRangoDatos.Rows
        '
        '  Si es una fecha (columna C)
        '
        If IsDate(oFila.Cells(1, 3).Value) Then
            '
            ' Obtenemos la fecha
            '
            mFecha = oFila.Cells(1, 3).Value
            '
            ' Si la fecha está contenida en el periodo de análisis
            '
            If vNewValue.Contiene(mFecha) Then
                    '
                    '   Obtenemos el Boleto del rango
                    '
                    oBoleto.Constructor oFila
                    '
                    '   Carga Apuestas
                    '
                    oBoleto.SetApuestas
                    If oBoleto.NumeroApuestas > 0 Then
                        '
                        '   Chequeamos el boleto si tiene apuestas
                        '
                        sCategoriaPrm = oCheckBoleto.ComprobarBoleto(oBoleto)
                        '
                        '
                        '
                        If Len(sCategoriaPrm) > 0 Then
                            '
                            '
                            '
                            oFila.Cells(1, 14).Value = sCategoriaPrm
                            oFila.Cells(1, 15).Value = oCheckBoleto.ImporteBoleto
                            oFila.Cells(1, 15).Interior.ColorIndex = COLOR_VERDE_CLARO
                        Else
                            oFila.Cells(1, 14).Value = ""
                            oFila.Cells(1, 15).Value = ""
                            oFila.Cells(1, 15).Interior.ColorIndex = xlColorIndexNone
                        End If
                    End If
            End If
        End If
    Next oFila
    '
    ' Inicializamos la memoria
    '
    Set oFila = Nothing
    Set oRangoDatos = Nothing
            
   On Error GoTo 0
    Exit Sub
ComprobarBoletos_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    Call HandleException(ErrNumber, ErrDescription, "Lot_06_ComprobarBoletos.ComprobarBoletos", ErrSource)
    Err.Raise ErrNumber, "Lot_06_ComprobarBoletos.ComprobarBoletos", ErrDescription
End Sub





'---------------------------------------------------------------------------------------
' Procedure : ComprobarApuestas
' Author    : CHARLY
' Date      : vie, 12/sep/2014 21:01:47
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub ComprobarApuestas(vNewValue As Periodo)
    
  On Error GoTo ComprobarApuestas_Error
    '
    '   Obtenemos el rango de comprobación
    '
    Set mDatos = New BdDatos
    '
    '   Ajustamos la fecha final al último registro de datos
    '
    If vNewValue.FechaFinal > mDatos.UltimoResultado Then
        vNewValue.FechaFinal = mDatos.UltimoResultado
    End If
    
    Set oRangoDatos = mDatos.GetApuestaInFechas(vNewValue)
    '
    '   Salimos si no hay datos
    '
    If oRangoDatos Is Nothing Then
        Exit Sub
    End If
    '
    ' Creamos el objeto sorteo
    '
    Set oSorteo = New Sorteo
    Set oApuesta = New Apuesta
    Set oSorteoEngine = New SorteoEngine
    Set oCheckBoleto = New ComprobarBoletos
    mFechaSorteo = #1/1/1900#
    '
    '   Para cada Apuesta
    '
    For Each oFila In oRangoDatos.Rows
        '
        '  Si es una fecha (columna D)
        '
        If IsDate(oFila.Cells(1, 4).Value) Then
            '
            ' Obtenemos la fecha
            '
            mFecha = oFila.Cells(1, 4).Value
            '
            ' Si la fecha está contenida en el periodo de análisis
            '
            If vNewValue.Contiene(mFecha) Then
                '
                ' si la fecha es distinta a la del sorteo, localizamos el sorteo
                '
                If mFecha <> mFechaSorteo Then
                    Set oSorteo = oSorteoEngine.GetSorteoByFecha(mFecha)
                    If oSorteo Is Nothing Then
                        mFechaSorteo = #1/1/1900#
                    Else
                        mFechaSorteo = oSorteo.Fecha
                    End If
                End If
                '
                '   Si encontramos el Sorteo
                '
                If Not (oSorteo Is Nothing) Then
                    '
                    '   Creamos la apuesta
                    '
                    oApuesta.Constructor oFila
                    '
                    'Coloreamos los Numeros acertados
                    '   Desde columna G a la Q
                    For i = 7 To 17
                        If IsNumeric(oFila.Cells(1, i).Value) Then
                            '
                            ' Obtenemos el número del rango
                            '
                            iNumero = oFila.Cells(1, i).Value
                            '
                            ' Si está contenido en el sorteo
                            '
                            If oSorteo.Combinacion.Contiene(iNumero) Then
                                '
                                ' Si no es el complementario lo colorea de verde
                                '
                                oFila.Cells(1, i).Interior.ColorIndex = COLOR_VERDE
                            Else
                                '
                                ' Si no está acertado lo deja sin color
                                '
                                oFila.Cells(1, i).Interior.ColorIndex = xlColorIndexNone
                            End If
                            '
                            ' para los sorteos de juego 6/49
                            '
                            If oSorteo.Complementario = iNumero Then
                                '
                                ' Si no es el complementario lo colorea de verde
                                '
                                oFila.Cells(1, i).Interior.ColorIndex = COLOR_AMARILLO
                            End If
                        End If
                    Next i
                    '
                    '   Comprobamos la apuesta
                    '
                    sCategoriaPrm = oCheckBoleto.ComprobarApuesta(oApuesta, False)
                    '
                    ' Obtenemos el premio y el coste
                    '
                    If oCheckBoleto.CatPremioApuesta <> Ninguna Then
                            oFila.Cells(1, 23).Value = oCheckBoleto.CategoriaPremioTxt
                            oFila.Cells(1, 23).Interior.ColorIndex = COLOR_VERDE_CLARO
                            oFila.Cells(1, 24).Value = sCategoriaPrm
                            oFila.Cells(1, 24).Interior.ColorIndex = COLOR_VERDE_CLARO
                        Else
                            '
                            ' si no lo tenemos ponemos las bolas acertadas
                            '
                            oFila.Cells(1, 23).Value = oCheckBoleto.CategoriaPremioTxt
                            oFila.Cells(1, 23).Interior.ColorIndex = xlColorIndexNone
                            oFila.Cells(1, 24).Value = ""   ' Categoria Premio
                            oFila.Cells(1, 24).Interior.ColorIndex = xlColorIndexNone
                    End If
                End If
            End If
        End If
    Next oFila
'
' Inicializamos la memoria
'
    Set oFila = Nothing
    Set oRangoDatos = Nothing
    
   On Error GoTo 0
   Exit Sub
ComprobarApuestas_Error:
   Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
   ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
   Call HandleException(ErrNumber, ErrDescription, ErrSource, "Lot_06_ComprobarBoletos.ComprobarApuestas")
   Err.Raise ErrNumber, "Lot_06_ComprobarBoletos.ComprobarApuestas", ErrDescription
End Sub


' *===========(EOF): Lot_06_ComprobarBoletos.bas
