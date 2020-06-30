Attribute VB_Name = "Lot_01_ComprobarBoletos"
'---------------------------------------------------------------------------------------
' Module    : Lot_06_ComprobarBoletos
' DateTime  : 06/09/2014 19:05
' Author    : Carlos Almela Baeza
' Purpose   : Caso de Uso Comprobar Boletos Apostados. Comprueba los pronosticos acertados
'---------------------------------------------------------------------------------------
Option Explicit
Option Base 0
'
'
'
'---------------------------------------------------------------------------------------
' Procedure : btn_ComprobarBoletos
' Author    : CHARLY
' Date      : vie, 12/sep/2014 21:01:36
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Sub btn_ComprobarBoletos()
    Dim oFrm        As frmSelPeriodo
  
  On Error GoTo btn_ComprobarBoletos_Error
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
                ComprobarBoletos oFrm.Periodo
                ComprobarApuestas oFrm.Periodo
                oFrm.Tag = BOTON_CERRAR
        End Select
    Loop
    '
    '  Se elimina de la memoria el formulario
    '
    Set oFrm = Nothing
   On Error GoTo 0
   Exit Sub
btn_ComprobarBoletos_Error:
     Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
   ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
   '   Audita el error
   Call HandleException(ErrNumber, ErrDescription, ErrSource, "Lot_06_ComprobarBoletos.btn_ComprobarBoletos")
   '   Informa del error
   Call MsgBox(ErrDescription, vbError Or vbSystemModal, NOMBRE_APLICACION)
   Call Trace("CERRAR")
End Sub

'---------------------------------------------------------------------------------------
' Procedure : ComprobarApuestas
' Author    : CHARLY
' Date      : vie, 12/sep/2014 21:01:47
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub ComprobarApuestas(vNewValue As Periodo)
    Dim oRangoDatos         As Range
    Dim oFila               As Range
    Dim mFecha              As Date
    Dim mFechaSorteo        As Date
    Dim oSorteo             As Sorteo
    Dim oApuesta            As Apuesta
    Dim oSorteoEngine       As SorteoEngine
    Dim oCUComprobar        As CU_ComprobarApuesta
    Dim oPremio             As Premio
    Dim eJuego              As Juego
    Dim i                   As Integer
    Dim iNumero             As Integer
    
On Error GoTo ComprobarApuestas_Error
        ThisWorkbook.Sheets("Apuestas").Activate
    Set oRangoDatos = ThisWorkbook.Sheets("Apuestas").Range("A1").CurrentRegion
    '
    ' Creamos el objeto sorteo
    '
    Set oSorteo = New Sorteo
    Set oApuesta = New Apuesta
    Set oSorteoEngine = New SorteoEngine
    Set oCUComprobar = New CU_ComprobarApuesta
    Set oPremio = New Premio
     mFechaSorteo = #1/1/1900#
    '
    '   Para cada Apuesta
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
                If Not (oSorteo Is Nothing) Then
                    '
                    '   Obtenemos la apuesta del rango
                    '
                    Set oApuesta = GetApuesta(oFila)
                    '
                    '  Enfrentamos apuesta a sorteo
                    '
                    Set oCUComprobar.MyApuesta = oApuesta
                    Set oCUComprobar.Sorteo = oSorteo
                    '
                    ' Obtenemos el premio de la apuesta
                    '
                    Set oPremio = oCUComprobar.GetPremio
                    '
                    '
                    oPremio.FechaSorteo = oSorteo.Fecha
                    oPremio.ModalidadJuego = oSorteo.Juego
                    eJuego = oSorteo.Juego
                    '
                    'Coloreamos los Numeros acertados
                    '
                    For i = 6 To 14
                        '
                        ' Obtenemos el número del rango
                        '
                        iNumero = oFila.Cells(1, i).Value
                        '
                        ' Si está contenido en el sorteo
                        '
                        If oSorteo.Combinacion.Contiene(iNumero) Then
                            '
                            ' para los sorteos de juego 6/49
                            '
                            If oSorteo.Complementario <> iNumero Then
                                '
                                ' Si no es el complementario lo colorea de verde
                                '
                                oFila.Cells(1, i).Interior.ColorIndex = COLOR_VERDE
                            Else
                                '
                                ' Si es el complmentario de amarillo
                                '
                                oFila.Cells(1, i).Interior.ColorIndex = COLOR_AMARILLO
                            End If
                        Else
                        '
                        ' Si no está acertado lo deja sin color
                        '
                            oFila.Cells(1, i).Interior.ColorIndex = xlColorIndexNone
                        End If
                    Next i
                    '
                    ' Obtenemos el premio y el coste
                    '
                    If oPremio.CategoriaPremio <> Ninguna Then
                            oFila.Cells(1, 15).Value = oPremio.LiteralCategoriaPremio
                            oFila.Cells(1, 15).Interior.ColorIndex = COLOR_VERDE_CLARO
                            oFila.Cells(1, 17).Value = oApuesta.Coste(eJuego)
                            oFila.Cells(1, 18).Value = oPremio.GetPremioEsperado
                            
                        Else
                            '
                            ' si no lo tenemos ponemos las bolas acertadas
                            '
                            oFila.Cells(1, 15).Value = oPremio.BolasAcertadas
                            oFila.Cells(1, 15).Interior.ColorIndex = xlColorIndexNone
                            oFila.Cells(1, 17).Value = oApuesta.Coste(eJuego)
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
   '   Audita el error
   Call HandleException(ErrNumber, ErrDescription, ErrSource, "Lot_06_ComprobarBoletos.ComprobarApuestas")
   '   Lanza el error
   Err.Raise ErrNumber, "Lot_06_ComprobarBoletos.ComprobarApuestas", ErrDescription

End Sub

'---------------------------------------------------------------------------------------
' Procedure : ComprobarBoletos
' Author    : CHARLY
' Date      : vie, 12/sep/2014 21:02:46
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub ComprobarBoletos(vNewValue As Periodo)
    Dim oRangoDatos         As Range
    Dim oFila               As Range
    Dim mFecha              As Date
    Dim mFechaSorteo        As Date
    Dim oSorteo             As Sorteo
    Dim oSorteoEngine       As SorteoEngine
    Dim oCUComprobar        As CU_ComprobarApuesta
    Dim oPremio             As Premio
    Dim eJuego              As Juego
    Dim i                   As Integer
    Dim iNumero             As Integer
    Dim oBoleto             As Boleto

   On Error GoTo ComprobarBoletos_Error
    ThisWorkbook.Sheets("Boletos").Activate
    Set oRangoDatos = ThisWorkbook.Sheets("Boletos").Range("A1").CurrentRegion
    '
    ' Creamos el objeto sorteo
    '
    Set oSorteo = New Sorteo
    Set oSorteoEngine = New SorteoEngine
    Set oBoleto = New Boleto
    mFechaSorteo = #1/1/1900#
    '
    '   Para cada Boleto
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
                End If
                If oSorteo Is Nothing Then
                    Debug.Print "Sorteo no encontrado para fecha => " & mFecha
                    mFechaSorteo = #1/1/1900#
                Else
                    mFechaSorteo = oSorteo.Fecha
                '
                '   Obtenemos el Boleto del rango
                '
                oBoleto.Constructor oFila
                If oBoleto.Reintegro = oSorteo.Reintegro Then
                    '
                    '   Si el juego es bonoloto y el boleto multiple
                    '   solo hay reintegro si es viernes
                    '
                    If oBoleto.Semanal And (oBoleto.Juego = "BL") Then
                        oFila.Cells(1, 11).Value = oBoleto.Coste
                    Else
                        If oBoleto.Semanal And (oBoleto.Juego = "LP") Then
                            oFila.Cells(1, 11).Value = oBoleto.Coste / 2
                        Else
                            oFila.Cells(1, 11).Value = oBoleto.Coste
                        End If
                    End If
                Else
                    oFila.Cells(1, 11).Value = ""
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
    '   Lanza el error
    Err.Raise ErrNumber, "Lot_06_ComprobarBoletos.ComprobarBoletos", ErrDescription
End Sub

'---------------------------------------------------------------------------------------
' Procedure : GetApuesta
' Author    : CHARLY
' Date      : mié, 10/sep/2014 00:02:52
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Function GetApuesta(vNewValue As Range) As Apuesta
    Dim mValor      As Variant
    Dim objResult   As Apuesta
    Dim i           As Integer
    Dim n           As Numero
    
   On Error GoTo GetApuesta_Error
    '
    '   Creamos el objeto
    Set objResult = New Apuesta
    '
    '   Extraemos el valor de la columna 1
    mValor = vNewValue.Value2(1, 1)
    '
    '
    '   Si no es un valor numérico Columna N
    If (Not IsNumeric(mValor)) _
    Or IsEmpty(mValor) Then
        Exit Function
    End If
    '
    '  Para cada Item de la fila se analiza la posición
    '  y se asigna a una propiedad del objeto
    For i = 1 To UBound(vNewValue.Value2, 2)
        '
        '   Obtenemos el valor de la celda
        mValor = vNewValue.Value2(1, i)
        '
        '   Segun su posición en la columna se asigna a una propiedad
        '
        '   1 - Id
        '   2 - idBoleto
        '   3 - Fecha
        '   4 - Juego
        '   5 - Semana
        '   6 To 14 -  N1  N2  N3  N4  N5  N6  N7  N8  N9
        '   15 - Aciertos
        '   16 - Metodo
        '   17 - Coste
        '   18 - Premio
        '
        Select Case i
            Case 1: objResult.EntidadNegocio.Id = CInt(mValor)  'Id
            Case 2: objResult.IdBoleto = CInt(mValor)           'IdBoleto
            Case 3: objResult.FechaAlta = CDate(mValor)         'Fecha Alta
            Case 6 To 14
                    If IsNumeric(mValor) And _
                    Not IsEmpty(mValor) Then
                        Set n = New Numero
                        n.Valor = CInt(mValor)
                        objResult.Combinacion.Add n
                        Set n = Nothing
                    End If
            Case 16: objResult.Metodo = mValor                   'MEtodo
        End Select
    Next i
    Set GetApuesta = objResult
    Set objResult = Nothing
            
   On Error GoTo 0
       Exit Function
            
GetApuesta_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    Call HandleException(ErrNumber, ErrDescription, "GetApuesta.Lot_06_ComprobarBoletos", ErrSource)
    '   Lanza el error
    Err.Raise ErrNumber, "GetApuesta.Lot_06_ComprobarBoletos", ErrDescription
 
End Function
