Attribute VB_Name = "Lot_10_ComprobarApuestas"
'---------------------------------------------------------------------------------------
' Module    : Lot_ComprobarApuestas
' Author    : Charly
' Date      : 22/10/2013
' Purpose   : Verificar Pronósticos
'---------------------------------------------------------------------------------------
Option Explicit

Private DB                     As New BdDatos           'Objeto Base de Datos
Private mInfo                  As InfoSorteo            'Información de sorteos

'---------------------------------------------------------------------------------------
' Procedure : btn_ComprobarApuestas
' Author    : Charly
' Date      : 22/10/2013
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Sub btn_ComprobarApuestas()
    Dim ofrmPeriodo         As frmSelPeriodo
    Dim oParamCU            As ParametrosComprobarApuestas
    
  On Error GoTo btn_ComprobarApuestas_Error
    '
    '  Focalizar la hoja de salida
    '
    Ir_A_Hoja "MisApuestas"
    '
    '  Definir el formulario
    '
    Set ofrmPeriodo = New frmSelPeriodo
    '
    '  Bucle de control del proceso
    '
    ofrmPeriodo.Tag = ESTADO_INICIAL
        Do While ofrmPeriodo.Tag <> BOTON_CERRAR
            
        ' Se inicializa el boton cerrar para salir del bucle
        ofrmPeriodo.Tag = BOTON_CERRAR
        
        ' Se muestra el formulario y queda a la espera de funciones
        ' pulsando el botón ejecutar
        ofrmPeriodo.Show vbModal
        
        'Se bifurca la función
        Select Case ofrmPeriodo.Tag
                                    ' El usuario ha cerrado el
            Case ""                 ' cuadro de dialogo con la [X]
                ofrmPeriodo.Tag = BOTON_CERRAR
            
            Case EJECUTAR
            
                Application.ScreenUpdating = False
                '
                '  Definir Parametros del proceso
                '
                Set oParamCU = New ParametrosComprobarApuestas
                '
                '   Tipo de comparación: 0 Todo, 1 Vigencia  ##TODO: incluir selección en formulario
                '
                oParamCU.TipoComparacion = 1
                '
                '   Obtiene los datos del formulario
                '
                Set oParamCU.IntervaloFechas = ofrmPeriodo.RangoFechas
                '
                '  Obtener los sorteos del periodo
                '
                GetSorteos oParamCU
                '
                ' Comprobar las apuestas para cada sorteo
                '
                ProcesoComprobarApuestas oParamCU
                '
                Application.ScreenUpdating = True
                Set oParamCU = Nothing
        End Select
    Loop
    '
    '  Se elimina de la memoria el formulario
    '
    Set ofrmPeriodo = Nothing
    '
    '  Se elimina de la memoria los parametros del proceso
    '
    Set oParamCU = Nothing
   On Error GoTo 0
   Exit Sub
btn_ComprobarApuestas_Error:
     Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
   ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
   Call HandleException(ErrNumber, ErrDescription, "Lot_ComprobarApuestas.btn_ComprobarApuestas")
   Call MsgBox(ErrDescription, vbError Or vbSystemModal, NOMBRE_APLICACION)
   Call Trace("CERRAR")
End Sub




'---------------------------------------------------------------------------------------
' Procedure : ProcesoComprobarApuestas
' Author    : CHARLY
' Date      : dom, 12/oct/2014 00:56:00
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Sub ProcesoComprobarApuestas(oParam As ParametrosComprobarApuestas)
    Dim oCurrentRango   As Range
    Dim oFila           As Range
    Dim oApuesta        As Apuesta
    Dim sKey            As String

   On Error GoTo ProcesoComprobarApuestas_Error
    '
    '  Nos posicionamos en la celda A2 donde está la tabla
    Range("A2").Select
    '
    ' Eliminamos Filtros, si existen
    '
    If (ActiveSheet.AutoFilterMode And _
        ActiveSheet.FilterMode) _
    Or ActiveSheet.FilterMode Then
        ActiveSheet.ShowAllData
    End If
    '
    '  Localizamos el rango de apuestas
    Set oCurrentRango = Range("A2").CurrentRegion
    '
    '  Inicializamos la clave de la colección apuestas
    sKey = 0
    '
    ' Escribimos la cabecera variable de la tabla
    '
    PonCabecera oParam
    '
    '   Recorremos el rango fila a fila
    '
    For Each oFila In oCurrentRango.Rows
        '
        '  Creo una apuesta nueva
        Set oApuesta = New Apuesta
        '
        '  Obtengo la apuesta a partir de la fila
        Set oApuesta = GetApuesta(oFila)
        '
        '  Si es una apuesta válida
        If Not (oApuesta Is Nothing) Then
            '
            '  Comprueba apuesta con sorteos
            '
            VerApuestasSorteos oApuesta, oParam
            '
            '   Visualizamos los resultados de la apuesta
            '
            VisualizaResultado oFila.Row, oParam
            '
            ' Se borra la apuesta
            Set oApuesta = Nothing
        End If
    Next oFila
    '
    ' Se seleccionan todas las celdas y se autoformatean
    '
    Cells.Select                'Selecciona todas las celdas de la hoja
    Cells.EntireColumn.AutoFit  'Autoajusta el tamaño de las columnas
    '
    ' Creamos un autofiltro
    '
    Range("A3").Select          'Se posiciona en la celda del primer número
    Selection.AutoFilter        'Crea un autofiltro
            
   On Error GoTo 0
       Exit Sub
            
ProcesoComprobarApuestas_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    Call HandleException(ErrNumber, ErrDescription, "Lot_01_ComprobarApuestas.ProcesoComprobarApuestas", ErrSource)
    Err.Raise ErrNumber, "Lot_01_ComprobarApuestas.ProcesoComprobarApuestas", ErrDescription
End Sub



'---------------------------------------------------------------------------------------
' Procedure : PonCabecera
' Author    : CHARLY
' Date      : 18/04/2014
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub PonCabecera(oParametros As ParametrosComprobarApuestas)
    Dim m_rgCabecera    As Range
    Dim i               As Integer
    Dim tmpFInicial     As Date
    Dim tmpFFinal       As Date
    Dim tmpFecha        As Date
   
   On Error GoTo PonCabecera_Error
    '
    ' Borramos el área de salida de la información
    '
    Set m_rgCabecera = Range("P2").CurrentRegion
    m_rgCabecera.Offset(0, 16).Delete
    Set mInfo = New InfoSorteo
    mInfo.Constructor JUEGO_DEFECTO
    '
    ' Calculamos el número de elementos de la cabecera
    '
    tmpFInicial = oParametros.IntervaloFechas.FechaInicial
    tmpFFinal = oParametros.IntervaloFechas.FechaFinal
    i = mInfo.GetSorteosEntreFechas(tmpFInicial, tmpFFinal) + 4
    ReDim m_sCampos(i)
    '
    '   Componemos los literales de la cabecera
    '
    i = 0
    For tmpFecha = tmpFInicial To tmpFFinal
        If mInfo.EsFechaSorteo(tmpFecha) Then
            m_sCampos(i) = Format(tmpFecha, "ddd, dd/MM/yyyy")
            i = i + 1
        End If
    Next tmpFecha
    m_sCampos(i) = "Costes": m_sCampos(i + 1) = "Premios":
    m_sCampos(i + 2) = "Dias": m_sCampos(i + 3) = "Puntuacion"
        
    Set m_rgCabecera = ActiveSheet.Range("Q2")
    For i = 0 To UBound(m_sCampos)
        With m_rgCabecera.Offset(0, i)
                .Value = m_sCampos(i)
                .Font.Name = "Arial"
                .Font.FontStyle = "Negrita"
                .Font.Size = 10
                .Font.ThemeColor = xlThemeColorDark1
                .Font.TintAndShade = 0
                .VerticalAlignment = xlBottom
                .Orientation = 90
                .Interior.Pattern = xlSolid
                .Interior.PatternColorIndex = xlColorIndexAutomatic
                .Interior.TintAndShade = 0
                .Interior.PatternTintAndShade = 0
                .Interior.color = 12419407   'azul
        End With
    Next i
   On Error GoTo 0
   Exit Sub

PonCabecera_Error:
   Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
   ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
   Call HandleException(ErrNumber, ErrDescription, "Lot_01_ComprobarApuestas.PonCabecera")
   Err.Raise ErrNumber, ErrSource, ErrDescription
End Sub



'---------------------------------------------------------------------------------------
' Procedure : GetSorteos
' Author    : Charly
' Date      : 10/11/2013
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub GetSorteos(oParametros As ParametrosComprobarApuestas)
    Dim rgSorteos As Range
    Dim oSorteo   As Sorteo
    Dim oFila     As Range
    Dim sKey      As String
  
  On Error GoTo GetSorteos_Error
    '
    '  Inicializamos la clave de la colección apuestas
    sKey = 0
    
    On Error Resume Next
    Set rgSorteos = DB.Resultados_Fechas(oParametros.IntervaloFechas.FechaInicial, _
                                         oParametros.IntervaloFechas.FechaFinal)
    
    If rgSorteos Is Nothing _
    Or Err.Number = 100 Then
        Exit Sub
    End If
    
    On Error GoTo GetSorteos_Error
    '
    '   Recorremos los sorteos
    '
    For Each oFila In rgSorteos.Rows
        
        Set oSorteo = New Sorteo
        oSorteo.Constructor oFila
        '
        '  Si es una apuesta válida
        If Not (oSorteo Is Nothing) Then
            '
            ' Se agrega a la colección
            sKey = Format(oSorteo.EntidadNegocio.Id, "000000")
            oParametros.ColSorteos.Add oSorteo, sKey
            '
            ' Se borra el objeto sorteo
            Set oSorteo = Nothing
        End If
    
    Next oFila
  
  On Error GoTo 0
    Exit Sub
GetSorteos_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    Call HandleException(ErrNumber, ErrDescription, "Lot_01_ComprobarApuestas.GetSorteos")
    Err.Raise ErrNumber, ErrSource, ErrDescription
End Sub

'---------------------------------------------------------------------------------------
' Procedure : VerApuestasSorteos
' Author    : CHARLY
' Date      : dom, 12/oct/2014 01:17:57
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub VerApuestasSorteos(oApuesta As Apuesta, oParam As ParametrosComprobarApuestas)
    Dim oCUComprobar    As CU_ComprobarApuesta
    Dim oEstdstk        As EstadisticasApuesta
    Dim oSorteo         As Sorteo
    Dim oPremio         As Premio
    Dim sKey            As String

   On Error GoTo VerApuestasSorteos_Error
    '
    '
    '
    Set oCUComprobar = New CU_ComprobarApuesta
    '
    '
    '
    oParam.InitColAciertos
    oParam.InitColEstadisticas
    '
    '  Creamos la estadistica de una apuesta
    '
    Set oEstdstk = New EstadisticasApuesta
    oEstdstk.IdApuesta = oApuesta.EntidadNegocio.Id
    '
    '  Comprobamos cada uno de los sorteos
    '
    For Each oSorteo In oParam.ColSorteos
        '
        '   Si la apuesta está vigente
        '
        If (oApuesta.FechaAlta <= oSorteo.Fecha And _
            oSorteo.Fecha <= oApuesta.FechaFinVigencia) Or _
            (oParam.TipoComparacion = 0) Then
            '
            '  Creamos el premio
            '
            Set oPremio = New Premio
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
            '
            If oPremio.BolasAcertadas > 0 Then
                '
                ' Construimos la clave de la coleccion con
                ' el id de la apuesta y la fecha del sorteo
                '
                sKey = Format(oSorteo.Fecha, "yyyy-MM-dd")
                oPremio.Key = sKey
                '
                ' Se agrega a la colección
                '
                oParam.ColAciertos.Add oPremio, sKey
                '
                '  acumulamos a la estadistica de la apuesta
                '
                With oEstdstk
                    .Costes = .Costes + oApuesta.Coste(oSorteo.Juego)
                    .DiasAciertos = .DiasAciertos + 1
                    .ImportePremios = .ImportePremios + oPremio.GetPremioEsperado
                    .Puntuacion = .Puntuacion + CalPuntuacion(oPremio.BolasAcertadas)
                End With
            Else
                oEstdstk.Costes = oEstdstk.Costes + oApuesta.Coste(oSorteo.Juego)
            End If
            '
            ' Se borra el premio
            Set oPremio = Nothing
        End If
    Next oSorteo
    '
    '  agregamos la estadistica de la apuesta a la colección
    '
    oParam.ColEstaditicas.Add oEstdstk
    '
    '  Se borra la estadistica
    '
    Set oEstdstk = Nothing
   On Error GoTo 0
       Exit Sub
            
VerApuestasSorteos_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    Call HandleException(ErrNumber, ErrDescription, "Lot_01_ComprobarApuestas.VerApuestasSorteos", ErrSource)
    '   Lanza el error
    Err.Raise ErrNumber, "Lot_01_ComprobarApuestas.VerApuestasSorteos", ErrDescription
End Sub

'---------------------------------------------------------------------------------------
' Procedure : VisualizaResultado
' Author    : CHARLY
' Date      : dom, 26/oct/2014 23:41:50
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub VisualizaResultado(datRow As Long, oParam As ParametrosComprobarApuestas)
    Dim oPremio         As Premio
    Dim oEstd           As EstadisticasApuesta
    Dim mFecha          As Date
    Dim xCol            As Long
    Dim sKey            As String
    
   On Error GoTo VisualizaResultado_Error
    '
    '
    '
    Set oPremio = New Premio
    '
    '
    '
    xCol = 17
    '
    '
    '
    For mFecha = oParam.IntervaloFechas.FechaInicial _
    To oParam.IntervaloFechas.FechaFinal
        If (Weekday(mFecha) <> 1) Then      ' Si no es Domingo
            '
            '
            '
            sKey = Format(mFecha, "yyyy-MM-dd")
            '
            ' Localizamos el premio de la apuesta para el sorteo
            ' en la coleccion de aciertos
            '
            On Error Resume Next
            Set oPremio = oParam.ColAciertos.Item(sKey)
            '
            ' Si lo encontramos Err = 0 sino Err = 5
            '
            If Err.Number = 0 Then
                '
                ' Volvemos a activar el controlador de error
                '
                On Error GoTo VisualizaResultado_Error
                '
                ' Si tenemos un premio coloreamos la celda
                '
                If oPremio.CategoriaPremio <> Ninguna Then
                    Cells(datRow, xCol).Value = oPremio.LiteralCategoriaPremio
                    Cells(datRow, xCol).Interior.ColorIndex = COLOR_VERDE_CLARO
                Else
                    '
                    ' si no lo tenemos ponemos las bolas acertadas
                    '
                     Cells(datRow, xCol).Value = oPremio.BolasAcertadas
                End If
            End If
            '
            ' Incrementamos la siguiente columna
            '
            xCol = xCol + 1
        End If
    Next mFecha
    '
    ' Al final localizamos las estadisticas de la apuesta
    '
    On Error Resume Next
    Set oEstd = oParam.ColEstaditicas.Item(1)
    '
    ' Si las encontramos
    '
    If Err.Number = 0 Then
        With Cells(datRow, xCol)
            .Value = oEstd.Costes
            .NumberFormat = FMT_IMPORTE
        End With
        With Cells(datRow, xCol + 1)
             .Value = oEstd.ImportePremios
            .NumberFormat = FMT_IMPORTE
        End With
        Cells(datRow, xCol + 2).Value = oEstd.DiasAciertos
        Cells(datRow, xCol + 3).Value = oEstd.Puntuacion
    End If
    Set oEstd = Nothing
            
   On Error GoTo 0
       Exit Sub
            
VisualizaResultado_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    Call HandleException(ErrNumber, ErrDescription, "Lot_01_ComprobarApuestas.VisualizaResultado", ErrSource)
    '   Lanza el error
    Err.Raise ErrNumber, "Lot_01_ComprobarApuestas.VisualizaResultado", ErrDescription
End Sub
'
'  ****************************************************************************************************
'                   FUNCIONES
'  ****************************************************************************************************
'

'---------------------------------------------------------------------------------------
' Procedure : GetApuesta
' Author    : Charly
' Date      : 10/11/2013
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Function GetApuesta(datFila As Range) As Apuesta
    Dim mValor As Variant
    Dim objResult As Apuesta
    Dim i As Integer
    Dim n As Numero
    
  On Error GoTo GetApuesta_Error
    '
    '   Creamos el objetp
    Set objResult = New Apuesta
    '
    '   Extraemos el valor de la columna 1
    mValor = datFila.Value2(1, 1)
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
    For i = 1 To UBound(datFila.Value2, 2)
        '
        '   Obtenemos el valor de la celda
        mValor = datFila.Value2(1, i)
        '
        '   Segun su posición en la columna se asigna a una propiedad
        Select Case i
            Case 1: objResult.EntidadNegocio.Id = CInt(mValor)
            Case 2: objResult.FechaAlta = CDate(mValor)
            Case 3: objResult.FechaFinVigencia = mInfo.AddDiasSorteo(objResult.FechaAlta, CInt(mValor))
            Case 4: objResult.Metodo = mValor
            Case 5 To 15
                If IsNumeric(mValor) And _
                Not IsEmpty(mValor) Then
                    Set n = New Numero
                    n.Valor = CInt(mValor)
                    objResult.Combinacion.Add n
                    Set n = Nothing
                End If
        End Select
    Next i
    Set GetApuesta = objResult
    Set objResult = Nothing
    
   On Error GoTo 0
   Exit Function

GetApuesta_Error:
     Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
   ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
   '   Audita el error
   Call HandleException(ErrNumber, ErrDescription, "Lot_01_ComprobarApuestas.GetApuesta")
   '   Lanza el Error
   Err.Raise ErrNumber, ErrSource, ErrDescription
End Function

'---------------------------------------------------------------------------------------
' Procedure : CalPuntuacion
' Author    : CHARLY
' Date      : 07/02/2014
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Function CalPuntuacion(datAciertos As Integer) As Integer
    Dim iResult As Integer
'
'TODO: Agregar la categoria del premio y puntuar diferente
'
   On Error GoTo CalPuntuacion_Error
    iResult = 0
    Select Case datAciertos
        Case 1: iResult = 200
        Case 2: iResult = 400
        Case 3: iResult = 800
        Case 4: iResult = 900
        Case 5: iResult = 1000
        Case 6: iResult = 1100
        Case 7: iResult = 1500
    End Select
    
    CalPuntuacion = iResult

   On Error GoTo 0
   Exit Function

CalPuntuacion_Error:
   Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
   ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
   '   Audita el error
   Call HandleException(ErrNumber, ErrDescription, "Lot_01_ComprobarApuestas.CalPuntuacion")
   '   Lanza el Error
   Err.Raise ErrNumber, ErrSource, ErrDescription
 
End Function

